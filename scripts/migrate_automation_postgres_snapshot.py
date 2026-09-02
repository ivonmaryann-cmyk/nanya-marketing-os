"""Copy the audited Docker PostgreSQL automation domain into an empty target.

This is a one-time, transaction-safe migration tool.  It intentionally keeps
the Docker source read-only and refuses to overwrite an already populated
automation domain.  It is not an application runtime synchronisation tool.
"""

from __future__ import annotations

import argparse
import os
import subprocess
from collections.abc import Iterable

import psycopg


SOURCE_DOCKER_CONTAINER = "nanya-marketing-os-ivon-dev-analysis-automation-postgres-1"
SOURCE_DATABASE = "nanya_automation"
SOURCE_USER = "nanya_automation"

# Migration helper/log tables are deliberately excluded.  They are neither
# source business data nor a valid input for a new target migration history.
BUSINESS_TABLES = (
    "automation_metadata",
    "automation_customers",
    "automation_customer_contacts",
    "automation_customer_events",
    "automation_customer_extraction_maps",
    "automation_customer_routing_rules",
    "automation_customer_routing_conditions",
    "automation_customer_spec_mappings",
    "mail_accounts",
    "mail_messages",
    "mail_attachments",
    "mail_attachment_texts",
    "mail_fetch_logs",
    "mail_fetch_tasks",
    "mail_fetch_task_messages",
    "mail_order_tasks",
    "mail_transcode_jobs",
    "order_change_tags",
    "order_change_tag_keywords",
    "order_mail_routing_rules",
    "order_mail_routing_rule_events",
    "order_mail_rule_groups",
    "order_mail_rule_keywords",
    "order_intake_cases",
    "order_intake_case_events",
    "order_entry_templates",
    "order_entry_template_lines",
    "order_entry_template_versions",
    "order_entry_template_tasks",
    "order_entry_detail_events",
    "order_interface_configs",
    "order_interface_config_versions",
    "order_interface_call_logs",
    "order_material_query_suggestions",
    "order_material_resolution_tasks",
)

PRIMARY_KEYS = {
    "automation_metadata": ("key",),
    "mail_attachment_texts": ("attachment_id",),
    "mail_fetch_task_messages": ("fetch_task_id", "mail_id"),
}


def _quote(identifier: str) -> str:
    if not identifier.replace("_", "").isalnum():
        raise ValueError(f"unsafe SQL identifier: {identifier}")
    return f'"{identifier}"'


def _source_command(args: argparse.Namespace, statement: str) -> list[str]:
    return [
        "docker", "exec", args.source_container,
        "psql", "-X", "-q", "-A", "-t", "-U", args.source_user, "-d", args.source_database,
        "-v", "ON_ERROR_STOP=1", "-c", statement,
    ]


def _source_columns(args: argparse.Namespace) -> dict[str, tuple[str, ...]]:
    sql = """
        SELECT table_name || E'\\t' || column_name
        FROM information_schema.columns
        WHERE table_schema='public'
        ORDER BY table_name, ordinal_position
    """
    output = subprocess.run(
        _source_command(args, sql), check=True, capture_output=True, text=True,
    ).stdout
    columns: dict[str, list[str]] = {}
    for line in output.splitlines():
        table, column = line.split("\t", 1)
        columns.setdefault(table, []).append(column)
    return {table: tuple(values) for table, values in columns.items()}


def _target_columns(target: psycopg.Connection, tables: Iterable[str]) -> dict[str, tuple[str, ...]]:
    with target.cursor() as cursor:
        cursor.execute(
            """
            SELECT table_name, column_name
            FROM information_schema.columns
            WHERE table_schema='public' AND table_name = ANY(%s)
            ORDER BY table_name, ordinal_position
            """,
            (list(tables),),
        )
        columns: dict[str, list[str]] = {}
        for table, column in cursor.fetchall():
            columns.setdefault(table, []).append(column)
    return {table: tuple(values) for table, values in columns.items()}


def _source_csv(args: argparse.Namespace, table: str, columns: tuple[str, ...]) -> bytes:
    primary_key = PRIMARY_KEYS.get(table, ("id",))
    selected = ",".join(_quote(column) for column in columns)
    order = ",".join(_quote(column) for column in primary_key)
    statement = (
        f"COPY (SELECT {selected} FROM {_quote(table)} ORDER BY {order}) "
        "TO STDOUT WITH (FORMAT CSV, HEADER true)"
    )
    return subprocess.run(
        _source_command(args, statement), check=True, capture_output=True,
    ).stdout


def _count_source_rows(args: argparse.Namespace, table: str) -> int:
    output = subprocess.run(
        _source_command(args, f"SELECT count(*) FROM {_quote(table)}"),
        check=True, capture_output=True, text=True,
    ).stdout.strip()
    return int(output)


def _count_target_rows(target: psycopg.Connection, table: str) -> int:
    with target.cursor() as cursor:
        cursor.execute(f"SELECT count(*) FROM {_quote(table)}")
        return int(cursor.fetchone()[0])


def _digest_sql(table: str, columns: tuple[str, ...]) -> str:
    primary_key = PRIMARY_KEYS.get(table, ("id",))
    selected = ",".join(_quote(column) for column in columns)
    ordering = ",".join(_quote(column) for column in primary_key)
    return (
        "SELECT count(*)::text || E'\\t' || "
        "COALESCE(md5(string_agg(md5(row_to_json(row_data)::text), '' "
        f"ORDER BY {ordering})), md5('')) "
        f"FROM (SELECT {selected} FROM {_quote(table)}) AS row_data"
    )


def _source_digest(args: argparse.Namespace, table: str, columns: tuple[str, ...]) -> tuple[int, str]:
    output = subprocess.run(
        _source_command(args, _digest_sql(table, columns)),
        check=True, capture_output=True, text=True,
    ).stdout.strip()
    count, digest = output.split("\t", 1)
    return int(count), digest


def _target_digest(
    target: psycopg.Connection,
    table: str,
    columns: tuple[str, ...],
) -> tuple[int, str]:
    with target.cursor() as cursor:
        cursor.execute(_digest_sql(table, columns))
        count, digest = cursor.fetchone()[0].split("\t", 1)
    return int(count), digest


def _reset_sequence(target: psycopg.Connection, table: str) -> None:
    with target.cursor() as cursor:
        cursor.execute(
            "SELECT setval(pg_get_serial_sequence(%s, 'id'), "
            f"GREATEST(COALESCE((SELECT MAX(id) FROM {_quote(table)}), 0), 1), "
            f"EXISTS(SELECT 1 FROM {_quote(table)}))",
            (table,),
        )


def migrate(args: argparse.Namespace) -> dict[str, int]:
    target_url = os.environ.get(args.target_url_env)
    if not target_url:
        raise RuntimeError(f"required environment variable is empty: {args.target_url_env}")

    source_columns = _source_columns(args)
    missing = [table for table in BUSINESS_TABLES if table not in source_columns]
    if missing:
        raise RuntimeError(f"source is missing scoped business tables: {', '.join(missing)}")

    with psycopg.connect(target_url, connect_timeout=10) as target:
        target_columns = _target_columns(target, BUSINESS_TABLES)
        with target.transaction():
            populated = {}
            for table in BUSINESS_TABLES:
                count = _count_target_rows(target, table)
                if count:
                    populated[table] = count
            if populated:
                detail = ", ".join(f"{table}={count}" for table, count in populated.items())
                raise RuntimeError(f"target automation tables are not empty; refusing overwrite: {detail}")

            counts: dict[str, int] = {}
            for table in BUSINESS_TABLES:
                source_fields = source_columns[table]
                target_fields = target_columns.get(table)
                if not target_fields:
                    raise RuntimeError(f"target is missing scoped business table: {table}")
                unexpected = set(source_fields) - set(target_fields)
                if unexpected:
                    raise RuntimeError(f"target is missing source fields for {table}: {sorted(unexpected)}")

                csv_data = _source_csv(args, table, source_fields)
                with target.cursor() as cursor, cursor.copy(
                    f"COPY {_quote(table)}({','.join(_quote(field) for field in source_fields)}) "
                    "FROM STDIN WITH (FORMAT CSV, HEADER true)"
                ) as copy:
                    copy.write(csv_data)
                counts[table] = _count_target_rows(target, table)
                source_count = _count_source_rows(args, table)
                if counts[table] != source_count:
                    raise RuntimeError(
                        f"row count mismatch after copy for {table}: "
                        f"source={source_count}, target={counts[table]}"
                    )

            for table in BUSINESS_TABLES:
                if "id" in source_columns[table]:
                    _reset_sequence(target, table)
    return counts


def verify(args: argparse.Namespace) -> dict[str, str]:
    """Compare source and target rows by ordered, field-level digest only."""
    target_url = os.environ.get(args.target_url_env)
    if not target_url:
        raise RuntimeError(f"required environment variable is empty: {args.target_url_env}")
    source_columns = _source_columns(args)
    with psycopg.connect(target_url, connect_timeout=10) as target:
        target_columns = _target_columns(target, BUSINESS_TABLES)
        results: dict[str, str] = {}
        for table in BUSINESS_TABLES:
            fields = source_columns[table]
            if not set(fields).issubset(target_columns.get(table, ())):
                raise RuntimeError(f"target is missing source fields for {table}")
            source_count, source_hash = _source_digest(args, table, fields)
            target_count, target_hash = _target_digest(target, table, fields)
            if (source_count, source_hash) != (target_count, target_hash):
                results[table] = f"mismatch source={source_count}/{source_hash} target={target_count}/{target_hash}"
            else:
                results[table] = f"ok count={source_count}"
    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target-url-env", default="AUTOMATION_DATABASE_URL")
    parser.add_argument("--source-container", default=SOURCE_DOCKER_CONTAINER)
    parser.add_argument("--source-database", default=SOURCE_DATABASE)
    parser.add_argument("--source-user", default=SOURCE_USER)
    parser.add_argument("--verify-only", action="store_true")
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    if arguments.verify_only:
        verification = verify(arguments)
        failures = {table: result for table, result in verification.items() if not result.startswith("ok ")}
        print("verification=" + ",".join(f"{table}:{result}" for table, result in verification.items()))
        if failures:
            raise SystemExit(1)
    else:
        result = migrate(arguments)
        print("migration_complete=" + ",".join(f"{table}:{count}" for table, count in result.items()))
