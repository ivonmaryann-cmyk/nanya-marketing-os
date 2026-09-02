from __future__ import annotations

import json
import sqlite3

from .spec import PRIMARY_KEYS, TABLES


OUTBOX_TABLE = "automation_migration_outbox"


def _json_object(prefix: str, columns: tuple[str, ...]) -> str:
    parts = ", ".join(f"'{column}', {prefix}.\"{column}\"" for column in columns)
    return f"json_object({parts})"


def ensure_sqlite_outbox(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS automation_migration_outbox (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id TEXT NOT NULL UNIQUE,
            source_table TEXT NOT NULL,
            operation TEXT NOT NULL,
            pk_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            attempts INTEGER NOT NULL DEFAULT 0,
            next_attempt_at TEXT,
            last_error_type TEXT NOT NULL DEFAULT '',
            processed_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_automation_outbox_pending
            ON automation_migration_outbox(processed_at, next_attempt_at, id);

        CREATE TABLE IF NOT EXISTS automation_runtime_flags (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        INSERT INTO automation_runtime_flags(key,value) VALUES ('suppress_outbox','0')
            ON CONFLICT(key) DO NOTHING;

        CREATE TABLE IF NOT EXISTS automation_shadow_differences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            metric_name TEXT NOT NULL,
            sqlite_count INTEGER NOT NULL DEFAULT 0,
            postgresql_count INTEGER NOT NULL DEFAULT 0,
            sqlite_hash TEXT NOT NULL DEFAULT '',
            postgresql_hash TEXT NOT NULL DEFAULT '',
            elapsed_ms INTEGER NOT NULL DEFAULT 0,
            is_match INTEGER NOT NULL DEFAULT 0,
            observed_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_automation_shadow_observed
            ON automation_shadow_differences(observed_at, is_match, id);
        """
    )
    trigger_version = connection.execute(
        "SELECT value FROM automation_runtime_flags WHERE key='outbox_trigger_version'"
    ).fetchone()
    if trigger_version and trigger_version[0] == "2":
        return
    for table in TABLES:
        keys = PRIMARY_KEYS[table]
        for suffix, operation, prefix in (("ai", "insert", "NEW"), ("au", "update", "NEW"), ("ad", "delete", "OLD")):
            trigger = f"automation_outbox_{table}_{suffix}"
            connection.execute(f'DROP TRIGGER IF EXISTS "{trigger}"')
            connection.execute(
                f"""
                CREATE TRIGGER IF NOT EXISTS "{trigger}"
                AFTER {operation.upper()} ON "{table}"
                WHEN COALESCE((SELECT value FROM automation_runtime_flags WHERE key='suppress_outbox'),'0') <> '1'
                BEGIN
                    INSERT INTO {OUTBOX_TABLE}
                        (event_id, source_table, operation, pk_json, created_at)
                    VALUES (
                        lower(hex(randomblob(16))), '{table}', '{operation}',
                        {_json_object(prefix, keys)}, strftime('%Y-%m-%dT%H:%M:%fZ', 'now')
                    );
                END
                """
            )
    _ensure_metadata_triggers(connection)
    connection.execute(
        "INSERT INTO automation_runtime_flags(key,value) VALUES ('outbox_trigger_version','2') "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value"
    )


def _ensure_metadata_triggers(connection: sqlite3.Connection) -> None:
    condition = (
        "NEW.key LIKE 'order_mail_rule_%' OR NEW.key LIKE 'order_change_%' "
        "OR NEW.key LIKE 'order_intake_rule_engine_%'"
    )
    for suffix, operation, prefix, when in (
        ("ai", "INSERT", "NEW", condition),
        ("au", "UPDATE", "NEW", condition),
        ("ad", "DELETE", "OLD", condition.replace("NEW.", "OLD.")),
    ):
        connection.execute(f'DROP TRIGGER IF EXISTS "automation_outbox_metadata_{suffix}"')
        connection.execute(
            f"""
            CREATE TRIGGER IF NOT EXISTS "automation_outbox_metadata_{suffix}"
            AFTER {operation} ON settings
            WHEN ({when}) AND COALESCE((SELECT value FROM automation_runtime_flags WHERE key='suppress_outbox'),'0') <> '1'
            BEGIN
                INSERT INTO {OUTBOX_TABLE}
                    (event_id, source_table, operation, pk_json, created_at)
                VALUES (
                    lower(hex(randomblob(16))), 'automation_metadata', '{operation.lower()}',
                    json_object('key', {prefix}.key), strftime('%Y-%m-%dT%H:%M:%fZ', 'now')
                );
            END
            """
        )


def parse_primary_key(raw: str, table: str) -> dict[str, object]:
    parsed = json.loads(raw)
    expected = ("key",) if table == "automation_metadata" else PRIMARY_KEYS[table]
    if not isinstance(parsed, dict) or set(parsed) != set(expected):
        raise ValueError(f"invalid primary key for {table}")
    return {key: parsed[key] for key in expected}
