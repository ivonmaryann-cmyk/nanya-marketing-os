from __future__ import annotations

import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

from .spec import PRIMARY_KEYS, TABLES


def _quote(identifier: str) -> str:
    if not identifier.replace("_", "").isalnum():
        raise ValueError(f"unsafe SQL identifier: {identifier}")
    return f'"{identifier}"'


def _columns(source: sqlite3.Connection, table: str) -> list[str]:
    return [row["name"] for row in source.execute(f"PRAGMA table_info({_quote(table)})")]


def _upsert_sql(table: str, columns: list[str]) -> str:
    keys = PRIMARY_KEYS[table]
    quoted_columns = ", ".join(_quote(column) for column in columns)
    placeholders = ", ".join("%s" for _ in columns)
    conflict = ", ".join(_quote(column) for column in keys)
    mutable = [column for column in columns if column not in keys]
    if mutable:
        update = ", ".join(f'{_quote(column)}=EXCLUDED.{_quote(column)}' for column in mutable)
        action = f"DO UPDATE SET {update}"
    else:
        action = "DO NOTHING"
    return f"INSERT INTO {_quote(table)} ({quoted_columns}) VALUES ({placeholders}) ON CONFLICT ({conflict}) {action}"


def copy_snapshot(snapshot: Path, target: Any) -> dict[str, int]:
    """Idempotently copy all scoped tables. The caller owns the PostgreSQL transaction."""
    counts: dict[str, int] = {}
    with closing(sqlite3.connect(snapshot)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only=ON")
        with target.cursor() as cursor:
            for table in TABLES:
                columns = _columns(source, table)
                rows = source.execute(f"SELECT * FROM {_quote(table)}").fetchall()
                if rows:
                    cursor.executemany(
                        _upsert_sql(table, columns),
                        [tuple(row[column] for column in columns) for row in rows],
                    )
                counts[table] = len(rows)

            metadata_rows = source.execute(
                "SELECT key, value FROM settings WHERE "
                "key LIKE 'order_mail_rule_%' OR key LIKE 'order_change_%' OR key LIKE 'order_intake_rule_engine_%'"
            ).fetchall()
            if metadata_rows:
                cursor.executemany(
                    "INSERT INTO automation_metadata(key,value,updated_at) VALUES (%s,%s,%s) "
                    "ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at",
                    [(row["key"], row["value"], row["value"]) for row in metadata_rows],
                )
            counts["automation_metadata"] = len(metadata_rows)
    _reset_identity_sequences(target)
    return counts


def _reset_identity_sequences(target: Any) -> None:
    for table, keys in PRIMARY_KEYS.items():
        if keys != ("id",):
            continue
        target.execute(
            "SELECT setval(pg_get_serial_sequence(%s, 'id'), "
            "GREATEST(COALESCE((SELECT MAX(id) FROM " + _quote(table) + "), 0), 1), "
            "EXISTS(SELECT 1 FROM " + _quote(table) + "))",
            (table,),
        )
