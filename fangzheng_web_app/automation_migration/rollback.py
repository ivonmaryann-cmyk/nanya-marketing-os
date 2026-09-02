from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .copy import _quote
from .outbox import ensure_sqlite_outbox
from .spec import PRIMARY_KEYS, TABLES


ROLLBACK_CONFIRMATION = "REPLAY-POSTGRESQL-CHANGES-TO-SQLITE"


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _sqlite_upsert_sql(table: str, columns: list[str]) -> str:
    keys = PRIMARY_KEYS[table]
    mutable = [column for column in columns if column not in keys]
    action = "DO NOTHING"
    if mutable:
        action = "DO UPDATE SET " + ",".join(
            f'{_quote(column)}=excluded.{_quote(column)}' for column in mutable
        )
    return (
        f'INSERT INTO {_quote(table)} ({",".join(_quote(column) for column in columns)}) '
        f'VALUES ({",".join("?" for _ in columns)}) '
        f'ON CONFLICT ({",".join(_quote(key) for key in keys)}) {action}'
    )


def apply_changes_to_sqlite(connection: sqlite3.Connection, changes: list[dict[str, Any]]) -> int:
    connection.execute(
        "UPDATE automation_runtime_flags SET value='1' WHERE key='suppress_outbox'"
    )
    applied = 0
    try:
        for change in changes:
            table = str(change["source_table"])
            if table not in TABLES:
                raise ValueError(f"change table is outside automation scope: {table}")
            operation = str(change["operation"])
            primary_key = change["pk_json"]
            if isinstance(primary_key, str):
                primary_key = json.loads(primary_key)
            if set(primary_key) != set(PRIMARY_KEYS[table]):
                raise ValueError(f"invalid primary key for {table}")
            if operation == "delete":
                where = " AND ".join(f'{_quote(key)}=?' for key in PRIMARY_KEYS[table])
                connection.execute(
                    f'DELETE FROM {_quote(table)} WHERE {where}',
                    tuple(primary_key[key] for key in PRIMARY_KEYS[table]),
                )
            elif operation in {"insert", "update"}:
                row = change["row_json"]
                if isinstance(row, str):
                    row = json.loads(row)
                actual_columns = {item["name"] for item in connection.execute(f"PRAGMA table_info({_quote(table)})")}
                if not isinstance(row, dict) or set(row) != actual_columns:
                    raise ValueError(f"row schema mismatch for {table}")
                columns = [item["name"] for item in connection.execute(f"PRAGMA table_info({_quote(table)})")]
                connection.execute(_sqlite_upsert_sql(table, columns), tuple(row[column] for column in columns))
            else:
                raise ValueError(f"unsupported change operation: {operation}")
            applied += 1
    finally:
        connection.execute(
            "UPDATE automation_runtime_flags SET value='0' WHERE key='suppress_outbox'"
        )
    return applied


def replay_change_log(
    sqlite_path: Path,
    database_url: str,
    backup_path: Path,
    backup_sha256: str,
    confirmation: str,
    *,
    batch_size: int = 100,
) -> dict[str, int]:
    if confirmation != ROLLBACK_CONFIRMATION:
        raise RuntimeError("rollback confirmation does not match")
    if not backup_path.is_file() or _file_sha256(backup_path) != backup_sha256:
        raise RuntimeError("rollback backup is missing or changed")
    if batch_size <= 0 or batch_size > 1000:
        raise ValueError("batch_size must be between 1 and 1000")
    with psycopg.connect(database_url, row_factory=dict_row) as target:
        changes = [dict(row) for row in target.execute(
            "SELECT id,source_table,operation,pk_json,row_json FROM automation_change_log "
            "WHERE replayed_at IS NULL ORDER BY id LIMIT %s", (batch_size,)
        ).fetchall()]
        if not changes:
            with target.transaction():
                target.execute(
                    "UPDATE automation_runtime_flags SET value='false',updated_at=CURRENT_TIMESTAMP "
                    "WHERE key='capture_changes'"
                )
            return {"selected": 0, "applied": 0, "remaining": 0}
        with closing(sqlite3.connect(sqlite_path)) as destination:
            destination.row_factory = sqlite3.Row
            ensure_sqlite_outbox(destination)
            destination.commit()
            destination.execute("BEGIN IMMEDIATE")
            try:
                applied = apply_changes_to_sqlite(destination, changes)
                destination.commit()
            except Exception:
                destination.rollback()
                raise
        with target.transaction():
            target.execute(
                "UPDATE automation_change_log SET replayed_at=CURRENT_TIMESTAMP WHERE id=ANY(%s)",
                ([change["id"] for change in changes],),
            )
            remaining = target.execute(
                "SELECT COUNT(*) total FROM automation_change_log WHERE replayed_at IS NULL"
            ).fetchone()["total"]
            if remaining == 0:
                target.execute(
                    "UPDATE automation_runtime_flags SET value='false',updated_at=CURRENT_TIMESTAMP "
                    "WHERE key='capture_changes'"
                )
        return {"selected": len(changes), "applied": applied, "remaining": int(remaining)}
