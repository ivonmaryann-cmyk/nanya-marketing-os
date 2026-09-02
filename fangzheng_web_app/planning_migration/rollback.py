from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .spec import PRIMARY_KEYS, TABLE_COLUMNS


ENABLE_CONFIRMATION = "ENABLE-PLANNING-CHANGE-CAPTURE"
ROLLBACK_CONFIRMATION = "REPLAY-PLANNING-CHANGES-TO-SQLITE"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _verify_backup(path: Path, expected_hash: str) -> None:
    if not path.is_file() or file_sha256(path) != expected_hash:
        raise RuntimeError("planning SQLite backup is missing or changed")


def enable_change_capture(database_url: str, backup_path: Path, backup_sha256: str, confirmation: str) -> None:
    if confirmation != ENABLE_CONFIRMATION:
        raise RuntimeError("planning capture confirmation does not match")
    _verify_backup(backup_path, backup_sha256)
    with psycopg.connect(database_url, row_factory=dict_row) as target, target.transaction():
        cursor = target.execute(
            "UPDATE planning_runtime_flags SET value='true',updated_at=CURRENT_TIMESTAMP "
            "WHERE key='capture_changes' AND value='false'"
        )
        if cursor.rowcount != 1:
            raise RuntimeError("planning change capture was already enabled or is unavailable")


def apply_changes_to_sqlite(connection: sqlite3.Connection, changes: list[dict[str, Any]]) -> int:
    applied = 0
    for change in changes:
        table = str(change["table_name"])
        if table not in TABLE_COLUMNS:
            raise ValueError("planning change table is outside the migration scope")
        operation = str(change["operation"])
        pk = change["pk_json"]
        row = change["row_json"]
        if isinstance(pk, str):
            pk = json.loads(pk)
        if isinstance(row, str):
            row = json.loads(row)
        keys = PRIMARY_KEYS[table]
        if not isinstance(pk, dict) or set(pk) != set(keys):
            raise ValueError("planning change primary key schema mismatch")
        where = " AND ".join(f'"{key}"=?' for key in keys)
        if operation == "delete":
            connection.execute(f'DELETE FROM "{table}" WHERE {where}', tuple(pk[key] for key in keys))
        elif operation in {"insert", "update"}:
            columns = TABLE_COLUMNS[table]
            if not isinstance(row, dict) or set(row) != set(columns):
                raise ValueError("planning change row schema mismatch")
            mutable = [column for column in columns if column not in keys]
            updates = ",".join(f'"{column}"=excluded."{column}"' for column in mutable)
            connection.execute(
                f'INSERT INTO "{table}"({",".join(columns)}) VALUES ({",".join("?" for _ in columns)}) '
                f'ON CONFLICT({",".join(keys)}) DO UPDATE SET {updates}',
                tuple(row[column] for column in columns),
            )
        else:
            raise ValueError(f"unsupported planning change operation: {operation}")
        applied += 1
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
        raise RuntimeError("planning rollback confirmation does not match")
    _verify_backup(backup_path, backup_sha256)
    if batch_size <= 0 or batch_size > 1000:
        raise ValueError("batch_size must be between 1 and 1000")
    with psycopg.connect(database_url, row_factory=dict_row) as target:
        changes = [dict(row) for row in target.execute(
            "SELECT id,table_name,operation,pk_json,row_json FROM planning_change_log "
            "WHERE replayed_at IS NULL ORDER BY id LIMIT %s", (batch_size,)
        ).fetchall()]
        if not changes:
            return {"selected": 0, "applied": 0, "remaining": 0}
        with closing(sqlite3.connect(sqlite_path)) as destination:
            destination.execute("PRAGMA foreign_keys=ON")
            destination.execute("BEGIN IMMEDIATE")
            try:
                applied = apply_changes_to_sqlite(destination, changes)
                destination.commit()
            except Exception:
                destination.rollback()
                raise
        with target.transaction():
            target.execute(
                "UPDATE planning_change_log SET replayed_at=CURRENT_TIMESTAMP WHERE id=ANY(%s)",
                ([change["id"] for change in changes],),
            )
            remaining = int(target.execute(
                "SELECT COUNT(*) AS total FROM planning_change_log WHERE replayed_at IS NULL"
            ).fetchone()["total"])
            if remaining == 0:
                target.execute(
                    "UPDATE planning_runtime_flags SET value='false',updated_at=CURRENT_TIMESTAMP "
                    "WHERE key='capture_changes'"
                )
        return {"selected": len(changes), "applied": applied, "remaining": remaining}
