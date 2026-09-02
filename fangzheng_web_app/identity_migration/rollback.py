from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .copy import USER_COLUMNS


ENABLE_CONFIRMATION = "ENABLE-IDENTITY-CHANGE-CAPTURE"
ROLLBACK_CONFIRMATION = "REPLAY-IDENTITY-CHANGES-TO-SQLITE"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def enable_change_capture(
    database_url: str, backup_path: Path, backup_sha256: str, confirmation: str
) -> None:
    if confirmation != ENABLE_CONFIRMATION:
        raise RuntimeError("identity capture confirmation does not match")
    if not backup_path.is_file() or file_sha256(backup_path) != backup_sha256:
        raise RuntimeError("identity SQLite backup is missing or changed")
    with psycopg.connect(database_url, row_factory=dict_row) as target, target.transaction():
        cursor = target.execute(
            "UPDATE identity_runtime_flags SET value='true',updated_at=CURRENT_TIMESTAMP "
            "WHERE key='capture_changes' AND value='false'"
        )
        if cursor.rowcount != 1:
            raise RuntimeError("identity change capture was already enabled or is unavailable")


def apply_changes_to_sqlite(connection: sqlite3.Connection, changes: list[dict[str, Any]]) -> int:
    applied = 0
    for change in changes:
        operation = str(change["operation"])
        employee_id = str(change["employee_id"])
        if operation == "delete":
            connection.execute("DELETE FROM users WHERE employee_id=?", (employee_id,))
        elif operation in {"insert", "update"}:
            row = change["row_json"]
            if isinstance(row, str):
                row = json.loads(row)
            if not isinstance(row, dict) or set(row) != set(USER_COLUMNS):
                raise ValueError("identity change row schema mismatch")
            mutable = [column for column in USER_COLUMNS if column != "employee_id"]
            updates = ",".join(f"{column}=excluded.{column}" for column in mutable)
            connection.execute(
                f"INSERT INTO users({','.join(USER_COLUMNS)}) VALUES ({','.join('?' for _ in USER_COLUMNS)}) "
                f"ON CONFLICT(employee_id) DO UPDATE SET {updates}",
                tuple(row[column] for column in USER_COLUMNS),
            )
        else:
            raise ValueError(f"unsupported identity change operation: {operation}")
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
        raise RuntimeError("identity rollback confirmation does not match")
    if not backup_path.is_file() or file_sha256(backup_path) != backup_sha256:
        raise RuntimeError("identity rollback backup is missing or changed")
    if batch_size <= 0 or batch_size > 1000:
        raise ValueError("batch_size must be between 1 and 1000")
    with psycopg.connect(database_url, row_factory=dict_row) as target:
        changes = [dict(row) for row in target.execute(
            "SELECT id,operation,employee_id,row_json FROM identity_change_log "
            "WHERE replayed_at IS NULL ORDER BY id LIMIT %s", (batch_size,)
        ).fetchall()]
        if not changes:
            return {"selected": 0, "applied": 0, "remaining": 0}
        with closing(sqlite3.connect(sqlite_path)) as destination:
            destination.execute("BEGIN IMMEDIATE")
            try:
                applied = apply_changes_to_sqlite(destination, changes)
                destination.commit()
            except Exception:
                destination.rollback()
                raise
        with target.transaction():
            target.execute(
                "UPDATE identity_change_log SET replayed_at=CURRENT_TIMESTAMP WHERE id=ANY(%s)",
                ([change["id"] for change in changes],),
            )
            remaining = int(target.execute(
                "SELECT COUNT(*) AS total FROM identity_change_log WHERE replayed_at IS NULL"
            ).fetchone()["total"])
            if remaining == 0:
                target.execute(
                    "UPDATE identity_runtime_flags SET value='false',updated_at=CURRENT_TIMESTAMP "
                    "WHERE key='capture_changes'"
                )
        return {"selected": len(changes), "applied": applied, "remaining": remaining}
