from __future__ import annotations

import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any


USER_COLUMNS = (
    "employee_id", "display_name", "department", "password_hash", "role", "enabled",
    "must_change_password", "created_at", "updated_at",
)


def copy_users(snapshot: Path, target: Any) -> int:
    with closing(sqlite3.connect(snapshot)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only=ON")
        actual = tuple(row["name"] for row in source.execute("PRAGMA table_info(users)"))
        if actual != USER_COLUMNS:
            raise RuntimeError("SQLite users schema does not match the audited migration schema")
        rows = source.execute("SELECT * FROM users ORDER BY employee_id").fetchall()
    if rows:
        columns = ",".join(USER_COLUMNS)
        placeholders = ",".join("%s" for _ in USER_COLUMNS)
        mutable = [column for column in USER_COLUMNS if column != "employee_id"]
        updates = ",".join(f"{column}=EXCLUDED.{column}" for column in mutable)
        with target.cursor() as cursor:
            cursor.executemany(
                f"INSERT INTO users({columns}) VALUES ({placeholders}) "
                f"ON CONFLICT(employee_id) DO UPDATE SET {updates}",
                [tuple(row[column] for column in USER_COLUMNS) for row in rows],
            )
    return len(rows)
