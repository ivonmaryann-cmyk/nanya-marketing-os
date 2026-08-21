from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

from .copy import USER_COLUMNS


def _digest(row: dict[str, Any]) -> str:
    payload = json.dumps(row, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def verify_users(snapshot: Path, target: Any) -> dict[str, Any]:
    with closing(sqlite3.connect(snapshot)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only=ON")
        source_rows = [dict(row) for row in source.execute("SELECT * FROM users ORDER BY employee_id")]
        integrity = source.execute("PRAGMA integrity_check").fetchone()[0]
    target_rows = [dict(row) for row in target.execute("SELECT * FROM users ORDER BY employee_id").fetchall()]
    source_map = {str(row["employee_id"]): _digest(row) for row in source_rows}
    target_map = {str(row["employee_id"]): _digest(row) for row in target_rows}
    source_passwords = {str(row["employee_id"]): row["password_hash"] for row in source_rows}
    target_passwords = {str(row["employee_id"]): row["password_hash"] for row in target_rows}

    def count(rows: list[dict[str, Any]], predicate) -> int:
        return sum(1 for row in rows if predicate(row))

    checks = {
        "source_count": len(source_rows),
        "target_count": len(target_rows),
        "primary_keys_match": source_map.keys() == target_map.keys(),
        "row_hashes_match": source_map == target_map,
        "password_hashes_match": source_passwords == target_passwords,
        "enabled_count_match": count(source_rows, lambda row: bool(row["enabled"])) == count(target_rows, lambda row: bool(row["enabled"])),
        "admin_count_match": count(source_rows, lambda row: row["role"] == "admin") == count(target_rows, lambda row: row["role"] == "admin"),
        "must_change_count_match": count(source_rows, lambda row: bool(row["must_change_password"])) == count(target_rows, lambda row: bool(row["must_change_password"])),
        "sqlite_integrity": integrity,
        "columns": list(USER_COLUMNS),
    }
    checks["ok"] = all((
        checks["source_count"] == checks["target_count"],
        checks["primary_keys_match"], checks["row_hashes_match"], checks["password_hashes_match"],
        checks["enabled_count_match"], checks["admin_count_match"],
        checks["must_change_count_match"], integrity == "ok",
    ))
    return checks
