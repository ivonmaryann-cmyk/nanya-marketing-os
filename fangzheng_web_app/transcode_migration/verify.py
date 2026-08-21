from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

from .spec import PRIMARY_KEYS, TABLES


def _digest(row: dict[str, Any]) -> str:
    payload = json.dumps(row, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def verify_snapshot(snapshot: Path, target: Any) -> dict[str, Any]:
    report: dict[str, Any] = {"ok": True, "tables": {}}
    with closing(sqlite3.connect(snapshot)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only=ON")
        for table in TABLES:
            source_rows = [dict(row) for row in source.execute(f'SELECT * FROM "{table}"')]
            target_rows = [dict(row) for row in target.execute(f'SELECT * FROM "{table}"').fetchall()]
            keys = PRIMARY_KEYS[table]
            source_map = {tuple(row[key] for key in keys): _digest(row) for row in source_rows}
            target_map = {tuple(row[key] for key in keys): _digest(row) for row in target_rows}
            result = {
                "source_count": len(source_rows),
                "target_count": len(target_rows),
                "primary_keys_match": source_map.keys() == target_map.keys(),
                "row_hashes_match": source_map == target_map,
            }
            result["ok"] = all((
                result["source_count"] == result["target_count"],
                result["primary_keys_match"],
                result["row_hashes_match"],
            ))
            report["tables"][table] = result
            report["ok"] = report["ok"] and result["ok"]
        report["sqlite_integrity"] = source.execute("PRAGMA integrity_check").fetchone()[0]
        report["ok"] = report["ok"] and report["sqlite_integrity"] == "ok"
    return report
