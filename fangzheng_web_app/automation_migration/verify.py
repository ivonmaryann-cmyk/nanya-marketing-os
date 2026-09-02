from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

from .spec import ATTACHMENT_COLUMNS, PRIMARY_KEYS, TABLES


def _normalized_hash(row: dict[str, Any], table: str) -> str:
    payload = json.dumps(row, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _source_rows(source: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    return [dict(row) for row in source.execute(f'SELECT * FROM "{table}"')]


def _target_rows(target: Any, table: str) -> list[dict[str, Any]]:
    return [dict(row) for row in target.execute(f'SELECT * FROM "{table}"').fetchall()]


def verify_snapshot(snapshot: Path, target: Any, *, check_files: bool = True) -> dict[str, Any]:
    report: dict[str, Any] = {"ok": True, "tables": {}, "attachments": {"missing": [], "mismatched": []}}
    with closing(sqlite3.connect(snapshot)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only=ON")
        for table in TABLES:
            source_rows = _source_rows(source, table)
            target_rows = _target_rows(target, table)
            keys = PRIMARY_KEYS[table]
            source_map = {tuple(row[key] for key in keys): _normalized_hash(row, table) for row in source_rows}
            target_map = {tuple(row[key] for key in keys): _normalized_hash(row, table) for row in target_rows}
            result = {
                "source_count": len(source_rows), "target_count": len(target_rows),
                "primary_keys_match": source_map.keys() == target_map.keys(),
                "row_hashes_match": source_map == target_map,
            }
            result["ok"] = all((result["source_count"] == result["target_count"], result["primary_keys_match"], result["row_hashes_match"]))
            report["tables"][table] = result
            report["ok"] = report["ok"] and result["ok"]

        integrity = source.execute("PRAGMA integrity_check").fetchone()[0]
        report["sqlite_integrity"] = integrity
        report["ok"] = report["ok"] and integrity == "ok"

        if check_files:
            for table, fields in ATTACHMENT_COLUMNS.items():
                for row in _source_rows(source, table):
                    path_value = row.get(fields[0])
                    if not path_value:
                        continue
                    path = Path(path_value)
                    if not path.is_file():
                        report["attachments"]["missing"].append({"table": table, "id": row.get("id")})
                        continue
                    if table == "mail_attachments":
                        if row.get("size_bytes") not in (None, 0) and path.stat().st_size != row["size_bytes"]:
                            report["attachments"]["mismatched"].append({"table": table, "id": row.get("id"), "field": "size_bytes"})
                        if row.get("sha256"):
                            digest = hashlib.sha256(path.read_bytes()).hexdigest()
                            if digest != row["sha256"]:
                                report["attachments"]["mismatched"].append({"table": table, "id": row.get("id"), "field": "sha256"})
            report["ok"] = report["ok"] and not report["attachments"]["missing"] and not report["attachments"]["mismatched"]
    return report
