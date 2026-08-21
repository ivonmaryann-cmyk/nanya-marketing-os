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
        source_rows_by_table: dict[str, list[dict[str, Any]]] = {}
        target_rows_by_table: dict[str, list[dict[str, Any]]] = {}
        for table in TABLES:
            source_rows = [dict(row) for row in source.execute(f'SELECT * FROM "{table}"')]
            target_rows = [dict(row) for row in target.execute(f'SELECT * FROM "{table}"').fetchall()]
            source_rows_by_table[table] = source_rows
            target_rows_by_table[table] = target_rows
            keys = PRIMARY_KEYS[table]
            source_map = {tuple(row[key] for key in keys): _digest(row) for row in source_rows}
            target_map = {tuple(row[key] for key in keys): _digest(row) for row in target_rows}
            result = {
                "source_count": len(source_rows), "target_count": len(target_rows),
                "primary_keys_match": source_map.keys() == target_map.keys(),
                "row_hashes_match": source_map == target_map,
            }
            result["ok"] = all((
                result["source_count"] == result["target_count"],
                result["primary_keys_match"], result["row_hashes_match"],
            ))
            report["tables"][table] = result
            report["ok"] = report["ok"] and result["ok"]
        source_settings = {row["key"]: row["value"] for row in source_rows_by_table["settings"]}
        target_settings = {row["key"]: row["value"] for row in target_rows_by_table["settings"]}
        source_ciphertexts = {row["id"]: row["api_key_ciphertext"] for row in source_rows_by_table["pdf_excel_ai_config_versions"]}
        target_ciphertexts = {row["id"]: row["api_key_ciphertext"] for row in target_rows_by_table["pdf_excel_ai_config_versions"]}
        report["admin_password_hash_match"] = source_settings.get("admin_password_hash") == target_settings.get("admin_password_hash")
        report["active_ai_version_match"] = source_settings.get("active_pdf_excel_ai_config_version") == target_settings.get("active_pdf_excel_ai_config_version")
        report["ai_ciphertexts_match"] = source_ciphertexts == target_ciphertexts
        report["sqlite_integrity"] = source.execute("PRAGMA integrity_check").fetchone()[0]
        report["ok"] = report["ok"] and all((
            report["admin_password_hash_match"], report["active_ai_version_match"],
            report["ai_ciphertexts_match"], report["sqlite_integrity"] == "ok",
        ))
    return report
