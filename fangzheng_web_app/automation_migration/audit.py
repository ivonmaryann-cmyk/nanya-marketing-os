from __future__ import annotations

import hashlib
import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

from .spec import TABLES
from ..mail_transcode_agent.mail_crypto import decrypt_text


def audit_snapshot(snapshot: Path, *, check_files: bool = True) -> dict[str, Any]:
    report: dict[str, Any] = {
        "ok": True, "integrity": "", "tables": {}, "duplicates": {},
        "orphans": {}, "mail_credentials": {}, "files": {},
    }
    with closing(sqlite3.connect(snapshot)) as connection:
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA query_only=ON")
        report["integrity"] = connection.execute("PRAGMA integrity_check").fetchone()[0]
        report["ok"] = report["integrity"] == "ok"
        for table in TABLES:
            report["tables"][table] = connection.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]

        unique_keys = (
            ("mail_messages", ("account_id", "folder", "uid")),
            ("order_intake_cases", ("employee_id", "mail_id")),
            ("order_change_tags", ("employee_id", "name")),
            ("order_entry_templates", ("case_id",)),
            ("order_entry_template_lines", ("template_id", "line_no")),
            ("order_entry_template_versions", ("template_id", "version_number")),
        )
        for table, columns in unique_keys:
            grouped = ", ".join(f'"{column}"' for column in columns)
            count = connection.execute(
                f'SELECT COUNT(*) FROM (SELECT 1 FROM "{table}" GROUP BY {grouped} HAVING COUNT(*)>1)'
            ).fetchone()[0]
            report["duplicates"][f"{table}({','.join(columns)})"] = count
            report["ok"] = report["ok"] and count == 0

        relations = (
            ("mail_attachments", "mail_id", "mail_messages", "id"),
            ("mail_attachment_texts", "attachment_id", "mail_attachments", "id"),
            ("mail_fetch_task_messages", "fetch_task_id", "mail_fetch_tasks", "id"),
            ("mail_fetch_task_messages", "mail_id", "mail_messages", "id"),
            ("mail_order_tasks", "mail_id", "mail_messages", "id"),
            ("order_intake_cases", "mail_id", "mail_messages", "id"),
            ("order_intake_case_events", "case_id", "order_intake_cases", "id"),
            ("order_mail_routing_rule_events", "rule_id", "order_mail_routing_rules", "id"),
            ("order_mail_rule_keywords", "group_id", "order_mail_rule_groups", "id"),
            ("order_change_tag_keywords", "tag_id", "order_change_tags", "id"),
            ("order_entry_templates", "case_id", "order_intake_cases", "id"),
            ("order_entry_template_lines", "template_id", "order_entry_templates", "id"),
            ("order_entry_template_versions", "template_id", "order_entry_templates", "id"),
        )
        for child, child_key, parent, parent_key in relations:
            name = f"{child}.{child_key}->{parent}.{parent_key}"
            count = connection.execute(
                f'SELECT COUNT(*) FROM "{child}" c LEFT JOIN "{parent}" p ON c."{child_key}"=p."{parent_key}" '
                f'WHERE p."{parent_key}" IS NULL'
            ).fetchone()[0]
            report["orphans"][name] = count
            report["ok"] = report["ok"] and count == 0

        credentials = connection.execute(
            "SELECT id, auth_code_ciphertext FROM mail_accounts"
        ).fetchall()
        empty = sum(not row["auth_code_ciphertext"] for row in credentials)
        decrypt_failed = sum(
            bool(row["auth_code_ciphertext"]) and not decrypt_text(row["auth_code_ciphertext"])
            for row in credentials
        )
        report["mail_credentials"] = {
            "checked": len(credentials), "empty_ciphertext": empty, "decrypt_failed": decrypt_failed,
        }
        report["ok"] = report["ok"] and not empty and not decrypt_failed

        if check_files:
            checked = missing = size_mismatch = hash_mismatch = 0
            rows = connection.execute(
                "SELECT id, stored_path, size_bytes, sha256 FROM mail_attachments WHERE TRIM(stored_path)<>''"
            ).fetchall()
            for row in rows:
                checked += 1
                path = Path(row["stored_path"])
                if not path.is_file():
                    missing += 1
                    continue
                if row["size_bytes"] and path.stat().st_size != row["size_bytes"]:
                    size_mismatch += 1
                if row["sha256"]:
                    digest = hashlib.sha256(path.read_bytes()).hexdigest()
                    if digest != row["sha256"]:
                        hash_mismatch += 1
            report["files"] = {
                "checked": checked, "missing": missing,
                "size_mismatch": size_mismatch, "hash_mismatch": hash_mismatch,
            }
            report["ok"] = report["ok"] and not any((missing, size_mismatch, hash_mismatch))
    return report
