from __future__ import annotations

import hashlib
import json
import sqlite3
import time
import uuid
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .outbox import ensure_sqlite_outbox


METRICS = {
    "mail_date_counts": "SELECT substr(received_at,1,10) AS date_key,COUNT(*) total FROM mail_messages GROUP BY date_key ORDER BY date_key",
    "mail_category_counts": "SELECT is_order,COUNT(*) total FROM mail_messages GROUP BY is_order ORDER BY is_order",
    "mail_list_order": "SELECT id,received_at FROM mail_messages ORDER BY received_at DESC,id DESC",
    "case_workflow_counts": "SELECT status,action_type,workflow_stage,COUNT(*) total FROM order_intake_cases GROUP BY status,action_type,workflow_stage ORDER BY status,action_type,workflow_stage",
    "routing_rules": "SELECT id,employee_id,name,enabled,priority,action_type,updated_at FROM order_mail_routing_rules ORDER BY employee_id,priority,id",
    "rule_keywords": "SELECT group_id,scope,keyword FROM order_mail_rule_keywords ORDER BY group_id,id",
    "template_headers": "SELECT id,case_id,employee_id,template_key,header_json,current_version,updated_at FROM order_entry_templates ORDER BY id",
    "template_lines": "SELECT template_id,line_no,values_json,sources_json,updated_at FROM order_entry_template_lines ORDER BY template_id,line_no",
    "template_versions": "SELECT template_id,version_number,header_json,lines_json,saved_by,saved_at FROM order_entry_template_versions ORDER BY template_id,version_number",
    "attachment_metadata": "SELECT id,mail_id,filename,content_type,size_bytes,sha256,is_inline,parse_status FROM mail_attachments ORDER BY id",
}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def _canonical(rows: list[dict[str, Any]]) -> str:
    payload = json.dumps(rows, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def run_shadow_comparison(sqlite_path: Path, database_url: str) -> dict[str, Any]:
    run_id = uuid.uuid4().hex
    results: list[dict[str, Any]] = []
    with closing(sqlite3.connect(sqlite_path)) as source, psycopg.connect(database_url, row_factory=dict_row) as target:
        source.row_factory = sqlite3.Row
        ensure_sqlite_outbox(source)
        source.commit()
        for name, sql in METRICS.items():
            started = time.perf_counter()
            sqlite_rows = [dict(row) for row in source.execute(sql).fetchall()]
            postgres_rows = [dict(row) for row in target.execute(sql).fetchall()]
            elapsed_ms = round((time.perf_counter() - started) * 1000)
            sqlite_hash, postgres_hash = _canonical(sqlite_rows), _canonical(postgres_rows)
            item = {
                "metric_name": name, "sqlite_count": len(sqlite_rows), "postgresql_count": len(postgres_rows),
                "sqlite_hash": sqlite_hash, "postgresql_hash": postgres_hash,
                "elapsed_ms": elapsed_ms, "is_match": sqlite_hash == postgres_hash,
            }
            results.append(item)
            source.execute(
                """INSERT INTO automation_shadow_differences
                   (run_id,metric_name,sqlite_count,postgresql_count,sqlite_hash,postgresql_hash,elapsed_ms,is_match,observed_at)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (run_id, name, len(sqlite_rows), len(postgres_rows), sqlite_hash, postgres_hash, elapsed_ms, int(item["is_match"]), _now()),
            )
        source.commit()
        target.execute(
            "INSERT INTO automation_shadow_runs(run_id,finished_at,status,metric_count,difference_count) "
            "VALUES (%s,CURRENT_TIMESTAMP,%s,%s,%s) ON CONFLICT(run_id) DO NOTHING",
            (run_id, "matched" if all(item["is_match"] for item in results) else "different", len(results), sum(not item["is_match"] for item in results)),
        )
        target.commit()
    return {
        "run_id": run_id, "metric_count": len(results),
        "difference_count": sum(not item["is_match"] for item in results), "metrics": results,
    }


def shadow_summary(sqlite_path: Path, *, days: int = 7) -> dict[str, Any]:
    if days <= 0:
        raise ValueError("days must be positive")
    with closing(sqlite3.connect(sqlite_path)) as source:
        source.row_factory = sqlite3.Row
        ensure_sqlite_outbox(source)
        source.commit()
        row = source.execute(
            """SELECT COUNT(DISTINCT substr(observed_at,1,10)) observed_days, COUNT(DISTINCT run_id) runs,
                      COALESCE(SUM(CASE WHEN is_match=0 THEN 1 ELSE 0 END),0) differences,
                      COALESCE(MAX(elapsed_ms),0) max_elapsed_ms
               FROM automation_shadow_differences
               WHERE julianday(observed_at)>=julianday('now', ?)""",
            (f"-{days} days",),
        ).fetchone()
        return dict(row)
