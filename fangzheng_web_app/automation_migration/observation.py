from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .shadow import shadow_summary
from .sync import outbox_status


def collect_observation(sqlite_path: Path, database_url: str) -> dict[str, Any]:
    with psycopg.connect(database_url, row_factory=dict_row) as target:
        database = dict(target.execute(
            """SELECT numbackends,xact_commit,xact_rollback,deadlocks,temp_files,
                      pg_database_size(current_database()) database_size_bytes
               FROM pg_stat_database WHERE datname=current_database()"""
        ).fetchone())
        change_log = dict(target.execute(
            "SELECT COUNT(*) total,COUNT(*) FILTER (WHERE replayed_at IS NULL) unreplayed FROM automation_change_log"
        ).fetchone())
        business = dict(target.execute(
            """SELECT (SELECT COUNT(*) FROM mail_messages) mail_messages,
                      (SELECT COUNT(*) FROM order_intake_cases) order_cases,
                      (SELECT COUNT(*) FROM order_entry_templates) templates"""
        ).fetchone())
    return {
        "observed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "health_ok": False,
        "health_review_note": "manual review required",
        "outbox": outbox_status(sqlite_path),
        "shadow_7d": shadow_summary(sqlite_path, days=7),
        "postgresql": database,
        "change_log": change_log,
        "business_counts": business,
    }


def evaluate_observation_reports(paths: list[Path], *, minimum_days: int = 7) -> dict[str, Any]:
    if minimum_days < 7 or minimum_days > 14:
        raise ValueError("minimum_days must be between 7 and 14")
    reports = [json.loads(path.read_text(encoding="utf-8")) for path in paths]
    days = {str(report.get("observed_at", ""))[:10] for report in reports if report.get("observed_at")}
    blockers: list[str] = []
    if len(days) < minimum_days:
        blockers.append(f"observation_days>={minimum_days}")
    if any(report.get("health_ok") is not True for report in reports):
        blockers.append("all_health_reviews_passed")
    if any(int(report.get("outbox", {}).get("pending", -1)) != 0 for report in reports):
        blockers.append("outbox_pending=0")
    if any(int(report.get("shadow_7d", {}).get("differences", -1)) != 0 for report in reports):
        blockers.append("shadow_differences=0")
    return {"passed": not blockers, "observed_days": len(days), "blockers": blockers}
