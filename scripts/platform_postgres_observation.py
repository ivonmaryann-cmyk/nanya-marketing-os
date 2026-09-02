"""Write a read-only daily health record for the unified PostgreSQL cutover."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fangzheng_web_app.database.config import (  # noqa: E402
    AutomationDatabaseConfig,
    ConfigurationDatabaseConfig,
    IdentityDatabaseConfig,
    PlanningDatabaseConfig,
    TranscodeDatabaseConfig,
)
from fangzheng_web_app.local_env import load_local_env  # noqa: E402


CONFIG_TYPES = (
    AutomationDatabaseConfig,
    IdentityDatabaseConfig,
    TranscodeDatabaseConfig,
    ConfigurationDatabaseConfig,
    PlanningDatabaseConfig,
)
TABLES = (
    "users",
    "settings",
    "jobs",
    "automation_customers",
    "mail_messages",
    "mail_attachments",
    "order_intake_cases",
    "order_entry_templates",
    "order_interface_call_logs",
)


def _latest_file_inventory_summary() -> dict[str, Any]:
    reports = sorted((PROJECT_ROOT / "outputs" / "migrations").glob("platform-file-inventory-*.json"))
    if not reports:
        return {"available": False}
    try:
        payload = json.loads(reports[-1].read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"available": False}
    summary = payload.get("summary") if isinstance(payload, dict) else {}
    return {
        "available": isinstance(summary, dict),
        "failed": int((summary or {}).get("failed") or 0),
        "pending": int((summary or {}).get("pending") or 0),
        "total": int((summary or {}).get("total") or 0),
    }


def build_observation() -> dict[str, Any]:
    load_local_env()
    configs = [config_type.from_env() for config_type in CONFIG_TYPES]
    urls = {config.database_url for config in configs}
    if {config.backend for config in configs} != {"postgresql"} or len(urls) != 1:
        raise RuntimeError("统一 PostgreSQL 配置不一致，拒绝生成观察通过记录")

    import psycopg

    with psycopg.connect(configs[0].database_url or "", connect_timeout=8) as connection:
        with connection.cursor() as cursor:
            cursor.execute("SELECT current_database(), current_user, version()")
            database_name, database_user, server_version = cursor.fetchone()
            table_counts: dict[str, int] = {}
            for table in TABLES:
                cursor.execute(f"SELECT COUNT(*) FROM {table}")
                table_counts[table] = int(cursor.fetchone()[0])
            cursor.execute(
                """
                SELECT COUNT(*)
                FROM pg_constraint con
                JOIN pg_namespace schema ON schema.oid = con.connamespace
                WHERE schema.nspname = 'public'
                  AND con.contype = 'f'
                  AND NOT con.convalidated
                """
            )
            unvalidated_foreign_keys = int(cursor.fetchone()[0])
            cursor.execute("SELECT COUNT(*) FROM jobs WHERE status IN ('queued', 'running')")
            active_jobs = int(cursor.fetchone()[0])

    file_inventory = _latest_file_inventory_summary()
    ok = unvalidated_foreign_keys == 0 and int(file_inventory.get("failed") or 0) == 0
    return {
        "observed_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "ok": ok,
        "profile": {
            "module_count": len(configs),
            "backend": "postgresql",
            "database": str(database_name),
            "database_user": str(database_user),
            "server_version": " ".join(str(server_version).split()[:2]),
        },
        "table_counts": table_counts,
        "active_jobs": active_jobs,
        "unvalidated_foreign_keys": unvalidated_foreign_keys,
        "file_inventory": file_inventory,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Record one read-only platform PostgreSQL observation")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = build_observation()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"ok": report["ok"], "output": str(args.output)}, ensure_ascii=False))
    return 0 if report["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
