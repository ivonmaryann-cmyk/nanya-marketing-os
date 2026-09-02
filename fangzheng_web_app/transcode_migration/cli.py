from __future__ import annotations

import argparse
import json
import logging
import os
from pathlib import Path

import psycopg
from psycopg.rows import dict_row

from ..automation_migration.snapshot import sqlite_snapshot
from ..paths import STORAGE_DIR
from .copy import copy_snapshot
from .rollback import ENABLE_CONFIRMATION, ROLLBACK_CONFIRMATION, enable_change_capture, replay_change_log
from .schema import apply_migrations
from .verify import verify_snapshot


LOG = logging.getLogger("transcode_migration")


def _configure_logging() -> None:
    path = Path(os.getenv("TRANSCODE_MIGRATION_LOG_PATH", "").strip() or STORAGE_DIR / "migration_logs" / "transcode_migration.log")
    path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler(), logging.FileHandler(path, encoding="utf-8")],
        force=True,
    )


def _url(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"{name} is required")
    return value


def migrate(sqlite_path: Path, database_url: str) -> dict:
    with sqlite_snapshot(sqlite_path) as snapshot, psycopg.connect(database_url, row_factory=dict_row) as target:
        with target.transaction():
            migrations = apply_migrations(target)
            copied = copy_snapshot(snapshot, target)
            verification = verify_snapshot(snapshot, target)
            if not verification["ok"]:
                raise RuntimeError("transcode verification failed; PostgreSQL transaction rolled back")
    LOG.info("transcode migration completed copied_counts=%s verification_ok=true", copied)
    return {"migrations": migrations, "copied": copied, "verification": verification}


def rollback_test_target(database_url: str, confirmation: str) -> None:
    if confirmation != "DROP-TRANSCODE-TEST-DATA":
        raise RuntimeError("test rollback confirmation does not match")
    with psycopg.connect(database_url, row_factory=dict_row) as target, target.transaction():
        rows = target.execute(
            "SELECT table_name FROM information_schema.tables WHERE table_schema='public' "
            "AND (table_name IN ('jobs','pp_transcode_base_rules','pp_transcode_customer_rules',"
            "'pp_transcode_rule_changes','pp_transcode_confirmation_items','transcode_model_configs',"
            "'transcode_agent_confirmation_items','transcode_agent_confirmation_events',"
            "'transcode_agent_pending_rules','transcode_agent_row_verifications','transcode_agent_rule_overrides',"
            "'transcode_customer_rule_changes','transcode_customer_rule_overrides',"
            "'transcode_rule_center_asset_overrides','transcode_rule_center_base_overrides',"
            "'transcode_rule_center_changes','transcode_rule_center_confirmation_overrides',"
            "'transcode_rule_center_lookup_overrides') OR table_name LIKE 'transcode_%')"
        ).fetchall()
        for row in rows:
            target.execute(f'DROP TABLE IF EXISTS "{row["table_name"]}" CASCADE')
        target.execute("DROP FUNCTION IF EXISTS transcode_capture_change() CASCADE")


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrate the 18 scoped jobs/transcode tables")
    subparsers = parser.add_subparsers(dest="command", required=True)
    migrate_parser = subparsers.add_parser("migrate")
    migrate_parser.add_argument("--sqlite", type=Path, required=True)
    migrate_parser.add_argument("--test-target", action="store_true")
    verify_parser = subparsers.add_parser("verify")
    verify_parser.add_argument("--sqlite", type=Path, required=True)
    capture_parser = subparsers.add_parser("enable-change-capture")
    capture_parser.add_argument("--backup", type=Path, required=True)
    capture_parser.add_argument("--backup-sha256", required=True)
    capture_parser.add_argument("--confirm", required=True)
    replay_parser = subparsers.add_parser("rollback-replay")
    replay_parser.add_argument("--sqlite", type=Path, required=True)
    replay_parser.add_argument("--backup", type=Path, required=True)
    replay_parser.add_argument("--backup-sha256", required=True)
    replay_parser.add_argument("--confirm", required=True)
    replay_parser.add_argument("--batch-size", type=int, default=100)
    drop_parser = subparsers.add_parser("rollback-test-target")
    drop_parser.add_argument("--confirm", required=True)
    args = parser.parse_args()
    _configure_logging()
    try:
        if args.command == "migrate":
            name = "TRANSCODE_TEST_DATABASE_URL" if args.test_target else "TRANSCODE_DATABASE_URL"
            result = migrate(args.sqlite, _url(name))
        elif args.command == "verify":
            with sqlite_snapshot(args.sqlite) as snapshot, psycopg.connect(
                _url("TRANSCODE_DATABASE_URL"), row_factory=dict_row
            ) as target:
                result = verify_snapshot(snapshot, target)
        elif args.command == "enable-change-capture":
            enable_change_capture(_url("TRANSCODE_DATABASE_URL"), args.backup, args.backup_sha256, args.confirm)
            result = {"ok": True, "confirmation": ENABLE_CONFIRMATION}
        elif args.command == "rollback-replay":
            result = replay_change_log(
                args.sqlite, _url("TRANSCODE_DATABASE_URL"), args.backup,
                args.backup_sha256, args.confirm, batch_size=args.batch_size,
            )
        else:
            rollback_test_target(_url("TRANSCODE_TEST_DATABASE_URL"), args.confirm)
            result = {"ok": True, "scope": "transcode test target"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("ok", result.get("verification", {}).get("ok", True)) else 2
    except Exception as exc:
        LOG.error("transcode migration command failed (%s); sensitive details were omitted", type(exc).__name__)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
