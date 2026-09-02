from __future__ import annotations

import argparse
import json
import logging
import os
from pathlib import Path

import psycopg
from psycopg.rows import dict_row

from ..automation_migration.snapshot import sqlite_snapshot
from .copy import copy_users
from .rollback import (
    ENABLE_CONFIRMATION,
    ROLLBACK_CONFIRMATION,
    enable_change_capture,
    replay_change_log,
)
from .schema import apply_migrations
from .verify import verify_users


LOG = logging.getLogger("identity_migration")


def _url(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"{name} is required")
    return value


def migrate(sqlite_path: Path, database_url: str) -> dict:
    with sqlite_snapshot(sqlite_path) as snapshot, psycopg.connect(database_url, row_factory=dict_row) as target:
        with target.transaction():
            migrations = apply_migrations(target)
            copied = copy_users(snapshot, target)
            verification = verify_users(snapshot, target)
            if not verification["ok"]:
                raise RuntimeError("identity verification failed; PostgreSQL transaction rolled back")
    return {"migrations": migrations, "copied": copied, "verification": verification}


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrate only the users/login table")
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
    rollback_parser = subparsers.add_parser("rollback-test-target")
    rollback_parser.add_argument("--confirm", required=True)
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    try:
        if args.command == "migrate":
            name = "IDENTITY_TEST_DATABASE_URL" if args.test_target else "IDENTITY_DATABASE_URL"
            result = migrate(args.sqlite, _url(name))
        elif args.command == "verify":
            with sqlite_snapshot(args.sqlite) as snapshot, psycopg.connect(
                _url("IDENTITY_DATABASE_URL"), row_factory=dict_row
            ) as target:
                result = verify_users(snapshot, target)
        elif args.command == "enable-change-capture":
            enable_change_capture(
                _url("IDENTITY_DATABASE_URL"), args.backup, args.backup_sha256, args.confirm
            )
            result = {"ok": True, "confirmation": ENABLE_CONFIRMATION}
        elif args.command == "rollback-replay":
            result = replay_change_log(
                args.sqlite, _url("IDENTITY_DATABASE_URL"), args.backup,
                args.backup_sha256, args.confirm, batch_size=args.batch_size,
            )
        else:
            if args.confirm != "DROP-IDENTITY-TEST-DATA":
                raise RuntimeError("test rollback confirmation does not match")
            with psycopg.connect(_url("IDENTITY_TEST_DATABASE_URL"), row_factory=dict_row) as target, target.transaction():
                target.execute("DROP TABLE IF EXISTS users CASCADE")
                target.execute("DROP TABLE IF EXISTS identity_change_log CASCADE")
                target.execute("DROP TABLE IF EXISTS identity_runtime_flags CASCADE")
                target.execute("DROP TABLE IF EXISTS identity_schema_migrations CASCADE")
                target.execute("DROP FUNCTION IF EXISTS identity_capture_user_change() CASCADE")
            result = {"ok": True, "scope": "identity test target"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("ok", result.get("verification", {}).get("ok", True)) else 2
    except Exception as exc:
        LOG.error("identity migration command failed (%s); sensitive details were omitted", type(exc).__name__)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
