from __future__ import annotations

import argparse
import json
import logging
import os
import tempfile
from pathlib import Path

import psycopg
from psycopg.rows import dict_row

from .copy import copy_snapshot
from .audit import audit_snapshot
from .schema import apply_migrations
from .snapshot import sqlite_snapshot
from .spec import TABLES
from .verify import verify_snapshot


LOG = logging.getLogger("automation_migration")


def _connect(url: str):
    return psycopg.connect(url, row_factory=dict_row)


def _database_url(test: bool = False) -> str:
    name = "AUTOMATION_TEST_DATABASE_URL" if test else "AUTOMATION_DATABASE_URL"
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"{name} is required")
    return value


def migrate(sqlite_path: Path, *, test_target: bool = False, check_files: bool = True) -> dict:
    with sqlite_snapshot(sqlite_path) as snapshot, _connect(_database_url(test_target)) as target:
        with target.transaction():
            migrations = apply_migrations(target)
            copied = copy_snapshot(snapshot, target)
            verification = verify_snapshot(snapshot, target, check_files=check_files)
            if not verification["ok"]:
                raise RuntimeError("verification failed; PostgreSQL transaction rolled back")
        return {"migrations": migrations, "copied": copied, "verification": verification}


def rollback_test_target(confirm: str) -> None:
    if confirm != "DROP-AUTOMATION-TEST-DATA":
        raise RuntimeError("test rollback requires --confirm DROP-AUTOMATION-TEST-DATA")
    with _connect(_database_url(test=True)) as target, target.transaction():
        for table in reversed(TABLES):
            target.execute(f'DROP TABLE IF EXISTS "{table}" CASCADE')
        target.execute("DROP TABLE IF EXISTS automation_metadata CASCADE")
        target.execute("DROP TABLE IF EXISTS automation_schema_migrations CASCADE")


def _write_reports(result: dict, report_path: Path) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    handle, temporary_name = tempfile.mkstemp(prefix=f".{report_path.name}.", dir=report_path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as stream:
            stream.write(payload)
        os.replace(temporary_name, report_path)
    except Exception:
        Path(temporary_name).unlink(missing_ok=True)
        raise
    markdown_path = report_path.with_suffix(".md")
    lines = ["# Order-mail automation migration report", "", f"- Result: {'passed' if result.get('ok', result.get('verification', {}).get('ok')) else 'failed'}", ""]
    if "copied" in result:
        lines.extend(["## Copied rows", "", *[f"- `{table}`: {count}" for table, count in result["copied"].items()]])
    elif "tables" in result:
        lines.extend(["## SQLite rows", "", *[f"- `{table}`: {count}" for table, count in result["tables"].items()]])
    markdown_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrate only order-mail automation tables")
    subparsers = parser.add_subparsers(dest="command", required=True)
    migrate_parser = subparsers.add_parser("migrate")
    migrate_parser.add_argument("--sqlite", type=Path, required=True)
    migrate_parser.add_argument("--test-target", action="store_true")
    migrate_parser.add_argument("--skip-file-check", action="store_true")
    migrate_parser.add_argument("--report", type=Path)
    audit_parser = subparsers.add_parser("audit-sqlite")
    audit_parser.add_argument("--sqlite", type=Path, required=True)
    audit_parser.add_argument("--skip-file-check", action="store_true")
    audit_parser.add_argument("--report", type=Path)
    rollback_parser = subparsers.add_parser("rollback-test-target")
    rollback_parser.add_argument("--confirm", required=True)
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    try:
        if args.command == "migrate":
            result = migrate(args.sqlite, test_target=args.test_target, check_files=not args.skip_file_check)
            if args.report:
                _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
        elif args.command == "audit-sqlite":
            with sqlite_snapshot(args.sqlite) as snapshot:
                result = audit_snapshot(snapshot, check_files=not args.skip_file_check)
            if args.report:
                _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0 if result["ok"] else 2
        else:
            rollback_test_target(args.confirm)
            print(json.dumps({"ok": True, "scope": "automation test target"}))
    except Exception as exc:
        error = {"ok": False, "error_type": type(exc).__name__}
        report_path = getattr(args, "report", None)
        if report_path:
            _write_reports(error, report_path)
        LOG.error("migration command failed (%s); sensitive details were omitted", type(exc).__name__)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
