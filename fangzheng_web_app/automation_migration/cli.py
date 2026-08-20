from __future__ import annotations

import argparse
import json
import logging
import os
import tempfile
import time
from pathlib import Path

import psycopg
from psycopg.rows import dict_row

from .copy import copy_snapshot
from .audit import audit_snapshot
from .cutover import enable_change_capture, prepare_cutover
from .observation import collect_observation, evaluate_observation_reports
from .schema import apply_migrations
from .performance import run_read_benchmark
from .preflight import evaluate_preflight
from .shadow import run_shadow_comparison, shadow_summary
from .snapshot import sqlite_snapshot
from .spec import TABLES
from .sync import outbox_status, process_outbox
from .rollback import replay_change_log
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


def _shadow_database_url() -> str:
    value = os.getenv("AUTOMATION_SHADOW_DATABASE_URL", "").strip()
    if not value:
        raise RuntimeError("AUTOMATION_SHADOW_DATABASE_URL is required")
    return value


def _required_url(name: str) -> str:
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
        target.execute("DROP TABLE IF EXISTS automation_shadow_runs CASCADE")
        target.execute("DROP TABLE IF EXISTS automation_migration_inbox CASCADE")
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
    passed = result.get("ok", result.get("passed", result.get("verification", {}).get("ok")))
    if passed is None and "difference_count" in result:
        passed = result["difference_count"] == 0
    lines = ["# Order-mail automation migration report", "", f"- Result: {'passed' if passed else 'failed'}", ""]
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
    status_parser = subparsers.add_parser("outbox-status")
    status_parser.add_argument("--sqlite", type=Path, required=True)
    sync_parser = subparsers.add_parser("sync-outbox")
    sync_parser.add_argument("--sqlite", type=Path, required=True)
    sync_parser.add_argument("--batch-size", type=int, default=100)
    sync_parser.add_argument("--watch", action="store_true")
    sync_parser.add_argument("--poll-seconds", type=int, default=5)
    shadow_parser = subparsers.add_parser("shadow-compare")
    shadow_parser.add_argument("--sqlite", type=Path, required=True)
    shadow_parser.add_argument("--report", type=Path)
    summary_parser = subparsers.add_parser("shadow-summary")
    summary_parser.add_argument("--sqlite", type=Path, required=True)
    summary_parser.add_argument("--days", type=int, default=7)
    benchmark_parser = subparsers.add_parser("benchmark-read")
    benchmark_parser.add_argument("--concurrency", type=int, default=30)
    benchmark_parser.add_argument("--iterations", type=int, default=5)
    benchmark_parser.add_argument("--report", type=Path)
    preflight_parser = subparsers.add_parser("preflight")
    preflight_parser.add_argument("--evidence", type=Path, required=True)
    preflight_parser.add_argument("--report", type=Path)
    prepare_parser = subparsers.add_parser("prepare-cutover")
    prepare_parser.add_argument("--sqlite", type=Path, required=True)
    prepare_parser.add_argument("--evidence", type=Path, required=True)
    prepare_parser.add_argument("--backup-dir", type=Path, required=True)
    prepare_parser.add_argument("--report", type=Path, required=True)
    capture_parser = subparsers.add_parser("enable-change-capture")
    capture_parser.add_argument("--manifest", type=Path, required=True)
    capture_parser.add_argument("--confirm", required=True)
    replay_parser = subparsers.add_parser("rollback-replay")
    replay_parser.add_argument("--sqlite", type=Path, required=True)
    replay_parser.add_argument("--backup", type=Path, required=True)
    replay_parser.add_argument("--backup-sha256", required=True)
    replay_parser.add_argument("--confirm", required=True)
    replay_parser.add_argument("--batch-size", type=int, default=100)
    observe_parser = subparsers.add_parser("observe")
    observe_parser.add_argument("--sqlite", type=Path, required=True)
    observe_parser.add_argument("--report", type=Path, required=True)
    observation_gate_parser = subparsers.add_parser("observation-gate")
    observation_gate_parser.add_argument("reports", nargs="+", type=Path)
    observation_gate_parser.add_argument("--minimum-days", type=int, default=7)
    observation_gate_parser.add_argument("--report", type=Path)
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
        elif args.command == "rollback-test-target":
            rollback_test_target(args.confirm)
            print(json.dumps({"ok": True, "scope": "automation test target"}))
        elif args.command == "outbox-status":
            print(json.dumps(outbox_status(args.sqlite), ensure_ascii=False, indent=2))
        elif args.command == "sync-outbox":
            if args.poll_seconds <= 0 or args.poll_seconds > 300:
                raise ValueError("poll-seconds must be between 1 and 300")
            while True:
                result = process_outbox(args.sqlite, _shadow_database_url(), batch_size=args.batch_size)
                print(json.dumps(result, ensure_ascii=False))
                if not args.watch:
                    return 0 if result["failed"] == 0 else 2
                time.sleep(args.poll_seconds)
        elif args.command == "shadow-compare":
            result = run_shadow_comparison(args.sqlite, _shadow_database_url())
            if args.report:
                _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0 if result["difference_count"] == 0 else 2
        elif args.command == "shadow-summary":
            print(json.dumps(shadow_summary(args.sqlite, days=args.days), ensure_ascii=False, indent=2))
        elif args.command == "benchmark-read":
            result = run_read_benchmark(
                _shadow_database_url(), concurrency=args.concurrency, iterations=args.iterations
            )
            if args.report:
                _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0 if not result["errors"] else 2
        elif args.command == "preflight":
            result = evaluate_preflight(json.loads(args.evidence.read_text(encoding="utf-8")))
            if args.report:
                _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0 if result["passed"] else 2
        elif args.command == "prepare-cutover":
            evidence = json.loads(args.evidence.read_text(encoding="utf-8"))
            result = prepare_cutover(
                args.sqlite,
                _required_url("AUTOMATION_CUTOVER_DATABASE_URL"),
                evidence,
                args.backup_dir,
            )
            _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
        elif args.command == "enable-change-capture":
            manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
            enable_change_capture(
                _required_url("AUTOMATION_CUTOVER_DATABASE_URL"), manifest, args.confirm
            )
            print(json.dumps({"ok": True, "change_capture": "enabled"}))
        elif args.command == "rollback-replay":
            result = replay_change_log(
                args.sqlite,
                _required_url("AUTOMATION_ROLLBACK_DATABASE_URL"),
                args.backup,
                args.backup_sha256,
                args.confirm,
                batch_size=args.batch_size,
            )
            print(json.dumps(result, ensure_ascii=False, indent=2))
        elif args.command == "observe":
            result = collect_observation(
                args.sqlite, _required_url("AUTOMATION_OBSERVATION_DATABASE_URL")
            )
            _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 2
        elif args.command == "observation-gate":
            result = evaluate_observation_reports(args.reports, minimum_days=args.minimum_days)
            if args.report:
                _write_reports(result, args.report)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0 if result["passed"] else 2
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
