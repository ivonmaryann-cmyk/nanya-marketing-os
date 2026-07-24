from __future__ import annotations

import argparse
import json
import re
import shutil
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .db import delete_terminal_jobs, list_expired_terminal_jobs, list_job_ids
from .paths import JOBS_DIR, PROJECT_DIR


DEFAULT_RETENTION_DAYS = 30
WORKER_LOG_PATTERN = re.compile(r"^job_(\d+)_worker\.log$")


class UnsafeCleanupPath(ValueError):
    pass


@dataclass
class CleanupReport:
    cutoff: str
    dry_run: bool
    jobs_selected: int = 0
    jobs_deleted: int = 0
    paths_selected: int = 0
    paths_deleted: int = 0
    orphan_logs_selected: int = 0
    orphan_logs_deleted: int = 0
    bytes_selected: int = 0
    errors: list[str] = field(default_factory=list)


def _resolved(path: Path) -> Path:
    return path.expanduser().resolve(strict=False)


def _path_size(path: Path) -> int:
    try:
        if path.is_symlink() or path.is_file():
            return path.stat().st_size
        if path.is_dir():
            return sum(
                child.stat().st_size
                for child in path.rglob("*")
                if child.is_file() and not child.is_symlink()
            )
    except OSError:
        return 0
    return 0


def _remove_path(path: Path) -> None:
    if path.is_symlink() or path.is_file():
        path.unlink(missing_ok=True)
    elif path.is_dir():
        shutil.rmtree(path)


def _job_targets(job: Any, jobs_root: Path) -> set[Path]:
    root = _resolved(jobs_root)
    employee_dir = _resolved(root / str(job["employee_id"]))
    if employee_dir == root or not employee_dir.is_relative_to(root):
        raise UnsafeCleanupPath(f"unsafe employee directory for job {job['id']}")

    targets: set[Path] = {employee_dir / f"job_{int(job['id'])}_worker.log"}
    for key in ("stored_input_path", "stored_result_path"):
        value = job[key]
        if not value:
            continue
        candidate = Path(str(value))
        if not candidate.is_absolute():
            candidate = PROJECT_DIR / candidate
        candidate = _resolved(candidate)
        if not candidate.is_relative_to(employee_dir):
            raise UnsafeCleanupPath(f"{key} is outside the employee job directory for job {job['id']}")

        relative = candidate.relative_to(employee_dir)
        if not relative.parts:
            raise UnsafeCleanupPath(f"{key} points to the employee directory for job {job['id']}")
        if len(relative.parts) > 1:
            targets.add(employee_dir / relative.parts[0])
        else:
            targets.add(candidate)
    return targets


def _cleanup_orphan_worker_logs(
    jobs_root: Path,
    cutoff: datetime,
    existing_job_ids: set[int],
    report: CleanupReport,
) -> None:
    if not jobs_root.exists():
        return
    cutoff_timestamp = cutoff.replace(tzinfo=timezone.utc).timestamp()
    for path in jobs_root.glob("*/job_*_worker.log"):
        match = WORKER_LOG_PATTERN.fullmatch(path.name)
        if not match or int(match.group(1)) in existing_job_ids:
            continue
        try:
            if path.stat().st_mtime >= cutoff_timestamp:
                continue
            report.orphan_logs_selected += 1
            report.bytes_selected += _path_size(path)
            if not report.dry_run:
                _remove_path(path)
                report.orphan_logs_deleted += 1
        except OSError as exc:
            report.errors.append(f"orphan log {path}: {exc}")


def cleanup_expired_jobs(
    *,
    retention_days: int = DEFAULT_RETENTION_DAYS,
    dry_run: bool = False,
    now: datetime | None = None,
    jobs_root: Path | None = None,
) -> CleanupReport:
    if retention_days < 1:
        raise ValueError("retention_days must be at least 1")

    current_time = now or datetime.utcnow()
    cutoff_time = current_time - timedelta(days=retention_days)
    cutoff = cutoff_time.replace(microsecond=0).isoformat()
    root = _resolved(jobs_root or JOBS_DIR)
    report = CleanupReport(cutoff=cutoff, dry_run=dry_run)
    expired_jobs = list_expired_terminal_jobs(cutoff)
    report.jobs_selected = len(expired_jobs)
    deleted_ids: list[int] = []

    for job in expired_jobs:
        try:
            targets = _job_targets(job, root)
        except UnsafeCleanupPath as exc:
            report.errors.append(str(exc))
            continue

        existing_targets = {path for path in targets if path.exists() or path.is_symlink()}
        report.paths_selected += len(existing_targets)
        report.bytes_selected += sum(_path_size(path) for path in existing_targets)
        if dry_run:
            continue
        try:
            for target in sorted(existing_targets, key=lambda item: len(item.parts), reverse=True):
                _remove_path(target)
                report.paths_deleted += 1
            employee_dir = root / str(job["employee_id"])
            try:
                employee_dir.rmdir()
            except OSError:
                pass
            deleted_ids.append(int(job["id"]))
        except OSError as exc:
            report.errors.append(f"job {job['id']}: {exc}")

    if not dry_run:
        report.jobs_deleted = delete_terminal_jobs(deleted_ids)

    existing_ids = list_job_ids()
    _cleanup_orphan_worker_logs(root, cutoff_time, existing_ids, report)
    return report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Delete terminal jobs older than the retention period.")
    parser.add_argument("--days", type=int, default=DEFAULT_RETENTION_DAYS, help="Retention period in days.")
    parser.add_argument("--dry-run", action="store_true", help="Report eligible data without deleting it.")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        report = cleanup_expired_jobs(retention_days=args.days, dry_run=args.dry_run)
    except ValueError as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=True))
        return 2
    print(json.dumps(asdict(report), ensure_ascii=True, indent=2))
    return 1 if report.errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
