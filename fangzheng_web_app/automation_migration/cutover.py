from __future__ import annotations

import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .preflight import evaluate_preflight
from .snapshot import sqlite_snapshot
from .sync import outbox_status
from .verify import verify_snapshot


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _authorization_token(manifest: dict[str, Any]) -> str:
    safe = {key: value for key, value in manifest.items() if key != "authorization_token"}
    payload = json.dumps(safe, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]


def prepare_cutover(
    sqlite_path: Path,
    database_url: str,
    evidence: dict[str, Any],
    backup_dir: Path,
) -> dict[str, Any]:
    gate = evaluate_preflight(evidence)
    if not gate["passed"]:
        raise RuntimeError("cutover preflight is blocked")
    owners = [str(evidence.get(name, "")).strip() for name in ("switch_owner", "review_owner", "rollback_owner")]
    if any(not owner for owner in owners) or len(set(owners)) != 3:
        raise RuntimeError("three distinct named cutover owners are required")
    if not str(evidence.get("maintenance_window", "")).strip():
        raise RuntimeError("maintenance_window is required")
    status = outbox_status(sqlite_path)
    if status["pending"] != 0:
        raise RuntimeError("outbox is not empty")

    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup_path = backup_dir / f"automation-cutover-{timestamp}.sqlite3"
    with sqlite_snapshot(sqlite_path) as snapshot:
        shutil.copy2(snapshot, backup_path)

    with psycopg.connect(database_url, row_factory=dict_row) as target:
        verification = verify_snapshot(backup_path, target, check_files=True)
        if not verification["ok"]:
            backup_path.unlink(missing_ok=True)
            raise RuntimeError("cutover data verification failed")
        versions = [row["version"] for row in target.execute(
            "SELECT version FROM automation_schema_migrations ORDER BY version"
        ).fetchall()]
        flag = target.execute(
            "SELECT value FROM automation_runtime_flags WHERE key='capture_changes'"
        ).fetchone()
        if not flag or flag["value"] != "false":
            raise RuntimeError("change capture must be disabled during preparation")

    final_status = outbox_status(sqlite_path)
    if final_status["pending"] != 0:
        raise RuntimeError("new outbox events appeared during cutover preparation")

    manifest = {
        "prepared_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "sqlite_backup": str(backup_path),
        "sqlite_backup_sha256": _sha256(backup_path),
        "schema_versions": versions,
        "outbox_status": final_status,
        "owners": {"switch": owners[0], "review": owners[1], "rollback": owners[2]},
        "maintenance_window": evidence["maintenance_window"],
        "data_verification_ok": True,
        "formal_switch_executed": False,
    }
    manifest["authorization_token"] = _authorization_token(manifest)
    return manifest


def enable_change_capture(database_url: str, manifest: dict[str, Any], confirmation: str) -> None:
    if manifest.get("formal_switch_executed") is not False:
        raise RuntimeError("invalid cutover manifest state")
    expected = _authorization_token(manifest)
    if confirmation != expected or confirmation != manifest.get("authorization_token"):
        raise RuntimeError("cutover authorization token does not match")
    backup = Path(str(manifest.get("sqlite_backup", "")))
    if not backup.is_file() or _sha256(backup) != manifest.get("sqlite_backup_sha256"):
        raise RuntimeError("cutover SQLite backup is missing or changed")
    with psycopg.connect(database_url, row_factory=dict_row) as target, target.transaction():
        cursor = target.execute(
            "UPDATE automation_runtime_flags SET value='true',updated_at=CURRENT_TIMESTAMP "
            "WHERE key='capture_changes' AND value='false'"
        )
        if cursor.rowcount != 1:
            raise RuntimeError("change capture was already enabled or is unavailable")
