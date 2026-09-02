from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
from typing import Any

from ..database.automation import automation_cursor
from ..local_env import load_local_env
from ..paths import STORAGE_DIR
from .local import LocalFileStorage, automation_object_key


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _entry(
    *,
    kind: str,
    record_id: int,
    legacy_path: str,
    expected_size: int = 0,
    expected_sha256: str = "",
    metadata: dict[str, Any] | None = None,
    managed: LocalFileStorage,
    legacy_root: Path = STORAGE_DIR,
) -> dict[str, Any]:
    source = Path(legacy_path)
    result: dict[str, Any] = {
        "kind": kind,
        "record_id": record_id,
        "legacy_path": legacy_path,
        "object_key": "",
        "expected_size": int(expected_size or 0),
        "expected_sha256": str(expected_sha256 or ""),
        "source_exists": source.is_file(),
        "source_size": 0,
        "source_sha256": "",
        "managed_exists": False,
        "managed_size": 0,
        "managed_sha256": "",
        "status": "pending",
        "error": "",
        "metadata": metadata or {},
    }
    try:
        result["object_key"] = automation_object_key(source, legacy_root=legacy_root)
    except ValueError as exc:
        result["status"] = "failed"
        result["error"] = str(exc)
        return result
    if not source.is_file():
        result["status"] = "failed"
        result["error"] = "legacy source is missing"
        return result
    result["source_size"] = source.stat().st_size
    result["source_sha256"] = _sha256(source)
    if expected_size and result["source_size"] != expected_size:
        result["status"] = "failed"
        result["error"] = "legacy size does not match metadata"
        return result
    if expected_sha256 and result["source_sha256"] != expected_sha256:
        result["status"] = "failed"
        result["error"] = "legacy checksum does not match metadata"
        return result
    object_key = result["object_key"]
    result["managed_exists"] = managed.exists(object_key=object_key)
    if not result["managed_exists"]:
        return result
    managed_path = managed.object_path(object_key)
    result["managed_size"] = managed_path.stat().st_size
    result["managed_sha256"] = managed.checksum(object_key=object_key)
    if (
        result["managed_size"] == result["source_size"]
        and result["managed_sha256"] == result["source_sha256"]
    ):
        result["status"] = "verified"
    else:
        result["status"] = "failed"
        result["error"] = "managed copy differs from legacy source"
    return result


def build_inventory() -> dict[str, Any]:
    managed = LocalFileStorage(
        Path(os.getenv("AUTOMATION_FILE_STORAGE_ROOT", "storage/automation_objects"))
    )
    entries: list[dict[str, Any]] = []
    with automation_cursor() as connection:
        messages = connection.execute(
            "SELECT id,account_id,folder,uid,eml_path FROM mail_messages "
            "WHERE TRIM(eml_path)<>'' ORDER BY id"
        ).fetchall()
        attachments = connection.execute(
            "SELECT id,mail_id,filename,content_type,size_bytes,sha256,stored_path,is_inline "
            "FROM mail_attachments WHERE TRIM(stored_path)<>'' ORDER BY id"
        ).fetchall()
        template_versions = int(connection.execute(
            "SELECT COUNT(*) AS total FROM order_entry_template_versions"
        ).fetchone()["total"])
    for row in messages:
        entries.append(_entry(
            kind="original_eml",
            record_id=int(row["id"]),
            legacy_path=str(row["eml_path"]),
            metadata={"account_id": int(row["account_id"]), "folder": row["folder"], "uid": row["uid"]},
            managed=managed,
        ))
    for row in attachments:
        entries.append(_entry(
            kind="inline_image" if row["is_inline"] else "attachment",
            record_id=int(row["id"]),
            legacy_path=str(row["stored_path"]),
            expected_size=int(row["size_bytes"] or 0),
            expected_sha256=str(row["sha256"] or ""),
            metadata={
                "mail_id": int(row["mail_id"]),
                "filename": row["filename"],
                "content_type": row["content_type"],
            },
            managed=managed,
        ))
    summary = {
        "total": len(entries),
        "verified": sum(entry["status"] == "verified" for entry in entries),
        "pending": sum(entry["status"] == "pending" for entry in entries),
        "failed": sum(entry["status"] == "failed" for entry in entries),
        "database_only_template_versions": template_versions,
    }
    return {"ok": summary["failed"] == 0, "summary": summary, "entries": entries}


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a read-only automation file migration inventory")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    load_local_env()
    report = build_inventory()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report["summary"], ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
