from __future__ import annotations

import hashlib
import io
import logging
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Mapping

from ..paths import STORAGE_DIR
from .base import FileStorage


LOG = logging.getLogger(__name__)


def _enabled(name: str, default: bool = False) -> bool:
    value = os.getenv(name, "true" if default else "false").strip().lower()
    return value in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class AutomationFileWrite:
    legacy_path: Path
    object_key: str
    managed_written: bool


class LocalFileStorage:
    def __init__(self, root: Path):
        self.root = root.resolve()

    def object_path(self, object_key: str) -> Path:
        candidate = (self.root / object_key).resolve()
        if candidate != self.root and self.root not in candidate.parents:
            raise ValueError("object key escapes storage root")
        return candidate

    def save(self, stream: BinaryIO, object_key: str, metadata: Mapping[str, str] | None = None) -> Path:
        target = self.object_path(object_key)
        target.parent.mkdir(parents=True, exist_ok=True)
        handle, temporary_name = tempfile.mkstemp(prefix=f".{target.name}.", dir=target.parent)
        try:
            with os.fdopen(handle, "wb") as temporary:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    temporary.write(chunk)
            os.replace(temporary_name, target)
        except Exception:
            Path(temporary_name).unlink(missing_ok=True)
            raise
        return target

    def open(self, *, object_key: str = "", legacy_path: str = "") -> BinaryIO:
        return self.resolve(object_key=object_key, legacy_path=legacy_path).open("rb")

    def resolve(self, *, object_key: str = "", legacy_path: str = "") -> Path:
        if object_key:
            path = self.object_path(object_key)
            if path.is_file():
                return path
        if legacy_path:
            path = Path(legacy_path)
            if path.is_file():
                return path
        raise FileNotFoundError("file is unavailable in configured or legacy storage")

    def exists(self, *, object_key: str = "", legacy_path: str = "") -> bool:
        try:
            self.resolve(object_key=object_key, legacy_path=legacy_path)
            return True
        except FileNotFoundError:
            return False

    def checksum(self, *, object_key: str = "", legacy_path: str = "") -> str:
        path = self.resolve(object_key=object_key, legacy_path=legacy_path)
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def delete(self, object_key: str) -> None:
        self.object_path(object_key).unlink(missing_ok=True)

    def temporary_download_url(self, object_key: str, expires_seconds: int = 300) -> str | None:
        if expires_seconds <= 0:
            raise ValueError("expires_seconds must be greater than zero")
        return None


class FallbackFileStorage(LocalFileStorage):
    """Read a managed object first, then preserve access to the legacy local path."""

    def resolve(self, *, object_key: str = "", legacy_path: str = "") -> Path:
        if object_key:
            managed = self.object_path(object_key)
            if managed.is_file():
                return managed
        if legacy_path and Path(legacy_path).is_file():
            LOG.info("automation attachment resolved through legacy storage fallback")
            return Path(legacy_path)
        raise FileNotFoundError("file is unavailable in configured or legacy storage")


def automation_object_key(legacy_path: str | Path, *, legacy_root: Path = STORAGE_DIR) -> str:
    root = legacy_root.resolve()
    path = Path(legacy_path).resolve()
    try:
        relative = path.relative_to(root)
    except ValueError as exc:
        raise ValueError("automation file is outside the legacy storage root") from exc
    if not relative.parts or relative.parts[0] != "mail_transcode":
        raise ValueError("automation object keys are limited to mail_transcode files")
    return relative.as_posix()


def save_automation_file(
    payload: bytes,
    legacy_path: str | Path,
    *,
    legacy_root: Path = STORAGE_DIR,
    managed_storage: FileStorage | None = None,
    dual_write: bool | None = None,
) -> AutomationFileWrite:
    path = Path(legacy_path)
    object_key = automation_object_key(path, legacy_root=legacy_root)
    enabled = _enabled("AUTOMATION_FILE_DUAL_WRITE_ENABLED") if dual_write is None else dual_write
    managed = managed_storage or LocalFileStorage(
        Path(os.getenv("AUTOMATION_FILE_STORAGE_ROOT", "storage/automation_objects"))
    )
    managed_existed = managed.exists(object_key=object_key)
    if enabled:
        managed.save(io.BytesIO(payload), object_key)
    try:
        relative = path.resolve().relative_to(legacy_root.resolve()).as_posix()
        LocalFileStorage(legacy_root).save(io.BytesIO(payload), relative)
    except Exception:
        if enabled and not managed_existed:
            managed.delete(object_key)
        raise
    return AutomationFileWrite(path.resolve(), object_key, enabled)


def resolve_attachment_path(
    legacy_path: str,
    object_key: str = "",
    *,
    legacy_root: Path = STORAGE_DIR,
    managed_root: Path | None = None,
) -> Path:
    root = managed_root or Path(os.getenv("AUTOMATION_FILE_STORAGE_ROOT", "storage/automation_objects"))
    if not object_key and legacy_path and _enabled("AUTOMATION_FILE_DUAL_READ_ENABLED"):
        try:
            object_key = automation_object_key(legacy_path, legacy_root=legacy_root)
        except ValueError:
            object_key = ""
    return FallbackFileStorage(root).resolve(object_key=object_key, legacy_path=legacy_path)


def is_allowed_automation_path(path: str | Path) -> bool:
    resolved = Path(path).resolve()
    roots = (
        STORAGE_DIR.resolve(),
        Path(os.getenv("AUTOMATION_FILE_STORAGE_ROOT", "storage/automation_objects")).resolve(),
    )
    return any(resolved == root or root in resolved.parents for root in roots)
