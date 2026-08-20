from __future__ import annotations

import hashlib
import logging
import os
import tempfile
from pathlib import Path
from typing import BinaryIO, Mapping


LOG = logging.getLogger(__name__)


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


def resolve_attachment_path(legacy_path: str, object_key: str = "") -> Path:
    root = Path(os.getenv("AUTOMATION_FILE_STORAGE_ROOT", "storage/automation_objects"))
    return FallbackFileStorage(root).resolve(object_key=object_key, legacy_path=legacy_path)
