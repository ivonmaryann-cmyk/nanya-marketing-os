from __future__ import annotations

import contextvars
import hashlib
import importlib.metadata
import json
import os
import secrets
import time
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from pathlib import Path
from threading import Lock
from typing import Any, Iterator

from .paths import PDF_EXCEL_CACHE_DIR


PARSER_CACHE_VERSION = "purchase_parser_cache_v1"
PIPELINE_CODE_VERSION = "purchase_order_speed_v4"
PARSER_CACHE_DIR = PDF_EXCEL_CACHE_DIR / "parsed_documents"
_CURRENT_SUMMARY: contextvars.ContextVar[dict[str, Any] | None] = contextvars.ContextVar(
    "purchase_performance_summary",
    default=None,
)
_PARSER_CACHE_LOCK = Lock()

def _positive_int_env(name: str, default: int) -> int:
    try:
        value = int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default
    return value if value > 0 else default


def parser_cache_enabled() -> bool:
    return os.getenv("PDF_EXCEL_PARSER_CACHE", "1").strip().lower() not in {"0", "false", "off", "no"}


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


@lru_cache(maxsize=1)
def dependency_fingerprint() -> str:
    packages = [
        "docling",
        "docling-core",
        "docling-ibm-models",
        "rapidocr-onnxruntime",
        "onnxruntime",
        "pdfplumber",
        "opencv-python-headless",
    ]
    versions: list[str] = []
    for package in packages:
        try:
            version = importlib.metadata.version(package)
        except importlib.metadata.PackageNotFoundError:
            version = "missing"
        versions.append(f"{package}={version}")
    return "|".join(versions)


@lru_cache(maxsize=1)
def pipeline_fingerprint() -> str:
    payload = f"{PIPELINE_CODE_VERSION}|{dependency_fingerprint()}"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]


def new_performance_summary(*, content_sha256: str = "", hash_ms: float = 0.0) -> dict[str, Any]:
    summary: dict[str, Any] = {
        "pipeline_fingerprint": pipeline_fingerprint(),
        "content_sha256": content_sha256,
        "stage_ms": {},
        "ocr_calls": {"page": 0, "cell": 0},
        "ocr_ms": {"page": 0.0, "cell": 0.0},
        "cache": {"parser_hit": False, "parser_key": "", "parser_reason": ""},
        "fast_path": "",
        "fallback_reasons": [],
    }
    if hash_ms > 0:
        summary["stage_ms"]["content_hash"] = round(hash_ms, 3)
    return summary


@contextmanager
def performance_context(summary: dict[str, Any]) -> Iterator[None]:
    token = _CURRENT_SUMMARY.set(summary)
    try:
        yield
    finally:
        _CURRENT_SUMMARY.reset(token)


def activate_performance(summary: dict[str, Any]):
    return _CURRENT_SUMMARY.set(summary)


def reset_performance(token) -> None:
    _CURRENT_SUMMARY.reset(token)


@contextmanager
def performance_stage(name: str) -> Iterator[None]:
    started = time.perf_counter()
    try:
        yield
    finally:
        add_stage_ms(name, (time.perf_counter() - started) * 1000)


def add_stage_ms(name: str, elapsed_ms: float, *, summary: dict[str, Any] | None = None) -> None:
    target = summary or _CURRENT_SUMMARY.get()
    if target is None:
        return
    stages = target.setdefault("stage_ms", {})
    stages[name] = round(float(stages.get(name, 0.0) or 0.0) + float(elapsed_ms), 3)


def record_ocr_call(kind: str, elapsed_ms: float) -> None:
    summary = _CURRENT_SUMMARY.get()
    if summary is None:
        return
    calls = summary.setdefault("ocr_calls", {})
    durations = summary.setdefault("ocr_ms", {})
    calls[kind] = int(calls.get(kind, 0) or 0) + 1
    durations[kind] = round(float(durations.get(kind, 0.0) or 0.0) + float(elapsed_ms), 3)


def append_fallback_reason(reason: str) -> None:
    summary = _CURRENT_SUMMARY.get()
    if summary is None or not reason:
        return
    reasons = summary.setdefault("fallback_reasons", [])
    if reason not in reasons:
        reasons.append(reason)


def set_fast_path(value: str) -> None:
    summary = _CURRENT_SUMMARY.get()
    if summary is not None:
        summary["fast_path"] = value


def set_cache_state(name: str, value: Any) -> None:
    summary = _CURRENT_SUMMARY.get()
    if summary is not None:
        summary.setdefault("cache", {})[name] = value


def _cache_key(content_sha256: str, suffix: str) -> str:
    payload = f"{content_sha256}|{suffix.lower()}|{pipeline_fingerprint()}"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _cache_path(content_sha256: str, suffix: str) -> Path:
    return PARSER_CACHE_DIR / f"{_cache_key(content_sha256, suffix)}.json"


def _scrub_transient_paths(value: Any) -> Any:
    if isinstance(value, list):
        return [_scrub_transient_paths(item) for item in value]
    if not isinstance(value, dict):
        return value
    return {
        key: _scrub_transient_paths(item)
        for key, item in value.items()
        if key not in {"image_path", "clean_image_path", "performance_summary"}
    }


def load_parser_cache(content_sha256: str, suffix: str) -> tuple[dict[str, Any] | None, str, str]:
    key = _cache_key(content_sha256, suffix)
    if not parser_cache_enabled():
        return None, key, "disabled"
    path = _cache_path(content_sha256, suffix)
    if not path.exists():
        return None, key, "miss"
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None, key, "corrupt"
    if payload.get("cache_version") != PARSER_CACHE_VERSION:
        return None, key, "version_mismatch"
    if payload.get("pipeline_fingerprint") != pipeline_fingerprint():
        return None, key, "fingerprint_mismatch"
    document = payload.get("document")
    if not isinstance(document, dict):
        return None, key, "invalid_document"
    try:
        os.utime(path, None)
    except OSError:
        pass
    return document, key, "hit"


def _prune_parser_cache() -> None:
    if not PARSER_CACHE_DIR.exists():
        return
    retention_days = _positive_int_env("PDF_EXCEL_CACHE_RETENTION_DAYS", 30)
    max_bytes = _positive_int_env("PDF_EXCEL_CACHE_MAX_MB", 2048) * 1024 * 1024
    cutoff = datetime.now(timezone.utc) - timedelta(days=retention_days)
    entries: list[tuple[Path, float, int]] = []
    for path in PARSER_CACHE_DIR.glob("*.json"):
        try:
            stat = path.stat()
        except OSError:
            continue
        modified = datetime.fromtimestamp(stat.st_mtime, timezone.utc)
        if modified < cutoff:
            try:
                path.unlink()
            except OSError:
                pass
            continue
        entries.append((path, stat.st_mtime, stat.st_size))
    total = sum(size for _path, _mtime, size in entries)
    for path, _mtime, size in sorted(entries, key=lambda item: item[1]):
        if total <= max_bytes:
            break
        try:
            path.unlink()
            total -= size
        except OSError:
            pass


def save_parser_cache(content_sha256: str, suffix: str, document: dict[str, Any]) -> bool:
    if not parser_cache_enabled() or not content_sha256:
        return False
    with _PARSER_CACHE_LOCK:
        try:
            PARSER_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            path = _cache_path(content_sha256, suffix)
            payload = {
                "cache_version": PARSER_CACHE_VERSION,
                "pipeline_fingerprint": pipeline_fingerprint(),
                "content_sha256": content_sha256,
                "created_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "document": _scrub_transient_paths(document),
            }
            temp_path = path.with_suffix(f".{os.getpid()}.{secrets.token_hex(4)}.tmp")
            temp_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            os.replace(temp_path, path)
            _prune_parser_cache()
            return True
        except Exception:
            return False
