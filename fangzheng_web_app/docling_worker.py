from __future__ import annotations

import json
import os
import secrets
import socket
import subprocess
import sys
import time
from datetime import datetime, timezone
from multiprocessing.connection import Client, Listener
from pathlib import Path
from typing import Any

from .paths import PDF_EXCEL_CACHE_DIR, PROJECT_DIR
from .purchase_performance import dependency_fingerprint


WORKER_PROTOCOL_VERSION = "docling_worker_v2"
WORKER_DIR = PDF_EXCEL_CACHE_DIR / "docling_worker"
DESCRIPTOR_PATH = WORKER_DIR / "worker.json"
START_LOCK_PATH = WORKER_DIR / "worker.start.lock"
LOG_PATH = WORKER_DIR / "worker.log"


def _enabled() -> bool:
    return os.getenv("PDF_EXCEL_DOCLING_WORKER", "1").strip().lower() not in {"0", "false", "off", "no"}


def _idle_seconds() -> int:
    try:
        return max(30, int(os.getenv("PDF_EXCEL_DOCLING_IDLE_SECONDS", "900")))
    except ValueError:
        return 900


def _request_timeout_seconds() -> int:
    try:
        return max(30, int(os.getenv("PDF_EXCEL_DOCLING_REQUEST_TIMEOUT", "180")))
    except ValueError:
        return 180


def _worker_fingerprint() -> str:
    return f"{WORKER_PROTOCOL_VERSION}|{dependency_fingerprint()}"


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def _read_descriptor() -> dict[str, Any] | None:
    try:
        descriptor = json.loads(DESCRIPTOR_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None
    if descriptor.get("fingerprint") != _worker_fingerprint():
        return None
    if not _pid_alive(int(descriptor.get("pid") or 0)):
        return None
    return descriptor


def _write_descriptor(payload: dict[str, Any]) -> None:
    WORKER_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = DESCRIPTOR_PATH.with_suffix(f".{os.getpid()}.tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    os.replace(temp_path, DESCRIPTOR_PATH)


def _discard_descriptor(expected: dict[str, Any]) -> None:
    try:
        current = json.loads(DESCRIPTOR_PATH.read_text(encoding="utf-8"))
        if int(current.get("pid") or 0) != int(expected.get("pid") or 0):
            return
        if int(current.get("port") or 0) != int(expected.get("port") or 0):
            return
        DESCRIPTOR_PATH.unlink()
    except Exception:
        pass


def _release_start_lock() -> None:
    try:
        START_LOCK_PATH.unlink()
    except OSError:
        pass


def _acquire_start_lock() -> bool:
    WORKER_DIR.mkdir(parents=True, exist_ok=True)
    try:
        descriptor = os.open(str(START_LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        try:
            if time.time() - START_LOCK_PATH.stat().st_mtime > 30:
                START_LOCK_PATH.unlink()
                return _acquire_start_lock()
        except OSError:
            pass
        return False
    with os.fdopen(descriptor, "w", encoding="ascii") as handle:
        handle.write(str(os.getpid()))
    return True


def _launch_worker() -> None:
    WORKER_DIR.mkdir(parents=True, exist_ok=True)
    log_handle = LOG_PATH.open("ab")
    kwargs: dict[str, Any] = {
        "cwd": PROJECT_DIR,
        "stdin": subprocess.DEVNULL,
        "stdout": log_handle,
        "stderr": subprocess.STDOUT,
        "close_fds": os.name != "nt",
    }
    if os.name == "nt":
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags |= subprocess.CREATE_NO_WINDOW
        kwargs["creationflags"] = creationflags
    else:
        kwargs["start_new_session"] = True
    try:
        subprocess.Popen(
            [sys.executable, "-m", "fangzheng_web_app.docling_worker", "--serve"],
            **kwargs,
        )
    finally:
        log_handle.close()


def _ensure_worker() -> dict[str, Any] | None:
    descriptor = _read_descriptor()
    if descriptor is not None:
        return descriptor
    owns_lock = _acquire_start_lock()
    if owns_lock:
        try:
            _launch_worker()
            deadline = time.monotonic() + 15
            while time.monotonic() < deadline:
                descriptor = _read_descriptor()
                if descriptor is not None:
                    return descriptor
                time.sleep(0.1)
            return None
        finally:
            _release_start_lock()
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        descriptor = _read_descriptor()
        if descriptor is not None:
            return descriptor
        time.sleep(0.1)
    return None


def request_docling_parse(path: Path, content_sha256: str = "") -> dict[str, Any] | None:
    if not _enabled():
        return None
    descriptor = _ensure_worker()
    if descriptor is None:
        return None
    try:
        connection = Client(
            (str(descriptor["host"]), int(descriptor["port"])),
            authkey=bytes.fromhex(str(descriptor["authkey"])),
        )
        try:
            connection.send(
                {
                    "action": "parse",
                    "path": str(path.resolve()),
                    "content_sha256": content_sha256,
                    "fingerprint": _worker_fingerprint(),
                }
            )
            if not connection.poll(_request_timeout_seconds()):
                return None
            response = connection.recv()
        finally:
            connection.close()
    except Exception:
        _discard_descriptor(descriptor)
        return None
    if not isinstance(response, dict) or not response.get("ok"):
        return None
    result = response.get("result")
    if not isinstance(result, dict):
        return None
    result["worker_used"] = True
    return result


def _serve() -> int:
    WORKER_DIR.mkdir(parents=True, exist_ok=True)
    authkey = secrets.token_bytes(32)
    listener = Listener(("127.0.0.1", 0), authkey=authkey)
    host, port = listener.address
    descriptor = {
        "protocol": WORKER_PROTOCOL_VERSION,
        "fingerprint": _worker_fingerprint(),
        "pid": os.getpid(),
        "host": host,
        "port": int(port),
        "authkey": authkey.hex(),
        "started_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    _write_descriptor(descriptor)
    listener._listener._socket.settimeout(1.0)  # type: ignore[attr-defined]
    last_activity = time.monotonic()
    try:
        while time.monotonic() - last_activity < _idle_seconds():
            try:
                connection = listener.accept()
            except socket.timeout:
                continue
            try:
                request = connection.recv()
                last_activity = time.monotonic()
                if not isinstance(request, dict) or request.get("fingerprint") != _worker_fingerprint():
                    connection.send({"ok": False, "error": "worker fingerprint mismatch"})
                    continue
                if request.get("action") != "parse":
                    connection.send({"ok": False, "error": "unsupported action"})
                    continue
                from .docling_parser import parse_pdf_with_docling_local

                result = parse_pdf_with_docling_local(
                    Path(str(request.get("path") or "")),
                    content_sha256=str(request.get("content_sha256") or ""),
                )
                connection.send({"ok": True, "result": result})
            except Exception as exc:
                try:
                    connection.send({"ok": False, "error": str(exc)})
                except Exception:
                    pass
            finally:
                connection.close()
    finally:
        listener.close()
        current = _read_descriptor()
        if current and int(current.get("pid") or 0) == os.getpid():
            try:
                DESCRIPTOR_PATH.unlink()
            except OSError:
                pass
    return 0


def main(argv: list[str] | None = None) -> int:
    args = argv or sys.argv[1:]
    if args != ["--serve"]:
        print("Usage: python -m fangzheng_web_app.docling_worker --serve")
        return 2
    return _serve()


if __name__ == "__main__":
    raise SystemExit(main())
