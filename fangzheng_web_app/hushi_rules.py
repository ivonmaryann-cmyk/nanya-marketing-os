from __future__ import annotations

import json
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import get_setting, set_setting
from .paths import DEFAULT_RULES_DIR, HUSHI_RULES_VERSIONS_DIR


HUSHI_RULE_FILES_DIRNAME = "files"
ALLOWED_RULE_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def _history_key() -> str:
    return "hushi_rule_history"


def _active_key() -> str:
    return "active_hushi_rule_version"


def _read_history() -> list[dict]:
    raw = get_setting(_history_key(), "[]") or "[]"
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _write_history(history: list[dict]) -> None:
    set_setting(_history_key(), json.dumps(history[:50], ensure_ascii=False))


def append_hushi_rule_history(entry: dict) -> None:
    history = _read_history()
    history.insert(0, entry)
    _write_history(history)


def get_hushi_rule_history() -> list[dict]:
    return _read_history()


def get_active_hushi_rule_version() -> str:
    version = get_setting(_active_key(), "") or ""
    if version and _has_hushi_rule_files(get_hushi_rule_dir(version)):
        return version
    version = ensure_default_hushi_rule_version()
    return version


def get_hushi_rule_dir(version: str | None = None) -> Path:
    rule_version = version or (get_setting(_active_key(), "") or "")
    return HUSHI_RULES_VERSIONS_DIR / rule_version / HUSHI_RULE_FILES_DIRNAME


def ensure_default_hushi_rule_version() -> str:
    active_version = get_setting(_active_key(), "") or ""
    if active_version and _has_hushi_rule_files(get_hushi_rule_dir(active_version)):
        return active_version

    seed_zip = DEFAULT_RULES_DIR / "hushi_rules.zip"
    if not seed_zip.exists():
        return ""

    version = datetime.now().strftime("hushi_bootstrap_%Y%m%d_%H%M%S")
    version_dir = HUSHI_RULES_VERSIONS_DIR / version
    files_dir = version_dir / HUSHI_RULE_FILES_DIRNAME
    files_dir.mkdir(parents=True, exist_ok=True)

    extract_hushi_zip(seed_zip, files_dir)
    copied = len([path for path in files_dir.iterdir() if path.is_file() and path.suffix.lower() in ALLOWED_RULE_EXTENSIONS])
    if copied == 0:
        shutil.rmtree(version_dir, ignore_errors=True)
        return ""

    validate_hushi_rule_dir(files_dir)
    set_setting(_active_key(), version)
    append_hushi_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": "system",
            "remark": "由内置沪士报价规则包初始化",
            "rule_file": f"{copied} 个报价 Excel",
        }
    )
    return version


def save_new_hushi_rule_version(rule_zip: FileStorage, *, updated_by: str, remark: str) -> str:
    version = datetime.now().strftime("hushi_rules_%Y%m%d_%H%M%S")
    version_dir = HUSHI_RULES_VERSIONS_DIR / version
    files_dir = version_dir / HUSHI_RULE_FILES_DIRNAME
    files_dir.mkdir(parents=True, exist_ok=True)

    original_name = secure_filename(rule_zip.filename or "hushi_rules.zip") or "hushi_rules.zip"
    zip_path = version_dir / original_name
    rule_zip.save(zip_path)
    extract_hushi_zip(zip_path, files_dir)
    validate_hushi_rule_dir(files_dir)

    set_setting(_active_key(), version)
    append_hushi_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "网页上传沪士报价规则包",
            "rule_file": rule_zip.filename or original_name,
        }
    )
    return version


def extract_hushi_zip(zip_path: Path, target_dir: Path) -> None:
    if zip_path.suffix.lower() != ".zip":
        raise ValueError("沪士规则请上传 .zip 压缩包")

    target_root = target_dir.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            if member.is_dir():
                continue
            member_path = Path(member.filename)
            if member_path.suffix.lower() not in ALLOWED_RULE_EXTENSIONS:
                continue
            safe_name = secure_filename(member_path.name) or member_path.name
            destination = (target_dir / safe_name).resolve()
            if not str(destination).startswith(str(target_root)):
                raise ValueError("ZIP 包中包含非法路径")
            with archive.open(member) as source, destination.open("wb") as target:
                shutil.copyfileobj(source, target)


def validate_hushi_rule_dir(rule_dir: Path) -> None:
    files = [
        path
        for path in rule_dir.iterdir()
        if path.is_file() and path.suffix.lower() in ALLOWED_RULE_EXTENSIONS
    ]
    if not files:
        raise ValueError("沪士规则包中没有可读取的 Excel 报价文件")


def _has_hushi_rule_files(rule_dir: Path) -> bool:
    return rule_dir.exists() and any(
        path.is_file() and path.suffix.lower() in ALLOWED_RULE_EXTENSIONS
        for path in rule_dir.iterdir()
    )
