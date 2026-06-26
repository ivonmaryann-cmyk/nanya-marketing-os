from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path

from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import get_setting, set_setting
from .excel_utils import load_workbook_compat, normalized_xlsx_source
from .paths import BOMIN_RULES_VERSIONS_DIR, DEFAULT_RULES_DIR


BOMIN_RULE_FILENAME = "bomin_price_rules.xlsx"
ALLOWED_RULE_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def _history_key() -> str:
    return "bomin_rule_history"


def _active_key() -> str:
    return "active_bomin_rule_version"


def _read_history() -> list[dict]:
    raw = get_setting(_history_key(), "[]") or "[]"
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _write_history(history: list[dict]) -> None:
    set_setting(_history_key(), json.dumps(history[:50], ensure_ascii=False))


def append_bomin_rule_history(entry: dict) -> None:
    history = _read_history()
    history.insert(0, entry)
    _write_history(history)


def get_bomin_rule_history() -> list[dict]:
    return _read_history()


def get_bomin_rule_file_path(version: str | None = None) -> Path:
    rule_version = version or (get_setting(_active_key(), "") or "")
    return BOMIN_RULES_VERSIONS_DIR / rule_version / BOMIN_RULE_FILENAME


def get_active_bomin_rule_version() -> str:
    version = get_setting(_active_key(), "") or ""
    if version and get_bomin_rule_file_path(version).exists():
        return version
    return ensure_default_bomin_rule_version()


def ensure_default_bomin_rule_version() -> str:
    active_version = get_setting(_active_key(), "") or ""
    if active_version and get_bomin_rule_file_path(active_version).exists():
        return active_version

    seed_file = DEFAULT_RULES_DIR / BOMIN_RULE_FILENAME
    if not seed_file.exists():
        return ""

    version = datetime.now().strftime("bomin_bootstrap_%Y%m%d_%H%M%S")
    version_dir = BOMIN_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    target = version_dir / BOMIN_RULE_FILENAME
    shutil.copy2(seed_file, target)
    validate_bomin_rule_file(target)

    set_setting(_active_key(), version)
    append_bomin_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": "system",
            "remark": "由内置博敏价格表初始化",
            "rule_file": BOMIN_RULE_FILENAME,
        }
    )
    return version


def save_new_bomin_rule_version(rule_file: FileStorage, *, updated_by: str, remark: str) -> str:
    original_name = secure_filename(rule_file.filename or BOMIN_RULE_FILENAME) or BOMIN_RULE_FILENAME
    if Path(original_name).suffix.lower() not in ALLOWED_RULE_EXTENSIONS:
        raise ValueError("博敏价格规则仅支持 .xlsx / .xls / .xlsm 文件")

    version = datetime.now().strftime("bomin_rules_%Y%m%d_%H%M%S")
    version_dir = BOMIN_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    uploaded_path = version_dir / original_name
    rule_file.save(uploaded_path)

    workbook = load_workbook_compat(uploaded_path, data_only=True)
    xlsx_path = normalized_xlsx_source(uploaded_path, workbook)
    target = version_dir / BOMIN_RULE_FILENAME
    if xlsx_path != target:
        shutil.copy2(xlsx_path, target)
    validate_bomin_rule_file(target)

    set_setting(_active_key(), version)
    append_bomin_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "网页上传博敏价格表",
            "rule_file": rule_file.filename or original_name,
        }
    )
    return version


def validate_bomin_rule_file(path: str | Path) -> None:
    workbook = load_workbook_compat(path, data_only=True)
    sheet_names = {name.strip().upper() for name in workbook.sheetnames}
    missing = {"CCL", "PP"} - sheet_names
    if missing:
        raise ValueError(f"博敏价格表缺少必需 sheet：{', '.join(sorted(missing))}")

