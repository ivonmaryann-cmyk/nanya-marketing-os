from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import get_setting, set_setting
from .paths import ENGINE_DIR, TRANSCODE_RULES_VERSIONS_DIR


TRANSCODE_RULE_FILENAME = "transcode_rules.xlsx"
TRANSCODE_RULE_SHEETS = ["胶系代码", "胶系类别", "编码规则", "特殊需求", "总芯厚转换", "客户下单与胶系基板转换"]
BUILTIN_TRANSCODE_RULE_PATH = ENGINE_DIR / TRANSCODE_RULE_FILENAME


def _history_key() -> str:
    return "transcode_rule_history"


def _active_key() -> str:
    return "active_transcode_rule_version"


def _read_history() -> list[dict]:
    import json

    raw = get_setting(_history_key(), "[]") or "[]"
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _write_history(history: list[dict]) -> None:
    import json

    set_setting(_history_key(), json.dumps(history[:50], ensure_ascii=False))


def append_transcode_rule_history(entry: dict) -> None:
    history = _read_history()
    history.insert(0, entry)
    _write_history(history)


def get_transcode_rule_history() -> list[dict]:
    return _read_history()


def validate_transcode_rule_file(path: Path) -> None:
    excel = pd.ExcelFile(path)
    missing = [name for name in TRANSCODE_RULE_SHEETS if name not in excel.sheet_names]
    if missing:
        raise ValueError(f"转码规则文件缺少 Sheet：{', '.join(missing)}")


def ensure_default_transcode_rule_version() -> str:
    active_version = get_setting(_active_key(), "")
    if active_version:
        version_dir = TRANSCODE_RULES_VERSIONS_DIR / active_version
        if (version_dir / TRANSCODE_RULE_FILENAME).exists():
            return active_version

    if not BUILTIN_TRANSCODE_RULE_PATH.exists():
        raise FileNotFoundError(f"未找到内置转码规则文件：{BUILTIN_TRANSCODE_RULE_PATH}")

    version = datetime.now().strftime("transcode_bootstrap_%Y%m%d_%H%M%S")
    version_dir = TRANSCODE_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    target = version_dir / TRANSCODE_RULE_FILENAME
    shutil.copy2(BUILTIN_TRANSCODE_RULE_PATH, target)
    validate_transcode_rule_file(target)

    set_setting(_active_key(), version)
    append_transcode_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": "system",
            "remark": "由内置转码规则初始化",
            "rule_file": TRANSCODE_RULE_FILENAME,
        }
    )
    return version


def get_active_transcode_rule_version() -> str:
    version = get_setting(_active_key(), "")
    if not version:
        version = ensure_default_transcode_rule_version()
    return version


def get_transcode_rule_file_path(version: str | None = None) -> Path:
    rule_version = version or get_active_transcode_rule_version()
    return TRANSCODE_RULES_VERSIONS_DIR / rule_version / TRANSCODE_RULE_FILENAME


def save_new_transcode_rule_version(rule_file: FileStorage, *, updated_by: str, remark: str) -> str:
    version = datetime.now().strftime("transcode_rules_%Y%m%d_%H%M%S")
    version_dir = TRANSCODE_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    rule_path = version_dir / TRANSCODE_RULE_FILENAME

    rule_file.save(rule_path)
    validate_transcode_rule_file(rule_path)

    set_setting(_active_key(), version)
    append_transcode_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "网页上传更新转码规则",
            "rule_file": secure_filename(rule_file.filename) if rule_file and rule_file.filename else TRANSCODE_RULE_FILENAME,
        }
    )
    return version


def save_new_transcode_rule_version_from_sheets(
    sheet_files: dict[str, FileStorage | None],
    *,
    updated_by: str,
    remark: str,
) -> str:
    """Create a full rule workbook from independently maintained sheets."""
    if not any(file_obj and file_obj.filename for file_obj in sheet_files.values()):
        raise ValueError("请至少上传一张需要更新的转码规则表")

    current_path = get_transcode_rule_file_path()
    current_excel = pd.ExcelFile(current_path)
    version = datetime.now().strftime("transcode_rules_%Y%m%d_%H%M%S")
    version_dir = TRANSCODE_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    rule_path = version_dir / TRANSCODE_RULE_FILENAME

    updated_sheets: list[str] = []
    with pd.ExcelWriter(rule_path, engine="openpyxl") as writer:
        for sheet_name in TRANSCODE_RULE_SHEETS:
            file_obj = sheet_files.get(sheet_name)
            if file_obj and file_obj.filename:
                if not file_obj.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
                    raise ValueError(f"{sheet_name} 仅支持 Excel 文件")
                df = pd.read_excel(file_obj, sheet_name=0)
                updated_sheets.append(sheet_name)
            else:
                if sheet_name not in current_excel.sheet_names:
                    raise ValueError(f"当前规则版本缺少 Sheet：{sheet_name}，请上传该规则表")
                df = pd.read_excel(current_path, sheet_name=sheet_name)
            df.to_excel(writer, index=False, sheet_name=sheet_name)

    validate_transcode_rule_file(rule_path)
    set_setting(_active_key(), version)
    append_transcode_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "网页上传更新转码规则",
            "rule_file": TRANSCODE_RULE_FILENAME,
            "updated_sheets": ", ".join(updated_sheets),
        }
    )
    return version
