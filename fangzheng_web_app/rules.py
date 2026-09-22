from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import append_rule_history, get_setting, set_setting
from .paths import DEFAULT_ACCOUNT_PKL, DEFAULT_PRICE_PKL, RULES_VERSIONS_DIR


PRICE_FILENAME = "price_rules.xlsx"
ACCOUNT_FILENAME = "account_rules.xlsx"

PRICE_REQUIRED_COLUMNS = {"CCL", "型号", "不含铜板厚/（mm)", "铜厚", "铜箔", "叠构"}
ACCOUNT_REQUIRED_COLUMNS = {"品名", "小片数量", "大板规格"}


def _clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df.dropna(how="all")


def rule_version_exists(version: str | None) -> bool:
    if not version:
        return False
    version_dir = RULES_VERSIONS_DIR / version
    return version_dir.exists() and all(
        (version_dir / name).exists() for name in [PRICE_FILENAME, ACCOUNT_FILENAME]
    )


def ensure_default_rule_version() -> str:
    active_version = get_setting("active_rule_version", "")
    if active_version:
        # Rule files live in host-local storage while the active version may
        # live in a shared database. Never replace that shared selection.
        return active_version

    price_pkl = DEFAULT_PRICE_PKL
    account_pkl = DEFAULT_ACCOUNT_PKL
    if not price_pkl.exists() or not account_pkl.exists():
        raise FileNotFoundError(f"未找到默认规则源文件：{price_pkl.parent}")

    version = datetime.now().strftime("bootstrap_%Y%m%d_%H%M%S")
    version_dir = RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)

    pd.read_pickle(price_pkl).to_excel(version_dir / PRICE_FILENAME, index=False, sheet_name="价格对账表")
    pd.read_pickle(account_pkl).to_excel(version_dir / ACCOUNT_FILENAME, index=False, sheet_name="基板对照表")

    set_setting("active_rule_version", version)
    append_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": "system",
            "remark": "由现有 pkl 数据初始化",
        }
    )
    return version


def get_active_rule_version() -> str:
    version = get_setting("active_rule_version", "")
    if version:
        return version
    return ensure_default_rule_version()


def activate_rule_version(version: str) -> str:
    """Make an existing, complete local rule pack the active Fangzheng version."""
    normalized_version = str(version or "").strip()
    if (
        not normalized_version
        or normalized_version in {".", ".."}
        or Path(normalized_version).name != normalized_version
    ):
        raise ValueError("规则版本格式无效")
    if not rule_version_exists(normalized_version):
        raise ValueError(
            "该规则版本在本机不完整，需同时存在价格对账表和基板对照表后才能启用"
        )

    price_path, account_path = get_rule_file_paths(normalized_version)
    validate_rule_files(price_path, account_path)
    set_setting("active_rule_version", normalized_version)
    return normalized_version


def get_rule_file_paths(version: str | None = None) -> tuple[Path, Path]:
    rule_version = version or get_active_rule_version()
    version_dir = RULES_VERSIONS_DIR / rule_version
    return version_dir / PRICE_FILENAME, version_dir / ACCOUNT_FILENAME


def _latest_available_rule_file(filename: str) -> Path | None:
    candidates = [
        path for path in RULES_VERSIONS_DIR.glob(f"*/{filename}")
        if path.is_file()
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda path: (path.parent.stat().st_mtime, path.parent.name))


def _read_price_excel(path: Path) -> pd.DataFrame:
    excel = pd.ExcelFile(path)
    if "方正价格" in excel.sheet_names:
        df = pd.read_excel(path, sheet_name="方正价格", header=17)
    elif "价格对账表" in excel.sheet_names:
        df = pd.read_excel(path, sheet_name="价格对账表", header=0)
    else:
        df = pd.read_excel(path, sheet_name=excel.sheet_names[0], header=0)
    return _clean_frame(df)


def _read_account_excel(path: Path) -> pd.DataFrame:
    excel = pd.ExcelFile(path)
    if "基板对照" in excel.sheet_names:
        df = pd.read_excel(path, sheet_name="基板对照", header=0)
    elif "基板对照表" in excel.sheet_names:
        df = pd.read_excel(path, sheet_name="基板对照表", header=0)
    elif "基板对账" in excel.sheet_names:
        df = pd.read_excel(path, sheet_name="基板对账", header=0)
    elif "基板对账表" in excel.sheet_names:
        df = pd.read_excel(path, sheet_name="基板对账表", header=0)
    else:
        df = pd.read_excel(path, sheet_name=excel.sheet_names[0], header=0)
    return _clean_frame(df)


def load_rule_dataframes(version: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    price_path, account_path = get_rule_file_paths(version)
    missing = [str(path) for path in (price_path, account_path) if not path.is_file()]
    if missing:
        raise ValueError(
            "方正当前生效报价文件在本机缺失，请恢复 storage 中对应版本文件或重新上传报价单："
            + "；".join(missing)
        )
    return _read_price_excel(price_path), _read_account_excel(account_path)


def validate_rule_files(price_path: Path, account_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    price_df = _read_price_excel(price_path)
    account_df = _read_account_excel(account_path)

    missing_price = PRICE_REQUIRED_COLUMNS - set(price_df.columns)
    if missing_price:
        raise ValueError(f"价格对账表缺少字段：{', '.join(sorted(missing_price))}")

    missing_account = ACCOUNT_REQUIRED_COLUMNS - set(account_df.columns)
    if missing_account:
        raise ValueError(f"基板对照表缺少字段：{', '.join(sorted(missing_account))}")

    if price_df.empty:
        raise ValueError("价格对账表为空")
    if account_df.empty:
        raise ValueError("基板对照表为空")

    return price_df, account_df


def save_new_rule_version(
    price_file: FileStorage | None,
    account_file: FileStorage | None,
    *,
    updated_by: str,
    remark: str,
) -> str:
    current_price, current_account = get_rule_file_paths()
    price_source = current_price if current_price.is_file() else _latest_available_rule_file(PRICE_FILENAME)
    account_source = current_account if current_account.is_file() else _latest_available_rule_file(ACCOUNT_FILENAME)
    reuse_missing: list[str] = []
    if not (price_file and price_file.filename) and price_source is None:
        reuse_missing.append("价格对账表")
    if not (account_file and account_file.filename) and account_source is None:
        reuse_missing.append("基板对照表")
    if reuse_missing:
        raise ValueError(
            "当前生效方正规则文件在本机缺失，且找不到可沿用的"
            + "、".join(reuse_missing)
            + "。请同时上传价格对账表和基板对照表以恢复规则版本。"
        )

    version = datetime.now().strftime("rules_%Y%m%d_%H%M%S")
    version_dir = RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    price_path = version_dir / PRICE_FILENAME
    account_path = version_dir / ACCOUNT_FILENAME

    if price_file and price_file.filename:
        price_file.save(price_path)
    else:
        shutil.copy2(price_source, price_path)

    if account_file and account_file.filename:
        account_file.save(account_path)
    else:
        shutil.copy2(account_source, account_path)

    validate_rule_files(price_path, account_path)
    set_setting("active_rule_version", version)
    append_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "网页上传更新规则",
            "price_file": secure_filename(price_file.filename) if price_file and price_file.filename else PRICE_FILENAME,
            "account_file": secure_filename(account_file.filename) if account_file and account_file.filename else ACCOUNT_FILENAME,
        }
    )
    return version
