from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import get_setting, set_setting
from .paths import ENGINE_DIR, SHENNAN_RULES_VERSIONS_DIR


SHENNAN_RULE_FILENAME = "shennan_price_rules.xls"
BUILTIN_SHENNAN_RULE_PATH = ENGINE_DIR / SHENNAN_RULE_FILENAME

PRICE_COLUMNS = [
    "CCL",
    "型号",
    "对应基板",
    "不含铜板厚/（mm)",
    "铜厚",
    "铜箔",
    "叠构",
    "尺寸",
    "规格",
    "RMB/SF",
    "每米单价",
    "每卷单价",
    '36"*48"',
    '40"*48"',
    '42"*48"',
    "备注",
]


def _history_key() -> str:
    return "shennan_rule_history"


def _active_key() -> str:
    return "active_shennan_rule_version"


def _read_history() -> list[dict]:
    raw = get_setting(_history_key(), "[]") or "[]"
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _write_history(history: list[dict]) -> None:
    set_setting(_history_key(), json.dumps(history[:50], ensure_ascii=False))


def append_shennan_rule_history(entry: dict) -> None:
    history = _read_history()
    history.insert(0, entry)
    _write_history(history)


def get_shennan_rule_history() -> list[dict]:
    return _read_history()


def ensure_default_shennan_rule_version() -> str:
    active_version = get_setting(_active_key(), "")
    if active_version:
        version_dir = SHENNAN_RULES_VERSIONS_DIR / active_version
        if (version_dir / SHENNAN_RULE_FILENAME).exists():
            return active_version

    if not BUILTIN_SHENNAN_RULE_PATH.exists():
        raise FileNotFoundError(f"未找到内置深南报价单：{BUILTIN_SHENNAN_RULE_PATH}")

    version = datetime.now().strftime("shennan_bootstrap_%Y%m%d_%H%M%S")
    version_dir = SHENNAN_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    target = version_dir / SHENNAN_RULE_FILENAME
    shutil.copy2(BUILTIN_SHENNAN_RULE_PATH, target)
    validate_shennan_rule_file(target)

    set_setting(_active_key(), version)
    append_shennan_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": "system",
            "remark": "由内置深南汇总报价单初始化",
            "rule_file": SHENNAN_RULE_FILENAME,
        }
    )
    return version


def get_active_shennan_rule_version() -> str:
    version = get_setting(_active_key(), "")
    if not version:
        version = ensure_default_shennan_rule_version()
    return version


def get_shennan_rule_file_path(version: str | None = None) -> Path:
    rule_version = version or get_active_shennan_rule_version()
    return SHENNAN_RULES_VERSIONS_DIR / rule_version / SHENNAN_RULE_FILENAME


def save_new_shennan_rule_version(rule_file: FileStorage, *, updated_by: str, remark: str) -> str:
    version = datetime.now().strftime("shennan_rules_%Y%m%d_%H%M%S")
    version_dir = SHENNAN_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    rule_path = version_dir / SHENNAN_RULE_FILENAME

    rule_file.save(rule_path)
    validate_shennan_rule_file(rule_path)

    set_setting(_active_key(), version)
    append_shennan_rule_history(
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "网页上传更新深南报价单",
            "rule_file": secure_filename(rule_file.filename) if rule_file and rule_file.filename else SHENNAN_RULE_FILENAME,
        }
    )
    return version


def validate_shennan_rule_file(path: Path) -> None:
    df = load_shennan_price_dataframe(path)
    missing = set(PRICE_COLUMNS) - set(df.columns)
    if missing:
        raise ValueError(f"深南报价单转换后缺少字段：{', '.join(sorted(missing))}")
    if df.empty:
        raise ValueError("深南报价单没有可读取的价格行")
    if not (df["CCL"].astype(str).str.strip() == "CCL").any():
        raise ValueError("深南报价单没有可读取的 CCL 价格行")
    if not (df["CCL"].astype(str).str.strip() == "PP").any():
        raise ValueError("深南报价单没有可读取的 PP 价格行")


def load_shennan_price_dataframe(path: Path | None = None) -> pd.DataFrame:
    rule_path = path or get_shennan_rule_file_path()
    excel = pd.ExcelFile(rule_path)
    rows: list[dict] = []
    for sheet_name in excel.sheet_names:
        clean_name = str(sheet_name).strip()
        if clean_name.upper().endswith(" CCL"):
            rows.extend(_parse_ccl_sheet(rule_path, clean_name))
        elif clean_name.upper().endswith(" PP"):
            rows.extend(_parse_pp_sheet(rule_path, clean_name))
    return pd.DataFrame(rows, columns=PRICE_COLUMNS).dropna(how="all")


def _parse_ccl_sheet(rule_path: Path, sheet_name: str) -> list[dict]:
    df = pd.read_excel(rule_path, sheet_name=sheet_name, header=None)
    header_row = _find_row(df, ("产品类别", "厚度", "组合"))
    if header_row is None:
        return []

    block_starts = _find_ccl_blocks(df.iloc[header_row])
    if not block_starts:
        return []

    subheader_row = header_row + 1
    data_start = header_row + 2
    rows: list[dict] = []
    for row_idx in range(data_start, len(df)):
        product = _clean_text(df.iat[row_idx, 0])
        thickness = _format_number(df.iat[row_idx, 1])
        laminate = _normalize_laminate(df.iat[row_idx, 3])
        if not product or not thickness or not laminate:
            continue

        for idx, block_start in enumerate(block_starts):
            block_end = block_starts[idx + 1] if idx + 1 < len(block_starts) else df.shape[1]
            label = _clean_text(df.iat[header_row, block_start])
            copper = _copper_from_block_label(label)
            if not copper:
                continue

            price_cols = _find_price_columns(df, subheader_row, block_start, block_end)
            rmb_sf = _number_or_none(df.iat[row_idx, price_cols["RMB/SF"]])
            if rmb_sf is None:
                continue

            row = {
                "CCL": "CCL",
                "型号": product,
                "对应基板": product,
                "不含铜板厚/（mm)": thickness,
                "铜厚": copper,
                "铜箔": _foil_from_block_label(label),
                "叠构": laminate,
                "尺寸": "",
                "规格": f"{product}_{thickness}_{copper}_{laminate}_{sheet_name}",
                "RMB/SF": rmb_sf,
                "每米单价": "",
                "每卷单价": "",
                '36"*48"': _number_or_none(df.iat[row_idx, price_cols['36"*48"']]),
                '40"*48"': _number_or_none(df.iat[row_idx, price_cols['40"*48"']]),
                '42"*48"': _number_or_none(df.iat[row_idx, price_cols['42"*48"']]),
                "备注": sheet_name,
            }
            rows.append(row)
    return rows


def _parse_pp_sheet(rule_path: Path, sheet_name: str) -> list[dict]:
    df = pd.read_excel(rule_path, sheet_name=sheet_name, header=None)
    header_row = _find_row(df, ("产品类别", "产品型号", "树脂含量"))
    if header_row is None:
        return []

    price_cols = _find_pp_price_columns(df, header_row)
    data_start = header_row + 2
    rows: list[dict] = []
    for row_idx in range(data_start, len(df)):
        product = _clean_text(df.iat[row_idx, 0])
        glass_type = _format_number(df.iat[row_idx, 1], integer_if_whole=True)
        rc_value = _normalize_rc(df.iat[row_idx, 2])
        rmb_sf = _number_or_none(df.iat[row_idx, price_cols["RMB/SF"]])
        if not product or not glass_type or not rc_value or rmb_sf is None:
            continue

        rows.append(
            {
                "CCL": "PP",
                "型号": product,
                "对应基板": product.rstrip("P"),
                "不含铜板厚/（mm)": glass_type,
                "铜厚": rc_value,
                "铜箔": "",
                "叠构": glass_type,
                "尺寸": "",
                "规格": f"{product}_{glass_type}_{rc_value}_{sheet_name}",
                "RMB/SF": rmb_sf,
                "每米单价": _number_or_none(df.iat[row_idx, price_cols["每米单价"]]),
                "每卷单价": _number_or_none(df.iat[row_idx, price_cols["每卷单价"]]),
                '36"*48"': "",
                '40"*48"': "",
                '42"*48"': "",
                "备注": sheet_name,
            }
        )
    return rows


def _find_row(df: pd.DataFrame, keywords: tuple[str, ...]) -> int | None:
    for row_idx in range(len(df)):
        row_text = " ".join(_clean_text(value) for value in df.iloc[row_idx].tolist())
        if all(keyword in row_text for keyword in keywords):
            return row_idx
    return None


def _find_ccl_blocks(header_row: pd.Series) -> list[int]:
    starts: list[int] = []
    for col_idx, value in enumerate(header_row.tolist()):
        text = _clean_text(value).upper()
        if "OZ" in text and "/" in text:
            starts.append(col_idx)
    return starts


def _find_price_columns(df: pd.DataFrame, subheader_row: int, start: int, end: int) -> dict[str, int]:
    columns = {"RMB/SF": start, '36"*48"': start + 1, '40"*48"': start + 2, '42"*48"': start + 3}
    for col_idx in range(start, min(end, df.shape[1])):
        text = _clean_text(df.iat[subheader_row, col_idx]).upper()
        if "36" in text and "48" in text:
            columns['36"*48"'] = col_idx
        elif "40" in text and "48" in text:
            columns['40"*48"'] = col_idx
        elif "42" in text and "48" in text:
            columns['42"*48"'] = col_idx
        elif "SF" in text:
            columns["RMB/SF"] = col_idx
    return columns


def load_shennan_surcharge_rules(path: Path | None = None) -> dict[str, dict]:
    rule_path = path or get_shennan_rule_file_path()
    excel = pd.ExcelFile(rule_path)
    rules: dict[str, dict] = {}
    for sheet_name in excel.sheet_names:
        clean_name = str(sheet_name).strip()
        if not clean_name.upper().endswith(" CCL"):
            continue
        glue = clean_name.rsplit(" ", 1)[0]
        df = pd.read_excel(rule_path, sheet_name=clean_name, header=None)
        text = "\n".join(_clean_text(value) for value in df.to_numpy().ravel())
        rtf_percent = _parse_rtf_percent_surcharge(text)
        if rtf_percent:
            rules.setdefault(glue, {})["RTF"] = rtf_percent
            rules.setdefault(glue, {})["RTF1"] = rtf_percent
        rtf2 = _parse_rtf2_surcharge(text)
        if rtf2:
            rules.setdefault(glue, {})["RTF2"] = rtf2
        rtf3 = _parse_rtf3_surcharge(text)
        if rtf3:
            rules.setdefault(glue, {})["RTF3"] = rtf3
    return rules


def _parse_rtf_percent_surcharge(text: str) -> dict | None:
    if "RTF" not in text or "RTF2" in text:
        return None
    if not re.search(r"(上调|加)\s*3\s*%", text):
        return None
    return {"type": "percent", "percent": 0.03}


def _parse_rtf2_surcharge(text: str) -> dict | None:
    if "RTF2" not in text:
        return None
    window_match = re.search(r"RTF2.{0,160}", text, re.IGNORECASE | re.DOTALL)
    window = window_match.group(0) if window_match else text
    values = [float(value) for value in re.findall(r"(?:增加|加)\s*(\d+(?:\.\d+)?)\s*元(?:/SF)?", window)]
    if len(values) < 6:
        return None
    return {
        "type": "per_sf",
        "H": {"single": values[0], "double": values[1]},
        "1": {"single": values[2], "double": values[3]},
        "2": {"single": values[4], "double": values[5]},
    }


def _parse_rtf3_surcharge(text: str) -> dict | None:
    if "RTF3" not in text:
        return None
    hoz_single = _extract_surcharge_value(text, r"HOZ\s*单面加\s*(\d+(?:\.\d+)?)\s*元/SF")
    hoz_double = _extract_surcharge_value(text, r"HOZ\s*双面加\s*(\d+(?:\.\d+)?)\s*元/SF")
    one_single = _extract_surcharge_value(text, r"1OZ\s*单面加\s*(\d+(?:\.\d+)?)\s*元/SF")
    one_double = _extract_surcharge_value(text, r"1OZ.*?双面加\s*(\d+(?:\.\d+)?)\s*元/SF")
    if any(value is None for value in [hoz_single, hoz_double, one_single, one_double]):
        values = [float(value) for value in re.findall(r"加\s*(\d+(?:\.\d+)?)\s*元/SF", text)]
        if len(values) >= 4:
            hoz_single, hoz_double, one_single, one_double = values[:4]
    if any(value is None for value in [hoz_single, hoz_double, one_single, one_double]):
        return None
    return {
        "type": "per_sf",
        "H": {"single": hoz_single, "double": hoz_double},
        "1": {"single": one_single, "double": one_double},
    }


def _extract_surcharge_value(text: str, pattern: str) -> float | None:
    match = re.search(pattern, text, re.IGNORECASE)
    return float(match.group(1)) if match else None


def _find_pp_price_columns(df: pd.DataFrame, header_row: int) -> dict[str, int]:
    columns = {"RMB/SF": 5, "每米单价": 6, "每卷单价": 7}
    for row_idx in [header_row, header_row + 1]:
        if row_idx >= len(df):
            continue
        for col_idx, value in enumerate(df.iloc[row_idx].tolist()):
            text = _clean_text(value).upper()
            if "SF" in text or "平方英尺" in text:
                columns["RMB/SF"] = col_idx
            elif "每米" in text:
                columns["每米单价"] = col_idx
            elif "每卷" in text:
                columns["每卷单价"] = col_idx
    return columns


def _copper_from_block_label(label: str) -> str | None:
    upper = label.upper().replace(" ", "")
    if upper.startswith("1OZ/1OZ"):
        return "1/1"
    if upper.startswith("2OZ/2OZ"):
        return "2/2"
    if upper.startswith("HOZ/HOZ"):
        return "H/H"
    return None


def _foil_from_block_label(label: str) -> str:
    upper = label.upper()
    if "RTF" in upper:
        return "RTF"
    return "HTE"


def _normalize_laminate(value) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    text = text.replace("×", "x").replace("*", "x").replace("X", "x")
    text = re.sub(r"\s+", "", text)
    match = re.match(r"^(\d+(?:\.0+)?)x(\d+(?:\.0+)?)$", text)
    if not match:
        return text
    left = _strip_numeric_suffix(match.group(1))
    right = _strip_numeric_suffix(match.group(2))
    return f"{left}x{right}"


def _normalize_rc(value) -> str:
    text = _clean_text(value).replace("％", "%")
    if not text:
        return ""
    text = text.replace("%", "").strip()
    match = re.match(r"^([<>≤≥≦≧]=?)?\s*(\d+(?:\.\d+)?)$", text)
    if not match:
        return text
    operator = match.group(1) or ""
    number = float(match.group(2))
    if number <= 1:
        number *= 100
    number_text = _strip_numeric_suffix(f"{number:.4f}".rstrip("0").rstrip("."))
    return f"{operator}{number_text}"


def _format_number(value, *, integer_if_whole: bool = False) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)):
        if integer_if_whole and float(value).is_integer():
            return str(int(value))
        return f"{float(value):.4f}".rstrip("0").rstrip(".")
    text = _clean_text(value)
    match = re.match(r"^\d+(?:\.0+)?$", text)
    if integer_if_whole and match:
        return _strip_numeric_suffix(text)
    return text


def _strip_numeric_suffix(text: str) -> str:
    return re.sub(r"\.0+$", "", text)


def _number_or_none(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _clean_text(value).replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").replace("\u3000", " ").strip()
