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
    "报价Sheet",
    "支持铜厚",
    "基准铜箔",
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
    if version:
        return version
    return ensure_default_shennan_rule_version()


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
        elif _looks_like_ccl_sheet(rule_path, clean_name):
            # NYHP-7300&7350 is a CCL-only sheet whose name does not carry
            # the CCL suffix.  Detect by its actual table structure instead.
            rows.extend(_parse_ccl_sheet(rule_path, clean_name))
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
                "报价Sheet": sheet_name,
                "支持铜厚": "|".join(_copper_options_from_block_label(label)),
                "基准铜箔": _foil_from_block_label(label),
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
                "报价Sheet": sheet_name,
                "支持铜厚": "",
                "基准铜箔": "",
            }
        )
    return rows


def _find_row(df: pd.DataFrame, keywords: tuple[str, ...]) -> int | None:
    for row_idx in range(len(df)):
        row_text = " ".join(_clean_text(value) for value in df.iloc[row_idx].tolist())
        if all(keyword in row_text for keyword in keywords):
            return row_idx
    return None


def _looks_like_ccl_sheet(rule_path: Path, sheet_name: str) -> bool:
    df = pd.read_excel(rule_path, sheet_name=sheet_name, header=None, nrows=20)
    return _find_row(df, ("产品类别", "厚度", "组合")) is not None


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
    """Read per-sheet notes instead of keeping model-specific pricing in code.

    The quote's footer is part of the commercial rule.  Keep its original
    relationship (target foil/copper -> quote-table baseline) so an uploaded
    quote can change the surcharge without a source-code release.
    """
    rule_path = path or get_shennan_rule_file_path()
    excel = pd.ExcelFile(rule_path)
    rules: dict[str, dict] = {}
    for sheet_name in excel.sheet_names:
        clean_name = str(sheet_name).strip()
        if not clean_name.upper().endswith(" CCL") and not _looks_like_ccl_sheet(rule_path, clean_name):
            continue
        glue = clean_name.rsplit(" ", 1)[0]
        df = pd.read_excel(rule_path, sheet_name=clean_name, header=None)
        text = "\n".join(_clean_text(value) for value in df.to_numpy().ravel())
        rule_set = rules.setdefault(glue, {"foil": {}, "copper": {}, "square_board": False})
        rule_set["square_board"] = "方板价格" in text and "加7%" in text.replace(" ", "")
        _parse_foil_note_rules(text, rule_set["foil"])
        _parse_copper_note_rules(text, rule_set["copper"])
    return rules


def _parse_foil_note_rules(text: str, rules: dict[str, dict]) -> None:
    normalized = re.sub(r"\s+", "", text.upper().replace("Ｏ", "O"))

    # Ordinary HTE sheets commonly state this as "指定使用 RTF 上调 3%".
    percent_match = re.search(r"(?:指定)?使用RTF铜箔.*?(?:上调|加)(\d+(?:\.\d+)?)%", normalized)
    if percent_match:
        percent = float(percent_match.group(1)) / 100
        rules["RTF"] = {"base": "HTE", "type": "percent", "percent": percent}
        rules["RTF1"] = {"base": "HTE", "type": "percent", "percent": percent}

    # Examples: "使用 HVLP2 铜箔时，在 RTF3 铜箔的基础上 ...".
    pattern = re.compile(
        r"使用(?P<target>RTF[1-4]?|HVLP[1-3]?)铜箔[^。\n]{0,120}?在(?P<base>RTF[1-4]?|HVLP[1-3]?)铜箔的基础上(?P<body>.*?)(?=(?:。|\n)\d+[、.]|客户回签|$)"
    )
    for match in pattern.finditer(normalized):
        adjustment = _parse_copper_adjustment_body(match.group("body"))
        if adjustment:
            rules[match.group("target")] = {
                "base": match.group("base"),
                "type": "per_sf",
                "values": adjustment,
            }

    # Some sheets omit "在 XXX 铜箔的基础上" but their column header gives
    # a single RTF baseline.  Retain the explicit target and resolve its base
    # against that selected column during calculation.
    for target, body in re.findall(r"(?:若)?使用(RTF[1-4]?|HVLP[1-3]?)铜箔[^。\n]*?(?:加价|增加|加|减价|减)[:：]?(.+?)(?=(?:使用|客户回签|$))", normalized):
        if target in rules:
            continue
        adjustment = _parse_copper_adjustment_body(body)
        if adjustment:
            rules[target] = {"base": "RTF", "type": "per_sf", "values": adjustment}

    # Wording in several sheets: RTF2 / RTF3 are the same price as the
    # quoted foil.  The baseline is resolved later from the selected column.
    if "RTF2铜箔和RTF同价" in normalized:
        rules["RTF2"] = {"base": "RTF", "type": "per_sf", "values": {}}
    if "RTF2和RTF3铜箔与HVLP铜箔同价" in normalized:
        rules.setdefault("RTF2", {"base": "HVLP1", "type": "per_sf", "values": {}})
        rules.setdefault("RTF3", {"base": "HVLP1", "type": "per_sf", "values": {}})


def _parse_copper_note_rules(text: str, rules: dict[str, dict]) -> None:
    normalized = re.sub(r"\s+", "", text.upper().replace("Ｏ", "O"))
    token_patterns = {
        "3/3": r"3/3铜箔在H/H加价([\d.]+)元/SF",
        "S3/S3": r"S3/S3铜箔在H/H加价([\d.]+)元/SF",
        "S4/S4": r"S4/S4铜箔在H/H加价([\d.]+)元/SF",
        "T/T": r"T/T铜箔在H/H加价([\d.]+)元/SF",
        "ST/ST": r"ST/ST铜箔在H/H加价([\d.]+)元/SF",
        "1.5/1.5": r"1\.5/1\.5铜箔2/2减价([\d.]+)元/SF",
        "2.5/2.5": r"2\.5/2\.5铜箔H/H加价([\d.]+)元/SF",
    }
    base_copper = {
        "3/3": "H/H", "S3/S3": "H/H", "S4/S4": "H/H", "T/T": "H/H",
        "ST/ST": "H/H", "1.5/1.5": "2/2", "2.5/2.5": "H/H",
    }
    for token, pattern in token_patterns.items():
        match = re.search(pattern, normalized)
        if not match:
            continue
        amount = float(match.group(1))
        if "减价" in match.group(0):
            amount = -amount
        rules[token] = {"base_copper": base_copper[token], "per_sf": amount}


def _parse_copper_adjustment_body(body: str) -> dict[str, dict[str, float]]:
    values: dict[str, dict[str, float]] = {}
    for source, key in (("HOZ", "H"), ("H", "H"), ("1OZ", "1"), ("2OZ", "2")):
        match = re.search(
            rf"{source}单面(?P<single_action>加|增加|减|减少)(?P<single>[\d.]+)元(?:/SF)?.*?"
            rf"双面(?P<double_action>加|增加|减|减少)(?P<double>[\d.]+)元(?:/SF)?",
            body,
        )
        if not match:
            continue
        single_sign = -1 if "减" in match.group("single_action") else 1
        double_sign = -1 if "减" in match.group("double_action") else 1
        values[key] = {
            "single": single_sign * float(match.group("single")),
            "double": double_sign * float(match.group("double")),
        }
    # Some quote sheets state one amount per published copper-price column,
    # without a single/double split (for example Hoz +2.25, 1OZ +3.75).  The
    # amount is then already the column's complete RMB/SF adjustment.
    for source, key in (("HOZ", "H"), ("1OZ", "1"), ("2OZ", "2")):
        if key in values:
            continue
        match = re.search(rf"{source}(?P<action>加|增加|减|减少)(?P<value>[\d.]+)元(?:/SF)?", body)
        if not match:
            continue
        sign = -1 if "减" in match.group("action") else 1
        values[key] = {"any": sign * float(match.group("value"))}
    return values


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


def _copper_options_from_block_label(label: str) -> list[str]:
    primary = _copper_from_block_label(label)
    if not primary:
        return []
    upper = label.upper().replace(" ", "")
    options = [primary]
    if primary == "1/1" and "H" in upper:
        options.append("H/1")
    if primary == "2/2":
        if "H/2" in upper or "HOZ/2OZ" in upper:
            options.append("H/2")
        if "1/2" in upper or "1OZ/2OZ" in upper:
            options.append("1/2")
    return options


def _foil_from_block_label(label: str) -> str:
    upper = label.upper()
    for foil in ("HVLP3", "HVLP2", "HVLP1", "RTF4", "RTF3", "RTF2", "RTF1", "RTF"):
        if foil in upper:
            return foil
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
