# -*- coding: utf-8 -*-
"""
营销转码自动化工具 v3.0
========================
修复清单：
1. mil→mm 先保留2位小数再×1000
2. 铜箔规格正则：1/1FR-4 → 11（不匹配FR-4）
3. 尺寸提取：过滤叠构数字（4位布号），优先带引号尺寸
4. NY3150（无后缀）→ 3B（同NY3150HF）
5. 结构代码：按PP层数映射（A=1-2层, B=3-4层, C=5-6层, D=7-10层）
6. 基板级别：普通TG → F1（丹凤标布）；NY2140普通TG → F1
7. 铜箔类型：HTE是常规铜，不映射为D；只有RTF/HVLP/VLP才特殊
8. 0.080mm 正则优先匹配显式mm单位

编码结构（基板）：
  第1-2位  : 胶系代码
  第3-7位  : 厚度代码（5位，mm×1000，先round2位小数）
  第8-9位  : 铜箔规格代码
  第10-17位: 尺寸代码（8位）
  第18位   : 胶水类别（Y/R）
  第19位   : 铜箔类型
  第20-21位: 基板级别
  第22位   : 总/芯厚（T/C）
  第23位   : 结构代码
"""

import re
import math
import zipfile
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Tuple, Dict, List

try:
    from .transcode_special_rules import get_structured_special_rules_path, is_structured_special_rules_enabled
except Exception:
    get_structured_special_rules_path = None
    is_structured_special_rules_enabled = None


# ============================================================
# 数据加载
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_RULE_FILENAME = "transcode_rules.xlsx"
DEFAULT_RULE_PATH = BASE_DIR / DEFAULT_RULE_FILENAME
CUSTOMER_ORDER_SHEET_NAME = "客户下单与胶系基板转换"
RULE_SHEET_NAMES = ["胶系代码", "胶系类别", "编码规则", "特殊需求", "总芯厚转换", CUSTOMER_ORDER_SHEET_NAME]
DEMAND_SHEET_NAME = "转码需求表"
SPEC_HEADER_KEYWORDS = ("客户规格", "规格", "物料描述", "描述", "品名规格", "品名")
SPEC_HEADER_EXCLUDE_KEYWORDS = ("编码", "代码", "料号", "编号", "数量", "单位", "日期", "客户代码", "客户名称")
CUSTOMER_HEADER_KEYWORDS = ("客户简称", "客户简码", "客户名称", "客户")
CUSTOMER_HEADER_EXCLUDE_KEYWORDS = ("规格", "编码", "代码", "料号", "编号", "数量", "单位", "日期")
CUSTOMER_CODE_HEADER_KEYWORDS = ("客户编号", "客户代码", "客户编码", "客户代号")
CONTEXT_HEADER_KEYWORDS = (
    "备注",
    "订单备注",
    "订单要求",
    "客户特殊要求",
    "客户特殊需求",
    "客户物料描述",
    "物料描述",
    "材料描述",
    "描述",
)
SPEC_CONTENT_RE = re.compile(
    r"(?:NY[\w\-\(\)]*|\d+\.?\d*\s*(?:mm|mil)|\d+\.?\d*\"?\s*[*xX×]\s*\d+\.?\d*\"?|FR-?4|HALOGEN|无卤|有卤|H\s*/\s*H|\d\s*/\s*\d|RTF|HVLP|HTE)",
    re.IGNORECASE,
)

def load_all_sheets(filepath: str) -> Dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(filepath)
    return {name: pd.read_excel(filepath, sheet_name=name, header=None)
            for name in xl.sheet_names}


def select_demand_sheet_name(sheets: Dict[str, pd.DataFrame]) -> str:
    if DEMAND_SHEET_NAME in sheets:
        return DEMAND_SHEET_NAME
    for name, df in sheets.items():
        if name not in RULE_SHEET_NAMES and not df.empty:
            return name
    for name in sheets:
        return name
    raise ValueError("需求文件中没有可读取的 Sheet")


def load_demand_sheets(filepath: str) -> Dict[str, pd.DataFrame]:
    sheets = load_all_sheets(filepath)
    demand_name = select_demand_sheet_name(sheets)
    return {DEMAND_SHEET_NAME: sheets[demand_name]}


def load_rule_sheets(rule_path: str = None) -> Dict[str, pd.DataFrame]:
    path = Path(rule_path) if rule_path else DEFAULT_RULE_PATH
    if not path.exists():
        raise FileNotFoundError(
            f"未找到转码规则文件：{path}\n"
            f"请将包含 {', '.join(RULE_SHEET_NAMES)} 的规则文件保存为 {DEFAULT_RULE_FILENAME}"
        )
    sheets = load_all_sheets(str(path))
    missing = [name for name in RULE_SHEET_NAMES if name not in sheets]
    if missing:
        raise ValueError(f"规则文件缺少 Sheet：{', '.join(missing)}")
    return {name: sheets[name] for name in RULE_SHEET_NAMES}


def load_transcode_inputs(demand_path: str, rule_path: str = None) -> Tuple[Dict[str, pd.DataFrame], Dict]:
    demand_sheets = load_demand_sheets(demand_path)
    rule_sheets = load_rule_sheets(rule_path)
    return demand_sheets, build_lookup_tables(rule_sheets)


def _clean_cell(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ").strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    if text.lower() in ("nan", "none"):
        return ""
    return text


def _normalize_customer_code(value) -> str:
    return re.sub(r"\s+", "", _clean_cell(value))


def _normalize_customer_name(value) -> str:
    return re.sub(r"\s+", "", _clean_cell(value))


def _looks_like_spec(value) -> bool:
    return bool(SPEC_CONTENT_RE.search(_clean_cell(value)))


def detect_spec_column(df: pd.DataFrame) -> int:
    """Return zero-based column index for customer specification text."""
    if df.empty or len(df.columns) == 0:
        raise ValueError("需求表为空，无法识别规格列")

    header_rows = min(3, len(df))
    for row_idx in range(header_rows):
        for col_idx in range(len(df.columns)):
            header = _clean_cell(df.iloc[row_idx, col_idx])
            if (
                any(keyword in header for keyword in SPEC_HEADER_KEYWORDS)
                and not any(keyword in header for keyword in SPEC_HEADER_EXCLUDE_KEYWORDS)
            ):
                return col_idx

    best_col = 18 if len(df.columns) > 18 else 0
    best_score = -1
    for col_idx in range(len(df.columns)):
        score = sum(1 for row_idx in range(1, min(len(df), 80)) if _looks_like_spec(df.iloc[row_idx, col_idx]))
        if score > best_score:
            best_score = score
            best_col = col_idx
    return best_col


def detect_customer_spec_column(df: pd.DataFrame) -> Optional[int]:
    """Return the explicit 客户规格 column when present."""
    if df.empty or len(df.columns) == 0:
        return None
    header_rows = min(3, len(df))
    for row_idx in range(header_rows):
        for col_idx in range(len(df.columns)):
            header = _clean_cell(df.iloc[row_idx, col_idx])
            if '客户规格' in header:
                return col_idx
    return None


def select_transcode_spec_column(df: pd.DataFrame) -> int:
    """Official transcode input is 客户规格; fallback to auto detection only when absent."""
    customer_spec_col = detect_customer_spec_column(df)
    if customer_spec_col is not None:
        return customer_spec_col
    return detect_spec_column(df)


def detect_customer_column(df: pd.DataFrame, spec_col: int = None) -> Optional[int]:
    """Return zero-based customer short-name column index if present."""
    if df.empty or len(df.columns) == 0:
        return None
    header_rows = min(3, len(df))
    for row_idx in range(header_rows):
        for col_idx in range(len(df.columns)):
            if spec_col is not None and col_idx == spec_col:
                continue
            header = _clean_cell(df.iloc[row_idx, col_idx])
            if (
                any(keyword in header for keyword in CUSTOMER_HEADER_KEYWORDS)
                and not any(keyword in header for keyword in CUSTOMER_HEADER_EXCLUDE_KEYWORDS)
            ):
                return col_idx
    if len(df.columns) > 3 and spec_col != 3:
        return 3
    return None


def detect_customer_code_column(df: pd.DataFrame) -> Optional[int]:
    """Return zero-based customer code column index if present."""
    if df.empty or len(df.columns) == 0:
        return None
    header_rows = min(3, len(df))
    for row_idx in range(header_rows):
        for col_idx in range(len(df.columns)):
            header = _clean_cell(df.iloc[row_idx, col_idx])
            if any(keyword in header for keyword in CUSTOMER_CODE_HEADER_KEYWORDS):
                return col_idx
    if len(df.columns) > 1:
        return 1
    return None


def detect_transcode_context_columns(df: pd.DataFrame, *exclude_cols: int | None) -> list[int]:
    """Return columns that contain remark/order/material-description context for special rules."""
    if df.empty or len(df.columns) == 0:
        return []
    excluded = {col for col in exclude_cols if col is not None}
    context_cols: list[int] = []
    header_rows = min(3, len(df))
    for row_idx in range(header_rows):
        for col_idx in range(len(df.columns)):
            if col_idx in excluded:
                continue
            header = _clean_cell(df.iloc[row_idx, col_idx])
            if header and any(keyword in header for keyword in CONTEXT_HEADER_KEYWORDS):
                if col_idx not in context_cols:
                    context_cols.append(col_idx)
    return context_cols


def build_context_text_from_row(row, context_cols: list[int]) -> str:
    values = []
    for col_idx in context_cols:
        if len(row) <= col_idx:
            continue
        value = _clean_cell(row.iloc[col_idx])
        if value:
            values.append(value)
    return " ".join(values)


def detect_append_result_column(df: pd.DataFrame) -> int:
    """Return the first blank column after the rightmost column containing any data."""
    if df.empty or len(df.columns) == 0:
        return 0
    last_used = -1
    for col_idx in range(len(df.columns)):
        has_value = any(_clean_cell(df.iloc[row_idx, col_idx]) for row_idx in range(len(df)))
        if has_value:
            last_used = col_idx
    return last_used + 1


def ensure_result_column(df: pd.DataFrame, preferred_col: int = None) -> Tuple[pd.DataFrame, int]:
    """Ensure output column exists and can hold text values."""
    df = df.copy()
    if preferred_col is None:
        preferred_col = detect_append_result_column(df)
    while len(df.columns) <= preferred_col:
        df[len(df.columns)] = ""
    df.iloc[:, preferred_col] = df.iloc[:, preferred_col].astype(object)
    return df, preferred_col


def build_lookup_tables(sheets: Dict[str, pd.DataFrame]) -> Dict:
    tables = {}

    # ── 胶系代码表 ──
    df_glue = sheets.get("胶系代码", pd.DataFrame())
    glue_exact_map = {}   # 完整名称(大写) → 代码
    glue_model_map = {}   # 基础型号(大写) → 代码（第一个匹配）
    for i in range(1, len(df_glue)):
        b = str(df_glue.iloc[i, 1]).strip() if pd.notna(df_glue.iloc[i, 1]) else ''
        d = str(df_glue.iloc[i, 3]).strip() if pd.notna(df_glue.iloc[i, 3]) else ''
        if not b or not d or b == 'nan' or d == 'nan':
            continue
        glue_exact_map[b.upper()] = d
        base = b.split()[0].upper()
        if base not in glue_model_map:
            glue_model_map[base] = d
    tables['glue_exact_map'] = glue_exact_map
    tables['glue_model_map'] = glue_model_map

    # ── 胶系类别表 ──
    df_cat = sheets.get("胶系类别", pd.DataFrame())
    glue_cat_map = {}
    for i in range(2, len(df_cat)):
        for col in [0, 1]:
            a = str(df_cat.iloc[i, col]).strip() if pd.notna(df_cat.iloc[i, col]) else ''
            c = str(df_cat.iloc[i, 2]).strip() if pd.notna(df_cat.iloc[i, 2]) else ''
            if a and c and a != 'nan' and c != 'nan':
                glue_cat_map[a.upper()] = c
    tables['glue_cat_map'] = glue_cat_map

    # ── 总芯厚转换表 ──
    df_thick = sheets.get("总芯厚转换", pd.DataFrame())
    thick_total_to_core = {}
    thick_core_to_total = {}
    for i in range(1, len(df_thick)):
        d = str(df_thick.iloc[i, 3]).strip() if len(df_thick.columns) > 3 and pd.notna(df_thick.iloc[i, 3]) else ''
        f = df_thick.iloc[i, 5] if len(df_thick.columns) > 5 and pd.notna(df_thick.iloc[i, 5]) else None
        g = df_thick.iloc[i, 6] if len(df_thick.columns) > 6 and pd.notna(df_thick.iloc[i, 6]) else None
        if d and d != 'nan':
            if f is not None:
                try:
                    thick_total_to_core[d.upper()] = float(f)
                except:
                    pass
            if g is not None:
                try:
                    thick_core_to_total[d.upper()] = float(g)
                except:
                    pass
    tables['thick_total_to_core'] = thick_total_to_core
    tables['thick_core_to_total'] = thick_core_to_total

    # ── 特殊需求表 ──
    df_sp = sheets.get("特殊需求", pd.DataFrame())
    special_by_name = {}
    special_by_code = {}
    for i in range(1, len(df_sp)):
        code = str(df_sp.iloc[i, 0]).strip() if pd.notna(df_sp.iloc[i, 0]) else ''
        name = str(df_sp.iloc[i, 2]).strip() if pd.notna(df_sp.iloc[i, 2]) else ''
        mat  = str(df_sp.iloc[i, 3]).strip() if pd.notna(df_sp.iloc[i, 3]) else ''
        cat  = str(df_sp.iloc[i, 5]).strip() if pd.notna(df_sp.iloc[i, 5]) else ''
        req  = str(df_sp.iloc[i, 6]).strip() if pd.notna(df_sp.iloc[i, 6]) else ''
        entry = {'mat': mat, 'cat': cat, 'req': req}
        if name and name != 'nan':
            special_by_name.setdefault(name, []).append(entry)
        if code and code != 'nan':
            special_by_code.setdefault(code, []).append(entry)
    tables['special_by_name'] = special_by_name
    tables['special_by_code'] = special_by_code

    # ── 编码规则表 ──
    df_rule = sheets.get("编码规则", pd.DataFrame())
    grade_code_map = {}    # K列代码 → L列说明（用于反查）
    grade_desc_to_code = {}  # 说明关键词 → K列代码
    struct_desc_to_code = {}
    for i in range(len(df_rule)):
        k = str(df_rule.iloc[i, 10]).strip() if len(df_rule.columns) > 10 and pd.notna(df_rule.iloc[i, 10]) else ''
        l = str(df_rule.iloc[i, 11]).strip() if len(df_rule.columns) > 11 and pd.notna(df_rule.iloc[i, 11]) else ''
        if k and l and k != 'nan' and l != 'nan' and len(k) <= 4:
            grade_code_map[k] = l
            grade_desc_to_code[l] = k
        n = str(df_rule.iloc[i, 13]).strip() if len(df_rule.columns) > 13 and pd.notna(df_rule.iloc[i, 13]) else ''
        o = str(df_rule.iloc[i, 14]).strip() if len(df_rule.columns) > 14 and pd.notna(df_rule.iloc[i, 14]) else ''
        if n and n != 'nan' and len(n) <= 2:
            struct_desc_to_code[o] = n
    tables['grade_code_map'] = grade_code_map
    tables['grade_desc_to_code'] = grade_desc_to_code
    tables['struct_desc_to_code'] = struct_desc_to_code

    # ── 客户下单与胶系基板转换 ──
    df_customer_order = sheets.get(CUSTOMER_ORDER_SHEET_NAME, pd.DataFrame())
    customer_order_rules = []
    for i in range(1, len(df_customer_order)):
        cust_code = _normalize_customer_code(df_customer_order.iloc[i, 0] if len(df_customer_order.columns) > 0 else "")
        cust_name = _normalize_customer_name(df_customer_order.iloc[i, 1] if len(df_customer_order.columns) > 1 else "")
        customer_glue = _clean_cell(df_customer_order.iloc[i, 2] if len(df_customer_order.columns) > 2 else "")
        glue_code = _clean_cell(df_customer_order.iloc[i, 3] if len(df_customer_order.columns) > 3 else "")
        grade_code = _clean_cell(df_customer_order.iloc[i, 4] if len(df_customer_order.columns) > 4 else "")
        structure_code = _clean_cell(df_customer_order.iloc[i, 5] if len(df_customer_order.columns) > 5 else "")
        if not cust_code or not cust_name or not customer_glue:
            continue
        parsed = _parse_customer_order_glue_rule(customer_glue)
        customer_order_rules.append({
            'row': i + 1,
            'cust_code': cust_code,
            'cust_name': cust_name,
            'raw_glue': customer_glue,
            'glue_key': parsed['glue_key'],
            'unsupported': parsed['unsupported'],
            'condition_op': parsed['condition_op'],
            'condition_mm': parsed['condition_mm'],
            'glue_code': glue_code,
            'grade_code': grade_code,
            'structure_code': structure_code,
        })
    tables['customer_order_rules'] = customer_order_rules
    tables['structured_special_rules'] = load_structured_special_rules()

    return tables


def load_structured_special_rules() -> list[dict]:
    """Load confirmed structured special rules saved by the web helper."""
    if get_structured_special_rules_path is None:
        return []
    if is_structured_special_rules_enabled is not None:
        try:
            if not is_structured_special_rules_enabled():
                return []
        except Exception:
            return []
    try:
        path = get_structured_special_rules_path()
    except Exception:
        return []
    if not path or not Path(path).exists():
        return []
    try:
        df = pd.read_excel(path, sheet_name="结构化特殊规则")
    except Exception:
        return []
    rules: list[dict] = []
    for _, row in df.iterrows():
        item = {str(col).strip(): _clean_cell(row[col]) for col in df.columns}
        if not item.get("规则ID") and not item.get("规则大类"):
            continue
        rules.append(item)
    return rules


# ============================================================
# 步骤1：胶系代码
# ============================================================

KNOWN_GLUE_MODELS = sorted([
    'NY3150HF IST改善压合窗口', 'NY3150HF IST改善', 'NY3150HF FR15',
    'NYHP-7350 LNB330', 'NYHP-7350D', 'NYHP-7300',
    'NYHP-7350', 'NY6300SL', 'NY6300SP', 'NY6300SN', 'NY6300S',
    'NY3188HF', 'NY3176HF', 'NY3170HF', 'NY3170M2', 'NY3170LK',
    'NY3170HC', 'NY3170M', 'NY3150HC', 'NY3150HF', 'NY3150',
    'NY6666SE', 'NY6666S', 'NY6666N', 'NY6666', 'NY6288',
    'NY6200M', 'NY6200', 'NY6180HF', 'NY6180LL', 'NY6180L', 'NY6180',
    'NY6600', 'NY6300', 'NY2600', 'NY2170H', 'NY2170',
    'NY2150H', 'NY2150M', 'NY2150', 'NY2140L', 'NY2140',
    'NY2110', 'NY1600', 'NY1140',
    'NY8888Q', 'NY8888N', 'NY8888', 'NY9999Q', 'NY9999N', 'NY9999',
    'NY8320', 'NY-P5P', 'NY-P5Q', 'NY-P5', 'NY-P4P', 'NY-P4N', 'NY-P4',
    'NY-P3P', 'NY-P3', 'NY-P2P', 'NY-P2H', 'NY-P2',
    'NY-P1P', 'NY-P1', 'NY-A3HF', 'NY-A2', 'NY-A1',
], key=len, reverse=True)

# 特殊映射：客户常省略后缀的型号 → 标准代码
GLUE_ALIAS = {
    'NY2150': '2B',
    'NY2150H': '2H',
    'NY6300S': '6C',
    'NY3150': '3B',   # 客户写NY3150通常指NY3150HF
    'NY3170': '3C',   # 客户写NY3170通常指NY3170HF
}


def _normalize_glue_key(value: str) -> str:
    """Normalize customer glue text for tolerant rule-table matching."""
    if not value:
        return ""
    key = str(value).upper().strip()
    key = key.replace("（", "(").replace("）", ")").replace("\xa0", " ")
    key = re.sub(r"\s+", "", key)
    key = re.sub(r"^(?:CCL|FR-?4)[-_]?", "", key)
    key = re.sub(r"[-_]?板厚.*$", "", key)
    key = re.sub(r"[-_]?FR-?4(?:[_-]?\d+)?$", "", key)
    key = re.sub(r"[-_]UL$", "", key)
    key = re.sub(r"[-_]RC.*$", "", key)
    if len(re.findall(r"[-_]", key)) >= 2:
        key = re.sub(r"[-_]\d+(?:\.\d+)?$", "", key)
    key = key.replace("-", "").replace("_", "")
    return key


def _parse_customer_order_glue_rule(value: str) -> dict:
    """Parse customer order glue rules; support exact glue plus optional board-thickness condition."""
    raw = _clean_cell(value)
    condition_op = None
    condition_mm = None
    base = raw
    m = re.search(r'(.+?)[\(（]\s*板厚\s*(<=|>=|<|>|≤|≥)\s*(\d+(?:\.\d+)?)\s*(?:MM|mm)?\s*[\)）]', raw)
    if m:
        base = m.group(1).strip()
        condition_op = m.group(2)
        condition_mm = float(m.group(3))
    unsupported = '/' in base or '／' in base
    return {
        'glue_key': _normalize_glue_key(base),
        'unsupported': unsupported,
        'condition_op': condition_op,
        'condition_mm': condition_mm,
    }


def _condition_matches(value: Optional[float], op: Optional[str], threshold: Optional[float]) -> bool:
    if op is None or threshold is None:
        return True
    if value is None:
        return False
    if op in ('<',):
        return value < threshold
    if op in ('<=', '≤'):
        return value <= threshold
    if op in ('>',):
        return value > threshold
    if op in ('>=', '≥'):
        return value >= threshold
    return False


def find_customer_order_override(cust_code: str, cust_name: str, glue_model: str,
                                 raw_thickness_mm: Optional[float], tables: dict) -> dict:
    rules = tables.get('customer_order_rules', [])
    code_key = _normalize_customer_code(cust_code)
    name_key = _normalize_customer_name(cust_name)
    glue_key = _normalize_glue_key(glue_model)
    if not code_key or not name_key or not glue_key:
        return {}

    matches = []
    for rule in rules:
        if rule.get('unsupported'):
            continue
        if rule.get('cust_code') != code_key or rule.get('cust_name') != name_key:
            continue
        if rule.get('glue_key') != glue_key:
            continue
        if not _condition_matches(raw_thickness_mm, rule.get('condition_op'), rule.get('condition_mm')):
            continue
        matches.append(rule)

    if not matches:
        return {}

    signatures = {
        (m.get('glue_code', ''), m.get('grade_code', ''), m.get('condition_op'), m.get('condition_mm'))
        for m in matches
    }
    if len(signatures) > 1:
        rows = ','.join(str(m.get('row')) for m in matches)
        return {'error': f'客户下单胶系基板转换规则重复：客户{cust_code}/{cust_name}，胶系{glue_model}，规则行{rows}'}

    match = matches[0]
    return {
        'row': match.get('row'),
        'raw_glue': match.get('raw_glue', ''),
        'glue_code': match.get('glue_code', ''),
        'grade_code': match.get('grade_code', ''),
    }


def _strict_glue_key(value: str) -> str:
    """Normalize text without dropping qualifier words such as UL配方 or 汽车板."""
    if not value:
        return ""
    key = str(value).upper().strip()
    key = key.replace("（", "(").replace("）", ")").replace("\xa0", " ")
    key = re.sub(r"\s+", "", key)
    key = re.sub(r"^(?:CCL|FR-?4)[-_]?", "", key)
    return key.replace("-", "").replace("_", "")


def _is_ascii_alnum(char: str) -> bool:
    return bool(char) and char.isascii() and char.isalnum()


def _is_ascii_alpha(char: str) -> bool:
    return bool(char) and char.isascii() and char.isalpha()


def _normalized_glue_key_matches(norm_key: str, norm_text: str) -> bool:
    """Match normalized NY model text without letting NY2140L hit NY2140 Lastra."""
    if not norm_key or not norm_text:
        return False
    start = 0
    while True:
        idx = norm_text.find(norm_key, start)
        if idx < 0:
            return False
        before = norm_text[idx - 1] if idx > 0 else ""
        after_idx = idx + len(norm_key)
        after = norm_text[after_idx] if after_idx < len(norm_text) else ""
        if not _is_ascii_alnum(before) and not _is_ascii_alpha(after):
            return True
        start = idx + 1


def _find_rule_name(rule_map: dict, target: str) -> Optional[str]:
    target_key = _strict_glue_key(target)
    for name in rule_map.keys():
        if _strict_glue_key(name) == target_key:
            return name
    return None


def _resolve_ny3150hc_rule(text: str, glue_exact_map: dict) -> Optional[str]:
    """NY3150HC has business-specific default and conditional variants."""
    text_upper = str(text or "").upper()
    if not re.search(r"\bNY-?3150HC\b", text_upper):
        return None

    if any(keyword in text for keyword in ("汽车专用", "汽车板", "车用")):
        return _find_rule_name(glue_exact_map, "NY3150HC 汽车板")

    has_ul = "UL" in text_upper or "UL配方" in text
    has_3mm = bool(re.search(r"(?:板厚\s*)?3(?:\.0+)?\s*MM|板厚\s*3(?:\.0+)?", text_upper, re.IGNORECASE))
    if has_ul and has_3mm:
        return _find_rule_name(glue_exact_map, "NY3150HC 板厚3.0MMUL配方")
    if has_ul:
        return _find_rule_name(glue_exact_map, "NY3150HC UL配方")

    # Plain NY3150HC uses the exact NY3150HC rule. Variants above remain explicit.
    return _find_rule_name(glue_exact_map, "NY3150HC")


def _glue_rule_matches_text(rule_name: str, text: str) -> bool:
    norm_text = _strict_glue_key(text)
    norm_name = _strict_glue_key(rule_name)
    if _normalized_glue_key_matches(norm_name, norm_text):
        return True

    base_match = re.match(r"(NY[A-Z0-9()\-]+)", str(rule_name).upper().strip())
    if not base_match:
        return False
    base = _strict_glue_key(base_match.group(1))
    if not _normalized_glue_key_matches(base, norm_text):
        return False

    qualifier = norm_name[len(base):]
    if not qualifier:
        return True
    if "汽车" in qualifier:
        prefix = qualifier.split("汽车", 1)[0]
        if prefix and prefix not in norm_text:
            return False
        return any(keyword in text for keyword in ("汽车专用", "汽车板", "车用"))
    if "UL配方" in qualifier:
        has_ul = "UL" in str(text).upper()
        if "板厚30MM" in qualifier or "板厚3.0MM" in str(rule_name).upper():
            return has_ul and bool(re.search(r"(?:板厚\s*)?3(?:\.0+)?\s*MM|板厚\s*3(?:\.0+)?", str(text).upper()))
        return has_ul
    return qualifier in norm_text


def _glue_candidates(glue_model: str) -> List[str]:
    """Generate possible glue spellings seen in customer descriptions."""
    raw = str(glue_model or "").upper().strip()
    tokens = [raw, raw.split()[0] if raw else ""]
    candidates = []
    for token in tokens:
        if not token:
            continue
        normalized = _normalize_glue_key(token)
        variants = [token, normalized]
        if normalized.endswith("P"):
            variants.append(normalized[:-1])
        if token.endswith("P"):
            variants.append(token[:-1])
        for item in variants:
            if item and item not in candidates:
                candidates.append(item)
    return candidates


def extract_glue_model(text: str) -> Optional[str]:
    """从规格文本中提取胶系型号"""
    text_upper = text.upper()
    for model in KNOWN_GLUE_MODELS:
        if model.upper() in text_upper:
            return model
    m = re.search(r'(NY[\w\-]+)', text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None


def extract_glue_model_from_rules(text: str, glue_exact_map: dict) -> Optional[str]:
    """Prefer the longest rule-table glue name that appears in the customer text."""
    customer_7402 = re.search(r"(?:TC|DS)[-_]?7402(LC)?", str(text or ""), re.IGNORECASE)
    if customer_7402:
        suffix = "7402LC" if customer_7402.group(1) else "7402"
        candidates = [
            name
            for name in glue_exact_map
            if suffix in str(name).upper()
            and (suffix.endswith("LC") or "7402LC" not in str(name).upper())
        ]
        if candidates:
            return max(candidates, key=len)
    ny3150hc_rule = _resolve_ny3150hc_rule(text, glue_exact_map)
    if ny3150hc_rule:
        return ny3150hc_rule

    matched = []
    for name in glue_exact_map.keys():
        if _glue_rule_matches_text(name, text):
            matched.append((len(_strict_glue_key(name)), name))
    if matched:
        matched.sort(reverse=True)
        return matched[0][1]
    glue_model = extract_glue_model(text)
    if glue_model:
        return glue_model
    desc_model = _infer_glue_model_from_description(text)
    if desc_model:
        return desc_model
    return None


def _infer_glue_model_from_description(text: str) -> Optional[str]:
    """Fallback only when no explicit NY model is recognized."""
    text_raw = str(text or "")
    text_upper = text_raw.upper()
    has_halogen_free = '无卤' in text_raw or 'HALOGEN FREE' in text_upper
    has_htg = 'HTG' in text_upper
    has_mtg = 'MTG' in text_upper
    has_tg150 = bool(re.search(r'TG\s*[≥>=:]?\s*150', text_upper))
    has_tg170 = bool(re.search(r'TG\s*[≥>=:]?\s*170', text_upper))

    if has_htg:
        return 'NY3170' if has_halogen_free else 'NY2170'
    if has_mtg:
        return 'NY3150HF' if has_halogen_free else 'NY2150'
    if has_tg150:
        return 'NY3150HF' if has_halogen_free else 'NY2150'
    if has_tg170:
        return 'NY2170'
    return None


def get_glue_code(glue_model: str, glue_exact_map: dict, glue_model_map: dict, cust_name: str = "") -> Optional[str]:
    """查询胶系代码"""
    if not glue_model:
        return None
    norm_exact_map = {_normalize_glue_key(k): v for k, v in glue_exact_map.items()}
    norm_model_map = {_normalize_glue_key(k): v for k, v in glue_model_map.items()}
    cust_text = str(cust_name or "")

    for key in _glue_candidates(glue_model):
        norm_key = _normalize_glue_key(key)
        if "深圳安比" in cust_text and norm_key == "NY2140":
            return "2A"
        if any(name in cust_text for name in ("深万基隆", "珠海益天")) and norm_key.startswith("NY2140"):
            return "2A"
        if ("湖奥士康" in cust_text or "景旺" in cust_text) and norm_key == "NYA1":
            return "RC"
        if any(name in cust_text for name in ("广华升鑫", "深华升鑫")) and norm_key == "NYA2":
            return "AL"
        if norm_key == "NY2150H" and "中富" in cust_text:
            return "2W"
        # 1. 别名映射（客户常见非标准写法优先）
        if key in GLUE_ALIAS:
            return GLUE_ALIAS[key]
        if norm_key in GLUE_ALIAS:
            return GLUE_ALIAS[norm_key]
        # 2. 精确匹配
        if key in glue_exact_map:
            return glue_exact_map[key]
        if norm_key in norm_exact_map:
            return norm_exact_map[norm_key]
        # 3. 基础型号匹配
        base = key.split()[0]
        norm_base = _normalize_glue_key(base)
        if base in glue_model_map:
            return glue_model_map[base]
        if norm_base in norm_model_map:
            return norm_model_map[norm_base]
        if base in GLUE_ALIAS:
            return GLUE_ALIAS[base]
        if norm_base in GLUE_ALIAS:
            return GLUE_ALIAS[norm_base]
        # 4. 包含关系匹配
        for k, v in glue_exact_map.items():
            norm_k = _normalize_glue_key(k)
            if norm_key and (norm_key in norm_k or norm_k in norm_key):
                return v
    return None


# ============================================================
# 步骤2：厚度代码
# ============================================================

def extract_thickness_mm(text: str) -> Tuple[Optional[float], str, str]:
    """
    提取厚度，返回 (mm值, 原始字符串, 单位)
    修复：优先匹配显式mm/MM单位，mil转换时先round2位小数
    """
    text = re.sub(r'(?<=\d),(?=\d)', '.', str(text or ''))

    # Explicit inner/core thickness overrides the outer nominal thickness:
    # 0.184mm (not including copper 0.114mm, 2116*1).
    m = re.search(r'[\(（]\s*(?:不含铜|芯厚)\s*[:：]?\s*(\d+(?:\.\d+)?)\s*(?:mm)?', text, re.IGNORECASE)
    if m:
        val = float(m.group(1))
        if 0.03 <= val <= 10:
            return val, m.group(0), 'mm'

    # Customer ERP format: CCL_NY3170M_3.5_C_...; the value before _C_ is mil.
    m = re.search(r'(?<![A-Z0-9])CCL_[^\s_]+_(\d+(?:\.\d+)?)_C_', text, re.IGNORECASE)
    if m:
        mil = float(m.group(1))
        # ERP thin cores through 6mil retain 0.001mm (6mil -> 0.152mm);
        # regular cores follow the established 0.01mm order convention.
        mm = round(mil * 0.0254, 3 if mil <= 6 else 2)
        return mm, m.group(1), 'mil'

    # FR-4/FR4 is a material marker and may be glued to the thickness:
    # FR-40.9mm -> FR-4 0.9mm, FR41.2mm -> FR4 1.2mm.
    text = re.sub(r'\b(FR\s*-?\s*4)(?=\d+\.\d+\s*[Mm]{2})', r'\1 ', text, flags=re.IGNORECASE)

    # 0. 带公差写法取中心值，包括 0.508+0.02/-0.05mm。
    m = re.search(
        r'(\d+(?:\.\d+)?)\s*\+\s*\d+(?:\.\d+)?\s*/\s*-\s*\d+(?:\.\d+)?\s*mm',
        text,
        re.IGNORECASE,
    )
    if m:
        return float(m.group(1)), m.group(0), 'mm'

    # 0.1 常规对称公差：56±3mil、0.200±0.025mm。
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:±|\+/-|＋/－)\s*\d+(?:\.\d+)?\s*(mil|mm)', text, re.IGNORECASE)
    if m:
        val = float(m.group(1))
        unit = m.group(2).lower()
        if unit == 'mil':
            return round(val * 0.0254, 2), m.group(0), 'mil'
        return val, m.group(0), 'mm'

    # 0.2 客户常同时写 mm 与 inch，以括号前的主值为准：1.00(0.039")。
    m = re.search(r'(?<![\d.])(\d+(?:\.\d+)?)\s*[\(（]\s*\d+(?:\.\d+)?\s*(?:"|″|”)\s*[\)）]', text)
    if m:
        val = float(m.group(1))
        if 0.03 <= val <= 10:
            return val, m.group(0), 'mm'

    # 1. 显式mm/MM单位（优先，避免被独立小数匹配干扰）
    m = re.search(r'(?<![\d.])(\d+(?:\.\d+)?)\s*[Mm]{2}(?![Mm])', text)
    if m:
        val = float(m.group(1))
        if 0.03 <= val <= 10:
            return val, m.group(0), 'mm'

    # 2. 小数英寸厚度：优先于后置的公差mil，避免0.0045" ... +/-0.5mil取到0.5mil。
    m = re.search(r'(?<![\d.])(\d+\.\d+)\s*(?:"|″|”)(?!\s*[*Xx×])', text)
    if m:
        inch = float(m.group(1))
        if 0 < inch <= 0.25:
            return round(inch * 25.4, 2), m.group(0), 'inch'

    # 3. mil单位：常规下单口径保留0.01mm。
    m = re.search(r'(\d+(?:\.\d+)?)\s*mil', text, re.IGNORECASE)
    if m:
        mil = float(m.group(1))
        return round(mil * 0.0254, 2), m.group(0), 'mil'

    # 3.5 FR4 followed by a bare integer before copper is a thickness in mm:
    # FR4 1 H/H -> 1mm, FR-4 2 1/1 -> 2mm.
    m = re.search(r'\bFR\s*-?\s*4\s+(\d+(?:\.\d+)?)\s+(?=[A-Z0-9.]{1,3}\s*/)', text, re.IGNORECASE)
    if m:
        val = float(m.group(1))
        if 0.05 <= val <= 10:
            return val, m.group(1), 'mm'

    # 4. 数字(含铜) 或 数字(不含铜) 格式
    m = re.search(r'(\d+\.\d+)\s*[\(（]\s*(?:含铜(?:厚(?:度)?)?|不含铜)\s*[\)）]', text)
    if m:
        return float(m.group(1)), m.group(0), 'mm'

    # Missing delimiter between a one-decimal thickness and micron copper:
    # 1.515/15 means 1.5mm + 15/15um, not 1.515mm.
    m = re.search(r'(?<![\d.])(\d+\.\d)(?=\d{2}\s*/\s*\d{2}(?!\d))', text)
    if m:
        val = float(m.group(1))
        if 0.05 <= val <= 3.5:
            return val, m.group(1), 'mm'

    # 5. 独立小数（去掉叠构括号后再匹配）
    text_no_struct = re.sub(r'[\(（][^)）]*[\)）]', ' ', text)
    # 去掉尺寸部分（W*H格式）
    text_no_size = re.sub(r'\d+(?:\.\d+)?\s*["\u201d]?\s*[*×]\s*\d+(?:\.\d+)?\s*["\u201d]?', '', text_no_struct)
    m = re.search(r'(?<![*×x\d\.\-])(\d+\.\d+)(?!\s*[*×x"\'英寸])', text_no_size)
    if m:
        val = float(m.group(1))
        if 0.05 <= val <= 3.5:
            return val, m.group(0), 'mm'

    return None, '', ''


def thickness_to_code(mm_val: float) -> str:
    """厚度mm值转5位代码：mm × 1000 取整，5位补零"""
    val_int = round(mm_val * 1000)
    return f"{val_int:05d}"


def get_thickness_mode(text: str) -> str:
    """判断厚度模式：含铜=总厚，不含铜=芯厚
    修复：无标注时不能默认总厚，需结合厚度值判断
    """
    text_upper = str(text or '').upper()
    if any(term in text_upper for term in (
        '不含铜', '不连铜', '不連銅', '芯厚',
        'CORE', 'EXCLUDING COPPER', 'WITHOUT COPPER', 'NO COPPER',
    )):
        return 'core'
    if any(term in text_upper for term in ('含铜', '總厚', '总厚', 'OVERALL', 'TOTAL')):
        return 'total'
    if '芯板' in text_upper:
        return 'core'
    return 'unknown'  # 无标注，由calc_order_thickness根据厚度值判断


def get_customer_thickness_mode_override(cust_name: str, text: str) -> Tuple[Optional[str], str]:
    """Customer-only total/core notation that should not become global parsing."""
    cust_text = str(cust_name or "")
    text_upper = str(text or "").upper()
    if "兴森快捷" in cust_text:
        if re.search(r'(?<![A-Z0-9])T\s*/\s*C(?![A-Z0-9])', text_upper):
            return "core", "兴森快捷：T/C 是芯厚"
        if re.search(r'(?<![A-Z0-9])D\s*/\s*C(?![A-Z0-9])', text_upper):
            return "total", "兴森快捷：D/C 是总厚"
    return None, ""


def get_special_thickness_mode(cust_name: str, cust_code: str,
                               special_by_name: dict, special_by_code: dict) -> Tuple[Optional[str], str]:
    """Extract a narrow default thickness mode from customer special requirements."""
    if "深圳安比" in str(cust_name or ""):
        return "total", "深圳安比：含铜为总厚且<0.8时按总芯厚转换表执行"

    entries = []
    if cust_code:
        entries.extend(special_by_code.get(str(cust_code).strip(), []))
    if cust_name:
        entries.extend(special_by_name.get(str(cust_name).strip(), []))

    default_markers = ('未备注', '未注明', '未标注', '没有备注', '没备注', '无备注', '默认')
    total_terms = ('含铜厚度', '总厚', '含铜')
    core_terms = ('不含铜', '芯厚')
    seen = set()
    for entry in entries:
        req = str(entry.get('req', '') or '').strip()
        if not req or req in seen:
            continue
        seen.add(req)
        req_norm = re.sub(r'\s+', '', req)
        total_patterns = (
            '订单全部是总厚', '订单默认总厚', '默认总厚', '默认含铜', '客户默认含铜',
            '未备注都是含铜厚度', '未注明都是含铜厚度', '未标注都是含铜厚度',
            '不含铜会备注', '不含铜会额外备注', '不含铜另备注',
            '芯厚会备注', '芯厚客户订单会备注'
        )
        core_patterns = (
            '默认不含铜', '默认芯厚', '未备注为芯厚', '未注明为芯厚',
            '未标注为芯厚', '未备注都是芯厚', '未注明都是芯厚'
        )
        if any(pattern in req_norm for pattern in total_patterns):
            return 'total', req
        if any(pattern in req_norm for pattern in core_patterns):
            return 'core', req
        parts = re.split(r'[，,；;。]', req)
        for part in parts:
            part = part.strip()
            if not any(marker in part for marker in default_markers):
                continue
            if any(term in part for term in total_terms) and '不含铜' not in part:
                return 'total', req
            if any(term in part for term in core_terms) and not any(term in part for term in ('含铜厚度', '总厚')):
                return 'core', req
    return None, ''


JD_CY_CUSTOMER_KEYWORDS = ('健鼎', '超颖', '超颍')
_JD_CY_THICKNESS_CACHE = None


def is_jd_cy_customer(cust_name: str) -> bool:
    return bool(cust_name) and any(keyword in str(cust_name) for keyword in JD_CY_CUSTOMER_KEYWORDS)


def _normalize_special_thickness_key(value: str) -> str:
    key = str(value or '').upper().replace('（', '(').replace('）', ')')
    return re.sub(r'\s+', '', key)


def _parse_mm_cell(value) -> Optional[float]:
    text = str(value or '').strip()
    m = re.search(r'\d+(?:\.\d+)?', text)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def _excel_col_index(cell_ref: str) -> int:
    letters = re.match(r'[A-Z]+', str(cell_ref or '').upper())
    if not letters:
        return 0
    idx = 0
    for ch in letters.group(0):
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _read_xlsx_rows_compatible(path: Path) -> List[List[str]]:
    """Read xlsx rows even when internal zip members use backslashes."""
    rows: List[List[str]] = []
    ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    with zipfile.ZipFile(path) as zf:
        name_map = {name.replace('\\', '/'): name for name in zf.namelist()}
        shared_strings: List[str] = []
        if 'xl/sharedStrings.xml' in name_map:
            shared_root = ET.fromstring(zf.read(name_map['xl/sharedStrings.xml']))
            for si in shared_root.findall('a:si', ns):
                shared_strings.append(''.join(t.text or '' for t in si.findall('.//a:t', ns)))

        workbook_root = ET.fromstring(zf.read(name_map['xl/workbook.xml']))
        rel_root = ET.fromstring(zf.read(name_map['xl/_rels/workbook.xml.rels']))
        rels = {rel.attrib.get('Id'): rel.attrib.get('Target', '') for rel in rel_root}

        for sheet in workbook_root.findall('a:sheets/a:sheet', ns):
            sheet_name = sheet.attrib.get('name', '')
            if '板厚' not in sheet_name and '换算' not in sheet_name:
                continue
            rel_id = sheet.attrib.get(f'{{{rel_ns}}}id')
            target = rels.get(rel_id, '').lstrip('/').replace('\\', '/')
            if not target:
                continue
            if not target.startswith('xl/'):
                target = f'xl/{target}'
            member = name_map.get(target)
            if not member:
                continue
            sheet_root = ET.fromstring(zf.read(member))
            for row in sheet_root.findall('.//a:row', ns):
                values: List[str] = []
                for cell in row.findall('a:c', ns):
                    col_idx = _excel_col_index(cell.attrib.get('r', ''))
                    while len(values) <= col_idx:
                        values.append('')
                    cell_type = cell.attrib.get('t')
                    if cell_type == 'inlineStr':
                        value = ''.join(t.text or '' for t in cell.findall('.//a:t', ns))
                    else:
                        value_node = cell.find('a:v', ns)
                        value = value_node.text if value_node is not None else ''
                        if cell_type == 's' and value != '':
                            try:
                                value = shared_strings[int(value)]
                            except (ValueError, IndexError):
                                value = ''
                    values[col_idx] = _clean_cell(value)
                if any(values):
                    rows.append(values)
    return rows


def _read_jd_cy_workbook_rows(path: Path) -> List[List[str]]:
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None)
        rows: List[List[str]] = []
        for sheet_name, df in sheets.items():
            if '板厚' not in str(sheet_name) and '换算' not in str(sheet_name):
                continue
            for _, row in df.iterrows():
                values = [_clean_cell(value) for value in row.tolist()]
                if any(values):
                    rows.append(values)
        return rows
    except Exception:
        return _read_xlsx_rows_compatible(path)


def _jd_cy_thickness_candidate_paths() -> List[Path]:
    docs_dir = BASE_DIR.parent / 'docs' / 'develop0707'
    paths: List[Path] = []
    if docs_dir.exists():
        for pattern in ('*健鼎*特殊要求*.xlsx', '*超颖*特殊要求*.xlsx', '*超颍*特殊要求*.xlsx'):
            for path in sorted(docs_dir.glob(pattern)):
                if path.is_file() and not path.name.startswith('~$') and path not in paths:
                    paths.append(path)
    return paths


def _load_jd_cy_thickness_map() -> Dict[Tuple[str, str], float]:
    """Load 健鼎/超颖 mil board thickness mapping from the provided special table."""
    global _JD_CY_THICKNESS_CACHE
    if _JD_CY_THICKNESS_CACHE is not None:
        return _JD_CY_THICKNESS_CACHE

    table = {}
    for path in _jd_cy_thickness_candidate_paths():
        try:
            rows = _read_jd_cy_workbook_rows(path)
        except Exception:
            continue
        board_idx, factory_idx, copper_idx = 1, 2, 3
        for row in rows:
            header_map = {value: idx for idx, value in enumerate(row) if value}
            if '板厚' in header_map and '厂内规格' in header_map and '铜厚' in header_map:
                board_idx = header_map['板厚']
                factory_idx = header_map['厂内规格']
                copper_idx = header_map['铜厚']
                continue
            board = _clean_cell(row[board_idx]) if len(row) > board_idx else ''
            factory_mm = _parse_mm_cell(row[factory_idx]) if len(row) > factory_idx else None
            copper = _clean_cell(row[copper_idx]) if len(row) > copper_idx else ''
            if not board or factory_mm is None or not copper:
                continue
            if 'MIL' not in board.upper():
                continue
            table[(_normalize_special_thickness_key(board), _normalize_special_thickness_key(copper))] = factory_mm

    _JD_CY_THICKNESS_CACHE = table
    return table


def _extract_special_mil_candidates(text: str, thick_raw: str) -> List[str]:
    candidates = []
    source = f'{text or ""} {thick_raw or ""}'
    pattern = re.compile(r'(\d+(?:\.\d+)?(?:\([^)]*\))?)\s*(?:(?:±|\+/-|＋/－)\s*\d+(?:\.\d+)?)?\s*MIL', re.IGNORECASE)
    for m in pattern.finditer(source):
        candidates.append(_normalize_special_thickness_key(f'{m.group(1)}mil'))
    # Prefer more specific forms such as 4(106*2)MIL over 4MIL.
    return sorted(set(candidates), key=len, reverse=True)


def _special_mil_value(board_key: str) -> Optional[float]:
    m = re.match(r'(\d+(?:\.\d+)?)', str(board_key or ''))
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


def lookup_jd_cy_mil_thickness(cust_name: str, text: str, thick_raw: str,
                               copper_spec: str) -> Optional[Tuple[float, float, str]]:
    """Return (factory_mm, mil_value, note) for 湖北健鼎/超颖电子 mil thickness specials."""
    if not is_jd_cy_customer(cust_name):
        return None
    if not copper_spec:
        return None
    table = _load_jd_cy_thickness_map()
    if not table:
        return None

    copper_key = _normalize_special_thickness_key(copper_spec)
    for board_key in _extract_special_mil_candidates(text, thick_raw):
        factory_mm = table.get((board_key, copper_key))
        mil_value = _special_mil_value(board_key) or 0.0
        if factory_mm is not None:
            note = f'健鼎/超颖板厚换算表：{board_key}+{copper_key}->{factory_mm:g}mm'
            return factory_mm, mil_value, note

        same_mil_values = {
            value
            for (candidate_board, candidate_copper), value in table.items()
            if candidate_copper == copper_key and _special_mil_value(candidate_board) == mil_value
        }
        if len(same_mil_values) == 1:
            factory_mm = next(iter(same_mil_values))
            note = f'健鼎/超颖板厚换算表：{board_key}+{copper_key}->同mil唯一值{factory_mm:g}mm'
            return factory_mm, mil_value, note
    return None


def get_front_core_back_total_rule(cust_name: str, cust_code: str,
                                   special_by_name: dict, special_by_code: dict) -> str:
    """Return the customer special note for specs where front thickness is core and back thickness is total."""
    entries = []
    if cust_code:
        entries.extend(special_by_code.get(str(cust_code).strip(), []))
    if cust_name:
        entries.extend(special_by_name.get(str(cust_name).strip(), []))

    seen = set()
    for entry in entries:
        req = str(entry.get('req', '') or '').strip()
        if not req or req in seen:
            continue
        seen.add(req)
        req_norm = re.sub(r'\s+', '', req)
        if '前面芯厚后面总厚' in req_norm:
            return req
    return ''


def extract_front_core_back_total_thickness(text: str, copper_spec: str) -> Optional[Tuple[float, str]]:
    """For specs like '1.49 3/3 1.7 ...', return the post-copper total thickness."""
    if not text or not copper_spec:
        return None

    text_norm = str(text).upper().replace('×', 'X')
    copper_pattern = re.escape(str(copper_spec).upper()).replace(r'\/', r'\s*/\s*')
    pattern = (
        r'(?<![\d.])(\d+(?:\.\d+)?)\s+'
        + copper_pattern +
        r'(?:\s+[A-Z]+(?:\s*/\s*[A-Z]+)?){0,3}'
        r'\s+(\d+(?:\.\d+)?)(?!\s*[*X])'
    )
    for match in re.finditer(pattern, text_norm, re.IGNORECASE):
        try:
            total_mm = float(match.group(2))
        except ValueError:
            continue
        if 0.05 <= total_mm <= 10:
            return total_mm, match.group(2)
    return None


def apply_customer_exact_inch_thickness(cust_name: str, thick_raw: str, glue_model: str = "") -> Optional[float]:
    """Customers whose decimal-inch board thickness is encoded from exact inch*25.4."""
    customer = str(cust_name or "")
    glue_key = _normalize_glue_key(glue_model)
    use_exact = (
        "广合" in customer
        or "依利安达" in customer
        or ("台湾敬鹏" in customer and glue_key == "NYHP7350")
    )
    if not use_exact:
        return None
    m = re.search(r'(\d+\.\d+)\s*(?:"|″|”)', str(thick_raw or ""))
    if not m:
        return None
    return float(m.group(1)) * 25.4


def calc_order_thickness(mm_val: float, copper_spec: str, mode: str,
                         thick_total_to_core: dict, thick_core_to_total: dict) -> Tuple[float, bool]:
    """计算实际下单厚度，返回 (下单厚度mm, 是否为总厚下单)
    修复：mode='unknown'时，<0.8直接按芯厚，>=0.8直接按总厚，不做转换
    """
    THRESHOLD = 0.8
    spec_key = copper_spec.upper() if copper_spec else ''

    if mode == 'unknown':
        # 无含铜/不含铜标注：直接按厚度值判断，不做转换
        if mm_val >= THRESHOLD:
            return mm_val, True   # 直接按总厚下单
        else:
            return mm_val, False  # 直接按芯厚下单

    if mode == 'total':
        if mm_val >= THRESHOLD:
            # 含铜且厚度>=0.8mm：直接按总厚下单
            return mm_val, True
        else:
            # 含铜且厚度<0.8mm：客户填的是总厚，但下单要按芯厚，必须做总厚→芯厚转换
            reduce = thick_total_to_core.get(spec_key, 0.0)
            core_val = round(mm_val - reduce, 6)
            return core_val, False
    else:  # core
        if mm_val < THRESHOLD:
            return mm_val, False
        else:
            add = thick_core_to_total.get(spec_key, 0.0)
            total_val = round(mm_val + add, 6)
            return total_val, True


# ============================================================
# 步骤3：铜箔规格
# ============================================================

# 铜箔规格合法值（每部分最多2个字符）
VALID_COPPER_SPECS = {
    'H/H', '1/1', '2/2', '3/3', '4/4', '5/5', '6/6',
    'H/1', '1/H', 'H/2', '2/H', 'H/3', '3/H', 'H/0',
    '1/2', '2/1', '1/3', '3/1', '1/4', '4/1', '1/6', '6/1',
    '2/3', '3/2', '2/0', '3/0', 'F/0', 'F/F', 'F/1',
    'T/T', 'J/J', 'J/0', '0/J', 'I/I', 'A/A', 'C/C', 'D/D',
    'E/E', '3/H', 'H/3', '3/6', '6/3', '4/4', '1/A', 'A/1',
    '1/1.5', '1/F',
}


MICRON_COPPER_MAP = {
    '12': 'T',
    '15': 'J',
    '17': 'H',
    '18': 'H',
    '28': 'K',
    '35': '1',
    '50': 'F',
    '61': 'S',
    '70': '2',
}


def _normalize_copper_part(part: str) -> str:
    part = str(part or "").upper().strip()
    part = re.sub(r'(?:UM|μM|U|OZ)$', '', part, flags=re.IGNORECASE)
    part = part.strip()
    if part == '0.5':
        return 'H'
    if part == '1.5':
        return 'F'
    if part == '2.5':
        return 'E'
    return MICRON_COPPER_MAP.get(part, part)


def extract_copper_spec(text: str) -> Optional[str]:
    """
    提取铜箔规格（如H/H, 1/1, 2/2, T/T, J/J, H/1等）
    修复：限制每部分最多2字符，并且后缀不能是FR-4等非铜箔标识
    """
    text_upper = text.upper().replace('ＯＺ', 'OZ')
    text_upper = re.sub(r'(?<![A-Z0-9])([A-Z0-9.]{1,3}\s*/\s*[A-Z0-9.]{1,3})0Z\b', r'\1OZ', text_upper, flags=re.IGNORECASE)
    # DK values belong to dielectric properties. Without masking, DK=3.53/C级
    # can be misread as the copper pair 53/C.
    text_upper = re.sub(r'\bDK\s*=\s*\d+(?:\.\d+)?', ' DKEPSILON ', text_upper, flags=re.IGNORECASE)
    # C/M级 is a board grade marker, not copper foil. Mask it before generic X/Y copper matching.
    text_upper = re.sub(r'(?<![A-Z0-9])C\s*/\s*M\s*(?=级|級|GRADE|\b)', ' CMGRADE ', text_upper, flags=re.IGNORECASE)
    # Split tightly written thickness + copper specs, e.g. 0.100mm1/1 -> 0.100mm 1/1.
    text_upper = re.sub(r'(MM|MIL)\s*(?=[A-Z0-9.]{1,3}\s*/)', r'\1 ', text_upper, flags=re.IGNORECASE)

    # Customer notation 1/H is normalized to the internal H/1 order.
    if re.search(r'(?<![A-Z0-9])1\s*/\s*H(?![A-Z0-9])', text_upper):
        return 'H/1'

    # Compact micron notation used as 012012 = 12um/12um.
    if re.search(r'(?<!\d)012012(?!\d)', text_upper):
        return 'T/T'

    # Slash-delimited customer specs may state one symmetric copper thickness,
    # e.g. /18um/NY6300S/. The surrounding separators keep this narrow.
    m = re.search(
        r'(?:^|[/;；])\s*(12|15|17|18|28|35|50|61|70)\s*(?:UM|ΜM|Μ|U|微米)\s*(?=$|[/;；])',
        text_upper,
        re.IGNORECASE,
    )
    if m:
        part = _normalize_copper_part(m.group(1))
        return f"{part}/{part}"

    # 35/00 means 35um on one side and no copper on the other.
    m = re.search(r'(?<![A-Z0-9])(12|15|17|18|28|35|50|61|70)\s*/\s*00(?![A-Z0-9])', text_upper)
    if m:
        return f"{_normalize_copper_part(m.group(1))}/0"

    # Missing delimiter between thickness and copper, e.g. 1.515/15.
    m = re.search(r'(?<![\d.])\d+\.\d(\d{2})\s*/\s*(\d{2})(?![A-Z0-9])', text_upper)
    if m and m.group(1) in MICRON_COPPER_MAP and m.group(2) in MICRON_COPPER_MAP:
        return f"{MICRON_COPPER_MAP[m.group(1)]}/{MICRON_COPPER_MAP[m.group(2)]}"

    # Customer minus-grade copper: 1-/1- -> K/K.
    m = re.search(r'(?<![A-Z0-9])1-\s*/\s*1-(?![A-Z0-9])', text_upper, re.IGNORECASE)
    if m:
        return 'K/K'

    # 1.5/1.5, 2.5/2.5 and asymmetric half-ounce notations.
    m = re.search(r'(?<![A-Z0-9])(\d(?:\.5)?)\s*/\s*(\d(?:\.5)?)(?![A-Z0-9.])', text_upper, re.IGNORECASE)
    if m and any('.5' in item for item in (m.group(1), m.group(2))):
        left = _normalize_copper_part(m.group(1))
        right = _normalize_copper_part(m.group(2))
        return f"{left}/{right}"

    # 1.5oz/1.5oz, 1oz/1.5oz, 1.5oz/1oz: both sides carry their own oz unit.
    m = re.search(r'(?<![A-Z0-9])(\d(?:\.5)?)\s*OZ\s*/\s*(\d(?:\.5)?)\s*OZ\b', text_upper, re.IGNORECASE)
    if m:
        left = _normalize_copper_part(m.group(1))
        right = _normalize_copper_part(m.group(2))
        return f"{left}/{right}"

    # 15μ/15μ, 15um/15um: both sides carry their own micron unit.
    m = re.search(r'(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*(?:UM|μM|μ|U|微米)\s*/\s*(\d+(?:\.\d+)?)\s*(?:UM|μM|μ|U|微米)', text_upper, re.IGNORECASE)
    if m:
        left = _normalize_copper_part(m.group(1))
        right = _normalize_copper_part(m.group(2))
        return f"{left}/{right}"

    # 15/15um、18/18um、35/35um、50/50um 这类微米铜厚写法。
    m = re.search(r'(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)(?:\s*(?:UM|μM|U|微米))', text_upper, re.IGNORECASE)
    if m:
        left = _normalize_copper_part(m.group(1))
        right = _normalize_copper_part(m.group(2))
        return f"{left}/{right}"

    # 2/2Oz、H/HOz、3/3oz 等 oz 后缀写法，去掉 oz 单位。
    m = re.search(r'(?<![A-Z0-9])([A-Z0-9.]{1,3})\s*/\s*([A-Z0-9.]{1,3})\s*OZ\b', text_upper, re.IGNORECASE)
    if m:
        left = _normalize_copper_part(m.group(1))
        right = _normalize_copper_part(m.group(2))
        return f"{left}/{right}"

    # 无单位但两侧是明确微米值时，同样按微米铜厚映射。
    m = re.search(r'(?<![A-Z0-9])(\d{2})\s*/\s*(\d{2})(?![A-Z0-9])', text_upper, re.IGNORECASE)
    if m and m.group(1) in MICRON_COPPER_MAP and m.group(2) in MICRON_COPPER_MAP:
        return f"{MICRON_COPPER_MAP[m.group(1)]}/{MICRON_COPPER_MAP[m.group(2)]}"

    # 先尝试匹配 X/YFR-4 格式（如1/1FR-4），取X/Y部分
    m = re.search(r'(?<![A-Z0-9])([A-Z0-9]{1,2})/([A-Z0-9]{1,2})FR[-s]?4', text_upper, re.IGNORECASE)
    if m:
        return f"{m.group(1).upper()}/{m.group(2).upper()}"
    # 优先匹配合法铜箔规格（每部分1-2字符，后面不跟字母数字）
    m = re.search(r'(?<![A-Z0-9])([A-Z0-9]{1,2})/([A-Z0-9]{1,2})(?![A-Z0-9])', text_upper, re.IGNORECASE)
    if m:
        spec = f"{m.group(1).upper()}/{m.group(2).upper()}"
        return spec
    # 宽松匹配（带括号/空格/连字符后缀）
    m = re.search(r'([A-Z0-9]{1,2})/([A-Z0-9]{1,2})[\s\(（\-]', text_upper, re.IGNORECASE)
    if m:
        return f"{m.group(1).upper()}/{m.group(2).upper()}"
    return None


def copper_spec_to_code(spec: str) -> str:
    """铜箔规格转代码"""
    if not spec:
        return ''
    return spec.replace('/', '')


# ============================================================
# 步骤4：尺寸代码
# ============================================================

STANDARD_SIZE_RANGES = [
    (37.0, 37.3, 49.0, 49.3, 37.3, 49.3),
    (41.0, 41.3, 49.0, 49.3, 41.3, 49.3),
    (43.0, 43.3, 49.0, 49.3, 43.3, 49.3),
    (74.0, 74.6, 49.0, 49.3, 74.3, 49.3),
    (82.0, 82.6, 49.0, 49.3, 82.3, 49.3),
    (86.0, 86.6, 49.0, 49.3, 86.3, 49.3),
]

FABRIC_CODES = ("0106", "106", "1035", "1037", "1078", "1080", "1086", "1506", "2116", "3313", "7628")

SPECIAL_CUSTOMER_SIZE_SIDE_MAP = {
    37.0: 37.3,
    41.0: 41.3,
    43.0: 43.3,
    74.0: 74.3,
    82.0: 82.3,
    86.0: 86.3,
    49.0: 49.3,
}

SPECIAL_CUSTOMER_SIZE_MM_SIDE_MAP = {
    940: 37.3,
    1041: 41.3,
    1092: 43.3,
    1880: 74.3,
    2082: 82.3,
    2184: 86.3,
    1245: 49.3,
}


def is_pp_or_rc_spec(text: str) -> bool:
    """Identify PP/RC style specs that are intentionally out of base-sheet transcoding scope."""
    raw = str(text or '')
    upper = raw.upper()
    if '半固化片' in raw:
        return True
    if re.search(r'(?<![A-Z0-9])PP(?![A-Z0-9])', upper):
        return True
    if re.search(r'P\s*片', raw, re.IGNORECASE):
        return True
    if re.search(r'R\s*/\s*C', upper):
        return True
    if 'RC' in upper:
        return True
    if '%' in raw or '％' in raw:
        return True
    fabric = "|".join(FABRIC_CODES)
    if re.search(rf'#\s*(?:{fabric})\b', upper, re.IGNORECASE):
        return True
    if re.search(r'#\s*\d{3,4}\b', upper, re.IGNORECASE):
        return True
    return False


def _strip_fabric_structures(text: str) -> str:
    fabric = "|".join(FABRIC_CODES)
    cleaned = re.sub(rf'\b\d+\s*[*X]\s*(?:{fabric})\b', ' ', text, flags=re.IGNORECASE)
    cleaned = re.sub(rf'\b(?:{fabric})\s*[*X]\s*\d+\b', ' ', cleaned, flags=re.IGNORECASE)
    return cleaned


def _map_special_customer_size_side(value: float) -> Optional[float]:
    if value >= 300:
        return SPECIAL_CUSTOMER_SIZE_MM_SIDE_MAP.get(int(round(value)))
    rounded = round(float(value), 2)
    for source, target in SPECIAL_CUSTOMER_SIZE_SIDE_MAP.items():
        if abs(rounded - source) <= 0.05 or abs(rounded - target) <= 0.05:
            return target
    return None


def get_customer_size_override(cust_name: str, text: str) -> Optional[Tuple[float, float, str]]:
    """Customer-specific size mappings that are explicit in confirmed CCL rules."""
    if "珠海超毅" in str(cust_name or ""):
        match = re.search(
            r"F\s*(\d+(?:\.\d+)?)\s*MM\s*[*X]\s*W\s*(\d+(?:\.\d+)?)\s*MM",
            str(text or ""),
            re.IGNORECASE,
        )
        if match:
            f_mm, w_mm = float(match.group(1)), float(match.group(2))
            return _ceil2(w_mm / 25.4), _ceil2(f_mm / 25.4), match.group(0)
    if "惠州威健" in str(cust_name or ""):
        normalized = str(text or "").replace("×", "*").replace("x", "*").replace("X", "*")
        if re.search(r"(?<!\d)2184\s*\*\s*1245(?!\d)", normalized):
            return 86.3, 49.0, "2184*1245"
    if "特创" not in str(cust_name or ""):
        return None
    text_norm = str(text or "").replace('×', 'X').replace('x', 'X').replace('＊', '*')
    text_no_fabric = _strip_fabric_structures(text_norm)
    pairs: List[Tuple[float, float, str]] = []

    for m in re.finditer(r'(?<!\d)(\d+(?:\.\d+)?)\s*[*X]\s*(\d+(?:\.\d+)?)(?!\d)', text_no_fabric):
        pairs.append((float(m.group(1)), float(m.group(2)), m.group(0)))

    nums = [
        float(m.group(0))
        for m in re.finditer(r'(?<!\d)\d{3,4}(?:\.\d+)?(?!\d)', text_no_fabric)
        if float(m.group(0)) >= 300
    ]
    for w, h in zip(nums, nums[1:]):
        pairs.append((w, h, f"{w:g} {h:g}"))

    for w, h, raw in reversed(pairs):
        mapped_w = _map_special_customer_size_side(w)
        mapped_h = _map_special_customer_size_side(h)
        if mapped_w is not None and mapped_h is not None:
            return mapped_w, mapped_h, raw
        if mapped_w is not None or mapped_h is not None:
            factory_w = mapped_w if mapped_w is not None else (_ceil2(w / 25.4) if w >= 300 else w)
            factory_h = mapped_h if mapped_h is not None else (_ceil2(h / 25.4) if h >= 300 else h)
            return factory_w, factory_h, raw
        mapped_w = _map_special_customer_size_side(h)
        mapped_h = _map_special_customer_size_side(w)
        if mapped_w is not None and mapped_h is not None:
            return mapped_w, mapped_h, raw
    return None


def _ceil2(value: float) -> float:
    return math.ceil((float(value) * 100) - 1e-9) / 100


def _normalize_latitude_49(value: float) -> Optional[float]:
    """Normalize the latitude/width dimension around the business 49.x size bands."""
    value = round(float(value) + 1e-9, 2)
    if value < 49.0:
        return value
    if value <= 49.09:
        return 49.0
    if value <= 49.30:
        return value
    if value <= 49.49:
        return None
    if value <= 49.60:
        return 49.6
    return None


def _apply_latitude_rule(w: float, h: float) -> Optional[Tuple[float, float]]:
    h_norm = _normalize_latitude_49(h)
    if h_norm is None:
        return None
    return round(w, 2), h_norm


def _normalize_standard_size(w: float, h: float, snap_to_standard: bool = False) -> Optional[Tuple[float, float]]:
    """Snap near-standard panel sizes into the business standard ranges."""
    def fit(value: float, low: float, high: float) -> float:
        if value < low:
            return low
        if value > high:
            return high
        return round(value, 2)

    def at_or_below_lower_standard(value: float, low: float) -> bool:
        return value <= low + 1e-9

    for w_min, w_max, h_min, h_max, std_w, std_h in STANDARD_SIZE_RANGES:
        if (w_min - 0.10) <= w <= (w_max + 0.10) and (h_min - 0.10) <= h <= (h_max + 0.10):
            if snap_to_standard:
                if at_or_below_lower_standard(w, w_min) and h <= (h_min + 0.09):
                    return _apply_latitude_rule(w_min, h_min)
                h_for_rule = h if h >= h_min else h_min
                return _apply_latitude_rule(std_w, h_for_rule)
            h_for_rule = h if h >= h_min else fit(h, h_min, h_max)
            return _apply_latitude_rule(fit(w, w_min, w_max), h_for_rule)
        if (w_min - 0.10) <= h <= (w_max + 0.10) and (h_min - 0.10) <= w <= (h_max + 0.10):
            if snap_to_standard:
                if at_or_below_lower_standard(h, w_min) and w <= (h_min + 0.09):
                    return _apply_latitude_rule(w_min, h_min)
                w_for_rule = w if w >= h_min else h_min
                return _apply_latitude_rule(std_w, w_for_rule)
            w_for_rule = w if w >= h_min else fit(w, h_min, h_max)
            return _apply_latitude_rule(fit(h, w_min, w_max), w_for_rule)
    return _apply_latitude_rule(round(w, 2), round(h, 2))


def _size_from_numbers(w: float, h: float, unit: str) -> Optional[Tuple[float, float]]:
    if unit == "mm":
        w = _ceil2(w / 25.4)
        h = _ceil2(h / 25.4)
        return _normalize_standard_size(w, h, snap_to_standard=True)
    return _normalize_standard_size(w, h, snap_to_standard=False)


def extract_size(text: str) -> Optional[Tuple[float, float]]:
    """
    提取尺寸（宽×高）
    支持 inch 标记（"、in、inch）、mm 标记、W/F 前缀和大数字无单位 mm 写法。
    """
    text_norm = str(text).replace('×', 'X').replace('x', 'X').replace('＊', '*')
    text_norm = re.sub(r'(?<=\d)\s+\.\s*(?=\d)', '.', text_norm)
    text_no_fabric = _strip_fabric_structures(text_norm)

    # 经纬标注 inch 写法：16"(纬)X20.4"(经) should output 经 x 纬 -> 20.4 x 16.
    wh_pattern = (
        r'(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)\s*[\(（]\s*纬\s*[\)）]\s*'
        r'[*X]\s*(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?\s*[\(（]\s*经\s*[\)）]'
    )
    matches = list(re.finditer(wh_pattern, text_norm, re.IGNORECASE))
    if matches:
        m = matches[-1]
        return _size_from_numbers(float(m.group(2)), float(m.group(1)), "inch")

    hw_pattern = (
        r'(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)\s*[\(（]\s*经\s*[\)）]\s*'
        r'[*X]\s*(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?\s*[\(（]\s*纬\s*[\)）]'
    )
    matches = list(re.finditer(hw_pattern, text_norm, re.IGNORECASE))
    if matches:
        m = matches[-1]
        return _size_from_numbers(float(m.group(1)), float(m.group(2)), "inch")

    # 经纬无括号写法：37.3经X49.3纬、49.3纬X37.3经。
    m = re.search(
        r'(\d+(?:\.\d+)?)\s*经\s*[*X]\s*(\d+(?:\.\d+)?)\s*纬',
        text_norm, re.IGNORECASE
    )
    if m:
        return _size_from_numbers(float(m.group(1)), float(m.group(2)), "inch")
    m = re.search(
        r'(\d+(?:\.\d+)?)\s*纬\s*[*X]\s*(\d+(?:\.\d+)?)\s*经',
        text_norm, re.IGNORECASE
    )
    if m:
        return _size_from_numbers(float(m.group(2)), float(m.group(1)), "inch")

    m = re.search(
        r'经向\s*(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?\s*[*X]\s*纬向\s*(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?',
        text_norm, re.IGNORECASE
    )
    if m:
        return _size_from_numbers(float(m.group(1)), float(m.group(2)), "inch")
    m = re.search(
        r'纬向\s*(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?\s*[*X]\s*经向\s*(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?',
        text_norm, re.IGNORECASE
    )
    if m:
        return _size_from_numbers(float(m.group(2)), float(m.group(1)), "inch")

    # 1. 显式 inch：41.00"X49.00"、37inX49in、37 inch X 49 inch。
    inch_pattern = (
        r'(?:\b[WF]\s*)?(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)\s*'
        r'(?:[*X])\s*(?:[WF]\s*)?(\d+(?:\.\d+)?)\s*(?:"|″|”|INCH|IN|英寸)?'
    )
    matches = list(re.finditer(inch_pattern, text_norm, re.IGNORECASE))
    if matches:
        m = matches[-1]
        return _size_from_numbers(float(m.group(1)), float(m.group(2)), "inch")

    # 2. 显式 mm：W463 X F616(mm)、463X616mm。
    mm_pattern = (
        r'(?:\b[WF]\s*)?(\d+(?:\.\d+)?)\s*(?:MM)?\s*'
        r'(?:[*X])\s*(?:[WF]\s*)?(\d+(?:\.\d+)?)\s*(?:MM)?\s*(?:\(\s*MM\s*\))?'
    )
    mm_matches = []
    for m in re.finditer(mm_pattern, text_norm, re.IGNORECASE):
        span_text = text_norm[m.start():m.end()].upper()
        w, h = float(m.group(1)), float(m.group(2))
        if ('MM' in span_text or w > 200 or h > 200) and 5 <= (w / 25.4) <= 100 and 5 <= (h / 25.4) <= 100:
            mm_matches.append((w, h))
    if mm_matches:
        return _size_from_numbers(*mm_matches[-1], unit="mm")

    # 3. 找所有 W*H 格式，过滤叠构。无单位且数值在英寸范围内时按 inch 兼容旧写法。
    all_sizes = re.findall(r'(?:\b[WF]\s*)?(\d+(?:\.\d+)?)\s*[*X]\s*(?:[WF]\s*)?(\d+(?:\.\d+)?)', text_norm, re.IGNORECASE)
    valid = []
    for w_str, h_str in all_sizes:
        w, h = float(w_str), float(h_str)
        # 过滤条件：
        # - 4位整数（如7628, 1080, 2116）是布号，不是尺寸
        # - 尺寸通常在10~100英寸范围内
        w_int_len = len(w_str.split('.')[0])
        h_int_len = len(h_str.split('.')[0])
        if w_int_len == 4 or h_int_len == 4:
            continue  # 叠构布号
        if w < 5 or h < 5:
            continue  # 太小，可能是层数
        if w > 200 or h > 200:
            continue  # 太大
        valid.append((w, h))

    if valid:
        return _size_from_numbers(*valid[-1], unit="inch")  # 取最后一个（通常尺寸在末尾）

    fabric = "|".join(FABRIC_CODES)
    fabric_size_matches = []
    fabric_pair_pattern = rf'(?<!\d)(\d{{2}}(?:\.\d+)?)\s+(\d{{2}}(?:\.\d+)?)(?=\s+(?:{fabric})\s*[*X]\s*\d+)'
    for m in re.finditer(fabric_pair_pattern, text_norm, re.IGNORECASE):
        w, h = float(m.group(1)), float(m.group(2))
        if 10 <= w <= 100 and 10 <= h <= 100:
            fabric_size_matches.append((w, h))
    if fabric_size_matches:
        return _size_from_numbers(*fabric_size_matches[-1], unit="inch")

    # 4. 无分隔符、无单位的大尺寸数字，默认按 mm：2082 1245 -> 82*49 inch。
    mm_pair_matches = []
    for m in re.finditer(r'(?<!\d)(\d{3,4})(?:\.\d+)?\s+(\d{3,4})(?:\.\d+)?(?!\d)', text_no_fabric):
        w, h = float(m.group(1)), float(m.group(2))
        if w < 300 or h < 300:
            continue
        w_in, h_in = w / 25.4, h / 25.4
        if 5 <= w_in <= 100 and 5 <= h_in <= 100:
            mm_pair_matches.append((w, h))
    if not mm_pair_matches:
        nums = [float(m.group(0)) for m in re.finditer(r'(?<!\d)\d{3,4}(?:\.\d+)?(?!\d)', text_no_fabric)]
        for w, h in zip(nums, nums[1:]):
            if w < 300 or h < 300:
                continue
            w_in, h_in = w / 25.4, h / 25.4
            if 5 <= w_in <= 100 and 5 <= h_in <= 100:
                mm_pair_matches.append((w, h))
    if mm_pair_matches:
        return _size_from_numbers(*mm_pair_matches[-1], unit="mm")

    return None


def size_to_code(w: float, h: float) -> str:
    """尺寸转8位代码：整数2位+小数2位"""
    def encode_dim(val: float) -> str:
        val = round(float(val) + 1e-9, 2)
        int_part = int(val)
        dec_part = int(round((val - int_part) * 100))
        if dec_part == 100:
            int_part += 1
            dec_part = 0
        return f"{int_part:02d}{dec_part:02d}"
    return encode_dim(w) + encode_dim(h)


def get_special_size(cust_name: str, w: float, h: float, special_by_name: dict) -> Optional[str]:
    """查询客户特殊尺寸"""
    entries = special_by_name.get(cust_name, [])
    w_int, h_int = int(w), int(h)
    for e in entries:
        req = e.get('req', '')
        patterns = [
            rf'{w_int}\s*[*×]\s*{h_int}[^\d]*(\d{{8}})',
            rf'{w_int}["\u201d]?\s*[*×]\s*{h_int}["\u201d]?[^\d]*(\d{{8}})',
        ]
        for pat in patterns:
            m = re.search(pat, req)
            if m:
                return m.group(1)
    return None


# ============================================================
# 步骤5：胶水类别
# ============================================================

def get_glue_category(glue_model: str, glue_cat_map: dict, glue_code: str = "") -> str:
    """查询胶系类别（普通/特殊）"""
    if not glue_model:
        return ''
    norm_model = _normalize_glue_key(glue_model)
    if norm_model in {"NY2140L", "NY3170HC"}:
        return "特殊"
    if str(glue_code or "").upper() in {"RV", "DA"}:
        return "特殊"

    norm_cat_map = {_normalize_glue_key(k): v for k, v in glue_cat_map.items()}
    for key in _glue_candidates(glue_model):
        norm_key = _normalize_glue_key(key)
        if key in glue_cat_map:
            return glue_cat_map[key]
        if norm_key in norm_cat_map:
            return norm_cat_map[norm_key]
        # 部分匹配
        for k, v in glue_cat_map.items():
            norm_k = _normalize_glue_key(k)
            if norm_key and (norm_key == norm_k or norm_key in norm_k or norm_k in norm_key):
                return v
    return '普通'  # 默认普通


# ============================================================
# 步骤6：铜箔类型代码
# ============================================================

# 特殊铜箔关键词 → 代码（按优先级，长词优先）
SPECIAL_COPPER_MAP = [
    ('HS2-M2-VSP', 'P'),
    ('HS2-M2-VS', 'P'),
    ('HVLP5', 'J'),
    ('HVLP4', 'Z'),
    ('HVLP3', 'K'),
    ('HVLP2', 'P'),
    ('HVLP1', 'O'),
    ('HVLP', 'O'),
    ('RTF4', 'G'),
    ('RTF3', 'A'),
    ('RTF2', 'B'),
    ('RTF1', 'R'),
    ('RTF', 'R'),
    ('VLP1', 'O'),
    ('VLP', 'L'),
    # 注意：HTE 是常规电解铜，不算特殊铜箔，不在此列表
]


def get_copper_type_code(text: str) -> str:
    """
    根据规格文本判断铜箔类型代码
    修复：HTE是常规铜，不映射为D；只有RTF/HVLP/VLP才特殊
    """
    text_upper = text.upper()
    if 'IGAV UV' in text_upper:
        return 'I'
    has_hs2_vsp = "HS2-M2-VSP" in text_upper or "HS2-M2-VS" in text_upper
    if has_hs2_vsp and "RTF" in text_upper:
        return "N"
    if re.search(r'RTF3\s*/\s*RTF(?!\d)', text_upper) or re.search(r'RTF3\s*\+\s*RTF(?!\d)', text_upper):
        return "T"
    for keyword, code in SPECIAL_COPPER_MAP:
        if keyword in text_upper:
            return code
    if '有水印' in text or '带水印' in text or re.search(r'(?<!NO\s)UV\s*(?:YES|Y)\b', text_upper):
        return 'Q'
    return 'W'  # 默认常规铜无水印


# ============================================================
# 步骤7：基板级别代码
# ============================================================

_COPPER_OZ_BY_CODE = {
    "0": 0.0,
    "J": 15 / 35,
    "H": 0.5,
    "K": 28 / 35,
    "1": 1.0,
    "F": 1.5,
    "2": 2.0,
}


def _max_copper_oz(copper_spec: Optional[str]) -> Optional[float]:
    if not copper_spec:
        return None
    values = []
    for part in str(copper_spec).split('/'):
        part = part.strip().upper()
        if part in _COPPER_OZ_BY_CODE:
            values.append(_COPPER_OZ_BY_CODE[part])
    if not values:
        return None
    return max(values)


def get_grade_code(text: str, cust_name: str, glue_model: str,
                   special_by_name: dict, grade_desc_to_code: dict = None) -> str:
    """
    获取基板级别代码（2位）
    当前阶段除非业务给出可执行规则，否则默认 A1。
    """
    grade_source = f"{text or ''} {glue_model or ''}"
    cust_text = str(cust_name or "")
    text_upper = str(text or "").upper()
    if "惠州威健" in cust_text and "CTI" in text_upper and "600" in text_upper and _normalize_glue_key(glue_model) == "NY1600":
        return "A2"
    if any(name in cust_text for name in ("广东依顿", "森德科技")) and (
        "耐CAF" in grade_source or "ANTI-CAF" in text_upper or "ANTI CAF" in text_upper
    ):
        return "AC"
    if any(keyword in grade_source for keyword in (
        "汽车专用", "汽车板", "车载板", "车用", "CAR BOARD", "CARBOARD", "AUTOMOTIVE",
    )):
        grade_desc_to_code = grade_desc_to_code or {}
        for desc, code in grade_desc_to_code.items():
            if "汽车" in str(desc):
                return code
        return "AC"
    if "HDI专用" in grade_source or "HDI 专用" in grade_source.upper():
        return "AD"
    if "MINILED" in text_upper or "MINI LED" in text_upper:
        return "AM"
    if re.search(r'(?<![A-Z0-9])LED(?![A-Z0-9])', text_upper):
        return "AM"
    if "TFT" in str(text).upper():
        return "AT"
    norm_glue = _normalize_glue_key(glue_model)
    if "中宝悦嘉" in str(cust_name or "") and norm_glue == "NY1600":
        return "A2"
    if any(name in str(cust_name or "") for name in ("广华升鑫", "深华升鑫")) and norm_glue == "NYA2":
        return "AC"
    if ("湖奥士康" in str(cust_name or "") or "景旺" in str(cust_name or "")) and norm_glue in {"NYA1", "NYA2"}:
        return "AC"
    if "崇达" in cust_text and norm_glue == "NY3150HC":
        return "AC"
    if "深万基隆" in cust_text and norm_glue.startswith("NY2140"):
        return "A1"
    if "珠海益天" in cust_text and norm_glue.startswith("NY2140"):
        return "F1"
    if norm_glue == "NY2140" and "江苏苏杭" not in cust_text:
        thickness_mm, _, _ = extract_thickness_mm(text)
        copper_spec = extract_copper_spec(text)
        max_oz = _max_copper_oz(copper_spec)
        if thickness_mm is not None and thickness_mm > 0.8 and max_oz is not None and max_oz <= 1:
            return "F1"
    if "世运" in str(cust_name or "") and norm_glue == "NY3170HC":
        return "AC"
    return 'A1'


# ============================================================
# 步骤9：结构代码
# ============================================================

def get_struct_code(text: str, cust_name: str) -> str:
    """
    获取结构代码（1位）
    A=第一结构(1-2层PP), B=第二结构(3-4层), C=第三结构(5-6层), D=第四结构(7-10层)
    E=健鼎/超颍第一结构, H=健鼎/超颍第二结构, I=健鼎/超颍第三结构
    N=单面板（J/J铜厚）
    """
    is_special_customer = any(name in cust_name for name in ['超颖', '超颍', '健鼎'])

    # 计算总PP层数
    total_layers = 0

    # 匹配叠构括号内容
    struct_match = re.search(r'[\(（]([^)）]+)[\)）]', text)
    if struct_match:
        struct_text = struct_match.group(1)
        # 格式1：数字*布号（如 4*7628, 1*2116）
        for n, code in re.findall(r'(\d+)\s*[*×]\s*(\d{4})', struct_text):
            total_layers += int(n)
        # 格式2：布号*数字（如 7628*2, 1080*1）
        if total_layers == 0:
            for code, n in re.findall(r'(\d{4})\s*[*×]\s*(\d+)', struct_text):
                total_layers += int(n)
        # 格式3：单个布号（如 7628）
        if total_layers == 0:
            codes = re.findall(r'\d{4}', struct_text)
            if codes:
                total_layers = len(codes)

    # 无括号叠构
    if total_layers == 0:
        for n, code in re.findall(r'(\d+)\s*[*×]\s*(\d{4})', text):
            total_layers += int(n)
        if total_layers == 0:
            for code, n in re.findall(r'(\d{4})\s*[*×]\s*(\d+)', text):
                total_layers += int(n)

    if is_special_customer:
        if total_layers == 0 or total_layers <= 2:
            return 'E'
        elif total_layers <= 5:
            return 'H'
        else:
            return 'I'
    else:
        # 单面板（J/J铜厚）
        copper = extract_copper_spec(text)
        if copper and copper.upper() in ('J/J', 'J/0', '0/J'):
            return 'N'

        if total_layers == 0 or total_layers <= 2:
            return 'A'
        elif total_layers <= 4:
            return 'B'
        elif total_layers <= 6:
            return 'C'
        else:
            return 'D'


def _split_structured_condition(value: str) -> list[str]:
    return [item.strip() for item in re.split(r'[/,，、|；;]+', str(value or '')) if item.strip()]


def _rule_yes(value: str) -> bool:
    return str(value or '').strip() == '是'


def _normalize_match_text(value: str) -> str:
    return re.sub(r'\s+', '', str(value or '')).upper()


def _structured_rule_customer_matches(rule: dict, cust_code: str, cust_name: str) -> bool:
    rule_code = _normalize_customer_code(rule.get('客户代码', ''))
    rule_name = _normalize_customer_name(rule.get('客户简称', ''))
    code_key = _normalize_customer_code(cust_code)
    name_key = _normalize_customer_name(cust_name)
    if rule_code and rule_code != code_key:
        return False
    if rule_name and rule_name not in name_key and name_key not in rule_name:
        return False
    return True


def _structured_rule_source_text(rule: dict, spec_text: str, context_text: str) -> str:
    source = str(rule.get('匹配来源字段', '') or '')
    combined = f'{spec_text or ""} {context_text or ""}'
    if '客户规格' in source and '整行' not in source:
        return spec_text or ''
    if any(term in source for term in ('备注', '物料描述', '客户特殊要求', '客户特殊需求')) and '整行' not in source:
        return context_text or ''
    return combined


def _structured_rule_conditions_match(rule: dict, spec_text: str, context_text: str,
                                      glue_model: str, copper_spec: str) -> bool:
    target_text = _normalize_match_text(_structured_rule_source_text(rule, spec_text, context_text))
    spec_norm = _normalize_match_text(spec_text)
    glue_norm = _normalize_match_text(glue_model)
    copper_norm = _normalize_match_text(copper_spec)

    exclude_keywords = _split_structured_condition(rule.get('排除关键词', ''))
    if any(_normalize_match_text(keyword) in target_text for keyword in exclude_keywords):
        return False

    keywords = _split_structured_condition(rule.get('关键词条件', ''))
    if keywords and not any(_normalize_match_text(keyword) in target_text for keyword in keywords):
        return False

    glue_condition = _normalize_match_text(rule.get('适用胶系', ''))
    if glue_condition and glue_condition not in glue_norm and glue_condition not in spec_norm:
        return False

    copper_condition = _normalize_match_text(rule.get('铜厚条件', ''))
    if copper_condition and copper_condition not in copper_norm and copper_condition not in spec_norm:
        return False

    for key in ('尺寸条件', '含不含铜条件', 'CTI条件'):
        condition = _normalize_match_text(rule.get(key, ''))
        if condition and condition not in spec_norm and condition not in target_text:
            return False
    return True


def _executable_structured_rules(tables: dict, cust_code: str, cust_name: str, spec_text: str,
                                 context_text: str, glue_model: str, copper_spec: str) -> list[dict]:
    rules = []
    for rule in tables.get('structured_special_rules', []):
        if str(rule.get('规则状态', '') or '').strip() == '归档':
            continue
        if not _rule_yes(rule.get('启用')):
            continue
        if not _rule_yes(rule.get('是否参与转码')):
            continue
        if _rule_yes(rule.get('待确认')):
            continue
        if str(rule.get('执行策略', '') or '').strip() not in ('', '参与转码'):
            continue
        if not _structured_rule_customer_matches(rule, cust_code, cust_name):
            continue
        if not _structured_rule_conditions_match(rule, spec_text, context_text, glue_model, copper_spec):
            continue
        rules.append(rule)
    return sorted(rules, key=lambda item: int(float(item.get('优先级') or 0)), reverse=True)


def apply_structured_special_overrides(tables: dict, cust_code: str, cust_name: str, spec_text: str,
                                       context_text: str, glue_model: str, copper_spec: str,
                                       copper_code: str, grade_code: str, tc_code: str,
                                       struct_code: str, errors: list[str], steps: dict) -> tuple[str, str, str, str]:
    field_map = [
        ('覆盖铜厚代码', 'copper_code', 'step3_copper_code', '无法识别铜箔规格'),
        ('覆盖基板级别', 'grade_code', 'step7_grade_code', ''),
        ('覆盖总芯厚', 'tc_code', 'step8_tc_code', ''),
        ('覆盖结构码', 'struct_code', 'step9_struct_code', ''),
    ]
    values = {
        'copper_code': copper_code,
        'grade_code': grade_code,
        'tc_code': tc_code,
        'struct_code': struct_code,
    }
    applied_fields: set[str] = set()
    notes: list[str] = []
    conflicts: list[str] = []

    for rule in _executable_structured_rules(tables, cust_code, cust_name, spec_text, context_text, glue_model, copper_spec):
        rule_id = rule.get('规则ID', '')
        for override_col, value_key, step_key, error_text in field_map:
            override_value = str(rule.get(override_col, '') or '').strip().upper()
            if not override_value:
                continue
            if value_key in applied_fields:
                if values[value_key] != override_value:
                    conflicts.append(f"{override_col}冲突：已用{values[value_key]}，忽略{rule_id or '未编号'}={override_value}")
                continue
            old_value = values[value_key]
            values[value_key] = override_value
            steps[step_key] = override_value
            applied_fields.add(value_key)
            if error_text:
                errors[:] = [err for err in errors if err != error_text]
            notes.append(f"{rule_id or '未编号'} {override_col}:{old_value}->{override_value} ({rule.get('目标动作内容') or rule.get('规则解释') or '结构化特殊规则'})")

    if notes:
        steps['structured_special_rules'] = notes
    if conflicts:
        steps['structured_special_rule_conflicts'] = conflicts
    return values['copper_code'], values['grade_code'], values['tc_code'], values['struct_code']


# ============================================================
# 主转码函数
# ============================================================

def transcode_row(s_text: str, e_text: str, d_cust_name: str, a_cust_code: str,
                  tables: dict, context_text: str = "") -> Tuple[str, dict, str]:
    """
    对单行进行转码
    返回：(编码结果, 步骤详情字典, 错误信息)
    """
    steps = {}
    errors = []

    # ── 步骤1：胶系代码 ──
    glue_model = extract_glue_model_from_rules(s_text, tables['glue_exact_map'])
    steps['glue_model'] = glue_model or '未识别'
    if not glue_model:
        errors.append('无法识别胶系型号')
        glue_code = '??'
    else:
        glue_code = get_glue_code(glue_model, tables['glue_exact_map'], tables['glue_model_map'], d_cust_name)
        if not glue_code:
            errors.append(f'胶系 {glue_model} 在代码表中未找到')
            glue_code = '??'
    if (
        "深南" in str(d_cust_name or "")
        and _normalize_glue_key(glue_model or "") == "NYP5Q"
        and ("考试板" in str(s_text or "") or "90022-4" in str(s_text or ""))
    ):
        glue_code = "CG"
        steps['glue_note'] = '深南NY-P5Q考试板按确认样本使用CG胶系'
    if "宜兴硅谷" in str(d_cust_name or "") and re.search(r"(?<![A-Z0-9])NY2150(?![A-Z0-9])", str(s_text or "").upper()):
        explicit_code = get_glue_code("NY2150", tables['glue_exact_map'], tables['glue_model_map'], d_cust_name)
        if explicit_code:
            glue_model = "NY2150"
            steps['glue_model'] = glue_model
            glue_code = explicit_code
    steps['step1_glue_code'] = glue_code
    customer_grade_override = ''

    # ── 步骤2：厚度代码 ──
    mm_val, thick_raw, unit = extract_thickness_mm(s_text)
    exact_inch_mm = apply_customer_exact_inch_thickness(d_cust_name, thick_raw, glue_model or "")
    if exact_inch_mm is not None:
        mm_val = exact_inch_mm
        unit = 'inch-exact'
    steps['thickness_raw'] = thick_raw
    steps['thickness_mm'] = mm_val
    steps['thickness_unit'] = unit
    copper_spec = extract_copper_spec(s_text)
    steps['copper_spec_raw'] = copper_spec or '未识别'
    customer_order_override = find_customer_order_override(
        a_cust_code, d_cust_name, glue_model or '', mm_val, tables
    )
    if customer_order_override.get('error'):
        errors.append(customer_order_override['error'])
        glue_code = '??'
        steps['customer_order_rule_error'] = customer_order_override['error']
        steps['step1_glue_code'] = glue_code
    elif customer_order_override:
        if customer_order_override.get('glue_code'):
            glue_code = customer_order_override['glue_code']
            steps['step1_glue_code'] = glue_code
        if customer_order_override.get('grade_code'):
            customer_grade_override = customer_order_override['grade_code']
        steps['customer_order_rule'] = (
            f"客户下单转换表第{customer_order_override.get('row')}行："
            f"{customer_order_override.get('raw_glue')}"
        )
    jd_cy_thickness = lookup_jd_cy_mil_thickness(d_cust_name, s_text, thick_raw, copper_spec or '')
    front_core_back_supported = _normalize_glue_key(glue_model or "") == "NY3150HC"
    front_core_back_note = get_front_core_back_total_rule(
        d_cust_name, a_cust_code,
        tables.get('special_by_name', {}),
        tables.get('special_by_code', {})
    ) if front_core_back_supported else ""
    customer_front_core_back_note = ''
    if "崇达" in str(d_cust_name or "") and _normalize_glue_key(glue_model or "") == "NY3150HC":
        customer_front_core_back_note = "崇达：NY3150HC规格前面芯厚、后面总厚"
    front_core_back_total = (
        extract_front_core_back_total_thickness(s_text, copper_spec or '')
        if (front_core_back_note or customer_front_core_back_note) else None
    )

    if jd_cy_thickness:
        order_mm, special_mil, special_note = jd_cy_thickness
        order_is_total = special_mil >= 31
        mm_val = order_mm
        steps['thickness_mm'] = mm_val
        steps['thickness_unit'] = '健鼎/超颖mil表'
        steps['thickness_mode'] = '总厚' if order_is_total else '芯厚'
        steps['thickness_mode_source'] = '客户特殊板厚表'
        steps['thickness_mode_note'] = special_note
        steps['order_mm'] = order_mm
        steps['order_type'] = '总厚' if order_is_total else '芯厚'
        thick_code = thickness_to_code(order_mm)
    elif front_core_back_total:
        order_mm, total_raw = front_core_back_total
        order_is_total = True
        mm_val = order_mm
        steps['thickness_raw'] = total_raw
        steps['thickness_mm'] = mm_val
        steps['thickness_unit'] = 'customer-front-core-back-total'
        steps['thickness_mode'] = 'total'
        steps['thickness_mode_source'] = 'customer-special-requirement'
        steps['thickness_mode_note'] = front_core_back_note or customer_front_core_back_note
        steps['order_mm'] = order_mm
        steps['order_type'] = 'total'
        thick_code = thickness_to_code(order_mm)
    elif mm_val is None:
        errors.append('无法识别厚度')
        thick_code = '?????'
        order_is_total = True
        order_mm = 0.0
    else:
        mode = get_thickness_mode(s_text)
        mode_source = '规格原文' if mode != 'unknown' else '通用阈值'
        if mode == 'unknown':
            customer_mode, customer_mode_note = get_customer_thickness_mode_override(d_cust_name, s_text)
            if customer_mode:
                mode = customer_mode
                mode_source = '客户特殊需求'
                steps['thickness_mode_note'] = customer_mode_note
        if mode == 'unknown':
            special_mode, special_note = get_special_thickness_mode(
                d_cust_name, a_cust_code,
                tables.get('special_by_name', {}),
                tables.get('special_by_code', {})
            )
            if special_mode:
                mode = special_mode
                mode_source = '客户特殊需求'
                steps['thickness_mode_note'] = special_note
        steps['thickness_mode'] = '总厚' if mode == 'total' else '芯厚'
        steps['thickness_mode_source'] = mode_source
        order_mm, order_is_total = calc_order_thickness(
            mm_val, copper_spec, mode,
            tables['thick_total_to_core'],
            tables['thick_core_to_total']
        )
        steps['order_mm'] = order_mm
        steps['order_type'] = '总厚' if order_is_total else '芯厚'
        thick_code = thickness_to_code(order_mm)
    yidun_mil = re.search(
        r"(?<!\d)(\d+(?:\.\d+)?)\s*(?:±\s*\d+(?:\.\d+)?)?\s*MIL\b",
        f"{s_text or ''} {context_text or ''}",
        re.IGNORECASE,
    )
    if "广东依顿" in str(d_cust_name or "") and yidun_mil and float(yidun_mil.group(1)) >= 4:
        mil_value = float(yidun_mil.group(1))
        if not order_is_total and mil_value < 31:
            order_mm = mil_value * 0.0254
            steps['thickness_mode_note'] = "广东依顿31mil以下芯厚按MIL精确换算"
        else:
            order_mm = round(float(order_mm) + 1e-12, 2)
            steps['thickness_mode_note'] = "广东依顿总厚下单值按0.01mm四舍五入"
        thick_code = thickness_to_code(order_mm)
        steps['order_mm'] = order_mm
    steps['step2_thick_code'] = thick_code

    # ── 步骤3：铜箔规格代码 ──
    if not copper_spec:
        errors.append('无法识别铜箔规格')
        copper_code = '??'
    else:
        copper_code = copper_spec_to_code(copper_spec)
        if "广东依顿" in str(d_cust_name or "") and re.search(r'(?<![A-Z0-9])1\s*/\s*H\s*OZ\b', s_text.upper()):
            copper_code = "H1"
            steps['copper_note'] = '广东依顿特殊写法：1/HOZ按H/1'
        if "深万基隆" in str(d_cust_name or "") and copper_spec.upper() == "H/H":
            copper_code = "JJ"
            steps['copper_note'] = '深万基隆特殊规则：H/H按J/J出货'
        if "常熟斗山" in str(d_cust_name or "") and copper_spec.upper() == "12/12" and "RTF" in str(s_text or "").upper():
            copper_code = "TT"
            steps['copper_note'] = '常熟斗山12um/12um RTF按T/T编码'
    steps['step3_copper_code'] = copper_code

    # ── 步骤4：尺寸代码 ──
    customer_size = get_customer_size_override(d_cust_name, s_text)
    if customer_size:
        w, h, raw_size = customer_size
        steps['size_w'] = w
        steps['size_h'] = h
        size_code = size_to_code(w, h)
        steps['size_note'] = f'客户特殊尺寸映射：{raw_size}->{w:g}x{h:g}'
    else:
        size = extract_size(s_text)
        if not size:
            errors.append('无法识别尺寸')
            size_code = '????????'
        else:
            w, h = size
            steps['size_w'] = w
            steps['size_h'] = h
            sp_size = get_special_size(d_cust_name, w, h, tables['special_by_name'])
            if sp_size:
                size_code = sp_size
                steps['size_note'] = f'特殊尺寸→{sp_size}'
            else:
                size_code = size_to_code(w, h)
                steps['size_note'] = '标准尺寸'
    steps['step4_size_code'] = size_code

    # ── 步骤5：胶水类别代码 ──
    glue_cat = get_glue_category(glue_model or '', tables['glue_cat_map'], glue_code)
    steps['glue_category'] = glue_cat
    glue_cat_code = 'Y' if glue_cat == '普通' else 'R'
    steps['step5_glue_cat_code'] = glue_cat_code

    # ── 步骤6：铜箔类型代码 ──
    copper_type_code = get_copper_type_code(s_text)
    steps['step6_copper_type_code'] = copper_type_code

    # ── 步骤7：基板级别代码 ──
    grade_code = get_grade_code(s_text, d_cust_name, glue_model or '',
                                tables['special_by_name'], tables.get('grade_desc_to_code', {}))
    if customer_grade_override:
        grade_code = customer_grade_override
        steps['grade_note'] = '客户下单转换表覆盖基板等级'
    if (
        "深南" in str(d_cust_name or "")
        and _normalize_glue_key(glue_model or "") == "NYP5Q"
        and ("考试板" in str(s_text or "") or "90022-4" in str(s_text or ""))
    ):
        grade_code = "D3"
        steps['grade_note'] = '深南NY-P5Q考试板按确认样本使用D3等级'
    steps['step7_grade_code'] = grade_code

    # ── 步骤8：总/芯厚代码 ──
    tc_code = 'T' if order_is_total else 'C'
    steps['step8_tc_code'] = tc_code

    # ── 步骤9：结构代码（多数仍暂用*代替，已确认客户特殊规则先输出） ──
    if is_jd_cy_customer(d_cust_name):
        struct_code = 'E'
        steps['structure_note'] = '湖北健鼎/超颖电子基板结构码按E'
    else:
        struct_code = '*'
    steps['step9_struct_code'] = struct_code

    copper_code, grade_code, tc_code, struct_code = apply_structured_special_overrides(
        tables, a_cust_code, d_cust_name, s_text, context_text or '',
        glue_model or '', copper_spec or '',
        copper_code, grade_code, tc_code, struct_code, errors, steps
    )
    if "深万基隆" in str(d_cust_name or "") and "芯板" in str(s_text or "") and _normalize_glue_key(glue_model or '').startswith("NY2140"):
        grade_code = "A1"
        steps['step7_grade_code'] = grade_code
        steps['grade_note'] = '深万基隆芯板样本按A1，避免NY2140 F1泛化规则覆盖'

    # ── 步骤10：非普通胶系后缀 ──
    suffix = ''
    if glue_cat_code == 'R':
        suffix = 'XXXXXX'
    steps['step10_suffix'] = suffix

    # ── 汇总 ──
    has_error = bool(errors) or '??' in glue_code or '?????' in thick_code or '??' in copper_code or '????????' in size_code
    if has_error:
        full_code = ''
    else:
        full_code = (glue_code + thick_code + copper_code +
                     size_code + glue_cat_code + copper_type_code +
                     grade_code + tc_code + struct_code + suffix)
    steps['final_code'] = full_code
    steps['errors'] = errors

    err_str = '; '.join(errors)
    return full_code if full_code else ('未识别：' + err_str), steps, err_str


# ============================================================
# 主处理函数
# ============================================================

def process_file(filepath: str, output_path: str = None, rule_path: str = None) -> Tuple[pd.DataFrame, list]:
    sheets, tables = load_transcode_inputs(filepath, rule_path)
    df = sheets['转码需求表'].copy()
    spec_col = select_transcode_spec_column(df)
    customer_col = detect_customer_column(df, spec_col)
    customer_code_col = detect_customer_code_column(df)
    context_cols = detect_transcode_context_columns(df, spec_col, customer_col, customer_code_col)
    df, result_col = ensure_result_column(df)
    results = []

    for i in range(1, len(df)):
        row = df.iloc[i]
        a_val = _clean_cell(row.iloc[customer_code_col]) if customer_code_col is not None and len(row) > customer_code_col else ''
        d_val = str(row.iloc[customer_col]).strip() if customer_col is not None and len(row) > customer_col and pd.notna(row.iloc[customer_col]) else ''
        e_val = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ''
        s_val = _clean_cell(row.iloc[spec_col])
        context_val = build_context_text_from_row(row, context_cols)
        cust_spec_val = _clean_cell(row.iloc[6]) if len(row) > 6 else ''
        normalized_spec_val = _clean_cell(row.iloc[7]) if len(row) > 7 else ''
        pp_check_text = ' '.join([s_val, cust_spec_val, normalized_spec_val])

        if is_pp_or_rc_spec(pp_check_text):
            df.iloc[i, result_col] = ''
            results.append({'行号': i+1, '客户': d_val, '规格': (cust_spec_val or s_val)[:50], '编码': '', '错误': '', '状态': 'PP'})
            continue

        if not s_val or s_val == 'nan':
            results.append({'行号': i+1, '客户': d_val, '规格': '', '编码': '', '错误': '无规格', '状态': '跳过'})
            continue

        if is_pp_or_rc_spec(s_val):
            df.iloc[i, result_col] = ''
            results.append({'行号': i+1, '客户': d_val, '规格': s_val[:50], '编码': '', '错误': '', '状态': 'PP'})
            continue

        code, steps, err = transcode_row(s_val, e_val, d_val, a_val, tables, context_val)
        if steps.get('final_code'):
            df.iloc[i, result_col] = steps['final_code']

        results.append({
            '行号': i+1, '客户': d_val, '规格': s_val[:60],
            '编码': steps.get('final_code', ''),
            '错误': err, '状态': '成功' if not err else '部分失败',
            '步骤': steps
        })

    if output_path:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='转码需求表', index=False, header=False)
        print(f"结果已保存到: {output_path}")

    return df, results


if __name__ == '__main__':
    import sys
    fp = sys.argv[1] if len(sys.argv) > 1 else '/home/ubuntu/excel_auto_tool/营销转码需求功能表.xlsx'
    _, results = process_file(fp)
    for r in results:
        if r['状态'] in ('跳过', 'PP'):
            continue
        print(f"行{r['行号']} [{r['客户']}] → {r['编码']}")
        if r['错误']:
            print(f"  错误: {r['错误']}")
