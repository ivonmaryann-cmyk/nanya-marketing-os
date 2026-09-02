from __future__ import annotations

import copy
import re
from decimal import Decimal, InvalidOperation
from typing import Any

from .purchase_field_rules import (
    clean_text,
    extract_key_values,
    find_detail_header_row,
    header_score,
    looks_like_detail_data,
    map_detail_row,
    normalize_date,
    normalize_number,
)


SEQ = "序号"
CODE = "物料编码"
NAME = "物料名称"
DESC = "说明"
QTY = "数量"
UNIT = "单位"
PRICE = "含税单价"
AMOUNT = "金额"
DATE = "交货日期"
REMARK = "备注"
DETAIL_FIELDS = [SEQ, CODE, NAME, DESC, QTY, UNIT, PRICE, AMOUNT, DATE, REMARK]
PP_GLASS_CODES = (
    "0106",
    "1027",
    "1035",
    "1037",
    "1067",
    "1078",
    "1080",
    "1086",
    "1506",
    "2113",
    "2116",
    "2313",
    "3313",
    "7628",
    "106",
)

HEADER_KEYS = {
    "customer": "客户",
    "supplier": "供应商",
    "order": "订单号",
    "contract": "合同编号",
    "date": "日期",
    "currency": "币别",
    "ship_to": "交货地点",
    "payment": "付款方式",
    "tax": "税率",
    "contact": "联系人",
    "tel": "电话",
}


def _money(value: Any) -> Decimal | None:
    number = normalize_number(value)
    if not number:
        return None
    try:
        return Decimal(number)
    except InvalidOperation:
        return None


def _compact(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def normalize_pp_spec_spacing(value: Any) -> str:
    text = clean_text(value)
    compact_text = re.sub(r"\s+", "", text)
    if not compact_text:
        return text
    glass_pattern = "|".join(PP_GLASS_CODES)
    match = re.fullmatch(
        rf"(?P<product>NY[A-Z0-9()/_-]*?P)(?P<glass>{glass_pattern})(?P<rc>\d{{1,3}}(?:\.\d+)?%)(?P<length>\d+(?:\.\d+)?(?:米|M)(?:/卷|/R))",
        compact_text,
        flags=re.I,
    )
    if not match:
        return text
    return " ".join(match.group(name) for name in ["product", "glass", "rc", "length"])


def normalize_order_spec_spacing(value: Any) -> str:
    """Repair only high-confidence spacing artifacts found in order tables."""
    text = normalize_pp_spec_spacing(value)
    text = re.sub(r"(?i)(/卷)(?=(?:耐|无卤|有卤|CAF))", r"\1 ", text)
    text = re.sub(r"([纬经])\s+向", r"\1向", text)
    text = re.sub(r"(?i)(?<=\d)m\s+m(?=(?:纬|经)向)", "mm", text)
    text = re.sub(r"(?i)\bFR\s*-\s*4\b", "FR-4", text)
    text = re.sub(r"高\s+速(?=材料)", "高速", text)
    return clean_text(text)


def _split_header_body_text(source_text: str) -> tuple[list[str], list[str]]:
    lines = [clean_text(line) for line in str(source_text or "").splitlines() if clean_text(line)]
    if not lines:
        return [], []
    header_index = None
    for index, line in enumerate(lines):
        text = _compact(line)
        has_identity = any(keyword in text for keyword in ["物料编码", "物料编号", "原料编码", "partno", "goodsno", "料件编号"])
        has_value = any(keyword in text for keyword in ["数量", "单价", "金额", "交货日期", "到货日期", "quantity", "unitprice", "amount"])
        if has_identity and has_value:
            header_index = index
            break
    if header_index is None:
        return lines[:20], lines[20:]
    return lines[:header_index], lines[header_index:]


def _first_match(patterns: list[str], text: str, flags: int = re.I) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags=flags)
        if match:
            return clean_text(match.group(1))
    return ""


def _trim_to_next_label(value: str) -> str:
    value = clean_text(value)
    value = re.split(
        r"\s+(?:订单编号|订单号|采购单号|订购日期|采购日期|制表|页码|联系人|电话|传真|交货方式|结算条件|付款方式|税率|币别|供应商|供应商代码|VendorCode|P\.?ODate|Currency|TaxRate|PaymentTerms|ContactPerson|Telephone)\s*[:：/]?",
        value,
        maxsplit=1,
    )[0]
    return clean_text(value)


def _looks_polluted_header_value(value: str) -> bool:
    text = clean_text(value)
    if not text:
        return False
    bad_keywords = [
        "交易条款",
        "知识产权",
        "对账单",
        "逾期",
        "供应商回签",
        "数量 单价 金额",
        "物料名称、物料编号",
        "为月结者",
        "Item No",
        "Part No",
        "Description",
    ]
    if any(keyword in text for keyword in bad_keywords):
        return True
    return len(text) > 90 and any(mark in text for mark in ["，", "。", "；", ";"])


def _valid_tax(value: Any) -> bool:
    return bool(re.fullmatch(r"\d{1,2}(?:\.\d+)?%", clean_text(value)))


def _cleanup_header_info(result: dict[str, str], header_one_line: str) -> dict[str, str]:
    contract_key = HEADER_KEYS["contract"]
    order_key = HEADER_KEYS["order"]
    tax_key = HEADER_KEYS["tax"]

    has_contract_label = bool(re.search(r"合同(?:编号|号)|Contract\s*No", header_one_line, flags=re.I))
    has_order_label = bool(re.search(r"订单(?:编号|号)|采购单号|PO\s*NO|P\.?O\s*NO|P\.ONO", header_one_line, flags=re.I))
    if result.get(contract_key) and result.get(contract_key) == result.get(order_key):
        if has_contract_label and not has_order_label:
            result.pop(order_key, None)
        elif has_order_label and not has_contract_label:
            result.pop(contract_key, None)

    if result.get(tax_key) and not _valid_tax(result.get(tax_key)):
        fallback_tax = _first_match(
            [
                r"(?:税率|含税|税种|TaxRate|Tax\s*Rate|VAT)[^\d%]{0,12}(\d{1,2}(?:\.\d+)?\s*%?)",
                r"VAT\s*(\d{1,2})[_A-Z]*",
            ],
            header_one_line,
        )
        if fallback_tax:
            fallback_tax = fallback_tax.replace(" ", "")
            result[tax_key] = fallback_tax if fallback_tax.endswith("%") else f"{fallback_tax}%"
        else:
            result.pop(tax_key, None)
    return result


def extract_header_info_from_header_region(source_text: str, existing: dict[str, Any] | None = None) -> dict[str, str]:
    header_lines, _body_lines = _split_header_body_text(source_text)
    base = {key: clean_text(value) for key, value in (existing or {}).items() if clean_text(value)}
    if not header_lines:
        return {key: value for key, value in base.items() if not _looks_polluted_header_value(value)}

    header_text = "\n".join(header_lines)
    header_one_line = " ".join(header_lines)
    extracted = extract_key_values(header_lines)
    result = {key: clean_text(value) for key, value in extracted.items() if clean_text(value)}

    order = _first_match(
        [
            r"(?:订单编号|订单号|采购单号|采购订单号|PO\s*NO|P\.O\s*NO|P\.ONO)\s*[/A-Za-z（）() ]*[:：]\s*([A-Za-z0-9_-]{5,})",
            r"\b(PO[A-Z0-9_-]{6,})\b",
        ],
        header_one_line,
    )
    if order:
        result[HEADER_KEYS["order"]] = order

    contract = _first_match(
        [
            r"(?:合同编号|合同号|Contract\s*No\.?)\s*[:：]\s*([A-Za-z0-9_-]{5,})",
        ],
        header_one_line,
    )
    if contract:
        result[HEADER_KEYS["contract"]] = contract
        if result.get(HEADER_KEYS["order"]) == contract and "订单" not in header_text:
            result.pop(HEADER_KEYS["order"], None)

    date_value = _first_match(
        [
            r"(?:订购日期|采购日期|订单日期|合同日期|P\.?ODate|Date)\s*[:：]\s*([20][0-9][0-9][-/年.]\s*\d{1,2}[-/月.]\s*\d{1,2})",
        ],
        header_one_line,
    )
    if date_value:
        result[HEADER_KEYS["date"]] = normalize_date(date_value) or date_value

    payment = _first_match(
        [
            r"(?:结算条件|付款方式|付款条件|Payment\s*Terms?|PaymentTerms)\s*[:：]\s*([^\n]+?)(?=\s+(?:联系人|ContactPerson|电话|Telephone|传真|页码|币别|Currency|税率|TaxRate|交货|供应商|$))",
        ],
        header_one_line,
    )
    if payment:
        result[HEADER_KEYS["payment"]] = _trim_to_next_label(payment)

    tax = _first_match(
        [
            r"(?:税率|含税|税种|TaxRate|Tax\s*Rate|VAT)[^\d%]{0,12}(\d{1,2}(?:\.\d+)?\s*%?)",
            r"VAT\s*(\d{1,2})[_A-Z]*",
        ],
        header_one_line,
    )
    if tax:
        tax = tax.replace(" ", "")
        result[HEADER_KEYS["tax"]] = tax if tax.endswith("%") else f"{tax}%"

    supplier = _first_match(
        [
            r"(?:供应商名称|供应商|卖方|公司名称|Vendor)\s*[:：]\s*([^\n]+?)(?=\s+(?:订单|供应商代码|VendorCode|订购日期|P\.?ODate|制表|联系人|ContactPerson|电话|Telephone|$))",
        ],
        header_one_line,
    )
    if supplier:
        result[HEADER_KEYS["supplier"]] = _trim_to_next_label(supplier)

    ship_to = _first_match(
        [
            r"(?:交货地址|送货地址|交货地点|Ship\s*to)\s*[:：]\s*([^\n]+?)(?=\s+(?:报价单|货物接收|联系人|电话|$))",
        ],
        header_one_line,
    )
    if ship_to:
        result[HEADER_KEYS["ship_to"]] = _trim_to_next_label(ship_to)

    contact = _first_match([r"(?:联系人|收货人|ContactPerson)\s*[:：]\s*([^\s，,]+)"], header_one_line)
    if contact:
        result[HEADER_KEYS["contact"]] = contact

    tel = _first_match([r"(?:电话/TEL|电话|TEL|Telephone)\s*[:：]\s*([0-9+\- ]{6,})"], header_one_line)
    if tel:
        result[HEADER_KEYS["tel"]] = clean_text(tel)

    for key, value in base.items():
        result.setdefault(key, value)
    if (
        result.get(HEADER_KEYS["contract"])
        and result.get(HEADER_KEYS["contract"]) == result.get(HEADER_KEYS["order"])
        and not re.search(r"合同(?:编号|号)|Contract\s*No", header_one_line, flags=re.I)
    ):
        result.pop(HEADER_KEYS["contract"], None)
    result = _cleanup_header_info(result, header_one_line)
    cleaned = {key: _trim_to_next_label(value) for key, value in result.items()}
    return {key: value for key, value in cleaned.items() if clean_text(value) and not _looks_polluted_header_value(value)}


def _non_empty_indexes(row: list[str]) -> list[int]:
    return [index for index, value in enumerate(row) if clean_text(value)]


def _is_fragment_row(row: list[str]) -> bool:
    non_empty = _non_empty_indexes(row)
    if not non_empty or len(non_empty) > 2:
        return False
    text = " ".join(clean_text(row[index]) for index in non_empty)
    return bool(re.search(r"\(RMB|RMB\)?|^\)$|^[）)]$|^[A-Za-z%/]+$", text, flags=re.I))


def _is_opening_price_fragment(row: list[str]) -> bool:
    text = " ".join(clean_text(row[index]) for index in _non_empty_indexes(row))
    return bool(re.search(r"\d+(?:\.\d+)?\s*\(\s*RMB$", text, flags=re.I))


def _close_unbalanced_rmb_cells(row: list[str]) -> tuple[list[str], bool]:
    changed = False
    fixed = list(row)
    for index, value in enumerate(fixed):
        text = clean_text(value)
        if re.search(r"\d+(?:\.\d+)?\s*\(\s*RMB$", text, flags=re.I):
            fixed[index] = f"{text})"
            changed = True
    return fixed, changed


def _is_header_like_row(row: list[str]) -> bool:
    text = _compact(" ".join(row))
    if not text:
        return False
    score, _mapping = header_score(row)
    if score >= 3:
        return True
    return any(
        keyword in text
        for keyword in [
            "goodsno",
            "goodschapellation",
            "goodsapellation",
            "unitprice",
            "arrivaldate",
            "quantity",
            "物料编码物料名称",
        ]
    )


def _is_summary_or_section_row(row: list[str]) -> bool:
    text = clean_text(" ".join(row))
    if not text:
        return True
    return any(keyword in text for keyword in ["合计", "总计", "小计", "以下空白", "交易条款", "付款方式", "送货地址", "其它说明", "其他说明"])


def _row_field(mapping: dict[int, str], field: str) -> int | None:
    for index, mapped in mapping.items():
        if mapped == field:
            return index
    return None


def _format_fcf_material_spec(value: str) -> str:
    laminate = re.fullmatch(
        r"NY(?P<series>\d{4}[A-Z])(?P<thickness>\d+(?:\.\d+)?mm)(?P<rest>1/1\([^)）]{1,20}[)）]\d+(?:\.\d+)?[*×xX]\d+(?:\.\d+)?)",
        value,
        flags=re.I,
    )
    if laminate:
        return (
            f"NY {laminate.group('series')} {laminate.group('thickness')}\n"
            f"{laminate.group('rest')}"
        )

    pp = re.fullmatch(
        r"NYPP(?P<series>\d{4}[A-Z])(?P<glass>\d{3,4})(?P<rest>RC\d{1,3}(?:\.\d+)?%\d+(?:\.\d+)?[\"”]?[*×xX]\d+(?:\.\d+)?(?:米|M))",
        value,
        flags=re.I,
    )
    if pp:
        return f"NY PP {pp.group('series')} {pp.group('glass')}\n{pp.group('rest')}"
    return ""


def _split_duplicated_material_code_spec(code_value: str, name_value: str) -> tuple[str, str] | None:
    compact_code = re.sub(r"\s+", "", clean_text(code_value))
    compact_name = re.sub(r"\s+", "", clean_text(name_value))
    if not compact_code or compact_code != compact_name:
        return None

    patterns = [
        re.compile(
            r"^(?P<code>\d{5}NY\d{5}-\d{4}-[A-Z0-9]+-[A-Z0-9]+)(?P<spec>NY\d{4}[A-Z]\d+(?:\.\d+)?mm1/1\([^)）]{1,20}[)）]\d+(?:\.\d+)?[*×xX]\d+(?:\.\d+)?)$",
            flags=re.I,
        ),
        re.compile(
            r"^(?P<code>NYPP\d{3,4}-\d{1,3}-\d{4})(?P<spec>NYPP\d{4}[A-Z]\d{3,4}RC\d{1,3}(?:\.\d+)?%\d+(?:\.\d+)?[\"”]?[*×xX]\d+(?:\.\d+)?(?:米|M))$",
            flags=re.I,
        ),
    ]
    for pattern in patterns:
        match = pattern.fullmatch(compact_code)
        if not match:
            continue
        formatted_spec = _format_fcf_material_spec(match.group("spec"))
        if formatted_spec:
            return match.group("code"), formatted_spec
    return None


def _repair_duplicated_code_spec_columns(
    rows: list[list[str]],
    header_index: int,
    mapping: dict[int, str],
) -> tuple[list[list[str]], list[str]]:
    code_col = _row_field(mapping, CODE)
    name_col = _row_field(mapping, NAME)
    if code_col is None or name_col is None:
        return rows, []

    normalized = [list(row) for row in rows]
    changed = False
    for row in normalized[header_index + 1 :]:
        if max(code_col, name_col) >= len(row):
            continue
        split = _split_duplicated_material_code_spec(row[code_col], row[name_col])
        if not split:
            continue
        row[code_col], row[name_col] = split
        changed = True
    warning = "已拆分重复粘连的物料编号与物料名/规格，并恢复规格分行。"
    return normalized, [warning] if changed else []


def _is_unit_token(value: Any) -> bool:
    text = clean_text(value).lower()
    return text in {
        "个",
        "件",
        "卷",
        "张",
        "片",
        "套",
        "箱",
        "包",
        "桶",
        "支",
        "只",
        "米",
        "公斤",
        "千克",
        "吨",
        "kg",
        "pcs",
        "pc",
    }


def _repair_shifted_unit_quantity_columns(
    rows: list[list[str]],
    header_index: int,
    mapping: dict[int, str],
) -> tuple[list[list[str]], list[str]]:
    headers = rows[header_index]
    requirement_cols = [
        index
        for index, header in enumerate(headers)
        if any(keyword in _compact(header) for keyword in ["物料要求", "产品要求", "materialrequirement"])
    ]
    unit_col = _row_field(mapping, UNIT)
    qty_col = _row_field(mapping, QTY)
    price_col = _row_field(mapping, PRICE)
    amount_col = _row_field(mapping, AMOUNT)
    if len(requirement_cols) != 1 or unit_col is None or qty_col is None:
        return rows, []

    requirement_col = requirement_cols[0]
    normalized = [list(row) for row in rows]
    changed = False
    for row in normalized[header_index + 1 :]:
        required_width = max(requirement_col, unit_col, qty_col, price_col or 0, amount_col or 0) + 1
        if len(row) < required_width:
            row.extend([""] * (required_width - len(row)))
        requirement = clean_text(row[requirement_col])
        unit = clean_text(row[unit_col])
        quantity = clean_text(row[qty_col])
        if not _is_unit_token(requirement):
            continue
        if not unit:
            row[requirement_col] = ""
            row[unit_col] = requirement
            changed = True
            continue
        shifted_quantity = _money(unit)
        if quantity or shifted_quantity is None or shifted_quantity <= 0:
            continue
        price = _money(row[price_col]) if price_col is not None else None
        amount = _money(row[amount_col]) if amount_col is not None else None
        if price is None or amount is None or abs(shifted_quantity * price - amount) > Decimal("0.05"):
            continue
        row[requirement_col] = ""
        row[unit_col] = requirement
        row[qty_col] = normalize_number(unit)
        changed = True
    warning = "已按单位词及数量×单价=金额关系校正左移的单位/数量。"
    return normalized, [warning] if changed else []


def _split_concatenated_quantity_price(value: Any, amount: Decimal) -> tuple[str, str] | None:
    compact = re.sub(r"[\s,]", "", clean_text(value))
    match = re.fullmatch(r"(\d{2,})(\.\d+)?", compact)
    if not match or amount <= 0:
        return None

    integer_part = match.group(1)
    decimal_part = match.group(2) or ""
    candidates: dict[tuple[Decimal, Decimal], tuple[str, str]] = {}
    for split_at in range(1, len(integer_part)):
        quantity_text = integer_part[:split_at]
        if len(quantity_text) > 1 and quantity_text.startswith("0"):
            continue
        price_text = f"{integer_part[split_at:]}{decimal_part}"
        quantity = Decimal(quantity_text)
        price = Decimal(price_text)
        if not (Decimal("1") <= quantity <= Decimal("100000")) or price <= 0:
            continue
        if abs(quantity * price - amount) > Decimal("0.05"):
            continue
        candidates[(quantity, price)] = (normalize_number(quantity_text), normalize_number(price_text))

    if len(candidates) != 1:
        return None
    return next(iter(candidates.values()))


def _repair_concatenated_quantity_price_columns(
    rows: list[list[str]],
    header_index: int,
    mapping: dict[int, str],
) -> tuple[list[list[str]], list[str]]:
    qty_col = _row_field(mapping, QTY)
    price_col = _row_field(mapping, PRICE)
    amount_col = _row_field(mapping, AMOUNT)
    if qty_col is None or price_col is None or amount_col is None:
        return rows, []

    normalized = [list(row) for row in rows]
    changed = False
    required_width = max(qty_col, price_col, amount_col) + 1
    for row in normalized[header_index + 1 :]:
        if len(row) < required_width:
            row.extend([""] * (required_width - len(row)))
        if clean_text(row[qty_col]):
            continue
        amount = _money(row[amount_col])
        if amount is None:
            continue
        split = _split_concatenated_quantity_price(row[price_col], amount)
        if split is None:
            continue
        row[qty_col], row[price_col] = split
        changed = True

    warning = "已按数量×单价=金额的唯一解拆分粘连的数量和单价。"
    return normalized, [warning] if changed else []


def _append_code_text(existing: str, addition: str) -> str:
    return re.sub(r"\s+", "", f"{clean_text(existing)}{clean_text(addition)}")


def _append_cell_text(existing: str, addition: str, *, field: str | None = None) -> str:
    existing = clean_text(existing)
    addition = clean_text(addition)
    if not addition:
        return existing
    if not existing:
        return _append_code_text("", addition) if field == CODE else addition
    if field == CODE:
        return _append_code_text(existing, addition)
    if addition in {")", "）"} and existing.endswith("("):
        return f"{existing}{addition}"
    if addition in {")", "）"}:
        return f"{existing}{addition}"
    if existing.endswith("(") or existing.endswith("（"):
        return f"{existing}{addition}"
    return clean_text(f"{existing} {addition}")


def _append_row_cell_text(existing: str, addition: str, column: int, mapping: dict[int, str]) -> str:
    return _append_cell_text(existing, addition, field=mapping.get(column))


def _merge_fragment_rows(rows: list[list[str]], header_index: int, mapping: dict[int, str]) -> tuple[list[list[str]], list[str]]:
    if header_index is None or header_index >= len(rows):
        return rows, []
    warnings: list[str] = []
    price_col = _row_field(mapping, PRICE)
    if price_col is None:
        return rows, warnings

    normalized: list[list[str]] = rows[: header_index + 1]
    pending: dict[int, str] = {}
    index = header_index + 1
    while index < len(rows):
        row = list(rows[index])
        if _is_fragment_row(row):
            for column in _non_empty_indexes(row):
                pending[column] = _append_row_cell_text(pending.get(column, ""), row[column], column, mapping)
            index += 1
            continue

        if pending:
            for column, value in pending.items():
                if column >= len(row):
                    row.extend([""] * (column + 1 - len(row)))
                row[column] = _append_row_cell_text(value, row[column], column, mapping) if row[column] else value
            pending = {}

        lookahead = index + 1
        while lookahead < len(rows) and _is_fragment_row(rows[lookahead]) and not _is_opening_price_fragment(rows[lookahead]):
            for column in _non_empty_indexes(rows[lookahead]):
                if column >= len(row):
                    row.extend([""] * (column + 1 - len(row)))
                row[column] = _append_row_cell_text(row[column], rows[lookahead][column], column, mapping)
            lookahead += 1

        if lookahead != index + 1:
            warnings.append("已合并 PDF 表格中的单元格碎片行。")
        row, closed_rmb = _close_unbalanced_rmb_cells(row)
        if closed_rmb:
            warnings.append("已修复未闭合的 RMB 单价括号。")
        normalized.append(row)
        index = lookahead

    return normalized, warnings


def _merge_multiline_detail_rows(rows: list[list[str]], header_index: int, mapping: dict[int, str]) -> tuple[list[list[str]], list[str]]:
    if header_index is None or header_index >= len(rows):
        return rows, []
    warnings: list[str] = []
    seq_col = _row_field(mapping, SEQ)
    code_col = _row_field(mapping, CODE)
    name_col = _row_field(mapping, NAME)
    qty_col = _row_field(mapping, QTY)
    amount_col = _row_field(mapping, AMOUNT)
    date_col = _row_field(mapping, DATE)

    merged: list[list[str]] = rows[: header_index + 1]
    for row in rows[header_index + 1 :]:
        current = list(row)
        if not merged or len(merged) <= header_index:
            merged.append(current)
            continue
        if _is_summary_or_section_row(current):
            merged.append(current)
            continue
        previous = merged[-1]
        previous_is_detail = looks_like_detail_data(previous)
        has_current_anchor = any(
            column is not None and column < len(current) and clean_text(current[column])
            for column in [qty_col, amount_col, date_col, _row_field(mapping, PRICE)]
        )
        has_current_seq = seq_col is not None and seq_col < len(current) and clean_text(current[seq_col])
        has_current_code = code_col is not None and code_col < len(current) and clean_text(current[code_col])
        if previous_is_detail and not has_current_seq and not has_current_code and not has_current_anchor:
            for column in [code_col, name_col]:
                if column is None or column >= len(current) or not clean_text(current[column]):
                    continue
                if column >= len(previous):
                    previous.extend([""] * (column + 1 - len(previous)))
                previous[column] = _append_row_cell_text(previous[column], current[column], column, mapping)
            warnings.append("已合并跨行明细规格文本。")
        else:
            merged.append(current)
    return merged, warnings


def _compact_material_code_column(rows: list[list[str]], header_index: int, mapping: dict[int, str]) -> tuple[list[list[str]], list[str]]:
    code_col = _row_field(mapping, CODE)
    if code_col is None:
        return rows, []
    normalized = [list(row) for row in rows]
    changed = False
    for row in normalized[header_index + 1 :]:
        if code_col >= len(row):
            continue
        original = clean_text(row[code_col])
        compacted = re.sub(r"\s+", "", original)
        if compacted != original:
            row[code_col] = compacted
            changed = True
    return normalized, ["已清理物料编码列内部空格。"] if changed else []


def _normalize_spec_columns(rows: list[list[str]], header_index: int) -> tuple[list[list[str]], list[str]]:
    normalized = [list(row) for row in rows]
    headers = normalized[header_index]
    spec_columns = [
        index
        for index, header in enumerate(headers)
        if "规格" in _compact(header)
    ]
    if not spec_columns:
        return normalized, []
    changed = False
    for row in normalized[header_index + 1 :]:
        for column in spec_columns:
            if column >= len(row):
                continue
            original = clean_text(row[column])
            formatted = normalize_order_spec_spacing(original)
            if formatted and formatted != original:
                row[column] = formatted
                changed = True
    return normalized, ["已恢复可确认 PP 型号/规格的语义空格。"] if changed else []


def _normalize_mapped_detail_specs(rows: list[dict[str, Any]]) -> bool:
    changed = False
    for row in rows:
        standard = row.get("standard") or {}
        description = clean_text(standard.get(DESC))
        formatted_description = normalize_order_spec_spacing(description)
        if formatted_description and formatted_description != description:
            standard[DESC] = formatted_description
            changed = True
        for header, value in (row.get("original") or {}).items():
            if "规格" not in _compact(header) and "备注" not in _compact(header):
                continue
            formatted_value = normalize_order_spec_spacing(value)
            if formatted_value and formatted_value != clean_text(value):
                row["original"][header] = formatted_value
                changed = True
        remark = clean_text(standard.get(REMARK))
        formatted_remark = normalize_order_spec_spacing(remark)
        if formatted_remark and formatted_remark != remark:
            standard[REMARK] = formatted_remark
            changed = True
    return changed


def _normalize_table_rows(rows: list[list[str]]) -> tuple[list[list[str]], dict[int, str], list[str]]:
    rows = [[clean_text(value) for value in row] for row in rows if any(clean_text(value) for value in row)]
    header_index, mapping = find_detail_header_row(rows)
    if header_index is None:
        return rows, {}, []
    rows, fragment_warnings = _merge_fragment_rows(rows, header_index, mapping)
    rows, multiline_warnings = _merge_multiline_detail_rows(rows, header_index, mapping)
    rows, duplicate_warnings = _repair_duplicated_code_spec_columns(rows, header_index, mapping)
    rows, shifted_warnings = _repair_shifted_unit_quantity_columns(rows, header_index, mapping)
    rows, concatenated_warnings = _repair_concatenated_quantity_price_columns(rows, header_index, mapping)
    rows, code_warnings = _compact_material_code_column(rows, header_index, mapping)
    rows, spec_warnings = _normalize_spec_columns(rows, header_index)
    normalized = rows[: header_index + 1]
    for row in rows[header_index + 1 :]:
        if _is_header_like_row(row) or _is_summary_or_section_row(row):
            continue
        if not looks_like_detail_data(row) and len(_non_empty_indexes(row)) <= 2:
            continue
        normalized.append(row)
    return (
        normalized,
        mapping,
        fragment_warnings
        + multiline_warnings
        + duplicate_warnings
        + shifted_warnings
        + concatenated_warnings
        + code_warnings
        + spec_warnings,
    )


def _subtotal_values(document: dict[str, Any], source_text: str) -> list[Decimal]:
    text_parts = [str(source_text or "")]
    for page in document.get("pages") or []:
        text_parts.extend(str(line) for line in page.get("text_lines") or [])
    for lines in (document.get("sections") or {}).values():
        text_parts.extend(str(line) for line in lines or [])
    values: list[Decimal] = []
    pattern = re.compile(
        r"(?:小计|合计|sub\s*-?\s*total)\s*(?:\([^)]*\))?\s*[:：]?\s*(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)",
        flags=re.I,
    )
    for match in pattern.finditer("\n".join(text_parts)):
        try:
            values.append(Decimal(match.group(1).replace(",", "")))
        except InvalidOperation:
            continue
    return sorted(set(values))


def _unique_positive_integer_solution(remaining: Decimal, prices: list[Decimal]) -> list[int] | None:
    if remaining <= 0 or not prices or len(prices) > 2 or any(price <= 0 for price in prices):
        return None
    if len(prices) == 1:
        quantity = remaining / prices[0]
        if quantity == quantity.to_integral_value() and Decimal(1) <= quantity <= Decimal(10000):
            return [int(quantity)]
        return None

    solutions: list[list[int]] = []
    max_first = min(int(remaining / prices[0]), 10000)
    for first in range(1, max_first + 1):
        rest = remaining - prices[0] * first
        if rest <= 0:
            break
        second = rest / prices[1]
        if second == second.to_integral_value() and Decimal(1) <= second <= Decimal(10000):
            solutions.append([first, int(second)])
            if len(solutions) > 1:
                return None
    return solutions[0] if len(solutions) == 1 else None


def _infer_missing_quantities_from_subtotal(
    document: dict[str, Any],
    source_text: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    tables = document.get("raw_detail_tables") or []
    subtotals = _subtotal_values(document, source_text)
    if len(tables) != 1 or len(subtotals) != 1:
        return [], []
    table = tables[0]
    rows = table.get("rows") or []
    header_index, mapping = find_detail_header_row(rows)
    if header_index is None:
        return [], []
    qty_columns = [column for column, field in mapping.items() if field == QTY]
    if len(qty_columns) != 1:
        return [], []
    qty_column = qty_columns[0]

    known_total = Decimal(0)
    missing: list[tuple[int, Decimal]] = []
    for row_index, row in enumerate(rows[header_index + 1 :], start=header_index + 1):
        if _is_header_like_row(row) or _is_summary_or_section_row(row):
            continue
        standard = map_detail_row(rows[header_index], row, mapping).get("standard") or {}
        if not clean_text(standard.get(CODE)) and not clean_text(standard.get(NAME)):
            continue
        price = _money(standard.get(PRICE))
        amount = _money(standard.get(AMOUNT))
        quantity = _money(standard.get(QTY))
        if amount is not None:
            known_total += amount
        elif quantity is not None and price is not None:
            known_total += quantity * price
        elif quantity is None and price is not None:
            missing.append((row_index, price))
        else:
            return [], []
    solution = _unique_positive_integer_solution(subtotals[0] - known_total, [price for _row, price in missing])
    if not solution or len(solution) != len(missing):
        return [], []

    actions: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    for (row_index, _price), quantity in zip(missing, solution):
        while len(rows[row_index]) <= qty_column:
            rows[row_index].append("")
        rows[row_index][qty_column] = str(quantity)
        action = {
            "type": "subtotal_unique_quantity",
            "row_index": row_index,
            "column_index": qty_column,
            "field": QTY,
            "value": str(quantity),
            "subtotal": str(subtotals[0]),
        }
        actions.append(action)
        issues.append(
            {
                "page_index": table.get("page_index", 0),
                "region": "明细表",
                "field": QTY,
                "raw_value": " ".join(rows[row_index]),
                "clean_value": str(quantity),
                "confidence": "唯一解",
                "message": f"数量由小计 {subtotals[0]} 与各行单价关系的唯一正整数解补齐。",
            }
        )
    table.setdefault("recovery_actions", []).extend(actions)
    return actions, issues


def _rebuild_detail_rows_from_tables(document: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    mapped_rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    for table in document.get("raw_detail_tables") or []:
        rows = table.get("rows") or []
        header_index, mapping = find_detail_header_row(rows)
        if header_index is None:
            continue
        raw_headers = rows[header_index]
        for row_offset, row in enumerate(rows[header_index + 1 :], start=header_index + 1):
            if not any(clean_text(value) for value in row):
                continue
            if _is_header_like_row(row) or _is_summary_or_section_row(row):
                continue
            mapped = map_detail_row(raw_headers, row, mapping)
            standard = mapped.get("standard") or {}
            q = _money(standard.get(QTY))
            price = _money(standard.get(PRICE))
            amount = _money(standard.get(AMOUNT))
            if amount is not None and q is not None and price is None:
                issues.append(
                    {
                        "page_index": table.get("page_index", 0),
                        "region": "明细表",
                        "field": PRICE,
                        "raw_value": " ".join(row),
                        "clean_value": "",
                        "confidence": table.get("confidence", 0),
                        "message": "明细行存在数量和金额，但单价为空，需复核。",
                    }
                )
            if q is not None and price is not None and amount is not None and abs(q * price - amount) > Decimal("0.05"):
                issues.append(
                    {
                        "page_index": table.get("page_index", 0),
                        "region": "明细表",
                        "field": AMOUNT,
                        "raw_value": " ".join(row),
                        "clean_value": str(q * price),
                        "confidence": table.get("confidence", 0),
                        "message": "数量×单价与金额不一致，需复核。",
                    }
                )
            mapped.update(
                {
                    "page_index": table.get("page_index", 0),
                    "table_index": table.get("table_index", 0),
                    "row_index": row_offset,
                    "raw_text": " ".join(row),
                    "confidence": table.get("confidence", 0),
                    "method": f"{table.get('method') or 'normalized'}_normalized",
                }
            )
            mapped_rows.append(mapped)
    return mapped_rows, issues


def _number_spans(text: str) -> list[dict[str, Any]]:
    spans: list[dict[str, Any]] = []
    for match in re.finditer(r"-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text):
        raw = match.group(0)
        try:
            value = Decimal(raw.replace(",", ""))
        except InvalidOperation:
            continue
        spans.append({"raw": raw, "value": value, "start": match.start(), "end": match.end()})
    return spans


def _find_amount_qty_price(text: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]] | None:
    numbers = _number_spans(text)
    amount_candidates = [
        item
        for item in numbers
        if abs(item["value"]) >= Decimal("1000") and ("." in item["raw"] or abs(item["value"]) >= Decimal("10000"))
    ]
    best: tuple[Decimal, dict[str, Any], dict[str, Any], dict[str, Any]] | None = None
    for amount in amount_candidates:
        for qty in numbers:
            if qty is amount or qty["value"] <= 0 or qty["value"] > Decimal("100000"):
                continue
            if qty["value"] != qty["value"].to_integral_value():
                continue
            for price in numbers:
                if price is amount or price is qty or price["value"] <= 0:
                    continue
                if abs(qty["value"] * price["value"] - amount["value"]) > Decimal("0.05"):
                    continue
                distance = Decimal(abs(qty["start"] - amount["start"]) + abs(price["start"] - amount["start"]))
                score = amount["value"] * Decimal("100000") - distance
                if best is None or score > best[0]:
                    best = (score, amount, qty, price)
    if best is None:
        return None
    _score, amount, qty, price = best
    return amount, qty, price


def _token_spans(text: str) -> list[dict[str, Any]]:
    return [
        {"raw": match.group(0), "start": match.start(), "end": match.end()}
        for match in re.finditer(r"\S+", text)
    ]


def _token_near_span(tokens: list[dict[str, Any]], span: dict[str, Any], *, before: bool) -> str:
    if before:
        candidates = [token for token in tokens if token["end"] <= span["start"]]
        candidates.sort(key=lambda token: token["end"], reverse=True)
    else:
        candidates = [token for token in tokens if token["start"] >= span["end"]]
        candidates.sort(key=lambda token: token["start"])
    for token in candidates[:2]:
        raw = clean_text(token["raw"])
        if re.fullmatch(r"[\u4e00-\u9fffA-Za-z]{1,6}", raw) and not normalize_number(raw):
            return raw
    return ""


def _extract_unit(text: str, qty_span: dict[str, Any]) -> str:
    tokens = _token_spans(text)
    nearby = _token_near_span(tokens, qty_span, before=False) or _token_near_span(tokens, qty_span, before=True)
    if nearby:
        return nearby
    match = re.search(r"(?<!\S)(张|卷|片|PCS|pcs|SET|set|㎡|米|kg|KG|个|套|包|公斤|平方米)(?!\S)", text)
    return match.group(1) if match else ""


def _extract_material_code(text: str) -> str:
    candidates = re.findall(r"\b[A-Za-z]{1,8}[A-Za-z0-9-]*\d[A-Za-z0-9-]*\b|\b\d+-\d+(?:-\d+)+\b", text)
    filtered = []
    for candidate in candidates:
        upper = candidate.upper()
        if upper in {"FR-4", "NYA1", "RMB"}:
            continue
        if len(candidate) < 4:
            continue
        filtered.append(candidate)
    if not filtered:
        return ""
    filtered.sort(key=lambda item: (not re.match(r"^[A-Za-z]\d", item), len(item)))
    return filtered[0]


def _remove_span_text(text: str, spans: list[dict[str, Any]]) -> str:
    pieces = []
    cursor = 0
    for span in sorted(spans, key=lambda item: item["start"]):
        pieces.append(text[cursor : span["start"]])
        cursor = span["end"]
    pieces.append(text[cursor:])
    return clean_text(" ".join("".join(pieces).split()))


def _looks_like_packed_header(text: str) -> bool:
    compact = _compact(text)
    return (
        any(keyword in compact for keyword in ["物料编码", "物料编号", "原料编码", "partno", "goodsno"])
        and any(keyword in compact for keyword in ["数量", "单价", "金额", "amount", "price"])
    )


def _packed_row_to_detail(row: list[str], *, row_index: int, table: dict[str, Any], sequence: int) -> dict[str, Any] | None:
    text = clean_text(" ".join(row))
    if not text or _looks_like_packed_header(text) or _is_summary_or_section_row([text]):
        return None
    trio = _find_amount_qty_price(text)
    if trio is None:
        return None
    amount_span, qty_span, price_span = trio
    code = _extract_material_code(text)
    unit = _extract_unit(text, qty_span)
    name_text = _remove_span_text(text, [amount_span, qty_span, price_span])
    if code:
        name_text = clean_text(re.sub(rf"\b{re.escape(code)}\b", " ", name_text, count=1))
    if unit:
        name_text = clean_text(re.sub(rf"(?<!\S){re.escape(unit)}(?!\S)", " ", name_text))
    standard = {field: "" for field in DETAIL_FIELDS}
    standard.update(
        {
            SEQ: str(sequence),
            CODE: code,
            NAME: name_text,
            QTY: normalize_number(qty_span["raw"]),
            UNIT: unit,
            PRICE: normalize_number(price_span["raw"]),
            AMOUNT: normalize_number(amount_span["raw"]),
        }
    )
    return {
        "original": {"原始正文行": text},
        "standard": standard,
        "page_index": table.get("page_index", 0),
        "table_index": table.get("table_index", 0),
        "row_index": row_index,
        "raw_text": text,
        "confidence": table.get("confidence", 0),
        "method": f"{table.get('method') or 'raw'}_packed_text_rebuild",
    }


def _rebuild_rows_from_packed_text(document: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    mapped_rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    sequence = 1
    for table in document.get("raw_detail_tables") or []:
        rows = table.get("rows") or []
        table_text = "\n".join(" ".join(row) for row in rows[:6])
        if not _looks_like_packed_header(table_text):
            continue
        for row_index, row in enumerate(rows):
            mapped = _packed_row_to_detail(row, row_index=row_index, table=table, sequence=sequence)
            if not mapped:
                continue
            mapped_rows.append(mapped)
            sequence += 1
    if mapped_rows:
        issues.append(
            {
                "page_index": mapped_rows[0].get("page_index", 0),
                "region": "明细表",
                "field": "表格列",
                "raw_value": "",
                "clean_value": f"重建 {len(mapped_rows)} 行",
                "confidence": "",
                "message": "检测到多列内容挤在同一列，已按订单正文行中的数量/单价/金额关系重建明细。",
            }
        )
    return mapped_rows, issues


def _row_quality(row: dict[str, Any]) -> int:
    standard = row.get("standard") or {}
    score = 0
    for field in [CODE, NAME, QTY, UNIT, PRICE, AMOUNT, DATE]:
        if clean_text(standard.get(field)):
            score += 1
    q = _money(standard.get(QTY))
    price = _money(standard.get(PRICE))
    amount = _money(standard.get(AMOUNT))
    if q is not None and price is not None and amount is not None and abs(q * price - amount) <= Decimal("0.05"):
        score += 2
    return score


def _strong_row_count(rows: list[dict[str, Any]]) -> int:
    count = 0
    for row in rows:
        standard = row.get("standard") or {}
        if clean_text(standard.get(CODE)) and clean_text(standard.get(QTY)) and clean_text(standard.get(AMOUNT)) and clean_text(standard.get(DATE)):
            count += 1
    return count


def _strong_name_score(rows: list[dict[str, Any]]) -> int:
    score = 0
    for row in rows:
        standard = row.get("standard") or {}
        if not clean_text(standard.get(CODE)):
            continue
        if not any(clean_text(standard.get(field)) for field in [QTY, AMOUNT, DATE]):
            continue
        score += len(clean_text(standard.get(NAME)))
    return score


def _strong_code_score(rows: list[dict[str, Any]]) -> int:
    score = 0
    for row in rows:
        standard = row.get("standard") or {}
        code = clean_text(standard.get(CODE))
        if not code:
            continue
        if not any(clean_text(standard.get(field)) for field in [QTY, AMOUNT, DATE]):
            continue
        score += len(re.sub(r"\s+", "", code))
    return score


def _should_replace_mapped_rows(old_rows: list[dict[str, Any]], new_rows: list[dict[str, Any]]) -> bool:
    if not new_rows:
        return False
    if not old_rows:
        return True
    old_strong = _strong_row_count(old_rows)
    new_strong = _strong_row_count(new_rows)
    if new_strong < old_strong:
        old_name_score = _strong_name_score(old_rows)
        new_name_score = _strong_name_score(new_rows)
        if len(new_rows) < len(old_rows) and new_name_score > old_name_score * 1.25:
            return True
        return False
    if new_strong > old_strong:
        return True
    if new_strong == old_strong and _strong_code_score(new_rows) > _strong_code_score(old_rows) * 1.25:
        return True
    old_score = sum(_row_quality(row) for row in old_rows)
    new_score = sum(_row_quality(row) for row in new_rows)
    if new_score > old_score:
        return True
    if new_strong == old_strong and _strong_name_score(new_rows) > _strong_name_score(old_rows) * 1.25:
        return True
    return new_strong == old_strong and len(new_rows) < len(old_rows)


def _is_mapped_non_detail_row(row: dict[str, Any]) -> bool:
    standard = row.get("standard") or {}
    text_parts = [str(value) for value in standard.values()]
    text_parts.extend(str(value) for value in row.get("raw_values") or [])
    text = clean_text(" ".join(text_parts))
    if not text:
        return True

    has_detail_anchor = any(clean_text(standard.get(field)) for field in [QTY, PRICE, AMOUNT, DATE])
    section_keywords = [
        "装运方式",
        "付款方式",
        "供应商签章",
        "Vendor Singnature",
        "Vendor Signature",
        "Shipments Method",
        "以下空白",
        "上述产品",
        "ROHS",
    ]
    if not has_detail_anchor and any(keyword.lower() in text.lower() for keyword in section_keywords):
        return True

    code = clean_text(standard.get(CODE))
    if not has_detail_anchor and code and any(keyword.lower() in code.lower() for keyword in section_keywords):
        return True

    seq = clean_text(standard.get(SEQ))
    name = clean_text(standard.get(NAME))
    has_sequence = bool(re.fullmatch(r"\d+(?:\.\d+)?", seq))
    has_code = bool(re.search(r"[A-Za-z]{1,8}\d{2,}|\d{3,}[A-Za-z]", code)) or bool(re.fullmatch(r"\d{5,}", code))
    qty = _money(standard.get(QTY))
    price = _money(standard.get(PRICE))
    amount = _money(standard.get(AMOUNT))
    date = clean_text(standard.get(DATE))
    if not (has_sequence or has_code) and not (name and qty is not None and (price is not None or amount is not None or date)):
        return True
    if qty is None and price is None and amount is None and not date:
        return True

    non_empty_standard = [value for value in standard.values() if clean_text(value)]
    return not has_detail_anchor and len(non_empty_standard) <= 1


def _filter_mapped_detail_rows(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    filtered = [row for row in rows if not _is_mapped_non_detail_row(row)]
    return filtered, len(rows) - len(filtered)


def _infer_uniform_tax_from_tables(tables: list[dict[str, Any]]) -> str:
    values: list[str] = []
    for table in tables:
        rows = table.get("rows") or []
        header_index, _mapping = find_detail_header_row(rows)
        if header_index is None:
            continue
        headers = rows[header_index]
        tax_columns = [
            index
            for index, header in enumerate(headers)
            if any(keyword in _compact(header) for keyword in ["税率", "taxrate", "tax"])
        ]
        for row in rows[header_index + 1 :]:
            for column in tax_columns:
                if column >= len(row):
                    continue
                text = clean_text(row[column])
                match = re.search(r"(\d{1,2}(?:\.\d+)?)\s*%?", text)
                if match:
                    value = match.group(1)
                    values.append(value if value.endswith("%") else f"{value}%")
    unique = sorted(set(values))
    return unique[0] if len(unique) == 1 else ""


def normalize_purchase_document(document: dict[str, Any], *, source_text: str = "") -> dict[str, Any]:
    normalized = copy.deepcopy(document)
    warnings: list[str] = list(normalized.get("normalization_warnings") or [])

    normalized_tables: list[dict[str, Any]] = []
    for table in normalized.get("raw_detail_tables") or []:
        table_copy = dict(table)
        rows, _mapping, table_warnings = _normalize_table_rows(table_copy.get("rows") or [])
        table_copy["rows"] = rows
        normalized_tables.append(table_copy)
        warnings.extend(table_warnings)
    normalized["raw_detail_tables"] = normalized_tables

    quantity_actions, quantity_issues = _infer_missing_quantities_from_subtotal(normalized, source_text)
    recovery_actions = [
        action
        for table in normalized.get("raw_detail_tables") or []
        for action in table.get("recovery_actions") or []
    ]
    if recovery_actions:
        normalized["recovery_actions"] = recovery_actions
    if quantity_actions:
        warnings.append("已按小计与单价关系的唯一正整数解补齐缺失数量。")

    rebuilt_rows, validation_issues = _rebuild_detail_rows_from_tables(normalized)
    validation_issues = quantity_issues + validation_issues
    old_rows = list(normalized.get("mapped_detail_rows") or [])
    if not rebuilt_rows and not old_rows:
        rebuilt_rows, packed_issues = _rebuild_rows_from_packed_text(normalized)
        if packed_issues:
            validation_issues.extend(packed_issues)
        if rebuilt_rows:
            warnings.append("已从挤压在同一列的订单正文行重建明细。")
    if _should_replace_mapped_rows(old_rows, rebuilt_rows):
        normalized["mapped_detail_rows"] = rebuilt_rows
    elif rebuilt_rows and old_rows:
        warnings.append("规范化明细重建结果未优于原结果，已保留原结构化明细。")
    if _normalize_mapped_detail_specs(list(normalized.get("mapped_detail_rows") or [])):
        warnings.append("已同步规范化结构化明细中的 PP 型号/规格空格。")
    filtered_rows, dropped_count = _filter_mapped_detail_rows(list(normalized.get("mapped_detail_rows") or []))
    if dropped_count:
        normalized["mapped_detail_rows"] = filtered_rows
        warnings.append(f"已过滤 {dropped_count} 行非明细区块文本。")
    if not normalized.get("mapped_detail_rows"):
        packed_rows, packed_issues = _rebuild_rows_from_packed_text(normalized)
        if packed_rows:
            normalized["mapped_detail_rows"] = packed_rows
            validation_issues.extend(packed_issues)
            if "已从挤压在同一列的订单正文行重建明细。" not in warnings:
                warnings.append("已从挤压在同一列的订单正文行重建明细。")
    if validation_issues:
        normalized.setdefault("issues", []).extend(validation_issues)

    if source_text:
        normalized["header_info"] = extract_header_info_from_header_region(source_text, normalized.get("header_info") or {})
    else:
        normalized["header_info"] = {
            key: value
            for key, value in (normalized.get("header_info") or {}).items()
            if clean_text(value) and not _looks_polluted_header_value(str(value))
        }
    if not clean_text((normalized.get("header_info") or {}).get(HEADER_KEYS["tax"])):
        inferred_tax = _infer_uniform_tax_from_tables(normalized.get("raw_detail_tables") or [])
        if inferred_tax:
            normalized.setdefault("header_info", {})[HEADER_KEYS["tax"]] = inferred_tax

    if warnings:
        seen: set[str] = set()
        normalized["normalization_warnings"] = []
        for warning in warnings:
            if warning not in seen:
                seen.add(warning)
                normalized["normalization_warnings"].append(warning)
        normalized.setdefault("warnings", []).extend(w for w in normalized["normalization_warnings"] if w not in normalized.get("warnings", []))
    return normalized
