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


def _append_cell_text(existing: str, addition: str) -> str:
    existing = clean_text(existing)
    addition = clean_text(addition)
    if not addition:
        return existing
    if not existing:
        return addition
    if addition in {")", "）"} and existing.endswith("("):
        return f"{existing}{addition}"
    if addition in {")", "）"}:
        return f"{existing}{addition}"
    if existing.endswith("(") or existing.endswith("（"):
        return f"{existing}{addition}"
    return clean_text(f"{existing} {addition}")


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
                pending[column] = _append_cell_text(pending.get(column, ""), row[column])
            index += 1
            continue

        if pending:
            for column, value in pending.items():
                if column >= len(row):
                    row.extend([""] * (column + 1 - len(row)))
                row[column] = _append_cell_text(value, row[column]) if row[column] else value
            pending = {}

        lookahead = index + 1
        while lookahead < len(rows) and _is_fragment_row(rows[lookahead]) and not _is_opening_price_fragment(rows[lookahead]):
            for column in _non_empty_indexes(rows[lookahead]):
                if column >= len(row):
                    row.extend([""] * (column + 1 - len(row)))
                row[column] = _append_cell_text(row[column], rows[lookahead][column])
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
                previous[column] = _append_cell_text(previous[column], current[column])
            warnings.append("已合并跨行明细规格文本。")
        else:
            merged.append(current)
    return merged, warnings


def _normalize_table_rows(rows: list[list[str]]) -> tuple[list[list[str]], dict[int, str], list[str]]:
    rows = [[clean_text(value) for value in row] for row in rows if any(clean_text(value) for value in row)]
    header_index, mapping = find_detail_header_row(rows)
    if header_index is None:
        return rows, {}, []
    rows, fragment_warnings = _merge_fragment_rows(rows, header_index, mapping)
    rows, multiline_warnings = _merge_multiline_detail_rows(rows, header_index, mapping)
    normalized = rows[: header_index + 1]
    for row in rows[header_index + 1 :]:
        if _is_header_like_row(row) or _is_summary_or_section_row(row):
            continue
        if not looks_like_detail_data(row) and len(_non_empty_indexes(row)) <= 2:
            continue
        normalized.append(row)
    return normalized, mapping, fragment_warnings + multiline_warnings


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


def _should_replace_mapped_rows(old_rows: list[dict[str, Any]], new_rows: list[dict[str, Any]]) -> bool:
    if not new_rows:
        return False
    if not old_rows:
        return True
    old_strong = _strong_row_count(old_rows)
    new_strong = _strong_row_count(new_rows)
    if new_strong < old_strong:
        return False
    old_score = sum(_row_quality(row) for row in old_rows)
    new_score = sum(_row_quality(row) for row in new_rows)
    if new_score > old_score:
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
    has_code = bool(re.search(r"[A-Za-z]{1,8}\d{2,}|\d{3,}[A-Za-z]", code))
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

    rebuilt_rows, validation_issues = _rebuild_detail_rows_from_tables(normalized)
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
