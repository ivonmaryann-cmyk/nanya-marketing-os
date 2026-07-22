from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from typing import Any


STANDARD_HEADERS = ["序号", "物料编码", "物料名称", "说明", "数量", "单位", "含税单价", "金额", "交货日期", "备注"]

DETAIL_ALIASES = {
    "序号": ["序号", "项次", "项目", "行号", "item", "item no", "no.", "no"],
    "物料编码": ["物料编码", "物料编号", "物料代码", "原料编码", "料件编号", "料号", "品号", "goods no", "goodsno", "part no", "part no.", "part", "p/n"],
    "物料名称": ["物料名称", "物料品名", "原料名称", "品名", "名称", "型号", "规格", "名称规格", "型号/规格", "description", "desc"],
    "说明": ["说明", "描述", "物料描述", "环保要求", "rohs", "remark", "comments", "comment"],
    "数量": ["数量", "采购量", "订购数量", "quantity", "qty"],
    "单位": ["单位", "计量单位", "unit", "uom"],
    "含税单价": ["含税单价", "单价", "单价rmb", "unit price", "price"],
    "金额": ["金额", "价税合计", "合计金额", "total amount", "amount", "total"],
    "交货日期": ["交货日期", "到货日期", "交期", "delivery date", "del. date", "delivery"],
    "备注": ["备注", "附注", "notes", "remark", "comments"],
}

HEADER_ALIASES = {
    "客户": ["客户", "采购方", "买方", "公司名称", "需方"],
    "供应商": ["供应商", "供方", "卖方", "vendor", "supplier"],
    "订单号": ["订单号", "采购单号", "采购订单号", "p.o. no", "po no", "pono", "p.ono"],
    "合同编号": ["合同编号", "合同号", "contract no", "contract"],
    "日期": ["日期", "订单日期", "采购日期", "p.o date", "date"],
    "币别": ["币别", "币种", "currency"],
    "交货地点": ["交货地点", "交货地址", "送货地址", "delivery to", "shipto", "ship to"],
    "付款方式": ["付款方式", "付款条件", "payment terms", "payment"],
    "税率": ["税率", "含税", "税种", "tax rate", "vat"],
    "联系人": ["联系人", "收货人", "contact"],
    "电话": ["电话", "tel", "telephone", "mobile"],
}

SECTION_KEYWORDS = {
    "备注": ["备注", "说明", "特别要求", "注意事项", "合计", "总计", "大写金额"],
    "条款": ["条款", "质量", "环保", "验收", "违约", "争议", "包装", "责任", "签订"],
    "付款信息": ["付款", "结算", "月结", "账期", "发票", "payment"],
    "收货信息": ["收货", "送货", "交货地点", "交货地址", "地址", "联系人", "电话", "ship"],
    "签核区": ["采购", "审核", "批准", "签核", "制单", "supplier confirmation"],
}

DATE_PATTERN = r"20\d{2}[年./-]\s*\d{1,2}[月./-]\s*\d{1,2}日?"
MONEY_PATTERN = r"-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?"


def clean_text(value: Any) -> str:
    text = str(value or "").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def compact(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def normalize_date(value: Any) -> str:
    text = clean_text(value)
    match = re.search(DATE_PATTERN, text)
    if not match:
        compact_month_day = re.search(r"(20\d{2})[-/](\d{2})(\d{2})(?!\d)", text)
        if compact_month_day:
            return f"{compact_month_day.group(1)}-{compact_month_day.group(2)}-{compact_month_day.group(3)}"
        compact_date = re.search(r"20\d{6}", text)
        if compact_date:
            raw = compact_date.group(0)
            return f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"
        return ""
    raw = match.group(0).replace("年", "-").replace("月", "-").replace("日", "")
    raw = raw.replace("/", "-").replace(".", "-")
    parts = [part.strip() for part in raw.split("-") if part.strip()]
    if len(parts) >= 3:
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}-{int(parts[2]):02d}"
    return raw


def normalize_number(value: Any) -> str:
    text = clean_text(value)
    text = re.sub(r"(?<=\d),\s+(?=\d{3}(?:\D|$))", ",", text)
    text = re.sub(r"(?<=\d)\.\s+(?=\d)", ".", text)
    match = re.search(MONEY_PATTERN, text)
    return match.group(0).replace(",", "") if match else ""


def normalize_unit_price(value: Any) -> tuple[str, str]:
    number = normalize_number(value)
    return number, ""


def decimal_or_none(value: Any) -> Decimal | None:
    number = normalize_number(value)
    if not number:
        return None
    try:
        return Decimal(number)
    except InvalidOperation:
        return None


def split_quantity_unit(value: Any) -> tuple[str, str]:
    text = clean_text(value)
    match = re.search(r"(-?\d+(?:,\d{3})*(?:\.\d+)?)\s*([\u4e00-\u9fffA-Za-z/%]+)?", text)
    if not match:
        return "", ""
    return match.group(1).replace(",", ""), clean_text(match.group(2) or "")


def classify_header_cell(value: Any) -> str:
    text = compact(value)
    if not text:
        return ""
    if any(keyword in text for keyword in ["orderno", "订单号", "采购单号"]):
        return ""
    if any(keyword in text for keyword in ["goodsno", "partno", "p/n", "物料编码", "原料编码", "料号", "品号"]):
        return "物料编码"
    if any(keyword in text for keyword in ["unitprice", "含税单价", "单价rmb"]):
        return "含税单价"
    for standard, aliases in DETAIL_ALIASES.items():
        if any(compact(alias) in text for alias in aliases):
            return standard
    return ""


def _is_explicit_material_name_header(value: Any) -> bool:
    text = compact(value)
    return any(
        keyword in text
        for keyword in ["物料名称", "物料品名", "原料名称", "产品名称", "goodsapellation", "goodsappellation"]
    )


def _is_explicit_spec_header(value: Any) -> bool:
    text = compact(value)
    return any(keyword in text for keyword in ["型号/规格", "型号规格", "规格型号", "型号及规格"])


def header_score(row: list[str]) -> tuple[int, dict[int, str]]:
    mapping: dict[int, str] = {}
    has_separate_material_name = any(_is_explicit_material_name_header(value) for value in row)
    for index, value in enumerate(row):
        standard = classify_header_cell(value)
        if has_separate_material_name and _is_explicit_spec_header(value):
            standard = "说明"
        if standard:
            mapping[index] = standard
    return len(set(mapping.values())), mapping


def find_detail_header_row(rows: list[list[str]]) -> tuple[int | None, dict[int, str]]:
    best_index: int | None = None
    best_mapping: dict[int, str] = {}
    best_score = 0
    for index, row in enumerate(rows[:12]):
        row_text = compact(" ".join(row))
        if any(keyword in row_text for keyword in ["大写金额", "总金额", "合计金额"]):
            continue
        score, mapping = header_score(row)
        has_identity = bool({"物料编码", "物料名称"} & set(mapping.values()))
        has_value = len({"数量", "含税单价", "金额", "交货日期"} & set(mapping.values())) >= 1
        if score > best_score and score >= 3 and has_identity and has_value:
            best_index = index
            best_mapping = mapping
            best_score = score
    return best_index, best_mapping


def map_detail_row(raw_headers: list[str], row: list[str], mapping: dict[int, str]) -> dict[str, Any]:
    original: dict[str, str] = {}
    standard = {header: "" for header in STANDARD_HEADERS}
    cleaning_notes: list[str] = []
    for index, raw_header in enumerate(raw_headers):
        value = clean_text(row[index] if index < len(row) else "")
        original[clean_text(raw_header) or f"列{index + 1}"] = value
        field = mapping.get(index)
        if field and not standard.get(field):
            standard[field] = value

    if standard["数量"] and not standard["单位"]:
        qty, unit = split_quantity_unit(standard["数量"])
        if qty:
            standard["数量"] = qty
        if unit:
            standard["单位"] = unit

    for numeric_field in ["数量", "含税单价", "金额"]:
        if numeric_field == "含税单价":
            number, note = normalize_unit_price(standard.get(numeric_field))
            if note:
                cleaning_notes.append(note)
        else:
            number = normalize_number(standard.get(numeric_field))
        if number:
            standard[numeric_field] = number
    date = normalize_date(standard.get("交货日期"))
    if not date:
        date = normalize_date(" ".join(clean_text(value) for value in row))
    if date:
        standard["交货日期"] = date
    _repair_split_material_code(standard)

    return {"original": original, "standard": standard, "cleaning_notes": cleaning_notes}


def _repair_split_material_code(standard: dict[str, str]) -> None:
    material_code = clean_text(standard.get("物料编码"))
    material_name = clean_text(standard.get("物料名称"))
    if not material_code:
        return

    material_parts = material_code.split()
    if len(material_parts) > 1 and re.search(r"\d", material_parts[0]):
        remainder_parts = material_parts[1:]
        if _looks_like_code_continuation(material_parts[0], remainder_parts):
            material_code = material_parts[0] + "".join(remainder_parts)
            standard["物料编码"] = material_code
        else:
            material_code = material_parts[0]
            standard["物料编码"] = material_code
            remainder = " ".join(remainder_parts)
            if remainder and remainder not in material_name:
                material_name = clean_text(f"{remainder} {material_name}")
                standard["物料名称"] = material_name

    if not re.search(r"[-_/]$", material_code):
        standard["物料编码"] = _compact_material_code(material_code)
        return
    match = re.match(r"^([A-Za-z0-9][A-Za-z0-9_-]{1,16})(?:\s+|$)(.*)$", material_name)
    if not match:
        standard["物料编码"] = _compact_material_code(material_code)
        return
    tail = match.group(1)
    if not _is_code_tail(tail):
        standard["物料编码"] = _compact_material_code(material_code)
        return
    standard["物料编码"] = _compact_material_code(f"{material_code}{tail}")
    standard["物料名称"] = clean_text(match.group(2))


def _compact_material_code(value: str) -> str:
    return re.sub(r"\s+", "", clean_text(value))


def _looks_like_code_continuation(prefix: str, parts: list[str]) -> bool:
    if not parts:
        return False
    if all(re.fullmatch(r"\d{1,8}", part) for part in parts):
        return True
    if re.search(r"[-_/]$", prefix) and all(_is_code_tail(part) for part in parts):
        return True
    return False


def _is_code_tail(value: str) -> bool:
    text = clean_text(value)
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{1,16}", text):
        return False
    if text.upper() in {"FR4", "NYA1", "NY2150", "SHNY", "RMB"}:
        return False
    return bool(re.search(r"\d", text))


def looks_like_detail_data(row: list[str]) -> bool:
    text = " ".join(clean_text(value) for value in row)
    has_number = len(re.findall(MONEY_PATTERN, text)) >= 1
    has_code = bool(re.search(r"\b[A-Za-z0-9][A-Za-z0-9_-]{5,}\b", text))
    has_date = bool(re.search(DATE_PATTERN, text) or re.search(r"20\d{6}", text))
    return has_number and (has_code or has_date)


def extract_key_values(lines: list[str]) -> dict[str, str]:
    result = {key: "" for key in HEADER_ALIASES}
    full_text = "\n".join(lines)
    for field, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            pattern = rf"{re.escape(alias)}\s*(?:\([^)）]*\)|（[^)）]*）)?\s*[:：]?\s*([^\n]+)"
            match = re.search(pattern, full_text, flags=re.I)
            if not match:
                continue
            value = match.group(1).strip()
            value = re.split(r"\s{2,}|(?:供应商|供方|订单号|采购单号|日期|币别|付款方式|税率)[:：]", value)[0].strip()
            if value:
                result[field] = value[:120]
                break

    if result["日期"]:
        result["日期"] = normalize_date(result["日期"])
    if not result.get("日期"):
        result["日期"] = normalize_date(full_text)
    if not result["订单号"]:
        match = re.search(r"\b([A-Z]{1,6}\d{5,}[A-Z0-9_-]*)\b", full_text, flags=re.I)
        if match:
            result["订单号"] = match.group(1)
    if not result["合同编号"]:
        match = re.search(r"\b(TL\d{6,}|[A-Z]{2,}\d{6,})\b", full_text)
        if match and match.group(1) != result.get("订单号"):
            result["合同编号"] = match.group(1)
    if not result["税率"]:
        match = re.search(r"(?:含税|税率|VAT|Tax Rate)[^\d%]{0,10}(\d{1,2}(?:\.\d+)?%?)", full_text, flags=re.I)
        if match:
            value = match.group(1)
            result["税率"] = value if value.endswith("%") else f"{value}%"
    return {key: value for key, value in result.items() if value}


def classify_section_line(line: str) -> str:
    text = compact(line)
    if not text:
        return ""
    for section, keywords in SECTION_KEYWORDS.items():
        if any(compact(keyword) in text for keyword in keywords):
            return section
    return ""
