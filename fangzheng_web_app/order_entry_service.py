from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from copy import copy
from decimal import Decimal, InvalidOperation
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .database import automation_cursor as db_cursor
from .customer_archive_service import get_customer, get_enabled_extraction_maps
from .customer_spec_mapping_service import build_customer_spec_match
from .db import utcnow
from .excel_utils import load_workbook_compat
from .file_storage import resolve_attachment_path
from .order_intake_service import get_case
from .paths import PACKAGE_DIR, PROJECT_DIR
from .pdf_excel_domestic_export import (
    build_domestic_template_data,
    infer_product_type_from_spec,
    normalize_product_type,
)
from .order_document_sources import build_mail_html_purchase_document
from .order_interface_service import invalidate_changed_order_matches, record_order_detail_event
from .order_price_validation_service import PRICE_REVIEW_SNAPSHOT_KEY, review_case_template_prices
from .purchase_field_rules import clean_text, normalize_date, normalize_number
from .purchase_factory_mapper import project_factory_document
from .pdf_excel_service import recognize_purchase_order_document


# The mail workspace and PDF/图片转Excel export intentionally share this
# exact workbook and field order.  The editable page remains the saved source
# of truth for downloads.
DOMESTIC_TEMPLATE_PATH = PACKAGE_DIR / "default_rules" / "order_entry" / "PDF转Excel内销录单模板.xlsx"

HEADER_FIELDS = (
    "order_type", "type_1", "type_2", "bill_to_customer_code",
    "ship_to_customer_code", "delivery_factory", "customer_order_number", "ledger",
    "tax_type", "customer_invoice_number", "commission_rate",
)
LINE_FIELDS = (
    "line_no", "material_status", "product_code", "product_name", "adhesive_code", "customer_product_code",
    "customer_spec", "customer_spec_match", "product_type", "delivery_date",
    "quantity", "price_before_tax", "unit_price", "origin",
    "customer_order_seq", "one_to_many", "remark",
)
PERSISTED_LINE_FIELDS = (*LINE_FIELDS, "old_product_name", "customer_order_number")
HEADER_LABELS = {
    "order_type": "单别", "type_1": "类型1", "type_2": "类型2",
    "bill_to_customer_code": "账款客户编号", "ship_to_customer_code": "送货客户编号",
    "delivery_factory": "送货厂别", "customer_order_number": "客户订单号", "ledger": "账套",
    "tax_type": "税种", "customer_invoice_number": "客户发票号",
    "commission_rate": "佣金比率",
}
LINE_LABELS = {
    "line_no": "项次", "customer_order_number": "客户订单号", "material_status": "料号状态", "product_code": "产品编号", "product_name": "品名", "old_product_name": "旧品名", "adhesive_code": "胶系编码",
    "customer_product_code": "客户产品编号", "customer_spec": "客户规格",
    "customer_spec_match": "客户规格匹配", "product_type": "产品类型（PP、基板）",
    "delivery_date": "出货日期", "quantity": "数量", "price_before_tax": "税前单价",
    "unit_price": "单价", "origin": "产地", "customer_order_seq": "客户订单序号",
    "one_to_many": "一对多", "remark": "备注",
}
REQUIRED_HEADER_FIELDS = {"order_type", "bill_to_customer_code", "ledger"}
REQUIRED_LINE_FIELDS = {"line_no", "customer_product_code", "quantity"}
MATERIAL_STATUS_VALUES = {"查询", "新增"}
DEFAULT_HEADER_VALUES = {
    "order_type": "220",
    "type_1": "1",
    "type_2": "1",
    "bill_to_customer_code": "",
    "ship_to_customer_code": "",
    "delivery_factory": "",
    "customer_order_number": "",
    "ledger": "KL01",
    "tax_type": "",
    "customer_invoice_number": "",
    "commission_rate": "",
}
MAX_ORDER_ATTACHMENT_BYTES = 20 * 1024 * 1024
MAX_ORDER_ATTACHMENT_ROWS = 20_000

# These fields are deliberately not inferred from an order e-mail.  产品编号、品名
# will later come from Nyeos 151; 产地和一对多 are business decisions.  Leaving a
# value blank is safer than presenting an unverified guess as a usable result.
MANUAL_ONLY_LINE_FIELDS = {"product_code", "product_name", "origin", "one_to_many"}

# The source documents use different wording.  Keep this mapping here instead
# of forcing business users to normalise their customers' Excel files first.
_ATTACHMENT_HEADERS = {
    "line_no": {"序号", "项次", "项目", "行号", "item", "no"},
    "customer_order_number": {"PO号", "PO单号", "客户订单号", "采购订单号", "订单号", "po no", "po number"},
    "product_code": {"产品编号", "物料编号", "物料编码", "料号", "品号", "厂内料号"},
    "product_name": {"品名", "物料名称", "名称", "产品名称"},
    "customer_product_code": {"客户产品编号", "客户料号", "客户物料编号", "客户产品码", "part no", "p/n"},
    "customer_spec": {"客户规格", "规格", "型号", "名称规格", "物料规格", "物料描述"},
    "customer_spec_match": {"客户规格匹配", "规格匹配"},
    "product_type": {"产品类型", "产品类型（pp、基板）", "品类"},
    "delivery_date": {"出货日期", "交货日期", "交期", "需求日", "需求交期", "要求交期", "计划交期", "供应商交期", "到货日期", "delivery date"},
    "quantity": {"数量", "采购量", "订购数量", "订单数量", "qty", "quantity"},
    "_quantity_unit": {"单位", "计量单位", "数量单位", "uom"},
    "price_before_tax": {"税前单价", "未税单价", "不含税单价"},
    "unit_price": {"单价", "含税单价", "unit price", "price"},
    "origin": {"产地", "原产地"},
    "customer_order_seq": {"客户订单序号", "订单序号", "客户项次"},
    "one_to_many": {"一对多"},
    "remark": {"备注", "说明", "订单备注", "需方备注", "供方备注"},
}


def _compact_key(value: Any) -> str:
    return re.sub(r"[\s:：_\-（）()]+", "", str(value or "")).lower()


_ATTACHMENT_HEADER_INDEX = {
    _compact_key(alias): field
    for field, aliases in _ATTACHMENT_HEADERS.items()
    for alias in aliases
}


def _json(value: str, fallback: Any) -> Any:
    try:
        return json.loads(value or "")
    except (TypeError, ValueError):
        return fallback


def normalize_customer_order_number(value: Any) -> str:
    """Normalize display prefixes while preserving the customer's real PO."""
    text = clean_text(value)
    if not text:
        return ""
    if text.startswith("暂无PO号-"):
        return ""
    text = re.sub(r"^(?:建价|估价|报价|询价)\s*[:：-]?\s*", "", text, flags=re.I)
    match = re.search(r"(?i)(PO(?:[-_][A-Z0-9]+)+)", text)
    if match:
        return match.group(1).upper().replace("_", "-")
    return text


def _case_for_template(
    case_id: int, employee_id: str, *, action_type: str = "new_order",
) -> dict[str, Any]:
    case = get_case(case_id, employee_id)
    if not case:
        raise ValueError("订单邮件不存在或无权操作")
    if case.get("action_type") != action_type:
        label = "录单" if action_type == "new_order" else "修改订单"
        raise ValueError(f"只有已分流为“{label}”的邮件才能打开此模板")
    return case


def _blank_line(line_no: int) -> dict[str, str]:
    return {
        field: str(line_no)
        if field in {"line_no", "customer_order_seq"}
        else "查询"
        if field == "material_status"
        else ""
        for field in PERSISTED_LINE_FIELDS
    }


def _line_sequence(values: dict[str, Any], fallback: int) -> str:
    """Keep the internal line number aligned with the customer order sequence."""
    sequence = clean_text(values.get("customer_order_seq"))
    if re.fullmatch(r"\d+", sequence or "") and int(sequence) > 0:
        return str(int(sequence))
    return str(fallback)


def _is_pp_spec(value: str) -> bool:
    """Use the same PP classification for quantity conversion and the template."""
    return infer_product_type_from_spec(value) == "PP"


def _meter_values(value: str) -> list[Decimal]:
    """Return unambiguous metre values, without mistaking 0.075MM for metres."""
    found = re.findall(r"(\d+(?:\.\d+)?)\s*(?:米|[mM])(?![A-Za-z])", clean_text(value))
    values: list[Decimal] = []
    for item in found:
        try:
            values.append(Decimal(item))
        except InvalidOperation:
            continue
    return values


def _decimal_text(value: Decimal) -> str:
    text = format(value.normalize(), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def _append_roll_remark(remark: str, roll_quantity: Decimal | None) -> str:
    if roll_quantity is None:
        return remark
    roll_text = f"{_decimal_text(roll_quantity)}卷"
    prefix, separator, suffix = clean_text(remark).partition("&")
    prefix_parts = [item.strip() for item in re.split(r"[；;\n]+", prefix) if item.strip()]
    if roll_text not in prefix_parts:
        prefix_parts.append(roll_text)
    return f"{'；'.join(prefix_parts)}&{suffix}" if separator else "；".join(prefix_parts)


def _apply_auto_extraction_policy(
    values: dict[str, str], *, quantity_unit: Any = "", product_context: Any = "",
) -> dict[str, str]:
    """Apply the conservative domestic order-entry extraction rules.

    A value is kept only when its meaning is explicit in the source.  In
    particular, PP conversion requires an explicit PP spec, a roll unit, one
    unambiguous metre value and one numeric order quantity.  All other cases
    remain for business completion.
    """
    result = dict(values)
    for field in MANUAL_ONLY_LINE_FIELDS:
        result[field] = ""

    spec = result.get("customer_spec", "")
    # A PO often puts “半固化片” in the material-name column and the metre
    # value in the description column.  Both are explicit source evidence, but
    # only the description is shown as the customer specification.
    if not _is_pp_spec(f"{spec} {clean_text(product_context)}"):
        return result

    metres = _meter_values(spec)
    distinct_metres = {value.normalize() for value in metres}
    meter = next(iter(distinct_metres), None) if len(distinct_metres) == 1 else None
    if meter is None:
        # PP small pieces and PP specifications without a per-roll length keep
        # the customer-provided quantity; they cannot be converted to metres.
        return result

    try:
        quantity = Decimal(str(result.get("quantity") or "").replace(",", ""))
    except InvalidOperation:
        return result
    if quantity < 0:
        return result
    unit = clean_text(quantity_unit)
    if "张" in unit:
        # PP 小片：客户明确以张计数，数量直接保留，不做米数换算。
        return result
    if Decimal("0") < quantity < Decimal("1") and meter is not None:
        # A fractional PP quantity still represents a roll fraction.  The
        # template is in metres, so convert it from rolls even when the source
        # unit column is absent.
        result["quantity"] = _decimal_text(quantity * meter)
        result["remark"] = _append_roll_remark(result.get("remark", ""), quantity)
        return result
    if "卷" not in unit:
        # A quantity of one or more is converted only when the source makes
        # the roll unit explicit; otherwise retain the source quantity.
        return result
    result["quantity"] = _decimal_text(quantity * meter)
    result["remark"] = _append_roll_remark(result.get("remark", ""), quantity)
    return result


def _normalize_body_delivery_date(value: Any, reference_date: Any = "") -> str:
    normalized = normalize_date(value)
    if normalized:
        return normalized
    match = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日?", clean_text(value))
    if not match:
        return ""
    year_match = re.search(r"(20\d{2})", clean_text(reference_date))
    year = int(year_match.group(1)) if year_match else datetime.now().year
    try:
        return datetime(year, int(match.group(1)), int(match.group(2))).date().isoformat()
    except ValueError:
        return ""


def _split_demand_delivery_rows(body_text: str, reference_date: Any = "") -> list[dict[str, Any]]:
    """Parse ERP mail rows headed by 下单日期/需求交期 when no HTML table survives."""
    text = " ".join(str(body_text or "").split())
    if "下单日期" not in text or not any(label in text for label in ("需求交期", "要求交期")):
        return []
    row_pattern = re.compile(
        r"(?P<order_date>\d{1,2}/\d{1,2})\s+"
        r"(?P<product_type>板材|PP)\s+\S+\s+"
        r"(?P<customer_product_code>[A-Za-z0-9]{8,})\s+"
        r"(?P<body>.*?)(?=\s+\d{1,2}/\d{1,2}\s+(?:板材|PP)\s+\S+\s+[A-Za-z0-9]{8,}|$)",
        re.IGNORECASE,
    )
    result: list[dict[str, Any]] = []
    for match in row_pattern.finditer(text):
        body = clean_text(match.group("body"))
        detail_match = re.match(
            r"(?P<customer_spec>.*?)\s+"
            r"(?P<quantity>\d+(?:\.\d+)?)\s+"
            r"(?P<unit>[A-Za-z]+|[\u4e00-\u9fff]+)\s+"
            r"(?P<delivery_date>(?:20\d{2}[年./-]\s*)?\d{1,2}月\d{1,2}日?|20\d{2}[-/.]\d{1,2}[-/.]\d{1,2})(?P<remark>.*)$",
            body,
        )
        if not detail_match:
            continue
        values = {
            **_blank_line(len(result) + 1),
            "customer_product_code": match.group("customer_product_code"),
            "customer_spec": detail_match.group("customer_spec"),
            "product_type": match.group("product_type"),
            "quantity": detail_match.group("quantity"),
            "delivery_date": _normalize_body_delivery_date(
                detail_match.group("delivery_date"), reference_date,
            ),
            "remark": clean_text(detail_match.group("remark")),
        }
        result.append(_line_entry(
            values,
            label="邮件正文需求交期表",
            reference=f"下单日期 {match.group('order_date')}",
            line_no=len(result) + 1,
            quantity_unit=detail_match.group("unit"),
        ))
    return result


def _split_po_purchase_table_rows(body_text: str, reference_date: Any = "") -> list[dict[str, Any]]:
    """Parse line-broken mail tables containing PR/PO/material columns."""
    lines = [" ".join(item.split()) for item in str(body_text or "").splitlines() if item.strip()]
    required_headers = {"序号", "PR号", "PO号", "物料编码", "物料描述", "需求数量", "需求交期"}
    if not required_headers.issubset(set(lines)):
        return []
    row_starts = [
        index for index in range(len(lines) - 4)
        if re.fullmatch(r"[1-9]\d*", lines[index])
        and re.fullmatch(r"PR[-_A-Z0-9]+", lines[index + 1], re.IGNORECASE)
        and "PO" in lines[index + 2].upper()
        and re.fullmatch(r"[A-Z0-9_-]{8,}", lines[index + 3], re.IGNORECASE)
    ]
    result: list[dict[str, Any]] = []
    for position, start in enumerate(row_starts):
        end = row_starts[position + 1] if position + 1 < len(row_starts) else len(lines)
        chunk = lines[start:end]
        date_index = next((
            index for index, value in enumerate(chunk[4:], start=4)
            if normalize_date(value) or re.fullmatch(r"\d{1,2}月\d{1,2}日?", value)
        ), -1)
        if date_index <= 4 or not re.fullmatch(r"-?\d+(?:\.\d+)?", chunk[date_index - 1].replace(",", "")):
            continue
        description = clean_text(" ".join(chunk[4:date_index - 1]))
        values = {
            **_blank_line(len(result) + 1),
            "customer_order_seq": chunk[0],
            "customer_order_number": normalize_customer_order_number(chunk[2]),
            "customer_product_code": chunk[3],
            "customer_spec": description,
            "quantity": chunk[date_index - 1].replace(",", ""),
            "delivery_date": _normalize_body_delivery_date(chunk[date_index], reference_date),
            "remark": clean_text(" ".join(chunk[date_index + 1:])),
        }
        result.append(_line_entry(
            values, label="邮件正文采购表", reference=f"序号 {chunk[0]}",
            line_no=len(result) + 1, product_context=description,
        ))
    return result


def _split_body_order_rows(body_text: str, reference_date: Any = "") -> list[dict[str, Any]]:
    """Extract simple ERP-style rows from the line-oriented mail body.

    This is a safe first path for HTML mail tables. Attachment-specific parsers
    will feed the same structure in the next extraction layer.
    """
    po_purchase_rows = _split_po_purchase_table_rows(body_text, reference_date)
    if po_purchase_rows:
        return po_purchase_rows
    demand_delivery_rows = _split_demand_delivery_rows(body_text, reference_date)
    if demand_delivery_rows:
        return demand_delivery_rows
    lines = [" ".join(item.split()) for item in str(body_text or "").splitlines() if item.strip()]
    positions = [index for index, value in enumerate(lines) if re.fullmatch(r"(?:HJ\d{8,}|[A-Z]{1,4}\d{6,}[A-Z0-9_-]*)", value, re.I)]
    result: list[dict[str, Any]] = []
    for line_no, start in enumerate(positions, start=1):
        end = positions[line_no] if line_no < len(positions) else len(lines)
        chunk = lines[start:end]
        if len(chunk) < 2:
            continue
        customer_code = next((item for item in chunk[1:] if re.fullmatch(r"[A-Z0-9]{8,}", item, re.I)), "")
        quantity = next((item for item in reversed(chunk) if re.fullmatch(r"\d+(?:\.\d+)?", item.replace(",", ""))), "")
        material_index = next((index for index, item in enumerate(chunk) if re.fullmatch(r"NY\d+[A-Z0-9]*", item, re.I)), -1)
        spec = ""
        if material_index >= 0:
            spec = " ".join(chunk[material_index:min(len(chunk), material_index + 9)])
        if not customer_code and not spec:
            continue
        values = {
            **_blank_line(len(result) + 1),
            "customer_product_code": customer_code,
            "customer_spec": spec,
            "quantity": quantity.replace(",", ""),
            "customer_order_number": chunk[0],
        }
        result.append(_line_entry(values, label="邮件正文", reference=f"第 {start + 1} 行附近", line_no=len(result) + 1))
    return result


def _source(label: str, reference: str, values: dict[str, str]) -> dict[str, dict[str, str]]:
    return {
        field: {"label": label, "reference": reference}
        for field, value in values.items()
        if value and field not in {"line_no", "material_status"}
    }


def _line_entry(
    values: dict[str, Any], *, label: str, reference: str, line_no: int = 1, quantity_unit: Any = "", product_context: Any = "",
) -> dict[str, Any]:
    line = _blank_line(line_no)
    for field in PERSISTED_LINE_FIELDS:
        value = values.get(field)
        if value not in (None, ""):
            line[field] = clean_text(value)
    line["line_no"] = str(line_no)
    if line["product_type"]:
        line["product_type"] = normalize_product_type(line["product_type"])
    if not line["product_type"]:
        line["product_type"] = infer_product_type_from_spec(
            f"{line['customer_spec']} {clean_text(product_context)}"
        )
    if line["delivery_date"]:
        line["delivery_date"] = normalize_date(line["delivery_date"]) or line["delivery_date"]
    for field in {"quantity", "price_before_tax", "unit_price"}:
        if line[field]:
            line[field] = normalize_number(line[field]) or line[field]
    line = _apply_auto_extraction_policy(line, quantity_unit=quantity_unit, product_context=product_context)
    return {"values": line, "sources": _source(label, reference, line)}


def _value_by_alias(mapping: dict[str, Any], *aliases: str) -> Any:
    """Return an explicit source value even when a document uses bilingual headings."""
    if not mapping:
        return ""
    wanted = {_compact_key(alias) for alias in aliases}
    for key, value in mapping.items():
        if _compact_key(key) in wanted:
            return value
    for key, value in mapping.items():
        compact = _compact_key(key)
        if any(alias and alias in compact for alias in wanted):
            return value
    return ""


def _tax_inclusive_unit_price(mapping: dict[str, Any]) -> Any:
    excluded_markers = {
        "不含税",
        "未税",
        "税前",
        "nottaxinclusive",
        "taxexclusive",
        "excludingtax",
        "excltax",
        "withouttax",
    }
    eligible = {
        key: value
        for key, value in mapping.items()
        if not any(marker in _compact_key(key) for marker in excluded_markers)
    }
    return _value_by_alias(eligible, "含税单价", "单价", "Unit Price")


def _mapping_source_value(source: dict[str, Any], source_label: str, transform_type: str) -> str:
    """Read explicitly named attachment columns for one customer mapping.

    The mapping UI intentionally uses the customer's visible column names
    (for example ``Material Code``), not technical column indices.  Concatenation
    accepts ``+`` or Chinese/English commas.  Missing parts make the whole
    result blank: the system must never present a partially guessed value.
    """
    labels = [item.strip() for item in re.split(r"[+，,]", str(source_label or "")) if item.strip()]
    if not labels:
        return ""
    delivery_aliases = (
        "交货日期", "要求交货日期", "要求交货期", "要求交期", "需求交期", "需求日", "交期",
    )
    values = [
        clean_text(
            _value_by_alias(source, *delivery_aliases)
            if _compact_key(label) in {_compact_key(alias) for alias in delivery_aliases}
            else _value_by_alias(source, label)
        )
        for label in labels
    ]
    if not all(values):
        return ""
    if transform_type == "concat":
        return " ".join(values)
    return values[0] if len(values) == 1 else ""


def _apply_customer_extraction_mappings(
    values: dict[str, Any], source: dict[str, Any], mappings: list[dict[str, Any]],
    *, source_kind: str = "attachment_table",
) -> dict[str, Any]:
    """Apply customer mappings registered for the current source adapter.

    Customer-maintained order-table mappings work for both attachment tables
    and HTML tables in the mail body.  Manual mappings explicitly blank the
    target.  For all other source types we leave the universal result untouched.
    """
    if not mappings:
        return values
    result = dict(values)
    for mapping in mappings:
        mapping_source = str(mapping.get("source_kind") or "")
        if mapping_source != source_kind and not (
            mapping_source == "attachment_table" and source_kind == "mail_html_table"
        ):
            continue
        target = str(mapping.get("target_field") or "")
        transform = str(mapping.get("transform_type") or "")
        if target not in PERSISTED_LINE_FIELDS:
            continue
        if transform == "manual":
            result[target] = ""
            continue
        value = _mapping_source_value(source, str(mapping.get("source_label") or ""), transform)
        # A configured customer mapping takes precedence only when it produces
        # an explicit source value.  Otherwise retain an already verified
        # universal value; no fallback guess is introduced.
        if value:
            if target == "product_type":
                value = normalize_product_type(value)
                if not value:
                    continue
            if target == "delivery_date":
                value = normalize_date(value)
                if not value:
                    continue
            result[target] = value
    return result


def _canonical_order_number(header: dict[str, Any]) -> str:
    """Keep the actual PO token and discard surrounding document decorations."""
    for key in (
        "客户订单号",
        "订单号",
        "订单编号",
        "采购订单号",
        "合同编号",
        "PO号",
        "PO No",
    ):
        raw = clean_text(_value_by_alias(header, key))
        if not raw:
            continue
        candidates = re.findall(r"[A-Za-z]{1,10}[A-Za-z0-9_-]*\d{4,}[A-Za-z0-9_-]*", raw)
        if candidates:
            return max(candidates, key=len).upper()
        if re.fullmatch(r"[A-Za-z0-9_-]{6,}", raw):
            return raw
    return ""


def _within_order_attachment_size(path: Path) -> bool:
    """Reject real oversized files while keeping parser adapters mockable."""
    try:
        return path.stat().st_size <= MAX_ORDER_ATTACHMENT_BYTES
    except OSError:
        # The recognizer can operate on a virtual/test adapter path; its own
        # error handling remains responsible for unavailable source files.
        return True


def _rows_from_excel(path: Path, filename: str, customer_mappings: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    """Read customer Excel attachments as structured data without OCR."""
    if not _within_order_attachment_size(path):
        return []
    try:
        book = load_workbook_compat(path, data_only=True)
    except Exception:
        return []
    try:
        result: list[dict[str, Any]] = []
        scanned_rows = 0
        for sheet in book.worksheets:
            rows: list[tuple[Any, ...]] = []
            for row in sheet.iter_rows(values_only=True):
                scanned_rows += 1
                if scanned_rows > MAX_ORDER_ATTACHMENT_ROWS:
                    return []
                rows.append(row)
            header_index = -1
            mapping: dict[int, str] = {}
            for index, row in enumerate(rows[:40]):
                candidate = {
                    column: _ATTACHMENT_HEADER_INDEX.get(_compact_key(value))
                    for column, value in enumerate(row) if _ATTACHMENT_HEADER_INDEX.get(_compact_key(value))
                }
                # A single heading such as "规格" is too ambiguous; require a
                # usable set of columns before treating it as an order table.
                if len(set(candidate.values())) >= 2 and ("quantity" in candidate.values() or "customer_product_code" in candidate.values()):
                    header_index, mapping = index, candidate
                    break
            if header_index < 0:
                continue
            for index, row in enumerate(rows[header_index + 1 :], start=header_index + 2):
                values = {field: row[column] if column < len(row) else "" for column, field in mapping.items()}
                if not any(clean_text(value) for value in values.values()):
                    continue
                text = " ".join(clean_text(value) for value in values.values())
                if any(token in text for token in ("合计", "总计", "小计")):
                    continue
                source_row = {
                    clean_text(rows[header_index][column]): row[column] if column < len(row) else ""
                    for column in range(len(rows[header_index])) if clean_text(rows[header_index][column])
                }
                values = _apply_customer_extraction_mappings(values, source_row, customer_mappings or [])
                result.append(_line_entry(
                    values,
                    label=f"附件：{filename}",
                    reference=f"{sheet.title} 第 {index} 行",
                    line_no=len(result) + 1,
                    quantity_unit=values.get("_quantity_unit", ""),
                ))
        return result
    finally:
        book.close()


def _line_from_pipeline_row(
    row: dict[str, Any], order_number: str, label: str, reference: str, line_no: int,
    customer_mappings: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    original = row.get("original") or {}
    standard = row.get("standard") or {}
    raw_spec = _value_by_alias(original, "物料描述", "Material Description", "名称规格", "客户规格", "物料规格", "规格", "型号")
    raw_before_tax_price = _value_by_alias(original, "不含税单价", "未税单价", "税前单价", "Not tax inclusive Unit Price")
    raw_unit_price = _tax_inclusive_unit_price(original)
    raw_quantity_unit = _value_by_alias(original, "单位", "计量单位", "Unit", "UOM")
    raw_material_name = _value_by_alias(original, "物料品名", "物料名称", "Material Name") or standard.get("物料名称") or ""
    values = {
        "line_no": standard.get("序号") or _value_by_alias(original, "序号", "No") or line_no,
        # 采购订单中的“物料编码 / Material Code”是客户提供的明确料号，
        # 可安全写入内销模板的“客户产品编号”；不从品名或规格推测编码。
        "customer_product_code": (
            _value_by_alias(
                original,
                "客户产品编号",
                "客户料号",
                "客户物料编号",
                "客户产品码",
                "物料编码",
                "Material Code",
            )
            or standard.get("物料编码")
            or ""
        ),
        "customer_spec": raw_spec or standard.get("说明") or standard.get("物料名称") or "",
        "delivery_date": standard.get("交货日期") or _value_by_alias(
            original, "交货日期", "出货日期", "交期", "需求日", "需求交期", "要求交期", "要求交货期", "要求交货日期", "计划交期", "供应商交期", "Delivery Date",
        ) or "",
        "quantity": standard.get("数量") or _value_by_alias(original, "数量", "Quantity", "Qty") or "",
        "price_before_tax": raw_before_tax_price or standard.get("不含税单价") or "",
        "unit_price": raw_unit_price or "",
        "customer_order_number": normalize_customer_order_number(
            _value_by_alias(original, "PO号", "PO单号", "客户订单号", "采购订单号", "订单号", "PO No", "PO Number")
            or order_number
        ),
        "remark": standard.get("备注") or _value_by_alias(original, "备注", "说明", "订单备注") or "",
    }
    values = _apply_customer_extraction_mappings(values, original, customer_mappings or [])
    entry = _line_entry(
        values,
        label=label,
        reference=reference,
        line_no=line_no,
        quantity_unit=raw_quantity_unit or standard.get("单位") or "",
        product_context=raw_material_name,
    )
    # Similar PP rows may intentionally retain a blank quantity. Their source
    # row identity still makes them distinct order lines and prevents collapse.
    entry["_source_identity"] = "|".join(clean_text(value) for value in (
        label, reference,
        _value_by_alias(original, "PO项目号", "Project No", "项目号"),
        _value_by_alias(original, "物料编码", "Material Code"), raw_spec,
    ))
    return entry


def _rows_from_pdf_or_image(path: Path, filename: str, customer_mappings: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    if not _within_order_attachment_size(path):
        return []
    try:
        document = recognize_purchase_order_document(
            {"stored_path": str(path), "original_filename": filename}
        )
        return _rows_from_shared_purchase_document(
            document,
            label=f"附件：{filename}",
            reference_prefix="识别明细",
            customer_mappings=customer_mappings,
        )
    except Exception as exc:
        # A PDF/image order must never silently fall back to the lower-fidelity
        # mail-body parser.  That used to look like a successful refresh while
        # producing different rows from PDF/图片转Excel.
        raise ValueError(f"附件《{filename}》未能按 PDF/图片转Excel 规则提取：{exc}") from exc


def _rows_from_shared_purchase_document(
    document: dict[str, Any], *, label: str, reference_prefix: str,
    customer_mappings: list[dict[str, Any]] | None = None,
    source_kind: str = "attachment_table",
) -> list[dict[str, Any]]:
    """Project one canonical purchase document into editable domestic rows.

    Every source adapter (PDF/image, HTML mail tables, and later Excel) must
    pass this boundary.  It deliberately delegates to the same factory
    projection and domestic-template exporter used by PDF/图片转Excel.
    """
    project_factory_document(document)
    domestic_data = build_domestic_template_data(document)
    result: list[dict[str, Any]] = []
    source_rows = document.get("mapped_detail_rows") or []
    for index, (line, source_row) in enumerate(
        zip(domestic_data["lines"], source_rows), start=1
    ):
        values = {field: clean_text(line.get(field)) for field in PERSISTED_LINE_FIELDS}
        values["line_no"] = str(index)
        original = source_row.get("original") or {}
        standard = source_row.get("standard") or {}
        values["customer_order_number"] = normalize_customer_order_number(
            _value_by_alias(
                original, "PO号", "PO单号", "客户订单号", "采购订单号", "订单号", "PO No", "PO Number",
            ) or values.get("customer_order_number")
        )
        values = _apply_customer_extraction_mappings(
            values,
            original,
            customer_mappings or [],
            source_kind=source_kind,
        )
        values = _apply_auto_extraction_policy(
            values,
            quantity_unit=(
                _value_by_alias(original, "单位", "计量单位", "Unit", "UOM")
                or standard.get("单位")
                or ""
            ),
            product_context=_value_by_alias(
                original, "物料品名", "物料名称", "材料名称", "Material Name",
            ) or standard.get("物料名称") or "",
        )
        reference = f"{reference_prefix}第 {index} 行"
        result.append(
            {
                "values": values,
                "sources": _source(label, reference, values),
                "_source_identity": "|".join(
                    clean_text(value)
                    for value in (
                        label,
                        reference,
                        values.get("customer_product_code"),
                        values.get("customer_spec"),
                    )
                ),
                "extracted_header": domestic_data["header"],
            }
        )
    return result


def _rows_from_mail_html(case: dict[str, Any]) -> list[dict[str, Any]]:
    """Use HTML mail order tables as another input to the common PO model.

    A non-table mail body intentionally returns no rows so the established
    line-oriented body fallback remains available.  Once a table is recognised
    as an order table, failures are surfaced instead of quietly substituting a
    lower-fidelity parser.
    """
    document = build_mail_html_purchase_document(
        str(case.get("body_html") or ""),
        str(case.get("body_text") or ""),
        source_name=f"邮件正文#{case.get('id') or ''}.html",
        reference_date=str(case.get("received_at") or case.get("sent_at") or ""),
    )
    if not document:
        return []
    try:
        return _rows_from_shared_purchase_document(
            document,
            label="邮件正文表格",
            reference_prefix="表格明细",
            customer_mappings=get_enabled_extraction_maps(case.get("customer_id")),
            source_kind="mail_html_table",
        )
    except Exception as exc:
        raise ValueError(f"邮件正文表格未能按统一订单规则提取：{exc}") from exc


def _rows_from_word(path: Path, filename: str, customer_mappings: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    """Convert Word attachments locally, then use the same PDF/OCR pipeline."""
    if not _within_order_attachment_size(path):
        return []
    soffice = shutil.which("soffice")
    if not soffice:
        return []
    with tempfile.TemporaryDirectory(prefix="order-entry-word-") as output_dir:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", output_dir, str(path)],
                check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=45,
            )
        except (OSError, subprocess.SubprocessError):
            return []
        pdf = Path(output_dir) / f"{path.stem}.pdf"
        return _rows_from_pdf_or_image(pdf, filename, customer_mappings=customer_mappings) if pdf.is_file() else []


def _attachment_rows(case: dict[str, Any]) -> list[dict[str, Any]]:
    customer_mappings = get_enabled_extraction_maps(case.get("customer_id"))
    attachment_groups: dict[str, list[dict[str, Any]]] = {
        "excel": [], "pdf_or_image": [], "word": [],
    }
    for attachment in case.get("attachments") or []:
        if attachment.get("is_inline"):
            continue
        filename = str(attachment.get("filename") or "")
        suffix = Path(filename or str(attachment.get("stored_path") or "")).suffix.lower()
        if suffix in {".xlsx", ".xlsm", ".xls"}:
            attachment_groups["excel"].append(attachment)
        elif suffix in {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}:
            attachment_groups["pdf_or_image"].append(attachment)
        elif suffix in {".doc", ".docx"}:
            attachment_groups["word"].append(attachment)

    # Structured Excel data is both faster and more reliable than PDF/OCR.
    # Only fall back to a lower-priority attachment type when the preferred
    # tier contains no usable order rows.
    for group_name in ("excel", "pdf_or_image", "word"):
        rows: list[dict[str, Any]] = []
        for attachment in attachment_groups[group_name]:
            try:
                path = resolve_attachment_path(str(attachment.get("stored_path") or ""))
            except FileNotFoundError:
                continue
            if not path.is_file():
                continue
            filename = str(attachment.get("filename") or path.name)
            if group_name == "excel":
                rows.extend(_rows_from_excel(path, filename, customer_mappings=customer_mappings))
            elif group_name == "pdf_or_image":
                rows.extend(_rows_from_pdf_or_image(path, filename, customer_mappings=customer_mappings))
            else:
                rows.extend(_rows_from_word(path, filename, customer_mappings=customer_mappings))
        if rows:
            return rows
    return []


def _line_signature(entry: dict[str, Any]) -> tuple[str, ...]:
    source_identity = _compact_key(entry.get("_source_identity"))
    if source_identity:
        return ("source", source_identity)
    values = entry["values"]
    return tuple(
        _compact_key(values.get(field))
        for field in ("customer_order_number", "customer_order_seq", "customer_product_code", "customer_spec", "quantity")
    )


def _format_template_remark(value: Any) -> str:
    """Keep derived roll/metre notes before ``&`` and source notes after it."""
    current = clean_text(value)
    if not current:
        return current
    parts = [part.strip() for part in re.split(r"[&；;\n]+", current) if part.strip()]
    generated = [part for part in parts if re.fullmatch(r"\d+(?:\.\d+)?卷", part)]
    source_notes = [part for part in parts if part not in generated]
    prefix = "；".join(generated)
    suffix = "；".join(source_notes)
    return f"{prefix}&{suffix}" if prefix or suffix else ""


def _merge_initial_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, ...]] = set()
    for entry in rows:
        signature = _line_signature(entry)
        # Do not collapse rows that have no identifying data: a blank source
        # row is still useful to the business user for manual completion.
        if any(signature) and signature in seen:
            continue
        if any(signature):
            seen.add(signature)
        sequence = _line_sequence(entry["values"], len(result) + 1)
        entry["values"]["line_no"] = sequence
        entry["values"]["customer_order_seq"] = sequence
        entry["values"]["remark"] = _format_template_remark(entry["values"].get("remark"))
        result.append(entry)
    return result


def _apply_customer_spec_matches(
    header: dict[str, Any], lines: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    customer_code = clean_text(header.get("bill_to_customer_code"))
    result: list[dict[str, Any]] = []
    for entry in lines:
        values = dict(entry.get("values") or {})
        values["material_status"] = clean_text(values.get("material_status")) or "查询"
        sources = dict(entry.get("sources") or {})
        match_context = json.dumps([
            customer_code,
            clean_text(values.get("product_type")),
            clean_text(values.get("customer_spec")),
        ], ensure_ascii=False, separators=(",", ":"))
        existing_source = sources.get("customer_spec_match") or {}
        manual_match = (
            customer_code
            and existing_source.get("label") == "人工修改"
            and existing_source.get("context") == match_context
        )
        if manual_match:
            result.append({**entry, "values": values, "sources": sources})
            continue
        matched = build_customer_spec_match(
            customer_code,
            values.get("product_type"),
            values.get("customer_spec"),
        ) if customer_code else ""
        values["customer_spec_match"] = matched
        if matched:
            sources["customer_spec_match"] = {
                "label": "客户规格对照表",
                "reference": f"客户编号 {customer_code} / {values.get('product_type') or '未填写产品类型'}",
                "context": match_context,
            }
        else:
            sources.pop("customer_spec_match", None)
        result.append({**entry, "values": values, "sources": sources})
    return result


def _apply_matched_customer_code(
    header: dict[str, str], case: dict[str, Any], *, overwrite: bool,
) -> dict[str, str]:
    customer_id = case.get("customer_id")
    if not customer_id:
        return header
    customer = get_customer(int(customer_id))
    customer_code = clean_text((customer or {}).get("customer_code"))
    if not customer_code:
        return header
    bill_to_code = clean_text(header.get("bill_to_customer_code"))
    if overwrite or not bill_to_code:
        bill_to_code = customer_code
        header["bill_to_customer_code"] = bill_to_code
    if overwrite or not clean_text(header.get("ship_to_customer_code")):
        header["ship_to_customer_code"] = bill_to_code
    return header


def _apply_customer_transit_days(
    case: dict[str, Any], lines: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Turn the customer's requested date into the internal ship date once."""
    customer_id = case.get("customer_id")
    if not customer_id:
        return lines
    customer = get_customer(int(customer_id))
    raw_days = clean_text((customer or {}).get("transit_days"))
    if not re.fullmatch(r"\d+", raw_days):
        return lines
    transit_days = int(raw_days)
    if transit_days <= 0:
        return lines

    result: list[dict[str, Any]] = []
    for entry in lines:
        values = dict(entry.get("values") or {})
        requested_date = normalize_date(values.get("delivery_date"))
        if not requested_date:
            result.append(entry)
            continue
        try:
            ship_date = (date.fromisoformat(requested_date) - timedelta(days=transit_days)).isoformat()
        except ValueError:
            result.append(entry)
            continue
        sources = dict(entry.get("sources") or {})
        values["delivery_date"] = ship_date
        sources["delivery_date"] = {
            "label": "客户运输天数倒推",
            "reference": f"客户需求日 {requested_date} - 运输天数 {transit_days} 天",
        }
        result.append({**entry, "values": values, "sources": sources})
    return result


def _initial_template_data(case: dict[str, Any]) -> tuple[dict[str, str], list[dict[str, Any]]]:
    attachment_rows = _attachment_rows(case)
    # An order attachment remains authoritative.  When there is no supported
    # attachment, an HTML mail table gets the *same* canonical mapping and
    # validation path; only then use the legacy line-oriented body fallback.
    rows = attachment_rows or _rows_from_mail_html(case) or _split_body_order_rows(
        str(case.get("body_text") or ""), case.get("received_at") or case.get("sent_at") or "",
    )
    rows = _merge_initial_rows(rows)
    header = dict(DEFAULT_HEADER_VALUES)
    extracted_header = next(
        (item.get("extracted_header") for item in rows if item.get("extracted_header")),
        {},
    )
    for field in HEADER_FIELDS:
        value = clean_text((extracted_header or {}).get(field))
        if value:
            header[field] = value
    _apply_matched_customer_code(header, case, overwrite=True)
    if not header["customer_order_number"]:
        header["customer_order_number"] = clean_text(
            (case.get("detected_fields") or {}).get("order_number")
        )
    if not header["customer_order_number"]:
        header["customer_order_number"] = f"暂无PO号-{uuid.uuid4()}"
    if rows:
        matched_lines = _apply_customer_spec_matches(header, rows)
        return header, _apply_customer_transit_days(case, matched_lines)
    fields = case.get("detected_fields") or {}
    specs = fields.get("specs") or []
    lines = [
        {
            "values": {
                **_blank_line(index + 1),
                "customer_spec": str(spec),
            },
            "sources": {"customer_spec": {"label": "邮件正文", "reference": "自动识别"}},
        }
        for index, spec in enumerate(specs)
    ] or [{"values": _blank_line(1), "sources": {}}]
    matched_lines = _apply_customer_spec_matches(header, lines)
    return header, _apply_customer_transit_days(case, matched_lines)


def _initial_lines(case: dict[str, Any]) -> list[dict[str, Any]]:
    return _initial_template_data(case)[1]


def _initial_order_change_template_data(case: dict[str, Any]) -> tuple[dict[str, str], list[dict[str, Any]]]:
    """Extract only the fields needed by the order-change workspace.

    The source-document parser is shared, but pricing, material resolution and
    all domestic-order fields are deliberately discarded before persistence.
    """
    header, extracted_lines = _initial_template_data(case)
    order_number = clean_text(header.get("customer_order_number"))
    lines: list[dict[str, Any]] = []
    for index, entry in enumerate(extracted_lines, start=1):
        source_values = entry.get("values") or {}
        values = _blank_line(index)
        values.update({
            "customer_order_number": clean_text(source_values.get("customer_order_number")) or order_number,
            "line_no": clean_text(source_values.get("line_no")) or str(index),
            "customer_product_code": clean_text(source_values.get("customer_product_code")),
            "customer_spec": clean_text(source_values.get("customer_spec")),
            "delivery_date": clean_text(source_values.get("delivery_date")),
            "quantity": clean_text(source_values.get("quantity")),
        })
        values["material_status"] = "查询"
        lines.append({"values": values, "sources": dict(entry.get("sources") or {})})
    return header, lines or [{"values": _blank_line(1), "sources": {}}]


def _initial_order_groups(
    header: dict[str, Any], lines: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Group extracted lines by their normalized PO in first-seen order."""
    header_order_number = normalize_customer_order_number(header.get("customer_order_number"))
    detected = [
        normalize_customer_order_number((entry.get("values") or {}).get("customer_order_number"))
        for entry in lines
    ]
    known_numbers = list(dict.fromkeys(value for value in detected if value))
    buckets: dict[str, list[dict[str, Any]]] = {}
    for entry, detected_number in zip(lines, detected):
        order_number = detected_number
        # A row-level PO is authoritative.  Empty rows form their own group
        # whenever the source also contains any explicit PO, so they are not
        # silently submitted under that PO.  A header PO still covers rows
        # when the source did not provide PO values per line at all.
        if not order_number and not known_numbers:
            order_number = header_order_number
        values = entry.get("values") or {}
        values["customer_order_number"] = order_number
        buckets.setdefault(order_number, []).append(entry)
    if not buckets:
        buckets[header_order_number] = [{"values": _blank_line(1), "sources": {}}]

    groups: list[dict[str, Any]] = []
    for sort_order, (order_number, group_lines) in enumerate(buckets.items()):
        group_header = {**DEFAULT_HEADER_VALUES, **header}
        group_header["customer_order_number"] = order_number
        groups.append({
            "group_key": uuid.uuid4().hex,
            "order_number": order_number,
            "header": group_header,
            "sort_order": sort_order,
            "status": "pending",
            "lines": group_lines,
        })
    return groups


def _ensure_template_groups(conn: Any, template: Any) -> list[Any]:
    """Lazily upgrade a legacy single-header template to one order group."""
    template_id = int(template["id"])
    groups = conn.execute(
        "SELECT * FROM order_entry_template_groups WHERE template_id=? ORDER BY sort_order,id",
        (template_id,),
    ).fetchall()
    if groups:
        return list(groups)
    header = {**DEFAULT_HEADER_VALUES, **_json(template["header_json"], {})}
    order_number = normalize_customer_order_number(header.get("customer_order_number"))
    header["customer_order_number"] = order_number
    submitted = conn.execute(
        """SELECT 1 FROM order_interface_call_logs
           WHERE template_id=? AND interface_key='domestic_order_entry' AND status='success' LIMIT 1""",
        (template_id,),
    ).fetchone()
    now = utcnow()
    cursor = conn.execute(
        """INSERT INTO order_entry_template_groups
           (template_id,group_key,order_number,header_json,sort_order,status,nyeos_order_number,
            erp_order_number,submitted_at,created_at,updated_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (
            template_id, f"legacy-{template_id}", order_number,
            json.dumps(header, ensure_ascii=False), 0,
            "submitted" if submitted else "pending", "", "", now if submitted else None,
            str(template["created_at"]), str(template["updated_at"]),
        ),
    )
    group_id = int(cursor.lastrowid)
    conn.execute(
        "UPDATE order_entry_template_lines SET group_id=? WHERE template_id=? AND group_id IS NULL",
        (group_id, template_id),
    )
    conn.execute(
        """UPDATE order_interface_call_logs SET order_group_id=?
           WHERE template_id=? AND interface_key='domestic_order_entry' AND order_group_id IS NULL""",
        (group_id, template_id),
    )
    return list(conn.execute(
        "SELECT * FROM order_entry_template_groups WHERE template_id=? ORDER BY sort_order,id",
        (template_id,),
    ).fetchall())


def _serialize_template(conn, template_id: int) -> dict[str, Any]:
    template = conn.execute("SELECT * FROM order_entry_templates WHERE id=?", (template_id,)).fetchone()
    if not template:
        raise ValueError("录单模板不存在")
    groups = _ensure_template_groups(conn, template)
    rows = conn.execute(
        "SELECT * FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no,id", (template_id,)
    ).fetchall()
    serialized_lines = [
        {
            "id": row["id"],
            "group_id": row["group_id"],
            "line_no": row["line_no"],
            "values": {**_blank_line(int(row["line_no"])), **_json(row["values_json"], {})},
            "sources": _json(row["sources_json"], {}),
        }
        for row in rows
    ]
    group_items = []
    for group in groups:
        group_header = {**DEFAULT_HEADER_VALUES, **_json(group["header_json"], {})}
        group_header["customer_order_number"] = str(group["order_number"] or "")
        group_lines = [line for line in serialized_lines if int(line.get("group_id") or 0) == int(group["id"])]
        group_items.append({
            "id": int(group["id"]),
            "group_key": str(group["group_key"]),
            "order_number": str(group["order_number"] or ""),
            "display_order_number": (
                str(group["order_number"] or "")
                or f"暂无PO号-{str(group['group_key'])[:8]}"
            ),
            "header": group_header,
            "sort_order": int(group["sort_order"] or 0),
            "status": str(group["status"] or "pending"),
            "submitted": str(group["status"] or "") == "submitted",
            "nyeos_order_number": str(group["nyeos_order_number"] or ""),
            "erp_order_number": str(group["erp_order_number"] or ""),
            "submitted_at": group["submitted_at"],
            "lines": group_lines,
        })
    primary_header = group_items[0]["header"] if group_items else {**DEFAULT_HEADER_VALUES, **_json(template["header_json"], {})}
    result = {
        **dict(template),
        "header": primary_header,
        "lines": serialized_lines,
        "groups": group_items,
    }
    result["lines"] = _apply_customer_spec_matches(result["header"], result["lines"])
    lines_by_id = {int(line["id"]): line for line in result["lines"]}
    for group in result["groups"]:
        group["lines"] = [lines_by_id[int(line["id"])] for line in group["lines"]]
    return result


def get_or_create_template(
    case_id: int, employee_id: str, *, action_type: str = "new_order",
) -> tuple[dict[str, Any], dict[str, Any]]:
    case = _case_for_template(case_id, employee_id, action_type=action_type)
    with db_cursor() as conn:
        existing = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        if existing:
            template_id = int(existing["id"])
            template = _serialize_template(conn, template_id)
            legacy_row = conn.execute(
                "SELECT header_json FROM order_entry_templates WHERE id=?", (template_id,)
            ).fetchone()
            legacy_header = _json(legacy_row["header_json"] if legacy_row else "", {})
            if len(template.get("groups") or []) == 1:
                for field in HEADER_FIELDS:
                    if field == "customer_order_number" and not normalize_customer_order_number(legacy_header.get(field)):
                        continue
                    if clean_text(legacy_header.get(field)):
                        template["header"][field] = clean_text(legacy_header.get(field))
            previous_header = dict(template["header"])
            _apply_matched_customer_code(template["header"], case, overwrite=False)
            if template["header"] != previous_header or template["header"] != legacy_header:
                conn.execute(
                    "UPDATE order_entry_templates SET header_json=?,updated_at=? WHERE id=?",
                    (json.dumps(template["header"], ensure_ascii=False), utcnow(), template_id),
                )
                for group in template.get("groups") or []:
                    if group.get("submitted"):
                        continue
                    group_header = dict(group.get("header") or {})
                    for field in ("bill_to_customer_code", "ship_to_customer_code"):
                        if not clean_text(group_header.get(field)):
                            group_header[field] = clean_text(template["header"].get(field))
                    conn.execute(
                        "UPDATE order_entry_template_groups SET header_json=?,updated_at=? WHERE id=?",
                        (json.dumps(group_header, ensure_ascii=False), utcnow(), int(group["id"])),
                    )
                template = _serialize_template(conn, template_id)
            return case, template
        now = utcnow()
        initial_header, initial_lines = (
            _initial_template_data(case)
            if action_type == "new_order"
            else _initial_order_change_template_data(case)
        )
        initial_groups = _initial_order_groups(initial_header, initial_lines)
        used_line_nos: set[int] = set()
        next_line_no = 1
        for group in initial_groups:
            for entry in group["lines"]:
                values = entry["values"]
                requested = (
                    clean_text(values.get("customer_order_seq"))
                    if action_type == "new_order"
                    else clean_text(values.get("line_no"))
                ) or clean_text(values.get("line_no"))
                line_no = int(requested) if re.fullmatch(r"[1-9]\d*", requested or "") else next_line_no
                if line_no in used_line_nos:
                    line_no = next_line_no
                used_line_nos.add(line_no)
                next_line_no = max(next_line_no, line_no + 1)
                values["line_no"] = str(line_no)
                values["customer_order_seq"] = clean_text(values.get("customer_order_seq")) or str(line_no)
        if action_type == "new_order":
            price_snapshot = review_case_template_prices(
                case, {"header": initial_header, "lines": [
                    entry for group in initial_groups for entry in group["lines"]
                ]}
            )
            initial_header[PRICE_REVIEW_SNAPSHOT_KEY] = price_snapshot
            for group in initial_groups:
                group["header"][PRICE_REVIEW_SNAPSHOT_KEY] = price_snapshot
        cursor = conn.execute(
            "INSERT INTO order_entry_templates(case_id,employee_id,header_json,created_at,updated_at) VALUES (?,?,?,?,?)",
            (case_id, employee_id, json.dumps(initial_header, ensure_ascii=False), now, now),
        )
        template_id = int(cursor.lastrowid)
        for group in initial_groups:
            group_cursor = conn.execute(
                """INSERT INTO order_entry_template_groups
                   (template_id,group_key,order_number,header_json,sort_order,status,nyeos_order_number,
                    erp_order_number,submitted_at,created_at,updated_at)
                   VALUES (?,?,?,?,?,'pending','','',NULL,?,?)""",
                (
                    template_id, group["group_key"], group["order_number"],
                    json.dumps(group["header"], ensure_ascii=False), group["sort_order"], now, now,
                ),
            )
            group_id = int(group_cursor.lastrowid)
            for entry in group["lines"]:
                values = entry["values"]
                line_no = int(values["line_no"])
                conn.execute(
                    """INSERT INTO order_entry_template_lines
                       (template_id,group_id,line_no,values_json,sources_json,created_at,updated_at)
                       VALUES (?,?,?,?,?,?,?)""",
                    (
                        template_id, group_id, line_no,
                        json.dumps(values, ensure_ascii=False), json.dumps(entry["sources"], ensure_ascii=False), now, now,
                    ),
                )
        record_order_detail_event(
            conn,
            case_id=case_id,
            template_id=template_id,
            employee_id=employee_id,
            event_type="template_extracted" if action_type == "new_order" else "order_change_template_extracted",
            title="已提取内销录单模板" if action_type == "new_order" else "已提取修改订单模板",
            detail={"line_count": len(initial_lines), "group_count": len(initial_groups), "source": "邮件正文和附件"},
            operated_by=employee_id,
        )
        return case, _serialize_template(conn, template_id)


def get_saved_template(case_id: int, employee_id: str) -> tuple[dict[str, Any], dict[str, Any] | None]:
    """Read a saved template without creating or modifying one."""
    case = _case_for_template(case_id, employee_id)
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        return case, _serialize_template(conn, int(row["id"])) if row else None


def _template_task_row(case_id: int, employee_id: str) -> dict[str, Any] | None:
    """Return the latest extraction task for one employee-owned mail case."""
    with db_cursor() as conn:
        row = conn.execute(
            """SELECT * FROM order_entry_template_tasks
               WHERE case_id=? AND employee_id=?
               ORDER BY id DESC LIMIT 1""",
            (case_id, employee_id),
        ).fetchone()
    return dict(row) if row else None


def _queue_template_extraction(case_id: int, employee_id: str, *, action_type: str) -> dict[str, Any]:
    """Create one background extraction task without doing document work in HTTP.

    The original attachment can require native PDF parsing or OCR.  Keeping that
    work out of the request is important: a click should always return at once,
    while the case page can accurately show the durable task state.
    """
    _case_for_template(case_id, employee_id, action_type=action_type)
    template_label = "内销模板" if action_type == "new_order" else "修改订单模板"
    with db_cursor() as conn:
        template = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
        if template:
            return {
                "status": "completed", "task_id": None, "template_id": int(template["id"]),
                "message": f"{template_label}已生成，可直接打开核对。",
            }
        active = conn.execute(
            """SELECT * FROM order_entry_template_tasks
               WHERE case_id=? AND employee_id=? AND status IN ('queued', 'running')
               ORDER BY id DESC LIMIT 1""",
            (case_id, employee_id),
        ).fetchone()
        if active:
            return {
                "status": str(active["status"]), "task_id": int(active["id"]),
                "template_id": None, "message": "订单信息正在后台提取，请稍候刷新。",
            }
        now = utcnow()
        cursor = conn.execute(
            """INSERT INTO order_entry_template_tasks
               (case_id,employee_id,status,message,started_at)
               VALUES (?,?,?,?,?)""",
                (case_id, employee_id, "queued", f"等待后台提取{template_label}", now),
        )
        task_id = int(cursor.lastrowid)
        # Preserve the user's first start, independently of worker start/retries.
        conn.execute(
            """INSERT INTO order_intake_case_events
               (case_id,employee_id,action,before_json,after_json,created_at)
               SELECT ?,?,'processing_started','{}','{}',?
               WHERE NOT EXISTS (SELECT 1 FROM order_intake_case_events
                                 WHERE case_id=? AND action='processing_started')""",
            (case_id, employee_id, now, case_id),
        )

    command = [
        sys.executable, "-m", "fangzheng_web_app.order_entry_template_worker",
        "--task-id", str(task_id), "--case-id", str(case_id),
        "--employee-id", employee_id, "--action-type", action_type,
    ]
    try:
        subprocess.Popen(
            command,
            cwd=str(PROJECT_DIR),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
    except Exception as exc:
        _complete_template_extraction_task(
            task_id, status="error", message=f"无法启动后台提取：{exc}", template_id=None,
        )
        raise ValueError("无法启动后台订单提取，请稍后重试") from exc
    return {
        "status": "queued", "task_id": task_id, "template_id": None,
        "message": f"已开始后台提取订单信息，完成后会自动显示{template_label}。",
    }


def queue_template_extraction(case_id: int, employee_id: str) -> dict[str, Any]:
    return _queue_template_extraction(case_id, employee_id, action_type="new_order")


def queue_order_change_template_extraction(case_id: int, employee_id: str) -> dict[str, Any]:
    return _queue_template_extraction(case_id, employee_id, action_type="order_change")


def _complete_template_extraction_task(
    task_id: int, *, status: str, message: str, template_id: int | None,
) -> None:
    with db_cursor() as conn:
        conn.execute(
            """UPDATE order_entry_template_tasks
               SET status=?, message=?, template_id=?, completed_at=? WHERE id=?""",
            (status, message, template_id, utcnow(), task_id),
        )


def run_template_extraction_task(
    task_id: int, case_id: int, employee_id: str, *, action_type: str = "new_order",
) -> None:
    """Worker entry point.  This process owns the expensive first extraction."""
    with db_cursor() as conn:
        updated = conn.execute(
            """UPDATE order_entry_template_tasks
               SET status='running', message='正在解析邮件正文与附件', started_at=?
               WHERE id=? AND case_id=? AND employee_id=? AND status='queued'""",
            (utcnow(), task_id, case_id, employee_id),
        )
    if not updated.rowcount:
        return
    try:
        _case, template = get_or_create_template(case_id, employee_id, action_type=action_type)
        _complete_template_extraction_task(
            task_id,
            status="completed",
            message="订单信息已提取，请核对并保存订单。" if action_type == "new_order" else "订单修改信息已提取，请核对并保存。",
            template_id=int(template["id"]),
        )
    except Exception as exc:
        _complete_template_extraction_task(
            task_id,
            status="error",
            message=f"订单提取失败：{str(exc)[:240]}",
            template_id=None,
        )


def _replace_backup(
    conn,
    *,
    template_id: int,
    header: dict[str, Any],
    lines: list[dict[str, Any]],
    employee_id: str,
    saved_at: str,
) -> None:
    """Keep exactly one recoverable snapshot before replacing a template."""
    conn.execute("DELETE FROM order_entry_template_versions WHERE template_id=?", (template_id,))
    conn.execute(
        "INSERT INTO order_entry_template_versions(template_id,version_number,header_json,lines_json,saved_by,saved_at) VALUES (?,?,?,?,?,?)",
        (
            template_id,
            1,
            json.dumps(header, ensure_ascii=False),
            json.dumps(lines, ensure_ascii=False),
            employee_id,
            saved_at,
        ),
    )


def reextract_template(
    case_id: int, employee_id: str, *, action_type: str = "new_order",
) -> dict[str, Any]:
    """Rebuild one saved template's detail rows with the current extraction rules.

    The customer/header section is business-maintained and is intentionally left
    untouched. Before replacing the current detail rows, the immediately prior
    contents replace the single backup snapshot.
    """
    case = _case_for_template(case_id, employee_id, action_type=action_type)
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id,current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
        if not row:
            raise ValueError("请先打开修改订单模板" if action_type == "order_change" else "请先打开录单模板")
        template_id = int(row["id"])
        previous = _serialize_template(conn, template_id)

    # Recognition can involve OCR and file conversion, so do it outside of the
    # database transaction.  It only reads the original mail and attachments.
    if action_type == "order_change":
        regenerated_header, regenerated_lines = _initial_order_change_template_data(case)
        price_review_snapshot = None
    else:
        regenerated_header, regenerated_lines = _initial_template_data(case)
        price_review_snapshot = None
    now = utcnow()
    previous_lines = [
        {"values": line.get("values") or {}, "sources": line.get("sources") or {}}
        for line in previous.get("lines") or []
    ]
    previous_header = {**DEFAULT_HEADER_VALUES, **(previous.get("header") or {})}
    # Refresh adopts the PDF/图片转Excel template defaults and its extracted
    # order number, while retaining any customer information the business user
    # has already entered in this mail workspace.
    next_header = {
        **regenerated_header,
        **{
            field: value
            for field, value in previous_header.items()
            if field != PRICE_REVIEW_SNAPSHOT_KEY and clean_text(value)
        },
    }
    if price_review_snapshot is not None:
        next_header[PRICE_REVIEW_SNAPSHOT_KEY] = price_review_snapshot
    backup_version = 1
    current_version = 1
    submitted_line_nos = {
        int(line["line_no"])
        for group in previous.get("groups") or [] if group.get("submitted")
        for line in group.get("lines") or []
    } if action_type == "new_order" else set()

    with db_cursor() as conn:
        stale_resolution_tasks = conn.execute(
            """SELECT id,line_no,status,correlation_id,external_task_id
               FROM order_material_resolution_tasks
               WHERE template_id=?
               ORDER BY line_no""",
            (template_id,),
        ).fetchall()
        stale_resolution_tasks = [
            task for task in stale_resolution_tasks
            if int(task["line_no"]) not in submitted_line_nos
        ]
        _replace_backup(
            conn,
            template_id=template_id,
            header=previous_header,
            lines=previous_lines,
            employee_id=employee_id,
            saved_at=now,
        )
        # The regenerated detail rows no longer represent the material query
        # input that produced these candidates. Keep the call/event history,
        # but remove the active state so it cannot be selected or backfilled.
        for task in stale_resolution_tasks:
            conn.execute("DELETE FROM order_material_resolution_tasks WHERE id=?", (int(task["id"]),))
        conn.execute(
            "UPDATE order_entry_templates SET header_json=?,current_version=?,updated_at=? WHERE id=?",
            (json.dumps(next_header, ensure_ascii=False), current_version, now, template_id),
        )
        if action_type == "order_change":
            conn.execute("DELETE FROM order_entry_template_lines WHERE template_id=?", (template_id,))
            for entry in regenerated_lines:
                values = entry["values"]
                conn.execute(
                    "INSERT INTO order_entry_template_lines(template_id,line_no,values_json,sources_json,created_at,updated_at) VALUES (?,?,?,?,?,?)",
                    (
                        template_id,
                        int(values["line_no"] or 0),
                        json.dumps(values, ensure_ascii=False),
                        json.dumps(entry["sources"], ensure_ascii=False),
                        now,
                        now,
                    ),
                )
            invalidate_changed_order_matches(conn, template_id, regenerated_lines)
        else:
            submitted_groups = [group for group in previous.get("groups") or [] if group.get("submitted")]
            submitted_by_order = {
                normalize_customer_order_number(group.get("order_number")): group
                for group in submitted_groups if normalize_customer_order_number(group.get("order_number"))
            }
            fresh_groups: list[dict[str, Any]] = []
            ordered_submitted_ids: set[int] = set()
            next_sort_order = 0
            for group in _initial_order_groups(next_header, regenerated_lines):
                submitted_group = submitted_by_order.get(group.get("order_number"))
                if submitted_group:
                    submitted_group["next_sort_order"] = next_sort_order
                    ordered_submitted_ids.add(int(submitted_group["id"]))
                else:
                    group["next_sort_order"] = next_sort_order
                    fresh_groups.append(group)
                next_sort_order += 1
            for submitted_group in submitted_groups:
                if int(submitted_group["id"]) not in ordered_submitted_ids:
                    submitted_group["next_sort_order"] = next_sort_order
                    next_sort_order += 1
            for submitted_group in submitted_groups:
                conn.execute(
                    "UPDATE order_entry_template_groups SET sort_order=?,updated_at=? WHERE id=?",
                    (submitted_group["next_sort_order"], now, int(submitted_group["id"])),
                )
            pending_ids = [
                int(group["id"]) for group in previous.get("groups") or []
                if not group.get("submitted") and group.get("id")
            ]
            if pending_ids:
                placeholders = ",".join("?" for _ in pending_ids)
                conn.execute(
                    f"DELETE FROM order_entry_template_lines WHERE template_id=? AND group_id IN ({placeholders})",
                    (template_id, *pending_ids),
                )
                conn.execute(
                    f"UPDATE order_interface_call_logs SET order_group_id=NULL "
                    f"WHERE template_id=? AND order_group_id IN ({placeholders})",
                    (template_id, *pending_ids),
                )
                conn.execute(
                    f"DELETE FROM order_entry_template_groups WHERE template_id=? AND id IN ({placeholders})",
                    (template_id, *pending_ids),
                )
            used_line_nos = {
                int(line["line_no"])
                for group in submitted_groups for line in group.get("lines") or []
            }
            next_line_no = max(used_line_nos, default=0) + 1
            for group in fresh_groups:
                for entry in group.get("lines") or []:
                    values = entry["values"]
                    requested = clean_text(values.get("customer_order_seq")) or clean_text(values.get("line_no"))
                    line_no = int(requested) if re.fullmatch(r"[1-9]\d*", requested or "") else next_line_no
                    if line_no in used_line_nos:
                        line_no = next_line_no
                    used_line_nos.add(line_no)
                    next_line_no = max(next_line_no, line_no + 1)
                    values["line_no"] = str(line_no)
                    values["customer_order_seq"] = clean_text(values.get("customer_order_seq")) or str(line_no)
                group_header = {**next_header, **(group.get("header") or {})}
                group_header[PRICE_REVIEW_SNAPSHOT_KEY] = review_case_template_prices(
                    case, {"header": group_header, "lines": group.get("lines") or []}
                )
                cursor = conn.execute(
                    """INSERT INTO order_entry_template_groups
                       (template_id,group_key,order_number,header_json,sort_order,status,
                        nyeos_order_number,erp_order_number,submitted_at,created_at,updated_at)
                       VALUES (?,?,?,?,?,'pending','','',NULL,?,?)""",
                    (
                        template_id, group["group_key"], group.get("order_number") or "",
                        json.dumps(group_header, ensure_ascii=False), group["next_sort_order"], now, now,
                    ),
                )
                group_id = int(cursor.lastrowid)
                for entry in group.get("lines") or []:
                    values = entry["values"]
                    line_no = int(values["line_no"])
                    conn.execute(
                        """INSERT INTO order_entry_template_lines
                           (template_id,group_id,line_no,values_json,sources_json,created_at,updated_at)
                           VALUES (?,?,?,?,?,?,?)""",
                        (
                            template_id, group_id, line_no,
                            json.dumps(values, ensure_ascii=False),
                            json.dumps(entry["sources"], ensure_ascii=False), now, now,
                        ),
                    )
            primary_group = conn.execute(
                """SELECT header_json FROM order_entry_template_groups
                   WHERE template_id=? ORDER BY sort_order,id LIMIT 1""",
                (template_id,),
            ).fetchone()
            if primary_group:
                conn.execute(
                    "UPDATE order_entry_templates SET header_json=? WHERE id=?",
                    (str(primary_group["header_json"]), template_id),
                )
        record_order_detail_event(
            conn,
            case_id=case_id,
            template_id=template_id,
            employee_id=employee_id,
            event_type=("order_change_template_reextracted" if action_type == "order_change" else "template_reextracted"),
            title=("已重新提取修改订单明细" if action_type == "order_change" else "已重新提取订单明细"),
            detail={
                "previous_line_count": len(previous_lines),
                "line_count": len(regenerated_lines),
                "cleared_material_resolution_tasks": [
                    {
                        "line_no": int(task["line_no"]),
                        "status": str(task["status"]),
                        "correlation_id": str(task["correlation_id"]),
                        "external_task_id": str(task["external_task_id"] or ""),
                    }
                    for task in stale_resolution_tasks
                ],
            },
            operated_by=employee_id,
        )
        template = _serialize_template(conn, template_id)

    return {
        "case_id": case_id,
        "subject": str(case.get("subject") or ""),
        "previous_line_count": len(previous_lines),
        "line_count": len(template.get("lines") or []),
        "backup_version": backup_version,
        "current_version": current_version,
        "template": template,
    }


def reextract_order_change_template(case_id: int, employee_id: str) -> dict[str, Any]:
    return reextract_template(case_id, employee_id, action_type="order_change")


def reextract_all_templates(employee_id: str) -> dict[str, Any]:
    """Batch re-extract every existing domestic template owned by one user."""
    with db_cursor() as conn:
        rows = conn.execute(
            """SELECT t.case_id
               FROM order_entry_templates t
               JOIN order_intake_cases c ON c.id=t.case_id
               WHERE t.employee_id=? AND c.action_type='new_order'
               ORDER BY t.id""",
            (employee_id,),
        ).fetchall()
    results = [reextract_template(int(row["case_id"]), employee_id) for row in rows]
    return {
        "template_count": len(results),
        "previous_line_count": sum(item["previous_line_count"] for item in results),
        "line_count": sum(item["line_count"] for item in results),
        "results": results,
    }


def template_progress(case_id: int, employee_id: str) -> dict[str, Any]:
    progress = _entry_progress(case_id, employee_id)
    with db_cursor() as conn:
        sent = conn.execute("SELECT 1 FROM order_entry_detail_events WHERE case_id=? AND operated_by=? AND event_type='order_reply_sent' LIMIT 1", (case_id, employee_id)).fetchone()
    progress['replied'] = bool(sent)
    progress['closed'] = bool(progress['completed'])
    progress['operation_label'] = progress['label']
    progress['task_status'] = (
        ('entry_replied' if sent else 'entry_pending_reply')
        if progress['completed']
        else ('pending_entry' if progress['created'] else 'pending_template_generation')
    )
    progress['label'] = {
        'pending_template_generation': '待生成模板',
        'pending_entry': '待录单',
        'entry_pending_reply': '录单完成待回复',
        'entry_replied': '录单完成已回复',
    }[progress['task_status']]
    progress['next_action'] = ('查看订单' if sent else '回复邮件') if progress['completed'] else '去录单'
    return progress


ORDER_CHANGE_LINE_FIELDS = (
    "customer_order_number", "line_no", "customer_product_code", "customer_spec",
    "delivery_date", "quantity",
)


def order_change_template_progress(case_id: int, employee_id: str) -> dict[str, Any]:
    """Return the lightweight extraction/save state for an order-change mail."""
    case = _case_for_template(case_id, employee_id, action_type="order_change")
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
    version = int(row["current_version"] or 0) if row else 0
    if row:
        if case.get("status") == "pending_reply":
            return {
                "created": True, "saved": True, "version": version,
                "stage": "pending_reply", "label": "待回复邮件",
                "next_action": "回复邮件", "step": 4,
            }
        return {
            "created": True, "saved": version > 0, "version": version,
            "stage": "saved" if version > 0 else "pending_template_save",
            "label": "修改模板已保存" if version > 0 else "待核对并保存修改模板",
            "next_action": "打开修改订单" if version > 0 else "核对并保存修改订单",
            "step": 3,
        }
    task = _template_task_row(case_id, employee_id)
    if task and task.get("status") in {"queued", "running"}:
        return {
            "created": False, "saved": False, "version": 0, "stage": "extracting",
            "label": "正在提取修改订单", "next_action": "正在后台提取订单信息…", "step": 2,
            "task_id": task["id"], "task_message": task.get("message") or "正在准备订单信息。",
        }
    if task and task.get("status") == "error":
        return {
            "created": False, "saved": False, "version": 0, "stage": "extraction_error",
            "label": "提取失败", "next_action": "重新提取修改订单", "step": 2,
            "task_id": task["id"], "task_message": task.get("message") or "请重新提取订单信息。",
        }
    return {
        "created": False, "saved": False, "version": 0, "stage": "pending_extraction",
        "label": "待提取修改订单", "next_action": "提取订单到修改模板", "step": 2,
    }


def get_order_change_template(case_id: int, employee_id: str) -> tuple[dict[str, Any], dict[str, Any] | None]:
    case = _case_for_template(case_id, employee_id, action_type="order_change")
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        return case, _serialize_template(conn, int(row["id"])) if row else None


def save_order_change_template(case_id: int, employee_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    """Save the six user-facing fields without entering the domestic order flow."""
    case = _case_for_template(case_id, employee_id, action_type="order_change")
    if case.get("status") == "pending_reply":
        raise ValueError("修改订单已提交 APS，不能再修改模板。")
    raw_lines = payload.get("lines") or []
    if not isinstance(raw_lines, list):
        raise ValueError("修改订单明细格式无效")
    lines: list[dict[str, Any]] = []
    used_line_nos: set[str] = set()
    for index, raw in enumerate(raw_lines, start=1):
        source = (raw or {}).get("sources") or {}
        raw_values = (raw or {}).get("values") or raw or {}
        values = _blank_line(index)
        values["material_status"] = "查询"
        for field in ORDER_CHANGE_LINE_FIELDS:
            values[field] = clean_text(raw_values.get(field))
        line_no = values["line_no"]
        values["line_no"] = line_no if re.fullmatch(r"[1-9]\d*", line_no or "") else str(index)
        if not any(values.get(field) for field in ORDER_CHANGE_LINE_FIELDS if field != "line_no"):
            continue
        if values["line_no"] in used_line_nos:
            raise ValueError(f"项次 {values['line_no']} 重复")
        used_line_nos.add(values["line_no"])
        lines.append({"values": values, "sources": source if isinstance(source, dict) else {}})
    if not lines:
        lines = [{"values": _blank_line(1), "sources": {}}]

    now = utcnow()
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id,current_version,header_json FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
        if not row:
            raise ValueError("请先提取修改订单模板")
        template_id = int(row["id"])
        previous = _serialize_template(conn, template_id)
        invalidate_changed_order_matches(conn, template_id, lines)
        header = {**DEFAULT_HEADER_VALUES, **_json(row["header_json"], {})}
        first_order_number = next((item["values"]["customer_order_number"] for item in lines if item["values"].get("customer_order_number")), "")
        if first_order_number:
            header["customer_order_number"] = first_order_number
        _replace_backup(
            conn, template_id=template_id, header=previous["header"], lines=previous["lines"],
            employee_id=employee_id, saved_at=now,
        )
        conn.execute(
            "UPDATE order_entry_templates SET header_json=?,current_version=?,updated_at=? WHERE id=?",
            (json.dumps(header, ensure_ascii=False), int(row["current_version"] or 0) + 1, now, template_id),
        )
        conn.execute("DELETE FROM order_entry_template_lines WHERE template_id=?", (template_id,))
        for entry in lines:
            values = entry["values"]
            conn.execute(
                "INSERT INTO order_entry_template_lines(template_id,line_no,values_json,sources_json,created_at,updated_at) VALUES (?,?,?,?,?,?)",
                (template_id, int(values["line_no"]), json.dumps(values, ensure_ascii=False),
                 json.dumps(entry["sources"], ensure_ascii=False), now, now),
            )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="order_change_template_saved", title="修改订单模板已保存",
            detail={"line_count": len(lines)}, operated_by=employee_id,
        )
        return _serialize_template(conn, template_id)


def _entry_progress(case_id: int, employee_id: str) -> dict[str, Any]:
    """Return the single read-only workflow state used by every order view.

    A successful domestic-entry interface call is the terminal business fact.
    A completed order has a saved template too, but must not be shown as still
    waiting for an interface submission on another page.
    """
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
        group_counts = conn.execute(
            """SELECT COUNT(*) AS total,
                      SUM(CASE WHEN g.status='submitted' THEN 1 ELSE 0 END) AS submitted
               FROM order_entry_template_groups g
               JOIN order_entry_templates t ON t.id=g.template_id
               WHERE t.case_id=? AND t.employee_id=?""",
            (case_id, employee_id),
        ).fetchone()
        if group_counts and int(group_counts["total"] or 0) > 0:
            completed = int(group_counts["submitted"] or 0) == int(group_counts["total"])
        else:
            completed = bool(conn.execute(
                """SELECT 1 FROM order_interface_call_logs
                   WHERE case_id=? AND employee_id=? AND interface_key='domestic_order_entry'
                     AND status='success'
                   LIMIT 1""",
                (case_id, employee_id),
            ).fetchone())
    version = int(row["current_version"] or 0) if row else 0
    if completed:
        return {
            "created": bool(row), "saved": bool(row and version > 0), "completed": True,
            "version": version, "stage": "completed", "label": "已完成",
            "next_action": "录单已完成", "step": 5,
        }
    if not row:
        task = _template_task_row(case_id, employee_id)
        if task and task.get("status") in {"queued", "running"}:
            return {
                "created": False, "saved": False, "completed": False, "version": 0,
                "stage": "extracting", "label": "正在提取订单",
                "next_action": "正在后台提取订单信息…", "step": 2,
                "task_id": task["id"], "task_status": task["status"],
                "task_message": task.get("message") or "正在准备订单信息。",
            }
        if task and task.get("status") == "error":
            return {
                "created": False, "saved": False, "completed": False, "version": 0,
                "stage": "extraction_error", "label": "提取失败",
                "next_action": "重新提取订单", "step": 2,
                "task_id": task["id"], "task_status": "error",
                "task_message": task.get("message") or "请重新提取订单信息。",
            }
        return {
            "created": False, "saved": False, "completed": False, "version": 0,
            "stage": "pending_extraction", "label": "待提取订单",
            "next_action": "提取订单到内销模板", "step": 2,
        }
    if version <= 0:
        return {
            "created": True, "saved": False, "completed": False, "version": 0,
            "stage": "pending_template_save", "label": "待核对并保存订单",
            "next_action": "核对并保存订单", "step": 3,
        }
    return {
        "created": True, "saved": True, "completed": False, "version": version,
        "stage": "pending_interface_submit", "label": "待批量料号查询",
        "next_action": "批量料号查询", "step": 4,
    }


def _clean_values(values: dict[str, Any], fields: tuple[str, ...]) -> dict[str, str]:
    return {field: str(values.get(field) or "").strip() for field in fields}


def _template_changes(
    previous_header: dict[str, Any],
    previous_lines: list[dict[str, Any]],
    header: dict[str, Any],
    lines: list[dict[str, Any]],
) -> list[dict[str, str]]:
    changes: list[dict[str, str]] = []
    for field in HEADER_FIELDS:
        before = str(previous_header.get(field) or "")
        after = str(header.get(field) or "")
        if before != after:
            changes.append({"field": HEADER_LABELS[field], "before": before, "after": after, "scope": "表头"})
    old_by_line = {
        str((entry.get("values") or {}).get("line_no") or index): entry.get("values") or {}
        for index, entry in enumerate(previous_lines, start=1)
    }
    for index, entry in enumerate(lines, start=1):
        values = entry.get("values") or {}
        before_values = old_by_line.get(str(values.get("line_no") or index), {})
        for field in LINE_FIELDS:
            if field == "line_no":
                continue
            before = str(before_values.get(field) or "")
            after = str(values.get(field) or "")
            if before != after:
                changes.append({
                    "field": LINE_LABELS[field], "before": before, "after": after,
                    "scope": f"第 {values.get('line_no') or index} 行",
                })
    return changes[:120]


def save_template(case_id: int, employee_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    _case_for_template(case_id, employee_id)
    raw_groups = payload.get("groups")
    legacy_payload = not isinstance(raw_groups, list)
    if legacy_payload:
        raw_groups = [{
            "group_key": "",
            "header": payload.get("header") or {},
            "lines": payload.get("lines") or [],
        }]
    if not raw_groups:
        raise ValueError("至少保留一个 PO 分组")

    prepared_groups: list[dict[str, Any]] = []
    used_order_numbers: set[str] = set()
    for group_index, raw_group in enumerate(raw_groups, start=1):
        if not isinstance(raw_group, dict):
            raise ValueError("PO 分组格式无效")
        raw_header = raw_group.get("header") or {}
        raw_lines = raw_group.get("lines") or []
        if not isinstance(raw_header, dict) or not isinstance(raw_lines, list):
            raise ValueError(f"第 {group_index} 个 PO 分组格式无效")
        header = dict(DEFAULT_HEADER_VALUES)
        for field in HEADER_FIELDS:
            if field in raw_header:
                header[field] = str(raw_header.get(field) or "").strip()
        order_number = normalize_customer_order_number(
            raw_group.get("order_number") or header.get("customer_order_number")
        )
        header["customer_order_number"] = order_number
        if order_number and order_number in used_order_numbers:
            raise ValueError(f"客户订单号 {order_number} 存在重复分组")
        if order_number:
            used_order_numbers.add(order_number)
        group_lines: list[dict[str, Any]] = []
        for line_index, raw in enumerate(raw_lines, start=1):
            values = _clean_values((raw or {}).get("values") or raw or {}, PERSISTED_LINE_FIELDS)
            if not any(values[field] for field in PERSISTED_LINE_FIELDS if field not in {"line_no", "material_status", "customer_order_number"}):
                continue
            values["customer_order_number"] = order_number
            values["material_status"] = values["material_status"] or "查询"
            if values["material_status"] not in MATERIAL_STATUS_VALUES:
                raise ValueError(f"第 {line_index} 行料号状态只能选择“查询”或“新增”")
            sources = (raw or {}).get("sources") or {}
            group_lines.append({"values": values, "sources": sources if isinstance(sources, dict) else {}})
        prepared_groups.append({
            "group_key": str(raw_group.get("group_key") or "").strip() or uuid.uuid4().hex,
            "order_number": order_number,
            "header": header,
            "lines": group_lines,
            "sort_order": group_index - 1,
            "provided_header_fields": set(raw_header),
        })

    now = utcnow()
    with db_cursor() as conn:
        template = conn.execute(
            "SELECT id,current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        if not template:
            raise ValueError("请先打开录单模板")
        template_id = int(template["id"])
        previous = _serialize_template(conn, template_id)
        existing_groups = {group["group_key"]: group for group in previous.get("groups") or []}
        if legacy_payload and previous.get("groups"):
            prepared_groups[0]["group_key"] = previous["groups"][0]["group_key"]
        incoming_keys = {group["group_key"] for group in prepared_groups}
        submitted_groups = [group for group in previous.get("groups") or [] if group.get("submitted")]
        missing_submitted = [group["order_number"] or "待分配" for group in submitted_groups if group["group_key"] not in incoming_keys]
        if missing_submitted:
            raise ValueError("已提交 PO 不允许删除：" + "、".join(missing_submitted))

        for group in prepared_groups:
            existing = existing_groups.get(group["group_key"])
            if legacy_payload and existing:
                for field in HEADER_FIELDS:
                    if field not in group["provided_header_fields"]:
                        group["header"][field] = clean_text(existing["header"].get(field))
                group["order_number"] = normalize_customer_order_number(group["header"].get("customer_order_number"))
            if existing and existing.get("submitted"):
                submitted_header = {field: clean_text(existing["header"].get(field)) for field in HEADER_FIELDS}
                incoming_header = {field: clean_text(group["header"].get(field)) for field in HEADER_FIELDS}
                submitted_lines = [line.get("values") or {} for line in existing["lines"]]
                incoming_lines = [line.get("values") or {} for line in group["lines"]]
                if incoming_header != submitted_header or incoming_lines != submitted_lines:
                    raise ValueError(f"PO {existing['order_number']} 已提交，不允许修改")
                continue
            snapshot = ((existing or {}).get("header") or {}).get(PRICE_REVIEW_SNAPSHOT_KEY)
            if isinstance(snapshot, dict):
                group["header"][PRICE_REVIEW_SNAPSHOT_KEY] = snapshot

        previous_header = {**DEFAULT_HEADER_VALUES, **(previous.get("header") or {})}
        previous_lines = [
            {"values": line.get("values") or {}, "sources": line.get("sources") or {}}
            for line in previous.get("lines") or []
        ]
        editable_groups = [group for group in prepared_groups if not (existing_groups.get(group["group_key"]) or {}).get("submitted")]
        editable_lines = [line for group in editable_groups for line in group["lines"]]
        primary_header = prepared_groups[0]["header"]
        changes = _template_changes(previous_header, previous_lines, primary_header, editable_lines)
        _replace_backup(
            conn,
            template_id=template_id,
            header=previous_header,
            lines=previous_lines,
            employee_id=employee_id,
            saved_at=now,
        )
        pending_group_ids = [int(group["id"]) for group in previous.get("groups") or [] if not group.get("submitted")]
        if pending_group_ids:
            placeholders = ",".join("?" for _ in pending_group_ids)
            conn.execute(
                f"DELETE FROM order_entry_template_lines WHERE template_id=? AND group_id IN ({placeholders})",
                (template_id, *pending_group_ids),
            )
        if pending_group_ids:
            placeholders = ",".join("?" for _ in pending_group_ids)
            conn.execute(
                f"UPDATE order_interface_call_logs SET order_group_id=NULL "
                f"WHERE template_id=? AND order_group_id IN ({placeholders})",
                (template_id, *pending_group_ids),
            )
        conn.execute(
            "DELETE FROM order_entry_template_groups WHERE template_id=? AND status<>'submitted'",
            (template_id,),
        )

        used_line_nos = {
            int(line["line_no"])
            for group in submitted_groups for line in group.get("lines") or []
        }
        next_line_no = max(used_line_nos, default=0) + 1
        for group in prepared_groups:
            existing = existing_groups.get(group["group_key"])
            if existing and existing.get("submitted"):
                continue
            group_cursor = conn.execute(
                """INSERT INTO order_entry_template_groups
                   (template_id,group_key,order_number,header_json,sort_order,status,nyeos_order_number,
                    erp_order_number,submitted_at,created_at,updated_at)
                   VALUES (?,?,?,?,?,'pending','','',NULL,?,?)""",
                (
                    template_id, group["group_key"], group["order_number"],
                    json.dumps(group["header"], ensure_ascii=False), group["sort_order"], now, now,
                ),
            )
            group_id = int(group_cursor.lastrowid)
            group["lines"] = _apply_customer_spec_matches(group["header"], group["lines"])
            for entry in group["lines"]:
                values = entry["values"]
                requested = clean_text(values.get("customer_order_seq")) or clean_text(values.get("line_no"))
                line_no = int(requested) if re.fullmatch(r"[1-9]\d*", requested or "") else next_line_no
                if line_no in used_line_nos:
                    line_no = next_line_no
                used_line_nos.add(line_no)
                next_line_no = max(next_line_no, line_no + 1)
                values["line_no"] = str(line_no)
                values["customer_order_seq"] = clean_text(values.get("customer_order_seq")) or str(line_no)
                conn.execute(
                    """INSERT INTO order_entry_template_lines
                       (template_id,group_id,line_no,values_json,sources_json,created_at,updated_at)
                       VALUES (?,?,?,?,?,?,?)""",
                    (
                        template_id, group_id, line_no, json.dumps(values, ensure_ascii=False),
                        json.dumps(entry["sources"], ensure_ascii=False), now, now,
                    ),
                )

        remaining_line_nos = {
            int(row["line_no"]) for row in conn.execute(
                "SELECT line_no FROM order_entry_template_lines WHERE template_id=?", (template_id,)
            ).fetchall()
        }
        saved_line_values = {
            int(row["line_no"]): _json(row["values_json"], {})
            for row in conn.execute(
                "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=?", (template_id,)
            ).fetchall()
        }
        stale_tasks = conn.execute(
            "SELECT line_no FROM order_material_resolution_tasks WHERE template_id=?", (template_id,)
        ).fetchall()
        cleared_material_resolution_tasks = [
            int(row["line_no"]) for row in stale_tasks
            if int(row["line_no"]) not in remaining_line_nos
            or not any(clean_text(saved_line_values[int(row["line_no"])].get(field)) for field in ("product_code", "product_name"))
        ]
        for line_no in cleared_material_resolution_tasks:
            conn.execute(
                "DELETE FROM order_material_resolution_tasks WHERE template_id=? AND line_no=?",
                (template_id, line_no),
            )

        conn.execute(
            "UPDATE order_entry_templates SET header_json=?,current_version=?,updated_at=? WHERE id=?",
            (json.dumps(primary_header, ensure_ascii=False), 1, now, template_id),
        )
        record_order_detail_event(
            conn,
            case_id=case_id,
            template_id=template_id,
            employee_id=employee_id,
            event_type="template_saved",
            title="保存内销录单模板",
            detail={
                "changes": changes,
                "line_count": len(remaining_line_nos),
                "group_count": len(prepared_groups),
                "cleared_material_resolution_tasks": cleared_material_resolution_tasks,
            },
            operated_by=employee_id,
        )
        return _serialize_template(conn, template_id)


def validation_issues(template: dict[str, Any]) -> list[str]:
    issues = [f"{HEADER_LABELS[field]}未填写" for field in REQUIRED_HEADER_FIELDS if not str(template["header"].get(field) or "").strip()]
    for line in template.get("lines") or []:
        for field in REQUIRED_LINE_FIELDS:
            if not str((line.get("values") or {}).get(field) or "").strip():
                issues.append(f"第 {line.get('line_no')} 行{LINE_LABELS[field]}未填写")
    return issues


def build_domestic_export(
    case_id: int, employee_id: str, group_key: str = "",
) -> tuple[BytesIO, str]:
    _case_for_template(case_id, employee_id)
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        if not row:
            raise ValueError("请先保存模板内容后再下载")
        template = _serialize_template(conn, int(row["id"]))
    groups = template.get("groups") or []
    if group_key:
        group = next((item for item in groups if item.get("group_key") == group_key), None)
        if not group:
            raise ValueError("当前 PO 分组不存在")
        template = {**template, "header": group["header"], "lines": group["lines"]}
    if int(template.get("current_version") or 0) <= 0:
        raise ValueError("请先保存模板内容后再下载")
    if not DOMESTIC_TEMPLATE_PATH.is_file():
        raise ValueError("内销录单模板文件未配置")
    book = load_workbook(DOMESTIC_TEMPLATE_PATH)
    sheet = book["内销"]
    for index, field in enumerate(HEADER_FIELDS, start=1):
        required_label = "（必填）" if field in REQUIRED_HEADER_FIELDS else "（选填）"
        sheet.cell(1, index).value = f"{HEADER_LABELS[field]}{required_label}"
        sheet.cell(2, index).value = template["header"].get(field) or None
    for index, field in enumerate(LINE_FIELDS, start=1):
        required_label = "（系统生成）" if field == "line_no" else "（必填）" if field in REQUIRED_LINE_FIELDS else "（选填）"
        sheet.cell(3, index).value = f"{LINE_LABELS[field]}{required_label}"
    required_rows = 3 + max(1, len(template["lines"]))
    style_source_row = 4
    while sheet.max_row < required_rows:
        target = sheet.max_row + 1
        for col in range(1, len(LINE_FIELDS) + 1):
            source, cell = sheet.cell(style_source_row, col), sheet.cell(target, col)
            cell._style = copy(source._style)
            cell.number_format = source.number_format
    for row in range(4, max(sheet.max_row, required_rows) + 1):
        values = (template["lines"][row - 4]["values"] if row - 4 < len(template["lines"]) else {})
        for col, field in enumerate(LINE_FIELDS, start=1):
            value = values.get(field) or None
            if field in {"quantity", "price_before_tax", "unit_price"} and value not in (None, ""):
                try:
                    value = float(value)
                except ValueError:
                    pass
            if field == "delivery_date" and value:
                try:
                    value = datetime.fromisoformat(str(value)).date()
                except ValueError:
                    pass
            sheet.cell(row, col).value = value
    data = BytesIO()
    book.save(data)
    book.close()
    data.seek(0)
    order_number = normalize_customer_order_number(template["header"].get("customer_order_number"))
    suffix = re.sub(r"[^0-9A-Za-z_-]+", "_", order_number).strip("_")
    suffix = f"_{suffix}" if suffix else ""
    return data, f"内销录单_邮件{case_id}{suffix}_v{template['current_version']}.xlsx"
