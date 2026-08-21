from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from copy import copy
from decimal import Decimal, InvalidOperation
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .database import automation_cursor as db_cursor
from .customer_archive_service import get_enabled_extraction_maps
from .db import utcnow
from .excel_utils import load_workbook_compat
from .file_storage import resolve_attachment_path
from .order_intake_service import get_case
from .paths import PACKAGE_DIR
from .purchase_field_rules import clean_text, normalize_date, normalize_number
from .purchase_order_pipeline import run_purchase_order_pipeline


DOMESTIC_TEMPLATE_PATH = PACKAGE_DIR / "default_rules" / "order_entry" / "营管151查询和录单导入模板.xlsx"

HEADER_FIELDS = (
    "order_type", "type_1", "type_2", "bill_to_customer_code",
    "ship_to_customer_code", "delivery_factory", "customer_order_number", "ledger",
)
LINE_FIELDS = (
    "line_no", "product_code", "product_name", "customer_product_code",
    "customer_spec", "delivery_date", "quantity", "price_before_tax",
    "unit_price", "origin", "customer_order_seq", "customer_order_number",
    "one_to_many", "remark",
)
HEADER_LABELS = {
    "order_type": "单别", "type_1": "类型1", "type_2": "类型2",
    "bill_to_customer_code": "账款客户编号", "ship_to_customer_code": "送货客户编号",
    "delivery_factory": "送货厂别", "customer_order_number": "客户订单号", "ledger": "账套",
}
LINE_LABELS = {
    "line_no": "项次", "product_code": "产品编号", "product_name": "品名",
    "customer_product_code": "客户产品编号", "customer_spec": "客户规格",
    "delivery_date": "出货日期", "quantity": "数量", "price_before_tax": "税前单价",
    "unit_price": "单价", "origin": "产地", "customer_order_seq": "客户订单序号",
    "customer_order_number": "客户订单号", "one_to_many": "一对多", "remark": "备注",
}
REQUIRED_HEADER_FIELDS = {"order_type", "bill_to_customer_code", "ledger"}
REQUIRED_LINE_FIELDS = {"line_no", "customer_product_code", "quantity"}

# These fields are deliberately not inferred from an order e-mail.  产品编号、品名
# will later come from Nyeos 151; 产地和一对多 are business decisions.  Leaving a
# value blank is safer than presenting an unverified guess as a usable result.
MANUAL_ONLY_LINE_FIELDS = {"product_code", "product_name", "origin", "one_to_many"}

# The source documents use different wording.  Keep this mapping here instead
# of forcing business users to normalise their customers' Excel files first.
_ATTACHMENT_HEADERS = {
    "line_no": {"序号", "项次", "项目", "行号", "item", "no"},
    "product_code": {"产品编号", "物料编号", "物料编码", "料号", "品号", "厂内料号"},
    "product_name": {"品名", "物料名称", "名称", "产品名称"},
    "customer_product_code": {"客户产品编号", "客户料号", "客户物料编号", "客户产品码", "part no", "p/n"},
    "customer_spec": {"客户规格", "规格", "型号", "名称规格", "物料规格", "物料描述"},
    "delivery_date": {"出货日期", "交货日期", "交期", "到货日期", "delivery date"},
    "quantity": {"数量", "采购量", "订购数量", "订单数量", "qty", "quantity"},
    "_quantity_unit": {"单位", "计量单位", "数量单位", "uom"},
    "price_before_tax": {"税前单价", "未税单价", "不含税单价"},
    "unit_price": {"单价", "含税单价", "unit price", "price"},
    "origin": {"产地", "原产地"},
    "customer_order_seq": {"客户订单序号", "订单序号", "客户项次"},
    "customer_order_number": {"客户订单号", "订单号", "采购订单号", "采购单号", "po号", "po no"},
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


def _case_for_template(case_id: int, employee_id: str) -> dict[str, Any]:
    case = get_case(case_id, employee_id)
    if not case:
        raise ValueError("录单邮件不存在或无权操作")
    if case.get("action_type") != "new_order":
        raise ValueError("只有已分流为“录单”的邮件才能提取到内销模板")
    return case


def _blank_line(line_no: int) -> dict[str, str]:
    return {field: str(line_no) if field == "line_no" else "" for field in LINE_FIELDS}


def _is_pp_spec(value: str) -> bool:
    """Only classify PP when the customer specification explicitly says so."""
    text = clean_text(value)
    return bool(re.search(r"(?:半固化片|(?<![A-Za-z0-9])PP(?![A-Za-z0-9]))", text, re.IGNORECASE))


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


def _append_remark(remark: str, meter: Decimal | None) -> str:
    if meter is None:
        return remark
    meter_text = f"PP米数：{_decimal_text(meter)}米"
    return "；".join(item for item in (remark, meter_text) if item)


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
    result["remark"] = _append_remark(result.get("remark", ""), meter)

    unit = clean_text(quantity_unit)
    if "张" in unit:
        # PP 小片：客户明确以张计数，数量直接保留，不做米数换算。
        return result
    if "卷" not in unit or meter is None:
        # PP 的单位或米数不明确，不能猜测换算关系。
        result["quantity"] = ""
        return result
    try:
        quantity = Decimal(str(result.get("quantity") or "").replace(",", ""))
    except InvalidOperation:
        result["quantity"] = ""
        return result
    if quantity < 0:
        result["quantity"] = ""
        return result
    result["quantity"] = _decimal_text(quantity * meter)
    return result


def _split_body_order_rows(body_text: str) -> list[dict[str, Any]]:
    """Extract simple ERP-style rows from the line-oriented mail body.

    This is a safe first path for HTML mail tables. Attachment-specific parsers
    will feed the same structure in the next extraction layer.
    """
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
    return {field: {"label": label, "reference": reference} for field, value in values.items() if value and field != "line_no"}


def _line_entry(
    values: dict[str, Any], *, label: str, reference: str, line_no: int = 1, quantity_unit: Any = "", product_context: Any = "",
) -> dict[str, Any]:
    line = _blank_line(line_no)
    for field in LINE_FIELDS:
        value = values.get(field)
        if value not in (None, ""):
            line[field] = clean_text(value)
    line["line_no"] = str(line_no)
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
    values = [clean_text(_value_by_alias(source, label)) for label in labels]
    if not all(values):
        return ""
    if transform_type == "concat":
        return " ".join(values)
    return values[0] if len(values) == 1 else ""


def _apply_customer_extraction_mappings(
    values: dict[str, Any], source: dict[str, Any], mappings: list[dict[str, Any]],
) -> dict[str, Any]:
    """Apply only a uniquely identified customer's attachment-table mappings.

    Manual mappings explicitly blank the target.  For all other mapping source
    types we leave the universal result untouched for now; that preserves the
    source-of-truth rule until a dedicated body/filename extractor exists.
    """
    if not mappings:
        return values
    result = dict(values)
    for mapping in mappings:
        if mapping.get("source_kind") != "attachment_table":
            continue
        target = str(mapping.get("target_field") or "")
        transform = str(mapping.get("transform_type") or "")
        if target not in LINE_FIELDS:
            continue
        if transform == "manual":
            result[target] = ""
            continue
        value = _mapping_source_value(source, str(mapping.get("source_label") or ""), transform)
        # A configured customer mapping takes precedence only when it produces
        # an explicit source value.  Otherwise retain an already verified
        # universal value; no fallback guess is introduced.
        if value:
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


def _rows_from_excel(path: Path, filename: str, customer_mappings: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    """Read customer Excel attachments as structured data without OCR."""
    try:
        book = load_workbook_compat(path, data_only=True)
    except Exception:
        return []
    try:
        result: list[dict[str, Any]] = []
        for sheet in book.worksheets:
            rows = list(sheet.iter_rows(values_only=True))
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
        "customer_product_code": _value_by_alias(original, "客户产品编号", "客户料号", "客户物料编号", "客户产品码") or "",
        "customer_spec": raw_spec or standard.get("说明") or standard.get("物料名称") or "",
        "delivery_date": standard.get("交货日期") or _value_by_alias(original, "交货日期", "出货日期", "交期", "Delivery Date") or "",
        "quantity": standard.get("数量") or _value_by_alias(original, "数量", "Quantity", "Qty") or "",
        "price_before_tax": raw_before_tax_price or standard.get("不含税单价") or "",
        "unit_price": raw_unit_price or "",
        "customer_order_number": order_number,
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
    try:
        document = run_purchase_order_pipeline({"stored_path": str(path), "original_filename": filename})
    except Exception:
        # An unreadable scanned attachment must not stop opening the template;
        # business can still fill the row manually from the downloaded source.
        return []
    header = document.get("header_info") or {}
    order_number = _canonical_order_number(header)
    rows = document.get("mapped_detail_rows") or []
    return [
        _line_from_pipeline_row(
            row, order_number, f"附件：{filename}", f"识别明细第 {index} 行", index,
            customer_mappings=customer_mappings,
        )
        for index, row in enumerate(rows, start=1)
    ]


def _rows_from_word(path: Path, filename: str, customer_mappings: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    """Convert Word attachments locally, then use the same PDF/OCR pipeline."""
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
    rows: list[dict[str, Any]] = []
    customer_mappings = get_enabled_extraction_maps(case.get("customer_id"))
    for attachment in case.get("attachments") or []:
        if attachment.get("is_inline"):
            continue
        try:
            path = resolve_attachment_path(str(attachment.get("stored_path") or ""))
        except FileNotFoundError:
            continue
        if not path.is_file():
            continue
        filename, suffix = str(attachment.get("filename") or path.name), path.suffix.lower()
        if suffix in {".xlsx", ".xlsm", ".xls"}:
            rows.extend(_rows_from_excel(path, filename, customer_mappings=customer_mappings))
        elif suffix in {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}:
            rows.extend(_rows_from_pdf_or_image(path, filename, customer_mappings=customer_mappings))
        elif suffix in {".doc", ".docx"}:
            rows.extend(_rows_from_word(path, filename, customer_mappings=customer_mappings))
    return rows


def _line_signature(entry: dict[str, Any]) -> tuple[str, ...]:
    source_identity = _compact_key(entry.get("_source_identity"))
    if source_identity:
        return ("source", source_identity)
    values = entry["values"]
    return tuple(_compact_key(values.get(field)) for field in ("customer_order_number", "customer_product_code", "customer_spec", "quantity"))


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
        entry["values"]["line_no"] = str(len(result) + 1)
        result.append(entry)
    return result


def _initial_lines(case: dict[str, Any]) -> list[dict[str, Any]]:
    rows = _split_body_order_rows(str(case.get("body_text") or ""))
    rows.extend(_attachment_rows(case))
    rows = _merge_initial_rows(rows)
    if rows:
        return rows
    fields = case.get("detected_fields") or {}
    specs = fields.get("specs") or []
    return [
        {
            "values": {
                **_blank_line(index + 1),
                "customer_spec": str(spec),
                "customer_order_number": str(fields.get("order_number") or ""),
            },
            "sources": {"customer_spec": {"label": "邮件正文", "reference": "自动识别"}},
        }
        for index, spec in enumerate(specs)
    ] or [{"values": _blank_line(1), "sources": {}}]


def _serialize_template(conn, template_id: int) -> dict[str, Any]:
    template = conn.execute("SELECT * FROM order_entry_templates WHERE id=?", (template_id,)).fetchone()
    if not template:
        raise ValueError("录单模板不存在")
    rows = conn.execute(
        "SELECT * FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no,id", (template_id,)
    ).fetchall()
    return {
        **dict(template),
        "header": _json(template["header_json"], {}),
        "lines": [
            {"id": row["id"], "line_no": row["line_no"], "values": _json(row["values_json"], {}), "sources": _json(row["sources_json"], {})}
            for row in rows
        ],
    }


def get_or_create_template(case_id: int, employee_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
    case = _case_for_template(case_id, employee_id)
    with db_cursor() as conn:
        existing = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        if existing:
            return case, _serialize_template(conn, int(existing["id"]))
        now = utcnow()
        cursor = conn.execute(
            "INSERT INTO order_entry_templates(case_id,employee_id,created_at,updated_at) VALUES (?,?,?,?)",
            (case_id, employee_id, now, now),
        )
        template_id = int(cursor.lastrowid)
        for entry in _initial_lines(case):
            values = entry["values"]
            conn.execute(
                "INSERT INTO order_entry_template_lines(template_id,line_no,values_json,sources_json,created_at,updated_at) VALUES (?,?,?,?,?,?)",
                (template_id, int(values["line_no"] or 0), json.dumps(values, ensure_ascii=False), json.dumps(entry["sources"], ensure_ascii=False), now, now),
            )
        return case, _serialize_template(conn, template_id)


def reextract_template(case_id: int, employee_id: str) -> dict[str, Any]:
    """Rebuild one saved template's detail rows with the current extraction rules.

    The customer/header section is business-maintained and is intentionally left
    untouched.  Before replacing the current detail rows, both the previous
    contents and the regenerated contents are stored as immutable versions so a
    bulk re-extraction never discards a recoverable copy.
    """
    case = _case_for_template(case_id, employee_id)
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id,current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
        if not row:
            raise ValueError("请先打开录单模板")
        template_id = int(row["id"])
        current_version = int(row["current_version"] or 0)
        previous = _serialize_template(conn, template_id)

    # Recognition can involve OCR and file conversion, so do it outside of the
    # database transaction.  It only reads the original mail and attachments.
    regenerated_lines = _initial_lines(case)
    now = utcnow()
    previous_lines = [
        {"values": line.get("values") or {}, "sources": line.get("sources") or {}}
        for line in previous.get("lines") or []
    ]
    previous_header = previous.get("header") or {}
    backup_version = current_version + 1
    regenerated_version = backup_version + 1

    with db_cursor() as conn:
        # Record an explicit pre-run snapshot even when an older manual-save
        # version exists.  This protects the exact database state the user
        # asked to reprocess.
        conn.execute(
            "INSERT INTO order_entry_template_versions(template_id,version_number,header_json,lines_json,saved_by,saved_at) VALUES (?,?,?,?,?,?)",
            (
                template_id,
                backup_version,
                json.dumps(previous_header, ensure_ascii=False),
                json.dumps(previous_lines, ensure_ascii=False),
                employee_id,
                now,
            ),
        )
        conn.execute(
            "INSERT INTO order_entry_template_versions(template_id,version_number,header_json,lines_json,saved_by,saved_at) VALUES (?,?,?,?,?,?)",
            (
                template_id,
                regenerated_version,
                json.dumps(previous_header, ensure_ascii=False),
                json.dumps(regenerated_lines, ensure_ascii=False),
                employee_id,
                now,
            ),
        )
        conn.execute(
            "UPDATE order_entry_templates SET current_version=?,updated_at=? WHERE id=?",
            (regenerated_version, now, template_id),
        )
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
        template = _serialize_template(conn, template_id)

    return {
        "case_id": case_id,
        "subject": str(case.get("subject") or ""),
        "previous_line_count": len(previous_lines),
        "line_count": len(template.get("lines") or []),
        "backup_version": backup_version,
        "current_version": regenerated_version,
        "template": template,
    }


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
    """Return read-only workflow facts; no business user maintains this state."""
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
    version = int(row["current_version"] or 0) if row else 0
    if not row:
        return {
            "created": False, "saved": False, "version": 0,
            "stage": "pending_extraction", "label": "待提取订单",
            "next_action": "提取订单到内销模板", "step": 2,
        }
    if version <= 0:
        return {
            "created": True, "saved": False, "version": 0,
            "stage": "pending_template_save", "label": "待保存模板",
            "next_action": "核对并保存内销模板", "step": 3,
        }
    return {
        "created": True, "saved": True, "version": version,
        "stage": "pending_interface_submit", "label": "待接口提交",
        "next_action": "等待接口接入后提交", "step": 4,
    }


def _clean_values(values: dict[str, Any], fields: tuple[str, ...]) -> dict[str, str]:
    return {field: str(values.get(field) or "").strip() for field in fields}


def save_template(case_id: int, employee_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    _case_for_template(case_id, employee_id)
    header = _clean_values(payload.get("header") or {}, HEADER_FIELDS)
    raw_lines = payload.get("lines") or []
    if not isinstance(raw_lines, list):
        raise ValueError("明细行格式无效")
    lines: list[dict[str, Any]] = []
    for index, raw in enumerate(raw_lines, start=1):
        values = _clean_values((raw or {}).get("values") or raw or {}, LINE_FIELDS)
        if not any(values[field] for field in LINE_FIELDS if field != "line_no"):
            continue
        values["line_no"] = str(index)
        sources = (raw or {}).get("sources") or {}
        lines.append({"values": values, "sources": sources if isinstance(sources, dict) else {}})
    if not lines:
        lines = [{"values": _blank_line(1), "sources": {}}]
    now = utcnow()
    with db_cursor() as conn:
        template = conn.execute(
            "SELECT id,current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        if not template:
            raise ValueError("请先打开录单模板")
        template_id = int(template["id"])
        next_version = int(template["current_version"] or 0) + 1
        conn.execute("UPDATE order_entry_templates SET header_json=?,current_version=?,updated_at=? WHERE id=?", (json.dumps(header, ensure_ascii=False), next_version, now, template_id))
        conn.execute("DELETE FROM order_entry_template_lines WHERE template_id=?", (template_id,))
        for entry in lines:
            values = entry["values"]
            conn.execute(
                "INSERT INTO order_entry_template_lines(template_id,line_no,values_json,sources_json,created_at,updated_at) VALUES (?,?,?,?,?,?)",
                (template_id, int(values["line_no"]), json.dumps(values, ensure_ascii=False), json.dumps(entry["sources"], ensure_ascii=False), now, now),
            )
        conn.execute(
            "INSERT INTO order_entry_template_versions(template_id,version_number,header_json,lines_json,saved_by,saved_at) VALUES (?,?,?,?,?,?)",
            (template_id, next_version, json.dumps(header, ensure_ascii=False), json.dumps(lines, ensure_ascii=False), employee_id, now),
        )
        return _serialize_template(conn, template_id)


def validation_issues(template: dict[str, Any]) -> list[str]:
    issues = [f"{HEADER_LABELS[field]}未填写" for field in REQUIRED_HEADER_FIELDS if not str(template["header"].get(field) or "").strip()]
    for line in template.get("lines") or []:
        for field in REQUIRED_LINE_FIELDS:
            if not str((line.get("values") or {}).get(field) or "").strip():
                issues.append(f"第 {line.get('line_no')} 行{LINE_LABELS[field]}未填写")
    return issues


def build_domestic_export(case_id: int, employee_id: str) -> tuple[BytesIO, str]:
    _case_for_template(case_id, employee_id)
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?", (case_id, employee_id)
        ).fetchone()
        if not row:
            raise ValueError("请先保存模板内容后再下载")
        template = _serialize_template(conn, int(row["id"]))
    if int(template.get("current_version") or 0) <= 0:
        raise ValueError("请先保存模板内容后再下载")
    if not DOMESTIC_TEMPLATE_PATH.is_file():
        raise ValueError("内销录单模板文件未配置")
    book = load_workbook(DOMESTIC_TEMPLATE_PATH)
    sheet = book["内销"]
    for index, field in enumerate(HEADER_FIELDS, start=1):
        sheet.cell(2, index).value = template["header"].get(field) or None
    required_rows = 3 + max(1, len(template["lines"]))
    style_source_row = 4
    while sheet.max_row < required_rows:
        target = sheet.max_row + 1
        for col in range(1, 15):
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
    return data, f"内销录单_邮件{case_id}_v{template['current_version']}.xlsx"
