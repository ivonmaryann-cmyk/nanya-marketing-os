from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from copy import copy
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .db import db_cursor, utcnow
from .excel_utils import load_workbook_compat
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
        values = _blank_line(len(result) + 1)
        values.update({
            "customer_product_code": customer_code,
            "product_name": f"南亚{chunk[material_index]}" if material_index >= 0 and "南亚" in chunk else "",
            "customer_spec": spec,
            "quantity": quantity.replace(",", ""),
            "customer_order_number": chunk[0],
        })
        sources = {
            key: {"label": "邮件正文", "reference": f"第 {start + 1} 行附近"}
            for key, value in values.items() if value and key != "line_no"
        }
        result.append({"values": values, "sources": sources})
    return result


def _source(label: str, reference: str, values: dict[str, str]) -> dict[str, dict[str, str]]:
    return {field: {"label": label, "reference": reference} for field, value in values.items() if value and field != "line_no"}


def _line_entry(values: dict[str, Any], *, label: str, reference: str, line_no: int = 1) -> dict[str, Any]:
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
    return {"values": line, "sources": _source(label, reference, line)}


def _rows_from_excel(path: Path, filename: str) -> list[dict[str, Any]]:
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
                result.append(_line_entry(values, label=f"附件：{filename}", reference=f"{sheet.title} 第 {index} 行", line_no=len(result) + 1))
        return result
    finally:
        book.close()


def _line_from_pipeline_row(row: dict[str, Any], order_number: str, label: str, reference: str, line_no: int) -> dict[str, Any]:
    original = row.get("original") or {}
    standard = row.get("standard") or {}
    values = {
        "line_no": standard.get("序号") or original.get("序号") or line_no,
        "product_code": standard.get("物料编码") or original.get("物料编码") or original.get("料件编号") or "",
        "product_name": standard.get("物料名称") or original.get("物料名称") or original.get("品名") or "",
        "customer_product_code": original.get("客户产品编号") or standard.get("物料编码") or original.get("物料编码") or "",
        "customer_spec": standard.get("说明") or standard.get("物料名称") or original.get("名称规格") or original.get("规格") or "",
        "delivery_date": standard.get("交货日期") or original.get("交货日期") or "",
        "quantity": standard.get("数量") or original.get("数量") or "",
        "unit_price": standard.get("含税单价") or original.get("含税单价") or original.get("单价") or "",
        "customer_order_number": order_number,
        "remark": standard.get("备注") or original.get("备注") or "",
    }
    return _line_entry(values, label=label, reference=reference, line_no=line_no)


def _rows_from_pdf_or_image(path: Path, filename: str) -> list[dict[str, Any]]:
    try:
        document = run_purchase_order_pipeline({"stored_path": str(path), "original_filename": filename})
    except Exception:
        # An unreadable scanned attachment must not stop opening the template;
        # business can still fill the row manually from the downloaded source.
        return []
    header = document.get("header_info") or {}
    order_number = clean_text(header.get("订单号"))
    rows = document.get("mapped_detail_rows") or []
    return [
        _line_from_pipeline_row(row, order_number, f"附件：{filename}", f"识别明细第 {index} 行", index)
        for index, row in enumerate(rows, start=1)
    ]


def _rows_from_word(path: Path, filename: str) -> list[dict[str, Any]]:
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
        return _rows_from_pdf_or_image(pdf, filename) if pdf.is_file() else []


def _attachment_rows(case: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for attachment in case.get("attachments") or []:
        if attachment.get("is_inline"):
            continue
        path = Path(str(attachment.get("stored_path") or ""))
        if not path.is_file():
            continue
        filename, suffix = str(attachment.get("filename") or path.name), path.suffix.lower()
        if suffix in {".xlsx", ".xlsm", ".xls"}:
            rows.extend(_rows_from_excel(path, filename))
        elif suffix in {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}:
            rows.extend(_rows_from_pdf_or_image(path, filename))
        elif suffix in {".doc", ".docx"}:
            rows.extend(_rows_from_word(path, filename))
    return rows


def _line_signature(entry: dict[str, Any]) -> tuple[str, str, str, str]:
    values = entry["values"]
    return tuple(_compact_key(values.get(field)) for field in ("customer_order_number", "customer_product_code", "customer_spec", "quantity"))


def _merge_initial_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str]] = set()
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


def template_progress(case_id: int, employee_id: str) -> dict[str, Any]:
    """Return read-only workflow facts; no business user maintains this state."""
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT current_version FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
    version = int(row["current_version"] or 0) if row else 0
    return {"created": bool(row), "saved": version > 0, "version": version}


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
