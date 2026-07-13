from __future__ import annotations

import re
import tempfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from .docling_parser import parse_pdf_with_docling
from .image_preprocess import preprocess_page_image
from .page_image_renderer import render_input_pages
from .pdf_native_parser import parse_pdf_native
from .purchase_field_rules import clean_text, classify_section_line, extract_key_values, find_detail_header_row, looks_like_detail_data, normalize_date, normalize_number
from .purchase_layout_cache import load_layout_cache, save_layout_cache
from .purchase_order_segmenter import _cached_ocr_regions, build_detail_rows_from_table, segment_purchase_page, segment_purchase_page_with_layout
from .purchase_result_normalizer import normalize_purchase_document
from .template_parser import identify_template


def _file_type(path: Path) -> str:
    return "pdf" if path.suffix.lower() == ".pdf" else "image"


def _native_pdf_text(file_item: dict[str, str]) -> str:
    if Path(file_item["stored_path"]).suffix.lower() != ".pdf":
        return ""
    try:
        native = parse_pdf_native(Path(file_item["stored_path"]))
    except Exception:
        return ""
    return clean_text(native.get("text"))


def _merge_sections(pages: list[dict[str, Any]]) -> dict[str, list[str]]:
    sections = {"备注": [], "条款": [], "付款信息": [], "收货信息": [], "签核区": []}
    seen: set[tuple[str, str]] = set()
    for page in pages:
        for section, lines in (page.get("sections") or {}).items():
            for line in lines:
                text = clean_text(line)
                key = (section, text)
                if text and key not in seen:
                    seen.add(key)
                    sections.setdefault(section, []).append(text)
    return sections


def _collect_lines(pages: list[dict[str, Any]], auxiliary_text: str) -> list[str]:
    lines: list[str] = []
    for page in pages:
        lines.extend(page.get("text_lines") or [])
    if auxiliary_text:
        lines.extend(line.strip() for line in auxiliary_text.splitlines() if line.strip())
    result: list[str] = []
    seen: set[str] = set()
    for line in lines:
        text = clean_text(line)
        key = text.lower().replace(" ", "")
        if text and key not in seen:
            seen.add(key)
            result.append(text)
    return result


def _quick_page_lines(page: dict[str, Any]) -> list[str]:
    image_path = Path(page.get("clean_image_path") or page["image_path"])
    try:
        regions = _cached_ocr_regions(image_path)
    except Exception:
        return []
    regions.sort(key=lambda item: ((item.get("bbox") or [0, 0, 0, 0])[1], (item.get("bbox") or [0, 0, 0, 0])[0]))
    return [clean_text(region.get("text")) for region in regions if clean_text(region.get("text"))]


def _detail_tables_from_pages(pages: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    raw_detail_tables: list[dict[str, Any]] = []
    mapped_rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    for page in pages:
        issues.extend(page.get("issues") or [])
        for table in page.get("tables") or []:
            if table.get("table_type") != "detail_table":
                continue
            raw_detail_tables.append(
                {
                    "page_index": table.get("page_index", 0),
                    "table_index": table.get("table_index", len(raw_detail_tables)),
                    "bbox": table.get("bbox"),
                    "rows": table.get("raw_rows") or [],
                    "method": table.get("method") or "grid_cell_ocr",
                    "confidence": table.get("confidence", 0),
                }
            )
            rows, table_issues = build_detail_rows_from_table(table)
            mapped_rows.extend(rows)
            issues.extend(table_issues)
    return raw_detail_tables, mapped_rows, issues


def _matrix_from_native_cells(cells: list[dict[str, Any]]) -> list[list[str]]:
    if not cells:
        return []
    row_indexes = sorted({int(cell.get("row_index") or 0) for cell in cells})
    column_indexes = sorted({int(cell.get("column_index") or 0) for cell in cells})
    rows: list[list[str]] = []
    for row_index in row_indexes:
        row: list[str] = []
        for column_index in column_indexes:
            value = next(
                (
                    clean_text(cell.get("text"))
                    for cell in cells
                    if int(cell.get("row_index") or 0) == row_index and int(cell.get("column_index") or 0) == column_index
                ),
                "",
            )
            row.append(value)
        if any(row):
            rows.append(row)
    while rows and sum(1 for value in rows[-1] if clean_text(value)) <= 1:
        rows.pop()
    return rows


def _valid_detail_standard(row: dict[str, Any]) -> bool:
    standard = row.get("standard") or {}
    return bool(
        clean_text(standard.get("物料编码"))
        and clean_text(standard.get("物料名称"))
        and clean_text(standard.get("数量"))
        and clean_text(standard.get("单位"))
        and clean_text(standard.get("交货日期"))
    )


def _decimal_value(value: Any) -> Decimal | None:
    number = normalize_number(value)
    if not number:
        return None
    try:
        return Decimal(number)
    except InvalidOperation:
        return None


def _date_value(value: Any) -> str:
    return normalize_date(value) or clean_text(value).replace("/", "-")


def _make_fast_row(
    *,
    page_index: int,
    original_headers: list[str],
    original_values: list[str],
    standard: dict[str, Any],
    method: str,
) -> dict[str, Any]:
    normalized = {
        "序号": clean_text(standard.get("序号")),
        "物料编码": clean_text(standard.get("物料编码")),
        "物料名称": clean_text(standard.get("物料名称")),
        "说明": clean_text(standard.get("说明")),
        "数量": normalize_number(standard.get("数量")),
        "单位": clean_text(standard.get("单位")),
        "含税单价": normalize_number(standard.get("含税单价")),
        "金额": normalize_number(standard.get("金额")),
        "交货日期": _date_value(standard.get("交货日期")),
        "备注": clean_text(standard.get("备注")),
    }
    return {
        "page_index": page_index,
        "row_index": int(normalized.get("序号") or 0) if str(normalized.get("序号") or "").isdigit() else 0,
        "method": method,
        "confidence": 0.96,
        "original": {header: clean_text(original_values[index] if index < len(original_values) else "") for index, header in enumerate(original_headers)},
        "standard": normalized,
        "cleaning_notes": [],
    }


def _amount_is_consistent(row: dict[str, Any]) -> bool:
    standard = row.get("standard") or {}
    quantity = _decimal_value(standard.get("数量"))
    price = _decimal_value(standard.get("含税单价"))
    amount = _decimal_value(standard.get("金额"))
    if quantity is None or price is None or amount is None:
        return True
    return abs(quantity * price - amount) <= Decimal("0.05")


def _fast_rows_are_reliable(rows: list[dict[str, Any]], *, min_rows: int = 2) -> bool:
    valid_rows = [row for row in rows if _valid_detail_standard(row)]
    if len(valid_rows) < min_rows or len(valid_rows) != len(rows):
        return False
    return all(_amount_is_consistent(row) for row in valid_rows)


def _native_lines(native: dict[str, Any]) -> list[tuple[int, str]]:
    lines: list[tuple[int, str]] = []
    for page in native.get("pages") or []:
        page_index = int(page.get("page_index") or 0)
        for raw_line in str(page.get("text") or "").splitlines():
            line = clean_text(raw_line)
            if line:
                lines.append((page_index, line))
    return lines


def _append_text_part(parts: list[str], text: str) -> None:
    text = clean_text(text)
    if not text:
        return
    if text in {"备注:", "备注："}:
        return
    if any(keyword in text for keyword in ["总金额", "合计", "大写金额", "供方必须", "一、", "二、", "三、"]):
        return
    if text not in parts:
        parts.append(text)


def _parse_jingwang_native_lines(lines: list[tuple[int, str]]) -> tuple[list[dict[str, Any]], list[list[str]]]:
    headers = ["序号", "物料编码", "物料名称", "说明", "数量", "单位", "含税单价", "金额", "交货日期", "备注"]
    pattern = re.compile(
        r"^(\d+(?:\.\d+)?)\s+(\d{6,})\s+(.+?)\s+"
        r"(\d+(?:,\d{3})*(?:\.\d+)?)\s+([\u4e00-\u9fffA-Za-z]+)\s+"
        r"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+"
        r"(20\d{2}[-/]\d{1,2}[-/]\d{1,2})(?:\s+(.*))?$"
    )
    rows: list[dict[str, Any]] = []
    raw_rows = [headers]
    pending_remark: list[str] = []
    current: dict[str, Any] | None = None
    current_values: list[str] = []
    in_detail = False

    def flush() -> None:
        nonlocal current, current_values
        if not current:
            return
        remark = clean_text(" ".join(current.pop("_remark_parts", [])))
        current["说明"] = remark
        row = _make_fast_row(
            page_index=int(current.pop("_page_index", 0)),
            original_headers=headers,
            original_values=current_values,
            standard=current,
            method="pdf_text_line_fast",
        )
        rows.append(row)
        raw_rows.append([row["standard"].get(header, "") for header in headers])
        current = None
        current_values = []

    for page_index, line in lines:
        if "Item No" in line and "Part No" in line:
            in_detail = True
            continue
        if not in_detail:
            continue
        if any(keyword in line for keyword in ["总金额", "备注：", "供方必须"]):
            flush()
            break
        match = pattern.match(line)
        if match:
            flush()
            current = {
                "_page_index": page_index,
                "_remark_parts": pending_remark[:],
                "序号": match.group(1),
                "物料编码": match.group(2),
                "物料名称": match.group(3),
                "数量": match.group(4),
                "单位": match.group(5),
                "含税单价": match.group(6),
                "金额": match.group(7),
                "交货日期": match.group(8),
                "备注": clean_text(match.group(9) or ""),
            }
            current_values = [
                current["序号"],
                current["物料编码"],
                current["物料名称"],
                "",
                current["数量"],
                current["单位"],
                current["含税单价"],
                current["金额"],
                current["交货日期"],
                current["备注"],
            ]
            pending_remark = []
            continue
        if current:
            _append_text_part(current["_remark_parts"], line)
        else:
            _append_text_part(pending_remark, line)
    flush()
    return rows, raw_rows


def _parse_talian_native_lines(lines: list[tuple[int, str]]) -> tuple[list[dict[str, Any]], list[list[str]]]:
    headers = ["No.", "原料编码", "原料名称", "规格", "数量", "单位", "单价RMB", "税率%", "金额", "交期", "备注"]
    pattern = re.compile(
        r"^(\d+)\s+([A-Za-z0-9_-]+)\s+"
        r"(\d+(?:,\d{3})*(?:\.\d+)?)\s*([\u4e00-\u9fffA-Za-z]+)\s+"
        r"(\d+(?:\.\d+)?)\s+(\d{1,2}(?:\.\d+)?)\s+"
        r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+"
        r"(20\d{2}[-/]\d{1,2}[-/]\d{1,2})(?:\s+(.*))?$"
    )
    rows: list[dict[str, Any]] = []
    raw_rows = [headers]
    pending_spec: list[str] = []
    current: dict[str, Any] | None = None
    after_detail_consumed = False
    in_detail = False

    def flush() -> None:
        nonlocal current, after_detail_consumed
        if not current:
            return
        specs = current.pop("_spec_parts", [])
        name = clean_text(" ".join(specs))
        current["物料名称"] = name
        values = [
            current["序号"],
            current["物料编码"],
            name,
            "",
            current["数量"],
            current["单位"],
            current["含税单价"],
            current.pop("_tax_rate", ""),
            current["金额"],
            current["交货日期"],
            current["备注"],
        ]
        row = _make_fast_row(
            page_index=int(current.pop("_page_index", 0)),
            original_headers=headers,
            original_values=values,
            standard=current,
            method="pdf_text_line_fast",
        )
        rows.append(row)
        raw_rows.append(values)
        current = None
        after_detail_consumed = False

    for page_index, line in lines:
        if "No." in line and "原料编码" in line and "单价" in line:
            in_detail = True
            continue
        if not in_detail:
            continue
        if any(keyword in line for keyword in ["合计", "总计", "二、", "三、", "备注："]):
            flush()
            break
        match = pattern.match(line)
        if match:
            flush()
            current = {
                "_page_index": page_index,
                "_spec_parts": pending_spec[:],
                "序号": match.group(1),
                "物料编码": match.group(2),
                "物料名称": "",
                "说明": "",
                "数量": match.group(3),
                "单位": match.group(4),
                "含税单价": match.group(5),
                "_tax_rate": match.group(6),
                "金额": match.group(7),
                "交货日期": match.group(8),
                "备注": clean_text(match.group(9) or ""),
            }
            after_detail_consumed = False
            pending_spec = []
            continue
        if current and not after_detail_consumed:
            _append_text_part(current["_spec_parts"], line)
            after_detail_consumed = True
        else:
            _append_text_part(pending_spec, line)
    flush()
    return rows, raw_rows


def _parse_chaoyue_native_lines(lines: list[tuple[int, str]]) -> tuple[list[dict[str, Any]], list[list[str]]]:
    headers = ["项次", "料件编号", "名称规格", "数量", "单位", "单价", "金额", "交货日期", "备注"]
    pattern = re.compile(
        r"^(\d+)\s+([A-Za-z0-9_-]{8,})\s+"
        r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s*([\u4e00-\u9fffA-Za-z]+)\s+"
        r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+"
        r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+"
        r"(20\d{2}[-/]\d{1,2}[-/]\d{1,2})(?:\s+(.*))?$"
    )
    rows: list[dict[str, Any]] = []
    raw_rows = [headers]
    current: dict[str, Any] | None = None

    def flush() -> None:
        nonlocal current
        if not current:
            return
        name = clean_text(" ".join(current.pop("_name_parts", [])))
        current["物料名称"] = name
        values = [
            current["序号"],
            current["物料编码"],
            name,
            current["数量"],
            current["单位"],
            current["含税单价"],
            current["金额"],
            current["交货日期"],
            current["备注"],
        ]
        row = _make_fast_row(
            page_index=int(current.pop("_page_index", 0)),
            original_headers=headers,
            original_values=values,
            standard=current,
            method="pdf_text_line_fast",
        )
        rows.append(row)
        raw_rows.append(values)
        current = None

    for page_index, line in lines:
        if any(keyword in line for keyword in ["税前总金额", "含税总金额", "合计：", "人民币金额"]):
            flush()
            continue
        match = pattern.match(line)
        if match:
            flush()
            current = {
                "_page_index": page_index,
                "_name_parts": [],
                "序号": match.group(1),
                "物料编码": match.group(2),
                "物料名称": "",
                "说明": "",
                "数量": match.group(3),
                "单位": match.group(4),
                "含税单价": match.group(5),
                "金额": match.group(6),
                "交货日期": match.group(7),
                "备注": clean_text(match.group(8) or ""),
            }
            continue
        if current:
            _append_text_part(current["_name_parts"], line)
    flush()
    return rows, raw_rows


def _native_text_purchase_document(file_item: dict[str, str], native: dict[str, Any]) -> dict[str, Any] | None:
    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    quality = native.get("text_quality") or {}
    if not quality.get("has_text") or not quality.get("readable"):
        return None

    text = clean_text(native.get("text"))
    lines = _native_lines(native)
    if "P.ONO" in text and "Part No" in text:
        mapped_rows, raw_rows = _parse_jingwang_native_lines(lines)
    elif "深圳市塔联科技有限公司" in text and "合同编号" in text and "原料编码" in text:
        mapped_rows, raw_rows = _parse_talian_native_lines(lines)
    elif "赣州市超跃科技" in text and "采购单号" in text and "料件编号" in text:
        mapped_rows, raw_rows = _parse_chaoyue_native_lines(lines)
    else:
        return None

    if not _fast_rows_are_reliable(mapped_rows):
        return None

    text_lines = [line for _page_index, line in lines]
    header_info = extract_key_values(text_lines)
    template = identify_template(source_file, "\n".join(text_lines))
    sections = {"备注": [], "条款": [], "付款信息": [], "收货信息": [], "签核区": []}
    for line in text_lines:
        section = classify_section_line(line)
        if section:
            sections.setdefault(section, []).append(line)

    warnings = list(native.get("warnings") or [])
    warnings.append("PDF 使用 pdfplumber 原生文本快速重建明细，未调用 Docling、未渲染页面图片、未调用 OCR。")
    return {
        "pipeline_version": "purchase_order_v1",
        "source_file": source_file,
        "file_type": "pdf",
        "parser_mode": "template_pdf_text_line_fast" if template else "pdf_text_line_fast",
        "template_id": template.template_id if template else "",
        "template_label": template.label if template else "",
        "page_count": int(native.get("page_count") or 1),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "header_info": header_info,
        "pages": [],
        "regions": [
            {
                "page_index": 0,
                "tables": [
                    {
                        "table_index": 0,
                        "table_type": "detail_table",
                        "bbox": [],
                        "row_count": max(0, len(raw_rows) - 1),
                        "method": "pdf_text_line_fast",
                    }
                ],
            }
        ],
        "raw_detail_tables": [
            {
                "page_index": 0,
                "table_index": 0,
                "bbox": [],
                "rows": raw_rows,
                "method": "pdf_text_line_fast",
                "confidence": 0.96,
            }
        ],
        "mapped_detail_rows": mapped_rows,
        "sections": sections,
        "issues": [],
        "warnings": warnings,
    }


def _native_purchase_document(file_item: dict[str, str], native: dict[str, Any]) -> dict[str, Any] | None:
    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    quality = native.get("text_quality") or {}
    if not quality.get("has_text") or not quality.get("readable"):
        return None

    raw_detail_tables: list[dict[str, Any]] = []
    mapped_rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    for page in native.get("pages") or []:
        for table in page.get("tables") or []:
            rows = _matrix_from_native_cells(table.get("cells") or [])
            if len(rows) < 3:
                continue
            table_doc = {
                "page_index": page.get("page_index", 0),
                "table_index": table.get("table_index", len(raw_detail_tables)),
                "table_type": "detail_table",
                "raw_rows": rows,
                "rows": rows,
                "bbox": [],
                "method": "pdf_text",
                "confidence": 1.0,
            }
            table_rows, table_issues = build_detail_rows_from_table(table_doc)
            valid_rows = [row for row in table_rows if _valid_detail_standard(row)]
            if table_issues or len(valid_rows) < 2 or len(valid_rows) != len(table_rows):
                continue
            raw_detail_tables.append(
                {
                    "page_index": page.get("page_index", 0),
                    "table_index": len(raw_detail_tables),
                    "bbox": [],
                    "rows": rows,
                    "method": "pdf_text_fast",
                    "confidence": 1.0,
                }
            )
            for row in valid_rows:
                row["method"] = "pdf_text_fast"
            mapped_rows.extend(valid_rows)
            issues.extend(table_issues)

    if len(mapped_rows) < 2:
        return None

    lines = [line.strip() for line in clean_text(native.get("text")).splitlines() if line.strip()]
    header_info = extract_key_values(lines)
    template = identify_template(source_file, "\n".join(lines))
    sections = {"备注": [], "条款": [], "付款信息": [], "收货信息": [], "签核区": []}
    for line in lines:
        section = classify_section_line(line)
        if section:
            sections.setdefault(section, []).append(line)

    warnings = list(native.get("warnings") or [])
    warnings.append("PDF 使用 pdfplumber 原生文字快路径解析，未调用 Docling、未渲染页面图片、未调用 OCR。")
    return {
        "pipeline_version": "purchase_order_v1",
        "source_file": source_file,
        "file_type": "pdf",
        "parser_mode": "template_pdf_text_fast" if template else "pdf_text_fast",
        "template_id": template.template_id if template else "",
        "template_label": template.label if template else "",
        "page_count": int(native.get("page_count") or 1),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "header_info": header_info,
        "pages": [],
        "regions": [
            {
                "page_index": table.get("page_index"),
                "tables": [
                    {
                        "table_index": table.get("table_index"),
                        "table_type": "detail_table",
                        "bbox": table.get("bbox"),
                        "row_count": len(table.get("rows") or []),
                        "method": table.get("method"),
                    }
                ],
            }
            for table in raw_detail_tables
        ],
        "raw_detail_tables": raw_detail_tables,
        "mapped_detail_rows": mapped_rows,
        "sections": sections,
        "issues": issues,
        "warnings": warnings,
    }


def _native_pdf_summary(path: Path) -> dict[str, Any]:
    try:
        return parse_pdf_native(path)
    except Exception:
        return {"text": "", "page_count": 0, "pages": [], "warnings": ["pdfplumber 轻量解析失败。"]}


def _docling_table_type(rows: list[list[str]]) -> str:
    header_index, _mapping = find_detail_header_row(rows)
    if header_index is not None:
        return "detail_table"
    data_rows = sum(1 for row in rows if looks_like_detail_data(row))
    if data_rows >= 2:
        return "detail_table"
    text = "\n".join(" ".join(clean_text(value) for value in row) for row in rows)
    hits = [classify_section_line(line) for line in text.splitlines()]
    if any(hit in {"付款信息", "收货信息"} for hit in hits):
        return "payment_shipping"
    if any(hit in {"备注", "条款", "签核区"} for hit in hits):
        return "terms_notes"
    if any(keyword in text for keyword in ["订单号", "采购单号", "供应商", "客户", "合同编号", "P.O", "Vendor"]):
        return "order_header"
    return "unknown"


def _sections_from_docling(lines: list[str], tables: list[dict[str, Any]]) -> dict[str, list[str]]:
    sections = {"备注": [], "条款": [], "付款信息": [], "收货信息": [], "签核区": []}
    seen: set[tuple[str, str]] = set()

    def add(section: str, text: str) -> None:
        text = clean_text(text)
        key = (section, text)
        if text and key not in seen:
            seen.add(key)
            sections.setdefault(section, []).append(text)

    for line in lines:
        section = classify_section_line(line)
        if section:
            add(section, line)

    for table in tables:
        table_type = table.get("table_type")
        if table_type == "detail_table":
            continue
        table_lines = [" ".join(clean_text(value) for value in row if clean_text(value)) for row in table.get("rows") or []]
        for line in table_lines:
            section = classify_section_line(line)
            if section:
                add(section, line)
            elif table_type == "payment_shipping":
                add("付款信息", line)
            elif table_type == "terms_notes":
                add("条款", line)
            elif table_type == "unknown":
                add("备注", f"Docling 未分类内容：{line}")
    return sections


def _docling_purchase_document(file_item: dict[str, str], native: dict[str, Any]) -> dict[str, Any] | None:
    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    docling_result = parse_pdf_with_docling(input_path)
    tables: list[dict[str, Any]] = []
    for table_index, table in enumerate(docling_result.get("tables") or []):
        rows = table.get("rows") or []
        if not rows:
            continue
        table_type = _docling_table_type(rows)
        tables.append(
            {
                "page_index": 0,
                "table_index": table_index,
                "table_type": table_type,
                "title": table.get("title") or f"Docling 表 {table_index + 1}",
                "rows": rows,
                "raw_rows": rows,
                "bbox": [],
                "method": "docling_markdown",
                "confidence": 0.92,
            }
        )

    detail_tables = [table for table in tables if table.get("table_type") == "detail_table"]
    if not detail_tables:
        return None

    native_text = clean_text(native.get("text"))
    lines = [clean_text(line) for line in docling_result.get("lines") or [] if clean_text(line)]
    if native_text:
        lines.extend(line.strip() for line in native_text.splitlines() if line.strip())
    header_info = extract_key_values(lines)
    template = identify_template(source_file, "\n".join(lines))
    sections = _sections_from_docling(lines, tables)

    raw_detail_tables: list[dict[str, Any]] = []
    mapped_rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    for table in detail_tables:
        raw_detail_tables.append(
            {
                "page_index": table.get("page_index", 0),
                "table_index": len(raw_detail_tables),
                "bbox": [],
                "rows": table.get("rows") or [],
                "method": "docling_markdown",
                "confidence": table.get("confidence", 0.92),
                "title": table.get("title", ""),
            }
        )
        rows, table_issues = build_detail_rows_from_table(table)
        for row in rows:
            row["method"] = "docling_markdown"
        mapped_rows.extend(rows)
        issues.extend(table_issues)

    parser_mode = "template_docling_markdown_pdf" if template else "docling_markdown_pdf"
    warnings = list(native.get("warnings") or [])
    if docling_result.get("error"):
        warnings.append(f"Docling 解析提示：{docling_result.get('error')}")
    warnings.append("PDF 使用 Docling/Markdown 快速分层解析，未渲染页面图片，未调用 OCR。")
    return {
        "pipeline_version": "purchase_order_v1",
        "source_file": source_file,
        "file_type": "pdf",
        "parser_mode": parser_mode,
        "template_id": template.template_id if template else "",
        "template_label": template.label if template else "",
        "page_count": int(native.get("page_count") or 1),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "header_info": header_info,
        "pages": [],
        "regions": [
            {
                "page_index": 0,
                "tables": [
                    {
                        "table_index": table.get("table_index"),
                        "table_type": table.get("table_type"),
                        "bbox": table.get("bbox"),
                        "row_count": len(table.get("rows") or []),
                        "method": "docling_markdown",
                    }
                    for table in tables
                ],
            }
        ],
        "raw_detail_tables": raw_detail_tables,
        "mapped_detail_rows": mapped_rows,
        "sections": sections,
        "issues": issues,
        "warnings": warnings,
    }


def run_purchase_order_pipeline(file_item: dict[str, str], work_dir: Path | None = None) -> dict[str, Any]:
    """Generic purchase-order recognizer.

    This is intentionally template-optional. Text PDFs use Docling/Markdown as
    the fast layered path; images and scanned PDFs fall back to page images,
    table-grid detection and cell OCR.
    """
    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    cleanup: tempfile.TemporaryDirectory[str] | None = None
    if work_dir is None:
        cleanup = tempfile.TemporaryDirectory()
        base_dir = Path(cleanup.name)
    else:
        base_dir = Path(work_dir)
        base_dir.mkdir(parents=True, exist_ok=True)

    try:
        if input_path.suffix.lower() == ".pdf":
            native = _native_pdf_summary(input_path)
            native_text = str(native.get("text") or "")
            native_text_document = _native_text_purchase_document(file_item, native)
            if native_text_document:
                return normalize_purchase_document(native_text_document, source_text=native_text)
            native_document = _native_purchase_document(file_item, native)
            if native_document:
                return normalize_purchase_document(native_document, source_text=native_text)
            docling_document = _docling_purchase_document(file_item, native)
            if docling_document:
                return normalize_purchase_document(docling_document, source_text=native_text)

        render_dir = base_dir / "pages"
        clean_dir = base_dir / "clean"
        rendered_pages = render_input_pages(file_item, render_dir)
        if not rendered_pages:
            raise RuntimeError("未能渲染页面图片。")

        pages: list[dict[str, Any]] = []
        clean_pages = [preprocess_page_image(rendered_page, clean_dir) for rendered_page in rendered_pages]
        layout_cache = None
        if len(clean_pages) == 1:
            quick_lines = _quick_page_lines(clean_pages[0])
            probe_lines = []
            seen_probe_lines: set[str] = set()
            for line in quick_lines:
                key = clean_text(line).lower().replace(" ", "")
                if key and key not in seen_probe_lines:
                    seen_probe_lines.add(key)
                    probe_lines.append(clean_text(line))
            auxiliary_text = _native_pdf_text(file_item)
            if auxiliary_text:
                for line in auxiliary_text.splitlines():
                    text = clean_text(line)
                    key = text.lower().replace(" ", "")
                    if text and key not in seen_probe_lines:
                        seen_probe_lines.add(key)
                        probe_lines.append(text)
            probe_header = extract_key_values(probe_lines)
            layout_cache = load_layout_cache(
                probe_header,
                probe_lines,
                int(clean_pages[0].get("width") or 0),
                int(clean_pages[0].get("height") or 0),
            )
            if layout_cache:
                pages.append(segment_purchase_page_with_layout(clean_pages[0], layout_cache))
            else:
                probe_page = segment_purchase_page(clean_pages[0])
                pages.append(probe_page)
        else:
            for clean_page in clean_pages:
                pages.append(segment_purchase_page(clean_page))

        auxiliary_text = _native_pdf_text(file_item)
        lines = _collect_lines(pages, auxiliary_text)
        header_info = extract_key_values(lines)
        template = identify_template(source_file, "\n".join(lines))
        raw_detail_tables, mapped_rows, issues = _detail_tables_from_pages(pages)
        sections = _merge_sections(pages)

        if not raw_detail_tables:
            issues.append(
                {
                    "page_index": "",
                    "region": "明细表",
                    "field": "",
                    "raw_value": "",
                    "clean_value": "",
                    "confidence": 0,
                    "message": "通用 pipeline 未恢复出明细表，建议回退旧流程或人工复核。",
                }
            )

        parser_mode = "generic_page_grid_ocr"
        if template and raw_detail_tables:
            parser_mode = "generic_page_grid_ocr_template_enhanced"

        document = {
            "pipeline_version": "purchase_order_v1",
            "source_file": source_file,
            "file_type": _file_type(input_path),
            "parser_mode": parser_mode,
            "template_id": template.template_id if template else "",
            "template_label": template.label if template else "",
            "page_count": len(rendered_pages),
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "header_info": header_info,
            "pages": pages,
            "regions": [
                {
                    "page_index": page.get("page_index"),
                    "tables": [
                        {
                            "table_index": table.get("table_index"),
                            "table_type": table.get("table_type"),
                            "bbox": table.get("bbox"),
                            "row_count": len(table.get("raw_rows") or []),
                        }
                        for table in page.get("tables") or []
                    ],
                }
                for page in pages
            ],
            "raw_detail_tables": raw_detail_tables,
            "mapped_detail_rows": mapped_rows,
            "sections": sections,
            "issues": issues,
            "warnings": [],
        }
        if layout_cache:
            document["layout_cache_hit"] = True
            document.setdefault("warnings", []).append("已命中历史版式缓存，使用缓存区域快速识别。")
        normalized_document = normalize_purchase_document(document, source_text="\n".join(lines))
        if not layout_cache and save_layout_cache(normalized_document):
            normalized_document.setdefault("warnings", []).append("已保存本次版式缓存，后续相同版式将优先快速识别。")
        return normalized_document
    finally:
        if cleanup is not None:
            cleanup.cleanup()
