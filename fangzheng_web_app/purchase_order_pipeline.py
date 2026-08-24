from __future__ import annotations

import re
import tempfile
import time
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from .docling_parser import parse_pdf_with_docling
from .image_preprocess import preprocess_page_image
from .page_image_renderer import render_input_pages
from .pdf_native_parser import parse_pdf_native
from .purchase_field_rules import clean_text, classify_section_line, extract_key_values, find_detail_header_row, looks_like_detail_data, normalize_date, normalize_number, normalize_pdf_table_cell
from .purchase_layout_cache import load_layout_cache, save_layout_cache
from .purchase_order_segmenter import _cached_grids, _cached_ocr_regions, build_detail_rows_from_table, segment_purchase_page, segment_purchase_page_with_layout
from .purchase_performance import (
    activate_performance,
    add_stage_ms,
    append_fallback_reason,
    file_sha256,
    load_parser_cache,
    new_performance_summary,
    performance_stage,
    reset_performance,
    save_parser_cache,
    set_cache_state,
    set_fast_path,
)
from .purchase_result_normalizer import normalize_purchase_document
from .template_parser import identify_template


def _file_type(path: Path) -> str:
    return "pdf" if path.suffix.lower() == ".pdf" else "image"


def _native_pdf_text(file_item: dict[str, str], native: dict[str, Any] | None = None) -> str:
    if Path(file_item["stored_path"]).suffix.lower() != ".pdf":
        return ""
    if native is None:
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
                    "recovery_actions": table.get("recovery_actions") or [],
                }
            )
            rows, table_issues = build_detail_rows_from_table(table)
            mapped_rows.extend(rows)
            issues.extend(table_issues)
    return raw_detail_tables, mapped_rows, issues


def _matrix_from_native_cells(cells: list[dict[str, Any]], *, drop_trailing_sparse: bool = True) -> list[list[str]]:
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
                    normalize_pdf_table_cell(cell.get("text"))
                    for cell in cells
                    if int(cell.get("row_index") or 0) == row_index and int(cell.get("column_index") or 0) == column_index
                ),
                "",
            )
            row.append(value)
        if any(row):
            rows.append(row)
    while drop_trailing_sparse and rows and sum(1 for value in rows[-1] if clean_text(value)) <= 1:
        rows.pop()
    return rows


def _native_grid_rows_from_page(page: dict[str, Any], expected_column_count: int) -> list[list[str]]:
    """Rebuild headerless continuation rows from their visible PDF grid."""
    if expected_column_count < 2:
        return []

    vertical_groups: dict[tuple[float, float], set[float]] = {}
    for line in page.get("lines") or []:
        if line.get("orientation") != "v":
            continue
        bbox = line.get("bbox") or []
        if len(bbox) != 4:
            continue
        x0, top, x1, bottom = (float(value) for value in bbox)
        if bottom - top < 12:
            continue
        key = (round(top, 1), round(bottom, 1))
        vertical_groups.setdefault(key, set()).add(round((x0 + x1) / 2, 2))

    rows: list[list[str]] = []
    for (top, bottom), positions in sorted(vertical_groups.items()):
        boundaries = sorted(positions)
        if len(boundaries) != expected_column_count + 1:
            continue
        if boundaries[-1] - boundaries[0] < float(page.get("width") or 0) * 0.5:
            continue

        row: list[str] = []
        for left, right in zip(boundaries, boundaries[1:]):
            words = []
            for word in page.get("words") or []:
                bbox = word.get("bbox") or []
                if len(bbox) != 4:
                    continue
                x0, word_top, x1, word_bottom = (float(value) for value in bbox)
                center_x = (x0 + x1) / 2
                center_y = (word_top + word_bottom) / 2
                if left <= center_x <= right and top <= center_y <= bottom:
                    words.append((word_top, x0, clean_text(word.get("text"))))
            words.sort(key=lambda item: (item[0], item[1]))
            row.append(normalize_pdf_table_cell("\n".join(value for _, _, value in words if value)))
        if any(row):
            rows.append(row)
    return rows


def _repair_split_native_headers(rows: list[list[str]]) -> list[list[str]]:
    """Restore headers that pdfplumber splits over adjacent table columns."""
    repaired = [list(row) for row in rows]
    for row in repaired:
        for index in range(len(row) - 1):
            if clean_text(row[index]) == "单" and clean_text(row[index + 1]) == "位":
                row[index] = "单位"
                row[index + 1] = ""
    return repaired


def _compact_native_material_code_column(rows: list[list[str]], header_index: int, mapping: dict[int, str]) -> list[list[str]]:
    """PDF table cells may wrap one material code across several visual lines."""
    repaired = [list(row) for row in rows]
    for row in repaired[header_index + 1 :]:
        for column, field in mapping.items():
            if field == "物料编码" and column < len(row):
                row[column] = re.sub(r"\s+", "", clean_text(row[column]))
    return repaired


def _native_coordinate_rows(rows: list[list[str]]) -> tuple[list[list[str]], str]:
    """Collapse coordinate-only continuation rows without guessing their owner."""
    header_index, mapping = find_detail_header_row(rows)
    if header_index is None:
        return rows, "未识别到可靠明细表头"
    reverse_mapping = {field: column for column, field in mapping.items()}
    required_columns = ["序号", "物料编码", "物料名称", "数量", "单位", "含税单价", "金额"]
    if any(field not in reverse_mapping for field in required_columns):
        return rows, "原生坐标表缺少快速路径必需列"

    sequence_column = reverse_mapping["序号"]
    code_column = reverse_mapping["物料编码"]
    name_column = reverse_mapping["物料名称"]
    quantity_column = reverse_mapping["数量"]
    detail_indexes: list[int] = []
    for row_index in range(header_index + 1, len(rows)):
        row = rows[row_index]
        sequence = normalize_number(row[sequence_column] if sequence_column < len(row) else "")
        quantity = normalize_number(row[quantity_column] if quantity_column < len(row) else "")
        if sequence and quantity and re.fullmatch(r"\d+", sequence):
            detail_indexes.append(row_index)
    if len(detail_indexes) < 2:
        return rows, "原生坐标表可靠明细不足两行"

    reconstructed = [list(row) for row in rows[: header_index + 1]]
    code_fragment_count = 0
    for position, row_index in enumerate(detail_indexes):
        previous_detail = detail_indexes[position - 1] if position > 0 else header_index
        next_detail = detail_indexes[position + 1] if position + 1 < len(detail_indexes) else len(rows)
        row = list(rows[row_index])
        while len(row) < len(rows[header_index]):
            row.append("")

        spec_parts: list[str] = []
        for fragment_row in rows[previous_detail + 1 : row_index]:
            fragment = clean_text(fragment_row[name_column] if name_column < len(fragment_row) else "")
            if fragment:
                spec_parts.append(fragment)
            for column, value in enumerate(fragment_row):
                if column in {code_column, name_column} or not clean_text(value):
                    continue
                while len(row) <= column:
                    row.append("")
                if clean_text(row[column]) and clean_text(row[column]) != clean_text(value):
                    return rows, "规格前置行包含冲突字段"
                row[column] = clean_text(value)

        code_parts = [clean_text(row[code_column])]
        for fragment_row in rows[row_index + 1 : next_detail]:
            fragment_text = " ".join(clean_text(value) for value in fragment_row)
            if any(keyword in fragment_text.lower() for keyword in ["合计", "总计", "总金额", "total"]):
                break
            fragment = clean_text(fragment_row[code_column] if code_column < len(fragment_row) else "")
            if fragment:
                code_parts.append(fragment)
                code_fragment_count += 1
            unexpected = [
                clean_text(value)
                for column, value in enumerate(fragment_row)
                if column not in {code_column, name_column} and clean_text(value)
            ]
            if unexpected:
                return rows, "编码续行包含无法归属的其他字段"

        current_name = clean_text(row[name_column])
        row[code_column] = re.sub(r"\s+", "", "".join(code_parts))
        row[name_column] = "\n".join([part for part in [*spec_parts, current_name] if part])
        reconstructed.append(row)
    if code_fragment_count == 0:
        return rows, "未检测到需要合并的物料编码续行"
    return reconstructed, ""


def _native_coordinate_rows_reliable(
    mapped_rows: list[dict[str, Any]],
    source_rows: list[list[str]],
) -> tuple[bool, str]:
    if len(mapped_rows) < 2:
        return False, "坐标快速路径明细不足两行"
    expected_sequence = 1
    amounts: list[Decimal] = []
    for row in mapped_rows:
        standard = row.get("standard") or {}
        sequence = normalize_number(standard.get("序号"))
        code = clean_text(standard.get("物料编码"))
        name = clean_text(standard.get("物料名称"))
        quantity = _decimal_value(standard.get("数量"))
        price = _decimal_value(standard.get("含税单价"))
        amount = _decimal_value(standard.get("金额"))
        if sequence != str(expected_sequence):
            return False, "明细项次不连续"
        expected_sequence += 1
        if not code or not name or quantity is None or price is None or amount is None or not clean_text(standard.get("单位")):
            return False, "明细必需字段不完整"
        if not re.fullmatch(r"[A-Za-z0-9._/-]{4,100}", code):
            return False, "物料编码仍包含不可靠字符"
        if quantity <= 0 or price < 0 or amount < 0 or abs(quantity * price - amount) > Decimal("0.05"):
            return False, "数量、单价和金额关系未通过校验"
        amounts.append(amount)

    source_text = "\n".join(" ".join(clean_text(value) for value in row) for row in source_rows)
    total_matches = re.findall(
        r"(?:含税总金额|价税合计|总金额|合计|total)\D{0,20}(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)",
        source_text,
        flags=re.I,
    )
    if total_matches:
        totals = {_decimal_value(value) for value in total_matches}
        if sum(amounts) not in totals:
            return False, "明细金额合计与订单总计不一致"
    return True, ""


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
            source_rows = _matrix_from_native_cells(table.get("cells") or [])
            if len(source_rows) < 3:
                continue
            rows, reconstruction_reason = _native_coordinate_rows(source_rows)
            if reconstruction_reason:
                append_fallback_reason(f"native_coordinate:{reconstruction_reason}")
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
            reliable, reliability_reason = _native_coordinate_rows_reliable(table_rows, source_rows)
            if table_issues or not reliable:
                append_fallback_reason(f"native_coordinate:{reliability_reason or '表格映射存在异常'}")
                continue
            raw_detail_tables.append(
                {
                    "page_index": page.get("page_index", 0),
                    "table_index": len(raw_detail_tables),
                    "bbox": [],
                    "rows": rows,
                    "method": "pdf_text_coordinate_fast",
                    "confidence": 1.0,
                }
            )
            for row in table_rows:
                row["method"] = "pdf_text_coordinate_fast"
            mapped_rows.extend(table_rows)
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
    warnings.append("PDF 使用 pdfplumber 原生坐标快速重建明细，未调用 Docling、未渲染页面图片、未调用 OCR。")
    return {
        "pipeline_version": "purchase_order_v1",
        "source_file": source_file,
        "file_type": "pdf",
        "parser_mode": "template_pdf_text_coordinate_fast" if template else "pdf_text_coordinate_fast",
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


def _native_table_has_required_fields(mapping: dict[int, str]) -> bool:
    required_fields = {
        "\u7269\u6599\u7f16\u7801",
        "\u6570\u91cf",
        "\u5355\u4f4d",
        "\u542b\u7a0e\u5355\u4ef7",
        "\u91d1\u989d",
        "\u4ea4\u8d27\u65e5\u671f",
    }
    return required_fields.issubset(set(mapping.values()))


def _native_table_rows_reliable(rows: list[dict[str, Any]], *, minimum_rows: int = 2) -> bool:
    if len(rows) < minimum_rows:
        return False
    for row in rows:
        standard = row.get("standard") or {}
        code = clean_text(standard.get("\u7269\u6599\u7f16\u7801"))
        description = clean_text(standard.get("\u7269\u6599\u540d\u79f0")) or clean_text(standard.get("\u8bf4\u660e"))
        quantity = _decimal_value(standard.get("\u6570\u91cf"))
        price = _decimal_value(standard.get("\u542b\u7a0e\u5355\u4ef7"))
        amount = _decimal_value(standard.get("\u91d1\u989d"))
        delivery_date = clean_text(standard.get("\u4ea4\u8d27\u65e5\u671f"))
        if not code or not description or quantity is None or price is None or amount is None or not delivery_date:
            return False
        if quantity <= 0 or price < 0 or amount < 0 or abs(quantity * price - amount) > Decimal("0.05"):
            return False
    return True


def _native_table_grand_total(raw_tables: list[dict[str, Any]]) -> Decimal | None:
    for table in raw_tables:
        for row in table.get("rows") or []:
            text = " ".join(clean_text(value) for value in row)
            if not re.search(r"grand\s+amount|amount\s+total|(?:合计|总计).*(?:金额|amount)|(?:金额|amount).*(?:合计|总计)", text, flags=re.I):
                continue
            values = [
                _decimal_value(value)
                for value in re.findall(r"-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text)
            ]
            values = [value for value in values if value is not None]
            if values:
                return values[-1]
    return None


def _native_table_purchase_document(file_item: dict[str, str], native: dict[str, Any]) -> dict[str, Any] | None:
    quality = native.get("text_quality") or {}
    if not quality.get("has_text") or not quality.get("readable"):
        return None

    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    raw_detail_tables: list[dict[str, Any]] = []
    mapped_rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    reusable_headers: list[str] | None = None

    for page in native.get("pages") or []:
        for table in page.get("tables") or []:
            source_rows = _repair_split_native_headers(
                _matrix_from_native_cells(table.get("cells") or [], drop_trailing_sparse=False)
            )
            if not source_rows:
                continue
            header_index, mapping = find_detail_header_row(source_rows)
            if header_index is not None and _native_table_has_required_fields(mapping):
                source_rows = _compact_native_material_code_column(source_rows, header_index, mapping)
                reusable_headers = list(source_rows[header_index])
                table_rows = [list(row) for row in source_rows[header_index:]]
            elif reusable_headers is not None:
                continuation_rows = (
                    [list(row) for row in source_rows]
                    if max((len(row) for row in source_rows), default=0) == len(reusable_headers)
                    else _native_grid_rows_from_page(page, len(reusable_headers))
                )
                if not continuation_rows:
                    continue
                table_rows = [list(reusable_headers), *continuation_rows]
            else:
                continue

            table_document = {
                "page_index": page.get("page_index", 0),
                "table_index": table.get("table_index", len(raw_detail_tables)),
                "table_type": "detail_table",
                "raw_rows": table_rows,
                "rows": table_rows,
                "bbox": [],
                "method": "pdf_native_table_fast",
                "confidence": 1.0,
            }
            table_rows_mapped, table_issues = build_detail_rows_from_table(table_document)
            if not _native_table_rows_reliable(table_rows_mapped, minimum_rows=1):
                return None
            for row in table_rows_mapped:
                row["method"] = "pdf_native_table_fast"
            raw_detail_tables.append(
                {
                    "page_index": table_document["page_index"],
                    "table_index": len(raw_detail_tables),
                    "bbox": [],
                    "rows": table_rows,
                    "method": "pdf_native_table_fast",
                    "confidence": 1.0,
                }
            )
            mapped_rows.extend(table_rows_mapped)
            issues.extend(table_issues)

    if not _native_table_rows_reliable(mapped_rows):
        return None
    grand_total = _native_table_grand_total(raw_detail_tables)
    if grand_total is not None:
        parsed_total = sum((_decimal_value((row.get("standard") or {}).get("\u91d1\u989d")) or Decimal(0) for row in mapped_rows), Decimal(0))
        if parsed_total != grand_total:
            return None

    lines = [line.strip() for line in clean_text(native.get("text")).splitlines() if line.strip()]
    header_info = extract_key_values(lines)
    template = identify_template(source_file, "\n".join(lines))
    sections = {"\u5907\u6ce8": [], "\u6761\u6b3e": [], "\u4ed8\u6b3e\u4fe1\u606f": [], "\u6536\u8d27\u4fe1\u606f": [], "\u7b7e\u6838\u533a": []}
    for line in lines:
        section = classify_section_line(line)
        if section:
            sections.setdefault(section, []).append(line)

    warnings = list(native.get("warnings") or [])
    warnings.append("PDF used validated native table extraction across pages.")
    return {
        "pipeline_version": "purchase_order_v1",
        "source_file": source_file,
        "file_type": "pdf",
        "parser_mode": "template_pdf_native_table_fast" if template else "pdf_native_table_fast",
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
                        "row_count": max(0, len(table.get("rows") or []) - 1),
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


def _native_detail_rows_missing_from_docling(
    native: dict[str, Any], mapped_rows: list[dict[str, Any]]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """Recover validated native rows that Docling omitted from its Markdown."""
    known = {
        (
            clean_text((row.get("standard") or {}).get("序号")),
            clean_text((row.get("standard") or {}).get("物料编码")),
        )
        for row in mapped_rows
    }
    recovered_tables: list[dict[str, Any]] = []
    recovered_rows: list[dict[str, Any]] = []
    recovered_issues: list[dict[str, Any]] = []
    for page in native.get("pages") or []:
        for table in page.get("tables") or []:
            source_rows = _repair_split_native_headers(
                _matrix_from_native_cells(table.get("cells") or [], drop_trailing_sparse=False)
            )
            header_index, mapping = find_detail_header_row(source_rows)
            if header_index is None or not _native_table_has_required_fields(mapping):
                continue
            table_rows = [list(row) for row in source_rows[header_index:]]
            table_document = {
                "page_index": page.get("page_index", 0),
                "table_index": table.get("table_index", 0),
                "table_type": "detail_table",
                "rows": table_rows,
                "raw_rows": table_rows,
                "method": "pdf_native_table_recovery",
                "confidence": 1.0,
            }
            candidates, issues = build_detail_rows_from_table(table_document)
            if not candidates or not _native_table_rows_reliable(candidates, minimum_rows=1):
                continue
            missing: list[dict[str, Any]] = []
            for row in candidates:
                standard = row.get("standard") or {}
                identity = (clean_text(standard.get("序号")), clean_text(standard.get("物料编码")))
                if not any(identity) or identity in known:
                    continue
                row["method"] = "pdf_native_table_recovery"
                missing.append(row)
                known.add(identity)
            if not missing:
                continue
            recovered_tables.append(
                {
                    "page_index": page.get("page_index", 0),
                    "table_index": table.get("table_index", len(recovered_tables)),
                    "bbox": [],
                    "rows": table_rows,
                    "method": "pdf_native_table_recovery",
                    "confidence": 1.0,
                    "title": "原生表格补充明细",
                }
            )
            recovered_rows.extend(missing)
            recovered_issues.extend(issues)
    return recovered_tables, recovered_rows, recovered_issues


def _docling_purchase_document(file_item: dict[str, str], native: dict[str, Any]) -> dict[str, Any] | None:
    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    docling_result = parse_pdf_with_docling(
        input_path,
        content_sha256=clean_text(file_item.get("content_sha256")),
    )
    set_cache_state("docling_hit", bool(docling_result.get("cache_hit")))
    set_cache_state("docling_worker_used", bool(docling_result.get("worker_used")))
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

    recovered_tables, recovered_rows, recovered_issues = _native_detail_rows_missing_from_docling(
        native, mapped_rows
    )
    raw_detail_tables.extend(recovered_tables)
    mapped_rows.extend(recovered_rows)
    issues.extend(recovered_issues)

    parser_mode = "template_docling_markdown_pdf" if template else "docling_markdown_pdf"
    warnings = list(native.get("warnings") or [])
    if docling_result.get("error"):
        warnings.append(f"Docling 解析提示：{docling_result.get('error')}")
    warnings.append("PDF 使用 Docling/Markdown 快速分层解析，未渲染页面图片，未调用 OCR。")
    if recovered_rows:
        warnings.append(f"Docling 漏行时已从原生 PDF 表格安全补充 {len(recovered_rows)} 条明细。")
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
    pipeline_started = time.perf_counter()
    input_path = Path(file_item["stored_path"])
    source_file = file_item.get("original_filename") or input_path.name
    content_sha256 = clean_text(file_item.get("content_sha256"))
    hash_ms = float(file_item.get("hash_ms") or 0.0)
    if not content_sha256:
        hash_started = time.perf_counter()
        content_sha256 = file_sha256(input_path)
        hash_ms = (time.perf_counter() - hash_started) * 1000
    summary = new_performance_summary(content_sha256=content_sha256, hash_ms=hash_ms)
    performance_token = activate_performance(summary)
    cleanup: tempfile.TemporaryDirectory[str] | None = None
    parser_cache_hit = False

    def finish(document: dict[str, Any], *, cacheable: bool = True) -> dict[str, Any]:
        if cacheable and not parser_cache_hit:
            with performance_stage("parser_cache_write"):
                saved = save_parser_cache(content_sha256, input_path.suffix, document)
            summary.setdefault("cache", {})["parser_saved"] = bool(saved)
        add_stage_ms("pipeline_total", (time.perf_counter() - pipeline_started) * 1000, summary=summary)
        document["performance_summary"] = summary
        return document

    try:
        with performance_stage("parser_cache_read"):
            cached_document, cache_key, cache_reason = load_parser_cache(content_sha256, input_path.suffix)
        summary["cache"].update(
            {
                "parser_hit": cached_document is not None,
                "parser_key": cache_key,
                "parser_reason": cache_reason,
            }
        )
        if cached_document is not None:
            parser_cache_hit = True
            cached_document["source_file"] = source_file
            cached_document["file_type"] = _file_type(input_path)
            cached_document["started_at"] = datetime.now().isoformat(timespec="seconds")
            for stale_key in ["ai_repair_summary", "factory_import", "factory_mapping_summary"]:
                cached_document.pop(stale_key, None)
            set_fast_path("parser_cache")
            return finish(cached_document, cacheable=False)

        if work_dir is None:
            cleanup = tempfile.TemporaryDirectory()
            base_dir = Path(cleanup.name)
        else:
            base_dir = Path(work_dir)
            base_dir.mkdir(parents=True, exist_ok=True)

        native: dict[str, Any] | None = None
        if input_path.suffix.lower() == ".pdf":
            with performance_stage("pdf_native"):
                native = _native_pdf_summary(input_path)
            native_text = str(native.get("text") or "")
            with performance_stage("pdf_native_table_fast_path"):
                native_table_document = _native_table_purchase_document(file_item, native)
            if native_table_document:
                set_fast_path("pdf_native_table_fast")
                with performance_stage("normalize"):
                    normalized = normalize_purchase_document(native_table_document, source_text=native_text)
                return finish(normalized)
            append_fallback_reason("pdf_native_table_fast_not_reliable")
            with performance_stage("pdf_native_fast_path"):
                native_text_document = _native_text_purchase_document(file_item, native)
            if native_text_document:
                set_fast_path("pdf_text_line_fast")
                with performance_stage("normalize"):
                    normalized = normalize_purchase_document(native_text_document, source_text=native_text)
                return finish(normalized)
            append_fallback_reason("pdf_text_line_fast_not_reliable")
            with performance_stage("pdf_native_table_fast_path"):
                native_document = _native_purchase_document(file_item, native)
            if native_document:
                set_fast_path("pdf_native_table_fast")
                with performance_stage("normalize"):
                    normalized = normalize_purchase_document(native_document, source_text=native_text)
                return finish(normalized)
            append_fallback_reason("pdf_native_table_fast_not_reliable")
            with performance_stage("docling"):
                docling_document = _docling_purchase_document(file_item, native)
            if docling_document:
                set_fast_path("docling")
                with performance_stage("normalize"):
                    normalized = normalize_purchase_document(docling_document, source_text=native_text)
                return finish(normalized)
            append_fallback_reason("docling_no_reliable_detail_table")

        render_dir = base_dir / "pages"
        clean_dir = base_dir / "clean"
        render_dpi = 180 if input_path.suffix.lower() == ".pdf" else 240
        with performance_stage("render"):
            rendered_pages = render_input_pages(file_item, render_dir, dpi=render_dpi)
        if not rendered_pages:
            raise RuntimeError("未能渲染页面图片。")

        pages: list[dict[str, Any]] = []
        with performance_stage("preprocess"):
            clean_pages = [preprocess_page_image(rendered_page, clean_dir) for rendered_page in rendered_pages]
        layout_cache = None
        with performance_stage("page_segmentation"):
            if len(clean_pages) == 1:
                clean_image_path = Path(clean_pages[0].get("clean_image_path") or clean_pages[0]["image_path"])
                with performance_stage("grid_detection"):
                    precheck_grids = _cached_grids(clean_image_path)
                has_clear_grid = any(
                    int(grid.get("column_count") or 0) >= 6 and int(grid.get("row_count") or 0) >= 2
                    for grid in precheck_grids
                )
                if has_clear_grid:
                    pages.append(segment_purchase_page(clean_pages[0]))
                else:
                    quick_lines = _quick_page_lines(clean_pages[0])
                    probe_lines = []
                    seen_probe_lines: set[str] = set()
                    for line in quick_lines:
                        key = clean_text(line).lower().replace(" ", "")
                        if key and key not in seen_probe_lines:
                            seen_probe_lines.add(key)
                            probe_lines.append(clean_text(line))
                    auxiliary_text = _native_pdf_text(file_item, native)
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
                        pages.append(segment_purchase_page(clean_pages[0]))
            else:
                for clean_page in clean_pages:
                    pages.append(segment_purchase_page(clean_page))

        auxiliary_text = _native_pdf_text(file_item, native)
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
        with performance_stage("normalize"):
            normalized_document = normalize_purchase_document(document, source_text="\n".join(lines))
        if not layout_cache and save_layout_cache(normalized_document):
            normalized_document.setdefault("warnings", []).append("已保存本次版式缓存，后续相同版式将优先快速识别。")
        set_fast_path(parser_mode)
        return finish(normalized_document)
    finally:
        reset_performance(performance_token)
        if cleanup is not None:
            cleanup.cleanup()
