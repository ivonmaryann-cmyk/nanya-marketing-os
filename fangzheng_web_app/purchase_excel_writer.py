from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .purchase_field_rules import STANDARD_HEADERS, clean_text, decimal_or_none, normalize_date


HEADER_FILL = "1F4E78"
SECTION_FILL = "D9EAF7"
LIGHT_FILL = "EEF3F8"
TITLE_FILL = "0B1F4D"
THIN_SIDE = Side(style="thin", color="A6A6A6")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


def _style_cell(cell, *, bold: bool = False, fill: str | None = None, color: str = "000000", align: str = "left") -> None:
    cell.font = Font(name="Microsoft YaHei", bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = THIN_BORDER
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def _style_row(ws, row: int, start_col: int, end_col: int, *, fill: str, color: str = "000000", bold: bool = True) -> None:
    for column in range(start_col, end_col + 1):
        _style_cell(ws.cell(row=row, column=column), bold=bold, fill=fill, color=color, align="center")


def _set_purchase_sheet_layout(ws) -> None:
    widths = {"A": 8, "B": 28, "C": 24, "D": 26, "E": 18, "F": 16, "G": 18, "H": 18, "I": 18, "J": 20, "K": 24, "L": 24}
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    for index in range(13, 31):
        ws.column_dimensions[get_column_letter(index)].width = 18
    ws.sheet_view.showGridLines = False


def _set_detail_sheet_layout(ws, col_count: int) -> None:
    for index in range(1, col_count + 1):
        letter = get_column_letter(index)
        ws.column_dimensions[letter].width = 16
    for index in range(1, min(col_count, 4) + 1):
        ws.column_dimensions[get_column_letter(index)].width = 18
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"


def _write_key_values(ws, row: int, header_info: dict[str, Any], max_col: int) -> int:
    pairs = [(key, value) for key, value in header_info.items() if clean_text(value)]
    if not pairs:
        return row
    title_row = row
    ws.cell(row=row, column=1, value="订单头信息")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    _style_row(ws, row, 1, max_col, fill=SECTION_FILL)
    row += 1

    for index in range(0, len(pairs), 2):
        left_key, left_value = pairs[index]
        ws.cell(row=row, column=1, value=left_key)
        ws.cell(row=row, column=2, value=left_value)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=max(2, max_col // 2))
        _style_cell(ws.cell(row=row, column=1), bold=True, fill=LIGHT_FILL, align="center")
        _style_cell(ws.cell(row=row, column=2), align="left")
        if index + 1 < len(pairs):
            right_key, right_value = pairs[index + 1]
            right_label_col = max_col // 2 + 1
            right_value_col = right_label_col + 1
            ws.cell(row=row, column=right_label_col, value=right_key)
            ws.cell(row=row, column=right_value_col, value=right_value)
            ws.merge_cells(start_row=row, start_column=right_value_col, end_row=row, end_column=max_col)
            _style_cell(ws.cell(row=row, column=right_label_col), bold=True, fill=LIGHT_FILL, align="center")
            _style_cell(ws.cell(row=row, column=right_value_col), align="left")
        for column in range(1, max_col + 1):
            ws.cell(row=row, column=column).border = THIN_BORDER
        row += 1
    return row + 1


def _write_raw_table(ws, row: int, table: dict[str, Any], max_col_hint: int) -> tuple[int, int]:
    rows = table.get("rows") or []
    if not rows:
        return row, 0
    max_col = max(max((len(source_row) for source_row in rows), default=1), max_col_hint)
    title = f"明细表 - 第 {int(table.get('page_index') or 0) + 1} 页"
    ws.cell(row=row, column=1, value=title)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    _style_row(ws, row, 1, max_col, fill=SECTION_FILL)
    row += 1
    start = row
    for row_values in rows:
        for column in range(1, max_col + 1):
            value = row_values[column - 1] if column <= len(row_values) else ""
            cell = ws.cell(row=row, column=column, value=value)
            _style_cell(cell, align="left")
        row += 1
    if row > start:
        _style_row(ws, start, 1, max_col, fill=HEADER_FILL, color="FFFFFF")
    for row_index in range(start + 1, row):
        ws.row_dimensions[row_index].height = 36
    return row + 1, max(0, len(rows) - 1)


def _rebuilt_detail_rows(document: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        detail
        for detail in document.get("mapped_detail_rows") or []
        if "packed_text_rebuild" in str(detail.get("method") or "")
    ]


def _write_rebuilt_detail_table(ws, row: int, document: dict[str, Any], max_col_hint: int) -> tuple[int, int]:
    details = _rebuilt_detail_rows(document)
    if not details:
        return row, 0
    headers = list(STANDARD_HEADERS)
    max_col = max(max_col_hint, len(headers))
    page_indexes = sorted({int(detail.get("page_index") or 0) + 1 for detail in details})
    page_text = "、".join(str(index) for index in page_indexes) if page_indexes else "1"
    ws.cell(row=row, column=1, value=f"明细表 - 第 {page_text} 页（已重建分列）")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    _style_row(ws, row, 1, max_col, fill=SECTION_FILL)
    row += 1

    for column, header in enumerate(headers, start=1):
        ws.cell(row=row, column=column, value=header)
    _style_row(ws, row, 1, max_col, fill=HEADER_FILL, color="FFFFFF")
    row += 1

    for detail in details:
        standard = detail.get("standard") or {}
        for column, header in enumerate(headers, start=1):
            value = _typed_value(header, standard.get(header, ""))
            cell = ws.cell(row=row, column=column, value=value)
            align = "right" if header in {"数量", "含税单价", "金额"} else "left"
            _style_cell(cell, align=align)
            number_format = _detail_number_format(f"标准-{header}")
            if number_format:
                cell.number_format = number_format
        ws.row_dimensions[row].height = 36
        row += 1
    return row + 1, len(details)


def _write_sections(ws, row: int, sections: dict[str, list[str]], max_col: int) -> int:
    for section in ["备注", "条款", "付款信息", "收货信息", "签核区"]:
        lines = [clean_text(line) for line in sections.get(section) or [] if clean_text(line)]
        if not lines:
            continue
        ws.cell(row=row, column=1, value=section)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
        _style_row(ws, row, 1, max_col, fill=SECTION_FILL)
        row += 1
        ws.cell(row=row, column=1, value="\n".join(lines))
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
        _style_cell(ws.cell(row=row, column=1), align="left")
        for column in range(1, max_col + 1):
            ws.cell(row=row, column=column).border = THIN_BORDER
        ws.row_dimensions[row].height = min(max(36, len(lines) * 20), 160)
        row += 2
    return row


def _typed_value(header: str, value: Any):
    text = clean_text(value)
    if not text:
        return ""
    if header in {"数量", "含税单价", "金额"}:
        number = decimal_or_none(text)
        if number is not None:
            return float(number) if number % 1 else int(number)
    if header == "交货日期":
        date = normalize_date(text)
        return date or text
    return text


def _standard_header(name: str) -> str:
    return f"标准-{name}"


def _detail_number_format(header: str) -> str | None:
    field = header.replace("标准-", "")
    if field == "数量":
        return "#,##0.###"
    if field == "含税单价":
        return "#,##0.000"
    if field == "金额":
        return "#,##0.00"
    if field == "交货日期":
        return "yyyy-mm-dd"
    return None


def _write_purchase_sheet(ws, documents: list[dict[str, Any]]) -> dict[str, int]:
    _set_purchase_sheet_layout(ws)
    row = 1
    detail_count = 0
    page_count = 0
    if not documents:
        ws.cell(row=1, column=1, value="未识别到可转换文件")
        return {"detail_count": 0, "page_count": 0}

    for index, document in enumerate(documents, start=1):
        raw_tables = document.get("raw_detail_tables") or []
        rebuilt_rows = _rebuilt_detail_rows(document)
        max_col = max([7, len(STANDARD_HEADERS) if rebuilt_rows else 0] + [max((len(r) for r in table.get("rows") or []), default=1) for table in raw_tables])
        title = clean_text(document.get("source_file")) or f"文件 {index}"
        ws.cell(row=row, column=1, value=f"{index}. {title}")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
        _style_row(ws, row, 1, max_col, fill=TITLE_FILL, color="FFFFFF")
        ws.row_dimensions[row].height = 24
        row += 1
        row = _write_key_values(ws, row, document.get("header_info") or {}, max_col)
        if rebuilt_rows:
            row, count = _write_rebuilt_detail_table(ws, row, document, max_col)
            detail_count += count
        elif raw_tables:
            for table in raw_tables:
                row, count = _write_raw_table(ws, row, table, max_col)
                detail_count += count
        else:
            ws.cell(row=row, column=1, value="未恢复出明细表")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
            _style_row(ws, row, 1, max_col, fill="FCE4D6")
            row += 2
        row = _write_sections(ws, row, document.get("sections") or {}, max_col)
        row += 1
        page_count += int(document.get("page_count") or 0)
    return {"detail_count": detail_count, "page_count": page_count}


def _detail_headers(documents: list[dict[str, Any]]) -> list[str]:
    original_headers: list[str] = []
    seen: set[str] = set()
    for document in documents:
        for row in document.get("mapped_detail_rows") or []:
            for header in (row.get("original") or {}).keys():
                header_text = clean_text(header)
                key = header_text.lower()
                if header_text and key not in seen:
                    seen.add(key)
                    original_headers.append(header_text)
    headers = []
    if len(documents) > 1:
        headers.extend(["来源文件", "页码", "订单号"])
    headers.extend(original_headers)
    headers.extend([f"标准-{header}" for header in STANDARD_HEADERS])
    return headers


def _write_detail_sheet(ws, documents: list[dict[str, Any]]) -> int:
    headers = _detail_headers(documents)
    if not headers:
        headers = STANDARD_HEADERS
    qty_col = headers.index(_standard_header("数量")) + 1 if _standard_header("数量") in headers else None
    price_col = headers.index(_standard_header("含税单价")) + 1 if _standard_header("含税单价") in headers else None
    _set_detail_sheet_layout(ws, len(headers))
    for column, header in enumerate(headers, start=1):
        ws.cell(row=1, column=column, value=header)
    _style_row(ws, 1, 1, len(headers), fill=HEADER_FILL, color="FFFFFF")

    row_index = 2
    for document in documents:
        header_info = document.get("header_info") or {}
        for detail in document.get("mapped_detail_rows") or []:
            values: dict[str, Any] = {}
            if len(documents) > 1:
                values.update(
                    {
                        "来源文件": document.get("source_file", ""),
                        "页码": int(detail.get("page_index") or 0) + 1,
                        "订单号": header_info.get("订单号", ""),
                    }
                )
            original = detail.get("original") or {}
            standard = detail.get("standard") or {}
            values.update(original)
            values.update({f"标准-{key}": value for key, value in standard.items()})
            for column, header in enumerate(headers, start=1):
                value = values.get(header, "")
                if header.startswith("标准-"):
                    field = header.replace("标准-", "")
                    if field == "金额" and not clean_text(value) and qty_col and price_col:
                        qty = decimal_or_none(standard.get("数量"))
                        price = decimal_or_none(standard.get("含税单价"))
                        if qty is not None and price is not None:
                            value = f"={get_column_letter(qty_col)}{row_index}*{get_column_letter(price_col)}{row_index}"
                        else:
                            value = _typed_value(field, value)
                    else:
                        value = _typed_value(field, value)
                cell = ws.cell(row=row_index, column=column, value=value)
                align = "right" if header.replace("标准-", "") in {"数量", "含税单价", "金额"} else "left"
                _style_cell(cell, align=align)
                number_format = _detail_number_format(header)
                if number_format:
                    cell.number_format = number_format
            row_index += 1
    return max(0, row_index - 2)


def _write_issue_sheet(ws, documents: list[dict[str, Any]]) -> int:
    headers = ["文件", "页码", "区域", "字段", "原始识别值", "清洗后值", "置信度", "问题说明"]
    for column, header in enumerate(headers, start=1):
        ws.cell(row=1, column=column, value=header)
    _style_row(ws, 1, 1, len(headers), fill="9E480E", color="FFFFFF")
    widths = [24, 8, 18, 18, 45, 30, 10, 45]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    row_index = 2
    for document in documents:
        for issue in document.get("issues") or []:
            values = [
                document.get("source_file", ""),
                "" if issue.get("page_index") in {"", None} else int(issue.get("page_index") or 0) + 1,
                issue.get("region", ""),
                issue.get("field", ""),
                issue.get("raw_value", ""),
                issue.get("clean_value", ""),
                issue.get("confidence", ""),
                issue.get("message", ""),
            ]
            for column, value in enumerate(values, start=1):
                _style_cell(ws.cell(row=row_index, column=column, value=value), align="left")
            ws.row_dimensions[row_index].height = 48
            row_index += 1
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    return max(0, row_index - 2)


def write_purchase_order_workbook(documents: list[dict[str, Any]], output_path: Path) -> dict[str, int]:
    wb = Workbook()
    purchase_ws = wb.active
    purchase_ws.title = "采购单"
    purchase_stats = _write_purchase_sheet(purchase_ws, documents)

    detail_ws = wb.create_sheet("明细数据")
    detail_count = _write_detail_sheet(detail_ws, documents)

    issue_count = sum(len(document.get("issues") or []) for document in documents)
    if issue_count:
        issue_ws = wb.create_sheet("识别日志")
        issue_count = _write_issue_sheet(issue_ws, documents)

    wb.save(output_path)
    return {
        "structured_count": detail_count,
        "cell_count": detail_count,
        "issue_count": issue_count,
        "page_count": purchase_stats.get("page_count", 0),
    }
