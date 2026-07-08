from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

from PIL import Image

from .cell_ocr import ocr_cell, ocr_image_regions
from .purchase_field_rules import (
    clean_text,
    classify_section_line,
    find_detail_header_row,
    looks_like_detail_data,
    map_detail_row,
    normalize_date,
    normalize_number,
)
from .table_structure_detector import detect_table_grids, iter_grid_cells


_GRID_CACHE: dict[str, list[dict[str, Any]]] = {}
_OCR_REGION_CACHE: dict[str, list[dict[str, Any]]] = {}


def _load_cv2():
    try:
        import cv2
        import numpy as np
    except Exception as exc:  # pragma: no cover - optional dependency
        raise RuntimeError("未安装 OpenCV，无法进行单元格空白检测。") from exc
    return cv2, np


def _image_cache_key(image_path: Path) -> str:
    digest = hashlib.sha1()
    with image_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cached_grids(image_path: Path) -> list[dict[str, Any]]:
    key = _image_cache_key(image_path)
    if key not in _GRID_CACHE:
        _GRID_CACHE[key] = detect_table_grids(image_path)
    return [dict(grid) for grid in _GRID_CACHE[key]]


def _cached_ocr_regions(image_path: Path) -> list[dict[str, Any]]:
    key = _image_cache_key(image_path)
    if key not in _OCR_REGION_CACHE:
        with Image.open(image_path) as image:
            _OCR_REGION_CACHE[key] = ocr_image_regions(image.convert("RGB"))
    return [dict(region) for region in _OCR_REGION_CACHE[key]]


def _is_blank_crop(image: Image.Image) -> bool:
    cv2, np = _load_cv2()
    gray = np.array(image.convert("L"))
    if gray.size == 0:
        return True
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    ink_ratio = float(np.count_nonzero(binary)) / float(binary.size)
    return ink_ratio < 0.0045


def _crop_cell(image: Image.Image, bbox: list[int], padding: int = 1) -> Image.Image:
    x0, y0, x1, y1 = bbox
    return image.crop(
        (
            max(x0 + padding, 0),
            max(y0 + padding, 0),
            min(x1 - padding, image.width),
            min(y1 - padding, image.height),
        )
    )


def _region_center(region: dict[str, Any]) -> tuple[float, float]:
    bbox = region.get("bbox") or [0, 0, 0, 0]
    return (float(bbox[0]) + float(bbox[2])) / 2, (float(bbox[1]) + float(bbox[3])) / 2


def _regions_in_bbox(regions: list[dict[str, Any]], bbox: list[int]) -> list[dict[str, Any]]:
    x0, y0, x1, y1 = bbox
    contained = []
    for region in regions:
        cx, cy = _region_center(region)
        if x0 <= cx <= x1 and y0 <= cy <= y1:
            contained.append(region)
    return contained


def _join_regions(regions: list[dict[str, Any]]) -> tuple[str, float]:
    if not regions:
        return "", 0.0
    ordered = sorted(regions, key=lambda item: ((item.get("bbox") or [0, 0, 0, 0])[1], (item.get("bbox") or [0, 0, 0, 0])[0]))
    text = " ".join(clean_text(region.get("text")) for region in ordered if clean_text(region.get("text")))
    confidences = [float(region.get("confidence") or 0) for region in ordered]
    confidence = sum(confidences) / len(confidences) if confidences else 0.0
    return clean_text(text), confidence


def _ocr_grid_rows(image_path: Path, grid: dict[str, Any], page_index: int, regions: list[dict[str, Any]]) -> tuple[list[list[str]], list[dict[str, Any]]]:
    image = Image.open(image_path).convert("RGB")
    cells: list[dict[str, Any]] = []
    matrix: dict[int, dict[int, str]] = {}
    for row_index, column_index, bbox in iter_grid_cells(grid):
        contained_regions = _regions_in_bbox(regions, bbox)
        if contained_regions:
            text, confidence = _join_regions(contained_regions)
            result = {"text": text, "confidence": confidence, "method": "page_ocr_region"}
        else:
            crop = _crop_cell(image, bbox)
            if _is_blank_crop(crop):
                result = {"text": "", "confidence": 1.0, "method": "manual_empty"}
            else:
                result = ocr_cell(crop)
        text = clean_text(result.get("text"))
        matrix.setdefault(row_index, {})[column_index] = text
        cells.append(
            {
                "page_index": page_index,
                "table_index": grid.get("table_index", 0),
                "row_index": row_index,
                "column_index": column_index,
                "text": text,
                "bbox": bbox,
                "confidence": result.get("confidence", 0),
                "method": result.get("method", "cell_ocr"),
            }
        )
    columns = sorted({column for row in matrix.values() for column in row})
    rows: list[list[str]] = []
    for row_index in sorted(matrix):
        row = [matrix[row_index].get(column, "") for column in columns]
        if any(clean_text(value) for value in row):
            rows.append(row)
    return rows, cells


def _column_index_for_x(x_positions: list[int], center_x: float) -> int | None:
    for column_index in range(len(x_positions) - 1):
        if x_positions[column_index] <= center_x <= x_positions[column_index + 1]:
            return column_index
    return None


def _cluster_regions_by_row(regions: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    if not regions:
        return []
    heights = [max(1.0, float((region.get("bbox") or [0, 0, 0, 0])[3]) - float((region.get("bbox") or [0, 0, 0, 0])[1])) for region in regions]
    tolerance = max(8.0, sorted(heights)[len(heights) // 2] * 0.75)
    rows: list[dict[str, Any]] = []
    for region in sorted(regions, key=lambda item: (_region_center(item)[1], _region_center(item)[0])):
        _cx, cy = _region_center(region)
        target = None
        for row in rows:
            if abs(cy - row["y"]) <= tolerance:
                target = row
                break
        if target is None:
            rows.append({"y": cy, "regions": [region]})
        else:
            target["regions"].append(region)
            target["y"] = (target["y"] * (len(target["regions"]) - 1) + cy) / len(target["regions"])
    return [row["regions"] for row in rows]


def _rows_from_region_grid(regions: list[dict[str, Any]], grid: dict[str, Any], page_index: int) -> tuple[list[list[str]], list[dict[str, Any]]]:
    bbox = grid.get("bbox") or [0, 0, 0, 0]
    x_positions = grid.get("x_positions") or []
    if len(x_positions) < 2:
        return [], []
    table_regions = _regions_in_bbox(regions, bbox)
    row_clusters = _cluster_regions_by_row(table_regions)
    rows: list[list[str]] = []
    cells: list[dict[str, Any]] = []
    max_columns = len(x_positions) - 1
    for row_index, row_regions in enumerate(row_clusters):
        buckets: dict[int, list[dict[str, Any]]] = {}
        for region in row_regions:
            cx, _cy = _region_center(region)
            column_index = _column_index_for_x(x_positions, cx)
            if column_index is None:
                continue
            buckets.setdefault(column_index, []).append(region)
        row = []
        for column_index in range(max_columns):
            column_regions = buckets.get(column_index, [])
            text, confidence = _join_regions(column_regions)
            row.append(text)
            if column_regions:
                xs = []
                ys = []
                for region in column_regions:
                    rb = region.get("bbox") or [0, 0, 0, 0]
                    xs.extend([float(rb[0]), float(rb[2])])
                    ys.extend([float(rb[1]), float(rb[3])])
                cell_bbox = [round(min(xs)), round(min(ys)), round(max(xs)), round(max(ys))]
            else:
                y_values = [(region.get("bbox") or [0, 0, 0, 0])[1] for region in row_regions]
                y2_values = [(region.get("bbox") or [0, 0, 0, 0])[3] for region in row_regions]
                cell_bbox = [x_positions[column_index], round(min(y_values)), x_positions[column_index + 1], round(max(y2_values))]
            cells.append(
                {
                    "page_index": page_index,
                    "table_index": grid.get("table_index", 0),
                    "row_index": row_index,
                    "column_index": column_index,
                    "text": text,
                    "bbox": cell_bbox,
                    "confidence": confidence,
                    "method": "page_ocr_row_cluster",
                }
            )
        if any(clean_text(value) for value in row):
            rows.append(row)
    return rows, cells


def _region_rows_are_better(region_rows: list[list[str]], cell_rows: list[list[str]], grid: dict[str, Any]) -> bool:
    if len(region_rows) <= len(cell_rows):
        return False
    header_index, _mapping = find_detail_header_row(region_rows)
    if header_index is None:
        return False
    if int(grid.get("row_count") or 0) > 4:
        cell_header_index, _cell_mapping = find_detail_header_row(cell_rows)
        if cell_header_index is not None:
            cell_detail_count = sum(1 for row in cell_rows[cell_header_index + 1 :] if looks_like_detail_data(row))
            if cell_detail_count >= 2:
                return False
    detail_like_count = sum(1 for row in region_rows[header_index + 1 :] if looks_like_detail_data(row))
    return detail_like_count >= 2 or len(region_rows) >= len(cell_rows) + 3


def _table_type(rows: list[list[str]]) -> str:
    header_index, _mapping = find_detail_header_row(rows)
    if header_index is not None:
        return "detail_table"
    data_rows = sum(1 for row in rows if looks_like_detail_data(row))
    if data_rows >= 2:
        return "detail_table"
    text = "\n".join(" ".join(row) for row in rows)
    section_hits = [classify_section_line(line) for line in text.splitlines()]
    if any(hit in {"付款信息", "收货信息"} for hit in section_hits):
        return "payment_shipping"
    if any(hit in {"备注", "条款", "签核区"} for hit in section_hits):
        return "terms_notes"
    return "unknown"


def _inside_any_table(region: dict[str, Any], grids: list[dict[str, Any]]) -> bool:
    bbox = region.get("bbox") or [0, 0, 0, 0]
    cx = (float(bbox[0]) + float(bbox[2])) / 2
    cy = (float(bbox[1]) + float(bbox[3])) / 2
    for grid in grids:
        x0, y0, x1, y1 = grid.get("bbox") or [0, 0, 0, 0]
        if x0 <= cx <= x1 and y0 <= cy <= y1:
            return True
    return False


def _looks_like_section_row(text: str) -> bool:
    keywords = [
        "交易条款",
        "付款方式",
        "付款条件",
        "送货地址",
        "收货地址",
        "其它说明",
        "其他说明",
        "供应商",
        "采购方",
        "装运方式",
        "质量保证",
        "环保",
        "违约",
    ]
    return any(keyword in text for keyword in keywords)


def _column_for(mapping: dict[int, str], field: str) -> int | None:
    for column, mapped_field in mapping.items():
        if mapped_field == field:
            return column
    return None


def _tokens(value: str) -> list[str]:
    return [part.strip() for part in clean_text(value).split() if part.strip()]


def _number_tokens(value: str) -> list[str]:
    import re

    text = clean_text(value)
    return [token.replace(",", "") for token in re.findall(r"-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text)]


def _date_tokens(value: str) -> list[str]:
    import re

    text = clean_text(value)
    raw_dates = re.findall(r"20\d{2}[./-]\d{1,2}[./-]\d{1,2}|20\d{6}", text)
    result: list[str] = []
    for raw in raw_dates:
        result.append(normalize_date(raw) or raw)
    return result


def _code_tokens(value: str) -> list[str]:
    import re

    tokens = _tokens(value)
    code_like = [token for token in tokens if re.search(r"[A-Za-z0-9]", token) and len(token) >= 4]
    return code_like or tokens


def _split_text_tokens(value: str, count: int) -> list[str]:
    tokens = _tokens(value)
    if count <= 1 or len(tokens) <= 1:
        return [clean_text(value)] + [""] * max(0, count - 1)
    if len(tokens) >= count:
        result = tokens[: count - 1]
        result.append(" ".join(tokens[count - 1 :]))
        return result
    return tokens + [""] * (count - len(tokens))


def _split_stacked_detail_row(row: list[str], mapping: dict[int, str]) -> list[list[str]]:
    """Split one OCR row that actually contains several detail rows.

    This handles purchase orders with vertical table lines but weak/missing
    horizontal lines: OCR returns one wide row with repeated codes, quantities,
    units, prices and dates. We split by the strongest repeated fields.
    """
    if not mapping:
        return [row]

    token_columns: dict[int, list[str]] = {}
    for column, field in mapping.items():
        value = row[column] if column < len(row) else ""
        if field in {"序号", "物料编码"}:
            token_columns[column] = _code_tokens(value)
        elif field in {"数量", "含税单价", "金额"}:
            token_columns[column] = _number_tokens(value)
        elif field == "交货日期":
            token_columns[column] = _date_tokens(value)
        elif field == "单位":
            token_columns[column] = _tokens(value)

    repeated_lengths = [len(values) for values in token_columns.values() if len(values) >= 2]
    if not repeated_lengths:
        return [row]
    target_count = max(repeated_lengths)
    if target_count < 2:
        return [row]

    split_rows: list[list[str]] = []
    for item_index in range(target_count):
        split_row = list(row)
        for column, field in mapping.items():
            value = row[column] if column < len(row) else ""
            values = token_columns.get(column)
            if values:
                split_row[column] = values[item_index] if item_index < len(values) else ""
                continue
            if field in {"物料名称", "说明", "备注"}:
                parts = _split_text_tokens(value, target_count)
                split_row[column] = parts[item_index] if item_index < len(parts) else ""
        split_rows.append(split_row)
    return split_rows


def segment_purchase_page(page: dict[str, Any]) -> dict[str, Any]:
    image_path = Path(page.get("clean_image_path") or page["image_path"])
    page_index = int(page.get("page_index", 0))
    grids = _cached_grids(image_path)
    tables: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []

    try:
        regions = _cached_ocr_regions(image_path)
    except Exception as exc:
        regions = []
        issues.append(
            {
                "page_index": page_index,
                "region": "页面 OCR",
                "field": "",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": f"页面文本 OCR 失败：{exc}",
            }
        )

    for grid in grids:
        cell_rows, cell_cells = _ocr_grid_rows(image_path, grid, page_index, regions)
        region_rows, region_cells = _rows_from_region_grid(regions, grid, page_index)
        if _region_rows_are_better(region_rows, cell_rows, grid):
            rows, cells, method = region_rows, region_cells, "page_ocr_row_cluster"
        else:
            rows, cells, method = cell_rows, cell_cells, "grid_cell_ocr"
        table_type = _table_type(rows)
        tables.append(
            {
                "page_index": page_index,
                "table_index": grid.get("table_index", len(tables)),
                "table_type": table_type,
                "bbox": grid.get("bbox"),
                "confidence": grid.get("confidence", 0),
                "raw_rows": rows,
                "cells": cells,
                "method": method,
            }
        )

    non_table_regions = [region for region in regions if not _inside_any_table(region, grids)]
    non_table_regions.sort(key=lambda item: ((item.get("bbox") or [0, 0, 0, 0])[1], (item.get("bbox") or [0, 0, 0, 0])[0]))
    text_lines = [clean_text(region.get("text")) for region in non_table_regions if clean_text(region.get("text"))]

    sections = {"备注": [], "条款": [], "付款信息": [], "收货信息": [], "签核区": []}
    for line in text_lines:
        section = classify_section_line(line)
        if section:
            sections.setdefault(section, []).append(line)

    for table in tables:
        for row in table.get("raw_rows") or []:
            row_text = clean_text(" ".join(row))
            if not row_text or not _looks_like_section_row(row_text):
                continue
            section = classify_section_line(row_text) or "条款"
            sections.setdefault(section, []).append(row_text)

    if not any(table.get("table_type") == "detail_table" for table in tables):
        issues.append(
            {
                "page_index": page_index,
                "region": "明细表",
                "field": "",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": "未检测到可靠的明细表格，已保留页面文本供复核。",
            }
        )

    return {
        "page_index": page_index,
        "width": page.get("width"),
        "height": page.get("height"),
        "image_path": str(image_path),
        "tables": tables,
        "text_regions": non_table_regions,
        "text_lines": text_lines,
        "sections": sections,
        "issues": issues,
    }


def build_detail_rows_from_table(table: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rows = table.get("raw_rows") or []
    header_index, mapping = find_detail_header_row(rows)
    issues: list[dict[str, Any]] = []
    if header_index is None:
        issues.append(
            {
                "page_index": table.get("page_index", 0),
                "region": "明细表",
                "field": "表头",
                "raw_value": "\n".join(" | ".join(row) for row in rows[:4]),
                "clean_value": "",
                "confidence": table.get("confidence", 0),
                "message": "未识别到可靠表头，明细数据保留原始列，标准字段尽量留空。",
            }
        )
        if not rows:
            return [], issues
        raw_headers = [f"列{index + 1}" for index in range(max(len(row) for row in rows))]
        data_start = 0
        mapping = {}
    else:
        raw_headers = rows[header_index]
        data_start = header_index + 1

    detail_rows: list[dict[str, Any]] = []
    for row_offset, row in enumerate(rows[data_start:], start=data_start):
        if not any(clean_text(value) for value in row):
            continue
        row_text = " ".join(row)
        if any(keyword in row_text for keyword in ["合计", "总计", "大写金额"]):
            continue
        if _looks_like_section_row(row_text):
            issues.append(
                {
                    "page_index": table.get("page_index", 0),
                    "region": "明细表",
                    "field": "行内容",
                    "raw_value": row_text[:500],
                    "clean_value": "",
                    "confidence": table.get("confidence", 0),
                    "message": "该行更像付款/收货/条款区域，已保留在采购单原始表格中，但不写入明细数据。",
                }
            )
            continue
        if header_index is not None and not looks_like_detail_data(row) and len([v for v in row if clean_text(v)]) <= 2:
            continue
        if table.get("method") == "docling_markdown":
            split_rows = [row]
        else:
            split_rows = _split_stacked_detail_row(row, mapping)
        if len(split_rows) > 1:
            issues.append(
                {
                    "page_index": table.get("page_index", 0),
                    "region": "明细表",
                    "field": "行切分",
                    "raw_value": row_text[:500],
                    "clean_value": f"拆分为 {len(split_rows)} 行",
                    "confidence": table.get("confidence", 0),
                    "message": "检测到同一 OCR 行内包含多条明细，已按物料/数量/日期等重复字段拆分。",
                }
            )
        for split_index, split_row in enumerate(split_rows):
            mapped = map_detail_row(raw_headers, split_row, mapping)
            for note in mapped.get("cleaning_notes") or []:
                issues.append(
                    {
                        "page_index": table.get("page_index", 0),
                        "region": "明细表",
                        "field": "含税单价",
                        "raw_value": " ".join(split_row),
                        "clean_value": mapped.get("standard", {}).get("含税单价", ""),
                        "confidence": table.get("confidence", 0),
                        "message": note,
                    }
                )
            mapped.update(
                {
                    "page_index": table.get("page_index", 0),
                    "table_index": table.get("table_index", 0),
                    "row_index": f"{row_offset}.{split_index}" if len(split_rows) > 1 else row_offset,
                    "raw_text": " ".join(split_row),
                    "confidence": table.get("confidence", 0),
                    "method": "grid_cell_ocr_split" if len(split_rows) > 1 else "grid_cell_ocr",
                }
            )
            detail_rows.append(mapped)
    return detail_rows, issues
