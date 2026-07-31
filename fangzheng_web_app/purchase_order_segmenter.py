from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any

from PIL import Image

from .cell_ocr import ocr_cell, ocr_image_regions
from .purchase_field_rules import (
    clean_text,
    classify_section_line,
    find_detail_header_row,
    header_score,
    looks_like_detail_data,
    map_detail_row,
    normalize_date,
    normalize_number,
)
from .table_structure_detector import detect_table_grids, iter_grid_cells


_GRID_CACHE: dict[str, list[dict[str, Any]]] = {}
_OCR_REGION_CACHE: dict[str, list[dict[str, Any]]] = {}
_TABLE_OCR_REGION_CACHE: dict[str, list[dict[str, Any]]] = {}


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


def _cached_table_ocr_regions(image_path: Path, grid: dict[str, Any]) -> list[dict[str, Any]]:
    image_key = _image_cache_key(image_path)
    bbox = [int(value) for value in (grid.get("bbox") or [0, 0, 0, 0])[:4]]
    cache_key = f"{image_key}:{','.join(str(value) for value in bbox)}"
    if cache_key in _TABLE_OCR_REGION_CACHE:
        return [dict(region) for region in _TABLE_OCR_REGION_CACHE[cache_key]]
    with Image.open(image_path) as image:
        image = image.convert("RGB")
        x0, y0, x1, y1 = bbox
        padding = 8
        crop_box = (
            max(0, x0 - padding),
            max(0, y0 - padding),
            min(image.width, x1 + padding),
            min(image.height, y1 + padding),
        )
        crop = image.crop(crop_box)
    regions = []
    for region in ocr_image_regions(crop):
        rb = region.get("bbox") or [0, 0, 0, 0]
        shifted = dict(region)
        shifted["bbox"] = [
            round(float(rb[0]) + crop_box[0], 2),
            round(float(rb[1]) + crop_box[1], 2),
            round(float(rb[2]) + crop_box[0], 2),
            round(float(rb[3]) + crop_box[1], 2),
        ]
        shifted["method"] = "table_ocr_region"
        regions.append(shifted)
    _TABLE_OCR_REGION_CACHE[cache_key] = regions
    return [dict(region) for region in regions]


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


def _ocr_grid_rows(
    image_path: Path,
    grid: dict[str, Any],
    page_index: int,
    regions: list[dict[str, Any]],
    *,
    max_cell_ocr_fallback: int | None = None,
) -> tuple[list[list[str]], list[dict[str, Any]]]:
    image = Image.open(image_path).convert("RGB")
    cells: list[dict[str, Any]] = []
    matrix: dict[int, dict[int, str]] = {}
    fallback_count = 0
    for row_index, column_index, bbox in iter_grid_cells(grid):
        contained_regions = _regions_in_bbox(regions, bbox)
        if contained_regions:
            text, confidence = _join_regions(contained_regions)
            result = {"text": text, "confidence": confidence, "method": "page_ocr_region"}
        else:
            crop = _crop_cell(image, bbox)
            if _is_blank_crop(crop):
                result = {"text": "", "confidence": 1.0, "method": "manual_empty"}
            elif max_cell_ocr_fallback is not None and fallback_count >= max_cell_ocr_fallback:
                result = {"text": "", "confidence": 0.0, "method": "cell_ocr_skipped"}
            else:
                fallback_count += 1
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


SPARSE_OCR_FIELDS = {"数量", "单位", "含税单价", "金额", "交货日期"}
HEADER_OCR_CORE_FIELDS = {"物料编码", "物料名称", "数量", "含税单价", "金额", "交货日期"}


def _valid_sparse_ocr_value(field: str, value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if field in {"数量", "含税单价", "金额"}:
        return normalize_number(text)
    if field == "交货日期":
        return normalize_date(text)
    if field == "单位":
        if normalize_number(text) == text or len(text) > 8:
            return ""
        return text
    return ""


def _repair_sparse_detail_cells(
    image_path: Path,
    rows: list[list[str]],
    cells: list[dict[str, Any]],
    *,
    max_repairs: int = 8,
) -> list[dict[str, Any]]:
    header_index, mapping = find_detail_header_row(rows)
    if header_index is None or max_repairs <= 0:
        return []

    cell_map = {
        (int(cell.get("row_index") or 0), int(cell.get("column_index") or 0)): cell
        for cell in cells
    }
    actions: list[dict[str, Any]] = []
    with Image.open(image_path) as source_image:
        image = source_image.convert("RGB")
        for row_index in range(header_index + 1, len(rows)):
            row = rows[row_index]
            if not _looks_like_structured_detail_row(row):
                continue
            for column_index, field in mapping.items():
                if field not in SPARSE_OCR_FIELDS:
                    continue
                if column_index < len(row) and clean_text(row[column_index]):
                    continue
                cell = cell_map.get((row_index, column_index))
                bbox = [int(round(float(value))) for value in (cell or {}).get("bbox", [])[:4]]
                if len(bbox) != 4:
                    continue
                crop = _crop_cell(image, bbox)
                if _is_blank_crop(crop):
                    continue
                result = ocr_cell(crop)
                clean_value = _valid_sparse_ocr_value(field, result.get("text"))
                if not clean_value:
                    continue
                while len(row) <= column_index:
                    row.append("")
                row[column_index] = clean_value
                if cell is not None:
                    cell.update(
                        {
                            "text": clean_value,
                            "confidence": result.get("confidence", 0),
                            "method": "cell_ocr_sparse_fallback",
                        }
                    )
                actions.append(
                    {
                        "type": "sparse_cell_ocr",
                        "row_index": row_index,
                        "column_index": column_index,
                        "field": field,
                        "value": clean_value,
                        "confidence": result.get("confidence", 0),
                    }
                )
                if len(actions) >= max_repairs:
                    return actions
    return actions


def _repair_uncertain_detail_headers(
    image_path: Path,
    rows: list[list[str]],
    cells: list[dict[str, Any]],
    *,
    max_repairs: int = 4,
) -> list[dict[str, Any]]:
    header_index, mapping = find_detail_header_row(rows)
    if max_repairs <= 0:
        return []
    if header_index is None:
        scored_rows = []
        for index, row in enumerate(rows[:8]):
            score, row_mapping = header_score(row)
            scored_rows.append((score, index, row_mapping))
        scored_rows.sort(reverse=True)
        if not scored_rows or scored_rows[0][0] < 2:
            return []
        _score, header_index, mapping = scored_rows[0]
    mapped_fields = set(mapping.values())
    has_identity = bool({"物料编码", "物料名称"} & mapped_fields)
    has_quantity = "数量" in mapped_fields
    has_price_or_amount = bool({"含税单价", "金额"} & mapped_fields)
    if has_identity and has_quantity and has_price_or_amount and len(mapped_fields) >= 5:
        return []

    cell_map = {
        (int(cell.get("row_index") or 0), int(cell.get("column_index") or 0)): cell
        for cell in cells
    }
    header_row = rows[header_index]
    before_score, before_mapping = header_score(header_row)
    actions: list[dict[str, Any]] = []
    with Image.open(image_path) as source_image:
        image = source_image.convert("RGB")
        for column_index in range(len(header_row)):
            cell = cell_map.get((header_index, column_index))
            confidence = float((cell or {}).get("confidence") or 0)
            if column_index in before_mapping and confidence >= 0.65:
                continue
            bbox = [int(round(float(value))) for value in (cell or {}).get("bbox", [])[:4]]
            if len(bbox) != 4:
                continue
            crop = _crop_cell(image, bbox)
            if _is_blank_crop(crop):
                continue
            result = ocr_cell(crop)
            candidate = clean_text(result.get("text"))
            if not candidate or candidate == clean_text(header_row[column_index]):
                continue
            trial_row = list(header_row)
            trial_row[column_index] = candidate
            trial_score, trial_mapping = header_score(trial_row)
            new_field = trial_mapping.get(column_index)
            improves_core = bool(new_field in HEADER_OCR_CORE_FIELDS and new_field not in set(before_mapping.values()))
            if trial_score <= before_score and not improves_core:
                continue
            old_value = clean_text(header_row[column_index])
            header_row[column_index] = candidate
            before_score, before_mapping = trial_score, trial_mapping
            if cell is not None:
                cell.update(
                    {
                        "text": candidate,
                        "confidence": result.get("confidence", 0),
                        "method": "cell_ocr_header_fallback",
                    }
                )
            actions.append(
                {
                    "type": "header_cell_ocr",
                    "row_index": header_index,
                    "column_index": column_index,
                    "old_value": old_value,
                    "value": candidate,
                    "field": new_field or "",
                    "confidence": result.get("confidence", 0),
                }
            )
            if len(actions) >= max_repairs:
                break
    return actions


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


def _looks_like_detail_header_text(text: str) -> bool:
    compact = clean_text(text).replace(" ", "").lower()
    return (
        any(keyword in compact for keyword in ["\u5e8f\u53f7", "\u9879\u6b21", "no."])
        and any(keyword in compact for keyword in ["\u4ea7\u54c1\u540d\u79f0", "\u7269\u6599\u540d\u79f0", "\u89c4\u683c", "part", "description"])
        and any(keyword in compact for keyword in ["\u6570\u91cf", "\u8ba2\u8d2d\u6570\u91cf", "qty", "quantity"])
        and any(keyword in compact for keyword in ["\u5355\u4ef7", "\u91d1\u989d", "price", "amount"])
    )


def _looks_like_table_end_text(text: str) -> bool:
    compact = clean_text(text).replace(" ", "")
    return any(keyword in compact for keyword in ["\u5408\u8ba1\u91d1\u989d", "\u5408\u8ba1(RMB", "\u5408\u8ba1\uff1a", "\u5408\u540c\u8ba2\u7acb", "\u5907\u6ce8\uff1a", "\u5236\u5355\uff1a"])


def _virtual_x_positions_from_header(header_regions: list[dict[str, Any]]) -> list[int]:
    ordered = sorted(header_regions, key=lambda item: _region_center(item)[0])
    if len(ordered) < 4:
        return []
    centers = [_region_center(region)[0] for region in ordered]
    positions = [int(max(0, min((region.get("bbox") or [0, 0, 0, 0])[0] for region in ordered) - 20))]
    for left, right in zip(centers, centers[1:]):
        positions.append(int(round((left + right) / 2)))
    right_edge = int(max((region.get("bbox") or [0, 0, 0, 0])[2] for region in ordered) + 80)
    positions.append(right_edge)
    return positions


def _normalize_virtual_grid_rows(rows: list[list[str]]) -> list[list[str]]:
    if len(rows) <= 1:
        return rows
    normalized: list[list[str]] = [rows[0]]
    for row in rows[1:]:
        current = list(row)
        if not current:
            continue
        first_cell = clean_text(current[0])
        match = re.match(r"^(\d+(?:\.\d+)?)\s+(.+)$", first_cell)
        if match and len(current) > 1:
            current[0] = match.group(1)
            current[1] = clean_text(f"{current[1]} {match.group(2)}")
            normalized.append(current)
            continue
        if first_cell and not re.fullmatch(r"\d+(?:\.\d+)?", first_cell) and normalized and len(normalized[-1]) > 1:
            previous = normalized[-1]
            previous[1] = clean_text(f"{previous[1]} {first_cell}")
            if len(current) > 1 and clean_text(current[1]):
                previous[1] = clean_text(f"{previous[1]} {current[1]}")
            continue
        normalized.append(current)
    return normalized


def _infer_text_table_from_regions(regions: list[dict[str, Any]], page_index: int) -> dict[str, Any] | None:
    row_clusters = _cluster_regions_by_row(regions)
    header_cluster_index: int | None = None
    for index, row_regions in enumerate(row_clusters):
        row_text, _confidence = _join_regions(row_regions)
        if _looks_like_detail_header_text(row_text):
            header_cluster_index = index
            break
    if header_cluster_index is None:
        return None

    header_regions = row_clusters[header_cluster_index]
    x_positions = _virtual_x_positions_from_header(header_regions)
    if len(x_positions) < 5:
        return None

    selected_clusters: list[list[dict[str, Any]]] = []
    for row_regions in row_clusters[header_cluster_index:]:
        row_text, _confidence = _join_regions(row_regions)
        if selected_clusters and _looks_like_table_end_text(row_text):
            break
        selected_clusters.append(row_regions)
    if len(selected_clusters) < 3:
        return None

    rows: list[list[str]] = []
    cells: list[dict[str, Any]] = []
    all_xs: list[float] = []
    all_ys: list[float] = []
    for row_index, row_regions in enumerate(selected_clusters):
        buckets: dict[int, list[dict[str, Any]]] = {}
        for region in row_regions:
            bbox = region.get("bbox") or [0, 0, 0, 0]
            cx = (float(bbox[0]) + float(bbox[2])) / 2
            column_index = _column_index_for_x(x_positions, cx)
            if column_index is None:
                continue
            buckets.setdefault(column_index, []).append(region)
            all_xs.extend([float(bbox[0]), float(bbox[2])])
            all_ys.extend([float(bbox[1]), float(bbox[3])])

        row: list[str] = []
        row_y1_values = [(region.get("bbox") or [0, 0, 0, 0])[1] for region in row_regions]
        row_y2_values = [(region.get("bbox") or [0, 0, 0, 0])[3] for region in row_regions]
        for column_index in range(len(x_positions) - 1):
            column_regions = buckets.get(column_index, [])
            text, confidence = _join_regions(column_regions)
            row.append(text)
            if column_regions:
                xs: list[float] = []
                ys: list[float] = []
                for region in column_regions:
                    rb = region.get("bbox") or [0, 0, 0, 0]
                    xs.extend([float(rb[0]), float(rb[2])])
                    ys.extend([float(rb[1]), float(rb[3])])
                cell_bbox = [round(min(xs)), round(min(ys)), round(max(xs)), round(max(ys))]
            else:
                cell_bbox = [
                    x_positions[column_index],
                    round(min(row_y1_values)),
                    x_positions[column_index + 1],
                    round(max(row_y2_values)),
                ]
            cells.append(
                {
                    "page_index": page_index,
                    "table_index": 0,
                    "row_index": row_index,
                    "column_index": column_index,
                    "text": text,
                    "bbox": cell_bbox,
                    "confidence": confidence,
                    "method": "text_region_virtual_grid",
                }
            )
        if any(clean_text(value) for value in row):
            rows.append(row)

    if not rows:
        return None
    rows = _normalize_virtual_grid_rows(rows)
    bbox = [
        round(min(all_xs)) if all_xs else x_positions[0],
        round(min(all_ys)) if all_ys else 0,
        round(max(all_xs)) if all_xs else x_positions[-1],
        round(max(all_ys)) if all_ys else 0,
    ]
    return {
        "page_index": page_index,
        "table_index": 0,
        "table_type": _table_type(rows),
        "bbox": bbox,
        "confidence": 0.82,
        "raw_rows": rows,
        "cells": cells,
        "method": "text_region_virtual_grid",
    }


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


def _looks_like_structured_detail_row(row: list[str]) -> bool:
    values = [clean_text(value) for value in row]
    if not any(values):
        return False
    first = values[0] if values else ""
    leading = " ".join(values[:4])
    has_sequence = bool(re.fullmatch(r"\d+(?:\.\d+)?", first))
    has_code = bool(re.search(r"[A-Za-z]{1,8}\d{2,}|\d{3,}[A-Za-z]", leading))
    has_numeric_code = bool(re.search(r"\b\d{5,}\b", leading))
    numeric_tokens = re.findall(r"(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", " ".join(values))
    return (has_sequence or has_code or has_numeric_code) and len(numeric_tokens) >= 2


def segment_purchase_page(page: dict[str, Any]) -> dict[str, Any]:
    image_path = Path(page.get("clean_image_path") or page["image_path"])
    page_index = int(page.get("page_index", 0))
    grids = _cached_grids(image_path)
    tables: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []

    has_clear_grid = any(int(grid.get("column_count") or 0) >= 6 and int(grid.get("row_count") or 0) >= 2 for grid in grids)

    regions: list[dict[str, Any]] = []
    if not has_clear_grid:
        try:
            regions = _cached_ocr_regions(image_path)
        except Exception as exc:
            issues.append(
                {
                    "page_index": page_index,
                    "region": "\u9875\u9762 OCR",
                    "field": "",
                    "raw_value": "",
                    "clean_value": "",
                    "confidence": 0,
                    "message": f"\u9875\u9762\u6587\u672c OCR \u5931\u8d25\uff1a{exc}",
                }
            )

    for grid in grids:
        grid_regions = regions
        max_cell_ocr_fallback = None
        if has_clear_grid and int(grid.get("column_count") or 0) >= 6:
            grid_regions = _cached_table_ocr_regions(image_path, grid)
            max_cell_ocr_fallback = 0
        cell_rows, cell_cells = _ocr_grid_rows(image_path, grid, page_index, grid_regions, max_cell_ocr_fallback=max_cell_ocr_fallback)
        region_rows, region_cells = _rows_from_region_grid(grid_regions, grid, page_index)
        if _region_rows_are_better(region_rows, cell_rows, grid):
            rows, cells, method = region_rows, region_cells, "page_ocr_row_cluster"
        else:
            rows, cells, method = cell_rows, cell_cells, "grid_cell_ocr"
        recovery_actions = _repair_uncertain_detail_headers(image_path, rows, cells)
        if method == "page_ocr_row_cluster":
            recovery_actions.extend(_repair_sparse_detail_cells(image_path, rows, cells))
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
                "recovery_actions": recovery_actions,
            }
        )

    inferred_table = None
    if not tables and not regions:
        try:
            regions = _cached_ocr_regions(image_path)
        except Exception as exc:
            issues.append(
                {
                    "page_index": page_index,
                    "region": "\u9875\u9762 OCR",
                    "field": "",
                    "raw_value": "",
                    "clean_value": "",
                    "confidence": 0,
                    "message": f"\u9875\u9762\u6587\u672c OCR \u5931\u8d25\uff1a{exc}",
                }
            )
    if not tables and regions:
        inferred_table = _infer_text_table_from_regions(regions, page_index)
        if inferred_table:
            tables.append(inferred_table)

    table_areas = list(grids)
    if inferred_table:
        table_areas.append({"bbox": inferred_table.get("bbox") or []})
    non_table_regions = [region for region in regions if not _inside_any_table(region, table_areas)]
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


def segment_purchase_page_with_layout(page: dict[str, Any], layout: dict[str, Any]) -> dict[str, Any]:
    image_path = Path(page.get("clean_image_path") or page["image_path"])
    page_index = int(page.get("page_index", 0))
    width = int(page.get("width") or 0)
    height = int(page.get("height") or 0)
    from .purchase_layout_cache import cache_bbox_to_page

    bbox = cache_bbox_to_page(layout, width, height)
    if not bbox:
        return segment_purchase_page(page)

    issues: list[dict[str, Any]] = []
    try:
        regions = _cached_ocr_regions(image_path)
    except Exception as exc:
        regions = []
        issues.append(
            {
                "page_index": page_index,
                "region": "\u9875\u9762 OCR",
                "field": "",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": f"\u7f13\u5b58\u7248\u5f0f\u5feb\u901f OCR \u5931\u8d25\uff1a{exc}",
            }
        )

    grid = {
        "table_index": 0,
        "bbox": bbox,
        "x_positions": [bbox[0], bbox[2]],
        "y_positions": [bbox[1], bbox[3]],
        "confidence": 0.9,
    }
    table_regions = _regions_in_bbox(regions, bbox)
    cached_method = str(layout.get("method") or "")
    if "text_region_virtual_grid" in cached_method:
        inferred_table = _infer_text_table_from_regions(table_regions, page_index)
        if inferred_table:
            inferred_table["bbox"] = bbox
            inferred_table["confidence"] = max(float(inferred_table.get("confidence") or 0), 0.9)
            inferred_table["method"] = "layout_cache_text_region_virtual_grid"
            non_table_regions = [region for region in regions if not _inside_any_table(region, [grid])]
            non_table_regions.sort(key=lambda item: ((item.get("bbox") or [0, 0, 0, 0])[1], (item.get("bbox") or [0, 0, 0, 0])[0]))
            text_lines = [clean_text(region.get("text")) for region in non_table_regions if clean_text(region.get("text"))]
            sections = {"\u5907\u6ce8": [], "\u6761\u6b3e": [], "\u4ed8\u6b3e\u4fe1\u606f": [], "\u6536\u8d27\u4fe1\u606f": [], "\u7b7e\u6838\u533a": []}
            for line in text_lines:
                section = classify_section_line(line)
                if section:
                    sections.setdefault(section, []).append(line)
            return {
                "page_index": page_index,
                "width": page.get("width"),
                "height": page.get("height"),
                "image_path": str(image_path),
                "tables": [inferred_table],
                "text_regions": non_table_regions,
                "text_lines": text_lines,
                "sections": sections,
                "issues": issues,
                "layout_cache_hit": True,
            }

    row_clusters = _cluster_regions_by_row(table_regions)
    raw_rows: list[list[str]] = []
    cells: list[dict[str, Any]] = []
    for row_index, row_regions in enumerate(row_clusters):
        text, confidence = _join_regions(row_regions)
        if not clean_text(text):
            continue
        raw_rows.append([text])
        xs: list[float] = []
        ys: list[float] = []
        for region in row_regions:
            rb = region.get("bbox") or [0, 0, 0, 0]
            xs.extend([float(rb[0]), float(rb[2])])
            ys.extend([float(rb[1]), float(rb[3])])
        cell_bbox = [round(min(xs)), round(min(ys)), round(max(xs)), round(max(ys))] if xs and ys else bbox
        cells.append(
            {
                "page_index": page_index,
                "table_index": 0,
                "row_index": row_index,
                "column_index": 0,
                "text": text,
                "bbox": cell_bbox,
                "confidence": confidence,
                "method": "layout_cache_row_cluster",
            }
        )

    if not raw_rows:
        issues.append(
            {
                "page_index": page_index,
                "region": "\u660e\u7ec6\u8868",
                "field": "",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": "\u7f13\u5b58\u7248\u5f0f\u547d\u4e2d\uff0c\u4f46\u660e\u7ec6\u533a\u57df\u672a\u8bc6\u522b\u5230\u6587\u672c\uff0c\u5df2\u56de\u9000\u5b8c\u6574\u8bc6\u522b\u3002",
            }
        )
        return segment_purchase_page(page)

    non_table_regions = [region for region in regions if not _inside_any_table(region, [grid])]
    non_table_regions.sort(key=lambda item: ((item.get("bbox") or [0, 0, 0, 0])[1], (item.get("bbox") or [0, 0, 0, 0])[0]))
    text_lines = [clean_text(region.get("text")) for region in non_table_regions if clean_text(region.get("text"))]
    sections = {"\u5907\u6ce8": [], "\u6761\u6b3e": [], "\u4ed8\u6b3e\u4fe1\u606f": [], "\u6536\u8d27\u4fe1\u606f": [], "\u7b7e\u6838\u533a": []}
    for line in text_lines:
        section = classify_section_line(line)
        if section:
            sections.setdefault(section, []).append(line)

    return {
        "page_index": page_index,
        "width": page.get("width"),
        "height": page.get("height"),
        "image_path": str(image_path),
        "tables": [
            {
                "page_index": page_index,
                "table_index": 0,
                "table_type": "detail_table",
                "bbox": bbox,
                "confidence": 0.9,
                "raw_rows": raw_rows,
                "cells": cells,
                "method": f"layout_cache_{layout.get('method') or 'row_cluster'}",
            }
        ],
        "text_regions": non_table_regions,
        "text_lines": text_lines,
        "sections": sections,
        "issues": issues,
        "layout_cache_hit": True,
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
        if header_index is not None and table.get("method") != "docling_markdown" and not _looks_like_structured_detail_row(row):
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
