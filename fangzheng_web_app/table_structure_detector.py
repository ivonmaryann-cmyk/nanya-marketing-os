from __future__ import annotations

from pathlib import Path
from typing import Any

from .image_preprocess import threshold_for_lines


def _load_cv2():
    try:
        import cv2
        import numpy as np
    except Exception as exc:  # pragma: no cover - optional dependency
        raise RuntimeError("未安装 OpenCV 表格检测组件。") from exc
    return cv2, np


def _dedupe_positions(values: list[int], tolerance: int = 10) -> list[int]:
    if not values:
        return []
    values = sorted(values)
    groups: list[list[int]] = [[values[0]]]
    for value in values[1:]:
        if abs(value - groups[-1][-1]) <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [int(round(sum(group) / len(group))) for group in groups]


def _line_masks(image_path: Path):
    cv2, np = _load_cv2()
    image, binary = threshold_for_lines(image_path)
    height, width = binary.shape[:2]
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(width // 28, 24), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(height // 35, 18)))
    horizontal = cv2.dilate(cv2.erode(binary, horizontal_kernel), horizontal_kernel)
    vertical = cv2.dilate(cv2.erode(binary, vertical_kernel), vertical_kernel)
    return image, horizontal, vertical, cv2.add(horizontal, vertical)


def _positions_from_mask(mask, *, axis: int, threshold_ratio: float, offset: int = 0) -> list[int]:
    cv2, np = _load_cv2()
    length = mask.shape[1] if axis == 1 else mask.shape[0]
    projection = np.sum(mask > 0, axis=axis)
    candidates = [int(index + offset) for index, value in enumerate(projection) if value > length * threshold_ratio]
    return _dedupe_positions(candidates)


def _grid_from_bbox(horizontal, vertical, bbox: tuple[int, int, int, int]) -> dict[str, Any] | None:
    x, y, w, h = bbox
    if w <= 80 or h <= 60:
        return None
    h_roi = horizontal[y : y + h, x : x + w]
    v_roi = vertical[y : y + h, x : x + w]
    y_positions = _positions_from_mask(h_roi, axis=1, threshold_ratio=0.35, offset=y)
    x_positions = _positions_from_mask(v_roi, axis=0, threshold_ratio=0.25, offset=x)
    if len(x_positions) < 2 or len(y_positions) < 2:
        return None
    rows = len(y_positions) - 1
    cols = len(x_positions) - 1
    if rows < 2 or cols < 3:
        return None
    return {
        "bbox": [min(x_positions), min(y_positions), max(x_positions), max(y_positions)],
        "x_positions": x_positions,
        "y_positions": y_positions,
        "row_count": rows,
        "column_count": cols,
        "confidence": round(min(0.98, 0.45 + rows * cols / 180), 4),
    }


def detect_table_grids(image_path: Path) -> list[dict[str, Any]]:
    """Detect bordered table grids and return cell boundary positions."""
    cv2, np = _load_cv2()
    image, horizontal, vertical, line_mask = _line_masks(image_path)
    height, width = line_mask.shape[:2]

    contours, _hierarchy = cv2.findContours(line_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    grids: list[dict[str, Any]] = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w < width * 0.35 or h < height * 0.08:
            continue
        grid = _grid_from_bbox(horizontal, vertical, (x, y, w, h))
        if grid:
            grids.append(grid)

    if not grids:
        grid = _grid_from_bbox(horizontal, vertical, (0, 0, width, height))
        if grid:
            grids.append(grid)

    grids.sort(key=lambda item: (item["bbox"][1], -item["row_count"] * item["column_count"]))
    deduped: list[dict[str, Any]] = []
    for grid in grids:
        gx0, gy0, gx1, gy1 = grid["bbox"]
        duplicate = False
        for existing in deduped:
            ex0, ey0, ex1, ey1 = existing["bbox"]
            ix0, iy0, ix1, iy1 = max(gx0, ex0), max(gy0, ey0), min(gx1, ex1), min(gy1, ey1)
            intersection = max(0, ix1 - ix0) * max(0, iy1 - iy0)
            area = max(1, (gx1 - gx0) * (gy1 - gy0))
            if intersection / area > 0.8:
                duplicate = True
                break
        if not duplicate:
            grid["table_index"] = len(deduped)
            deduped.append(grid)
    return deduped


def iter_grid_cells(grid: dict[str, Any]):
    x_positions = grid.get("x_positions") or []
    y_positions = grid.get("y_positions") or []
    for row_index in range(len(y_positions) - 1):
        y0, y1 = y_positions[row_index], y_positions[row_index + 1]
        if y1 - y0 < 8:
            continue
        for column_index in range(len(x_positions) - 1):
            x0, x1 = x_positions[column_index], x_positions[column_index + 1]
            if x1 - x0 < 8:
                continue
            yield row_index, column_index, [x0, y0, x1, y1]
