from __future__ import annotations

from pathlib import Path
from typing import Any

from PIL import Image

from .cell_ocr import ocr_cell, ocr_image_regions


def _load_cv2():
    try:
        import cv2
        import numpy as np
    except Exception as exc:  # pragma: no cover - depends on optional runtime package
        raise RuntimeError("未安装 OpenCV 表格检测组件，请先安装 requirements.txt 中的 opencv-python-headless。") from exc
    return cv2, np


def _dedupe_positions(values: list[int], tolerance: int = 8) -> list[int]:
    if not values:
        return []
    values = sorted(values)
    groups: list[list[int]] = [[values[0]]]
    for value in values[1:]:
        if abs(value - groups[-1][-1]) <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [int(sum(group) / len(group)) for group in groups]


def _detect_grid(image_path: Path) -> tuple[list[int], list[int]]:
    cv2, np = _load_cv2()
    image_bytes = np.fromfile(str(image_path), dtype=np.uint8)
    image = cv2.imdecode(image_bytes, cv2.IMREAD_COLOR)
    if image is None:
        raise RuntimeError(f"无法读取图片：{image_path}")

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 21, -2)
    height, width = binary.shape[:2]

    horizontal = binary.copy()
    vertical = binary.copy()
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(width // 32, 18), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(height // 32, 18)))
    horizontal = cv2.dilate(cv2.erode(horizontal, horizontal_kernel), horizontal_kernel)
    vertical = cv2.dilate(cv2.erode(vertical, vertical_kernel), vertical_kernel)

    horizontal_projection = np.sum(horizontal > 0, axis=1)
    vertical_projection = np.sum(vertical > 0, axis=0)
    y_candidates = [int(idx) for idx, value in enumerate(horizontal_projection) if value > width * 0.35]
    x_candidates = [int(idx) for idx, value in enumerate(vertical_projection) if value > height * 0.18]
    return _dedupe_positions(x_candidates), _dedupe_positions(y_candidates)


def _is_blank_crop(image: Image.Image) -> bool:
    cv2, np = _load_cv2()
    gray = np.array(image.convert("L"))
    if gray.size == 0:
        return True
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    ink_ratio = float(np.count_nonzero(binary)) / float(binary.size)
    return ink_ratio < 0.006


def parse_image_tables(image_path: Path, *, page_index: int = 0) -> dict[str, Any]:
    image = Image.open(image_path).convert("RGB")
    words = ocr_image_regions(image)
    x_positions, y_positions = _detect_grid(image_path)
    warnings: list[str] = []
    if len(x_positions) < 2 or len(y_positions) < 2:
        warnings.append("未检测到完整表格线，无法进行单元格局部 OCR。")
        return {
            "page_index": page_index,
            "width": image.width,
            "height": image.height,
            "words": words,
            "tables": [{"table_index": 0, "cells": []}],
            "warnings": warnings,
        }

    cells: list[dict[str, Any]] = []
    padding = 3
    table_index = 0
    for row_index in range(len(y_positions) - 1):
        y0, y1 = y_positions[row_index], y_positions[row_index + 1]
        if y1 - y0 < 10:
            continue
        for column_index in range(len(x_positions) - 1):
            x0, x1 = x_positions[column_index], x_positions[column_index + 1]
            if x1 - x0 < 10:
                continue
            crop_box = (max(x0 + padding, 0), max(y0 + padding, 0), min(x1 - padding, image.width), min(y1 - padding, image.height))
            crop = image.crop(crop_box)
            if _is_blank_crop(crop):
                result = {"text": "", "confidence": 1.0, "method": "manual_empty"}
            else:
                contained_words = []
                for word in words:
                    wx0, wy0, wx1, wy1 = word["bbox"]
                    center_x = (wx0 + wx1) / 2
                    center_y = (wy0 + wy1) / 2
                    if x0 <= center_x <= x1 and y0 <= center_y <= y1:
                        contained_words.append(word)
                if contained_words:
                    result = {
                        "text": " ".join(word["text"] for word in contained_words),
                        "confidence": sum(float(word.get("confidence") or 0) for word in contained_words) / len(contained_words),
                        "method": "page_ocr",
                    }
                else:
                    result = ocr_cell(crop)
            cells.append(
                {
                    "page_index": page_index,
                    "table_index": table_index,
                    "row_index": row_index,
                    "column_index": column_index,
                    "text": result["text"],
                    "bbox": [x0, y0, x1, y1],
                    "confidence": result["confidence"],
                    "method": result["method"],
                }
            )

    return {
        "page_index": page_index,
        "width": image.width,
        "height": image.height,
        "words": words,
        "tables": [{"table_index": table_index, "cells": cells}],
        "warnings": warnings,
    }
