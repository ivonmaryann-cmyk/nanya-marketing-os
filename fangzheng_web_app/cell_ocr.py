from __future__ import annotations

from functools import lru_cache
from threading import Lock
from typing import Any

import numpy as np
from PIL import Image, ImageOps


class OcrUnavailableError(RuntimeError):
    pass


_OCR_LOCK = Lock()


@lru_cache(maxsize=1)
def _ocr_engine():
    try:
        from rapidocr_onnxruntime import RapidOCR
    except Exception as exc:  # pragma: no cover - depends on optional runtime package
        raise OcrUnavailableError(
            "未安装离线 OCR 组件 rapidocr-onnxruntime，请先安装 requirements.txt 中的依赖。"
        ) from exc
    return RapidOCR()


def _normalize_crop(image: Image.Image) -> Image.Image:
    gray = ImageOps.grayscale(image)
    bbox = gray.getbbox()
    if bbox:
        gray = gray.crop(bbox)
    width, height = gray.size
    scale = 2 if max(width, height) < 260 else 1
    if scale > 1:
        gray = gray.resize((width * scale, height * scale), Image.Resampling.LANCZOS)
    return ImageOps.autocontrast(gray)


def ocr_cell(image: Image.Image) -> dict[str, Any]:
    normalized = _normalize_crop(image)
    if normalized.width <= 2 or normalized.height <= 2:
        return {"text": "", "confidence": 1.0, "method": "manual_empty"}

    array = np.array(normalized.convert("RGB"))
    with _OCR_LOCK:
        engine = _ocr_engine()
        result, _elapsed = engine(array)
    if not result:
        return {"text": "", "confidence": 0.0, "method": "cell_ocr"}

    parts: list[str] = []
    confidences: list[float] = []
    for item in result:
        if len(item) < 3:
            continue
        text = str(item[1] or "").strip()
        if text:
            parts.append(text)
            try:
                confidences.append(float(item[2]))
            except (TypeError, ValueError):
                pass
    confidence = sum(confidences) / len(confidences) if confidences else 0.0
    return {"text": " ".join(parts).strip(), "confidence": confidence, "method": "cell_ocr"}


def ocr_image_regions(image: Image.Image) -> list[dict[str, Any]]:
    normalized = ImageOps.autocontrast(image.convert("RGB"))
    array = np.array(normalized)
    with _OCR_LOCK:
        engine = _ocr_engine()
        result, _elapsed = engine(array)
    regions: list[dict[str, Any]] = []
    if not result:
        return regions
    for item in result:
        if len(item) < 3:
            continue
        box = item[0]
        text = str(item[1] or "").strip()
        if not text:
            continue
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        try:
            confidence = float(item[2])
        except (TypeError, ValueError):
            confidence = 0.0
        regions.append(
            {
                "text": text,
                "bbox": [round(min(xs), 2), round(min(ys), 2), round(max(xs), 2), round(max(ys), 2)],
                "confidence": confidence,
                "method": "page_ocr",
            }
        )
    return regions
