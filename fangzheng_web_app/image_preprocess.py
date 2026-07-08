from __future__ import annotations

from pathlib import Path
from typing import Any

from PIL import Image, ImageOps


def _load_cv2():
    try:
        import cv2
        import numpy as np
    except Exception as exc:  # pragma: no cover - optional dependency
        raise RuntimeError("未安装 OpenCV，无法进行图片预处理。") from exc
    return cv2, np


def preprocess_page_image(page: dict[str, Any], output_dir: Path) -> dict[str, Any]:
    """Create a stable high-contrast page image while preserving the original page."""
    output_dir.mkdir(parents=True, exist_ok=True)
    source_path = Path(page["image_path"])
    output_path = output_dir / f"page-{int(page.get('page_index', 0)) + 1:03d}-clean.png"

    with Image.open(source_path) as image:
        image = ImageOps.exif_transpose(image).convert("RGB")
        image = ImageOps.autocontrast(image)
        image.save(output_path)

    return {
        **page,
        "clean_image_path": str(output_path),
    }


def threshold_for_lines(image_path: Path):
    """Return an inverted binary image suitable for line morphology."""
    cv2, np = _load_cv2()
    image_bytes = np.fromfile(str(image_path), dtype=np.uint8)
    image = cv2.imdecode(image_bytes, cv2.IMREAD_COLOR)
    if image is None:
        raise RuntimeError(f"无法读取图片：{image_path}")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 25, -2)
    return image, binary
