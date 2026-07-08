from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}


def _page_record(image_path: Path, page_index: int) -> dict[str, Any]:
    with Image.open(image_path) as image:
        width, height = image.size
    return {
        "page_index": page_index,
        "image_path": str(image_path),
        "width": width,
        "height": height,
    }


def _render_pdf_with_poppler(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    pdftoppm = shutil.which("pdftoppm") or shutil.which("pdftoppm.cmd")
    if not pdftoppm:
        return []
    prefix = output_dir / "page"
    try:
        subprocess.run(
            [pdftoppm, "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError):
        return []
    return sorted(output_dir.glob("page-*.png"))


def _render_pdf_with_pdfium(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    try:
        import pypdfium2 as pdfium
    except Exception as exc:  # pragma: no cover - optional dependency
        raise RuntimeError("PDF 页面渲染失败，且 pypdfium2 不可用。") from exc

    scale = max(float(dpi) / 72.0, 1.0)
    pdf = pdfium.PdfDocument(str(pdf_path))
    image_paths: list[Path] = []
    try:
        for index in range(len(pdf)):
            page = pdf[index]
            bitmap = page.render(scale=scale)
            image = bitmap.to_pil()
            image_path = output_dir / f"page-{index + 1:03d}.png"
            image.save(image_path)
            image_paths.append(image_path)
    finally:
        pdf.close()
    return image_paths


def _normalize_image_input(image_path: Path, output_dir: Path) -> Path:
    output_path = output_dir / "page-001.png"
    with Image.open(image_path) as image:
        image = ImageOps.exif_transpose(image).convert("RGB")
        image.save(output_path)
    return output_path


def render_input_pages(file_item: dict[str, str], output_dir: Path, *, dpi: int = 240) -> list[dict[str, Any]]:
    """Render PDF/image input into normalized page images for the generic pipeline."""
    output_dir.mkdir(parents=True, exist_ok=True)
    input_path = Path(file_item["stored_path"])
    suffix = input_path.suffix.lower()

    if suffix == ".pdf":
        image_paths = _render_pdf_with_poppler(input_path, output_dir, dpi)
        if not image_paths:
            image_paths = _render_pdf_with_pdfium(input_path, output_dir, dpi)
    elif suffix in IMAGE_EXTENSIONS:
        image_paths = [_normalize_image_input(input_path, output_dir)]
    else:
        raise ValueError(f"不支持的文件类型：{file_item.get('original_filename') or input_path.name}")

    return [_page_record(path, index) for index, path in enumerate(image_paths)]
