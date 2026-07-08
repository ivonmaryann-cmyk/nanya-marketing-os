from __future__ import annotations

import hashlib
import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any

from .paths import STORAGE_DIR


DOCLING_CACHE_VERSION = "docling_markdown_no_ocr_accurate_content_v1"
DOCLING_CACHE_DIR = STORAGE_DIR / "pdf_excel_cache" / "docling"


def _file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cache_key(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(_file_digest(path).encode("ascii"))
    digest.update(DOCLING_CACHE_VERSION.encode("ascii"))
    return digest.hexdigest()


def _read_cache(path: Path) -> dict[str, Any] | None:
    cache_path = DOCLING_CACHE_DIR / f"{_cache_key(path)}.json"
    if not cache_path.exists():
        return None
    try:
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if cached.get("cache_version") != DOCLING_CACHE_VERSION:
        return None
    return {
        "markdown": str(cached.get("markdown") or ""),
        "lines": cached.get("lines") or [],
        "tables": cached.get("tables") or [],
        "error": cached.get("error"),
        "cache_hit": True,
    }


def _write_cache(path: Path, result: dict[str, Any]) -> None:
    if result.get("error") or not result.get("markdown"):
        return
    try:
        DOCLING_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_path = DOCLING_CACHE_DIR / f"{_cache_key(path)}.json"
        payload = {
            "cache_version": DOCLING_CACHE_VERSION,
            "source_name": path.name,
            "markdown": result.get("markdown") or "",
            "lines": result.get("lines") or [],
            "tables": result.get("tables") or [],
            "error": result.get("error"),
        }
        cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception:
        return


@lru_cache(maxsize=1)
def _converter():
    try:
        from docling.datamodel.base_models import InputFormat
        from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode
        from docling.document_converter import DocumentConverter
        from docling.document_converter import PdfFormatOption
    except Exception as exc:  # pragma: no cover - optional dependency
        raise RuntimeError(f"Docling 不可用：{exc}") from exc
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_ocr = False
    pipeline_options.force_backend_text = True
    pipeline_options.do_table_structure = True
    pipeline_options.generate_page_images = False
    pipeline_options.generate_picture_images = False
    pipeline_options.table_structure_options.mode = TableFormerMode.ACCURATE
    return DocumentConverter(
        allowed_formats=[InputFormat.PDF],
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)},
    )


def pdf_to_markdown(path: Path) -> tuple[str, str | None]:
    result = parse_pdf_with_docling(path)
    return str(result.get("markdown") or ""), result.get("error")


def parse_pdf_with_docling(path: Path) -> dict[str, Any]:
    cached = _read_cache(path)
    if cached is not None:
        return cached
    try:
        result = _converter().convert(str(path))
        document = result.document
        markdown = ""
        if hasattr(document, "export_to_markdown"):
            markdown = str(document.export_to_markdown() or "")
        elif hasattr(document, "export_to_text"):
            markdown = str(document.export_to_text() or "")
        else:
            return {"markdown": "", "lines": [], "tables": [], "error": "Docling 未返回可导出的 Markdown/Text。"}
        parsed = {
            "markdown": markdown,
            "lines": _markdown_lines(markdown),
            "tables": _markdown_tables(markdown),
            "error": None,
            "cache_hit": False,
        }
        _write_cache(path, parsed)
        return parsed
    except Exception as exc:
        return {"markdown": "", "lines": [], "tables": [], "error": str(exc), "cache_hit": False}


def _markdown_lines(markdown: str) -> list[str]:
    lines: list[str] = []
    in_code_block = False
    for raw_line in str(markdown or "").splitlines():
        line = raw_line.strip()
        if line.startswith("```"):
            in_code_block = not in_code_block
            continue
        if in_code_block or not line:
            continue
        if _looks_like_table_line(line) or _is_separator_row(_split_markdown_row(line)):
            continue
        text = line.strip(" #*\t")
        if text:
            lines.append(text)
    return lines


def _markdown_tables(markdown: str) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    block: list[str] = []
    title = ""
    previous_text = ""
    in_code_block = False

    def flush() -> None:
        nonlocal block, title
        rows = _rows_from_table_block(block)
        if rows:
            tables.append(
                {
                    "table_index": len(tables),
                    "title": title or f"Docling 表 {len(tables) + 1}",
                    "rows": rows,
                    "method": "docling_markdown",
                }
            )
        block = []
        title = ""

    for raw_line in str(markdown or "").splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()
        if stripped.startswith("```"):
            if block:
                flush()
            in_code_block = not in_code_block
            continue
        if in_code_block:
            continue
        if _looks_like_table_line(stripped):
            if not block:
                title = previous_text
            block.append(stripped)
            continue
        if block:
            flush()
        if stripped:
            previous_text = stripped.strip(" #*\t")
    if block:
        flush()
    return tables


def _looks_like_table_line(line: str) -> bool:
    return line.count("|") >= 2 and bool(line.strip(" |"))


def _split_markdown_row(line: str) -> list[str]:
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    cells: list[str] = []
    current: list[str] = []
    escaped = False
    for char in stripped:
        if escaped:
            current.append(char)
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == "|":
            cells.append("".join(current).strip())
            current = []
            continue
        current.append(char)
    cells.append("".join(current).strip())
    return [cell.replace("\\|", "|").strip() for cell in cells]


def _is_separator_row(cells: list[str]) -> bool:
    if not cells:
        return False
    non_empty = [cell.strip() for cell in cells if cell.strip()]
    if not non_empty:
        return False
    return all(re.fullmatch(r":?-{2,}:?", cell.replace(" ", "")) for cell in non_empty)


def _rows_from_table_block(block: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    max_cols = 0
    for line in block:
        cells = _split_markdown_row(line)
        if _is_separator_row(cells):
            continue
        if not any(cells):
            continue
        rows.append(cells)
        max_cols = max(max_cols, len(cells))
    if len(rows) < 2 or max_cols < 2:
        return []
    return [(row + [""] * max_cols)[:max_cols] for row in rows]
