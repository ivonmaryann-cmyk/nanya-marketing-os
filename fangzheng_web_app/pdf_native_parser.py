from __future__ import annotations

import re
from pathlib import Path
from statistics import median
from typing import Any

import pdfplumber


READABLE_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
MOJIBAKE_RE = re.compile(r"[�]|(?:[锟鎿]{1,})")


def text_quality(text: str) -> dict[str, Any]:
    visible = [char for char in text if not char.isspace()]
    if not visible:
        return {"has_text": False, "readable": False, "score": 0.0, "reason": "no_text"}
    cjk_count = len(READABLE_CJK_RE.findall(text))
    bad_count = len(MOJIBAKE_RE.findall(text))
    ascii_count = sum(1 for char in visible if char.isascii() and char.isprintable())
    score = (cjk_count * 1.5 + ascii_count * 0.2) / max(len(visible), 1)
    readable = cjk_count >= 8 and bad_count / max(len(visible), 1) < 0.03
    if cjk_count == 0 and ascii_count > 20 and bad_count == 0:
        readable = True
    reason = "readable" if readable else "text_layer_unreadable"
    return {
        "has_text": True,
        "readable": readable,
        "score": round(score, 4),
        "cjk_count": cjk_count,
        "bad_count": bad_count,
        "reason": reason,
    }


def _cluster_words_to_cells(words: list[dict[str, Any]], *, page_index: int, table_index: int = 0) -> list[dict[str, Any]]:
    if not words:
        return []
    words = sorted(words, key=lambda word: (float(word["top"]), float(word["x0"])))
    heights = [float(word["bottom"]) - float(word["top"]) for word in words]
    tolerance = max(3.0, (median(heights) if heights else 8.0) * 0.65)

    rows: list[list[dict[str, Any]]] = []
    row_tops: list[float] = []
    for word in words:
        top = float(word["top"])
        target_index = None
        for index, row_top in enumerate(row_tops):
            if abs(top - row_top) <= tolerance:
                target_index = index
                break
        if target_index is None:
            row_tops.append(top)
            rows.append([word])
        else:
            rows[target_index].append(word)
            row_tops[target_index] = (row_tops[target_index] + top) / 2

    x_positions: list[float] = []
    for row in rows:
        for word in sorted(row, key=lambda value: float(value["x0"])):
            x = float(word["x0"])
            if not any(abs(x - existing) <= 18 for existing in x_positions):
                x_positions.append(x)
    x_positions.sort()

    cells: list[dict[str, Any]] = []
    for row_index, row in enumerate(rows):
        row_words = sorted(row, key=lambda value: float(value["x0"]))
        for word in row_words:
            x0 = float(word["x0"])
            column_index = min(range(len(x_positions)), key=lambda idx: abs(x_positions[idx] - x0)) if x_positions else 0
            text = str(word.get("text") or "").strip()
            cells.append(
                {
                    "page_index": page_index,
                    "table_index": table_index,
                    "row_index": row_index,
                    "column_index": column_index,
                    "text": text,
                    "bbox": [round(float(word["x0"]), 2), round(float(word["top"]), 2), round(float(word["x1"]), 2), round(float(word["bottom"]), 2)],
                    "confidence": 1.0,
                    "method": "pdf_text" if text else "manual_empty",
                }
            )
    return cells


def _tables_to_cells(tables: list[list[list[str | None]]], *, page_index: int, method: str) -> list[dict[str, Any]]:
    cells: list[dict[str, Any]] = []
    for table_index, table in enumerate(tables):
        for row_index, row in enumerate(table):
            for column_index, value in enumerate(row):
                text = str(value or "").strip()
                cells.append(
                    {
                        "page_index": page_index,
                        "table_index": table_index,
                        "row_index": row_index,
                        "column_index": column_index,
                        "text": text,
                        "bbox": None,
                        "confidence": 1.0,
                        "method": method if text else "manual_empty",
                    }
                )
    return cells


def _page_words(words: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for word in words:
        text = str(word.get("text") or "").strip()
        if not text:
            continue
        result.append(
            {
                "text": text,
                "bbox": [round(float(word["x0"]), 2), round(float(word["top"]), 2), round(float(word["x1"]), 2), round(float(word["bottom"]), 2)],
                "confidence": 1.0,
                "method": "pdf_text",
            }
        )
    return result


def _page_lines(page) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for line in page.lines:
        items.append(
            {
                "bbox": [
                    round(float(line.get("x0", 0)), 2),
                    round(float(line.get("top", line.get("y0", 0))), 2),
                    round(float(line.get("x1", 0)), 2),
                    round(float(line.get("bottom", line.get("y1", 0))), 2),
                ],
                "orientation": "h" if abs(float(line.get("y0", 0)) - float(line.get("y1", 0))) < abs(float(line.get("x0", 0)) - float(line.get("x1", 0))) else "v",
            }
        )
    for rect in page.rects:
        x0 = round(float(rect.get("x0", 0)), 2)
        x1 = round(float(rect.get("x1", 0)), 2)
        top = round(float(rect.get("top", 0)), 2)
        bottom = round(float(rect.get("bottom", 0)), 2)
        if abs(x1 - x0) < 1.5 or abs(bottom - top) < 1.5:
            items.append({"bbox": [x0, top, x1, bottom], "orientation": "v" if abs(x1 - x0) < 1.5 else "h"})
    return items


def parse_pdf_native(path: Path) -> dict[str, Any]:
    pages: list[dict[str, Any]] = []
    all_text_parts: list[str] = []
    warnings: list[str] = []

    with pdfplumber.open(path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            all_text_parts.append(text)
            words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
            table_settings = [
                {"vertical_strategy": "lines", "horizontal_strategy": "text"},
                {"vertical_strategy": "text", "horizontal_strategy": "lines"},
                {"vertical_strategy": "text", "horizontal_strategy": "text"},
            ]
            tables: list[list[list[str | None]]] = []
            for settings in table_settings:
                try:
                    tables = page.extract_tables(table_settings=settings)
                except Exception:
                    tables = []
                if tables:
                    break
            if tables:
                cells = _tables_to_cells(tables, page_index=page_index, method="pdf_text")
            else:
                cells = _cluster_words_to_cells(words, page_index=page_index)
                warnings.append(f"第 {page_index + 1} 页未检测到完整表格，已按文字坐标重建行列。")
            pages.append(
                {
                    "page_index": page_index,
                    "width": float(page.width),
                    "height": float(page.height),
                    "text": text,
                    "words": _page_words(words),
                    "lines": _page_lines(page),
                    "char_count": len(page.chars),
                    "line_count": len(page.lines),
                    "rect_count": len(page.rects),
                    "curve_count": len(page.curves),
                    "tables": [{"table_index": 0, "cells": cells}],
                }
            )

    full_text = "\n".join(all_text_parts)
    quality = text_quality(full_text)
    if quality["has_text"] and not quality["readable"]:
        warnings.append("PDF 存在文字层，但文字可读性偏低；已优先保留原生坐标解析结果，必要时可用局部 OCR 复核。")
    return {
        "page_count": len(pages),
        "text": full_text,
        "text_quality": quality,
        "pages": pages,
        "warnings": warnings,
    }
