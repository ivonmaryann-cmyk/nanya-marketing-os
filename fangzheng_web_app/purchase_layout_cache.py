from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from .paths import PDF_EXCEL_LAYOUT_CACHE_DIR
from .purchase_field_rules import clean_text


LAYOUT_CACHE_VERSION = "purchase_layout_cache_v1"


def _compact(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def _safe_token(value: str) -> str:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()[:16]
    return digest


def _page_ratio(width: int, height: int) -> float:
    return round(float(width or 1) / float(height or 1), 3)


def _bbox_to_ratio(bbox: list[int] | tuple[int, int, int, int], width: int, height: int) -> list[float]:
    if not bbox or width <= 0 or height <= 0:
        return []
    x0, y0, x1, y1 = [float(value) for value in bbox[:4]]
    return [round(x0 / width, 5), round(y0 / height, 5), round(x1 / width, 5), round(y1 / height, 5)]


def _bbox_from_ratio(ratio_bbox: list[float], width: int, height: int) -> list[int]:
    if len(ratio_bbox) != 4:
        return []
    return [
        max(0, min(width, round(ratio_bbox[0] * width))),
        max(0, min(height, round(ratio_bbox[1] * height))),
        max(0, min(width, round(ratio_bbox[2] * width))),
        max(0, min(height, round(ratio_bbox[3] * height))),
    ]


def layout_signature(header_info: dict[str, Any], lines: list[str], width: int, height: int) -> str:
    customer = clean_text(header_info.get("客户") or "")
    supplier = clean_text(header_info.get("供应商") or "")
    header_text = " ".join(clean_text(line) for line in lines[:20])
    company_hits = re.findall(r"[\u4e00-\u9fffA-Za-z0-9（）()]{4,40}(?:公司|有限公司|厂)", header_text)
    identity = (company_hits[0] if company_hits else "") or customer or supplier
    if not identity:
        return ""
    text = "|".join([_compact(identity), str(_page_ratio(width, height))])
    return _safe_token(text)


def _cache_path(signature: str) -> Path:
    return PDF_EXCEL_LAYOUT_CACHE_DIR / f"{signature}.json"


def load_layout_cache(header_info: dict[str, Any], lines: list[str], width: int, height: int) -> dict[str, Any] | None:
    signature = layout_signature(header_info, lines, width, height)
    if not signature:
        return None
    path = _cache_path(signature)
    if not path.exists():
        return None
    try:
        cache = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if cache.get("version") != LAYOUT_CACHE_VERSION:
        return None
    if abs(float(cache.get("page_ratio") or 0) - _page_ratio(width, height)) > 0.04:
        return None
    return cache


def save_layout_cache(document: dict[str, Any]) -> bool:
    if document.get("file_type") not in {"image", "pdf"}:
        return False
    page_count = int(document.get("page_count") or 0)
    if page_count != 1:
        return False
    rows = document.get("mapped_detail_rows") or []
    if len(rows) < 2:
        return False
    source_tables = document.get("raw_detail_tables") or []
    if not source_tables:
        return False

    pages = document.get("pages") or []
    page = pages[0] if pages else {}
    width = int(page.get("width") or 0)
    height = int(page.get("height") or 0)
    lines = []
    for page_item in pages:
        lines.extend(page_item.get("text_lines") or [])
    header_info = document.get("header_info") or {}
    signature = layout_signature(header_info, lines, width, height)
    if not signature:
        return False

    table = source_tables[0]
    bbox = table.get("bbox") or []
    bbox_ratio = _bbox_to_ratio(bbox, width, height)
    if not bbox_ratio:
        return False
    method = "packed_text_rebuild" if any("packed_text_rebuild" in str(row.get("method") or "") for row in rows) else str(table.get("method") or "")
    if not method:
        return False

    payload = {
        "version": LAYOUT_CACHE_VERSION,
        "signature": signature,
        "source_file": document.get("source_file", ""),
        "customer": header_info.get("客户", ""),
        "supplier": header_info.get("供应商", ""),
        "page_width": width,
        "page_height": height,
        "page_ratio": _page_ratio(width, height),
        "detail_bbox_ratio": bbox_ratio,
        "method": method,
        "row_count": len(rows),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    PDF_EXCEL_LAYOUT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    _cache_path(signature).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return True


def cache_bbox_to_page(cache: dict[str, Any], width: int, height: int) -> list[int]:
    return _bbox_from_ratio(cache.get("detail_bbox_ratio") or [], width, height)
