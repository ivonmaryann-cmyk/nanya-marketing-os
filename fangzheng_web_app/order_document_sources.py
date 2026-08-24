"""Source adapters for the shared purchase-order recognition model.

PDF/image recognition and mail HTML tables have different ways of obtaining
cells, but they must produce the same normalised purchase document before any
factory mapping, validation, or domestic-template export is applied.
"""
from __future__ import annotations

import re
from copy import deepcopy
from html.parser import HTMLParser
from typing import Any

from .purchase_field_rules import clean_text, find_detail_header_row
from .purchase_result_normalizer import normalize_purchase_document


ORDER_SOURCE_ADAPTER_RULES = {
    "version": "order_source_adapters_v1",
    "adapters": {
        "mail_html_table": {
            "label": "邮件正文 HTML 表格",
            # This identifier is also used by customer-specific mapping rules;
            # keeping it stable lets a future page edit the same rule set.
            "source_kind": "mail_html_table",
            "table_selection": "使用共享订单字段规则识别同时包含明细标识和数量、价格、金额或交期的表格。",
            "confidence": 1.0,
        },
        "pdf_image": {
            "label": "PDF/图片订单文件",
            "source_kind": "file_ocr",
            "table_selection": "由 PDF/图片识别管线输出结构化明细表。",
        },
    },
}


def order_source_adapter_catalog() -> dict[str, Any]:
    """Return data-only adapter rules for a future visual rule-maintenance page."""
    return deepcopy(ORDER_SOURCE_ADAPTER_RULES)


class _HtmlOrderTableParser(HTMLParser):
    """Small dependency-free HTML table reader for mail bodies.

    It intentionally keeps source cell text intact. Normalisation happens in
    ``normalize_purchase_document`` together with PDF/image source tables.
    """

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._stack: list[dict[str, Any]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "table":
            self._stack.append({"rows": [], "row": None, "cell": None})
        elif not self._stack:
            return
        elif tag == "tr":
            self._stack[-1]["row"] = []
        elif tag in {"td", "th"} and self._stack[-1].get("row") is not None:
            self._stack[-1]["cell"] = []
        elif tag == "br" and self._stack[-1].get("cell") is not None:
            self._stack[-1]["cell"].append("\n")

    def handle_data(self, data: str) -> None:
        if self._stack and self._stack[-1].get("cell") is not None:
            self._stack[-1]["cell"].append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if not self._stack:
            return
        current = self._stack[-1]
        if tag in {"td", "th"} and current.get("cell") is not None:
            current["row"].append(clean_text("".join(current["cell"])))
            current["cell"] = None
        elif tag == "tr" and current.get("row") is not None:
            if any(clean_text(value) for value in current["row"]):
                current["rows"].append(current["row"])
            current["row"] = None
            current["cell"] = None
        elif tag == "table":
            finished = self._stack.pop()
            rows = finished.get("rows") or []
            if rows:
                self.tables.append(rows)


def parse_mail_html_tables(body_html: str) -> list[list[list[str]]]:
    """Read table cells from a mail body without using OCR or file parsing."""
    if not clean_text(body_html):
        return []
    parser = _HtmlOrderTableParser()
    try:
        parser.feed(body_html)
        parser.close()
    except Exception:
        return []
    return parser.tables


def _table_text(rows: list[list[str]]) -> str:
    return "\n".join("\t".join(clean_text(value) for value in row) for row in rows)


def _header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", clean_text(value).lower())


def _order_number_from_tables(tables: list[list[list[str]]]) -> str:
    aliases = {
        "订单号", "采购单号", "采购订单号", "po单号", "po编号", "pono", "ponumber", "pono",
    }
    for rows in tables:
        header_index, _mapping = find_detail_header_row(rows)
        if header_index is None:
            continue
        headers = rows[header_index]
        for column, header in enumerate(headers):
            if _header_key(header) not in aliases:
                continue
            for row in rows[header_index + 1 :]:
                value = clean_text(row[column] if column < len(row) else "")
                if value:
                    return value
    return ""


def build_mail_html_purchase_document(
    body_html: str,
    body_text: str = "",
    *,
    source_name: str = "邮件正文.html",
) -> dict[str, Any] | None:
    """Adapt HTML order tables into the same normalised purchase document as files.

    The returned document has the usual ``mapped_detail_rows`` shape and may
    therefore be sent unchanged through ``project_factory_document`` and
    ``build_domestic_template_data``.
    """
    tables = parse_mail_html_tables(body_html)
    raw_detail_tables: list[dict[str, Any]] = []
    for table_index, rows in enumerate(tables):
        header_index, _mapping = find_detail_header_row(rows)
        if header_index is None:
            continue
        raw_detail_tables.append(
            {
                "page_index": 0,
                "table_index": table_index,
                "rows": rows,
                "method": "mail_html_table",
                "confidence": ORDER_SOURCE_ADAPTER_RULES["adapters"]["mail_html_table"]["confidence"],
                "recovery_actions": [],
            }
        )
    if not raw_detail_tables:
        return None

    source_text = "\n".join(
        item for item in [clean_text(body_text), *(_table_text(rows) for rows in tables)] if item
    )
    order_number = _order_number_from_tables(tables)
    document: dict[str, Any] = {
        "pipeline_version": "purchase_order_v1",
        "source_file": source_name,
        "file_type": "mail_html",
        "parser_mode": "mail_html_table",
        "template_id": "",
        "template_label": "",
        "page_count": 0,
        "header_info": {"订单号": order_number} if order_number else {},
        "pages": [],
        "raw_detail_tables": raw_detail_tables,
        "mapped_detail_rows": [],
        "issues": [],
        "warnings": [],
        "source_adapter": {
            "kind": "mail_html_table",
            "rule_version": ORDER_SOURCE_ADAPTER_RULES["version"],
            "table_count": len(raw_detail_tables),
        },
    }
    normalized = normalize_purchase_document(document, source_text=source_text)
    if not normalized.get("mapped_detail_rows"):
        return None
    return normalized
