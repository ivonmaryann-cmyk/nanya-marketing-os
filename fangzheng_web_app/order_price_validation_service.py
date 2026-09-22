from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Mapping

from .bomin_service import calculate_bomin_quote
from .calculator_service import calculate_fangzheng_quote
from .customer_archive_service import get_customer
from .price_calculation_customer_mapping import (
    PRICE_TAX_MODE_EXCLUSIVE,
    PRICE_TAX_MODE_INCLUSIVE,
    resolve_customer_price_calculation,
)
from .price_calculation_service import calculate_price_quote
from .shennan_service import calculate_shennan_quote
from .hushi_service import calculate_hushi_quote


TARGET_FIELD_BY_TAX_MODE = {
    PRICE_TAX_MODE_INCLUSIVE: ("unit_price", "单价"),
    PRICE_TAX_MODE_EXCLUSIVE: ("price_before_tax", "税前单价"),
}
PRICE_REVIEW_SNAPSHOT_KEY = "_price_review"


class PriceMismatchConfirmationRequired(ValueError):
    def __init__(self, review: dict[str, Any]) -> None:
        self.review = review
        super().__init__(f"发现 {len(review['mismatches'])} 项价格与报价单不一致，请确认后继续录单")


def _decimal(value: Any, *, places: Decimal = Decimal("0.01")) -> Decimal | None:
    text = str(value or "").strip().replace(",", "")
    if not text:
        return None
    try:
        return Decimal(text).quantize(places, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def _display_price(value: Decimal, *, places: Decimal = Decimal("0.01")) -> str:
    decimals = max(-places.as_tuple().exponent, 0)
    return format(value, f".{decimals}f")


def _quote_places(values: Mapping[str, Any]) -> Decimal:
    remark = str(values.get("remark") or "")
    spec = str(values.get("customer_spec") or "")
    is_roll = "卷" in remark or bool(re.search(r"\d+(?:\.\d+)?\s*m\s*/\s*roll\b", spec, re.IGNORECASE))
    return Decimal("0.0001") if is_roll else Decimal("0.01")


def _empty_review() -> dict[str, Any]:
    return {
        "association": {},
        "tax_mode": "unknown",
        "target_field": "",
        "target_label": "",
        "by_line": {},
        "mismatches": [],
    }


def _calculate_quote(customer_key: str, spec: str, quantity: Any) -> dict[str, Any]:
    if customer_key == "fangzheng":
        return calculate_fangzheng_quote(spec)
    if customer_key == "bomin":
        return calculate_bomin_quote(spec)
    if customer_key == "shennan":
        return calculate_shennan_quote(spec)
    if customer_key == "hushi":
        return calculate_hushi_quote(spec)
    return calculate_price_quote(customer_key, spec, quantity=quantity)


def review_template_prices(
    customer: Mapping[str, Any] | None, lines: list[Mapping[str, Any]],
) -> dict[str, Any]:
    """Review saved template prices without mutating the template or its rules."""
    association = resolve_customer_price_calculation(customer)
    tax_mode = str(association.get("price_tax_mode") or "unknown")
    target = TARGET_FIELD_BY_TAX_MODE.get(tax_mode)
    review = _empty_review()
    review.update({
        "association": association,
        "tax_mode": tax_mode,
        "target_field": target[0] if target else "",
        "target_label": target[1] if target else "",
    })
    if not association.get("matched") or not target:
        return review

    target_field, target_label = target
    for line in lines:
        values = line.get("values") if isinstance(line.get("values"), Mapping) else line
        values = values if isinstance(values, Mapping) else {}
        line_no = int(line.get("line_no") or values.get("line_no") or 0)
        places = _quote_places(values)
        template_price = _decimal(values.get(target_field), places=places)
        item: dict[str, Any] = {
            "line_no": line_no,
            "field": target_field,
            "field_label": target_label,
            "status": "not_checked",
            "template_price": str(values.get(target_field) or "").strip(),
            "quote_price": "",
            "note": "",
        }
        spec = str(values.get("customer_spec") or "").strip()
        if not spec:
            item["note"] = "客户规格为空，未核价"
        else:
            try:
                quote = _calculate_quote(str(association["price_customer_key"]), spec, values.get("quantity"))
                quote_price = _decimal(quote.get("price"), places=places)
                if str(quote.get("status") or "") != "成功" or quote_price is None:
                    item["note"] = str(quote.get("error") or quote.get("note") or "报价未命中")
                else:
                    item["quote_price"] = _display_price(quote_price, places=places)
                    if template_price is None:
                        item["status"] = "suggested"
                        item["note"] = "报价单计算"
                    else:
                        item["note"] = str(quote.get("note") or "")
                        item["status"] = "matched" if template_price == quote_price else "mismatch"
            except Exception as exc:
                item["note"] = f"核价失败：{exc}"
        review["by_line"][line_no] = item
        if item["status"] == "mismatch":
            review["mismatches"].append(item)
    return review


def review_cached_template_prices(
    snapshot: Mapping[str, Any] | None, lines: list[Mapping[str, Any]],
) -> dict[str, Any]:
    """Compare current values with saved quote results without recalculating."""
    if not isinstance(snapshot, Mapping):
        return _empty_review()
    target_field = str(snapshot.get("target_field") or "")
    target_label = str(snapshot.get("target_label") or "")
    cached_by_line = snapshot.get("by_line")
    if not target_field or not isinstance(cached_by_line, Mapping):
        return _empty_review()

    review = _empty_review()
    review.update({
        "association": dict(snapshot.get("association") or {}),
        "tax_mode": str(snapshot.get("tax_mode") or "unknown"),
        "target_field": target_field,
        "target_label": target_label,
    })
    for line in lines:
        values = line.get("values") if isinstance(line.get("values"), Mapping) else line
        values = values if isinstance(values, Mapping) else {}
        line_no = int(line.get("line_no") or values.get("line_no") or 0)
        places = _quote_places(values)
        cached = cached_by_line.get(str(line_no), cached_by_line.get(line_no, {}))
        if not isinstance(cached, Mapping):
            continue
        quote_price = _decimal(cached.get("quote_price"), places=places)
        item = {
            "line_no": line_no,
            "field": target_field,
            "field_label": target_label,
            "status": "not_checked",
            "template_price": str(values.get(target_field) or "").strip(),
            "quote_price": _display_price(quote_price, places=places) if quote_price is not None else "",
            "note": str(cached.get("note") or ""),
        }
        if quote_price is not None:
            template_price = _decimal(values.get(target_field), places=places)
            if template_price is None:
                item["status"] = "suggested"
                item["note"] = "报价单计算"
            else:
                item["status"] = "matched" if template_price == quote_price else "mismatch"
        review["by_line"][line_no] = item
        if item["status"] == "mismatch":
            review["mismatches"].append(item)
    return review


def review_case_template_prices(case: Mapping[str, Any], template: Mapping[str, Any]) -> dict[str, Any]:
    customer_id = case.get("customer_id") if isinstance(case, Mapping) else None
    customer = get_customer(int(customer_id)) if customer_id else None
    lines = template.get("lines") if isinstance(template, Mapping) else []
    return review_template_prices(customer, list(lines or []))
