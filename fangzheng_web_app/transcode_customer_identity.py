from __future__ import annotations

import re
import unicodedata
from typing import Any


CUSTOMER_ALIAS_GROUPS = (
    frozenset({"广东依顿", "广州伊顿"}),
)


def normalize_customer_name(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    return re.sub(r"\s+", "", text)


def customer_names_match(left: Any, right: Any) -> bool:
    left_name = normalize_customer_name(left)
    right_name = normalize_customer_name(right)
    if not left_name or not right_name:
        return False
    if left_name in right_name or right_name in left_name:
        return True
    return any(
        left_name in {normalize_customer_name(item) for item in group}
        and right_name in {normalize_customer_name(item) for item in group}
        for group in CUSTOMER_ALIAS_GROUPS
    )


def is_yidun_customer(value: Any) -> bool:
    return any(customer_names_match(value, alias) for alias in CUSTOMER_ALIAS_GROUPS[0])
