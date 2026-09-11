from __future__ import annotations

"""Resolve customer master data to the applicable price-calculation rule."""

from typing import Any, Mapping


PRICE_TAX_MODE_UNKNOWN = "unknown"
PRICE_TAX_MODE_INCLUSIVE = "tax_inclusive"
PRICE_TAX_MODE_EXCLUSIVE = "tax_exclusive"
PRICE_TAX_MODE_LABELS = {
    PRICE_TAX_MODE_UNKNOWN: "未知",
    PRICE_TAX_MODE_INCLUSIVE: "含税",
    PRICE_TAX_MODE_EXCLUSIVE: "未税",
}

PRICE_CALCULATION_ASSOCIATIONS = (
    {"key": "fangzheng", "label": "方正", "group_names": ("方正集团",), "short_names": ("方正",)},
    {"key": "bomin", "label": "博敏", "group_names": ("博敏集团",), "short_names": ("博敏",)},
    {"key": "shennan", "label": "深南", "group_names": ("深南集团",), "short_names": ("深南",)},
    {"key": "hushi", "label": "沪士", "group_names": ("沪士集团",), "short_names": ("沪士",)},
    {"key": "jingwang", "label": "景旺", "group_names": ("景旺集团",), "short_names": ("景旺",)},
    {"key": "plin", "label": "普林", "group_names": ("天津普林",), "short_names": ("普林",)},
    {"key": "hanyu", "label": "瀚宇", "group_names": ("江苏瀚宇",), "short_names": ("瀚宇",)},
    {"key": "wutong", "label": "吴通", "group_names": ("苏州吴通",), "short_names": ("吴通",)},
    {"key": "eaton", "label": "依顿", "group_names": ("广东依顿",), "short_names": ("依顿", "伊顿")},
    {"key": "taixing", "label": "泰兴", "group_names": ("泰兴电路",), "short_names": ("泰兴",)},
    {"key": "aoshikang", "label": "奥士康", "group_names": ("ASK集团",), "short_names": ("奥士康",)},
    {"key": "mingyang", "label": "明阳", "group_names": ("明阳集团",), "short_names": ("明阳",)},
    {"key": "lejian", "label": "乐健", "group_names": ("乐健集团",), "short_names": ("乐健",)},
    {"key": "guanghe", "label": "广合", "group_names": ("广合集团",), "short_names": ("广合",)},
    {"key": "shengyi", "label": "生益", "group_names": ("生益集团",), "short_names": ("生益",)},
    {"key": "guigu", "label": "硅谷", "group_names": ("宜兴硅谷",), "short_names": ("硅谷",)},
    {"key": "techuang", "label": "特创", "group_names": ("特创集团",), "short_names": ("特创",)},
    {"key": "zhongfu", "label": "中富", "group_names": ("中富集团",), "short_names": ("中富",)},
    {"key": "huaxingyu", "label": "华兴宇", "group_names": ("川华兴宇",), "short_names": ("华兴宇",)},
    {"key": "dongxun", "label": "东讯", "group_names": ("东讯集团",), "short_names": ("东讯",)},
    {"key": "suhang", "label": "苏杭", "group_names": ("苏杭集团",), "short_names": ("苏杭",)},
    {"key": "yingchuangli", "label": "英创力", "group_names": ("川英创力",), "short_names": ("英创力",)},
    {"key": "zhongjing", "label": "中京", "group_names": ("中京集团",), "short_names": ("中京",)},
    {"key": "kexiang", "label": "科翔", "group_names": ("科翔集团",), "short_names": ("科翔",)},
    {"key": "junya", "label": "骏亚", "group_names": ("骏亚集团",), "short_names": ("骏亚",)},
    {"key": "chaoying", "label": "超颖", "group_names": ("定颖集团",), "short_names": ("超颖",)},
)

PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY = {
    **{
        association["key"]: PRICE_TAX_MODE_UNKNOWN
        for association in PRICE_CALCULATION_ASSOCIATIONS
    },
    "bomin": PRICE_TAX_MODE_INCLUSIVE,
    "jingwang": PRICE_TAX_MODE_INCLUSIVE,
    "wutong": PRICE_TAX_MODE_INCLUSIVE,
    "eaton": PRICE_TAX_MODE_INCLUSIVE,
    "aoshikang": PRICE_TAX_MODE_INCLUSIVE,
    "mingyang": PRICE_TAX_MODE_INCLUSIVE,
    "lejian": PRICE_TAX_MODE_INCLUSIVE,
    "shengyi": PRICE_TAX_MODE_INCLUSIVE,
    "guigu": PRICE_TAX_MODE_INCLUSIVE,
    "techuang": PRICE_TAX_MODE_INCLUSIVE,
    "zhongfu": PRICE_TAX_MODE_INCLUSIVE,
    "dongxun": PRICE_TAX_MODE_INCLUSIVE,
    "huaxingyu": PRICE_TAX_MODE_INCLUSIVE,
    "suhang": PRICE_TAX_MODE_INCLUSIVE,
    "zhongjing": PRICE_TAX_MODE_INCLUSIVE,
    "junya": PRICE_TAX_MODE_INCLUSIVE,
    "kexiang": PRICE_TAX_MODE_INCLUSIVE,
    "fangzheng": PRICE_TAX_MODE_EXCLUSIVE,
    "plin": PRICE_TAX_MODE_EXCLUSIVE,
    "hanyu": PRICE_TAX_MODE_EXCLUSIVE,
    "guanghe": PRICE_TAX_MODE_EXCLUSIVE,
    "hushi": PRICE_TAX_MODE_EXCLUSIVE,
    "chaoying": PRICE_TAX_MODE_EXCLUSIVE,
}


def _text(value: Any) -> str:
    return str(value or "").strip()


def _result(association: Mapping[str, Any], source: str, value: str) -> dict[str, str | bool]:
    customer_key = _text(association["key"])
    tax_mode = PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY.get(customer_key, PRICE_TAX_MODE_UNKNOWN)
    return {
        "matched": True,
        "price_customer_key": customer_key,
        "price_customer_label": _text(association["label"]),
        "price_tax_mode": tax_mode,
        "price_tax_mode_label": PRICE_TAX_MODE_LABELS.get(tax_mode, PRICE_TAX_MODE_LABELS[PRICE_TAX_MODE_UNKNOWN]),
        "match_source": source,
        "match_value": value,
        "reason": f"{source}“{value}”匹配{association['label']}价格规则",
    }


def resolve_customer_price_calculation(customer: Mapping[str, Any] | None) -> dict[str, str | bool]:
    """Resolve a customer record without mutating customer master data."""
    if not customer:
        return {
            "matched": False,
            "price_customer_key": "",
            "price_customer_label": "",
            "price_tax_mode": PRICE_TAX_MODE_UNKNOWN,
            "price_tax_mode_label": PRICE_TAX_MODE_LABELS[PRICE_TAX_MODE_UNKNOWN],
            "match_source": "",
            "match_value": "",
            "reason": "客户档案不存在，无法关联价格规则",
        }
    short_name = _text(customer.get("customer_short_name"))
    group_name = _text(customer.get("group_name"))
    if short_name:
        for association in PRICE_CALCULATION_ASSOCIATIONS:
            if short_name in association["short_names"]:
                return _result(association, "客户简称", short_name)
    if group_name:
        for association in PRICE_CALCULATION_ASSOCIATIONS:
            if group_name in association["group_names"]:
                return _result(association, "所属集团", group_name)
    return {
        "matched": False,
        "price_customer_key": "",
        "price_customer_label": "",
        "price_tax_mode": PRICE_TAX_MODE_UNKNOWN,
        "price_tax_mode_label": PRICE_TAX_MODE_LABELS[PRICE_TAX_MODE_UNKNOWN],
        "match_source": "",
        "match_value": "",
        "reason": "客户简称和所属集团均未配置价格计算关联",
    }


def resolve_customer_price_calculation_by_id(customer_id: int | None) -> dict[str, str | bool]:
    from .customer_archive_service import get_customer

    return resolve_customer_price_calculation(get_customer(int(customer_id)) if customer_id else None)
