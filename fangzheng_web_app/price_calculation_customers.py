from __future__ import annotations


PRICE_CALCULATION_CUSTOMERS = [
    {
        "key": "jingwang",
        "label": "景旺",
        "enabled": True,
        "rule_label": "景旺报价表",
        "test_label": "景旺测试数据",
    },
    {
        "key": "plin",
        "label": "普林",
        "enabled": True,
        "rule_label": "普林报价表",
        "test_label": "普林测试数据",
    },
    {
        "key": "hanyu",
        "label": "瀚宇",
        "enabled": True,
        "rule_label": "瀚宇报价表",
        "test_label": "瀚宇测试数据",
    },
    {
        "key": "wutong",
        "label": "吴通",
        "enabled": True,
        "rule_label": "吴通报价表",
        "test_label": "吴通测试数据",
    },
    {
        "key": "eaton",
        "label": "伊顿",
        "enabled": True,
        "rule_label": "伊顿报价表",
        "test_label": "伊顿测试数据",
    },
    {
        "key": "taixing",
        "label": "泰兴",
        "enabled": True,
        "rule_label": "泰兴报价表",
        "test_label": "泰兴测试数据",
    },
    {
        "key": "aoshikang",
        "label": "奥士康",
        "enabled": True,
        "rule_label": "奥士康报价表",
        "test_label": "奥士康测试数据",
    },
    {"key": "shenghong", "label": "胜宏", "enabled": False, "rule_label": "胜宏报价表", "test_label": "胜宏测试数据"},
    {"key": "bomin", "label": "博敏", "enabled": False, "rule_label": "博敏报价表", "test_label": "博敏测试数据"},
]


def get_price_customer(customer_key: str | None) -> dict:
    key = (customer_key or "").strip().lower()
    for customer in PRICE_CALCULATION_CUSTOMERS:
        if customer["key"] == key:
            return customer
    raise ValueError("未知价格计算客户")


def enabled_price_customer(customer_key: str | None) -> dict:
    customer = get_price_customer(customer_key)
    if not customer.get("enabled"):
        raise ValueError(f"{customer['label']}价格计算暂未接入")
    return customer


def default_price_customer_key() -> str:
    for customer in PRICE_CALCULATION_CUSTOMERS:
        if customer.get("enabled"):
            return customer["key"]
    return PRICE_CALCULATION_CUSTOMERS[0]["key"]
