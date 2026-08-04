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
    {
        "key": "mingyang",
        "label": "明阳",
        "enabled": True,
        "rule_label": "明阳报价表",
        "test_label": "明阳测试数据",
    },
    {
        "key": "lejian",
        "label": "乐健",
        "enabled": True,
        "rule_label": "乐健报价表",
        "test_label": "乐健测试数据",
    },
    {
        "key": "guanghe",
        "label": "广合",
        "enabled": True,
        "rule_label": "广合报价表",
        "test_label": "广合测试数据",
    },
    {
        "key": "shengyi",
        "label": "生益",
        "enabled": True,
        "rule_label": "生益报价表",
        "test_label": "生益测试数据",
    },
    {
        "key": "techuang",
        "label": "特创",
        "enabled": True,
        "rule_label": "特创报价表",
        "test_label": "特创测试数据",
    },
    {
        "key": "zhongfu",
        "label": "中富",
        "enabled": True,
        "rule_label": "中富报价表",
        "test_label": "中富测试数据",
    },
    {
        "key": "huaxingyu",
        "label": "华兴宇",
        "enabled": True,
        "rule_label": "华兴宇报价表",
        "test_label": "华兴宇测试数据",
    },
    {
        "key": "dongxun",
        "label": "东讯",
        "enabled": True,
        "rule_label": "东讯报价表",
        "test_label": "东讯测试数据",
    },
    {
        "key": "suhang",
        "label": "苏杭",
        "enabled": True,
        "rule_label": "苏杭报价表",
        "test_label": "苏杭测试数据",
    },
    {
        "key": "yingchuangli",
        "label": "英创力",
        "enabled": True,
        "rule_label": "英创力报价表",
        "test_label": "英创力测试数据",
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
