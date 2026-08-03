from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_confirmation_policy import (
    load_confirmation_policy_rules,
    match_confirmation_policy_rules,
)


def main() -> None:
    rules = load_confirmation_policy_rules()
    assert len(rules) == 2, rules
    assert not any(
        token in str(rule)
        for rule in rules
        for token in ("历史正确码", "测试正确码", "结果对比", "人工答案")
    ), rules
    assert not match_confirmation_policy_rules("惠州智恩", "NY-A2 0.25mm")
    assert not match_confirmation_policy_rules("广东依顿", "NY2150H ANTI-CAF")
    assert not match_confirmation_policy_rules("广东依顿", "NY2170 16±1.5MIL ANTI-CAF")
    assert not match_confirmation_policy_rules("湖奥士康", "NY2150 耐CAF")
    mixed = match_confirmation_policy_rules("广东依顿", "NY2150H 1/2OZ", "HW订单")
    assert [item["rule_id"] for item in mixed] == ["CPR-YIDUN-HW-MIXED-COPPER"], mixed
    conflict = match_confirmation_policy_rules("广州伊顿", "NY2150H 2/2OZ", "华为 下汽车板")
    assert [item["rule_id"] for item in conflict] == ["CPR-YIDUN-AUTO-HW-CONFLICT"], conflict
    print("unified confirmation policy rules smoke passed")


if __name__ == "__main__":
    main()
