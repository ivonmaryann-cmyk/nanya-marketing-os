from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import calculate_transcode_agent_quote


def main() -> None:
    cases = [
        (
            "AY专用条件",
            "广东依顿",
            "覆铜板 NY2150H 56mil 1/1 37*49 HTE",
            "H/H 8张布 汽车板",
            "成功",
            "AY",
        ),
        (
            "HW双面达到阈值",
            "广东依顿",
            "覆铜板 NY2150H 39mil 2/2OZ 37*49 HTE",
            "HW订单",
            "成功",
            "AP",
        ),
        (
            "HW双面低于阈值",
            "广东依顿",
            "覆铜板 NY2150H 39mil 1/1OZ 37*49 HTE",
            "华为订单",
            "成功",
            "A1",
        ),
        (
            "汽车板口语",
            "广州伊顿",
            "覆铜板 NY2150H 39mil 1/1OZ 37*49 HTE",
            "要汽板",
            "成功",
            "AC",
        ),
    ]
    for name, customer, spec, remark, expected_status, expected_grade in cases:
        result = calculate_transcode_agent_quote(
            spec,
            customer=customer,
            customer_code="103901",
            order_remark=remark,
        )
        assert result["status"] == expected_status, (name, result)
        code = result["result"] or result["candidate_code"]
        assert code[19:21] == expected_grade, (name, code, result)

    mixed = calculate_transcode_agent_quote(
        "覆铜板 NY2150H 39mil 1/2OZ 37*49 HTE",
        customer="广东依顿",
        customer_code="103901",
        order_remark="HW订单",
    )
    assert mixed["status"] == "待确认", mixed
    assert "CPR-YIDUN-HW-MIXED-COPPER" in mixed["error"], mixed
    assert mixed["candidate_code"], mixed

    conflict = calculate_transcode_agent_quote(
        "覆铜板 NY2150H 39mil 2/2OZ 37*49 HTE",
        customer="广州伊顿",
        customer_code="103901",
        order_remark="HW汽车板",
    )
    assert conflict["status"] == "待确认", conflict
    assert "规则冲突" in conflict["error"] and "AC/AP" in conflict["error"], conflict
    assert conflict["candidate_code"], conflict
    print("Yidun latest semantic rules smoke: PASS deterministic=4 pending=2")


if __name__ == "__main__":
    main()
