from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import _calculate_transcode_agent_analysis


def _analyze(customer: str, spec: str, remark: str = "") -> dict:
    analysis, _, _ = _calculate_transcode_agent_analysis(
        spec,
        customer=customer,
        order_remark=remark,
    )
    return analysis


def _scores(analysis: dict) -> dict[str, int | None]:
    return {
        str(item.get("field_key")): item.get("score")
        for item in analysis.get("field_evidence") or []
    }


def main() -> None:
    yidun_39 = _analyze(
        "广东依顿",
        '覆铜板 39±3MIL 2/2OZ 37*49" TG≥150 HTE 不连铜 NY2150H 5张7628 ANTI-CAF',
    )
    assert yidun_39["status"] == "成功", yidun_39
    assert yidun_39["formal_code"], yidun_39
    assert yidun_39["overall_score"] == 100, yidun_39

    yidun_16 = _analyze(
        "广东依顿",
        '覆铜板 16±1.5MIL 2/2OZ 37.3*49" TG≥150 HTE 不连铜 NY2150H 2张7628 ANTI-CAF',
    )
    assert yidun_16["status"] == "成功", yidun_16
    assert yidun_16["overall_score"] == 100, yidun_16

    yidun_2170 = _analyze(
        "广东依顿",
        '覆铜板 16±1.5MIL H/HOZ 82*49" TG≥170 HTE 不连铜 NY2170 ANTI-CAF',
    )
    assert yidun_2170["status"] == "成功", yidun_2170
    assert yidun_2170["overall_score"] == 100, yidun_2170

    jingwang = _analyze(
        "龙川景旺",
        "CCL NY2150 0.43mm 3/3 (不含铜) 86inX49.3in 7628*1+1080*1+7628*1 HTE",
        "基板级别下AP",
    )
    assert jingwang["status"] == "成功", jingwang
    assert jingwang["overall_score"] == 100, jingwang
    assert jingwang["engine_steps"]["step7_grade_code"] == "AP", jingwang
    assert jingwang["formal_code"][19:21] == "AP", jingwang["formal_code"]

    yidun_auto = _analyze(
        "广东依顿",
        '覆铜板 36±3MIL 2/2OZ 74*49" TG≥150 HTE 不连铜 NY2150H ANTI-CAF',
        "下汽车板",
    )
    assert yidun_auto["status"] == "成功", yidun_auto
    assert yidun_auto["engine_steps"]["step7_grade_code"] == "AC", yidun_auto

    explicit_conflict = _analyze(
        "普通客户",
        'NY2150 0.8mm 1/1 37*49 HTE',
        "基板级别下A1，等级下AC",
    )
    assert explicit_conflict["status"] == "待确认", explicit_conflict
    assert explicit_conflict["decision_state"] == "规则冲突", explicit_conflict
    assert "订单备注同时指定多个等级" in explicit_conflict["reason"], explicit_conflict

    print("unified confirmation decision smoke passed")


if __name__ == "__main__":
    main()
