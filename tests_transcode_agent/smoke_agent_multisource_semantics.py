from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import (
    _load_runtime,
    _refresh_analysis_after_semantic_overrides,
    analyze_spec,
)
from fangzheng_web_app.transcode_customer_identity import customer_names_match
from fangzheng_web_app.transcode_semantic_overrides import apply_confirmed_semantic_overrides
from fangzheng_web_app.transcode_semantic_rules import parse_semantic_rule_workbook
from fangzheng_web_app.transcode_semantic_shadow import evaluate_semantic_shadow


RULE_PATH = ROOT / "fangzheng_web_app/default_rules/transcode_semantic/transcode_semantic_rules.xlsx"


def main() -> None:
    assert customer_names_match("广东依顿", "广州伊顿")
    rules, summary = parse_semantic_rule_workbook(RULE_PATH)
    assert summary["rule_count"] == 76
    assert sum(rule.get("source_column") != "CCL特殊规则" for rule in rules) == 21
    test_yidun_automotive_order_remark(rules)
    test_ten_confirmed_samples(rules)
    print("multisource semantic smoke passed rules=76 samples=10")


def test_yidun_automotive_order_remark(rules) -> None:
    engine, tables, agent_rules, mapping_tables, *_ = _load_runtime()
    spec = "覆铜板 36±3MIL 2/2OZ 74*49 TG>=150 HTE 不连铜 NY2150H ANTI-CAF"
    analysis = analyze_spec(
        engine,
        tables,
        agent_rules,
        spec,
        agent_mapping_tables=mapping_tables,
        customer="广州伊顿",
        customer_code="",
        context_text="下汽车板",
        parse_fallback_text="",
    )
    steps = analysis["engine_steps"]
    observations = {
        "订单备注": {"available": True, "value": "下汽车板", "sources": ["测试"]},
        "客户规格": {"available": True, "value": spec, "sources": ["测试"]},
        "订单规格": {"available": True, "value": spec, "sources": ["测试"]},
        "订单规格/订单备注": {
            "available": True,
            "value": f"{spec} 下汽车板",
            "sources": ["测试"],
        },
        "胶系": {"available": True, "value": steps.get("glue_model"), "sources": ["解析"]},
        "基板厚度": {"available": True, "value": steps.get("thickness_mm"), "sources": ["解析"]},
        "铜箔规格": {"available": True, "value": steps.get("copper_spec_raw"), "sources": ["解析"]},
    }
    evaluations = evaluate_semantic_shadow(
        rules,
        customer_code="",
        customer_name="广州伊顿",
        observations=observations,
        spec=spec,
    )
    applied, conflicts = apply_confirmed_semantic_overrides(
        engine,
        tables,
        analysis,
        evaluations,
        allow_order_remark_priority=True,
    )
    assert not conflicts
    _refresh_analysis_after_semantic_overrides(analysis, applied, conflicts)
    assert analysis["engine_steps"]["step7_grade_code"] == "AC"
    assert analysis["engine_steps"]["step8_tc_code"] == "T"
    assert analysis["candidate_code"].startswith("2H010502274004900RWACT")


def test_ten_confirmed_samples(rules) -> None:
    samples = [
        ("广东依顿", "NY2150H 36MIL 2/2OZ 74*49 HTE", "下汽车板", "grade", "AC"),
        ("广州伊顿", "NY2150H 36MIL 2/2OZ 74*49 HTE", "普通板", "grade", "A1"),
        ("东莞康源", "NY2150 0.8mm 1/1 37*49 HTE", "", "grade", "AC"),
        ("益明正宏", "NY-A2 0.8mm 1/1 37*49 HTE", "", "grade", "AC"),
        ("韩国HW", "NY2150 0.8mm 1/1 37*49 HTE", "", "grade", "AC"),
        ("加宏CA", "NY2150 0.8mm 1/1 37*49 HTE", "", "copper_type", "Q"),
        ("西班牙G", "NY2150 0.8mm 1/1 37*49 有水印", "", "copper_type", "I"),
        ("江苏洲旭", "NY3170M 0.8mm 1/1 37*49 双面", "", "copper_type", "R"),
        ("江苏洲旭", "NY3170M 0.8mm 1/1 37*49", "", "copper_type", "W"),
        ("赣州深联", "NY2150 0.8mm 1/1 27.3*49 HTE", "", "size", "27304930"),
    ]
    step_keys = {
        "grade": "step7_grade_code",
        "copper_type": "step6_copper_type_code",
        "size": "step4_size_code",
    }
    for customer, spec, remark, field, expected in samples:
        analysis = _run_sample(rules, customer, spec, remark)
        assert analysis["engine_steps"][step_keys[field]] == expected


def _run_sample(rules, customer: str, spec: str, remark: str):
    engine, tables, agent_rules, mapping_tables, *_ = _load_runtime()
    analysis = analyze_spec(
        engine,
        tables,
        agent_rules,
        spec,
        agent_mapping_tables=mapping_tables,
        customer=customer,
        customer_code="",
        context_text=remark,
        parse_fallback_text="",
    )
    steps = analysis["engine_steps"]
    observations = {
        "订单备注": {"available": True, "value": remark, "sources": ["测试"]},
        "客户规格": {"available": True, "value": spec, "sources": ["测试"]},
        "订单规格": {"available": True, "value": spec, "sources": ["测试"]},
        "订单规格/订单备注": {"available": True, "value": f"{spec} {remark}", "sources": ["测试"]},
        "胶系": {"available": True, "value": steps.get("glue_model"), "sources": ["解析"]},
        "基板厚度": {"available": True, "value": steps.get("thickness_mm"), "sources": ["解析"]},
        "铜箔规格": {"available": True, "value": steps.get("copper_spec_raw"), "sources": ["解析"]},
    }
    evaluations = evaluate_semantic_shadow(
        rules,
        customer_code="",
        customer_name=customer,
        observations=observations,
        spec=spec,
    )
    applied, conflicts = apply_confirmed_semantic_overrides(
        engine,
        tables,
        analysis,
        evaluations,
        allow_order_remark_priority=True,
    )
    assert not conflicts
    _refresh_analysis_after_semantic_overrides(analysis, applied, conflicts)
    return analysis


if __name__ == "__main__":
    main()
