from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_order_semantic_model import (
    build_model_rule_evaluations,
    should_normalize_order,
)
from fangzheng_web_app.transcode_semantic_overrides import apply_confirmed_semantic_overrides
from fangzheng_web_app.transcode_agent_service import _build_field_evidence


CASES = [
    ("grade_intent", "AC", "下汽车板"),
    ("grade_intent", "AM", "MINILED产品"),
    ("grade_intent", "AP", "能源板"),
    ("grade_intent", "A1", "普通板"),
    ("copper_type", "RTF", "反转铜箔"),
    ("copper_type", "HTE", "常规铜箔"),
    ("print_mark", "有水印", "要求印字"),
    ("glue", "NY2140", "按普通TG料"),
    ("glue", "NY2150", "TG150料"),
    ("total_core", "total", "按含铜总厚"),
]


class FakeEngine:
    @staticmethod
    def get_copper_type_code(value):
        return {"RTF": "R", "HTE": "W"}[str(value).upper()]

    @staticmethod
    def get_glue_code(value, exact, model, customer):
        return exact.get(str(value).upper())


def evaluation(index: int, target: str, value: str) -> dict:
    return {
        "rule_id": f"TSR-TEST-{index:02d}",
        "status": "未命中",
        "business_field": "基板级别",
        "target_fields": [target],
        "normalized_values": [value],
        "source_text": "已批准订单备注规则",
        "condition_results": [
            {"field": "订单备注", "matched": False},
            {"field": "胶系", "matched": True},
        ],
        "observed_inputs": {"订单备注": {"available": True, "value": "口语备注"}},
    }


def model_result(items: list[dict], confidence: str = "high") -> dict:
    return {
        "schema_version": "1.0",
        "task_type": "order_normalization",
        "material_scope": "CCL",
        "semantic_items": items,
        "ambiguities": [],
        "missing_inputs": [],
        "model_confidence": confidence,
    }


def main() -> None:
    evaluations = [evaluation(i, target, value) for i, (target, value, _) in enumerate(CASES, 1)]
    items = [
        {
            "target_field": target,
            "normalized_value": value,
            "confidence": "high",
            "evidence_text": evidence,
        }
        for target, value, evidence in CASES
    ]
    matched, notes = build_model_rule_evaluations(model_result(items), evaluations)
    assert len(matched) == len(CASES), (len(matched), notes)
    assert should_normalize_order(evaluations, "有备注")
    assert not should_normalize_order(evaluations, "")

    low, _ = build_model_rule_evaluations(model_result(items, confidence="medium"), evaluations)
    assert not low
    blocked = evaluation(99, "grade_intent", "AC")
    blocked["condition_results"][1]["matched"] = False
    blocked_match, _ = build_model_rule_evaluations(model_result([items[0]]), [blocked])
    assert not blocked_match

    analysis = {
        "customer": "测试客户",
        "engine_steps": {"step7_grade_code": "A1", "errors": []},
        "applied_rules": [{"field": "grade_code", "rule_id": "AGENT-OLD"}],
    }
    grade_evaluation = next(item for item in matched if item["normalized_values"] == ["AC"])
    applied, conflicts = apply_confirmed_semantic_overrides(
        FakeEngine(),
        {"grade_code_map": {"A1": "A级", "AC": "汽车板"}},
        analysis,
        [grade_evaluation],
        allow_order_remark_priority=True,
    )
    assert not conflicts
    assert analysis["engine_steps"]["step7_grade_code"] == "AC"
    assert applied and applied[0]["source"] == "模型标准化+已批准语义规则"
    evidence = _build_field_evidence(analysis["engine_steps"], [], applied, [])
    grade_evidence = next(item for item in evidence if item["field_key"] == "grade")
    assert grade_evidence["score"] == 98, grade_evidence
    structure_evidence = next(item for item in evidence if item["field_key"] == "structure")
    assert structure_evidence["score"] is None
    assert not structure_evidence["gate"]
    print("order semantic activation smoke passed classic_cases=10 safe_binding=true")


if __name__ == "__main__":
    main()
