from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_rule_compiler import (
    compile_semantic_candidate,
    evaluate_semantic_compilation,
)


def main() -> None:
    test_valid_candidate_is_deterministic()
    test_runtime_field_aliases_are_canonicalized()
    test_deterministic_asset_is_action_not_condition()
    test_explicit_mapping_removes_default_priority_false_conflict()
    test_explicit_mapping_is_deterministic()
    test_ambiguity_waits_for_business()
    test_incomplete_low_confidence_rule_waits_for_business()
    test_target_field_mismatch_fails()
    print("semantic rule compiler smoke passed")


def _candidate() -> dict:
    return {
        "candidate_id": "MSR-0007-02",
        "customer_code": "105007",
        "customer_name": "湖奥士康",
        "source_row": 7,
        "business_field": "基板级别",
        "source_text": "当备注中有X0A0/car/汽车板字样时=AC",
        "required_input_fields": "订单备注",
    }


def _valid_result() -> dict:
    return {
        "schema_version": "1.0",
        "task_type": "rule_structure",
        "material_scope": "CCL",
        "semantic_items": [
            {
                "semantic_type": "keyword_present",
                "target_field": "grade_intent",
                "normalized_value": "automotive",
                "stated_target_value": "AC",
                "conditions": [
                    {
                        "field": "订单备注",
                        "operator": "contains_any",
                        "value": ["X0A0", "car", "汽车板"],
                        "source_scope": "订单备注",
                    }
                ],
                "source_field": "CCL特殊规则",
                "evidence_text": "当备注中有X0A0/car/汽车板字样时=AC",
                "confidence": "high",
                "deterministic_preferred": False,
            }
        ],
        "ambiguities": [],
        "missing_inputs": [],
        "model_confidence": "high",
    }


def test_valid_candidate_is_deterministic() -> None:
    class FakeClient:
        def normalize(self, **kwargs):
            assert kwargs["task_type"] == "rule_structure"
            assert kwargs["customer_code"] == "105007"
            assert kwargs["task_context"]["business_field"] == "基板级别"
            return _valid_result()

    compiled = compile_semantic_candidate(_candidate(), FakeClient())
    assert compiled.status == "程序校验通过"
    assert compiled.recommended_execution_mode == "结构化后可确定性执行"
    assert compiled.target_field_summary == "grade_intent"


def test_ambiguity_waits_for_business() -> None:
    result = _valid_result()
    result["ambiguities"] = [
        {
            "field": "订单备注",
            "reason": "电源板和能源板是任一命中还是同时命中未说明",
            "evidence_text": "",
        }
    ]
    compiled = evaluate_semantic_compilation(_candidate(), result)
    assert compiled.status == "待业务确认"
    assert compiled.recommended_execution_mode == "待业务确认"


def test_runtime_field_aliases_are_canonicalized() -> None:
    result = _valid_result()
    condition = result["semantic_items"][0]["conditions"][0]
    condition["field"] = "order_remark"
    condition["source_scope"] = "CCL特殊规则"
    compiled = evaluate_semantic_compilation(_candidate(), result)
    normalized = compiled.model_result["semantic_items"][0]["conditions"][0]
    assert normalized["field"] == "订单备注"
    assert normalized["source_scope"] == "订单备注"
    assert "order_remark->订单备注" in compiled.normalization_notes


def test_deterministic_asset_is_action_not_condition() -> None:
    candidate = _candidate()
    candidate["business_field"] = "基板厚度"
    result = _valid_result()
    item = result["semantic_items"][0]
    item["target_field"] = "total_core"
    item["conditions"] = [
        {"field": "订单备注", "operator": "contains_any", "value": ["含铜"], "source_scope": "订单备注"},
        {"field": "基板厚度", "operator": "lt", "value": 0.8, "source_scope": "订单基板厚度"},
        {
            "field": "基板厚度",
            "operator": "present",
            "value": "总厚转芯厚",
            "source_scope": "available_deterministic_assets",
        },
    ]
    compiled = evaluate_semantic_compilation(candidate, result)
    conditions = compiled.model_result["semantic_items"][0]["conditions"]
    assert len(conditions) == 2
    assert conditions[1]["source_scope"] == "基板厚度"
    assert compiled.model_result["semantic_items"][0]["normalized_value"] == (
        "core_after_total_to_core_conversion"
    )
    assert compiled.status == "程序校验通过"


def test_target_field_mismatch_fails() -> None:
    result = _valid_result()
    result["semantic_items"][0]["target_field"] = "size"
    compiled = evaluate_semantic_compilation(_candidate(), result)
    assert compiled.status == "程序校验失败"
    assert "不一致" in compiled.validation_result


def test_incomplete_low_confidence_rule_waits_for_business() -> None:
    result = _valid_result()
    result["semantic_items"][0]["conditions"] = []
    result["ambiguities"] = [
        {"field": "订单备注", "reason": "只有字段名，没有规则内容", "evidence_text": ""}
    ]
    result["model_confidence"] = "low"
    compiled = evaluate_semantic_compilation(_candidate(), result)
    assert compiled.status == "待业务确认"
    assert "缺少执行条件" in compiled.business_question


def test_explicit_mapping_is_deterministic() -> None:
    candidate = _candidate()
    candidate["business_field"] = "胶系"
    candidate["source_text"] = "TG150=NY2150"
    result = _valid_result()
    item = result["semantic_items"][0]
    item.update(
        semantic_type="explicit_fact",
        target_field="glue",
        normalized_value="NY2150",
        stated_target_value="NY2150",
        evidence_text="TG150=NY2150",
    )
    item["conditions"] = [
        {"field": "胶系", "operator": "equals", "value": "TG150", "source_scope": "胶系"}
    ]
    compiled = evaluate_semantic_compilation(candidate, result)
    assert compiled.status == "程序校验通过"
    assert compiled.recommended_execution_mode == "结构化后可确定性执行"


def test_explicit_mapping_removes_default_priority_false_conflict() -> None:
    candidate = _candidate()
    candidate["business_field"] = "胶系"
    candidate["source_text"] = "未写胶系时=NY2140，TG150=NY2150"
    result = _valid_result()
    result["semantic_items"] = [
        {
            **result["semantic_items"][0],
            "semantic_type": "default_when_missing",
            "target_field": "glue",
            "normalized_value": "NY2140",
            "stated_target_value": "NY2140",
            "conditions": [
                {"field": "胶系", "operator": "missing", "value": None, "source_scope": "胶系"}
            ],
            "evidence_text": "未写胶系时=NY2140",
        },
        {
            **result["semantic_items"][0],
            "semantic_type": "keyword_present",
            "target_field": "glue",
            "normalized_value": "NY2150",
            "stated_target_value": "NY2150",
            "conditions": [
                {"field": "胶系", "operator": "equals", "value": "TG150", "source_scope": "胶系"}
            ],
            "evidence_text": "TG150=NY2150",
        },
    ]
    result["ambiguities"] = [
        {"field": "glue", "reason": "明示与默认可能存在优先级冲突", "evidence_text": ""}
    ]
    compiled = evaluate_semantic_compilation(candidate, result)
    assert compiled.status == "程序校验通过"
    assert not compiled.model_result["ambiguities"]


if __name__ == "__main__":
    main()
