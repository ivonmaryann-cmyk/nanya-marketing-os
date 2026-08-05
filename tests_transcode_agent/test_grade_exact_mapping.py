from __future__ import annotations

from fangzheng_web_app.transcode_agent_engine import get_grade_code, match_exact_grade_desc
from fangzheng_web_app.transcode_agent_service import (
    _business_field_evidence,
    _load_runtime,
    _score_field,
    analyze_spec,
)
from fangzheng_web_app.transcode_semantic_overrides import apply_confirmed_semantic_overrides


GRADE_DESC_TO_CODE = {
    "汽车板": "AC",
    "汽车板考试板": "AQ",
    "Module": "AJ",
    "电源板": "AP",
    "基板经纬不分": "NN",
    "Mini LED": "AM",
    "HDI专用": "AD",
}


def test_match_exact_grade_desc_uses_longest_phrase_first():
    assert match_exact_grade_desc("汽车板考试板", GRADE_DESC_TO_CODE) == ("AQ", "汽车板考试板")
    assert match_exact_grade_desc("汽车板", GRADE_DESC_TO_CODE) == ("AC", "汽车板")
    assert match_exact_grade_desc("汽车", GRADE_DESC_TO_CODE) is None
    assert match_exact_grade_desc("qiche", GRADE_DESC_TO_CODE) is None


def test_get_grade_code_applies_exact_spec_mappings():
    assert get_grade_code("NY2150 1.0 1/1 41*49 汽车板考试板", "", "", {}, GRADE_DESC_TO_CODE) == "AQ"
    assert get_grade_code("NY2150 1.0 1/1 41*49 汽车板", "", "", {}, GRADE_DESC_TO_CODE) == "AC"
    assert get_grade_code("NY2150 1.0 1/1 41*49 Module", "", "", {}, GRADE_DESC_TO_CODE) == "AJ"
    assert get_grade_code("NY2150 1.0 1/1 41*49 电源板", "", "", {}, GRADE_DESC_TO_CODE) == "AP"
    assert get_grade_code("NY2150 1.0 1/1 41*49 基板经纬不分", "", "", {}, GRADE_DESC_TO_CODE) == "NN"
    assert get_grade_code("NY2150 1.0 1/1 41*49 Mini LED", "", "", {}, GRADE_DESC_TO_CODE) == "AM"
    assert get_grade_code("NY2150 1.0 1/1 41*49 HDI专用", "", "", {}, GRADE_DESC_TO_CODE) == "AD"


def test_get_grade_code_rejects_fuzzy_and_pinyin():
    assert get_grade_code("NY2150 1.0 1/1 41*49 汽车", "", "", {}, GRADE_DESC_TO_CODE) == "A1"
    assert get_grade_code("NY2150 1.0 1/1 41*49 车", "", "", {}, GRADE_DESC_TO_CODE) == "A1"
    assert get_grade_code("NY2150 1.0 1/1 41*49 qiche", "", "", {}, GRADE_DESC_TO_CODE) == "A1"
    assert get_grade_code("NY2150 1.0 1/1 41*49 qicheban", "", "", {}, GRADE_DESC_TO_CODE) == "A1"


def test_default_a1_without_explicit_rule_is_not_100():
    score, hit_type, _source, _note = _score_field(
        "grade",
        "A1",
        {
            "glue_model": "NY2150",
            "thickness_raw": "1.0",
            "thickness_mm": 1.0,
            "copper_spec_raw": "1/1",
            "size_w": 41,
            "size_h": 49,
        },
        [],
        {},
        [],
    )
    assert score == 99
    assert hit_type == "解析来源待确认"


def test_default_a1_spec_goes_to_confirmation():
    engine, tables, agent_rules, _mapping_tables, _base_version, _agent_version = _load_runtime()
    analysis = analyze_spec(
        engine,
        tables,
        agent_rules,
        "NY2150 1.0 1/1 41*49",
        customer="",
        customer_code="",
    )
    assert analysis["engine_steps"]["step7_grade_code"] == "A1"
    assert analysis["status"] == "待确认"
    assert analysis["formal_code"] == ""


def test_business_field_evidence_is_readable_without_internal_ids():
    analysis = {
        "engine_steps": {
            "agent_glue_name": "NY2170",
            "glue_model": "NY2170",
            "step1_glue_code": "2C",
            "thickness_raw": "16±1.5MIL",
            "step2_thick_code": "00406",
            "copper_spec_raw": "H/H",
            "step3_copper_code": "HH",
            "size_w": 82,
            "size_h": 49,
            "step4_size_code": "82004900",
            "glue_category": "普通",
            "step5_glue_cat_code": "Y",
            "step6_copper_type_code": "W",
            "step7_grade_code": "AC",
            "order_type": "芯厚",
            "step8_tc_code": "C",
        },
        "order_semantic_model": {
            "source_fields": {"订单备注": "qiche"},
            "result": {
                "semantic_items": [
                    {
                        "target_field": "grade_intent",
                        "stated_target_value": "qiche",
                        "normalized_value": "汽车板",
                    }
                ]
            },
        },
        "applied_rules": [],
    }

    glue = _business_field_evidence(analysis, {"field_key": "glue", "code": "2C"})
    thickness = _business_field_evidence(
        analysis,
        {"field_key": "thickness", "code": "00406"},
    )
    copper_type = _business_field_evidence(
        analysis,
        {"field_key": "copper_type", "code": "W"},
    )
    grade = _business_field_evidence(
        analysis,
        {
            "field_key": "grade",
            "code": "AC",
            "score": 98,
            "hit_type": "模型语义标准化",
        },
    )

    assert glue == "规格识别出NY2170 → 胶系代码2C"
    assert thickness == "规格厚度16±1.5MIL → 厚度码00406"
    assert copper_type == "未指定特殊铜箔类型，按业务确认常规 HTE/W"
    assert "订单备注「qiche」" in grade
    assert "基板级别=AC" in grade
    for text in (glue, thickness, copper_type, grade):
        assert "TSR-" not in text
        assert "TGM-" not in text
        assert "transcode_rules.xlsx" not in text
        assert "BASE-" not in text


def test_semantic_override_cannot_overwrite_exact_spec_grade():
    engine, tables, _agent_rules, _mapping_tables, _base_version, _agent_version = _load_runtime()
    analysis = {
        "engine_steps": {
            "step7_grade_code": "AJ",
            "grade_note": "基板级别写法：MODULE→AJ",
        },
        "applied_rules": [],
        "customer": "测试客户",
        "spec": "Module",
    }
    evaluations = [
        {
            "status": "命中",
            "target_fields": ["grade_intent"],
            "normalized_values": ["AC"],
            "priority": 100,
            "business_field": "基板级别",
            "source_column": "订单备注",
            "rule_id": "TSR-QICHE",
            "source_text": "qiche=汽车板",
            "model_normalized": True,
            "condition_results": [
                {
                    "field": "订单备注",
                    "operator": "contains_any",
                    "value": ["qiche"],
                    "matched": True,
                }
            ],
            "observed_inputs": {
                "订单备注": {"available": True, "value": "qiche", "sources": ["备注"]}
            },
        }
    ]

    applied, conflicts = apply_confirmed_semantic_overrides(
        engine,
        tables,
        analysis,
        evaluations,
        allow_order_remark_priority=True,
    )

    assert applied == []
    assert any("不得覆盖" in item for item in conflicts)
    assert analysis["engine_steps"]["step7_grade_code"] == "AJ"
