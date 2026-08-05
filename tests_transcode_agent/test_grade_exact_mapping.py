from __future__ import annotations

from fangzheng_web_app.transcode_agent_engine import get_grade_code, match_exact_grade_desc
from fangzheng_web_app.transcode_agent_service import _load_runtime
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
