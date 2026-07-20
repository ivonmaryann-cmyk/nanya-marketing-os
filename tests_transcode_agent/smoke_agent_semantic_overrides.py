from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_overrides import (
    apply_confirmed_semantic_overrides,
    get_semantic_override_mode,
)


class FakeEngine:
    @staticmethod
    def get_copper_type_code(value):
        return {"RTF": "R", "HTE": "W"}[str(value).upper()]

    @staticmethod
    def get_glue_code(value, exact, model, customer):
        return exact.get(str(value).upper())


def main() -> None:
    assert get_semantic_override_mode({}) == "enforce"
    tables = {
        "grade_code_map": {"A1": "A级", "AC": "汽车板", "AM": "MINILED", "AP": "电源板"},
        "glue_exact_map": {"NY2140": "2A", "NY2150": "2B"},
        "glue_model_map": {},
    }
    analysis = {
        "customer": "测试客户",
        "engine_steps": {
            "step1_glue_code": "??",
            "step6_copper_type_code": "W",
            "step7_grade_code": "A1",
            "errors": ["无法识别胶系型号"],
        },
    }
    applied, conflicts = apply_confirmed_semantic_overrides(
        FakeEngine(),
        tables,
        analysis,
        [
            _evaluation("TSR-GRADE", "grade_intent", "AC"),
            _evaluation("TSR-GLUE", "glue", "NY2140"),
            _evaluation("TSR-SIZE", "size", "external_lookup:测试表"),
        ],
    )
    assert not conflicts
    assert analysis["engine_steps"]["step1_glue_code"] == "2A"
    assert analysis["engine_steps"]["step7_grade_code"] == "AC"
    assert analysis["engine_steps"]["errors"] == []
    assert {item["rule_id"] for item in applied} == {"TSR-GRADE", "TSR-GLUE"}

    conflict_analysis = {"customer": "测试", "engine_steps": {"step7_grade_code": "A1", "errors": []}}
    applied, conflicts = apply_confirmed_semantic_overrides(
        FakeEngine(),
        tables,
        conflict_analysis,
        [
            _evaluation("TSR-AC", "grade_intent", "AC"),
            _evaluation("TSR-AM", "grade_intent", "AM"),
        ],
    )
    assert not applied
    assert conflicts and "AC/AM" in conflicts[0]
    print("semantic overrides smoke passed grade/glue/unsupported-skip/conflict-block")


def _evaluation(rule_id, target, value):
    return {
        "rule_id": rule_id,
        "status": "命中",
        "target_fields": [target],
        "normalized_values": [value],
        "source_text": "测试规则",
        "business_field": "CCL特殊规则",
    }


if __name__ == "__main__":
    main()
