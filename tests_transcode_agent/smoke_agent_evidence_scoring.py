from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_evidence_scoring import (
    evidence_gate_decision,
    evaluate_evidence_score_shadow,
    get_evidence_gate_mode,
    get_evidence_score_runtime_mode,
    load_evidence_score_matrix,
)


def main() -> None:
    matrix = load_evidence_score_matrix()
    assert get_evidence_score_runtime_mode({}) == "shadow"
    assert get_evidence_score_runtime_mode({"TRANSCODE_EVIDENCE_SCORE_MODE": "off"}) == "off"
    assert get_evidence_gate_mode({}) == "enforce"
    assert get_evidence_gate_mode({"TRANSCODE_EVIDENCE_GATE_MODE": "enforce"}) == "enforce"

    default_grade = evaluate_evidence_score_shadow(
        _analysis(_field("grade", "基板级别", "A1", 86, "默认规则", "基础规则")),
        semantic_evaluations=[],
        observations={},
        matrix=matrix,
    )
    assert default_grade["shadow_score"] == 86
    assert default_grade["shadow_decision"] == "需标注"
    assert default_grade["field_reviews"][0]["verdict"] == "supported"

    exact_grade = evaluate_evidence_score_shadow(
        _analysis(_field("grade", "基板级别", "AC", 86, "默认规则", "基础规则")),
        semantic_evaluations=[_semantic("grade_intent", "AC", "命中")],
        observations=_observations(),
        matrix=matrix,
    )
    assert exact_grade["shadow_score"] == 98
    assert exact_grade["shadow_decision"] == "通过"
    assert exact_grade["field_reviews"][0]["verdict"] == "supported"

    contradicted = evaluate_evidence_score_shadow(
        _analysis(_field("grade", "基板级别", "A1", 86, "默认规则", "基础规则")),
        semantic_evaluations=[_semantic("grade_intent", "AP", "命中")],
        observations=_observations(),
        matrix=matrix,
    )
    assert contradicted["shadow_score"] == 0
    assert contradicted["field_reviews"][0]["verdict"] == "contradicted"

    missing = evaluate_evidence_score_shadow(
        _analysis(_field("grade", "基板级别", "A1", 86, "默认规则", "基础规则")),
        semantic_evaluations=[_semantic("grade_intent", "AP", "缺少输入", ["订单备注"])],
        observations=_observations(),
        matrix=matrix,
    )
    assert missing["shadow_score"] == 60
    assert missing["field_reviews"][0]["verdict"] == "missing_evidence"
    assert missing["hard_blockers"] == ["基板级别:missing_evidence"]

    agent_rule = evaluate_evidence_score_shadow(
        _analysis(_field("grade", "基板级别", "AC", 99, "Agent规则覆盖", "TAR-00001")),
        semantic_evaluations=[],
        observations=_observations(),
        matrix=matrix,
    )
    assert agent_rule["shadow_score"] == 99
    assert agent_rule["field_reviews"][0]["verdict"] == "supported"
    assert agent_rule["model_called"] is False

    gate_analysis = {"overall_score": 94, "evidence_score_shadow": missing}
    assert evidence_gate_decision(gate_analysis, mode="shadow")["blocked"] is False
    enforce = evidence_gate_decision(gate_analysis, mode="enforce")
    assert enforce["blocked"] is False
    assert enforce["effective_score"] == 94
    assert enforce["ignored_optional_missing_rules"] == 1

    print("evidence scoring smoke passed default/exact/contradicted/missing/agent-rule verified")


def _analysis(field: dict) -> dict:
    return {"overall_score": field["score"], "field_evidence": [field]}


def _field(field_key: str, label: str, code: str, score: int, hit_type: str, source: str) -> dict:
    return {
        "field_key": field_key,
        "field": label,
        "value": code,
        "code": code,
        "score": score,
        "gate": True,
        "hit_type": hit_type,
        "source": source,
        "evidence": "smoke evidence",
        "rule_id": "",
    }


def _semantic(target: str, value: str, status: str, missing_fields: list[str] | None = None) -> dict:
    return {
        "rule_id": "TSR-SMOKE",
        "target_fields": [target],
        "normalized_values": [value],
        "status": status,
        "missing_fields": missing_fields or [],
        "evidence_texts": ["smoke semantic evidence"],
    }


def _observations() -> dict:
    return {"订单备注": {"available": True, "value": "汽车板", "sources": ["订单备注"]}}


if __name__ == "__main__":
    main()
