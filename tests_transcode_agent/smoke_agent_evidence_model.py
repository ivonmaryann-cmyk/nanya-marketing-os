from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_evidence_model import (
    build_evidence_review_request,
    get_evidence_model_max_calls,
    get_evidence_model_runtime_mode,
    review_evidence_shadow,
)
from fangzheng_web_app.transcode_evidence_scoring import load_evidence_score_matrix
from fangzheng_web_app.transcode_evidence_scoring import evidence_gate_decision


class FakeClient:
    def __init__(self, result=None, error: Exception | None = None):
        self.result = result
        self.error = error
        self.calls = []

    def review_evidence(self, **kwargs):
        self.calls.append(kwargs)
        if self.error:
            raise self.error
        return self.result


def main() -> None:
    assert get_evidence_model_runtime_mode({}) == "off"
    assert get_evidence_model_runtime_mode({"TRANSCODE_EVIDENCE_MODEL_MODE": "shadow"}) == "shadow"
    assert get_evidence_model_max_calls({}) == 50
    assert get_evidence_model_max_calls({"TRANSCODE_EVIDENCE_MODEL_MAX_CALLS": "12"}) == 12
    matrix = load_evidence_score_matrix()
    analysis = _analysis()
    request = build_evidence_review_request(
        analysis,
        score_shadow=analysis["evidence_score_shadow"],
        semantic_evaluations=[],
    )
    assert request is not None
    assert request["requested_fields"] == ["grade"]
    assert set(request["payload"]["candidate_fields"]) == {"grade"}

    client = FakeClient(_supported_result())
    merged = review_evidence_shadow(
        analysis,
        semantic_evaluations=[],
        matrix=matrix,
        client=client,
    )
    review = merged["field_reviews"][0]
    assert len(client.calls) == 1
    assert review["program_verdict"] == "ambiguous"
    assert review["model_verdict"] == "supported"
    assert review["verdict"] == "supported"
    assert review["shadow_score"] == 95
    assert merged["model_call_count"] == 1
    assert merged["shadow_decision"] == "通过"
    gated = evidence_gate_decision({"overall_score": 90, "evidence_score_shadow": merged}, mode="enforce")
    assert gated["blocked"] is True
    assert gated["program_evidence_score"] == 80

    low_result = _supported_result()
    low_result["model_confidence"] = "medium"
    low = review_evidence_shadow(
        _analysis(),
        semantic_evaluations=[],
        matrix=matrix,
        client=FakeClient(low_result),
    )
    assert low["field_reviews"][0]["verdict"] == "ambiguous"
    assert low["field_reviews"][0]["shadow_score"] == 80
    assert low["field_reviews"][0]["model_accepted"] is False

    failed = review_evidence_shadow(
        _analysis(),
        semantic_evaluations=[],
        matrix=matrix,
        client=FakeClient(error=RuntimeError("timeout")),
    )
    assert failed["shadow_score"] == 80
    assert failed["model_error"] == "timeout"
    assert failed["field_reviews"][0]["model_called"] is True
    print("evidence model smoke passed selection/high-confidence merge/low-confidence fallback/error fallback")


def _analysis():
    shadow = {
        "mode": "shadow",
        "threshold": 90,
        "current_score": 86,
        "shadow_score": 80,
        "score_delta": -6,
        "shadow_decision": "需标注",
        "source_fields": {"订单规格": "NY2150 汽车板 0.8mm"},
        "hard_blockers": [],
        "field_reviews": [
            {
                "field_key": "grade",
                "field": "基板级别",
                "candidate_value": "AC",
                "candidate_code": "AC",
                "current_score": 86,
                "shadow_score": 80,
                "score_delta": -6,
                "verdict": "ambiguous",
                "source_field": "基础规则",
                "evidence_text": "默认A1不构成高置信证据",
                "reason": "默认规则",
                "hit_type": "默认规则",
                "rule_id": "",
                "semantic_rule_ids": ["TSR-SMOKE"],
                "semantic_evidence": [],
                "model_called": False,
            }
        ],
    }
    return {"evidence_score_shadow": shadow}


def _supported_result():
    return {
        "schema_version": "1.0",
        "field_reviews": [
            {
                "field": "grade",
                "verdict": "supported",
                "source_field": "订单规格",
                "evidence_text": "汽车板",
                "reason": "订单规格明确包含汽车板",
            }
        ],
        "hard_blockers": [],
        "model_confidence": "high",
    }


if __name__ == "__main__":
    main()
