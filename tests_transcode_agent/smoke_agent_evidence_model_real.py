from __future__ import annotations

import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_service import DeepSeekSemanticClient, load_semantic_model_config


def main() -> None:
    env = dict(os.environ)
    env["TRANSCODE_SEMANTIC_MODEL_MODE"] = "shadow"
    config = load_semantic_model_config(environ=env)
    result = DeepSeekSemanticClient(config).review_evidence(
        source_fields={"订单规格": "NY2150 汽车板 0.8mm 1/1 37*49 HTE"},
        normalized_semantics={
            "evaluations": [
                {
                    "rule_id": "TSR-REAL-SMOKE",
                    "status": "命中",
                    "target_fields": ["grade_intent"],
                    "normalized_values": ["AC"],
                    "evidence_texts": ["汽车板"],
                }
            ]
        },
        candidate_fields={"grade": {"value": "AC", "code": "AC"}},
        field_evidence=[
            {
                "field": "grade",
                "program_verdict": "ambiguous",
                "candidate_value": "AC",
                "candidate_code": "AC",
                "source": "默认规则",
                "evidence": "",
                "reason": "需要原文证据",
                "rule_id": "TSR-REAL-SMOKE",
            }
        ],
        relevant_rules=[
            {
                "rule_id": "TSR-REAL-SMOKE",
                "target_fields": ["grade_intent"],
                "normalized_values": ["AC"],
                "evidence_texts": ["汽车板"],
            }
        ],
    )
    reviews = result.get("field_reviews") or []
    assert len(reviews) == 1 and reviews[0].get("field") == "grade", result
    print(
        "real evidence model smoke passed "
        f"model={config.model} verdict={reviews[0].get('verdict')} "
        f"confidence={result.get('model_confidence')}"
    )


if __name__ == "__main__":
    main()
