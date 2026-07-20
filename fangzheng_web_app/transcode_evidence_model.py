from __future__ import annotations

import os
from copy import deepcopy
from dataclasses import dataclass
from typing import Any

from .transcode_evidence_scoring import apply_model_evidence_review
from .transcode_semantic_service import DeepSeekSemanticClient, load_semantic_model_config


EVIDENCE_MODEL_MODES = {"off", "shadow"}
MODEL_REVIEWABLE_VERDICTS = {"ambiguous", "missing_evidence"}


@dataclass(frozen=True)
class EvidenceModelRuntime:
    mode: str
    client: DeepSeekSemanticClient | None
    model: str = ""
    load_error: str = ""


def get_evidence_model_runtime_mode(environ: dict[str, str] | None = None) -> str:
    env = os.environ if environ is None else environ
    mode = str(env.get("TRANSCODE_EVIDENCE_MODEL_MODE") or "off").strip().lower()
    return mode if mode in EVIDENCE_MODEL_MODES else "off"


def get_evidence_model_max_calls(environ: dict[str, str] | None = None) -> int:
    env = os.environ if environ is None else environ
    try:
        value = int(str(env.get("TRANSCODE_EVIDENCE_MODEL_MAX_CALLS") or "50").strip())
    except ValueError:
        return 50
    return max(0, min(value, 500))


def load_evidence_model_runtime(environ: dict[str, str] | None = None) -> EvidenceModelRuntime:
    mode = get_evidence_model_runtime_mode(environ)
    if mode == "off":
        return EvidenceModelRuntime(mode="off", client=None)
    env = dict(os.environ if environ is None else environ)
    env["TRANSCODE_SEMANTIC_MODEL_MODE"] = "shadow"
    try:
        config = load_semantic_model_config(environ=env)
        return EvidenceModelRuntime(
            mode="shadow",
            client=DeepSeekSemanticClient(config),
            model=config.model,
        )
    except Exception as exc:
        return EvidenceModelRuntime(mode="shadow", client=None, load_error=str(exc))


def review_evidence_shadow(
    analysis: dict[str, Any],
    *,
    semantic_evaluations: list[dict[str, Any]],
    matrix: dict[str, Any],
    client: DeepSeekSemanticClient,
) -> dict[str, Any]:
    score_shadow = analysis.get("evidence_score_shadow") or {}
    request = build_evidence_review_request(
        analysis,
        score_shadow=score_shadow,
        semantic_evaluations=semantic_evaluations,
    )
    if not request:
        return score_shadow
    try:
        result = client.review_evidence(**request["payload"])
        _validate_requested_reviews(result, request["requested_fields"])
        return apply_model_evidence_review(score_shadow, result, matrix=matrix)
    except Exception as exc:
        fallback = deepcopy(score_shadow)
        fallback["model_called"] = True
        fallback["model_call_count"] = 1
        fallback["model_error"] = str(exc)
        for review in fallback.get("field_reviews") or []:
            if review.get("verdict") not in MODEL_REVIEWABLE_VERDICTS or not review.get("semantic_rule_ids"):
                continue
            review["program_verdict"] = review.get("verdict", "")
            review["program_shadow_score"] = review.get("shadow_score", 0)
            review["model_called"] = True
            review["model_accepted"] = False
            review["model_reason"] = f"模型调用失败：{exc}"
        fallback["runtime_effect"] = "模型证据审查失败，保留程序影子评分，不影响正式转码"
        return fallback


def build_evidence_review_request(
    analysis: dict[str, Any],
    *,
    score_shadow: dict[str, Any],
    semantic_evaluations: list[dict[str, Any]],
) -> dict[str, Any] | None:
    source_fields = {
        str(key): str(value)
        for key, value in (score_shadow.get("source_fields") or {}).items()
        if str(key).strip() and str(value).strip()
    }
    eligible = [
        item
        for item in score_shadow.get("field_reviews") or []
        if item.get("verdict") in MODEL_REVIEWABLE_VERDICTS
        and item.get("semantic_rule_ids")
    ]
    if not source_fields or not eligible:
        return None
    requested_fields = [str(item.get("field_key") or "") for item in eligible]
    candidate_fields = {
        field: {
            "value": item.get("candidate_value", ""),
            "code": item.get("candidate_code", ""),
        }
        for field, item in zip(requested_fields, eligible)
    }
    field_evidence = [
        {
            "field": field,
            "program_verdict": item.get("verdict", ""),
            "candidate_value": item.get("candidate_value", ""),
            "candidate_code": item.get("candidate_code", ""),
            "source": item.get("source_field", ""),
            "evidence": item.get("evidence_text", ""),
            "reason": item.get("reason", ""),
            "rule_id": item.get("rule_id", ""),
        }
        for field, item in zip(requested_fields, eligible)
    ]
    normalized_semantics = {
        "evaluations": [
            {
                "rule_id": item.get("rule_id", ""),
                "status": item.get("status", ""),
                "target_fields": item.get("target_fields") or [],
                "normalized_values": item.get("normalized_values") or [],
                "missing_fields": item.get("missing_fields") or [],
                "evidence_texts": item.get("evidence_texts") or [],
            }
            for item in semantic_evaluations
            if item.get("status") in {"命中", "缺少输入", "条件错误"}
        ]
    }
    relevant_rules = normalized_semantics["evaluations"]
    return {
        "requested_fields": requested_fields,
        "payload": {
            "source_fields": source_fields,
            "normalized_semantics": normalized_semantics,
            "candidate_fields": candidate_fields,
            "field_evidence": field_evidence,
            "relevant_rules": relevant_rules,
        },
    }


def _validate_requested_reviews(result: dict[str, Any], requested_fields: list[str]) -> None:
    returned = [str(item.get("field") or "") for item in result.get("field_reviews") or []]
    if len(returned) != len(set(returned)):
        raise ValueError("模型证据审查返回了重复字段")
    if set(returned) != set(requested_fields):
        raise ValueError(
            f"模型证据审查字段不完整或越权：requested={sorted(requested_fields)} returned={sorted(returned)}"
        )
