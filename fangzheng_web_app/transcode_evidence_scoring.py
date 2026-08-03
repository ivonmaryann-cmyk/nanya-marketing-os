from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from pathlib import Path
from typing import Any


EVIDENCE_SCORE_MODES = {"off", "shadow"}
EVIDENCE_GATE_MODES = {"shadow", "enforce"}
VERDICTS = {"supported", "contradicted", "ambiguous", "missing_evidence"}
DEFAULT_SCORE_MATRIX_PATH = (
    Path(__file__).resolve().parents[1]
    / "model_skills/marketing-transcode-semantics/references/evidence_score_matrix.json"
)


def get_evidence_score_runtime_mode(environ: dict[str, str] | None = None) -> str:
    env = os.environ if environ is None else environ
    mode = str(env.get("TRANSCODE_EVIDENCE_SCORE_MODE") or "shadow").strip().lower()
    return mode if mode in EVIDENCE_SCORE_MODES else "off"


def get_evidence_gate_mode(environ: dict[str, str] | None = None) -> str:
    env = os.environ if environ is None else environ
    mode = str(env.get("TRANSCODE_EVIDENCE_GATE_MODE") or "enforce").strip().lower()
    return mode if mode in EVIDENCE_GATE_MODES else "shadow"


def evidence_gate_decision(analysis: dict[str, Any], *, mode: str) -> dict[str, Any]:
    shadow = analysis.get("evidence_score_shadow") or {}
    current_score = int(analysis.get("overall_score") or 0)
    shadow_score = int(shadow.get("shadow_score") or 0)
    threshold = int(shadow.get("threshold") or 100)
    reviews = shadow.get("field_reviews") or []
    formal_reviews = [
        item
        for item in reviews
        if not (
            item.get("verdict") == "missing_evidence"
            and item.get("semantic_rule_ids")
        )
    ]
    program_evidence_score = min(
        [int(item.get("program_shadow_score", item.get("shadow_score", 0)) or 0) for item in formal_reviews]
        or [current_score]
    )
    final_evidence_score = min(
        [int(item.get("shadow_score") or 0) for item in formal_reviews]
        or [current_score]
    )
    effective_score = min(current_score, final_evidence_score, program_evidence_score) if reviews else current_score
    blockers = [
        blocker
        for blocker in shadow.get("hard_blockers") or []
        if not str(blocker).endswith(":missing_evidence")
    ]
    should_block = mode == "enforce" and (effective_score < threshold or bool(blockers))
    return {
        "mode": mode,
        "threshold": threshold,
        "current_score": current_score,
        "evidence_score": shadow_score,
        "formal_evidence_score": final_evidence_score,
        "program_evidence_score": program_evidence_score,
        "effective_score": effective_score,
        "blockers": blockers,
        "ignored_optional_missing_rules": len(reviews) - len(formal_reviews),
        "blocked": should_block,
        "rule": "有效分=min(确定性分,证据分)；模型和证据只能维持或降低正式出码资格",
    }


def load_evidence_score_matrix(path: str | Path = DEFAULT_SCORE_MATRIX_PATH) -> dict[str, Any]:
    matrix_path = Path(path)
    payload = json.loads(matrix_path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0":
        raise ValueError(f"证据评分矩阵版本无效：{matrix_path}")
    if int(payload.get("gate_threshold") or 0) <= 0:
        raise ValueError("证据评分矩阵缺少有效门禁")
    verdict_scores = payload.get("verdict_scores") or {}
    if set(verdict_scores) != VERDICTS:
        raise ValueError("证据评分矩阵的verdict_scores不完整")
    if matrix_path.resolve() == DEFAULT_SCORE_MATRIX_PATH.resolve():
        from .transcode_rule_center import merge_score_matrix

        payload = merge_score_matrix(payload)
    return payload


def evaluate_evidence_score_shadow(
    analysis: dict[str, Any],
    *,
    semantic_evaluations: list[dict[str, Any]],
    observations: dict[str, dict[str, Any]],
    matrix: dict[str, Any] | None = None,
) -> dict[str, Any]:
    score_matrix = matrix or load_evidence_score_matrix()
    threshold = int(score_matrix["gate_threshold"])
    reviews: list[dict[str, Any]] = []
    for evidence in analysis.get("field_evidence") or []:
        if not evidence.get("gate"):
            continue
        reviews.append(
            _review_field(
                evidence,
                semantic_evaluations=semantic_evaluations,
                observations=observations,
                matrix=score_matrix,
            )
        )
    shadow_score = min([int(item["shadow_score"]) for item in reviews] or [0])
    current_score = int(analysis.get("overall_score") or 0)
    blockers = [
        f"{item['field']}:{item['verdict']}"
        for item in reviews
        if item["verdict"] in {"contradicted", "missing_evidence"}
    ]
    decision = "通过" if shadow_score >= threshold and not blockers else "需标注"
    return {
        "mode": "shadow",
        "threshold": threshold,
        "current_score": current_score,
        "shadow_score": shadow_score,
        "score_delta": shadow_score - current_score,
        "shadow_decision": decision,
        "field_reviews": reviews,
        "hard_blockers": blockers,
        "source_fields": _source_fields(observations),
        "model_called": False,
        "runtime_effect": "影子评分只用于对比，不覆盖当前确定性分和100分正式码门禁",
    }


def empty_evidence_score_shadow(*, current_score: int = 0, reason: str = "") -> dict[str, Any]:
    return {
        "mode": "off",
        "threshold": 100,
        "current_score": int(current_score or 0),
        "shadow_score": 0,
        "score_delta": -int(current_score or 0),
        "shadow_decision": "未评估",
        "field_reviews": [],
        "hard_blockers": [],
        "source_fields": {},
        "model_called": False,
        "runtime_effect": reason or "证据影子评分未启用",
    }


def apply_model_evidence_review(
    score_shadow: dict[str, Any],
    model_result: dict[str, Any],
    *,
    matrix: dict[str, Any] | None = None,
) -> dict[str, Any]:
    score_matrix = matrix or load_evidence_score_matrix()
    merged = deepcopy(score_shadow)
    confidence = str(model_result.get("model_confidence") or "").strip().lower()
    accepted = confidence in set(score_matrix.get("model_accepted_confidences") or ["high"])
    model_reviews = {
        str(item.get("field") or ""): item
        for item in model_result.get("field_reviews") or []
    }
    for review in merged.get("field_reviews") or []:
        field_key = str(review.get("field_key") or "")
        model_review = model_reviews.get(field_key)
        if not model_review:
            continue
        review["program_verdict"] = review.get("verdict", "")
        review["program_shadow_score"] = review.get("shadow_score", 0)
        review["model_called"] = True
        review["model_verdict"] = model_review.get("verdict", "")
        review["model_source_field"] = model_review.get("source_field", "")
        review["model_evidence_text"] = model_review.get("evidence_text", "")
        review["model_reason"] = model_review.get("reason", "")
        review["model_confidence"] = confidence
        review["model_accepted"] = accepted
        if not accepted:
            continue
        verdict = str(model_review.get("verdict") or "")
        current_score = int(review.get("current_score") or 0)
        if verdict == "supported":
            shadow_score = max(current_score, int(score_matrix.get("model_supported_score") or current_score))
        else:
            shadow_score = _score_for_verdict(
                verdict,
                current_score,
                score_matrix,
                semantic_exact=False,
            )
        review["verdict"] = verdict
        review["shadow_score"] = shadow_score
        review["score_delta"] = shadow_score - current_score
        review["reason"] = f"模型证据审查：{model_review.get('reason', '')}"

    reviews = merged.get("field_reviews") or []
    threshold = int(merged.get("threshold") or score_matrix.get("gate_threshold") or 100)
    shadow_score = min([int(item.get("shadow_score") or 0) for item in reviews] or [0])
    current_score = int(merged.get("current_score") or 0)
    blockers = [
        f"{item.get('field', '')}:{item.get('verdict', '')}"
        for item in reviews
        if item.get("verdict") in {"contradicted", "missing_evidence"}
    ]
    merged.update(
        {
            "shadow_score": shadow_score,
            "score_delta": shadow_score - current_score,
            "shadow_decision": "通过" if shadow_score >= threshold and not blockers else "需标注",
            "hard_blockers": blockers,
            "model_called": True,
            "model_call_count": 1,
            "model_confidence": confidence,
            "model_hard_blockers": list(model_result.get("hard_blockers") or []),
            "runtime_effect": "模型证据审查仅更新影子对比，不能单独取得100分正式出码资格",
        }
    )
    return merged


def _review_field(
    evidence: dict[str, Any],
    *,
    semantic_evaluations: list[dict[str, Any]],
    observations: dict[str, dict[str, Any]],
    matrix: dict[str, Any],
) -> dict[str, Any]:
    field_key = str(evidence.get("field_key") or "")
    current_score = int(evidence.get("score") or 0)
    related = _related_semantic_evaluations(field_key, semantic_evaluations, matrix)
    matched = [item for item in related if item.get("status") == "命中"]
    missing = [item for item in related if item.get("status") == "缺少输入"]
    errors = [item for item in related if item.get("status") == "条件错误"]

    verdict = ""
    reason = ""
    semantic_rule_ids: list[str] = []
    semantic_evidence: list[str] = []
    semantic_exact = False
    if matched:
        semantic_rule_ids = [str(item.get("rule_id") or "") for item in matched]
        semantic_evidence = [
            str(text)
            for item in matched
            for text in item.get("evidence_texts") or []
            if str(text).strip()
        ]
        comparison = _compare_semantic_values(field_key, evidence, matched, matrix)
        verdict = comparison["verdict"]
        reason = comparison["reason"]
        semantic_exact = bool(comparison["exact"])
    elif missing:
        verdict = "missing_evidence"
        semantic_rule_ids = [str(item.get("rule_id") or "") for item in missing]
        missing_fields = sorted(
            {
                str(field)
                for item in missing
                for field in item.get("missing_fields") or []
                if str(field).strip()
            }
        )
        reason = f"已批准语义规则依赖缺失输入：{'、'.join(missing_fields)}"
    elif errors:
        verdict = "ambiguous"
        semantic_rule_ids = [str(item.get("rule_id") or "") for item in errors]
        reason = "语义条件评估出错，不能提高证据分"
    else:
        verdict, reason = _base_verdict(evidence, matrix)

    shadow_score = _score_for_verdict(
        verdict,
        current_score,
        matrix,
        semantic_exact=semantic_exact,
    )
    return {
        "field_key": field_key,
        "field": evidence.get("field", ""),
        "candidate_value": evidence.get("value", ""),
        "candidate_code": evidence.get("code", ""),
        "current_score": current_score,
        "shadow_score": shadow_score,
        "score_delta": shadow_score - current_score,
        "verdict": verdict,
        "source_field": evidence.get("source", ""),
        "evidence_text": evidence.get("evidence", ""),
        "reason": reason,
        "hit_type": evidence.get("hit_type", ""),
        "rule_id": evidence.get("rule_id", ""),
        "semantic_rule_ids": semantic_rule_ids,
        "semantic_evidence": list(dict.fromkeys(semantic_evidence)),
        "model_called": False,
    }


def _related_semantic_evaluations(
    field_key: str,
    evaluations: list[dict[str, Any]],
    matrix: dict[str, Any],
) -> list[dict[str, Any]]:
    target_map = matrix.get("semantic_target_fields") or {}
    return [
        item
        for item in evaluations
        if any(target_map.get(str(target)) == field_key for target in item.get("target_fields") or [])
    ]


def _compare_semantic_values(
    field_key: str,
    evidence: dict[str, Any],
    matched: list[dict[str, Any]],
    matrix: dict[str, Any],
) -> dict[str, Any]:
    candidate = _canonical_candidate_value(field_key, evidence)
    semantic_values: list[str] = []
    for item in matched:
        target = str((item.get("target_fields") or [""])[0])
        for value in item.get("normalized_values") or []:
            canonical = _canonical_semantic_value(target, value, matrix)
            if canonical:
                semantic_values.append(canonical)
    unique_values = sorted(set(semantic_values))
    if not unique_values or not candidate:
        return {
            "verdict": "ambiguous",
            "reason": "语义规则已命中，但当前阶段无法将语义值与候选字段做唯一对比",
            "exact": False,
        }
    if len(unique_values) > 1:
        return {
            "verdict": "ambiguous",
            "reason": f"同一字段命中多个语义值：{'、'.join(unique_values)}",
            "exact": False,
        }
    if candidate == unique_values[0]:
        return {
            "verdict": "supported",
            "reason": f"已批准语义规则与候选字段一致：{candidate}",
            "exact": True,
        }
    return {
        "verdict": "contradicted",
        "reason": f"已批准语义值={unique_values[0]}，候选字段={candidate}",
        "exact": False,
    }


def _base_verdict(evidence: dict[str, Any], matrix: dict[str, Any]) -> tuple[str, str]:
    hit_type = str(evidence.get("hit_type") or "")
    source = str(evidence.get("source") or "")
    hit_verdict = (matrix.get("hit_type_verdicts") or {}).get(hit_type)
    if hit_verdict:
        if hit_verdict == "supported":
            return "supported", f"命中方式“{hit_type}”已确认为可执行基础证据"
        return str(hit_verdict), f"命中方式“{hit_type}”不构成直接高置信证据"
    for keyword, verdict in (matrix.get("source_keyword_verdicts") or {}).items():
        if str(keyword) in source:
            if verdict == "supported":
                return "supported", f"证据来源“{source}”已确认为可执行基础证据"
            return str(verdict), f"证据来源“{source}”依赖默认或阈值推断"
    return "supported", "现有确定性解析或已批准映射提供了直接证据"


def _score_for_verdict(
    verdict: str,
    current_score: int,
    matrix: dict[str, Any],
    *,
    semantic_exact: bool,
) -> int:
    if verdict not in VERDICTS:
        raise ValueError(f"不支持的证据结论：{verdict}")
    if verdict == "supported" and semantic_exact:
        return max(current_score, int(matrix.get("semantic_supported_score") or current_score))
    config = (matrix.get("verdict_scores") or {}).get(verdict) or {}
    mode = config.get("mode")
    value = int(config.get("value") or 0)
    if mode == "preserve":
        return current_score
    if mode == "fixed":
        return value
    if mode == "cap":
        return min(current_score, value)
    raise ValueError(f"证据评分模式无效：{mode}")


def _canonical_candidate_value(field_key: str, evidence: dict[str, Any]) -> str:
    code = _normalize(evidence.get("code"))
    raw = _normalize(evidence.get("value"))
    if field_key in {"grade", "total_core", "copper_type"}:
        return code
    return raw or code


def _canonical_semantic_value(target_field: str, value: Any, matrix: dict[str, Any]) -> str:
    raw = str(value or "").strip()
    aliases = (matrix.get("semantic_value_aliases") or {}).get(target_field) or {}
    if raw in aliases:
        return _normalize(aliases[raw])
    if raw.startswith("external_lookup:"):
        return ""
    if target_field in {"grade_intent", "glue"}:
        return _normalize(raw)
    return ""


def _source_fields(observations: dict[str, dict[str, Any]]) -> dict[str, str]:
    return {
        str(field): str(item.get("value") or "")
        for field, item in observations.items()
        if item.get("available")
    }


def _normalize(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def format_field_reviews(value: list[dict[str, Any]]) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def format_source_fields(value: dict[str, Any]) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
