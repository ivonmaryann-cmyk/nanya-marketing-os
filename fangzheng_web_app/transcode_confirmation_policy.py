from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any


DEFAULT_CONFIRMATION_RULE_PATH = (
    Path(__file__).resolve().parents[1]
    / "model_skills/marketing-transcode-semantics/references/confirmation_policy_rules.json"
)
CONFIRMATION_BASIS_TYPES = {
    "rule_conflict",
    "condition_missing",
    "non_unique_mapping",
    "approved_pending_rule",
}
FORBIDDEN_RUNTIME_REFERENCES = ("历史正确码", "测试正确码", "结果对比", "人工答案")


@lru_cache(maxsize=4)
def load_confirmation_policy_rules(
    path: str = str(DEFAULT_CONFIRMATION_RULE_PATH),
) -> list[dict[str, Any]]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0" or not isinstance(payload.get("rules"), list):
        raise ValueError("人工确认策略规则表格式无效")
    from .transcode_rule_center import merge_confirmation_policy_overrides

    rules: list[dict[str, Any]] = []
    for raw_rule in merge_confirmation_policy_overrides(payload["rules"]):
        rule = dict(raw_rule or {})
        if str(rule.get("status") or "").strip() != "approved":
            continue
        rule_id = str(rule.get("rule_id") or "").strip()
        basis_type = str(rule.get("basis_type") or "").strip()
        if not rule_id or basis_type not in CONFIRMATION_BASIS_TYPES:
            raise ValueError(f"人工确认策略规则缺少合法ID或依据类型：{rule_id or '未命名'}")
        searchable = " ".join(
            str(rule.get(key) or "")
            for key in ("reason", "source", "business_basis")
        )
        if any(term in searchable for term in FORBIDDEN_RUNTIME_REFERENCES):
            raise ValueError(f"人工确认策略不得引用测试正确码或结果对比：{rule_id}")
        rules.append(rule)
    return rules


def match_confirmation_policy_rules(customer: str, *texts: str) -> list[dict[str, Any]]:
    customer_norm = _normalize(customer)
    text_norm = _normalize(" ".join(str(text or "") for text in texts))
    matches: list[dict[str, Any]] = []
    for rule in load_confirmation_policy_rules():
        customer_aliases = rule.get("customers") or [rule.get("customer")]
        if customer_aliases and not any(
            alias and _normalize(alias) in customer_norm
            for alias in customer_aliases
        ):
            continue
        required = [_normalize(value) for value in rule.get("contains_all") or []]
        if required and not all(value in text_norm for value in required):
            continue
        groups = [
            [_normalize(value) for value in group]
            for group in rule.get("contains_any_groups") or []
        ]
        if groups and not all(any(value in text_norm for value in group) for group in groups):
            continue
        excluded = [_normalize(value) for value in rule.get("not_contains_any") or []]
        if excluded and any(value in text_norm for value in excluded):
            continue
        threshold = rule.get("copper_pair_mixed_threshold")
        if threshold is not None and not _has_mixed_copper_threshold(text_norm, float(threshold)):
            continue
        if required or groups or excluded or threshold is not None:
            matches.append(rule)
    return matches


def apply_confirmation_rules_to_evidence(
    field_evidence: list[dict[str, Any]],
    rules: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    rules_by_field: dict[str, list[dict[str, Any]]] = {}
    for rule in rules:
        for field_key in _confirmation_field_keys(rule):
            rules_by_field.setdefault(field_key, []).append(rule)
    for evidence in field_evidence:
        field_rules = rules_by_field.get(str(evidence.get("field_key") or ""))
        if not field_rules or not evidence.get("gate"):
            continue
        evidence["score"] = min(int(evidence.get("score") or 0), 99)
        evidence["decision_state"] = _basis_state(field_rules[0].get("basis_type"))
        evidence["hit_type"] = evidence["decision_state"]
        evidence["source"] = ",".join(str(rule.get("rule_id") or "") for rule in field_rules)
        evidence["evidence"] = "；".join(str(rule.get("reason") or "") for rule in field_rules)
        evidence["rule_id"] = evidence["source"]
        evidence["rule_type"] = "统一人工确认策略"
    return field_evidence


def decide_confirmation(
    *,
    errors: list[str],
    conflicts: list[str],
    candidate_code: str,
    field_evidence: list[dict[str, Any]],
) -> dict[str, Any]:
    gate_evidence = [item for item in field_evidence if item.get("gate")]
    gate_scores = [int(item.get("score") or 0) for item in gate_evidence]
    overall_score = min(gate_scores) if gate_scores else 0
    low_fields = [item for item in gate_evidence if int(item.get("score") or 0) < 100]
    triggers: list[dict[str, Any]] = []
    for item in low_fields:
        triggers.append(
            {
                "type": str(item.get("decision_state") or _decision_state_from_evidence(item)),
                "field_key": str(item.get("field_key") or ""),
                "field": str(item.get("field") or ""),
                "reason": str(item.get("evidence") or ""),
                "rule_id": str(item.get("rule_id") or ""),
            }
        )
    if errors:
        return _decision("失败", "", overall_score, "; ".join(errors), "基础字段缺失", triggers)
    if not candidate_code:
        return _decision("失败", "", overall_score, "无法生成候选编码", "基础字段缺失", triggers)
    if conflicts:
        triggers.extend(
            {"type": "规则冲突", "field_key": "", "field": "", "reason": text, "rule_id": ""}
            for text in conflicts
        )
        return _decision(
            "待确认",
            "",
            min(overall_score, 60),
            "规则冲突：" + "; ".join(conflicts),
            "规则冲突",
            triggers,
        )
    if low_fields:
        labels = list(dict.fromkeys(str(item.get("field") or "") for item in low_fields))
        reasons = list(
            dict.fromkeys(
                (
                    f"[{item.get('rule_id')}]{item.get('evidence')}"
                    if item.get("rule_id")
                    else str(item.get("evidence") or "")
                )
                for item in low_fields
                if item.get("evidence")
            )
        )
        reason = f"需人工确认字段：{', '.join(labels)}"
        if reasons:
            reason += "；" + "；".join(reasons)
        return _decision("待确认", "", overall_score, reason, "字段不确定", triggers)
    return _decision("成功", candidate_code, overall_score, "", "确定", triggers)


def _decision(
    status: str,
    formal_code: str,
    overall_score: int,
    reason: str,
    state: str,
    triggers: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "status": status,
        "formal_code": formal_code,
        "overall_score": overall_score,
        "reason": reason,
        "decision_state": state,
        "confirmation_triggers": triggers,
    }


def _confirmation_field_keys(rule: dict[str, Any]) -> list[str]:
    explicit = [str(value) for value in rule.get("field_keys") or [] if str(value)]
    if explicit:
        return explicit
    return {
        "胶系": "glue",
        "基板厚度": "thickness",
        "厚度": "thickness",
        "铜箔规格": "copper",
        "铜厚": "copper",
        "基板尺寸": "size",
        "尺寸": "size",
        "胶水类别": "glue_category",
        "铜箔类型": "copper_type",
        "基板级别": "grade",
        "总/芯厚": "total_core",
    }.get(str(rule.get("field") or "").strip(), "").split()


def _basis_state(value: Any) -> str:
    return {
        "rule_conflict": "规则冲突",
        "condition_missing": "条件不足",
        "non_unique_mapping": "多值不唯一",
        "approved_pending_rule": "规则待确认",
    }.get(str(value or ""), "规则待确认")


def _decision_state_from_evidence(evidence: dict[str, Any]) -> str:
    hit_type = str(evidence.get("hit_type") or "")
    if "模型" in hit_type:
        return "模型推断"
    if "冲突" in hit_type:
        return "规则冲突"
    if "未识别" in hit_type:
        return "条件不足"
    return "规则待确认"


def _normalize(value: Any) -> str:
    return re.sub(r"[\s_]+", "", str(value or "")).upper()


def _has_mixed_copper_threshold(text: str, threshold: float) -> bool:
    matches = re.finditer(
        r"(1\.5|0\.5|H|F|\d+(?:\.\d+)?)\s*/\s*"
        r"(1\.5|0\.5|H|F|\d+(?:\.\d+)?)",
        text,
    )
    aliases = {"H": 0.5, "F": 1.5}
    for match in matches:
        values: list[float] = []
        for token in match.groups():
            try:
                values.append(aliases[token] if token in aliases else float(token))
            except ValueError:
                values = []
                break
        if values and min(values) < threshold <= max(values):
            return True
    return False
