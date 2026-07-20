from __future__ import annotations

import json
import os
import re
import unicodedata
from typing import Any


RUNTIME_MODES = {"off", "shadow"}
SHADOW_STATUS_MATCHED = "命中"
SHADOW_STATUS_NOT_MATCHED = "未命中"
SHADOW_STATUS_MISSING_INPUT = "缺少输入"
SHADOW_STATUS_ERROR = "条件错误"


def get_semantic_rule_runtime_mode(environ: dict[str, str] | None = None) -> str:
    env = os.environ if environ is None else environ
    mode = str(env.get("TRANSCODE_SEMANTIC_RULE_RUNTIME_MODE") or "shadow").strip().lower()
    return mode if mode in RUNTIME_MODES else "off"


def evaluate_semantic_shadow(
    rules: list[dict[str, Any]],
    *,
    customer_code: str,
    customer_name: str,
    observations: dict[str, dict[str, Any]],
    excel_row: int | None = None,
    spec: str = "",
) -> list[dict[str, Any]]:
    evaluations: list[dict[str, Any]] = []
    for rule in sorted(rules, key=lambda item: int(item.get("priority") or 0), reverse=True):
        if not rule.get("enabled") or not _customer_matches(rule, customer_code, customer_name):
            continue
        evaluations.append(
            _evaluate_rule(
                rule,
                observations,
                excel_row=excel_row,
                customer_code=customer_code,
                customer_name=customer_name,
                spec=spec,
            )
        )
    return evaluations


def _evaluate_rule(
    rule: dict[str, Any],
    observations: dict[str, dict[str, Any]],
    *,
    excel_row: int | None,
    customer_code: str,
    customer_name: str,
    spec: str,
) -> dict[str, Any]:
    missing_fields: list[str] = []
    condition_results: list[dict[str, Any]] = []
    errors: list[str] = []
    for condition in rule.get("conditions") or []:
        field = str(condition.get("field") or "").strip()
        observation = observations.get(field) or {"available": False, "value": "", "sources": []}
        if not observation.get("available"):
            if field and field not in missing_fields:
                missing_fields.append(field)
            condition_results.append(
                {
                    "field": field,
                    "operator": condition.get("operator", ""),
                    "expected": condition.get("value"),
                    "actual": "",
                    "matched": False,
                    "reason": "输入字段不存在",
                }
            )
            continue
        try:
            matched = _condition_matches(condition, observation.get("value"), rule.get("source_text", ""))
            reason = "条件成立" if matched else "条件不成立"
        except (TypeError, ValueError) as exc:
            matched = False
            reason = f"条件解析失败：{exc}"
            errors.append(f"{field}:{exc}")
        condition_results.append(
            {
                "field": field,
                "operator": condition.get("operator", ""),
                "expected": condition.get("value"),
                "actual": observation.get("value"),
                "matched": matched,
                "reason": reason,
            }
        )

    if errors:
        status = SHADOW_STATUS_ERROR
    elif missing_fields:
        status = SHADOW_STATUS_MISSING_INPUT
    elif condition_results and all(item["matched"] for item in condition_results):
        status = SHADOW_STATUS_MATCHED
    else:
        status = SHADOW_STATUS_NOT_MATCHED

    observed = {
        field: {
            "value": observations.get(field, {}).get("value", ""),
            "sources": observations.get(field, {}).get("sources", []),
            "available": bool(observations.get(field, {}).get("available")),
        }
        for field in dict.fromkeys(
            str(condition.get("field") or "").strip() for condition in rule.get("conditions") or []
        )
        if field
    }
    return {
        "row": excel_row,
        "customer_code": customer_code,
        "customer": customer_name,
        "spec": spec,
        "rule_id": rule.get("rule_id", ""),
        "source_candidate_id": rule.get("source_candidate_id", ""),
        "business_field": rule.get("business_field", ""),
        "target_fields": rule.get("target_fields", []),
        "normalized_values": rule.get("normalized_values", []),
        "stated_target_values": rule.get("stated_target_values", []),
        "status": status,
        "missing_fields": missing_fields,
        "condition_results": condition_results,
        "observed_inputs": observed,
        "source_text": rule.get("source_text", ""),
        "evidence_texts": rule.get("evidence_texts", []),
        "model": rule.get("model", ""),
        "note": "影子结果只用于证据观察，不覆盖编码或评分",
    }


def _condition_matches(condition: dict[str, Any], actual: Any, source_text: str) -> bool:
    operator = str(condition.get("operator") or "").strip()
    expected = condition.get("value")
    actual_text = _normalize(actual)
    if operator == "contains_any":
        values = _as_list(expected)
        return any(_normalize(value) in actual_text for value in values if _normalize(value))
    if operator == "contains_all":
        values = [value for value in _as_list(expected) if _normalize(value)]
        return bool(values) and all(_normalize(value) in actual_text for value in values)
    if operator == "not_contains":
        values = _as_list(expected)
        return all(_normalize(value) not in actual_text for value in values if _normalize(value))
    if operator == "equals":
        return actual_text == _normalize(expected)
    if operator == "not_equals":
        return actual_text != _normalize(expected)
    if operator == "in":
        return actual_text in {_normalize(value) for value in _as_list(expected)}
    if operator == "not_in":
        return actual_text not in {_normalize(value) for value in _as_list(expected)}
    if operator in {"lt", "lte", "gt", "gte"}:
        actual_number = _to_float(actual)
        expected_number = _to_float(expected)
        if operator == "lt":
            return actual_number < expected_number
        if operator == "lte":
            return actual_number <= expected_number
        if operator == "gt":
            return actual_number > expected_number
        return actual_number >= expected_number
    if operator == "missing":
        return not actual_text
    if operator == "present":
        return bool(actual_text)
    if operator == "char_at":
        position, values = _parse_char_at(expected, source_text)
        if not actual_text or not values:
            return False
        if position == -1:
            actual_char = actual_text[-1]
        elif position > 0 and len(actual_text) >= position:
            actual_char = actual_text[position - 1]
        else:
            return False
        return actual_char in {_normalize(value) for value in values}
    raise ValueError(f"不支持的操作符{operator}")


def _parse_char_at(value: Any, source_text: str) -> tuple[int, list[Any]]:
    if isinstance(value, dict):
        position = int(value.get("position") or 0)
        return position, _as_list(value.get("values"))
    source = str(source_text or "")
    if "最后一位" in str(value or "") or "最后一位" in source or "末位" in source:
        match = re.search(r"(?:=|是)\s*([A-Za-z0-9])", str(value or "") + source)
        return -1, [match.group(1)] if match else _as_list(value)
    position_match = re.search(r"第\s*(\d+)\s*码", source)
    position = int(position_match.group(1)) if position_match else 0
    return position, _as_list(value)


def _customer_matches(rule: dict[str, Any], customer_code: str, customer_name: str) -> bool:
    rule_codes = set(re.findall(r"\d+", str(rule.get("customer_code") or "")))
    actual_code = str(customer_code or "").strip()
    if rule_codes and actual_code:
        return actual_code in rule_codes
    rule_name = _normalize(rule.get("customer_name"))
    actual_name = _normalize(customer_name)
    return bool(rule_name and actual_name and (rule_name in actual_name or actual_name in rule_name))


def _as_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    return [value]


def _normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    return re.sub(r"\s+", "", text)


def _to_float(value: Any) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(value or ""))
    if not match:
        raise ValueError(f"无法转换为数值：{value}")
    return float(match.group(0))


def format_condition_results(value: list[dict[str, Any]]) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def format_observed_inputs(value: dict[str, Any]) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
