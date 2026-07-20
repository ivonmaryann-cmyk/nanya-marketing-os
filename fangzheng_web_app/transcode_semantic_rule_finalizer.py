from __future__ import annotations

import json
from collections import defaultdict
from copy import deepcopy
from pathlib import Path
from typing import Any


PASS_STATUS = "程序校验通过"
EXECUTION_MODE = "结构化后可确定性执行"
VALID_GRADE_CODES = {
    "AC",
    "AD",
    "AH",
    "AL",
    "AM",
    "AP",
    "AT",
    "A1",
    "A2",
    "A5",
    "F1",
}
DEFAULT_OVERRIDE_PATH = (
    Path(__file__).resolve().parents[1]
    / "model_skills/customer-special-rule-maintenance/references/semantic_rule_atomic_overrides.json"
)


def build_atomic_semantic_rule_rows(
    audit_payload: dict[str, Any],
    *,
    approval_basis: str,
    note: str,
    overrides: dict[str, Any] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    items = list(audit_payload.get("items") or [])
    model = str(audit_payload.get("model") or "").strip()
    prompt_sha256 = str(audit_payload.get("prompt_sha256") or "").strip()
    atomic_rows: list[dict[str, Any]] = []
    pending_rows: list[dict[str, Any]] = []
    override_config = overrides if overrides is not None else load_atomic_overrides()
    parent_index = 0
    for candidate in items:
        if candidate.get("status") != PASS_STATUS:
            pending_rows.append(candidate)
            continue
        parent_index += 1
        semantic_items = list((candidate.get("model_result") or {}).get("semantic_items") or [])
        if not semantic_items:
            raise ValueError(f"{candidate.get('candidate_id')} 没有语义项")
        for item_index, semantic_item in enumerate(semantic_items, start=1):
            rule_id = f"TSR-{parent_index:05d}"
            if len(semantic_items) > 1:
                rule_id = f"{rule_id}-{item_index:02d}"
            row = _build_atomic_row(
                candidate,
                semantic_item,
                rule_id=rule_id,
                model=model,
                prompt_sha256=prompt_sha256,
                approval_basis=approval_basis,
                note=note,
                item_index=item_index,
                overrides=override_config,
            )
            atomic_rows.append(row)
    validate_atomic_semantic_rule_rows(atomic_rows)
    return atomic_rows, pending_rows


def load_atomic_overrides(path: str | Path = DEFAULT_OVERRIDE_PATH) -> dict[str, Any]:
    override_path = Path(path)
    payload = json.loads(override_path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0":
        raise ValueError(f"原子规则修正配置版本无效：{override_path}")
    return payload


def validate_atomic_semantic_rule_rows(rows: list[dict[str, Any]]) -> None:
    seen_rule_ids: set[str] = set()
    for row_number, row in enumerate(rows, start=2):
        rule_id = str(row.get("规则ID") or "").strip()
        if not rule_id or rule_id in seen_rule_ids:
            raise ValueError(f"第{row_number}行规则ID为空或重复：{rule_id}")
        seen_rule_ids.add(rule_id)
        semantic_types = _split_values(row.get("语义类型"))
        target_fields = _split_values(row.get("目标字段"))
        normalized_values = _split_values(row.get("标准语义值"))
        if len(semantic_types) != 1:
            raise ValueError(f"第{row_number}行必须只有一个语义类型")
        if len(target_fields) != 1:
            raise ValueError(f"第{row_number}行必须只有一个目标字段")
        if len(normalized_values) != 1:
            raise ValueError(f"第{row_number}行必须只有一个标准语义值")
        conditions = _parse_conditions(row.get("条件JSON"), row_number)
        validate_atomic_conditions(conditions, context=f"第{row_number}行")


def validate_atomic_conditions(conditions: list[dict[str, Any]], *, context: str) -> None:
    if not conditions:
        raise ValueError(f"{context}没有条件")
    serialized: set[str] = set()
    by_field: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for condition in conditions:
        if not isinstance(condition, dict):
            raise ValueError(f"{context}条件不是对象")
        field = str(condition.get("field") or "").strip()
        operator = str(condition.get("operator") or "").strip()
        if not field or not operator:
            raise ValueError(f"{context}条件缺少field或operator")
        fingerprint = json.dumps(condition, ensure_ascii=False, sort_keys=True)
        if fingerprint in serialized:
            raise ValueError(f"{context}存在重复条件：{field}/{operator}")
        serialized.add(fingerprint)
        by_field[field].append(condition)

    for field, field_conditions in by_field.items():
        operators = {str(item.get("operator") or "") for item in field_conditions}
        if "missing" in operators and len(operators) > 1:
            raise ValueError(f"{context}{field}的missing与其他条件互斥")
        equal_values = {
            _normalized_json_value(item.get("value"))
            for item in field_conditions
            if item.get("operator") == "equals"
        }
        if len(equal_values) > 1:
            raise ValueError(f"{context}{field}存在多个互斥equals")
        contains_values = {
            value
            for item in field_conditions
            if item.get("operator") in {"contains_any", "contains_all"}
            for value in _condition_values(item.get("value"))
        }
        excluded_values = {
            value
            for item in field_conditions
            if item.get("operator") == "not_contains"
            for value in _condition_values(item.get("value"))
        }
        overlap = contains_values & excluded_values
        if overlap:
            raise ValueError(f"{context}{field}同时要求包含和不包含：{sorted(overlap)}")


def _build_atomic_row(
    candidate: dict[str, Any],
    semantic_item: dict[str, Any],
    *,
    rule_id: str,
    model: str,
    prompt_sha256: str,
    approval_basis: str,
    note: str,
    item_index: int,
    overrides: dict[str, Any],
) -> dict[str, Any]:
    item = deepcopy(semantic_item)
    candidate_id = str(candidate.get("candidate_id") or "").strip()
    conditions = _normalize_conditions(
        candidate_id,
        item_index,
        item.get("conditions") or [],
        overrides,
    )
    normalized_value = _canonical_normalized_value(candidate, item, item_index, overrides)
    priority = _priority_for(item, conditions)
    row = {
        "规则ID": rule_id,
        "来源候选ID": candidate_id,
        "启用": "是",
        "客户代码": str(candidate.get("customer_code") or "").strip(),
        "客户简称": str(candidate.get("customer_name") or "").strip(),
        "来源行号": candidate.get("source_row"),
        "业务字段": str(candidate.get("business_field") or "").strip(),
        "规则原文": str(candidate.get("source_text") or "").strip(),
        "语义类型": str(item.get("semantic_type") or "").strip(),
        "目标字段": str(item.get("target_field") or "").strip(),
        "标准语义值": normalized_value,
        "原文目标值": str(item.get("stated_target_value") or "").strip(),
        "条件JSON": json.dumps(conditions, ensure_ascii=False, separators=(",", ":")),
        "所需订单字段": "；".join(dict.fromkeys(str(item.get("field") or "").strip() for item in conditions)),
        "执行方式": EXECUTION_MODE,
        "优先级": priority,
        "模型版本": model,
        "提示词SHA256": prompt_sha256,
        "原文证据": str(item.get("evidence_text") or "").strip(),
        "业务确认": "确认",
        "确认依据": approval_basis,
        "备注": note,
    }
    validate_atomic_semantic_rule_rows([row])
    return row


def _canonical_normalized_value(
    candidate: dict[str, Any],
    item: dict[str, Any],
    item_index: int,
    overrides: dict[str, Any],
) -> str:
    candidate_id = str(candidate.get("candidate_id") or "").strip()
    override_key = _override_key(candidate_id, item_index)
    value_override = (overrides.get("normalized_value_overrides") or {}).get(override_key)
    if str(value_override or "").strip():
        return str(value_override).strip()
    target_field = str(item.get("target_field") or "").strip()
    normalized = str(item.get("normalized_value") or "").strip()
    stated = str(item.get("stated_target_value") or "").strip()
    if target_field == "grade_intent":
        if stated in VALID_GRADE_CODES:
            return stated
        if normalized.upper() == "AUTOMOTIVE" or normalized == "汽车板":
            return "AC"
        if normalized.upper() == "MINILED":
            return "AM"
    if target_field == "total_core":
        mapping = {
            "总厚": "total",
            "total": "total",
            "芯厚": "core",
            "core": "core",
            "core_thickness": "core",
            "需要进行芯总厚转换": "core_after_total_to_core_conversion",
            "执行总芯厚转换": "core_after_total_to_core_conversion",
            "默认总厚（订单未备注芯厚）且基板厚度<0.8时需执行总芯厚转换": "core_after_total_to_core_conversion",
        }
        if normalized in mapping:
            return mapping[normalized]
    return normalized or stated


def _normalize_conditions(
    candidate_id: str,
    item_index: int,
    conditions: list[dict[str, Any]],
    overrides: dict[str, Any],
) -> list[dict[str, Any]]:
    override_key = _override_key(candidate_id, item_index)
    replacements = overrides.get("condition_replacements") or {}
    additions = overrides.get("condition_additions") or {}
    if override_key in replacements:
        result = [deepcopy(condition) for condition in replacements[override_key]]
    else:
        result = [deepcopy(condition) for condition in conditions]
    result.extend(deepcopy(condition) for condition in additions.get(override_key, []))
    return _deduplicate_conditions(result)


def _priority_for(item: dict[str, Any], conditions: list[dict[str, Any]]) -> int:
    semantic_type = str(item.get("semantic_type") or "")
    base = 90 if semantic_type in {"default_when_missing", "keyword_absent"} else 100
    return base + max(0, len(conditions) - 1) * 10


def _override_key(candidate_id: str, item_index: int) -> str:
    return f"{candidate_id}#{item_index}"


def _deduplicate_conditions(conditions: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[str] = set()
    result: list[dict[str, Any]] = []
    for condition in conditions:
        fingerprint = json.dumps(condition, ensure_ascii=False, sort_keys=True)
        if fingerprint in seen:
            continue
        seen.add(fingerprint)
        result.append(condition)
    return result


def _parse_conditions(value: Any, row_number: int) -> list[dict[str, Any]]:
    if isinstance(value, list):
        return value
    try:
        parsed = json.loads(str(value or "[]"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"第{row_number}行条件JSON无效") from exc
    if not isinstance(parsed, list):
        raise ValueError(f"第{row_number}行条件JSON必须是数组")
    return parsed


def _condition_values(value: Any) -> set[str]:
    values = value if isinstance(value, list) else [value]
    return {str(item).strip().upper() for item in values if str(item or "").strip()}


def _normalized_json_value(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def _split_values(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item or "").strip()]
    return [item.strip() for item in str(value or "").replace(",", "；").split("；") if item.strip()]
