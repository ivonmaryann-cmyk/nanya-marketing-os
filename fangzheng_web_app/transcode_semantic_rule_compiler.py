from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from typing import Any, Protocol

from .transcode_agent_rules import GRADE_CODES


BUSINESS_FIELD_TARGET_FIELDS = {
    "胶系": {"glue"},
    "基板厚度": {"thickness", "total_core"},
    "铜箔规格": {"copper"},
    "基板尺寸": {"size"},
    "胶水类别": {"glue_category"},
    "铜箔类型+印字/非印字": {"copper_type", "print_mark"},
    "基板级别": {"grade_intent"},
    "总/芯厚": {"total_core"},
}

DETERMINISTIC_SEMANTIC_TYPES = {
    "keyword_present",
    "keyword_absent",
    "default_when_missing",
    "exclusion_set",
    "position_code",
    "alias_mapping",
    "comparison",
    "multi_condition",
    "external_lookup",
    "explicit_fact",
}

RUNTIME_INPUT_FIELDS = {
    "订单备注",
    "客户规格",
    "客户料品名称",
    "客户物料编码",
    "品号/物料编号",
    "胶系",
    "基板厚度",
    "铜箔规格",
    "订单规格",
    "订单规格/订单备注",
}

RUNTIME_FIELD_ALIASES = {
    "remark": "订单备注",
    "order_note": "订单备注",
    "order_remark": "订单备注",
    "订单备注栏": "订单备注",
    "客户规格描述": "客户规格",
    "customer_spec": "客户规格",
    "客户料品": "客户料品名称",
    "料品名称": "客户料品名称",
    "customer_material_name": "客户料品名称",
    "物料编码": "客户物料编码",
    "material_code": "客户物料编码",
    "customer_material_code": "客户物料编码",
    "品号": "品号/物料编号",
    "物料编号": "品号/物料编号",
    "glue": "胶系",
    "glue_model": "胶系",
    "thickness": "基板厚度",
    "订单基板厚度": "基板厚度",
    "copper": "铜箔规格",
    "order_spec": "订单规格",
    "订单备注/订单规格": "订单规格/订单备注",
}


class SemanticRuleClient(Protocol):
    def normalize(
        self,
        *,
        task_type: str,
        source_fields: dict[str, Any],
        customer_code: str = "",
        customer_name: str = "",
        relevant_rules: list[dict[str, Any]] | None = None,
        task_context: dict[str, Any] | None = None,
    ) -> dict[str, Any]: ...


@dataclass(frozen=True)
class SemanticRuleCompilation:
    candidate_id: str
    customer_code: str
    customer_name: str
    source_row: int
    business_field: str
    source_text: str
    required_input_fields: str
    status: str
    recommended_execution_mode: str
    validation_result: str
    business_question: str
    model_confidence: str
    semantic_type_summary: str
    target_field_summary: str
    normalization_notes: str
    raw_model_result: dict[str, Any]
    model_result: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    def model_json(self) -> str:
        return json.dumps(self.model_result, ensure_ascii=False, separators=(",", ":"))


def compile_semantic_candidate(
    candidate: dict[str, Any],
    client: SemanticRuleClient,
) -> SemanticRuleCompilation:
    source_text = str(candidate.get("source_text") or "").strip()
    if not source_text:
        return _failed_compilation(candidate, "候选原文为空")

    model_result = client.normalize(
        task_type="rule_structure",
        source_fields={"CCL特殊规则": source_text},
        customer_code=str(candidate.get("customer_code") or "").strip(),
        customer_name=str(candidate.get("customer_name") or "").strip(),
        task_context={
            "candidate_id": str(candidate.get("candidate_id") or "").strip(),
            "business_field": str(candidate.get("business_field") or "").strip(),
            "required_input_fields": str(candidate.get("required_input_fields") or "").strip(),
            "purpose": "客户特殊规则离线JSON化",
            "related_candidate_texts": candidate.get("related_candidate_texts") or [],
            "available_deterministic_assets": [
                "总芯厚转换表：订单口径为总厚且厚度<0.8时，执行总厚转芯厚",
                "含铜/不含铜口径：含铜为总厚，不含铜为芯厚",
            ],
        },
    )
    return evaluate_semantic_compilation(candidate, model_result)


def evaluate_semantic_compilation(
    candidate: dict[str, Any],
    model_result: dict[str, Any],
) -> SemanticRuleCompilation:
    raw_model_result = json.loads(json.dumps(model_result, ensure_ascii=False))
    model_result, normalization_notes = canonicalize_model_result(
        model_result,
        source_text=str(candidate.get("source_text") or "").strip(),
    )
    errors: list[str] = []
    questions: list[str] = []
    items = model_result.get("semantic_items")
    if not isinstance(items, list):
        return _failed_compilation(candidate, "semantic_items不是数组", model_result)

    expected_targets = BUSINESS_FIELD_TARGET_FIELDS.get(
        str(candidate.get("business_field") or ""), set()
    )
    semantic_types: list[str] = []
    target_fields: list[str] = []
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            errors.append(f"semantic_items[{index}]不是对象")
            continue
        semantic_type = str(item.get("semantic_type") or "").strip()
        target_field = str(item.get("target_field") or "").strip()
        if semantic_type and semantic_type not in semantic_types:
            semantic_types.append(semantic_type)
        if target_field and target_field not in target_fields:
            target_fields.append(target_field)
        if target_field in {
            "unknown",
            "combination_structure",
            "copper_vendor",
            "glass_vendor",
            "formula_code",
        }:
            errors.append(f"第{index + 1}项目标字段不属于当前CCL范围：{target_field}")
        elif expected_targets and target_field not in expected_targets:
            errors.append(
                f"第{index + 1}项目标字段{target_field}"
                f"与业务字段{candidate.get('business_field')}不一致"
            )
        _validate_condition_completeness(item, index, errors, questions)
        _validate_stated_target(item, index, errors, questions)

    for ambiguity in model_result.get("ambiguities") or []:
        if isinstance(ambiguity, dict):
            reason = str(ambiguity.get("reason") or "存在未说明的业务歧义").strip()
            questions.append(reason)
    for missing in model_result.get("missing_inputs") or []:
        if isinstance(missing, dict):
            field = str(missing.get("field") or "所需订单字段").strip()
            reason = str(missing.get("reason") or "规则定义缺少输入字段").strip()
            questions.append(f"{field}：{reason}")

    confidence = str(model_result.get("model_confidence") or "low").strip()
    if not items:
        questions.append("模型未拆出可执行语义项")
    if confidence == "low":
        questions.append("模型整体置信度为low，需要人工复核")

    if errors and questions and confidence == "low":
        status = "待业务确认"
        execution_mode = "待业务确认"
        validation_result = "结构未形成可执行条件；需要补充业务规则"
        questions.extend(errors)
    elif errors:
        status = "程序校验失败"
        execution_mode = "无法解析"
        validation_result = "失败：" + "；".join(_deduplicate(errors))
    elif questions:
        status = "待业务确认"
        execution_mode = "待业务确认"
        validation_result = "结构校验通过；存在待确认项"
    else:
        status = "程序校验通过"
        execution_mode = recommend_execution_mode(items)
        validation_result = f"通过；建议执行方式={execution_mode}"

    return SemanticRuleCompilation(
        candidate_id=str(candidate.get("candidate_id") or "").strip(),
        customer_code=str(candidate.get("customer_code") or "").strip(),
        customer_name=str(candidate.get("customer_name") or "").strip(),
        source_row=_to_int(candidate.get("source_row")),
        business_field=str(candidate.get("business_field") or "").strip(),
        source_text=str(candidate.get("source_text") or "").strip(),
        required_input_fields=str(candidate.get("required_input_fields") or "").strip(),
        status=status,
        recommended_execution_mode=execution_mode,
        validation_result=validation_result,
        business_question="；".join(_deduplicate(questions)),
        model_confidence=confidence,
        semantic_type_summary="；".join(semantic_types),
        target_field_summary="；".join(target_fields),
        normalization_notes="；".join(normalization_notes),
        raw_model_result=raw_model_result,
        model_result=model_result,
    )


def canonicalize_model_result(
    model_result: dict[str, Any],
    *,
    source_text: str = "",
) -> tuple[dict[str, Any], list[str]]:
    canonical = json.loads(json.dumps(model_result, ensure_ascii=False))
    notes: list[str] = []
    for item_index, item in enumerate(canonical.get("semantic_items") or []):
        if not isinstance(item, dict):
            continue
        raw_source_field = str(item.get("source_field") or "").strip()
        evidence_text = str(item.get("evidence_text") or "").strip()
        if (
            raw_source_field != "CCL特殊规则"
            and evidence_text
            and evidence_text in source_text
        ):
            item["source_field"] = "CCL特殊规则"
            notes.append(
                f"第{item_index + 1}项source_field：{raw_source_field}->CCL特殊规则"
            )
        normalized_conditions: list[dict[str, Any]] = []
        for condition_index, condition in enumerate(item.get("conditions") or []):
            if not isinstance(condition, dict):
                continue
            raw_scope = str(condition.get("source_scope") or "").strip()
            algorithm_value = str(condition.get("value") or "").strip()
            if raw_scope == "available_deterministic_assets" and "总厚转芯厚" in algorithm_value:
                item["normalized_value"] = "core_after_total_to_core_conversion"
                notes.append(
                    f"第{item_index + 1}项第{condition_index + 1}个条件转换为总厚转芯厚动作"
                )
                continue
            raw_field = str(condition.get("field") or "").strip()
            field = _canonical_runtime_field(raw_field)
            if field != raw_field:
                condition["field"] = field
                notes.append(
                    f"第{item_index + 1}项第{condition_index + 1}个条件field：{raw_field}->{field}"
                )
            scope = _canonical_runtime_field(raw_scope)
            if (
                raw_scope in {"CCL特殊规则", "order_field", "客户标准化字段"}
                and field in RUNTIME_INPUT_FIELDS
            ):
                scope = field
            elif field in RUNTIME_INPUT_FIELDS and scope in RUNTIME_INPUT_FIELDS and scope != field:
                scope = field
            if scope != raw_scope:
                condition["source_scope"] = scope
                notes.append(
                    f"第{item_index + 1}项第{condition_index + 1}个条件source_scope："
                    f"{raw_scope}->{scope}"
                )
            normalized_conditions.append(condition)
        item["conditions"] = normalized_conditions
    if _has_explicit_over_default_glue(canonical.get("semantic_items") or []):
        kept_ambiguities = []
        for ambiguity in canonical.get("ambiguities") or []:
            reason = str(ambiguity.get("reason") or "") if isinstance(ambiguity, dict) else ""
            field = str(ambiguity.get("field") or "") if isinstance(ambiguity, dict) else ""
            if field in {"glue", "胶系"} and ("优先级" in reason or "冲突" in reason):
                notes.append("胶系明示映射优先于缺失默认，移除伪冲突")
                continue
            kept_ambiguities.append(ambiguity)
        canonical["ambiguities"] = kept_ambiguities
    return canonical, _deduplicate(notes)


def recommend_execution_mode(items: list[dict[str, Any]]) -> str:
    if items and all(_is_deterministic_item(item) for item in items):
        return "结构化后可确定性执行"
    return "运行时需要模型"


def _is_deterministic_item(item: dict[str, Any]) -> bool:
    semantic_type = str(item.get("semantic_type") or "")
    if semantic_type not in DETERMINISTIC_SEMANTIC_TYPES:
        return False
    conditions = item.get("conditions")
    if not isinstance(conditions, list) or not conditions:
        return semantic_type == "external_lookup"
    return all(_condition_has_executable_value(condition) for condition in conditions)


def _condition_has_executable_value(condition: Any) -> bool:
    if not isinstance(condition, dict):
        return False
    operator = str(condition.get("operator") or "")
    field = str(condition.get("field") or "").strip()
    source_scope = str(condition.get("source_scope") or "").strip()
    if not field or not source_scope:
        return False
    value = condition.get("value")
    if operator in {"missing", "present"}:
        return True
    return value is not None and value != "" and value != [] and value != {}


def _validate_condition_completeness(
    item: dict[str, Any],
    item_index: int,
    errors: list[str],
    questions: list[str],
) -> None:
    conditions = item.get("conditions")
    if not isinstance(conditions, list):
        errors.append(f"第{item_index + 1}项conditions不是数组")
        return
    semantic_type = str(item.get("semantic_type") or "")
    if not conditions and semantic_type not in {"external_lookup", "out_of_scope"}:
        errors.append(f"第{item_index + 1}项缺少执行条件")
        return
    for condition_index, condition in enumerate(conditions):
        if not _condition_has_executable_value(condition):
            errors.append(f"第{item_index + 1}项第{condition_index + 1}个条件不完整")
            continue
        field = str(condition.get("field") or "").strip()
        source_scope = str(condition.get("source_scope") or "").strip()
        if field not in RUNTIME_INPUT_FIELDS:
            questions.append(
                f"第{item_index + 1}项第{condition_index + 1}个条件字段未明确：{field}"
            )
        if source_scope not in RUNTIME_INPUT_FIELDS:
            questions.append(
                f"第{item_index + 1}项第{condition_index + 1}个条件来源未明确：{source_scope}"
            )


def _validate_stated_target(
    item: dict[str, Any],
    item_index: int,
    errors: list[str],
    questions: list[str],
) -> None:
    target_field = str(item.get("target_field") or "")
    stated_target = str(item.get("stated_target_value") or "").strip().upper()
    if (
        target_field == "grade_intent"
        and stated_target
        and re.fullmatch(r"[A-Z][A-Z0-9]?", stated_target)
        and stated_target not in GRADE_CODES
    ):
        errors.append(f"第{item_index + 1}项基板级别值未在有效代码表：{stated_target}")
    if target_field == "grade_intent" and not stated_target:
        normalized = str(item.get("normalized_value") or "").strip()
        if not normalized:
            questions.append(f"第{item_index + 1}项缺少基板级别业务意图或原文目标值")


def _failed_compilation(
    candidate: dict[str, Any],
    reason: str,
    model_result: dict[str, Any] | None = None,
) -> SemanticRuleCompilation:
    return SemanticRuleCompilation(
        candidate_id=str(candidate.get("candidate_id") or "").strip(),
        customer_code=str(candidate.get("customer_code") or "").strip(),
        customer_name=str(candidate.get("customer_name") or "").strip(),
        source_row=_to_int(candidate.get("source_row")),
        business_field=str(candidate.get("business_field") or "").strip(),
        source_text=str(candidate.get("source_text") or "").strip(),
        required_input_fields=str(candidate.get("required_input_fields") or "").strip(),
        status="程序校验失败",
        recommended_execution_mode="无法解析",
        validation_result=f"失败：{reason}",
        business_question="",
        model_confidence="low",
        semantic_type_summary="",
        target_field_summary="",
        normalization_notes="",
        raw_model_result=model_result or {},
        model_result=model_result or {},
    )


def _deduplicate(values: list[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        text = str(value or "").strip()
        if text and text not in result:
            result.append(text)
    return result


def _to_int(value: Any) -> int:
    try:
        return int(float(str(value or "0").strip()))
    except ValueError:
        return 0


def _canonical_runtime_field(value: str) -> str:
    text = str(value or "").strip()
    return RUNTIME_FIELD_ALIASES.get(text, RUNTIME_FIELD_ALIASES.get(text.lower(), text))


def _has_explicit_over_default_glue(items: list[Any]) -> bool:
    has_default = False
    has_explicit = False
    for item in items:
        if not isinstance(item, dict) or item.get("target_field") != "glue":
            continue
        semantic_type = str(item.get("semantic_type") or "")
        conditions = item.get("conditions") or []
        if semantic_type == "default_when_missing" and any(
            isinstance(condition, dict) and condition.get("operator") == "missing"
            for condition in conditions
        ):
            has_default = True
        if semantic_type in {"explicit_fact", "keyword_present", "alias_mapping"} and any(
            isinstance(condition, dict)
            and condition.get("operator") in {"equals", "contains_any", "contains_all", "in"}
            for condition in conditions
        ):
            has_explicit = True
    return has_default and has_explicit
