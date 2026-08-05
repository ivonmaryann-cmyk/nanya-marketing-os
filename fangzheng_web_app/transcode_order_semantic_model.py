from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from typing import Any

from .db import get_setting
from .transcode_semantic_service import DeepSeekSemanticClient, load_semantic_model_config
from .transcode_model_config import load_user_model_config


GLOBAL_ORDER_MODEL_SETTING = "transcode_agent_order_model_global_enabled"


@dataclass(frozen=True)
class OrderSemanticRuntime:
    mode: str
    client: DeepSeekSemanticClient | None
    model: str = ""
    max_calls: int = 0
    load_error: str = ""


def is_order_model_globally_enabled() -> bool:
    value = str(get_setting(GLOBAL_ORDER_MODEL_SETTING, "0") or "0").strip().lower()
    return value in {"1", "true", "yes", "on"}


def load_order_semantic_runtime(employee_id: str = "") -> OrderSemanticRuntime:
    try:
        if is_order_model_globally_enabled():
            config = _global_order_model_config(employee_id)
            return OrderSemanticRuntime(
                mode="active",
                client=DeepSeekSemanticClient(config),
                model=config.model,
                max_calls=config.max_order_calls,
            )
        config = (
            load_user_model_config(employee_id).to_runtime_config()
            if str(employee_id or "").strip()
            else load_semantic_model_config()
        )
    except Exception as exc:
        return OrderSemanticRuntime(mode="off", client=None, load_error=str(exc))
    if not config.enabled:
        return OrderSemanticRuntime(mode="off", client=None, max_calls=config.max_order_calls)
    return OrderSemanticRuntime(
        mode=config.mode,
        client=DeepSeekSemanticClient(config),
        model=config.model,
        max_calls=config.max_order_calls,
    )


def _global_order_model_config(employee_id: str):
    from .transcode_semantic_service import SemanticModelConfig

    if str(employee_id or "").strip():
        user = load_user_model_config(employee_id)
        if user.api_key:
            return SemanticModelConfig(
                api_key=user.api_key,
                base_url=user.base_url,
                model=user.model,
                mode="active",
                timeout_seconds=user.timeout_seconds,
                max_order_calls=user.max_order_calls,
            )
    base = load_semantic_model_config()
    return SemanticModelConfig(
        api_key=base.api_key,
        base_url=base.base_url,
        model=base.model,
        mode="active",
        timeout_seconds=base.timeout_seconds,
        max_order_calls=base.max_order_calls,
    )


def should_normalize_order(
    semantic_evaluations: list[dict[str, Any]],
    order_remark: str,
) -> bool:
    if not str(order_remark or "").strip():
        return False
    return any(
        item.get("rule_id")
        and item.get("status") != "条件错误"
        and _evaluation_uses_order_remark(item)
        for item in semantic_evaluations
    )


def normalize_order_shadow(
    runtime: OrderSemanticRuntime,
    *,
    customer_code: str,
    customer_name: str,
    source_fields: dict[str, str],
    semantic_evaluations: list[dict[str, Any]],
) -> dict[str, Any]:
    if runtime.client is None:
        raise RuntimeError(runtime.load_error or "DeepSeek订单语义模型未启用")
    relevant_rules = [_compact_rule(item) for item in semantic_evaluations if item.get("rule_id")]
    return runtime.client.normalize(
        task_type="order_normalization",
        source_fields=source_fields,
        customer_code=customer_code,
        customer_name=customer_name,
        relevant_rules=relevant_rules,
        task_context={
            "runtime_mode": "shadow",
            "instruction": "只做订单口语标准化，不输出制造码，不覆盖正式转码结果",
        },
    )


def build_model_rule_evaluations(
    normalized: dict[str, Any],
    semantic_evaluations: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[str]]:
    """Bind model-normalized facts back to approved rules.

    The model may normalize wording, but a formal override is only allowed when
    the normalized target/value maps to an approved customer rule and every
    non-remark condition of that rule is already deterministically satisfied.
    """
    if str(normalized.get("model_confidence") or "") != "high":
        return [], ["模型整体置信度非high"]
    if normalized.get("ambiguities"):
        return [], ["模型返回语义歧义"]

    matched: list[dict[str, Any]] = []
    notes: list[str] = []
    for item in normalized.get("semantic_items") or []:
        if str(item.get("confidence") or "") != "high":
            notes.append(f"跳过非high语义：{item.get('target_field') or '未知字段'}")
            continue
        target = _normalize(item.get("target_field"))
        value = _normalize(item.get("normalized_value"))
        if not target or not value:
            continue
        candidates = [
            evaluation
            for evaluation in semantic_evaluations
            if _evaluation_uses_order_remark(evaluation)
            and len(evaluation.get("target_fields") or []) == 1
            and len(evaluation.get("normalized_values") or []) == 1
            and _normalize((evaluation.get("target_fields") or [""])[0]) == target
            and _normalize((evaluation.get("normalized_values") or [""])[0]) == value
            and _non_remark_conditions_match(evaluation)
        ]
        if not candidates:
            notes.append(
                f"未找到已批准规则：{item.get('target_field')}={item.get('normalized_value')}"
            )
            continue
        for evaluation in candidates:
            linked = dict(evaluation)
            linked["status"] = "命中"
            linked["model_normalized"] = True
            linked["model_evidence_text"] = str(item.get("evidence_text") or "")
            linked["model_confidence"] = str(item.get("confidence") or "")
            linked["note"] = "模型仅完成订单备注语义标准化；覆盖值来自已批准语义规则"
            matched.append(linked)
    unique: dict[str, dict[str, Any]] = {}
    for item in matched:
        unique[str(item.get("rule_id") or "")] = item
    return list(unique.values()), notes


def build_order_semantic_cache_key(
    customer_code: str,
    customer_name: str,
    source_fields: dict[str, str],
    semantic_evaluations: list[dict[str, Any]],
) -> str:
    payload = {
        "customer_code": str(customer_code or "").strip(),
        "customer_name": str(customer_name or "").strip(),
        "source_fields": source_fields,
        "rule_ids": sorted(
            str(item.get("rule_id") or "")
            for item in semantic_evaluations
            if item.get("rule_id")
        ),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def source_fields_from_observations(
    observations: dict[str, dict[str, Any]],
) -> dict[str, str]:
    return {
        str(field): str(item.get("value") or "").strip()
        for field, item in observations.items()
        if item.get("available") and str(item.get("value") or "").strip()
    }


def order_remark_source_fields(
    observations: dict[str, dict[str, Any]],
) -> dict[str, str]:
    observation = observations.get("订单备注") or {}
    value = str(observation.get("value") or "").strip()
    return {"订单备注": value} if value else {}


def _compact_rule(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "rule_id": item.get("rule_id", ""),
        "business_field": item.get("business_field", ""),
        "target_fields": item.get("target_fields") or [],
        "normalized_values": item.get("normalized_values") or [],
        "stated_target_values": item.get("stated_target_values") or [],
        "status": item.get("status", ""),
        "missing_fields": item.get("missing_fields") or [],
        "source_text": item.get("source_text", ""),
        "evidence_texts": item.get("evidence_texts") or [],
    }


def _evaluation_uses_order_remark(item: dict[str, Any]) -> bool:
    return any(
        str(condition.get("field") or "").strip() == "订单备注"
        for condition in item.get("condition_results") or []
    ) or "订单备注" in (item.get("observed_inputs") or {})


def _non_remark_conditions_match(item: dict[str, Any]) -> bool:
    return all(
        bool(condition.get("matched"))
        for condition in item.get("condition_results") or []
        if str(condition.get("field") or "").strip() != "订单备注"
    )


def _normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    return re.sub(r"\s+", "", text)
