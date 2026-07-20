from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any

from .transcode_semantic_service import DeepSeekSemanticClient, load_semantic_model_config


@dataclass(frozen=True)
class OrderSemanticRuntime:
    mode: str
    client: DeepSeekSemanticClient | None
    model: str = ""
    max_calls: int = 0
    load_error: str = ""


def load_order_semantic_runtime() -> OrderSemanticRuntime:
    try:
        config = load_semantic_model_config()
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


def should_normalize_order(
    semantic_evaluations: list[dict[str, Any]],
    source_fields: dict[str, str],
) -> bool:
    if not source_fields:
        return False
    return any(item.get("rule_id") and item.get("status") != "条件错误" for item in semantic_evaluations)


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
