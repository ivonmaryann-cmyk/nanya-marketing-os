from __future__ import annotations

from dataclasses import dataclass

from .db import get_transcode_model_config, save_transcode_model_config
from .transcode_semantic_service import (
    DeepSeekSemanticClient,
    SemanticModelConfig,
    SemanticModelConfigError,
)


DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-v4-pro"
DEFAULT_TIMEOUT_SECONDS = 120.0
DEFAULT_MAX_ORDER_CALLS = 50


@dataclass(frozen=True)
class UserModelConfig:
    employee_id: str
    enabled: bool = False
    base_url: str = DEFAULT_BASE_URL
    api_key: str = ""
    model: str = DEFAULT_MODEL
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS
    max_order_calls: int = DEFAULT_MAX_ORDER_CALLS

    @property
    def key_configured(self) -> bool:
        return bool(self.api_key)

    def to_runtime_config(self) -> SemanticModelConfig:
        if self.enabled and not self.api_key:
            raise SemanticModelConfigError("开启模型前请先配置 API Key。")
        return SemanticModelConfig(
            api_key=self.api_key,
            base_url=self.base_url,
            model=self.model,
            mode="active" if self.enabled else "off",
            timeout_seconds=self.timeout_seconds,
            max_order_calls=self.max_order_calls,
        )


def load_user_model_config(employee_id: str) -> UserModelConfig:
    employee_id = str(employee_id or "").strip()
    row = get_transcode_model_config(employee_id) if employee_id else None
    if not row:
        return UserModelConfig(employee_id=employee_id)
    return UserModelConfig(
        employee_id=employee_id,
        enabled=bool(row["enabled"]),
        base_url=str(row["base_url"] or DEFAULT_BASE_URL).strip(),
        api_key=str(row["api_key"] or "").strip(),
        model=str(row["model"] or DEFAULT_MODEL).strip(),
        timeout_seconds=float(row["timeout_seconds"] or DEFAULT_TIMEOUT_SECONDS),
        max_order_calls=int(row["max_order_calls"] or DEFAULT_MAX_ORDER_CALLS),
    )


def update_user_model_config(
    employee_id: str,
    *,
    enabled: bool,
    base_url: str,
    api_key: str | None,
    model: str,
) -> UserModelConfig:
    employee_id = str(employee_id or "").strip()
    base_url = str(base_url or "").strip().rstrip("/")
    model = str(model or "").strip()
    if not employee_id:
        raise ValueError("未找到当前登录用户。")
    if not base_url.startswith("https://"):
        raise ValueError("模型 API URL 必须使用 HTTPS。")
    if not model:
        raise ValueError("请填写模型名称。")
    current = load_user_model_config(employee_id)
    next_key = current.api_key if api_key is None else str(api_key).strip()
    if enabled and not next_key:
        raise ValueError("开启模型前请先配置 API Key。")
    save_transcode_model_config(
        employee_id,
        enabled=enabled,
        base_url=base_url,
        api_key=api_key,
        model=model,
        timeout_seconds=current.timeout_seconds,
        max_order_calls=current.max_order_calls,
    )
    return load_user_model_config(employee_id)


def test_user_model_connection(config: UserModelConfig) -> dict[str, str]:
    if not config.api_key:
        raise ValueError("请先填写并保存 API Key。")
    runtime = config.to_runtime_config()
    if not runtime.enabled:
        runtime = SemanticModelConfig(
            api_key=runtime.api_key,
            base_url=runtime.base_url,
            model=runtime.model,
            mode="active",
            timeout_seconds=min(runtime.timeout_seconds, 20),
            max_order_calls=runtime.max_order_calls,
        )
    result = DeepSeekSemanticClient(runtime).normalize(
        task_type="order_normalization",
        customer_code="",
        customer_name="",
        source_fields={"订单备注": "模型连接测试"},
        relevant_rules=[],
        task_context={
            "runtime_mode": "connection_test",
            "instruction": "这是连接测试，不生成制造编码。",
        },
    )
    return {
        "model": runtime.model,
        "confidence": str(result.get("model_confidence") or ""),
        "status": "连接成功",
    }
