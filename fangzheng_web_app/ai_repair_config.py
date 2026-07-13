from __future__ import annotations

import os
from dataclasses import dataclass

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - optional local convenience dependency
    load_dotenv = None


if load_dotenv is not None:
    load_dotenv()


def _env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int, *, minimum: int, maximum: int) -> int:
    raw = os.environ.get(name, "").strip()
    try:
        value = int(raw)
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(value, maximum))


@dataclass(frozen=True)
class AiRepairConfig:
    enabled: bool
    api_key: str
    base_url: str
    model: str
    timeout_seconds: int
    max_rows: int

    @property
    def available(self) -> bool:
        return self.enabled and bool(self.api_key)

    def safe_status(self) -> str:
        if not self.enabled:
            return "AI补缺未启用。"
        if not self.api_key:
            return "AI补缺已配置为启用，但缺少 DEEPSEEK_API_KEY，已自动跳过。"
        return f"AI补缺已启用：model={self.model}，timeout={self.timeout_seconds}s，max_rows={self.max_rows}。"


def get_ai_repair_config() -> AiRepairConfig:
    return AiRepairConfig(
        enabled=_env_flag("PDF_EXCEL_AI_REPAIR_ENABLED", default=False),
        api_key=os.environ.get("DEEPSEEK_API_KEY", "").strip(),
        base_url=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip().rstrip("/"),
        model=os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash").strip() or "deepseek-v4-flash",
        timeout_seconds=_env_int("PDF_EXCEL_AI_REPAIR_TIMEOUT_SECONDS", 45, minimum=5, maximum=300),
        max_rows=_env_int("PDF_EXCEL_AI_REPAIR_MAX_ROWS", 12, minimum=1, maximum=50),
    )
