from __future__ import annotations

import json
import urllib.error
import urllib.request
from typing import Any

from .ai_repair_config import AiRepairConfig


class DeepSeekRepairError(RuntimeError):
    pass


def request_repair_json(config: AiRepairConfig, payload: dict[str, Any]) -> dict[str, Any]:
    if not config.available:
        raise DeepSeekRepairError("AI repair is not available.")

    messages = [
        {
            "role": "system",
            "content": (
                "你是采购订单明细修复助手。只根据用户提供的订单正文、明细表头、原始行文本和已有字段补缺。"
                "不要处理订单头、付款条款、签核区。不要凭空编造。只返回严格 JSON。"
            ),
        },
        {
            "role": "user",
            "content": json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        },
    ]
    request_body = {
        "model": config.model,
        "messages": messages,
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    data = json.dumps(request_body, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        f"{config.base_url}/chat/completions",
        data=data,
        headers={
            "Authorization": f"Bearer {config.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=config.timeout_seconds) as response:
            response_text = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")[:500]
        raise DeepSeekRepairError(f"DeepSeek HTTP {exc.code}: {detail}") from exc
    except Exception as exc:
        raise DeepSeekRepairError(f"DeepSeek request failed: {exc}") from exc

    try:
        response_json = json.loads(response_text)
        content = response_json["choices"][0]["message"]["content"]
        return json.loads(content)
    except Exception as exc:
        raise DeepSeekRepairError("DeepSeek did not return valid repair JSON.") from exc
