from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any


DEFAULT_RISK_PATH = (
    Path(__file__).resolve().parents[1]
    / "model_skills/marketing-transcode-semantics/references/pending_formal_risks.json"
)


@lru_cache(maxsize=4)
def load_pending_formal_risks(path: str = str(DEFAULT_RISK_PATH)) -> list[dict[str, Any]]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0" or not isinstance(payload.get("risks"), list):
        raise ValueError("待确认正式风险表格式无效")
    return payload["risks"]


def match_pending_formal_risk(customer: str, *texts: str) -> dict[str, Any] | None:
    customer_norm = _normalize(customer)
    text_norm = _normalize(" ".join(str(text or "") for text in texts))
    for risk in load_pending_formal_risks():
        if _normalize(risk.get("customer")) not in customer_norm:
            continue
        required = [_normalize(value) for value in risk.get("contains_all") or []]
        if required and all(value in text_norm for value in required):
            return risk
    return None


def _normalize(value: Any) -> str:
    return re.sub(r"[\s_]+", "", str(value or "")).upper()
