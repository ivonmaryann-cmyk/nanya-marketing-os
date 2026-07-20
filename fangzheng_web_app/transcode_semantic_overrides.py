from __future__ import annotations

import os
from typing import Any


SEMANTIC_OVERRIDE_MODES = {"off", "enforce"}


def get_semantic_override_mode(environ: dict[str, str] | None = None) -> str:
    env = os.environ if environ is None else environ
    mode = str(env.get("TRANSCODE_SEMANTIC_OVERRIDE_MODE") or "enforce").strip().lower()
    return mode if mode in SEMANTIC_OVERRIDE_MODES else "off"


def apply_confirmed_semantic_overrides(
    engine,
    tables: dict[str, Any],
    analysis: dict[str, Any],
    evaluations: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[str]]:
    matched = [item for item in evaluations if item.get("status") == "命中"]
    proposed: dict[str, list[tuple[str, dict[str, Any]]]] = {}
    for item in matched:
        targets = list(item.get("target_fields") or [])
        values = list(item.get("normalized_values") or [])
        if len(targets) != 1 or len(values) != 1:
            continue
        resolved = _resolve_override(engine, tables, targets[0], values[0], analysis)
        if resolved is None:
            continue
        field, code = resolved
        proposed.setdefault(field, []).append((code, item))

    conflicts: list[str] = []
    applied: list[dict[str, Any]] = []
    steps = analysis.get("engine_steps") or {}
    protected_fields = {
        str(item.get("field") or "")
        for item in (analysis.get("applied_rules") or [])
        if str(item.get("field") or "")
    }
    for field, candidates in proposed.items():
        # Confirmed deterministic/customer mappings have higher priority than
        # semantic normalization. The model layer may fill a gap, not overwrite it.
        if field in protected_fields:
            continue
        values = sorted({code for code, _ in candidates})
        if len(values) != 1:
            conflicts.append(f"{field}: 已批准语义规则同字段多值 {'/'.join(values)}")
            continue
        code = values[0]
        item = candidates[0][1]
        step_key = {
            "glue_code": "step1_glue_code",
            "copper_type_code": "step6_copper_type_code",
            "grade_code": "step7_grade_code",
        }[field]
        old = str(steps.get(step_key, "") or "")
        if old == code:
            continue
        steps[step_key] = code
        if field == "glue_code":
            steps["glue_model"] = str((item.get("normalized_values") or [""])[0])
            errors = [error for error in steps.get("errors") or [] if "胶系" not in str(error)]
            steps["errors"] = errors
        applied.append(
            {
                "rule_id": item.get("rule_id", ""),
                "field": field,
                "old": old,
                "new": code,
                "text": item.get("source_text", ""),
                "source": "已批准语义规则",
                "source_row": "",
                "source_field": item.get("business_field", ""),
                "rule_type": "已批准模型语义规则",
            }
        )
    return applied, conflicts


def _resolve_override(engine, tables: dict[str, Any], target: str, value: Any, analysis: dict[str, Any]):
    raw = str(value or "").strip()
    if target == "grade_intent":
        code = raw.upper()
        valid_codes = {str(item).upper() for item in (tables.get("grade_code_map") or {}).keys()}
        return ("grade_code", code) if code and code in valid_codes else None
    if target == "print_mark" and raw == "有水印":
        return "copper_type_code", "Q"
    if target == "copper_type" and raw.upper() in {"RTF", "HTE"}:
        return "copper_type_code", engine.get_copper_type_code(raw)
    if target == "glue":
        code = engine.get_glue_code(
            raw,
            tables.get("glue_exact_map") or {},
            tables.get("glue_model_map") or {},
            str(analysis.get("customer") or ""),
        )
        return ("glue_code", str(code)) if code else None
    return None
