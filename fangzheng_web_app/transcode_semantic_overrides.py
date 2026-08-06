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
    *,
    allow_order_remark_priority: bool = False,
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
    exact_grade_note = str((steps or {}).get("grade_note") or "")
    exact_grade_hit = exact_grade_note.startswith("基板级别写法：")
    exact_grade_code = str((steps or {}).get("step7_grade_code") or "").strip().upper()
    protected_fields = {
        str(item.get("field") or "")
        for item in (analysis.get("applied_rules") or [])
        if str(item.get("field") or "")
    }
    existing_semantic_values: dict[str, set[str]] = {}
    for existing in analysis.get("applied_rules") or []:
        if str(existing.get("rule_type") or "") not in {
            "已批准模型语义规则",
            "客户人工长期规则",
        }:
            continue
        existing_semantic_values.setdefault(str(existing.get("field") or ""), set()).add(
            str(existing.get("new") or "")
        )
    for field, candidates in proposed.items():
        # Confirmed deterministic/customer mappings have higher priority than
        # semantic normalization. The model layer may fill a gap, not overwrite it.
        top_priority = max(int(candidate.get("priority") or 0) for _, candidate in candidates)
        candidates = [
            (code, candidate)
            for code, candidate in candidates
            if int(candidate.get("priority") or 0) == top_priority
        ]
        item = candidates[0][1]
        uses_order_remark = _uses_order_remark(item)
        if field in protected_fields and not (allow_order_remark_priority and uses_order_remark):
            continue
        values = sorted({code for code, _ in candidates})
        if len(values) != 1:
            conflicts.append(f"{field}: 已批准语义规则同字段多值 {'/'.join(values)}")
            continue
        code = values[0]
        if field == "tc_code":
            explicit_total_core = _explicit_total_core_from_spec(analysis.get("spec"))
            if explicit_total_core and explicit_total_core != code:
                # Customer defaults such as "默认含铜，不含铜会备注" may fill
                # a missing value, but may never overwrite an explicit value in
                # the uploaded specification itself.
                continue
        prior_semantic = {value for value in existing_semantic_values.get(field, set()) if value}
        if prior_semantic and code not in prior_semantic:
            conflicts.append(
                f"{field}: 确定性订单语义={ '/'.join(sorted(prior_semantic)) }，"
                f"模型标准化={code}"
            )
            continue
        if field == "grade_code" and exact_grade_hit and exact_grade_code and code != exact_grade_code:
            conflicts.append(
                f"grade_code: 规格已按基板级别写法精确命中{exact_grade_code}，"
                f"订单备注模型语义{code}不得覆盖"
            )
            continue
        step_key = {
            "glue_code": "step1_glue_code",
            "thickness_code": "step2_thick_code",
            "copper_code": "step3_copper_code",
            "size_code": "step4_size_code",
            "glue_category_code": "step5_glue_cat_code",
            "copper_type_code": "step6_copper_type_code",
            "grade_code": "step7_grade_code",
            "tc_code": "step8_tc_code",
        }[field]
        old = str(steps.get(step_key, "") or "")
        if old == code:
            continue
        steps[step_key] = code
        if field == "glue_code":
            steps["glue_model"] = str((item.get("normalized_values") or [""])[0])
            errors = [error for error in steps.get("errors") or [] if "胶系" not in str(error)]
            steps["errors"] = errors
        elif field == "glue_category_code":
            steps["glue_category"] = "普通" if code == "Y" else "特殊"
        elif field == "tc_code":
            steps["order_type"] = "总厚" if code == "T" else "芯厚"
        is_manual_long_term = (
            str(item.get("model") or "").strip() == "确认中心人工规则"
            or str(item.get("source_column") or "").strip() == "确认中心"
        )
        # 模型归一化属于语义推断，命中后保留为待人工确认；
        # 只有订单备注原文直接命中规则时，才允许作为100分正式依据。
        source = (
            "已确认人工长期规则"
            if is_manual_long_term and not item.get("model_normalized")
            else "模型标准化+已确认人工长期规则"
            if is_manual_long_term and item.get("model_normalized")
            else "模型标准化+已批准语义规则"
            if item.get("model_normalized")
            else "已批准语义规则"
        )
        applied.append(
            {
                "rule_id": item.get("rule_id", ""),
                "field": field,
                "old": old,
                "new": code,
                "text": item.get("source_text", ""),
                "source": source,
                "source_row": "",
                "source_field": item.get("business_field", ""),
                "source_column": item.get("source_column", ""),
                "rule_type": "客户人工长期规则" if is_manual_long_term else "已批准模型语义规则",
                "model_normalized": bool(item.get("model_normalized")),
            }
        )
    return applied, conflicts


def _explicit_total_core_from_spec(value: Any) -> str:
    text = str(value or "").upper()
    if any(term in text for term in (
        "不含铜", "不连铜", "不連銅", "芯厚",
        "EXCLUDING COPPER", "WITHOUT COPPER", "NO COPPER",
    )):
        return "C"
    if any(term in text for term in ("含铜", "总厚", "總厚", "OVERALL", "TOTAL")):
        return "T"
    return ""


def _uses_order_remark(item: dict[str, Any]) -> bool:
    return any(
        str(condition.get("field") or "").strip() == "订单备注"
        for condition in item.get("condition_results") or []
    ) or "订单备注" in (item.get("observed_inputs") or {})


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
    if target == "copper_type" and raw.upper() in {"IGAV", "IGAV UV"}:
        return "copper_type_code", "I"
    if target == "glue":
        upper = raw.upper()
        known_codes = {
            str(value or "").strip().upper()
            for mapping_name in ("glue_exact_map", "glue_model_map")
            for value in (tables.get(mapping_name) or {}).values()
            if str(value or "").strip()
        }
        if upper in known_codes:
            return "glue_code", upper
        code = engine.get_glue_code(
            raw,
            tables.get("glue_exact_map") or {},
            tables.get("glue_model_map") or {},
            str(analysis.get("customer") or ""),
        )
        return ("glue_code", str(code)) if code else None
    if target == "total_core":
        normalized = raw.lower()
        if normalized in {"core", "芯厚", "c"}:
            return "tc_code", "C"
        if normalized in {"total", "总厚", "t"}:
            return "tc_code", "T"
    if target == "glue_category" and raw.upper() in {"Y", "R"}:
        return "glue_category_code", raw.upper()
    if target == "copper":
        compact = raw.upper().replace("/", "").replace(" ", "")
        if compact in {"FF", "JJ", "KK", "HH", "TT", "11", "22", "33"}:
            return "copper_code", compact
    if target == "thickness" and raw.isdigit() and len(raw) == 5:
        return "thickness_code", raw
    if target == "size":
        if raw.isdigit() and len(raw) == 8:
            return "size_code", raw
        if raw == "height_plus_0.3":
            steps = analysis.get("engine_steps") or {}
            try:
                width = float(steps.get("size_w"))
                height = float(steps.get("size_h")) + 0.3
            except (TypeError, ValueError):
                return None
            return "size_code", engine.size_to_code(width, height)
    return None
