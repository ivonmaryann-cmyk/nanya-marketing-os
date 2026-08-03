from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import (  # noqa: E402
    _apply_agent_rules,
    _load_runtime,
    _norm_match,
)


def main() -> None:
    rules = [
        rule
        for rule in _load_runtime()[2]
        if rule.get("启用") == "是"
        and rule.get("待确认") != "是"
        and rule.get("强制执行") == "是"
        and rule.get("物料类别") == "CCL"
    ]
    failures = []
    for rule in rules:
        spec, steps = _synthetic_case(rule)
        applied, conflicts = _apply_agent_rules(
            [rule],
            str(rule.get("客户代码") or ""),
            str(rule.get("客户简称") or ""),
            spec,
            "",
            steps,
            [],
        )
        expected_field = str(rule.get("覆盖字段") or "")
        expected_value = str(rule.get("覆盖值") or "").upper()
        matching = [item for item in applied if item.get("field") == expected_field]
        if conflicts or not matching or matching[-1].get("new") != expected_value:
            failures.append(
                {
                    "rule_id": rule.get("规则ID"),
                    "customer": rule.get("客户简称"),
                    "field": expected_field,
                    "expected": expected_value,
                    "spec": spec,
                    "applied": applied,
                    "conflicts": conflicts,
                }
            )
    assert not failures, failures[:20]
    print(f"all confirmed Agent rules smoke passed active={len(rules)} reachable={len(rules)}")


def _synthetic_case(rule: dict) -> tuple[str, dict]:
    glue = _first_condition(rule.get("条件胶系"))
    thickness_mm, thickness_raw = _thickness_value(rule.get("条件厚度"))
    copper = _copper_value(rule)
    size = _first_condition(rule.get("条件尺寸"), separators=r"[,，;；]+") or "37*49"
    keyword = str(rule.get("条件关键词") or "").strip()
    parts = []
    if (
        keyword
        and "没有Q" not in keyword
        and keyword != "__NO_TG__"
        and _norm_match(keyword) != _norm_match(glue)
    ):
        parts.append(keyword)
    parts.extend([glue, thickness_raw, copper, size])
    if not any(
        rule.get(key)
        for key in ("条件胶系", "条件厚度", "条件铜厚", "条件尺寸", "条件关键词")
    ):
        parts.append(str(rule.get("条件文本") or ""))
    spec = " ".join(part for part in parts if part)

    required_code = re.search(
        r"胶系代码\s*[:：=]\s*([A-Z0-9]{2})",
        str(rule.get("条件文本") or ""),
        re.IGNORECASE,
    )
    size_parts = size.split("*", 1)
    steps = {
        "glue_model": glue,
        "thickness_mm": thickness_mm,
        "thickness_raw": thickness_raw,
        "copper_spec_raw": copper,
        "size_w": size_parts[0] if len(size_parts) == 2 else "37",
        "size_h": size_parts[1] if len(size_parts) == 2 else "49",
        "step1_glue_code": required_code.group(1).upper() if required_code else "??",
        "step2_thick_code": "?????",
        "step3_copper_code": "??",
        "step4_size_code": "????????",
        "step5_glue_cat_code": "Y",
        "step6_copper_type_code": "W",
        "step7_grade_code": "A1",
        "step8_tc_code": "T",
        "step9_struct_code": "*",
    }
    return spec, steps


def _first_condition(value, *, separators: str = r"[/,，;；]+") -> str:
    return next(
        (item.strip() for item in re.split(separators, str(value or "")) if item.strip()),
        "",
    )


def _thickness_value(value) -> tuple[float, str]:
    source = str(value or "").strip().upper()
    if not source:
        return 0.8, "0.8mm"
    if "43MIL" in source:
        return 1.0922, "43MIL"
    match = re.search(r"(>=|<=|>|<|≥|≤)?(\d+(?:\.\d+)?)", source)
    if not match:
        return 0.8, source
    operator, number = match.groups()
    actual = float(number)
    if operator in (">", "≥"):
        actual += 0.1
    elif operator in ("<", "≤"):
        actual -= 0.1
    return actual, str(actual)


def _copper_value(rule: dict) -> str:
    source = str(rule.get("条件铜厚") or "").upper()
    if (
        rule.get("覆盖字段") == "copper_code"
        and rule.get("覆盖值") == "FF"
        and rule.get("条件关键词") == "5"
    ):
        return "R/R"
    if "<1OZ" in source:
        return "H/H"
    for token in ("W/W", "R/R", "H/H", "1/1", "2/2"):
        if token in source:
            return token
    return "1/1"


if __name__ == "__main__":
    main()
