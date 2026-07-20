from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_rules import parse_semantic_rule_workbook
from fangzheng_web_app.transcode_semantic_shadow import (
    SHADOW_STATUS_MATCHED,
    SHADOW_STATUS_MISSING_INPUT,
    SHADOW_STATUS_NOT_MATCHED,
    evaluate_semantic_shadow,
    get_semantic_rule_runtime_mode,
)


RULE_WORKBOOK = ROOT / "docs/develop0707/营销转码Agent最终模型语义规则表_20260710.xlsx"


def main() -> None:
    rules, summary = parse_semantic_rule_workbook(RULE_WORKBOOK)
    assert len(rules) == 51, summary
    assert get_semantic_rule_runtime_mode({}) == "shadow"
    assert get_semantic_rule_runtime_mode({"TRANSCODE_SEMANTIC_RULE_RUNTIME_MODE": "off"}) == "off"

    matched = evaluate_semantic_shadow(
        rules,
        customer_code="103890",
        customer_name="珠海景旺",
        observations=_observations("能源板订单"),
        excel_row=2,
        spec="NY3150HC 0.8mm 1/1 37*49 HVLP",
    )
    assert _status(matched, "TSR-00001") == SHADOW_STATUS_MATCHED, matched

    missing = evaluate_semantic_shadow(
        rules,
        customer_code="103890",
        customer_name="珠海景旺",
        observations=_observations(None),
        excel_row=2,
        spec="NY3150HC 0.8mm 1/1 37*49 HVLP",
    )
    assert _status(missing, "TSR-00001") == SHADOW_STATUS_MISSING_INPUT, missing
    assert "订单备注" in missing[0]["missing_fields"], missing

    automotive = evaluate_semantic_shadow(
        rules,
        customer_code="104354",
        customer_name="常熟敬鹏",
        observations=_observations("", glue="NY2150"),
        excel_row=3,
        spec="NY2150 0.8mm 1/1 37*49",
    )
    assert _status(automotive, "TSR-00007") == SHADOW_STATUS_MATCHED, automotive

    excluded = evaluate_semantic_shadow(
        rules,
        customer_code="104354",
        customer_name="常熟敬鹏",
        observations=_observations("", glue="NY2140"),
        excel_row=4,
        spec="NY2140 0.8mm 1/1 37*49",
    )
    assert _status(excluded, "TSR-00007") == SHADOW_STATUS_NOT_MATCHED, excluded

    material_code = evaluate_semantic_shadow(
        rules,
        customer_code="104158",
        customer_name="江苏瀚宇",
        observations=_observations("", customer_material_code="ABC123Q"),
        excel_row=5,
        spec="NY2150 0.8mm 1/1 37*49",
    )
    assert _status(material_code, "TSR-00026") == SHADOW_STATUS_MATCHED, material_code
    assert all("formal_code" not in item and "score" not in item for item in matched + automotive)

    material_631 = evaluate_semantic_shadow(
        rules,
        customer_code="104443",
        customer_name="无新美亚",
        observations=_observations("", customer_material_code="631"),
        excel_row=6,
        spec="NY2150 0.8mm 1/1 37*49",
    )
    material_statuses = _candidate_statuses(material_631, "MSR-0023-02")
    assert material_statuses.count(SHADOW_STATUS_MATCHED) == 1, material_631
    assert material_statuses.count(SHADOW_STATUS_NOT_MATCHED) == 1, material_631

    default_total = evaluate_semantic_shadow(
        rules,
        customer_code="103996",
        customer_name="深三德盈",
        observations=_observations("普通订单"),
        excel_row=7,
        spec="NY2150 0.8mm 1/1 37*49",
    )
    assert _matched_values(default_total, "MSR-0181-01") == ["total"], default_total
    no_copper = evaluate_semantic_shadow(
        rules,
        customer_code="103996",
        customer_name="深三德盈",
        observations=_observations("订单不含铜"),
        excel_row=8,
        spec="NY2150 0.8mm 1/1 37*49",
    )
    assert _matched_values(no_copper, "MSR-0181-01") == ["core"], no_copper

    anbi_core = evaluate_semantic_shadow(
        rules,
        customer_code="203012",
        customer_name="深圳安比",
        observations=_observations("含铜", thickness=0.7),
        excel_row=9,
        spec="NY2150 0.7mm 1/1 37*49",
    )
    assert _matched_values(anbi_core, "MSR-0220-01") == ["core"], anbi_core
    anbi_total = evaluate_semantic_shadow(
        rules,
        customer_code="203012",
        customer_name="深圳安比",
        observations=_observations("含铜", thickness=0.8),
        excel_row=10,
        spec="NY2150 0.8mm 1/1 37*49",
    )
    assert _matched_values(anbi_total, "MSR-0220-01") == ["total"], anbi_total

    print("semantic shadow smoke passed rules=51 matched/missing/not-matched/char-at verified")


def _observations(
    order_remark: str | None,
    *,
    glue: str = "NY3150HC",
    customer_material_code: str = "",
    thickness: float = 0.8,
) -> dict:
    remark_available = order_remark is not None
    return {
        "订单备注": {
            "available": remark_available,
            "value": order_remark or "",
            "sources": ["订单备注"] if remark_available else [],
        },
        "订单规格": {"available": True, "value": "NY3150HC 0.8mm", "sources": ["规格"]},
        "订单规格/订单备注": {
            "available": True,
            "value": f"NY3150HC 0.8mm {order_remark or ''}",
            "sources": ["规格"] + (["订单备注"] if remark_available else []),
        },
        "胶系": {"available": True, "value": glue, "sources": ["规格解析/胶系"]},
        "基板厚度": {"available": True, "value": thickness, "sources": ["规格解析/基板厚度"]},
        "客户物料编码": {
            "available": True,
            "value": customer_material_code,
            "sources": ["客户产品编号"],
        },
        "客户料品名称": {"available": True, "value": "", "sources": ["品名"]},
        "客户规格": {"available": True, "value": "", "sources": ["客户规格"]},
    }


def _status(evaluations: list[dict], rule_id: str) -> str:
    for item in evaluations:
        if item.get("rule_id") == rule_id:
            return str(item.get("status") or "")
    raise AssertionError((rule_id, evaluations))


def _candidate_statuses(evaluations: list[dict], candidate_id: str) -> list[str]:
    return [
        str(item.get("status") or "")
        for item in evaluations
        if item.get("source_candidate_id") == candidate_id
    ]


def _matched_values(evaluations: list[dict], candidate_id: str) -> list[str]:
    return sorted(
        str(value)
        for item in evaluations
        if item.get("source_candidate_id") == candidate_id and item.get("status") == SHADOW_STATUS_MATCHED
        for value in item.get("normalized_values", [])
    )


if __name__ == "__main__":
    main()
