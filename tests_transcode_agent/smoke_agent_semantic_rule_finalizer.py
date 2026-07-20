from __future__ import annotations

import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_rule_finalizer import (
    build_atomic_semantic_rule_rows,
    validate_atomic_conditions,
    validate_atomic_semantic_rule_rows,
)


AUDIT_JSON = ROOT / "docs/develop0707/客户特殊规则模型JSON化结果_20260710.json"


def main() -> None:
    payload = json.loads(AUDIT_JSON.read_text(encoding="utf-8"))
    rows, pending = build_atomic_semantic_rule_rows(
        payload,
        approval_basis="P2-2B smoke",
        note="P2-2B smoke",
    )
    validate_atomic_semantic_rule_rows(rows)
    assert len(rows) == 51
    assert len({row["来源候选ID"] for row in rows}) == 39
    assert len(pending) == 3
    assert {item["candidate_id"] for item in pending} == {
        "MSR-0003-01",
        "MSR-0185-02",
        "MSR-0212-02",
    }

    material_rules = _candidate_rows(rows, "MSR-0023-02")
    assert len(material_rules) == 2
    assert {row["标准语义值"] for row in material_rules} == {"core", "total"}
    assert {_first_condition_value(row) for row in material_rules} == {"631", "632"}

    glue_rules = _candidate_rows(rows, "MSR-0144-01")
    assert len(glue_rules) == 4
    assert {row["标准语义值"] for row in glue_rules} == {"NY2140", "NY2150", "NY2170"}

    for candidate_id in ["MSR-0181-01", "MSR-0185-01", "MSR-0212-01"]:
        total_core_rules = _candidate_rows(rows, candidate_id)
        assert len(total_core_rules) == 2
        assert {row["标准语义值"] for row in total_core_rules} == {"total", "core"}

    assert _single(rows, "MSR-0007-03")["标准语义值"] == "AM"
    assert _single(rows, "MSR-0044-01")["标准语义值"] == "AP"
    assert _single(rows, "MSR-0100-01")["标准语义值"] == "AC"
    assert _single(rows, "MSR-0189-01")["标准语义值"] == "RTF"
    assert _single(rows, "MSR-0023-01")["标准语义值"] == "external_lookup:新美亚规格尺寸对照表"

    anbi_rules = _candidate_rows(rows, "MSR-0220-01")
    assert len(anbi_rules) == 3
    assert {row["标准语义值"] for row in anbi_rules} == {"total", "core"}
    total_rule = next(row for row in anbi_rules if row["标准语义值"] == "total")
    total_conditions = json.loads(total_rule["条件JSON"])
    assert any(item["operator"] == "gte" and item["value"] == 0.8 for item in total_conditions)
    assert any(item["operator"] == "not_contains" and item["value"] == "不含铜" for item in total_conditions)

    _assert_invalid_conditions()
    print("semantic rule finalizer smoke passed atomic_rules=51 candidates=39 pending=3")


def _assert_invalid_conditions() -> None:
    invalid_sets = [
        [
            {"field": "订单备注", "operator": "contains_any", "value": "不含铜"},
            {"field": "订单备注", "operator": "not_contains", "value": "不含铜"},
        ],
        [
            {"field": "客户物料编码", "operator": "equals", "value": "631"},
            {"field": "客户物料编码", "operator": "equals", "value": "632"},
        ],
        [
            {"field": "订单备注", "operator": "missing", "value": None},
            {"field": "订单备注", "operator": "contains_any", "value": "不含铜"},
        ],
    ]
    for conditions in invalid_sets:
        try:
            validate_atomic_conditions(conditions, context="smoke")
        except ValueError:
            continue
        raise AssertionError(conditions)


def _candidate_rows(rows: list[dict], candidate_id: str) -> list[dict]:
    return [row for row in rows if row["来源候选ID"] == candidate_id]


def _single(rows: list[dict], candidate_id: str) -> dict:
    matched = _candidate_rows(rows, candidate_id)
    assert len(matched) == 1, (candidate_id, matched)
    return matched[0]


def _first_condition_value(row: dict) -> str:
    return str(json.loads(row["条件JSON"])[0]["value"])


if __name__ == "__main__":
    main()
