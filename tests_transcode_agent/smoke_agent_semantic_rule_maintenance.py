from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_rule_maintenance import (
    build_semantic_candidates,
    classify_draft_workbook,
    requires_semantic_model,
    split_ccl_rule_fields,
)


DRAFT_PATH = ROOT / "docs/develop0707/客户特殊规则结构化草稿_按原表行_20260708.xlsx"
SKILL_DIR = ROOT / "model_skills/customer-special-rule-maintenance"


def main() -> None:
    test_rule_field_split()
    test_model_routing_examples()
    test_full_draft_classification()
    test_multisource_candidates()
    test_skill_cli()
    print("semantic rule maintenance smoke passed")


def test_rule_field_split() -> None:
    fields = split_ccl_rule_fields(
        "胶系：NY2150=2B；基板厚度：0.8以下芯厚；"
        "基板级别：当备注中有汽车板字样时=AC；组合结构：Y/Z"
    )
    assert fields["胶系"] == "NY2150=2B"
    assert "汽车板" in fields["基板级别"]
    assert fields["组合结构"] == "Y/Z"


def test_model_routing_examples() -> None:
    assert requires_semantic_model("当备注中有汽车板字样时=AC")
    assert requires_semantic_model("NY1600及NY2140外，其余均下汽车板料号")
    assert requires_semantic_model("客户规格没有Q时=AP")
    assert requires_semantic_model("订单备注第5码A或S时=AC")
    assert not requires_semantic_model("0.8含以上=总厚")
    assert not requires_semantic_model("参考健鼎超颖板厚表进行换算")

    candidates = build_semantic_candidates(
        customer_code="105007",
        customer_name="湖奥士康",
        source_row=7,
        ccl_rule=(
            "胶系：NY-A1=RC；铜箔规格：R/R=F/F；"
            "基板级别：当备注中有X0A0/X0B0/car/汽车板字样时=AC，"
            "当备注中有MINILED字样时=AM；"
            "组合结构：当胶系=NY2140L时=A"
        ),
    )
    assert len(candidates) == 2
    assert all(item.business_field == "基板级别" for item in candidates)
    assert all(item.required_input_fields == "订单备注" for item in candidates)


def test_full_draft_classification() -> None:
    result = classify_draft_workbook(DRAFT_PATH)
    assert result["row_count"] == 219
    assert result["candidate_count"] >= 40
    assert result["candidate_customer_count"] >= 30
    assert sum(result["path_counts"].values()) == 219
    assert any(item["customer_name"] == "湖奥士康" for item in result["candidates"])
    assert any(
        item["customer_name"] == "常熟敬鹏" and item["semantic_type"] == "排除集合"
        for item in result["candidates"]
    )
    assert not any(item["business_field"] == "组合结构" for item in result["candidates"])


def test_multisource_candidates() -> None:
    result = classify_draft_workbook(DRAFT_PATH)
    source_columns = {item["source_column"] for item in result["candidates"]}
    assert source_columns == {"CCL特殊规则", "通用特殊规则", "非影响转码备注"}
    assert any(
        item["customer_name"] == "广东依顿"
        and item["source_column"] == "非影响转码备注"
        and item["business_field"] == "基板级别"
        for item in result["candidates"]
    )
    assert any(
        item["customer_name"] == "吉安生益"
        and item["source_column"] == "通用特殊规则"
        and "HTE" in item["source_text"]
        for item in result["candidates"]
    )


def test_skill_cli() -> None:
    assert (SKILL_DIR / "SKILL.md").exists()
    assert (SKILL_DIR / "references/规则分流标准.md").exists()
    assert (SKILL_DIR / "references/维护SOP.md").exists()
    script = SKILL_DIR / "scripts/classify_customer_rules.py"
    with tempfile.TemporaryDirectory() as temp_dir:
        output = Path(temp_dir) / "classification.json"
        subprocess.run(
            [sys.executable, str(script), str(DRAFT_PATH), "--output", str(output)],
            cwd=ROOT,
            check=True,
            capture_output=True,
            text=True,
        )
        data = json.loads(output.read_text(encoding="utf-8"))
        assert data["row_count"] == 219


if __name__ == "__main__":
    main()
