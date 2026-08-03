from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_rules import parse_customer_special_master
from fangzheng_web_app.transcode_agent_service import _apply_agent_rules


DRAFT_PATH = ROOT / "docs/develop0707/客户特殊规则结构化草稿_按原表行_20260708.xlsx"


def main() -> None:
    rules, summary = parse_customer_special_master(DRAFT_PATH)
    assert summary["source_type"] == "confirmed_structured_draft"
    assert len(rules) == 196, len(rules)

    _assert_applies(
        rules,
        customer_code="103890",
        customer_name="珠海景旺",
        spec="NY3150HC 0.8mm 1/1 37*49 HVLP",
        context="",
        steps=_steps(glue_model="NY3150HC", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "RV", "copper_type_code": "O"},
    )
    _assert_applies(
        rules,
        customer_code="103891",
        customer_name="方正F7",
        spec="NY2150 0.6mm 1/1 37*49 HS2-M2-VSP",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.6, thickness_raw="0.6mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "2B", "tc_code": "C", "copper_type_code": "P"},
    )
    _assert_applies(
        rules,
        customer_code="101003",
        customer_name="上海山崎",
        spec="NY1140 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY1140", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "2A"},
    )
    _assert_applies(
        rules,
        customer_code="104373",
        customer_name="江苏广谦",
        spec="NY2150 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"tc_code": "T"},
    )
    _assert_applies(
        rules,
        customer_code="104373",
        customer_name="江苏广谦",
        spec="NY2150 0.6mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.6, thickness_raw="0.6mm", copper="1/1", size=("37", "49")),
        expected={"tc_code": "C"},
    )
    _assert_applies(
        rules,
        customer_code="133025",
        customer_name="广德三生",
        spec="NY2150 0.6mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.6, thickness_raw="0.6mm", copper="1/1", size=("37", "49")),
        expected={"tc_code": "C"},
    )
    _assert_applies(
        rules,
        customer_code="106008",
        customer_name="川英创力",
        spec="TG140 0.8mm W/W 37*49",
        context="",
        steps=_steps(glue_model="", thickness_mm=0.8, thickness_raw="0.8mm", copper="W/W", size=("37", "49")),
        expected={"glue_code": "2A", "copper_code": "FF"},
    )
    _assert_applies(
        rules,
        customer_code="106008",
        customer_name="川英创力",
        spec="TG150 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "2B"},
    )
    _assert_applies(
        rules,
        customer_code="106008",
        customer_name="川英创力",
        spec="TG170 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "2C"},
    )
    _assert_applies(
        rules,
        customer_code="106011",
        customer_name="川华兴宇",
        spec="NY2150 0.9mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.9, thickness_raw="0.9mm", copper="1/1", size=("37", "49")),
        expected={"grade_code": "AC"},
    )
    _assert_applies(
        rules,
        customer_code="106011",
        customer_name="川华兴宇",
        spec="NY2150 0.91mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.91, thickness_raw="0.91mm", copper="1/1", size=("37", "49")),
        expected={"grade_code": "A1"},
    )
    _assert_applies(
        rules,
        customer_code="106011",
        customer_name="川华兴宇",
        spec="NY2170 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2170", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"grade_code": "AC"},
    )
    _assert_applies(
        rules,
        customer_code="122021",
        customer_name="黄石广合",
        spec="NY2150 43mil 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=1.0922, thickness_raw="43mil", copper="1/1", size=("37", "49")),
        expected={"grade_code": "AT"},
    )
    _assert_applies(
        rules,
        customer_code="103786",
        customer_name="惠州泰和",
        spec="NY3150HF 1.2mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY3150HF", thickness_mm=1.2, thickness_raw="1.2mm", copper="1/1", size=("37", "49")),
        expected={"grade_code": "AT"},
    )
    _assert_applies(
        rules,
        customer_code="123114",
        customer_name="赣州景旺",
        spec="NY3150HC 0.8mm 1/1 37*49 HVLP",
        context="",
        steps=_steps(glue_model="NY3150HC", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "RV", "copper_type_code": "O", "grade_code": "A1"},
    )
    _assert_applies(
        rules,
        customer_code="103990",
        customer_name="广华升鑫、深华升鑫",
        spec="NY-A2 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="NY-A2", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "AL", "grade_code": "AC"},
    )
    _assert_applies(
        rules,
        customer_code="103844",
        customer_name="万安裕维",
        spec="TG150 无卤素 CTI600 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "3H"},
    )
    _assert_applies(
        rules,
        customer_code="103844",
        customer_name="万安裕维",
        spec="TG150 无卤素 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "3B"},
    )
    _assert_applies(
        rules,
        customer_code="103844",
        customer_name="万安裕维",
        spec="TG130 0.8mm 1/1 37*49",
        context="",
        steps=_steps(glue_model="", thickness_mm=0.8, thickness_raw="0.8mm", copper="1/1", size=("37", "49")),
        expected={"glue_code": "2A"},
    )
    _assert_applies(
        rules,
        customer_code="103576",
        customer_name="珠海乐健",
        spec="NY2150 1.1mm H/H 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=1.1, thickness_raw="1.1mm", copper="H/H", size=("37", "49")),
        expected={"grade_code": "A1"},
    )
    _assert_applies(
        rules,
        customer_code="103576",
        customer_name="珠海乐健",
        spec="NY2150 1.5mm H/H 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=1.5, thickness_raw="1.5mm", copper="H/H", size=("37", "49")),
        expected={"grade_code": "A1"},
    )
    _assert_applies(
        rules,
        customer_code="123123",
        customer_name="赣州乐健",
        spec="NY2150 0.8mm H/H 37*49",
        context="",
        steps=_steps(glue_model="NY2150", thickness_mm=0.8, thickness_raw="0.8mm", copper="H/H", size=("37", "49")),
        expected={"grade_code": "AC"},
    )

    print("agent rule application smoke passed")


def _steps(*, glue_model: str, thickness_mm: float, thickness_raw: str, copper: str, size: tuple[str, str]) -> dict:
    return {
        "glue_model": glue_model,
        "thickness_mm": thickness_mm,
        "thickness_raw": thickness_raw,
        "copper_spec_raw": copper,
        "size_w": size[0],
        "size_h": size[1],
        "step1_glue_code": "??",
        "step2_thick_code": "?????",
        "step3_copper_code": "??",
        "step4_size_code": "????????",
        "step5_glue_cat_code": "Y",
        "step6_copper_type_code": "W",
        "step7_grade_code": "A1",
        "step8_tc_code": "T",
        "step9_struct_code": "*",
    }


def _assert_applies(
    rules: list[dict],
    *,
    customer_code: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
    expected: dict[str, str],
) -> None:
    errors = ["无法识别胶系型号", "无法识别厚度", "无法识别铜箔规格", "无法识别尺寸"]
    applied, conflicts = _apply_agent_rules(rules, customer_code, customer_name, spec, context, steps, errors)
    assert not conflicts, conflicts
    assert applied, f"no rule applied for {customer_name}: {spec}"
    step_map = {
        "glue_code": "step1_glue_code",
        "thickness_code": "step2_thick_code",
        "copper_code": "step3_copper_code",
        "size_code": "step4_size_code",
        "glue_category_code": "step5_glue_cat_code",
        "copper_type_code": "step6_copper_type_code",
        "grade_code": "step7_grade_code",
        "tc_code": "step8_tc_code",
    }
    for field, expected_value in expected.items():
        assert steps[step_map[field]] == expected_value, (customer_name, field, expected_value, steps[step_map[field]], applied)


if __name__ == "__main__":
    main()
