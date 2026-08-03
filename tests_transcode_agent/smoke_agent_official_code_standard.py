from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_rules import GRADE_CODES
from fangzheng_web_app.transcode_agent_service import _load_runtime, analyze_spec
from fangzheng_web_app.transcode_agent_standard import load_base_code_standard


def main() -> None:
    engine, tables, rules, mappings, _, _ = _load_runtime()
    standard = load_base_code_standard()
    assert standard["high_speed_mil_to_mm"][15.0] == 0.381
    assert standard["standard_mm_size_aliases"] == {940.0: 37.0, 1245.0: 49.0}

    # NYG-ATD-002-A1中的标准尺寸是编码值，生产能力区间不能当作放大映射。
    assert engine.extract_size("NY3170M 0.152*940*1245MM") == (37.0, 49.0)
    assert engine.extract_size('NY3170M 0.152mm 37.3"*49.3"') == (37.3, 49.3)

    # 规范mil表只作用于高频/高速材料，常规材料继续使用既有下单口径。
    high_speed = analyze_spec(
        engine,
        tables,
        rules,
        "NY6600 15mil H/H 41*49 RTF 不含铜",
        agent_mapping_tables=mappings,
        customer="规范测试客户",
    )
    assert high_speed["engine_steps"]["thickness_mm"] == 0.381, high_speed
    assert high_speed["engine_steps"]["step2_thick_code"] == "00381", high_speed

    regular = analyze_spec(
        engine,
        tables,
        rules,
        "NY3150HF 15mil H/H 41*49 HTE 不含铜",
        agent_mapping_tables=mappings,
        customer="规范测试客户",
    )
    assert regular["engine_steps"]["thickness_mm"] == 0.38, regular
    assert regular["engine_steps"]["step2_thick_code"] == "00380", regular

    m6_below = analyze_spec(
        engine,
        tables,
        rules,
        'NY3170HF 4mil 106*2 不含铜 6/6 43"x49" RTF 无卤',
        agent_mapping_tables=mappings,
        customer="惠州泰和",
    )
    assert m6_below["engine_steps"]["thickness_mm"] == 0.10, m6_below
    assert m6_below["engine_steps"]["step2_thick_code"] == "00100", m6_below

    # 第19码组合直接读取transcode_rules.xlsx/编码规则，不由模型推断。
    copper_cases = {
        "HVLP1+RTF 无水印": "C",
        "RTF2+HTE 无水印": "F",
        "HTE+RTF2 无水印": "U",
        "RTF3+RTF 无水印": "T",
        "RTF 有水印": "Y",
        "HTE 无水印": "W",
    }
    for text, expected in copper_cases.items():
        actual = engine.get_copper_type_code(text, tables["copper_type_rules"])
        assert actual == expected, (text, expected, actual)
    assert engine.extract_copper_spec("22um/22um") == "I/I"

    # 扩充的是合法值白名单，不代表没有条件时自动选择这些等级。
    official_grade_codes = {
        "A1", "A2", "A3", "A4", "A5", "A6", "A8", "A9",
        "AC", "AD", "AH", "AJ", "AL", "AM", "AN", "AP", "AQ", "AT", "AW",
        "AY", "B3", "C1", "D1", "D2", "D3", "D4", "D5", "F1", "NN", "PG", "S1", "T1", "X",
    }
    assert official_grade_codes <= GRADE_CODES

    # 通用总/芯厚边界：0.8mm含以上为总厚，以下为芯厚。
    assert engine.calc_order_thickness(0.8, "H/H", "unknown", {}, {}) == (0.8, True)
    assert engine.calc_order_thickness(0.79, "H/H", "unknown", {}, {}) == (0.79, False)

    print("official code standard smoke: PASS")


if __name__ == "__main__":
    main()
