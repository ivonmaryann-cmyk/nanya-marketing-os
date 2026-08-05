from __future__ import annotations

from fangzheng_web_app.transcode_agent_engine import extract_size
from fangzheng_web_app.transcode_agent_service import calculate_transcode_agent_quote


def _code22(result: dict) -> str:
    return str(result.get("result") or result.get("candidate_code") or "")[:22]


def test_ny3150hf_without_tft_uses_latest_global_rule_even_with_automotive_descriptor():
    result = calculate_transcode_agent_quote(
        "南亚板材 NY3150HF 汽车专用 (1.0mm)39mil H/H 74*49 TG150 无卤素",
        customer="江苏广谦",
        customer_code="104373",
    )

    code = _code22(result)
    assert result["status"] == "成功"
    assert code[:2] == "AH"
    assert code[17] == "R"
    assert code[19:21] == "AC"


def test_explicit_no_copper_in_spec_wins_over_customer_default_total_semantic_rule():
    result = calculate_transcode_agent_quote(
        "FR4 0.1mm H/H 不含铜 43*49 NY2170",
        customer="深三德盈",
        customer_code="103996",
    )

    code = _code22(result)
    assert result["status"] == "待确认"
    assert code[21] == "C"
    assert "TSR-00031-01" not in str(result.get("note") or "")


def test_decimal_thickness_followed_by_inch_token_is_not_accepted_as_panel_size():
    spec = "覆铜板 FR4 NY3170M 0.102 H/H RTF/RTF 0.13786IN 49IN 1067*2 黄色 无水印 大料 无卤"
    assert extract_size(spec) is None

    result = calculate_transcode_agent_quote(
        spec,
        customer="崇达一厂",
        customer_code="103769",
    )
    assert result["status"] != "成功"
    assert not result.get("result")
