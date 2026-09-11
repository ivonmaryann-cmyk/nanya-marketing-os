from __future__ import annotations

import unittest

from fangzheng_web_app.price_calculation_service import (
    JingwangRules,
    PpRule,
    _norm_jingwang_pp_product,
    _parse_jingwang_pp_rule_row,
    calculate_jingwang_spec,
)


class JingwangPriceCalculationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rules = JingwangRules(
            pp_rows=[PpRule(469, "NY-A3HF", "1080", 67.0, 68.0, 36.37)],
            ccl_rows=[],
        )

    def test_pp_quote_row_accepts_a3hf_without_pp_suffix(self) -> None:
        row = _parse_jingwang_pp_rule_row(
            469,
            {2: "31.63", 3: "36.37", 4: "NY-A3HF", 5: "RC67-68%", 6: "1080"},
        )

        self.assertIsNotNone(row)
        self.assertEqual("NY-A3HF", row.product)
        self.assertEqual(36.37, row.price)

    def test_pp_product_matching_ignores_trailing_p(self) -> None:
        self.assertEqual("NY-A3HF", _norm_jingwang_pp_product("NY-A3HFP"))
        self.assertEqual("NY6300S", _norm_jingwang_pp_product("NY6300SP"))

    def test_a3hfp_narrow_roll_returns_quoted_per_meter_price(self) -> None:
        result = calculate_jingwang_spec(
            "PP NY-A3HFP 1080 RC68% 24.41IN 300m 无卤",
            self.rules,
        )

        self.assertEqual("成功", result.status)
        self.assertEqual("PP", result.material_type)
        self.assertEqual(36.37, result.price)
        self.assertEqual("24.41IN", result.width)
        self.assertEqual("300m", result.roll_length)
        self.assertEqual(469, result.rule_row)
