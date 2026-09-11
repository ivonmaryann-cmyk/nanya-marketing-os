from __future__ import annotations

import unittest

from fangzheng_web_app.price_calculation_extended import (
    ExtCclRule,
    ExtPpRule,
    ExtRules,
    _guanghe_ccl_cut_price,
    calculate_extended_spec,
)


class GuiguPriceCalculationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rules = ExtRules(
            "guigu",
            [ExtPpRule(15, "NY6300PP", "NY6300", "1035", 72, 72, 300, None, 82.58)],
            [],
        )

    def test_pp_with_only_latitudinal_width_returns_quote_per_meter(self) -> None:
        result = calculate_extended_spec(
            "guigu", "NY6300P 1035 RC72% 49.5(纬)inch 硅谷专用", self.rules
        )

        self.assertEqual(("成功", "PP", 82.58, "300m"), (result.status, result.material_type, result.price, result.roll_length))
        self.assertIn("每米价", result.note)

    def test_pp_with_two_explicit_piece_dimensions_is_converted(self) -> None:
        result = calculate_extended_spec(
            "guigu", "NY6300P 1035 RC72% 18.3inch*21.3(纬)inch", self.rules
        )

        self.assertEqual(("成功", "PP", 17.04), (result.status, result.material_type, result.price))
        self.assertIn("小片公式", result.note)

    def test_ccl_17_by_49_uses_fifth_of_86_by_49_parent(self) -> None:
        row = ExtCclRule(10, "NY2150", "NY2150", 0.1, None, "1/1", "HTE", "1*1080", {"43": 200, "SF": 20})

        result = _guanghe_ccl_cut_price(row, 17, 49)

        self.assertEqual({"ok": True, "price": 80.0, "label": "86*49/5", "formula": "200*2/5"}, result)
