from __future__ import annotations

import unittest

from fangzheng_web_app.price_calculation_extended import (
    ExtCclRule,
    ExtPpRule,
    ExtRules,
    calculate_extended_spec,
)


class ZhongfuPriceCalculationTests(unittest.TestCase):
    def test_pp_matches_chinese_resin_content_and_roll_format(self) -> None:
        rules = ExtRules(
            "zhongfu",
            [ExtPpRule(12, "NY2170HP", "NY2170HP", "1080", 73, 73, 300, 49.5, 38.7)],
            [],
        )

        result = calculate_extended_spec("zhongfu", "NY2170HP 1080 含量73.00% 卷300Mx49.50", rules)

        self.assertEqual(("成功", "PP", 38.7, 12), (result.status, result.material_type, result.price, result.rule_row))

    def test_total_thickness_selects_the_copper_included_ccl_quote(self) -> None:
        rules = ExtRules(
            "zhongfu",
            [],
            [
                ExtCclRule(223, "NY2150H", "NY2150H", 1.5, None, "H/H", "HTE", "8*7628", {"41": 369.86}, "不含铜"),
                ExtCclRule(226, "NY2150H", "NY2150H", 1.5, None, "H/H", "HTE", "8*7628", {"41": 340.36}, "含铜"),
            ],
        )

        result = calculate_extended_spec(
            "zhongfu", "NY2150H 1.500mm 总厚 H/H HTE+HTE A 无水印 41.00x49.00 AH", rules
        )

        self.assertEqual(("成功", "CCL", 340.36, 226), (result.status, result.material_type, result.price, result.rule_row))
