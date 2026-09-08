from __future__ import annotations

import unittest
from pathlib import Path

from fangzheng_web_app.price_calculation_extended import calculate_extended_spec, load_extended_rules


RULE_PATH = Path("fangzheng_web_app/default_rules/price_calculation/chaoying/price_rules.xlsx")
BASE_SPEC = 'CCL Halogen Free MTG PN NY2150 15mil 1/1 2*7628 Normal {size} FR-4'


class ChaoyingPriceCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rules = load_extended_rules("chaoying", RULE_PATH)

    def test_direct_ccl_price_uses_quoted_small_sheet_column(self) -> None:
        result = calculate_extended_spec("chaoying", BASE_SPEC.format(size='27"*48"'), self.rules)

        self.assertEqual(result.status, "成功")
        self.assertEqual(result.price, 138.87)
        self.assertEqual(result.rule_row, 80)

    def test_double_width_ccl_uses_corresponding_single_sheet_price_times_two(self) -> None:
        result = calculate_extended_spec("chaoying", BASE_SPEC.format(size='86"*49"'), self.rules)

        self.assertEqual(result.status, "成功")
        self.assertEqual(result.price, 432.04)
        self.assertIn("43*49×2", result.note)

    def test_pp_small_piece_is_not_quoted(self) -> None:
        result = calculate_extended_spec(
            "chaoying", "PP Halogen Free MTG PN NY2150P 2116 56% 300mm*400mm", self.rules
        )

        self.assertEqual((result.status, result.price), ("失败", "未匹配"))
        self.assertEqual(result.note, "超颖PP小片不报价")

    def test_pp_uses_the_per_meter_price_below_the_ccl_table(self) -> None:
        result = calculate_extended_spec(
            "chaoying", 'PP Halogen MTG PN NY-A1P 1080 66% 48" (1250 mm*300m)', self.rules
        )

        self.assertEqual((result.status, result.price), ("成功", 27.94))
        self.assertEqual((result.rule_row, result.size_column), (95, "RMB/M"))

    def test_ccl_does_not_ignore_stack(self) -> None:
        result = calculate_extended_spec("chaoying", BASE_SPEC.format(size='27"*48"').replace("2*7628", "1*7628"), self.rules)

        self.assertEqual((result.status, result.price), ("失败", "未匹配"))
        self.assertIn("叠构不匹配", result.note)

    def test_ccl_without_stack_uses_the_only_matching_quote_row(self) -> None:
        result = calculate_extended_spec(
            "chaoying", 'CCL Halogen MTG PN NY2150 47mil 1/1 Normal 28"*48" FR-4', self.rules
        )

        self.assertEqual((result.status, result.price), ("成功", 187.34))
        self.assertEqual(result.rule_row, 127)
        self.assertIn("唯一报价行叠构=6*7628", result.note)

    def test_ccl_loads_product_from_sheet_header_when_title_is_generic(self) -> None:
        result = calculate_extended_spec(
            "chaoying", "NY3170LK 2.3mil H/H (1078*1) RTF 27*48", self.rules
        )

        self.assertEqual((result.status, result.price), ("成功", 91.17))
        self.assertEqual((result.rule_row, result.size_column), (8, "27*48"))
