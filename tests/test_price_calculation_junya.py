from __future__ import annotations

from pathlib import Path
import unittest

from openpyxl import load_workbook

from fangzheng_web_app.price_calculation_extended import calculate_extended_spec, load_extended_rules
from fangzheng_web_app.price_calculation_service import process_price_workbook


RULE_PATH = Path("fangzheng_web_app/default_rules/price_calculation/junya/price_rules.xlsx")
TEST_DATA_PATH = Path("fangzheng_web_app/default_rules/price_calculation/junya/test_data.xlsx")


class JunyaPriceCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rules = load_extended_rules("junya", RULE_PATH)

    def test_pp_returns_quote_per_meter_price_for_resin_content_range(self) -> None:
        result = calculate_extended_spec("junya", "NY2150P 2116 300M/卷 含量55%", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual("PP", result.material_type)
        self.assertEqual(35.5, result.price)
        self.assertEqual("300m", result.roll_length)
        self.assertIn("每米价格", result.note)

    def test_pp_matches_the_next_resin_content_range(self) -> None:
        result = calculate_extended_spec("junya", "NY2150P 2116 300M/卷 含量57%", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual(36.5, result.price)

    def test_pp_ignores_length_when_matching(self) -> None:
        result = calculate_extended_spec("junya", "NY2150P 2116 含量55%", self.rules)

        self.assertEqual(("成功", 35.5), (result.status, result.price))

    def test_ccl_matches_both_values_of_slash_thickness(self) -> None:
        first = calculate_extended_spec("junya", "NY2150 0.127mm(不含铜)1/1 FR-4 82*49 1*2116", self.rules)
        second = calculate_extended_spec("junya", "NY2150 0.13mm(不含铜)1/1 FR-4 82*49 1*2116", self.rules)

        self.assertEqual(("成功", 396.0), (first.status, first.price))
        self.assertEqual(("成功", 396.0), (second.status, second.price))

    def test_ccl_applies_37_and_43_size_factors(self) -> None:
        base = "NY2150 0.13mm(不含铜)1/1 FR-4 {size} 1*2116"
        result_37 = calculate_extended_spec("junya", base.format(size="37*49"), self.rules)
        result_43 = calculate_extended_spec("junya", base.format(size="43*49"), self.rules)

        self.assertEqual(("成功", 178.2), (result_37.status, result_37.price))
        self.assertEqual(("成功", 207.9), (result_43.status, result_43.price))

    def test_ny3150_hc_uses_ny3150_hf_quote_sheet(self) -> None:
        result = calculate_extended_spec(
            "junya",
            "NY3150HC 0.45mm(不含铜)1/1 FR-4 41*49 1*1080+2*7628",
            self.rules,
        )

        self.assertEqual(("成功", 258.0), (result.status, result.price))
        self.assertIn("NY3150HC按NY3150HF报价行取价", result.note)

    def test_ny3150_hc_pp_uses_ny3150_hf_hc_quote_sheet(self) -> None:
        result = calculate_extended_spec(
            "junya", "半固化片PP 2116（上海南亚NY3150HC）49.5*300M 55%", self.rules
        )

        self.assertEqual(("成功", 38.2, 57), (result.status, result.price, result.rule_row))

    def test_ccl_examples_use_the_requested_quote_rows(self) -> None:
        ny2140 = calculate_extended_spec(
            "junya", "NY2140 1.1(含铜)1/1 FR-4 86*49 6*7628上海无水印普通TG", self.rules
        )
        ny2150h = calculate_extended_spec(
            "junya", "NY2150H 1.5(含铜)1/1 FR-4 82.3*49.3 8*7628上海南亚无水印TG150", self.rules
        )
        ny2150h_without_copper = calculate_extended_spec(
            "junya", "NY2150H 1.5(不含铜)1/1 FR-4 82.3*49.3 8*7628上海南亚无水印TG150", self.rules
        )

        self.assertEqual(("成功", 680.4, 34), (ny2140.status, ny2140.price, ny2140.rule_row))
        self.assertEqual(("成功", 800.0, 41), (ny2150h.status, ny2150h.price, ny2150h.rule_row))
        self.assertEqual(("成功", 800.0, 41), (ny2150h_without_copper.status, ny2150h_without_copper.price, ny2150h_without_copper.rule_row))

    def test_ccl_j_j_uses_the_dedicated_junya_quote_column(self) -> None:
        result = calculate_extended_spec(
            "junya",
            'NY1600 1.5(含铜）J/J FR-4 74.3"*49.3" 8*7628 上海南亚无水印普通TG CTI≥600',
            self.rules,
        )

        self.assertEqual(("成功", 568.8, 37), (result.status, result.price, result.rule_row))

    def test_missing_quote_returns_unmatched(self) -> None:
        result = calculate_extended_spec("junya", "NY2150P 2116 300M/卷 含量61%", self.rules)

        self.assertEqual("失败", result.status)
        self.assertEqual("未匹配", result.price)
        self.assertIn("未匹配骏亚PP报价", result.note)

    def test_pp_small_piece_returns_unmatched(self) -> None:
        result = calculate_extended_spec("junya", "NY2150P 2116 含量55% 300M 400mm*500mm", self.rules)

        self.assertEqual("失败", result.status)
        self.assertEqual("未匹配", result.price)
        self.assertIn("PP小片不参与报价", result.note)

    def test_batch_uses_customer_spec_and_layout_stack(self) -> None:
        workbook = load_workbook(TEST_DATA_PATH)
        results = process_price_workbook(workbook, "junya", self.rules)
        row_five = next(item for item in results if item["row"] == 5)

        self.assertEqual("成功", row_five["status"])
        self.assertEqual(35.5, next(item for item in results if item["row"] == 4)["price"])
        self.assertEqual(516.0, row_five["price"])
        row_thirteen = next(item for item in results if item["row"] == 13)
        self.assertEqual(("成功", 320.0, 17), (row_thirteen["status"], row_thirteen["price"], row_thirteen["rule_row"]))

    def test_ccl_prefers_dimension_over_stack_pair(self) -> None:
        result = calculate_extended_spec(
            "junya", "NY3170M 0.305(不含铜)2/2(1506*2) FR-4 37*49 2*1506", self.rules
        )

        self.assertNotIn("尺寸未匹配", result.note)
