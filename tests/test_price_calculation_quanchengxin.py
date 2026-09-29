from __future__ import annotations

from pathlib import Path
import unittest

from fangzheng_web_app.excel_utils import load_workbook_compat
from fangzheng_web_app.price_calculation_customer_mapping import (
    PRICE_TAX_MODE_INCLUSIVE,
    resolve_customer_price_calculation,
)
from fangzheng_web_app.price_calculation_extended import calculate_extended_spec, load_extended_rules


RULE_PATH = Path("fangzheng_web_app/default_rules/price_calculation/quanchengxin/price_rules.xlsx")
TEST_DATA_PATH = Path("fangzheng_web_app/default_rules/price_calculation/quanchengxin/test_data.xlsx")


class QuanchengxinPriceCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rules = load_extended_rules("quanchengxin", RULE_PATH)

    def test_loads_all_quote_sheets_and_dynamic_conditions(self) -> None:
        self.assertEqual(
            {"NY2150", "NY2170", "NY2170H", "NY3150HF", "NY3170M", "NY3170HF", "NY3170M2", "NY6666SE"},
            set(self.rules.ccl_notes),
        )
        self.assertGreaterEqual(len(self.rules.ccl_rows), 340)
        self.assertGreaterEqual(len(self.rules.pp_rows), 120)
        self.assertEqual(0.05, self.rules.ccl_notes["NY2150"]["rtf_markup"])
        self.assertIsNone(self.rules.ccl_notes["NY3170M"]["rtf_markup"])
        self.assertEqual("RTF", self.rules.ccl_notes["NY3170M"]["standard_foil"])

    def test_calculates_standard_ccl_from_the_quote_row(self) -> None:
        result = calculate_extended_spec(
            "quanchengxin",
            "NY2150 0.076mm 1/1oz HTE 82*49 (1080*1)",
            self.rules,
        )

        self.assertEqual("成功", result.status)
        self.assertEqual(470.0, result.price)
        self.assertEqual("82", result.size_column)
        self.assertIn("含13%增值税", result.note)

    def test_uses_rtf_quote_rows_without_an_extra_markup(self) -> None:
        result = calculate_extended_spec(
            "quanchengxin",
            "NY3170M 0.076mm H/H RTF 82*49 (1080*1)",
            self.rules,
        )

        self.assertEqual("成功", result.status)
        self.assertEqual(464.0, result.price)
        self.assertIn("报价专用行", result.note)

    def test_uses_sheet_rtf_markup_only_when_a_direct_row_is_missing(self) -> None:
        result = calculate_extended_spec(
            "quanchengxin",
            "NY2150 0.076mm H/H RTF 82*49 (1080*1)",
            self.rules,
        )

        self.assertEqual("成功", result.status)
        self.assertEqual(431.0, result.price)
        self.assertIn("上调5%", result.note)

    def test_supports_one_h_as_one_one(self) -> None:
        result = calculate_extended_spec(
            "quanchengxin",
            "NY2150 0.10mm 1/Hoz 43*49 (2116*1)",
            self.rules,
        )

        self.assertEqual("成功", result.status)
        self.assertEqual(245.0, result.price)
        self.assertIn("1/Hoz", result.note)

    def test_applies_confirmed_size_mappings(self) -> None:
        oversized = calculate_extended_spec("quanchengxin", "NY2150 0.10mm 1/1oz 82.3*49.3 (2116*1)", self.rules)
        small = calculate_extended_spec("quanchengxin", "NY2150 0.10mm 1/1oz 43*49 (2116*1)", self.rules)
        narrow = calculate_extended_spec("quanchengxin", "NY2150 0.10mm 1/1oz 37*43 (2116*1)", self.rules)

        self.assertEqual(("成功", 466.0, "82"), (oversized.status, oversized.price, oversized.size_column))
        self.assertEqual(("成功", 245.0, "42"), (small.status, small.price, small.size_column))
        self.assertEqual(("成功", 207.0, "SF"), (narrow.status, narrow.price, narrow.size_column))

    def test_calculates_common_pp_and_rejects_special_or_piece_specs(self) -> None:
        roll = calculate_extended_spec("quanchengxin", "NY2150P 7628 RC49% 150米", self.rules)
        special = calculate_extended_spec("quanchengxin", "NY2150P 7628 RC49% 200米", self.rules)
        piece = calculate_extended_spec("quanchengxin", "PP NY2150P 7628 RC49% 20*20IN", self.rules)

        self.assertEqual(("成功", 46.0, "RMB/M"), (roll.status, roll.price, roll.size_column))
        self.assertEqual(("失败", "未匹配"), (special.status, special.price))
        self.assertEqual(("失败", "未匹配"), (piece.status, piece.price))

    def test_rejects_missing_stack_when_the_quote_has_multiple_candidates(self) -> None:
        result = calculate_extended_spec("quanchengxin", "NY3170M2 0.20mm 2/2 RTF 82*49", self.rules)

        self.assertEqual(("失败", "未匹配"), (result.status, result.price))
        self.assertIn("多个精确候选行", result.note)

    def test_packaged_test_specs_are_recognized_or_explicitly_rejected(self) -> None:
        ws = load_workbook_compat(TEST_DATA_PATH, data_only=True).active
        results = [
            calculate_extended_spec("quanchengxin", str(ws.cell(row_idx, 1).value), self.rules)
            for row_idx in range(2, ws.max_row + 1)
            if ws.cell(row_idx, 1).value
        ]

        self.assertEqual(37, len(results))
        self.assertEqual(34, sum(result.status == "成功" for result in results))
        self.assertEqual(3, sum(result.status == "失败" for result in results))

    def test_resolves_shenzhen_and_hubei_quanchengxin_customers(self) -> None:
        shenzhen = resolve_customer_price_calculation({"customer_short_name": "", "group_name": "深圳全成信"})
        hubei = resolve_customer_price_calculation({"customer_short_name": "湖全成信", "group_name": ""})

        self.assertEqual(("quanchengxin", PRICE_TAX_MODE_INCLUSIVE), (shenzhen["price_customer_key"], shenzhen["price_tax_mode"]))
        self.assertEqual(("quanchengxin", "客户简称"), (hubei["price_customer_key"], hubei["match_source"]))


if __name__ == "__main__":
    unittest.main()
