from __future__ import annotations

from pathlib import Path
import unittest

from fangzheng_web_app.price_calculation_extended import calculate_extended_spec, load_extended_rules


RULE_PATH = Path("fangzheng_web_app/default_rules/price_calculation/yibo/price_rules.xls")


class YiboPriceCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rules = load_extended_rules("yibo", RULE_PATH)

    def test_loads_ccl_and_pp_sections_from_each_quote_sheet(self) -> None:
        self.assertGreater(len(self.rules.ccl_rows), 200)
        self.assertGreater(len(self.rules.pp_rows), 100)

    def test_calculates_pp_roll_from_exact_quote_row(self) -> None:
        result = calculate_extended_spec("yibo", "PP NY2170HP 2116 RC59% 49.5IN*300M", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual("PP", result.material_type)
        self.assertEqual(12952.73, result.price)
        self.assertEqual("Per Roll", result.size_column)

    def test_calculates_pp_piece_using_the_yibo_formula(self) -> None:
        result = calculate_extended_spec("yibo", 'PP NY2150P 2116 RC58% 21.5"*24.5"', self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual(11.54, result.price)
        self.assertIn("21.5×25.4/1000", result.note)
        self.assertEqual("PCS", result.size_column)

    def test_uses_the_highest_yield_ccl_parent_for_piece_price(self) -> None:
        result = calculate_extended_spec("yibo", "CCL NY2150 0.076mm H/H 1080*1 18*24IN HTE", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual(36.15, result.price)
        self.assertIn("74*49开8片", result.note)

    def test_allows_only_the_explicit_rtf2_markup_note(self) -> None:
        result = calculate_extended_spec("yibo", "CCL NY3170M 0.076mm H/H 1080*1 37*49IN RTF2", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual(149.84, result.price)
        self.assertIn("RTF2 按 RTF 加6%", result.note)

    def test_honors_the_sheet_note_limiting_pnl_to_six_pieces(self) -> None:
        result = calculate_extended_spec("yibo", "CCL NY3170M 0.076mm H/H 1080*1 18*24IN RTF", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual(36.4, result.price)
        self.assertIn("37*49开4片", result.note)

    def test_does_not_approximately_match_ccl_thickness(self) -> None:
        result = calculate_extended_spec("yibo", "CCL NY2170H 0.100mm 2/2oz 1086*1 41*49IN HTE/HTE", self.rules)

        self.assertEqual("失败", result.status)
        self.assertEqual("未匹配", result.price)

    def test_supports_hyphenated_ny_p_series_models(self) -> None:
        result = calculate_extended_spec("yibo", "PP NY-P3P 1080 RC67% 18*24IN", self.rules)

        self.assertEqual("成功", result.status)
        self.assertEqual(46.01, result.price)


if __name__ == "__main__":
    unittest.main()
