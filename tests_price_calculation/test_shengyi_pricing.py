from __future__ import annotations

import math
import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fangzheng_web_app.price_calculation_extended import (  # noqa: E402
    ExtCclRule,
    ExtRules,
    _shengyi_adjusted_sf,
    _shengyi_copper_aliases,
    _shengyi_foil_compatible,
    _shengyi_thickness_matches,
    calculate_extended_spec,
    load_extended_rules,
)
from fangzheng_web_app.price_calculation_rules import (  # noqa: E402
    get_active_price_rule_version,
    get_price_rule_file_path,
)


def money(value: float) -> float:
    return round(float(value) + 1e-9, 2)


class ShengyiPricingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.version = get_active_price_rule_version("shengyi")
        cls.rules = load_extended_rules(
            "shengyi",
            get_price_rule_file_path("shengyi", cls.version),
        )

    def ccl_row(self) -> ExtCclRule:
        return next(
            row
            for row in self.rules.ccl_rows
            if row.sheet == "NY2150"
            and row.thickness_mm is not None
            and abs(row.thickness_mm - 0.076) <= 0.001
            and row.copper == "1/1"
            and row.stack == "1*1080"
            and row.prices.get("SF") is not None
        )

    def test_ccl_dimension_formula_branches(self) -> None:
        row = self.ccl_row()
        sf = float(row.prices["SF"])
        expected = {
            "24.00X49.00": money(money(sf * 12) * 2 / 3),
            "24.20X49.00": money(money(sf * 12) * 2 / 3),
            "24.50X49.00": money(money(sf * 12) * 2 / 3),
            "27.00X49.00": money(money(sf * 13.33) * 2 / 3),
            "27.20X49.00": money(money(sf * 13.33) * 2 / 3),
            "27.50X49.00": money(money(sf * 13.33) * 2 / 3),
            "28.00X49.00": money(money(sf * 14) * 2 / 3),
            "28.20X49.00": money(money(sf * 14) * 2 / 3),
            "28.50X49.00": money(money(sf * 14) * 2 / 3),
            "37.00X49.00": money(sf * 12),
            "41.00X49.00": money(sf * 13.333),
            "43.00X49.00": money(sf * 14),
            "37.00X43.00": money(money(money(sf * 12) * 43 / 48) * 1.07),
            "41.00X43.00": money(money(money(sf * 13.33) * 43 / 48) * 1.07),
        }
        for size, expected_price in expected.items():
            with self.subTest(size=size):
                result = calculate_extended_spec(
                    "shengyi",
                    f"FR-4 NY2150 1/1 0.076mm {size} 1X1080",
                    self.rules,
                )
                self.assertEqual(result.status, "成功")
                self.assertEqual(result.price, expected_price)

        narrow = calculate_extended_spec(
            "shengyi",
            "FR-4 NY2150 1/1 0.076mm 37.00X43.00 1X1080",
            self.rules,
        )
        self.assertEqual(narrow.size_column, "37*43")
        self.assertIn("*1.07", narrow.note)

    def test_ccl_small_piece_formula(self) -> None:
        row = self.ccl_row()
        sf = float(row.prices["SF"])
        result = calculate_extended_spec(
            "shengyi",
            "FR-4 NY2150 1/1 0.076mm 20.00X10.00 1X1080",
            self.rules,
        )
        self.assertEqual(result.status, "成功")
        self.assertEqual(result.price, money(sf * 20 * 48 / math.floor(49.5 / 10) / 144))
        self.assertIn("20*48/4/144", result.note)

    def test_pp_per_meter_prices_keep_roll_length_matching(self) -> None:
        for length in (150, 200, 300):
            row = next(
                row
                for row in self.rules.pp_rows
                if row.product == "NY2150P"
                and row.length == length
                and row.sf_price is not None
                and row.rc_min is not None
            )
            rc = row.rc_min
            result = calculate_extended_spec(
                "shengyi",
                f"PP NY2150 {row.glass} RC={rc:g} 49.5X{length}M",
                self.rules,
            )
            with self.subTest(length=length):
                self.assertEqual(result.status, "成功")
                self.assertEqual(result.price, money(row.price))
                self.assertEqual(result.roll_length, f"{length}m")

    def test_pp_small_piece_uses_per_sf_price(self) -> None:
        source = next(
            row
            for row in self.rules.pp_rows
            if row.product == "NY2150P"
            and row.glass == "7628"
            and row.sf_price is not None
            and row.rc_min is not None
        )
        result = calculate_extended_spec(
            "shengyi",
            f"PP NY2150 7628 RC={source.rc_min:g} 20.00X10.00",
            self.rules,
        )
        matched = next(row for row in self.rules.pp_rows if row.excel_row == result.rule_row and row.sheet == result.size_column)
        self.assertEqual(result.status, "成功")
        self.assertEqual(result.price, money(float(matched.sf_price) * 20 * 48 / 4 / 144))
        self.assertIn("*48/4/144", result.note)

    def test_copper_aliases(self) -> None:
        self.assertEqual(_shengyi_copper_aliases("1/H"), {"1/1", "1/H", "H/1"})
        self.assertEqual(_shengyi_copper_aliases("H/2"), {"1/2", "2/1", "H/2", "2/H"})

    def test_foil_adjustments(self) -> None:
        def row(sheet: str, foil: str) -> ExtCclRule:
            return ExtCclRule(1, sheet, "NY2150", 0.076, 3.0, "1/1", foil, "1*1080", {"SF": 10.0})

        self.assertEqual(_shengyi_adjusted_sf(row("NY2150", "HTE"), "RTF", "1/1")["sf"], 10.5)
        self.assertEqual(_shengyi_adjusted_sf(row("NY-A1", "HTE"), "RTF", "1/1")["sf"], 10.5)
        self.assertEqual(_shengyi_adjusted_sf(row("NY3170M", "RTF"), "RTF2", "1/1")["sf"], 11.68)
        self.assertEqual(_shengyi_adjusted_sf(row("NY6200", "RTF"), "RTF2", "1/1")["sf"], 11.68)
        self.assertEqual(_shengyi_adjusted_sf(row("NY3170M2", "RTF2"), "RTF", "1/1")["sf"], 10.0)
        self.assertEqual(_shengyi_adjusted_sf(row("NY6300(C)", "HVLP1"), "RTF2", "1/1")["sf"], 7.74)
        self.assertEqual(_shengyi_adjusted_sf(row("NY6300S", "RTF2"), "RTF", "1/1")["sf"], 10.0)
        self.assertEqual(_shengyi_adjusted_sf(row("NY6300S", "RTF2"), "RTF3", "1/1")["sf"], 11.3)
        self.assertEqual(_shengyi_adjusted_sf(row("NY6300SN", "RTF2"), "RTF4", "1/1")["sf"], 13.0)
        self.assertEqual(_shengyi_adjusted_sf(row("NY6300SN", "RTF2"), "HVLP1", "1/1")["sf"], 13.6)

    def test_thickness_exact_row_wins_over_fallback(self) -> None:
        exact = ExtCclRule(10, "exact", "NY2150", 0.100, 3.94, "1/1", "HTE", "1*1080", {"SF": 10.0})
        nearby = ExtCclRule(11, "nearby", "NY2150", 0.102, 4.02, "1/1", "HTE", "1*1080", {"SF": 20.0})
        result = calculate_extended_spec(
            "shengyi",
            "FR-4 NY2150 1/1 0.100mm 37.00X49.00 1X1080",
            ExtRules("shengyi", [], [exact, nearby]),
        )
        self.assertEqual(result.rule_row, 10)
        self.assertEqual(result.price, 120.0)

    def test_thickness_fallback_is_limited_to_three_microns(self) -> None:
        nearby = ExtCclRule(11, "nearby", "NY2150", 0.102, 4.02, "1/1", "HTE", "1*1080", {"SF": 20.0})
        within = calculate_extended_spec(
            "shengyi",
            "FR-4 NY2150 1/1 0.100mm 37.00X49.00 1X1080",
            ExtRules("shengyi", [], [nearby]),
        )
        self.assertEqual(within.rule_row, 11)
        self.assertIn("厚度0.1mm未找到精确报价", within.note)
        self.assertTrue(_shengyi_thickness_matches(nearby, None, 0.100, 0.003))
        self.assertFalse(_shengyi_thickness_matches(nearby, None, 0.360, 0.003))

    def test_thickness_examples_and_foil_combinations(self) -> None:
        row = ExtCclRule(1, "test", "NY2150", 0.178, 7.0, "1/1", "HTE", "1*1080", {"SF": 10.0})
        self.assertTrue(_shengyi_thickness_matches(row, None, 0.180, 0.003))
        self.assertFalse(_shengyi_thickness_matches(row, None, 0.360, 0.003))
        self.assertTrue(_shengyi_foil_compatible("HVLP1", "HVLP1/HS1"))
        self.assertTrue(_shengyi_foil_compatible("RTF3/RTF", "RTF3"))
        self.assertFalse(_shengyi_foil_compatible("RTF", "RTF2"))

    def test_1027_stack_is_loaded_for_shengyi_ccl(self) -> None:
        result = calculate_extended_spec(
            "shengyi",
            "FR-4 NY6666N 1/2 RTF 0.076mm 43.00X49.00 2X1027",
            self.rules,
        )
        self.assertIsNotNone(result.rule_row)
        self.assertEqual(result.size_column, "43*49")


if __name__ == "__main__":
    unittest.main()
