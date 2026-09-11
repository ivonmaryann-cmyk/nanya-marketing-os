import unittest

from fangzheng_web_app.shennan_service import (
    _base_foil_for_copper_token,
    _calculate_ccl_surcharge_for_context,
    _copper_from_shennan_token,
    _foil_detail_from_shennan_token,
    _laminate_from_structure,
    _normalize_laminate_key,
    GLASS_CODE_MAP,
)


class ShennanCopperPricingTests(unittest.TestCase):
    def test_copper_tokens_follow_confirmed_mapping(self) -> None:
        cases = {
            "1/1": ("1/1", "", ""),
            "2/2": ("2/2", "", ""),
            "3/3": ("3/3", "", ""),
            "S1/S1": ("1/1", "RTF1", "RTF"),
            "S2/S2": ("2/2", "RTF", "RTF"),
            "R21/R21": ("1/1", "RTF2", "RTF"),
            "R31/R31": ("1/1", "RTF3", "RTF"),
            "HV1/HV1": ("1/1", "HVLP1", "HTE"),
            "HV21/HV21": ("1/1", "HVLP2", "HTE"),
        }

        for token, expected in cases.items():
            with self.subTest(token=token):
                self.assertEqual(_copper_from_shennan_token(token), expected[0])
                self.assertEqual(_foil_detail_from_shennan_token(token), expected[1])
                self.assertEqual(_base_foil_for_copper_token(token), expected[2])

    def test_per_sf_surcharge_reuses_original_size_cell_formula(self) -> None:
        selected = {
            "copper": "1/1",
            "prices": {"RMB/SF": 10.0, '40"*48"': 150.0},
            "foil_adjustments": [{
                "target": "RTF2",
                "type": "per_sf",
                "values": {"1": {"single": 0.5, "double": 1.0}},
            }],
            "copper_adjustment": {},
        }
        result = _calculate_ccl_surcharge_for_context(
            {"copper_token": "R21/R21"},
            "NYTEST",
            selected,
            {},
            {"size_col": '40"*48"', "multiplier": 2, "qty": 4, "tail_factor": 0.9},
            67.5,
        )

        self.assertIsNotNone(result)
        self.assertAlmostEqual(result["adjusted_price"], 74.25)
        self.assertAlmostEqual(result["amount"], 6.75)
        self.assertIn("原40\"*48\"单元格公式系数15", result["rule_text"])

    def test_percent_surcharge_is_applied_to_sf_before_size_formula(self) -> None:
        selected = {
            "copper": "1/1",
            "prices": {"RMB/SF": 10.0, '36"*48"': 125.0},
            "foil_adjustments": [{"target": "RTF1", "type": "percent", "percent": 0.03}],
            "copper_adjustment": {},
        }
        result = _calculate_ccl_surcharge_for_context(
            {"copper_token": "S1/S1"},
            "NYTEST",
            selected,
            {},
            {"size_col": '36"*48"', "multiplier": 1, "qty": 1, "tail_factor": 1.0},
            125.0,
        )

        self.assertIsNotNone(result)
        self.assertAlmostEqual(result["adjusted_price"], 128.75)
        self.assertAlmostEqual(result["amount"], 3.75)
        self.assertIn("RTF1：SF×(1+3%)", result["rule_text"])
        self.assertIn("SF 10→10.3", result["rule_text"])

    def test_laminate_tokens_keep_their_order_and_are_not_merged(self) -> None:
        self.assertEqual(
            {key: GLASS_CODE_MAP[key] for key in "ABCDEFGHIJKLMNOPQRSTUVZ"},
            {
                "A": "106", "B": "1065", "C": "1067", "D": "1078", "E": "1080",
                "F": "1086", "G": "2112", "H": "2113", "I": "2313", "J": "3313",
                "K": "2116", "L": "2165", "M": "1500", "N": "1501", "O": "1504",
                "P": "1506", "Q": "1652", "R": "6700", "S": "7627", "T": "7628",
                "U": "7629", "V": "7630", "Z": "1037",
            },
        )
        self.assertEqual(GLASS_CODE_MAP["LVW"], "1035")
        self.assertEqual(_laminate_from_structure("2T1K2T"), "7628x2+2116x1+7628x2")
        self.assertEqual(
            _normalize_laminate_key("7628×2+2116×1+7628×2"),
            "7628x2+2116x1+7628x2",
        )
        self.assertNotEqual(
            _normalize_laminate_key("7628×2+2116×1+7628×2"),
            _normalize_laminate_key("7628×4+2116×1"),
        )

    def test_column_level_foil_adjustment_does_not_assume_single_or_double(self) -> None:
        selected = {
            "copper": "1/1",
            "prices": {"RMB/SF": 10.0, '36"*48"': 120.0},
            "foil_adjustments": [{
                "target": "HVLP3",
                "type": "per_sf",
                "values": {"1": {"any": 3.75}},
            }],
            "copper_adjustment": {},
        }
        result = _calculate_ccl_surcharge_for_context(
            {"copper_token": "HV3/HV3"},
            "NYTEST",
            selected,
            {},
            {"size_col": '36"*48"', "multiplier": 1, "qty": 1, "tail_factor": 1.0},
            120.0,
        )
        self.assertIsNotNone(result)
        self.assertAlmostEqual(result["adjusted_price"], 165.0)


if __name__ == "__main__":
    unittest.main()
