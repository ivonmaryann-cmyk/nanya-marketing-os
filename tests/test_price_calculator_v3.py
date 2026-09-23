from __future__ import annotations

import unittest

import pandas as pd

from fangzheng_web_app import price_calculator_v3 as calculator


PRICE_COLUMNS = [
    "CCL", "型号", "不含铜板厚/（mm)", "铜厚", "铜箔", "叠构", "RMB/SF",
    '36"*48"', '40"*48"', '42"*48"',
]


class FangzhengPriceCalculatorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.price_rules = pd.DataFrame(
            [
                ["CCL", "NY6300(C)", "0.089", "1/1", "HVLP1", "1037x2", 24.2, 290.4, 322.59, 338.8],
                ["PP", "NY6300P(C)", "106", "77", "49.5", "300", 6.06, 23858.27, None, None],
                ["PP", "NY317HFP", "2116", "53", "49.5", "200", 2.7, 7086.61, None, None],
            ],
            columns=PRICE_COLUMNS,
        )
        self.account_rules = pd.DataFrame(columns=["品名", "小片数量", "大板规格"])

    def test_ccl_accepts_unquoted_standard_size(self) -> None:
        price, _note, error = calculator.calculate_price(
            "NY6300(C) 0.089mm 1/1 37*49(1037*2)(HVLP1)",
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(290.4, price)

    def test_ccl_falls_back_from_hvlp1_to_hvlp_when_quote_uses_legacy_name(self) -> None:
        self.price_rules.loc[len(self.price_rules)] = [
            "CCL", "NY6300S", "0.203", "H/H", "HVLP", "3313x2", 19.88, 238.56, 265.0, 278.32,
        ]

        price, _note, error = calculator.calculate_price(
            'NY6300S 0.203mm H/H 37"*49"(3313*2)(HVLP1)(无卤素)',
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(238.56, price)

    def test_ccl_falls_back_from_hvlp_to_hvlp1_when_quote_uses_standard_name(self) -> None:
        self.price_rules.loc[len(self.price_rules)] = [
            "CCL", "NY6300", "0.127", "H/H", "HVLP1", "2116x1", 16.64, 199.68, 221.81, 232.96,
        ]

        price, _note, error = calculator.calculate_price(
            'NY6300 0.127mm H/H 43"x49"有卤 HVLP 1x2116',
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(232.96, price)

    def test_ccl_prefers_exact_hvlp1_quote_over_legacy_hvlp_fallback(self) -> None:
        self.price_rules.loc[len(self.price_rules)] = [
            "CCL", "NY6300S", "0.203", "H/H", "HVLP", "3313x2", 19.88, 238.56, 265.0, 278.32,
        ]
        self.price_rules.loc[len(self.price_rules)] = [
            "CCL", "NY6300S", "0.203", "H/H", "HVLP1", "3313x2", 20.0, 240.0, 266.0, 279.0,
        ]

        price, _note, error = calculator.calculate_price(
            'NY6300S 0.203mm H/H 37"*49"(3313*2)(HVLP1)(无卤素)',
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(240.0, price)

    def test_ccl_normalizes_special_dual_foil_codes(self) -> None:
        self.assertEqual(
            "HVLP2/RTF",
            calculator.parse_foil_type("(HS2-M2-VSP/RTF)", "H/2"),
        )
        self.assertEqual(
            "HTE/HVLP2",
            calculator.parse_foil_type("(HTE/HS2-M2-VSP)", "H/2"),
        )

    def test_ccl_matches_normalized_special_dual_foil_codes(self) -> None:
        for foil, spec in (
            ("HVLP2/RTF", 'NY-P1 0.064mm H/2 41"*49"(1080*1)(HS2-M2-VSP/RTF)(有卤素)'),
            ("HTE/HVLP2", 'NY-P1 0.064mm H/2 41"*49"(1080*1)(HTE/HS2-M2-VSP)(有卤素)'),
        ):
            with self.subTest(foil=foil):
                self.price_rules.loc[len(self.price_rules)] = [
                    "CCL", "NY-P1", "0.064", "2/2", foil, "1080x1", 10.0, 120.0, 132.0, 143.0,
                ]
                price, _note, error = calculator.calculate_price(
                    spec,
                    self.price_rules,
                    self.account_rules,
                )

                self.assertIsNone(error)
                self.assertEqual(132.0, price)

    def test_pp_accepts_unquoted_piece_size(self) -> None:
        price, _note, error = calculator.calculate_price(
            "NY6300P(C) 106 RC77% 30.1*24.5",
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertGreater(price, 0)

    def test_pp_alias_matches_abbreviated_quote_model(self) -> None:
        price, _note, error = calculator.calculate_price(
            'NY3170HFP 2116 RC53% 21.7"*24.6" 无卤',
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(9.76, price)

    def test_pp_roll_accepts_chinese_resin_content_without_width(self) -> None:
        self.price_rules.loc[len(self.price_rules)] = [
            "PP", "NY2170", "2116", ">57", "48", "200", 2.011, 5278.22, None, None,
        ]
        spec = "NY2170 2116 200M/卷 黄色 含量58%"

        price, note, error = calculator.calculate_price(
            spec,
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(5276.86, price)
        self.assertIn("公式=2.011×200×13.12 = 5276.86", note)
        self.assertEqual(5276.86, calculator.calculate_pp_roll_price(spec, self.price_rules))

    def test_pp_roll_uses_sf_roll_length_and_13_12_with_two_decimals(self) -> None:
        self.price_rules.loc[len(self.price_rules)] = [
            "PP", "NY6300P(C)", "2116", "57", "49.5", "200", 6.96, None, None, None,
        ]

        price, note, error = calculator.calculate_price(
            'PP NY6300P(C) 2116 RC57% 49.5"*200M/Roll(有卤素)',
            self.price_rules,
            self.account_rules,
        )

        self.assertIsNone(error)
        self.assertEqual(18263.04, price)
        self.assertTrue(note.endswith("= 18263.04"))

    def test_pp_roll_accepts_standard_glass_types_after_glue_model(self) -> None:
        cases = [
            ("NY2150P 7628 RC50% 49.5\" 有卤 CAF 150M/卷", "NY2150P", "7628", "50", "150"),
            ("NY3150HCP 7628 RC48% 49.5\" 无卤 CAF 150M/卷", "NY3150HCP", "7628", "48", "150"),
            ("NY2170P 7628 RC51% 49.5\" 有卤 CAF 150M/卷", "NY2170P", "7628", "51", "150"),
            ("NY-A2P 7628 RC50% 49.5\" 有卤 CAF 150M/卷 (汽车板)", "NY-A2P", "7628", "50", "150"),
            ('PP NY6300P(C) 1027 RC72% 49.5\"*300M/Roll (有卤素)', "NY6300P(C)", "1027", "72", "300"),
            ('PP NY6300SP 1037 RC73% 49.5\"*300M/Roll (无卤素)', "NY6300SP", "1037", "73", "300"),
        ]
        for spec, model, glass, rc, length in cases:
            with self.subTest(spec=spec):
                self.price_rules.loc[len(self.price_rules)] = [
                    "PP", model, glass, rc, "49.5", length, 2.0, 100.0, None, None,
                ]
                self.assertEqual((model, glass), calculator._extract_pp_glue_and_laminate(spec))
                price, _note, error = calculator.calculate_price(
                    spec,
                    self.price_rules,
                    self.account_rules,
                )
                self.assertIsNone(error)
                self.assertGreater(price, 0)
