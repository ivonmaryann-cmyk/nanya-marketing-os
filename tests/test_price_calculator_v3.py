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
