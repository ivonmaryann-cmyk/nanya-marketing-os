import unittest

from fangzheng_web_app.bomin_service import PriceRow, RuleBook, calculate_single, parse_spec


class BominPpRollPriceTests(unittest.TestCase):
    def setUp(self):
        self.row = PriceRow(12, {
            'product': 'NY2170P', 'glass': '2116', 'rc': 0.58,
            'length_m': 200, 'per_m': 10, 'per_roll': 1999.126,
        })
        self.rules = RuleBook([], [self.row])

    def test_roll_and_unsized_pp_use_quote_roll_price(self):
        for spec in ('PP NY2170P 2116 RC 58% 49.5"*200M', 'PP NY2170P 2116 RC 58%'):
            with self.subTest(spec=spec):
                result = calculate_single(parse_spec(spec), self.rules)
                self.assertEqual('成功', result.status)
                self.assertEqual(1999.13, result.price)
                self.assertIn('Per Roll', result.note)

    def test_roll_does_not_require_per_m(self):
        self.row.values['per_m'] = None
        result = calculate_single(parse_spec('PP NY2170P 2116 RC 58% 49.5"*200M'), self.rules)
        self.assertEqual(1999.13, result.price)

    def test_missing_roll_price_does_not_return_per_m(self):
        self.row.values['per_roll'] = None
        result = calculate_single(parse_spec('PP NY2170P 2116 RC 58% 49.5"*200M'), self.rules)
        self.assertEqual('失败', result.status)
        self.assertIsNone(result.price)
        self.assertIn('Per Roll', result.note)

    def test_small_piece_still_uses_per_m(self):
        result = calculate_single(parse_spec('PP NY2170P 2116 RC 58% 经20IN*纬24IN'), self.rules)
        self.assertEqual('成功', result.status)
        self.assertEqual(2.54, result.price)
        self.assertIn('PP 小片', result.note)


if __name__ == '__main__':
    unittest.main()
