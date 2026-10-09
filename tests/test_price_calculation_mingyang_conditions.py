import unittest

from openpyxl import Workbook

from fangzheng_web_app.price_calculation_extended import (
    ExtCclRule, ExtRules, _parse_mingyang_conditions, _mingyang_price_columns, calculate_extended_spec,
)


class MingyangConditionsTests(unittest.TestCase):
    def test_month_prefixed_price_headers(self):
        self.assertEqual(
            {1: "41", 2: "37", 3: "43"},
            _mingyang_price_columns(["10月41*49\n(RMB含税）", "2026年10月37×49".replace("×", "*"), "43*49", "41*49（旧）", "10月价格"]),
        )

    def test_high_speed_nominal_thickness_with_tolerance(self):
        row = ExtCclRule(17, "通用CCL 高速", "NY3170M", 0.102, None, "1/1", "RTF", "1*2116", {"41": 250})
        result = calculate_extended_spec("mingyang", 'CCL NY3170M 0.102±0.013mm 不含铜 1/1 RTF 41*49" 2116*1', ExtRules("mingyang", [], [row]))
        self.assertEqual("成功", result.status)
        self.assertEqual(250, result.price)
        self.assertEqual(17, result.rule_row)

    def test_special_models_and_exact_thickness_priority(self):
        for product in ("NY6300(C)", "NY-P2", "NY-P3(C)", "NY-P5Q"):
            with self.subTest(product=product):
                rows = [
                    ExtCclRule(16, "通用CCL 高速", product.replace("-", ""), 0.1, None, "1/1", "RTF", "1*2116", {"41": 200}),
                    ExtCclRule(17, "通用CCL 高速", product.replace("-", ""), 0.102, None, "1/1", "RTF", "1*2116", {"41": 250}),
                ]
                result = calculate_extended_spec("mingyang", f"CCL {product} 0.102±0.013mm 1/1 RTF 41*49 2116*1", ExtRules("mingyang", [], rows))
                self.assertEqual(250, result.price)
                self.assertEqual(17, result.rule_row)

    def rules(self, amount=195, markup=3):
        ws = Workbook().active
        ws.append(["说明："])
        ws.append([f"加价原则在41*49的价格基础上进行，价格最后四舍五入保留整数。2/2铜厚的在H/H的基础上加{amount}元，1.5/1.5（W/W）铜厚的在H/H的基础上加145元，RTF铜箔在HTE上加{markup}%，2/2以上RTF铜价另议"])
        ws.append(["37*49在41*49的基础上乘以0.9，43*49在41*49的基础上剩以1.05"])
        row = ExtCclRule(9, "通用CCL", "NY2170", 0.2, None, "H/H", "HTE", "1*7628", {"41": 100})
        return ExtRules("mingyang", [], [row], {row.sheet: _parse_mingyang_conditions(ws)})

    def calculate(self, copper, rules, foil="HTE"):
        return calculate_extended_spec("mingyang", f"NY2170 0.2mm {copper} {foil} 41*49 (1*7628)", rules)

    def test_quote_changes_adjustment(self):
        self.assertEqual(295, self.calculate("2/2", self.rules()).price)
        self.assertEqual(280, self.calculate("2/2", self.rules(180)).price)
        self.assertEqual(245, self.calculate("W/W", self.rules()).price)

    def test_quote_changes_rtf_percentage(self):
        self.assertEqual(252, self.calculate("W/W", self.rules(), "RTF").price)
        self.assertEqual(257, self.calculate("W/W", self.rules(markup=5), "RTF").price)

    def test_negotiable_and_missing_conditions(self):
        self.assertEqual("失败", self.calculate("2/2", self.rules(), "RTF").status)
        rules = self.rules()
        rules.ccl_notes.clear()
        self.assertEqual("失败", self.calculate("2/2", rules).status)

    def test_direct_quote_priority(self):
        rules = self.rules()
        rules.ccl_rows.append(ExtCclRule(10, "通用CCL", "NY2170", 0.2, None, "2/2", "RTF", "1*7628", {"41": 333}))
        self.assertEqual(333, self.calculate("2/2", rules, "RTF").price)


if __name__ == "__main__":
    unittest.main()
