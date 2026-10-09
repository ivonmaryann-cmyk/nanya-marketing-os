import unittest

from openpyxl import Workbook

from fangzheng_web_app.price_calculation_extended import (
    ExtCclRule, ExtRules, _parse_mingyang_conditions, calculate_extended_spec,
)


class MingyangConditionsTests(unittest.TestCase):
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
