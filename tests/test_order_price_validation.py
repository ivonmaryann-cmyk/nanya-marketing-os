from __future__ import annotations

import unittest
from unittest.mock import patch

from fangzheng_web_app.order_price_validation_service import (
    review_cached_template_prices,
    review_template_prices,
)
from fangzheng_web_app.price_calculation_customer_mapping import (
    PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY,
    PRICE_TAX_MODE_EXCLUSIVE,
    PRICE_TAX_MODE_INCLUSIVE,
)


class OrderPriceValidationTests(unittest.TestCase):
    def _customer(self) -> dict[str, str]:
        return {"customer_short_name": "方正", "group_name": "方正集团"}

    def test_tax_inclusive_quote_checks_unit_price(self) -> None:
        with patch.dict(PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY, {"fangzheng": PRICE_TAX_MODE_INCLUSIVE}, clear=True), patch(
            "fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote",
            return_value={"status": "成功", "price": 12.5, "note": "命中报价"},
        ):
            review = review_template_prices(self._customer(), [{
                "line_no": 1,
                "values": {"customer_spec": "规格A", "quantity": "100", "unit_price": "12.49", "price_before_tax": "10"},
            }])

        item = review["by_line"][1]
        self.assertEqual(item["field"], "unit_price")
        self.assertEqual(item["status"], "mismatch")
        self.assertEqual(item["quote_price"], "12.50")

    def test_tax_exclusive_quote_checks_price_before_tax_and_rounds_to_cents(self) -> None:
        with patch.dict(PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY, {"fangzheng": PRICE_TAX_MODE_EXCLUSIVE}, clear=True), patch(
            "fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote",
            return_value={"status": "成功", "price": "10.004", "note": "命中报价"},
        ):
            review = review_template_prices(self._customer(), [{
                "line_no": 1,
                "values": {"customer_spec": "规格A", "price_before_tax": "10.00", "unit_price": "11.30"},
            }])

        item = review["by_line"][1]
        self.assertEqual(item["field"], "price_before_tax")
        self.assertEqual(item["status"], "matched")

    def test_unknown_tax_mode_does_not_calculate_or_block(self) -> None:
        with patch.dict(PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY, {}, clear=True), patch(
            "fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote"
        ) as quote:
            review = review_template_prices(self._customer(), [{
                "line_no": 1,
                "values": {"customer_spec": "规格A", "unit_price": "12.50"},
            }])

        quote.assert_not_called()
        self.assertFalse(review["mismatches"])
        self.assertEqual(review["tax_mode"], "unknown")

    def test_empty_target_price_uses_successful_quote_as_a_suggestion(self) -> None:
        with patch.dict(PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY, {"fangzheng": PRICE_TAX_MODE_INCLUSIVE}, clear=True), patch(
            "fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote",
            return_value={"status": "成功", "price": "12.5", "note": "命中报价"},
        ):
            review = review_template_prices(self._customer(), [{
                "line_no": 1,
                "values": {"customer_spec": "规格A", "quantity": "100", "unit_price": ""},
            }])

        item = review["by_line"][1]
        self.assertEqual(item["field"], "unit_price")
        self.assertEqual(item["status"], "suggested")
        self.assertEqual(item["quote_price"], "12.50")
        self.assertEqual(item["note"], "报价单计算")
        self.assertFalse(review["mismatches"])

    def test_roll_quote_is_written_with_four_decimal_places(self) -> None:
        with patch.dict(PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY, {"fangzheng": PRICE_TAX_MODE_EXCLUSIVE}, clear=True), patch(
            "fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote",
            return_value={"status": "成功", "price": "91.3386", "note": "命中报价"},
        ):
            review = review_template_prices(self._customer(), [{
                "line_no": 35,
                "values": {
                    "customer_spec": 'PP NY6300P(C) 2116 RC57% 49.5"*200M/Roll(有卤素)',
                    "remark": "1卷", "price_before_tax": "91.3384",
                },
            }])

        self.assertEqual(review["by_line"][35]["quote_price"], "91.3386")

    def test_chaoying_suggests_its_tax_exclusive_price_before_tax(self) -> None:
        customer = {"customer_short_name": "超颖", "group_name": "定颖集团"}
        with patch(
            "fangzheng_web_app.order_price_validation_service.calculate_price_quote",
            return_value={"status": "成功", "price": "18.88"},
        ):
            review = review_template_prices(customer, [{
                "line_no": 1,
                "values": {"customer_spec": "PP 规格A", "quantity": "11", "price_before_tax": "", "unit_price": ""},
            }])

        item = review["by_line"][1]
        self.assertEqual(item["field"], "price_before_tax")
        self.assertEqual(item["status"], "suggested")
        self.assertEqual(item["quote_price"], "18.88")

    def test_unavailable_quote_is_not_a_mismatch(self) -> None:
        with patch.dict(PRICE_QUOTE_TAX_MODE_BY_CUSTOMER_KEY, {"fangzheng": PRICE_TAX_MODE_INCLUSIVE}, clear=True), patch(
            "fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote",
            return_value={"status": "失败", "price": None, "error": "未命中报价"},
        ):
            review = review_template_prices(self._customer(), [{
                "line_no": 1,
                "values": {"customer_spec": "规格A", "unit_price": "12.50"},
            }])

        self.assertEqual(review["by_line"][1]["status"], "not_checked")
        self.assertFalse(review["mismatches"])

    def test_cached_quote_rechecks_current_price_without_recalculating(self) -> None:
        snapshot = {
            "association": {"matched": True, "price_customer_key": "fangzheng"},
            "tax_mode": PRICE_TAX_MODE_INCLUSIVE,
            "target_field": "unit_price",
            "target_label": "单价",
            "by_line": {
                "1": {"line_no": 1, "quote_price": "12.50", "note": "命中报价"},
            },
        }
        with patch("fangzheng_web_app.order_price_validation_service.calculate_fangzheng_quote") as quote:
            review = review_cached_template_prices(snapshot, [{
                "line_no": 1,
                "values": {"unit_price": "12.00", "customer_spec": "已改变也不重新计算"},
            }])

        quote.assert_not_called()
        self.assertEqual(review["by_line"][1]["status"], "mismatch")
        self.assertEqual(review["mismatches"][0]["quote_price"], "12.50")
