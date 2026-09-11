from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from flask import Flask

from fangzheng_web_app.price_calculation_customer_mapping import (
    PRICE_CALCULATION_ASSOCIATIONS,
    PRICE_TAX_MODE_EXCLUSIVE,
    PRICE_TAX_MODE_INCLUSIVE,
    PRICE_TAX_MODE_UNKNOWN,
    resolve_customer_price_calculation,
    resolve_customer_price_calculation_by_id,
)
from fangzheng_web_app.routes import bp


class PriceCalculationCustomerMappingTests(unittest.TestCase):
    def test_every_configured_group_resolves_to_its_price_rule(self) -> None:
        for association in PRICE_CALCULATION_ASSOCIATIONS:
            for group_name in association["group_names"]:
                with self.subTest(group_name=group_name):
                    result = resolve_customer_price_calculation({"group_name": group_name})

                    self.assertTrue(result["matched"])
                    self.assertEqual(result["price_customer_key"], association["key"])
                    self.assertEqual(result["match_source"], "所属集团")

    def test_short_name_mapping_overrides_group_mapping(self) -> None:
        result = resolve_customer_price_calculation(
            {"customer_short_name": "泰兴", "group_name": "深南集团"}
        )

        self.assertTrue(result["matched"])
        self.assertEqual(result["price_customer_key"], "taixing")
        self.assertEqual(result["match_source"], "客户简称")

    def test_taixing_and_shennan_groups_stay_separate(self) -> None:
        taixing = resolve_customer_price_calculation({"group_name": "泰兴电路"})
        shennan = resolve_customer_price_calculation({"group_name": "深南集团"})

        self.assertEqual(taixing["price_customer_key"], "taixing")
        self.assertEqual(shennan["price_customer_key"], "shennan")

    def test_quote_tax_modes_follow_the_configured_customer_list(self) -> None:
        inclusive = resolve_customer_price_calculation({"customer_short_name": "依顿"})
        exclusive = resolve_customer_price_calculation({"customer_short_name": "方正"})
        unknown = resolve_customer_price_calculation({"customer_short_name": "泰兴"})

        self.assertEqual(inclusive["price_customer_label"], "依顿")
        self.assertEqual(inclusive["price_tax_mode"], PRICE_TAX_MODE_INCLUSIVE)
        self.assertEqual(exclusive["price_tax_mode"], PRICE_TAX_MODE_EXCLUSIVE)
        self.assertEqual(unknown["price_tax_mode"], PRICE_TAX_MODE_UNKNOWN)

    def test_legacy_eaton_short_name_keeps_resolving_to_yidun(self) -> None:
        result = resolve_customer_price_calculation({"customer_short_name": "伊顿"})

        self.assertEqual(result["price_customer_key"], "eaton")
        self.assertEqual(result["price_customer_label"], "依顿")

    def test_empty_or_unknown_customer_returns_unmatched_reason(self) -> None:
        for customer in ({}, {"group_name": "未知集团"}):
            with self.subTest(customer=customer):
                result = resolve_customer_price_calculation(customer)

                self.assertFalse(result["matched"])
                self.assertEqual(result["price_customer_key"], "")
                self.assertTrue("未配置" in result["reason"] or "不存在" in result["reason"])

    def test_customer_id_lookup_delegates_to_customer_master_data(self) -> None:
        with patch(
            "fangzheng_web_app.customer_archive_service.get_customer",
            return_value={"group_name": "骏亚集团"},
        ):
            result = resolve_customer_price_calculation_by_id(12)

        self.assertEqual(result["price_customer_key"], "junya")


class CustomerArchivePriceCalculationDisplayTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = Flask(
            __name__,
            template_folder=str(Path(__file__).parents[1] / "templates"),
            static_folder=str(Path(__file__).parents[1] / "static"),
        )
        self.app.config.update(SECRET_KEY="test-secret", TESTING=True)
        self.app.register_blueprint(bp)
        self.client = self.app.test_client()
        with self.client.session_transaction() as user_session:
            user_session["employee_id"] = "employee-a"

    def test_customer_detail_shows_price_calculation_match_and_source(self) -> None:
        workspace = {
            "customer": {
                "id": 7,
                "customer_code": "100001",
                "customer_name": "测试客户",
                "customer_short_name": "泰兴",
                "group_name": "深南集团",
                "status": "active",
            },
            "contacts": [],
            "rules": [],
            "mappings": [],
        }
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.routes.list_customers", return_value=[]
        ), patch("fangzheng_web_app.routes.customer_summary", return_value={"total": 1, "active": 1, "disabled": 0}), patch(
            "fangzheng_web_app.routes.get_customer_workspace", return_value=workspace
        ):
            response = self.client.get("/customers?customer_id=7")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("价格计算关联", html)
        self.assertIn("泰兴价格规则", html)
        self.assertIn("客户简称“泰兴”", html)


if __name__ == "__main__":
    unittest.main()
