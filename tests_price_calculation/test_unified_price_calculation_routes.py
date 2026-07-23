from __future__ import annotations

import sys
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch

from flask import Flask


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import routes  # noqa: E402


class UnifiedPriceCalculationRouteTests(unittest.TestCase):
    def setUp(self):
        self.app = Flask(
            __name__,
            template_folder=str(ROOT / "templates"),
            static_folder=str(ROOT / "static"),
        )
        self.app.config.update(TESTING=True, SECRET_KEY="test")
        self.app.register_blueprint(routes.bp)
        self.client = self.app.test_client()
        self.stack = ExitStack()
        self.stack.enter_context(patch.object(routes, "require_login", return_value=None))
        self.stack.enter_context(patch.object(routes, "current_employee", return_value="tester"))
        self.list_jobs = self.stack.enter_context(patch.object(routes, "list_jobs", return_value=[]))
        self.stack.enter_context(patch.object(routes, "get_active_rule_version", return_value="fangzheng-v1"))
        self.stack.enter_context(patch.object(routes, "get_active_bomin_rule_version", return_value="bomin-v1"))
        self.stack.enter_context(patch.object(routes, "get_active_shennan_rule_version", return_value="shennan-v1"))
        self.stack.enter_context(patch.object(routes, "get_active_hushi_rule_version", return_value="hushi-v1"))
        self.stack.enter_context(patch.object(routes, "get_active_price_rule_version", return_value="customer-v1"))

    def tearDown(self):
        self.stack.close()

    def test_dashboard_contains_only_unified_price_card(self):
        price_keys = {
            card["key"]
            for card in routes.FUNCTION_CARDS
            if card["key"] in {"fangzheng", "bomin", "shennan", "hushi", "price_calculation"}
        }
        self.assertEqual(price_keys, {"price_calculation"})

    def test_legacy_feature_urls_redirect_to_unified_selection(self):
        for calculator_key in ("fangzheng", "bomin", "shennan", "hushi"):
            with self.subTest(calculator_key=calculator_key):
                response = self.client.get(f"/features/{calculator_key}")
                self.assertEqual(response.status_code, 302)
                self.assertTrue(response.location.endswith(f"/features/price-calculation?calculator_key={calculator_key}"))

    def test_default_selection_is_fangzheng_and_supports_browser_memory(self):
        response = self.client.get("/features/price-calculation")
        body = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("方正价格计算：", body)
        self.assertIn('action="/jobs"', body)
        self.assertIn('data-endpoint="/api/fangzheng/quote"', body)
        self.assertIn('const selectedCalculator = "fangzheng";', body)
        self.assertIn('price-calculation-calculator-key', body)

    def test_special_calculators_use_their_existing_endpoints(self):
        expected = {
            "fangzheng": ("/jobs", "/api/fangzheng/quote", "/admin/rules"),
            "bomin": ("/bomin/jobs", "/api/bomin/quote", "/admin/bomin-rules"),
            "shennan": ("/shennan/jobs", "/api/shennan/quote", "/admin/shennan-rules"),
            "hushi": ("/hushi/jobs", "/api/hushi/quote", "/admin/hushi-rules"),
        }
        for calculator_key, (upload_url, quote_url, admin_url) in expected.items():
            with self.subTest(calculator_key=calculator_key):
                response = self.client.get(f"/features/price-calculation?calculator_key={calculator_key}")
                body = response.get_data(as_text=True)
                self.assertEqual(response.status_code, 200)
                self.assertIn(f'action="{upload_url}"', body)
                self.assertIn(f'data-endpoint="{quote_url}"', body)
                self.assertIn(f'href="{admin_url}"', body)

    def test_special_selection_reads_existing_feature_history(self):
        self.client.get("/features/price-calculation?calculator_key=bomin")
        self.list_jobs.assert_called_once_with("tester", limit=20, feature="bomin")

    def test_customer_key_compatibility_uses_existing_customer_service(self):
        response = self.client.get("/features/price-calculation?customer_key=shengyi")
        body = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("生益价格计算：", body)
        self.assertIn('action="/price-calculation/jobs"', body)
        self.assertIn('data-endpoint="/api/price-calculation/quote"', body)
        self.assertIn('name="customer_key" id="upload-customer-key" value="shengyi"', body)

    def test_jingwang_variant_remains_available(self):
        response = self.client.get("/features/price-calculation?calculator_key=jingwang&quote_variant=old")
        body = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn('name="quote_variant"', body)
        self.assertIn('value="old" selected', body)

    def test_legacy_upload_validation_returns_to_unified_page(self):
        for endpoint, calculator_key in (
            ("/jobs", "fangzheng"),
            ("/bomin/jobs", "bomin"),
            ("/shennan/jobs", "shennan"),
            ("/hushi/jobs", "hushi"),
        ):
            with self.subTest(endpoint=endpoint):
                response = self.client.post(endpoint)
                self.assertEqual(response.status_code, 302)
                self.assertTrue(response.location.endswith(f"/features/price-calculation?calculator_key={calculator_key}"))

    def test_cancel_return_keeps_the_special_calculator_selected(self):
        job = {"id": 42, "feature": "hushi", "employee_id": "tester"}
        with patch.object(routes, "get_job", return_value=job):
            response = self.client.get("/jobs/42/cancel")
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith("/features/price-calculation?calculator_key=hushi&job_id=42"))


if __name__ == "__main__":
    unittest.main()
