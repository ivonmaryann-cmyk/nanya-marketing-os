"""No-network P2 contract tests; runnable without pytest."""
from __future__ import annotations

import json
import os
import unittest
from unittest.mock import patch

from flask import Flask

from fangzheng_web_app.connector_api import bp


class ConnectorP2Test(unittest.TestCase):
    def setUp(self):
        self.previous = os.environ.get("CONNECTOR_API_TOKEN")
        os.environ["CONNECTOR_API_TOKEN"] = "connector-p2-test-token"
        app = Flask(__name__)
        app.testing = True
        app.register_blueprint(bp)
        self.client = app.test_client()
        self.headers = {"Authorization": "Bearer connector-p2-test-token", "X-Nanya-Employee-Id": "23582"}

    def tearDown(self):
        if self.previous is None:
            os.environ.pop("CONNECTOR_API_TOKEN", None)
        else:
            os.environ["CONNECTOR_API_TOKEN"] = self.previous

    def call(self, name, arguments):
        return self.client.post("/api/connector/v1/mcp", headers=self.headers, json={
            "jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": name, "arguments": arguments},
        })

    def test_material_query_uses_existing_service_context(self):
        with patch("fangzheng_web_app.connector_api.build_material_query", return_value={"mode": "mock", "status": "success"}) as call:
            response = self.call("marketing.material.query", {"case_id": 9, "line_nos": [1, 2]})
        self.assertEqual(response.status_code, 200)
        call.assert_called_once_with(9, "23582", "23582", line_nos={1, 2})
        self.assertFalse(json.loads(response.get_json()["result"]["content"][0]["text"]).get("secret"))

    def test_prepare_does_not_call_submit_service(self):
        preview = {"can_submit": True, "payload": {"sctoDataList": []}, "interface": {"mode": "mock"}}
        with patch("fangzheng_web_app.connector_api.prepare_domestic_order_entry", return_value=preview) as prepare, patch("fangzheng_web_app.connector_api.build_domestic_order_entry") as submit:
            response = self.call("marketing.order.prepare_entry", {"case_id": 9})
        self.assertEqual(response.status_code, 200)
        prepare.assert_called_once_with(9, "23582")
        submit.assert_not_called()

    def test_submit_requires_explicit_true_confirmation(self):
        with patch("fangzheng_web_app.connector_api.build_domestic_order_entry") as submit:
            denied = self.call("marketing.order.submit_entry", {"case_id": 9, "confirm": False})
            accepted = self.call("marketing.order.submit_entry", {"case_id": 9, "confirm": True})
        self.assertTrue(denied.get_json()["result"]["isError"])
        submit.assert_called_once_with(9, "23582", "23582")
        self.assertEqual(accepted.status_code, 200)

    def test_reply_draft_cannot_send_mail(self):
        with patch("fangzheng_web_app.connector_api.build_order_reply_draft", return_value={"sent": False, "body": "draft"}) as draft:
            response = self.call("marketing.order.reply_draft", {"case_id": 9})
        self.assertEqual(response.status_code, 200)
        draft.assert_called_once_with(9, employee_id="23582")


if __name__ == "__main__":
    unittest.main()
