from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_entry_service import get_or_create_template, save_template
from fangzheng_web_app.order_interface_service import (
    build_domestic_order_entry_mock,
    build_material_query_mock,
    get_interface_config,
    is_domestic_order_entry_completed,
    get_order_detail_records,
    get_material_resolution_states,
    list_interface_configs,
    process_material_created_callback,
    save_interface_config,
    test_interface_config,
    validate_domestic_order_entry,
)
from fangzheng_web_app.order_intake_service import bootstrap_cases, list_cases


class OrderInterfaceMaintenanceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "interfaces.sqlite3")
        self.db_patch.start()
        db.init_db()
        self.account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            self.account_id, folder="INBOX", uid="interface-1", message_id="<interface-1@example.com>",
            subject="采购订单", sender="buyer@example.com", sent_at="2026-08-19 09:00:00",
            received_at="2026-08-19 09:00:00", body_html="", body_text="PO-20260819",
            eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", self.account_id)
        self.case_id = list_cases("employee-a", "2026-08-19", "new_order", self.account_id)[0]["id"]

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_default_configs_are_seeded_and_versioned(self) -> None:
        configs = list_interface_configs()
        self.assertEqual({item["interface_key"] for item in configs}, {"material_batch_query", "domestic_order_entry"})
        material = get_interface_config("material_batch_query")
        self.assertEqual(material["mode"], "mock")
        saved = save_interface_config("material_batch_query", {
            "display_name": material["display_name"],
            "description": material["description"],
            "enabled": "1",
            "mode": "mock",
            "method": "POST",
            "endpoint_url": material["endpoint_url"],
            "request_mapping": "{}",
            "response_mapping": "{}",
            "mock_scenarios": "{}",
        }, "23582")
        self.assertEqual(saved["config_version"], 2)
        self.assertEqual(saved["endpoint_url"], material["endpoint_url"])

    def test_mock_interface_test_uses_the_full_endpoint_url(self) -> None:
        result = test_interface_config({
            "interface_key": "material_batch_query",
            "mode": "mock",
            "method": "POST",
            "endpoint_url": "https://mock.nouya.local/material/batch-query",
        })
        self.assertTrue(result["ok"])
        self.assertEqual(result["status_code"], 200)
        self.assertEqual(result["request"]["items"][0]["customer_part_no"], "TEST-001")

    def test_template_events_and_mock_query_do_not_overwrite_manual_fields(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "product_code": "MANUAL-001",
                "customer_product_code": "CUST-002", "customer_spec": "FR-4 1.6",
                "quantity": "20",
            }}],
        })
        result = build_material_query_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual(len(result["items"]), 1)
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertTrue(any(item["event_type"] == "template_extracted" for item in details["events"]))
        self.assertTrue(any(item["event_type"] == "template_saved" for item in details["events"]))
        self.assertTrue(any(item["event_type"] == "material_query_mock" for item in details["events"]))
        self.assertEqual(len(details["calls"]), 1)
        self.assertTrue(details["events"][0]["trace_id"].startswith("E-"))
        self.assertRegex(details["events"][0]["occurred_at"], r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
        self.assertTrue(details["calls"][0]["trace_id"].startswith("I-"))
        self.assertEqual(details["calls"][0]["interface_label"], "批量料号查询")
        self.assertEqual(details["calls"][0]["endpoint_url"], "https://mock.nouya.local/material/batch-query")
        self.assertEqual(details["calls"][0]["method"], "POST")
        with db.db_cursor() as conn:
            value = conn.execute(
                "SELECT values_json FROM order_entry_template_lines"
            ).fetchone()["values_json"]
        self.assertIn("MANUAL-001", value)

    def test_mock_error_and_timeout_are_preserved_as_failed_calls(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        failed = build_material_query_mock(
            self.case_id, "employee-a", "employee-a", scenario="business_error"
        )
        timeout = build_material_query_mock(
            self.case_id, "employee-a", "employee-a", scenario="timeout"
        )
        self.assertEqual(failed["status"], "failed")
        self.assertEqual(timeout["status"], "failed")
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertEqual([item["status"] for item in details["calls"]], ["failed", "failed"])

    def test_domestic_entry_mock_records_the_submission_without_changing_template(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "product_code": "MANUAL-001", "customer_product_code": "CUST-001",
                "customer_spec": "FR-4 1.6", "quantity": "20",
            }}],
        })
        result = build_domestic_order_entry_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual(result["status"], "success")
        self.assertTrue(result["entry_no"].startswith("MOCK-SO-"))
        self.assertTrue(is_domestic_order_entry_completed(self.case_id, "employee-a"))
        with db.db_cursor() as conn:
            case = conn.execute(
                "SELECT status,workflow_stage,erp_prepare_status,completed_at FROM order_intake_cases WHERE id=?",
                (self.case_id,),
            ).fetchone()
        self.assertEqual(
            (case["status"], case["workflow_stage"], case["erp_prepare_status"]),
            ("archived", "completed", "submitted"),
        )
        self.assertTrue(case["completed_at"])
        with self.assertRaisesRegex(ValueError, "不能重复提交"):
            build_domestic_order_entry_mock(self.case_id, "employee-a", "employee-a")
        with self.assertRaisesRegex(ValueError, "不能再次请求料号查询接口"):
            build_material_query_mock(self.case_id, "employee-a", "employee-a")
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertEqual(details["calls"][0]["interface_key"], "domestic_order_entry")
        self.assertTrue(any(item["event_type"] == "domestic_order_entry_mock" for item in details["events"]))

    def test_material_mock_backfills_then_callback_completes_creation(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [
                {"values": {"line_no": "1", "customer_product_code": "CUST-001", "customer_spec": "规格一", "customer_spec_match": "匹配一", "quantity": "1"}},
                {"values": {"line_no": "2", "customer_product_code": "CUST-002", "customer_spec": "规格二", "customer_spec_match": "匹配二", "quantity": "1"}},
                {"values": {"line_no": "3", "customer_product_code": "CUST-003", "customer_spec": "规格三", "customer_spec_match": "匹配三", "quantity": "1"}},
            ],
        })
        result = build_material_query_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual([item["status"] for item in result["items"]], ["matched", "creating", "failed"])
        states = get_material_resolution_states(self.case_id, "employee-a")["items"]
        self.assertEqual([item["status"] for item in states], ["resolved", "waiting_callback", "failed"])
        self.assertTrue(validate_domestic_order_entry(self.case_id, "employee-a"))
        waiting = next(item for item in states if item["status"] == "waiting_callback")
        process_material_created_callback(
            waiting["correlation_id"], product_code="CALLBACK-002", product_name="回调品名", source="test",
        )
        states = get_material_resolution_states(self.case_id, "employee-a")["items"]
        self.assertEqual(states[1]["status"], "resolved")
        with db.db_cursor() as conn:
            values = conn.execute(
                "SELECT values_json FROM order_entry_template_lines WHERE line_no=2"
            ).fetchone()["values_json"]
        self.assertIn("CALLBACK-002", values)
