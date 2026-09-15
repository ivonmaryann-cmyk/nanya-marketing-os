from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_entry_service import (
    get_order_change_template,
    order_change_template_progress,
    queue_order_change_template_extraction,
    run_template_extraction_task,
    save_order_change_template,
)
from fangzheng_web_app.order_interface_service import (
    _flatten_order_info_candidates,
    _match_order_change_line,
    _store_order_change_matches,
    get_order_change_matches,
    get_order_detail_records,
    query_order_info,
    select_order_change_candidate,
)
from fangzheng_web_app.order_intake_service import bootstrap_cases, list_cases


class OrderChangeTemplateMarkupTests(unittest.TestCase):
    def test_match_details_do_not_repeat_original_template_values(self) -> None:
        markup = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_order_change_template.html"
        ).read_text(encoding="utf-8")

        self.assertNotIn("原模板明细", markup)
        self.assertNotIn("oc-original", markup)
        self.assertIn("ERP 匹配明细", markup)

    def test_erp_candidate_display_fields_hide_unneeded_headers(self) -> None:
        route_source = (
            Path(__file__).resolve().parents[1]
            / "fangzheng_web_app"
            / "routes.py"
        ).read_text(encoding="utf-8")

        self.assertIn('("scta01", "NYEOS订单号")', route_source)
        self.assertIn('("sctb35", "项次")', route_source)
        self.assertNotIn('("sctb04", "项次")', route_source)
        self.assertNotIn('("scta11", "送货客户ID")', route_source)
        self.assertNotIn('("scta38", "单头客户订单号")', route_source)
        self.assertNotIn('("peag04", "主数据品名规格")', route_source)

    def test_change_submit_modal_is_hidden_until_the_button_is_clicked(self) -> None:
        markup = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_order_change_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn('id="orderChangeModal" hidden', markup)
        self.assertIn('id="submitOrderChange" type="button">提交APS</button>', markup)
        self.assertIn('.oc-modal-mask[hidden]{display:none}', markup)
        self.assertIn("open.addEventListener('click',()=>{modal.hidden=false", markup)


class OrderChangeTemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "change.sqlite3")
        self.db_patch.start()
        db.init_db()
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            account_id, folder="INBOX", uid="change-1", message_id="<change-1@example.com>",
            subject="采购订单", sender="buyer@example.com", sent_at="2026-08-19 09:00:00",
            received_at="2026-08-19 09:00:00", body_html="", body_text="PO-CHANGE-001", eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        self.case_id = list_cases("employee-a", "2026-08-19", "new_order", account_id)[0]["id"]
        with db.db_cursor() as conn:
            conn.execute("UPDATE order_intake_cases SET action_type='order_change' WHERE id=?", (self.case_id,))

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_change_template_extracts_in_background_and_saves_the_six_fields(self) -> None:
        with patch("fangzheng_web_app.order_entry_service.subprocess.Popen"):
            queued = queue_order_change_template_extraction(self.case_id, "employee-a")
        self.assertEqual(queued["status"], "queued")
        self.assertEqual(order_change_template_progress(self.case_id, "employee-a")["stage"], "extracting")

        header = {"customer_order_number": "PO-CHANGE-001"}
        lines = [{"values": {
            "line_no": "8", "customer_product_code": "CUST-8", "customer_spec": "NY2150 1.0mm",
            "delivery_date": "2026-09-25", "quantity": "100",
        }, "sources": {}}]
        with patch("fangzheng_web_app.order_entry_service.review_case_template_prices") as price_review, patch(
            "fangzheng_web_app.order_entry_service._initial_template_data", return_value=(header, lines)
        ):
            run_template_extraction_task(
                queued["task_id"], self.case_id, "employee-a", action_type="order_change",
            )
        price_review.assert_not_called()

        _case, extracted = get_order_change_template(self.case_id, "employee-a")
        self.assertEqual(extracted["header"]["customer_order_number"], "PO-CHANGE-001")
        self.assertEqual(extracted["lines"][0]["values"]["customer_product_code"], "CUST-8")
        self.assertEqual(extracted["lines"][0]["values"]["unit_price"], "")
        extracted_details = get_order_detail_records(self.case_id, "employee-a")
        self.assertTrue(any(item["event_type"] == "order_change_template_extracted" for item in extracted_details["events"]))

        saved = save_order_change_template(self.case_id, "employee-a", {"lines": [{"values": {
            "customer_order_number": "PO-CHANGE-002", "line_no": "99", "customer_product_code": "CUST-9",
            "customer_spec": "NY2170 0.8mm", "delivery_date": "2026-10-01", "quantity": "80",
        }}]})
        values = saved["lines"][0]["values"]
        self.assertEqual(values["line_no"], "99")
        self.assertEqual(values["customer_order_number"], "PO-CHANGE-002")
        self.assertEqual(values["delivery_date"], "2026-10-01")
        self.assertEqual(order_change_template_progress(self.case_id, "employee-a")["stage"], "saved")
        saved_details = get_order_detail_records(self.case_id, "employee-a")
        self.assertTrue(any(item["event_type"] == "order_change_template_saved" for item in saved_details["events"]))

    def test_change_template_rejects_a_non_change_case(self) -> None:
        with db.db_cursor() as conn:
            conn.execute("UPDATE order_intake_cases SET action_type='new_order' WHERE id=?", (self.case_id,))
        with self.assertRaisesRegex(ValueError, "修改订单"):
            queue_order_change_template_extraction(self.case_id, "employee-a")

    def test_order_info_query_only_sends_customer_order_numbers(self) -> None:
        with patch("fangzheng_web_app.order_entry_service.subprocess.Popen"):
            queued = queue_order_change_template_extraction(self.case_id, "employee-a")
        with patch("fangzheng_web_app.order_entry_service._initial_template_data", return_value=(
            {"customer_order_number": "PO-DEFAULT"},
            [{"values": {
                "line_no": "1", "customer_order_number": "PO-QUERY-001",
                "customer_product_code": "CUST-QUERY-001", "quantity": "10",
            }, "sources": {}}],
        )):
            run_template_extraction_task(queued["task_id"], self.case_id, "employee-a", action_type="order_change")
        _case, template = get_order_change_template(self.case_id, "employee-a")
        result = query_order_info(
            self.case_id, int(template["id"]), "employee-a", "employee-a",
            ["PO-QUERY-001", "PO-QUERY-001", ""],
        )
        self.assertEqual(result["mode"], "mock")
        details = get_order_detail_records(self.case_id, "employee-a")
        call = details["calls"][0]
        self.assertEqual(call["interface_key"], "order_info_query")
        self.assertEqual(call["request"], {"orderNumberList": ["PO-QUERY-001"]})
        self.assertEqual(call["interface_label"], "查询订单信息")
        matches = get_order_change_matches(self.case_id, int(template["id"]), "employee-a")
        self.assertEqual(matches[1]["status"], "matched")
        self.assertEqual(matches[1]["selected_candidate"]["scta39"], "220-MOCK-0001")

        save_order_change_template(self.case_id, "employee-a", {"lines": [{"values": {
            "customer_order_number": "PO-QUERY-001", "line_no": "1",
            "customer_product_code": "", "quantity": "99",
        }}]})
        self.assertEqual(get_order_change_matches(self.case_id, int(template["id"]), "employee-a"), {})

    def test_order_change_matching_uses_priority_and_deduplicates_erp_rows(self) -> None:
        first = {
            "scta39": "ERP-1", "scta01": "ORDER-1", "sctb02": "PART-1", "sctb14": "Cust-A",
            "sctb15": "PO-A", "sctb35": 10, "sctb05": 100,
        }
        second = {
            "scta39": "ERP-2", "scta01": "ORDER-2", "sctb02": "PART-2", "sctb14": "CUST-A",
            "sctb15": "po-a", "sctb35": 20, "sctb05": 200,
        }
        payload = {"data": {"orderList": [
            {"scta01": "ORDER-1", "scta39": "ERP-1", "sctbList": [first, dict(first)]},
            {"scta01": "ORDER-2", "scta39": "ERP-2", "sctbList": [second]},
        ]}}
        candidates = _flatten_order_info_candidates(payload)
        self.assertEqual(len(candidates), 2)
        by_item = _match_order_change_line({
            "customer_order_number": " PO-A ", "customer_product_code": "cust-a",
            "line_no": "10", "quantity": "200.0",
        }, candidates)
        self.assertEqual(by_item["match_level"], "order_part_item")
        self.assertEqual(by_item["selected"]["scta39"], "ERP-1")
        by_quantity = _match_order_change_line({
            "customer_order_number": "PO-A", "customer_product_code": "CUST-A",
            "line_no": "99", "quantity": "200.0",
        }, candidates)
        self.assertEqual(by_quantity["match_level"], "order_part_quantity")
        self.assertEqual(by_quantity["selected"]["scta39"], "ERP-2")
        multiple = _match_order_change_line({
            "customer_order_number": "PO-A", "customer_product_code": "CUST-A",
            "line_no": "99", "quantity": "999",
        }, candidates)
        self.assertEqual(multiple["status"], "multiple")
        self.assertEqual(len(multiple["candidates"]), 2)

    def test_multiple_match_selection_is_persisted(self) -> None:
        with patch("fangzheng_web_app.order_entry_service.subprocess.Popen"):
            queued = queue_order_change_template_extraction(self.case_id, "employee-a")
        with patch("fangzheng_web_app.order_entry_service._initial_template_data", return_value=(
            {"customer_order_number": "PO-MULTI"},
            [{"values": {
                "line_no": "9", "customer_order_number": "PO-MULTI",
                "customer_product_code": "CUST-MULTI", "quantity": "99",
            }, "sources": {}}],
        )):
            run_template_extraction_task(queued["task_id"], self.case_id, "employee-a", action_type="order_change")
        _case, template = get_order_change_template(self.case_id, "employee-a")
        initial = query_order_info(
            self.case_id, int(template["id"]), "employee-a", "employee-a", ["PO-MULTI"],
        )
        response = {"data": {"orderList": [{
            "scta01": "ORDER-M", "scta39": "ERP-M",
            "sctbList": [
                {"sctb02": "PART-1", "sctb14": "CUST-MULTI", "sctb15": "PO-MULTI", "sctb35": 1, "sctb05": 1},
                {"sctb02": "PART-2", "sctb14": "CUST-MULTI", "sctb15": "PO-MULTI", "sctb35": 2, "sctb05": 2},
            ],
        }]}}
        with db.db_cursor() as conn:
            _store_order_change_matches(
                conn, case_id=self.case_id, template_id=int(template["id"]), employee_id="employee-a",
                call_id=initial["call_id"], response_payload=response,
            )
        pending = get_order_change_matches(self.case_id, int(template["id"]), "employee-a")[9]
        self.assertEqual(pending["status"], "multiple")
        selected = pending["candidates"][1]
        select_order_change_candidate(
            self.case_id, int(template["id"]), "employee-a", 9, selected["candidate_key"], "employee-a",
        )
        persisted = get_order_change_matches(self.case_id, int(template["id"]), "employee-a")[9]
        self.assertEqual(persisted["status"], "matched")
        self.assertEqual(persisted["selected_candidate"]["sctb02"], "PART-2")
