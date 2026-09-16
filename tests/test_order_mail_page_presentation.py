from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flask import Flask, render_template, session

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_intake_service import bootstrap_cases, get_case, list_cases
from fangzheng_web_app.routes import (
    _filter_order_cases_by_mail_content,
    _filter_order_cases_by_nyeos_order_number,
    _filter_order_cases_by_status,
    _order_automation_list_progress,
    _order_mail_status_key,
    _template_customer_order_numbers,
    order_automation_reply_query_order,
)


class OrderMailPagePresentationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "presentation.sqlite3")
        self.db_patch.start()
        db.init_db()

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_list_exposes_concise_sender_summary_and_first_attachment(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_id, _ = mail_store.upsert_message(
            account_id,
            folder="INBOX",
            uid="3001",
            message_id="<3001@example.com>",
            subject="采购订单 PO-3001",
            sender='"buyer@customer.com" <buyer@customer.com>',
            sent_at="2026-08-18 09:00:00",
            received_at="2026-08-18 09:00:00",
            body_html="",
            body_text="请确认采购订单 PO-3001，并于本周回复交期。",
            eml_path="",
            is_order=1,
        )
        mail_store.replace_attachments(
            mail_id,
            [{
                "filename": "PO-3001.pdf",
                "content_type": "application/pdf",
                "size_bytes": 2048,
                "sha256": "test",
                "stored_path": "/tmp/PO-3001.pdf",
                "is_inline": 0,
                "parse_status": "pending",
            }],
        )
        bootstrap_cases("employee-a", account_id)

        case = list_cases("employee-a", "2026-08-18", "all", account_id)[0]

        self.assertEqual(case["sender_display"], "buyer@customer.com")
        self.assertEqual(case["sender_email"], "")
        self.assertIn("PO-3001", case["summary"])
        self.assertEqual(case["first_attachment_name"], "PO-3001.pdf")
        self.assertEqual(case["first_attachment_type"], "application/pdf")

        detail = get_case(case["id"], "employee-a")
        self.assertTrue(detail["attachments"][0]["previewable"])

    def test_display_status_key_uses_entry_progress_and_merges_archived_into_completed(self) -> None:
        self.assertEqual(
            _order_mail_status_key(
                {"status": "pending_review"},
                {"stage": "pending_interface_submit"},
            ),
            "pending_interface_submit",
        )
        self.assertEqual(_order_mail_status_key({"status": "archived"}, None), "completed")
        self.assertEqual(_order_mail_status_key({"status": "on_hold"}, None), "on_hold")

        cases = [
            {"id": 1, "status": "pending_review"},
            {"id": 2, "status": "archived"},
            {"id": 3, "status": "on_hold"},
        ]
        progresses = {1: {"stage": "pending_interface_submit"}}
        self.assertEqual(
            [item["id"] for item in _filter_order_cases_by_status(
                cases, progresses, "pending_interface_submit"
            )],
            [1],
        )
        self.assertEqual(
            [item["id"] for item in _filter_order_cases_by_status(cases, progresses, "completed")],
            [2],
        )

    def test_nyeos_order_number_search_is_case_insensitive_and_partial(self) -> None:
        cases = [{"id": 1}, {"id": 2}, {"id": 3}]
        numbers = {1: "SA2608270003", 2: "SA2608270018"}

        self.assertEqual(
            [item["id"] for item in _filter_order_cases_by_nyeos_order_number(cases, numbers, "270003")],
            [1],
        )
        self.assertEqual(
            [item["id"] for item in _filter_order_cases_by_nyeos_order_number(cases, numbers, "sa260827")],
            [1, 2],
        )

    def test_mail_content_search_matches_title_or_body_case_insensitively(self) -> None:
        cases = [
            {"id": 1, "subject": "DJ订单变更", "body_text": "请调整交期"},
            {"id": 2, "subject": "普通通知", "body_text": "客户要求修改数量"},
            {"id": 3, "subject": "报价", "body_text": "请确认"},
        ]

        self.assertEqual(
            [item["id"] for item in _filter_order_cases_by_mail_content(cases, "dj订单")],
            [1],
        )
        self.assertEqual(
            [item["id"] for item in _filter_order_cases_by_mail_content(cases, "修改数量")],
            [2],
        )
        self.assertEqual(
            _filter_order_cases_by_mail_content(cases, "  "), cases,
        )

    def test_order_change_list_exposes_reply_only_after_aps_submission(self) -> None:
        case = {"id": 8, "action_type": "order_change"}
        with patch(
            "fangzheng_web_app.routes.order_change_template_progress",
            return_value={"stage": "pending_reply", "label": "待回复邮件"},
        ):
            self.assertEqual(
                _order_automation_list_progress(case, "employee-a"),
                {
                    "stage": "pending_reply", "label": "待回复邮件",
                    "completed": True, "replied": False,
                },
            )
        with patch(
            "fangzheng_web_app.routes.order_change_template_progress",
            return_value={"stage": "saved", "label": "修改模板已保存"},
        ):
            self.assertIsNone(_order_automation_list_progress(case, "employee-a"))

    def test_reply_query_collects_distinct_header_and_line_order_numbers(self) -> None:
        self.assertEqual(
            _template_customer_order_numbers({
                "header": {"customer_order_number": " PO-001 "},
                "lines": [
                    {"values": {"customer_order_number": "PO-001"}},
                    {"values": {"customer_order_number": "PO-002"}},
                    {"values": {"customer_order_number": ""}},
                ],
            }),
            ["PO-001", "PO-002"],
        )

    def test_reply_query_control_is_only_rendered_for_new_orders(self) -> None:
        app = Flask(__name__, template_folder=str(Path(__file__).parents[1] / "templates"))
        app.secret_key = "test-secret"
        app.jinja_env.globals["url_for"] = lambda *_args, **_kwargs: "/test"
        context = {
            "draft": {"to": "buyer@example.com", "cc": "", "subject": "Re: PO", "body_html": "<p>正文</p>"},
            "reply_history": [], "smtp_ready": False, "reply_back_endpoint": "main.order_automation_entry_template",
            "order_template": {}, "line_count": 0, "reply_lines": [], "return_context": {"query": {}},
        }
        with app.test_request_context("/reply"):
            new_order_html = render_template(
                "order_automation_reply.html", case={"id": 1, "action_type": "new_order", "subject": "PO"},
                **context,
            )
            change_html = render_template(
                "order_automation_reply.html", case={"id": 2, "action_type": "order_change", "subject": "PO"},
                **context,
            )
        self.assertIn('id="orderQueryOpen"', new_order_html)
        self.assertIn('id="orderQueryDrawer"', new_order_html)
        self.assertNotIn('id="orderQueryOpen"', change_html)
        self.assertNotIn('id="orderQueryDrawer"', change_html)

    def test_reply_query_route_uses_saved_template_order_numbers(self) -> None:
        app = Flask(__name__)
        app.secret_key = "test-secret"
        template = {
            "id": 9,
            "header": {"customer_order_number": "PO-001"},
            "lines": [
                {"values": {"customer_order_number": "PO-001"}},
                {"values": {"customer_order_number": "PO-002"}},
            ],
        }
        with app.test_request_context("/order-automation/cases/7/reply/query-order", method="POST"), patch(
            "fangzheng_web_app.routes.require_login", return_value=None,
        ), patch(
            "fangzheng_web_app.routes.current_employee", return_value="employee-a",
        ), patch(
            "fangzheng_web_app.routes.get_order_intake_case",
            return_value={"id": 7, "action_type": "new_order"},
        ), patch(
            "fangzheng_web_app.routes.get_saved_order_entry_template",
            return_value=({"id": 7}, template),
        ), patch(
            "fangzheng_web_app.routes.query_order_info_readonly",
            return_value={"mode": "mock", "order_count": 0, "orders": [], "not_found": []},
        ) as query:
            response = order_automation_reply_query_order(7)

        self.assertTrue(response.get_json()["ok"])
        query.assert_called_once_with(7, 9, "employee-a", "employee-a", ["PO-001", "PO-002"])

    def test_optimized_templates_render_with_list_and_detail_data(self) -> None:
        app = Flask(__name__, template_folder=str(Path(__file__).parents[1] / "templates"))
        app.secret_key = "test-secret"
        app.jinja_env.globals["url_for"] = lambda *_args, **_kwargs: "/test"
        account = {"id": 7, "email": "orders@example.com"}
        list_case = {
            "id": 1,
            "subject": "采购订单 PO-3001",
            "sender_display": "采购部",
            "sender_email": "buyer@example.com",
            "sent_at": "2026-08-18 09:00:00",
            "received_at": "",
            "customer_match_status": "matched",
            "customer_name": "测试客户",
            "routing_matches": [{"scope": "subject", "keyword": "采购订单"}],
            "change_tags": [],
            "summary": "请确认订单并回复交期。",
            "first_attachment_name": "PO-3001.pdf",
            "first_attachment_type": "application/pdf",
            "attachment_count": 1,
            "action_type": "new_order",
            "routing_state": "routed",
            "status": "pending_review",
        }
        with app.test_request_context("/order-automation"):
            session["employee_id"] = "employee-a"
            list_html = render_template(
                "order_automation.html",
                selected_account=account,
                mail_accounts=[account],
                latest_fetch_task={"status": "completed", "completed_at": "2026-08-18T10:00:00", "new_count": 1},
                fetch_tasks=[],
                selected_date="2026-08-18",
                selected_action="all",
                selected_mail_status="pending_interface_submit",
                selected_mail_query="确认订单",
                selected_order_number="SA2608270003",
                previous_date="2026-08-17",
                next_date="2026-08-19",
                date_counts=[],
                cases=[list_case],
                nyeos_order_numbers={1: "SA2608270003"},
                entry_progresses={1: {"next_action": "提取订单信息"}},
                mail_status_filter_labels={
                    "pending_interface_submit": "订单信息确认",
                    "completed": "已完成",
                },
                counts={"total": 1, "needs_business_routing": 0, "new_order": 1, "order_change": 0, "quotation": 0, "unrouted": 0},
                work_summary={"active_total": 1, "needs_routing": 0, "completed_today": 0, "pending": 1, "in_progress": 0, "awaiting_confirmation": 0, "on_hold": 0, "by_type": {"new_order": 1, "order_change": 0, "quotation": 0}},
                action_labels={"unclassified": "暂不分流", "new_order": "录单", "order_change": "修改订单", "quotation": "报价"},
                scope_labels={"subject": "邮件主题"},
                total_cases=1,
                page=1,
                per_page=20,
                total_pages=1,
                page_start=0,
            )
            detail_html = render_template(
                "order_automation_case.html",
                case={
                    **list_case,
                    "sender": "采购部 <buyer@example.com>",
                    "status": "pending_triage",
                    "routing_reason": "明确分流依据匹配",
                    "routing_source": "keyword_rule",
                    "customer_match_detail": "唯一匹配",
                    "body_text": "请确认订单并回复交期。",
                    "display_html": "<p>请确认订单并回复交期。</p>",
                    "handling_note": "",
                    "attachments": [{"id": 1, "filename": "PO-3001.pdf", "content_type": "application/pdf", "size_bytes": 2048, "parse_status": "parsed", "is_inline": 0, "previewable": True}],
                },
                entry_progress={"step": 2, "next_action": "提取订单信息", "label": "待提取订单"},
                nyeos_order_number="SA2608270003",
                return_context={"url": "/test", "values": {"category": "all"}, "query": {}},
                status_labels={"pending_triage": "待处理"},
                action_labels={"unclassified": "暂不分流", "new_order": "录单", "order_change": "修改订单", "quotation": "报价"},
                match_status_labels={"matched": "已匹配负责客户"},
            )

        self.assertIn("同步新邮件", list_html)
        self.assertIn("补抓近 30 天邮件", list_html)
        self.assertIn('aria-label="邮件状态"', list_html)
        self.assertIn('<option value="pending_interface_submit" selected>订单信息确认</option>', list_html)
        self.assertIn('name="return_mail_status" value="pending_interface_submit"', list_html)
        self.assertIn("已匹配客户：测试客户", list_html)
        self.assertIn("NYEOS订单号：SA2608270003", list_html)
        self.assertIn('name="order_no" value="SA2608270003"', list_html)
        self.assertIn('name="mail_query" value="确认订单"', list_html)
        self.assertIn("业务分流与进度", detail_html)
        self.assertIn("NYEOS订单号", detail_html)
        self.assertIn("SA2608270003", detail_html)
        self.assertIn("查看纯文本邮件正文", detail_html)
        self.assertIn('id="routingToggle"', detail_html)
        self.assertIn('aria-expanded="false"', detail_html)
        self.assertIn('id="sourceMailHtml"', detail_html)


if __name__ == "__main__":
    unittest.main()
