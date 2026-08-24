from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flask import Flask, render_template, session

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_intake_service import bootstrap_cases, get_case, list_cases


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
                previous_date="2026-08-17",
                next_date="2026-08-19",
                date_counts=[],
                cases=[list_case],
                entry_progresses={1: {"next_action": "提取订单信息"}},
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
                return_context={"url": "/test", "values": {"category": "all"}, "query": {}},
                status_labels={"pending_triage": "待处理"},
                action_labels={"unclassified": "暂不分流", "new_order": "录单", "order_change": "修改订单", "quotation": "报价"},
                match_status_labels={"matched": "已匹配负责客户"},
            )

        self.assertIn("同步新邮件", list_html)
        self.assertIn("补抓近 30 天邮件", list_html)
        self.assertIn("已匹配客户：测试客户", list_html)
        self.assertIn("业务分流与进度", detail_html)
        self.assertIn("查看清洗后的 HTML 正文", detail_html)


if __name__ == "__main__":
    unittest.main()
