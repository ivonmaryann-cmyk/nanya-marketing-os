from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent.mail_html_parser import extract_order_fields, safe_display_html
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_intake_service import (
    bootstrap_cases,
    business_today,
    classify_mail,
    list_cases,
    list_change_tags,
    list_universal_rules,
    save_universal_rule,
    save_universal_rule_scope,
    update_case,
    update_routing,
    work_summary,
)


class OrderIntakeRoutingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "routing.sqlite3")
        self.db_patch.start()
        db.init_db()

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_explainable_routing_priority(self) -> None:
        self.assertEqual(classify_mail("订单变更待确认")[0], "order_change")
        self.assertEqual(classify_mail("RFQ开始报价通知")[0], "quotation")
        self.assertEqual(classify_mail("新增采购订单 PO-001")[0], "new_order")
        self.assertEqual(classify_mail("请回复交期")[0], "delivery")

    def test_manual_routing_is_saved_and_filtered_by_mail_date(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_id, _ = mail_store.upsert_message(
            account_id,
            folder="INBOX",
            uid="2001",
            message_id="<2001@example.com>",
            subject="RFQ开始报价通知",
            sender="customer@example.com",
            sent_at="2026-08-18 09:00:00",
            received_at="2026-08-18 09:00:00",
            body_html="",
            body_text="",
            eml_path="",
            is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        cases = list_cases("employee-a", "2026-08-18", "quotation", account_id)
        self.assertEqual([item["mail_id"] for item in cases], [mail_id])

        updated = update_routing(cases[0]["id"], "employee-a", "new_order")
        self.assertEqual(updated["action_type"], "new_order")
        self.assertEqual(updated["routing_source"], "manual")
        self.assertEqual(list_cases("employee-a", "2026-08-17", "all", account_id), [])

    def test_universal_rule_routes_by_body_keyword(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_id, _ = mail_store.upsert_message(
            account_id, folder="INBOX", uid="2002", message_id="<2002@example.com>",
            subject="普通通知", sender="buyer@customer.com",
            sent_at="2026-08-18 09:00:00", received_at="2026-08-18 09:00:00",
            body_html="", body_text="请尽快确认报价", eml_path="", is_order=1,
        )
        rule = next(rule for rule in list_universal_rules("employee-a") if rule["action_type"] == "quotation")
        bootstrap_cases("employee-a", account_id)
        case = list_cases("employee-a", "2026-08-18", "quotation", account_id)[0]
        self.assertEqual(case["mail_id"], mail_id)
        self.assertEqual(case["routing_state"], "routed")
        self.assertTrue(any(hit["keyword"] == "报价" for hit in case["routing_matches"]))

    def test_cross_category_keyword_hits_need_business_routing(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            account_id, folder="INBOX", uid="2003", message_id="<2003@example.com>",
            subject="采购订单", sender="buyer@bominelec.com",
            sent_at="2026-08-18 09:00:00", received_at="2026-08-18 09:00:00",
            body_html="", body_text="请确认报价", eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        cases = list_cases("employee-a", "2026-08-18", "needs_business_routing", account_id)
        self.assertEqual(len(cases), 1)
        self.assertEqual(cases[0]["routing_state"], "needs_business_routing")

    def test_one_search_scope_can_be_saved_without_overwriting_others(self) -> None:
        rule = next(rule for rule in list_universal_rules("employee-a") if rule["action_type"] == "new_order")
        before_subjects = [item["keyword"] for item in rule["keywords"] if item["scope"] == "subject"]

        updated = save_universal_rule_scope("employee-a", rule["id"], "body", ["订单明细", "下单通知"])

        self.assertEqual([item["keyword"] for item in updated["keywords"] if item["scope"] == "subject"], before_subjects)
        self.assertEqual([item["keyword"] for item in updated["keywords"] if item["scope"] == "body"], ["订单明细", "下单通知"])
        self.assertFalse(any(item["action_type"] == "unclassified" for item in list_universal_rules("employee-a")))

    def test_change_item_promotes_mail_to_order_change(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            account_id, folder="INBOX", uid="2004", message_id="<2004@example.com>",
            subject="取消", sender="buyer@customer.com",
            sent_at="2026-08-18 09:00:00", received_at="2026-08-18 09:00:00",
            body_html="", body_text="", eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        cases = list_cases("employee-a", "2026-08-18", "order_change", account_id)
        self.assertEqual(len(cases), 1)
        self.assertIn("取消 / 暂停", cases[0]["change_tags"])

    def test_generic_specification_word_does_not_become_an_order_change(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            account_id, folder="INBOX", uid="2005", message_id="<2005@example.com>",
            subject="资料回复", sender="buyer@customer.com",
            sent_at="2026-08-18 09:00:00", received_at="2026-08-18 09:00:00",
            body_html="", body_text="请查阅产品规格与型号说明", eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        cases = list_cases("employee-a", "2026-08-18", "all", account_id)
        self.assertEqual(cases[0]["action_type"], "unclassified")
        self.assertEqual(cases[0]["change_tags"], [])

    def test_temporary_unrouted_state_cannot_be_saved_as_a_rule(self) -> None:
        with self.assertRaisesRegex(ValueError, "明确分流结果"):
            save_universal_rule(
                "employee-a",
                {"name": "不应保存", "action_type": "unclassified", "keywords": [{"scope": "subject", "keyword": "测试"}]},
            )

    def test_completed_task_is_counted_by_completion_day(self) -> None:
        today = business_today().isoformat()
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            account_id, folder="INBOX", uid="2006", message_id="<2006@example.com>",
            subject="采购订单", sender="buyer@customer.com",
            sent_at=f"{today} 09:00:00", received_at=f"{today} 09:00:00",
            body_html="", body_text="", eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        case = list_cases("employee-a", today, "new_order", account_id)[0]
        update_case(case["id"], "employee-a", {"status": "archived"})

        summary = work_summary("employee-a", account_id)
        self.assertEqual(summary["by_type"]["new_order"], 0)
        self.assertEqual(summary["completed_today"], 1)

    def test_mail_html_is_displayed_without_active_or_remote_content(self) -> None:
        rendered = safe_display_html(
            '<meta charset="utf-8"><style>.x{background:url(https://tracker.example/a)}</style>'
            '<p class="x">订单<br>明细</p><img src="https://tracker.example/p.png">'
            '<script>alert(1)</script>'
        )
        self.assertIn("订单<br>明细", rendered)
        self.assertIn('class="x"', rendered)
        self.assertNotIn("tracker.example", rendered)
        self.assertNotIn("script", rendered)
        self.assertNotIn("<img", rendered)

    def test_email_style_order_number_and_spec_are_suggested(self) -> None:
        fields = extract_order_fields(
            "HJ20260818013\nA1A15022449YNNYZ002\n南亚\nNY2150\nFR-4 1.5\n±\n0.075MM 2/2"
        )
        self.assertEqual(fields["order_number"], "HJ20260818013")
        self.assertIn("NY2150", fields["spec"])


if __name__ == "__main__":
    unittest.main()
