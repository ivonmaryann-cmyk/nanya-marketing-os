from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.database.sql import sqlite_to_postgresql
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_entry_service import (
    get_or_create_template,
    queue_template_extraction,
    run_template_extraction_task,
    save_template,
    template_progress,
)
from fangzheng_web_app.order_intake_service import bootstrap_cases, list_cases


class OrderEntryTemplateTaskTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "tasks.sqlite3")
        self.db_patch.start()
        db.init_db()
        self.account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            self.account_id, folder="INBOX", uid="task-1", message_id="<task-1@example.com>",
            subject="采购订单", sender="buyer@example.com", sent_at="2026-08-19 09:00:00",
            received_at="2026-08-19 09:00:00", body_html="", body_text="PO-20260819",
            eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", self.account_id)
        self.case_id = list_cases("employee-a", "2026-08-19", "new_order", self.account_id)[0]["id"]

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_first_extraction_is_queued_then_worker_marks_template_ready(self) -> None:
        with patch("fangzheng_web_app.order_entry_service.subprocess.Popen") as popen:
            queued = queue_template_extraction(self.case_id, "employee-a")
            duplicate = queue_template_extraction(self.case_id, "employee-a")
        self.assertEqual(queued["status"], "queued")
        self.assertEqual(duplicate["task_id"], queued["task_id"])
        self.assertEqual(popen.call_count, 1)
        self.assertEqual(template_progress(self.case_id, "employee-a")["stage"], "extracting")

        header = {"customer_order_number": "PO-20260819"}
        lines = [{"values": {"line_no": "1", "customer_product_code": "CUST-1", "quantity": "20"}, "sources": {}}]
        with patch("fangzheng_web_app.order_entry_service._initial_template_data", return_value=(header, lines)):
            run_template_extraction_task(queued["task_id"], self.case_id, "employee-a")

        progress = template_progress(self.case_id, "employee-a")
        self.assertTrue(progress["created"])
        self.assertEqual(progress["stage"], "pending_template_save")
        with db.db_cursor() as conn:
            task = conn.execute("SELECT status,template_id FROM order_entry_template_tasks").fetchone()
        self.assertEqual(task["status"], "completed")
        self.assertIsNotNone(task["template_id"])

    def test_postgresql_adapter_returns_the_new_task_identity(self) -> None:
        statement, returns_identity = sqlite_to_postgresql(
            "INSERT INTO order_entry_template_tasks(case_id,employee_id,status,message,started_at) VALUES (?,?,?,?,?)"
        )
        self.assertTrue(returns_identity)
        self.assertTrue(statement.endswith("RETURNING id"))

    def test_failed_worker_keeps_a_retryable_error_state(self) -> None:
        with patch("fangzheng_web_app.order_entry_service.subprocess.Popen"):
            queued = queue_template_extraction(self.case_id, "employee-a")
        with patch("fangzheng_web_app.order_entry_service.get_or_create_template", side_effect=RuntimeError("PDF 无法读取")):
            run_template_extraction_task(queued["task_id"], self.case_id, "employee-a")
        progress = template_progress(self.case_id, "employee-a")
        self.assertEqual(progress["stage"], "extraction_error")
        self.assertIn("PDF 无法读取", progress["task_message"])

    def test_successful_entry_is_the_terminal_progress_for_all_views(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        saved = save_template(self.case_id, "employee-a", {
            "header": {"order_type": "220", "bill_to_customer_code": "C001", "ledger": "KL01"},
            "lines": [{"values": {
                "line_no": "1", "customer_product_code": "CUST-1", "quantity": "20",
            }}],
        })
        with db.db_cursor() as conn:
            conn.execute(
                """INSERT INTO order_interface_call_logs
                   (case_id,template_id,employee_id,interface_key,status,created_at)
                   VALUES (?,?,?,?,?,?)""",
                (self.case_id, saved["id"], "employee-a", "domestic_order_entry", "success", db.utcnow()),
            )
        progress = template_progress(self.case_id, "employee-a")
        self.assertTrue(progress["completed"])
        self.assertEqual(progress["stage"], "completed")
        self.assertEqual(progress["label"], "已完成")

    def test_saved_template_enters_material_lookup(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"order_type": "220", "bill_to_customer_code": "C001", "ledger": "KL01"},
            "lines": [{"values": {
                "line_no": "1", "customer_product_code": "CUST-1", "quantity": "20",
            }}],
        })

        progress = template_progress(self.case_id, "employee-a")

        self.assertEqual(progress["stage"], "pending_interface_submit")
        self.assertEqual(progress["label"], "待批量料号查询")
        self.assertEqual(progress["next_action"], "批量料号查询")
