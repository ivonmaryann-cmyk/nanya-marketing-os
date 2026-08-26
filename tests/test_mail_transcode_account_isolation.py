from __future__ import annotations

import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_fetch_service, mail_store, smtp_service
from fangzheng_web_app.order_intake_service import bootstrap_cases


class _FakeImapClient:
    def __init__(self) -> None:
        self.selected = None
        self.search_args = None
        self.logged_out = False

    def select(self, mailbox, readonly=False):
        self.selected = (mailbox, readonly)
        return "OK", [b""]

    def search(self, *args):
        self.search_args = args
        return "OK", [b""]

    def logout(self):
        self.logged_out = True
        return "BYE", [b""]


class _FakeSmtpClient:
    def __init__(self) -> None:
        self.messages = []
        self.closed = False

    def noop(self):
        return 250, b"OK"

    def send_message(self, message, from_addr=None, to_addrs=None):
        self.messages.append((message, from_addr, to_addrs))
        return {}

    def quit(self):
        self.closed = True


class MailTranscodeAccountIsolationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "mail-test.sqlite3"
        self.db_patch = patch.object(db, "DATABASE_PATH", self.db_path)
        self.db_patch.start()
        db.init_db()

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_accounts_and_auth_codes_are_scoped_to_owner(self) -> None:
        owner_account = mail_store.create_or_update_account(
            "owner@example.com",
            owner_employee_id="employee-a",
            auth_code="owner-auth-code",
        )
        mail_store.create_or_update_account(
            "other@example.com",
            owner_employee_id="employee-b",
            auth_code="other-auth-code",
        )

        self.assertEqual(
            [account["email"] for account in mail_store.list_accounts(owner_employee_id="employee-a")],
            ["owner@example.com"],
        )
        self.assertIsNone(mail_store.get_account(owner_account, owner_employee_id="employee-b"))
        self.assertIsNone(
            mail_store.get_account_auth_code(owner_account, owner_employee_id="employee-b")
        )
        self.assertEqual(
            mail_store.get_account_auth_code(owner_account, owner_employee_id="employee-a"),
            "owner-auth-code",
        )

    def test_edit_updates_email_and_delete_removes_only_owned_configuration(self) -> None:
        account_id = mail_store.create_or_update_account(
            "before@example.com",
            owner_employee_id="employee-a",
            auth_code="owner-auth-code",
        )

        updated_id = mail_store.create_or_update_account(
            "after@example.com",
            owner_employee_id="employee-a",
            account_id=account_id,
            imap_host="imaphz.qiye.163.com",
            imap_port=993,
            auth_code="",
            enabled=0,
        )

        self.assertEqual(updated_id, account_id)
        updated = mail_store.get_account(account_id, owner_employee_id="employee-a")
        self.assertEqual(updated["email"], "after@example.com")
        self.assertEqual(updated["enabled"], 0)
        self.assertEqual(
            mail_store.get_account_auth_code(account_id, owner_employee_id="employee-a"),
            "owner-auth-code",
        )
        with self.assertRaisesRegex(ValueError, "无权删除"):
            mail_store.delete_account(account_id, owner_employee_id="employee-b")

        self.assertEqual(
            mail_store.delete_account(account_id, owner_employee_id="employee-a"),
            "after@example.com",
        )
        self.assertIsNone(mail_store.get_account(account_id, owner_employee_id="employee-a"))

    def test_fetch_uses_read_only_and_only_yesterday_and_today(self) -> None:
        account_id = mail_store.create_or_update_account(
            "owner@example.com",
            owner_employee_id="employee-a",
            auth_code="owner-auth-code",
        )
        client = _FakeImapClient()

        with patch.object(mail_fetch_service, "_connect", return_value=client):
            result = mail_fetch_service.fetch_latest_order_mails(
                account_id,
                owner_employee_id="employee-a",
                created_by="employee-a",
            )

        self.assertEqual(client.selected, ("INBOX", True))
        self.assertEqual(
            client.search_args,
            (
                None,
                "SINCE",
                (date.today() - timedelta(days=1)).strftime("%d-%b-%Y"),
                "BEFORE",
                (date.today() + timedelta(days=1)).strftime("%d-%b-%Y"),
            ),
        )
        self.assertTrue(client.logged_out)
        self.assertEqual(result["fetched"], 0)
        self.assertEqual(result["new_count"], 0)
        self.assertEqual(result["duplicate_count"], 0)
        self.assertIn("新增 0 封", result["message"])

    def test_connection_test_only_opens_inbox_read_only(self) -> None:
        account_id = mail_store.create_or_update_account(
            "owner@example.com",
            owner_employee_id="employee-a",
            auth_code="owner-auth-code",
        )
        client = _FakeImapClient()

        with patch.object(mail_fetch_service, "_connect", return_value=client):
            result = mail_fetch_service.test_imap_connection(
                account_id, owner_employee_id="employee-a"
            )

        self.assertEqual(client.selected, ("INBOX", True))
        self.assertTrue(client.logged_out)
        self.assertIn("未抓取或保存任何邮件", result["message"])

    def test_queue_returns_immediately_and_prevents_duplicate_active_tasks(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="owner-auth-code"
        )
        with patch("fangzheng_web_app.mail_transcode_agent.mail_fetch_service.subprocess.Popen") as popen:
            result = mail_fetch_service.queue_latest_order_mails(
                account_id, owner_employee_id="employee-a", created_by="employee-a"
            )
        self.assertTrue(popen.called)
        task = mail_store.get_fetch_task(result["fetch_task_id"], owner_employee_id="employee-a")
        self.assertEqual(task["status"], "queued")
        with self.assertRaisesRegex(ValueError, "正在运行"):
            mail_fetch_service.queue_latest_order_mails(
                account_id, owner_employee_id="employee-a", created_by="employee-a"
            )

    def test_repeated_uid_is_recorded_but_not_created_twice(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com",
            owner_employee_id="employee-a",
            auth_code="owner-auth-code",
        )
        first_batch = mail_store.create_fetch_task(account_id, created_by="employee-a")
        mail_id, is_new = mail_store.upsert_message(
            account_id,
            folder="INBOX",
            uid="1001",
            message_id="<same@example.com>",
            subject="新订单",
            sender="customer@example.com",
            sent_at="2026-08-18 09:00:00",
            received_at="2026-08-18 09:00:00",
            body_html="",
            body_text="PO-001",
            eml_path="/tmp/1001.eml",
            is_order=1,
            fetch_task_id=first_batch,
        )
        mail_store.record_fetch_task_message(first_batch, mail_id, is_new=is_new)

        second_batch = mail_store.create_fetch_task(account_id, created_by="employee-a")
        repeated_id, repeated_is_new = mail_store.upsert_message(
            account_id,
            folder="INBOX",
            uid="1001",
            message_id="<same@example.com>",
            subject="新订单",
            sender="customer@example.com",
            sent_at="2026-08-18 09:00:00",
            received_at="2026-08-18 09:00:00",
            body_html="",
            body_text="PO-001",
            eml_path="/tmp/1001.eml",
            is_order=1,
            fetch_task_id=second_batch,
        )
        mail_store.record_fetch_task_message(second_batch, repeated_id, is_new=repeated_is_new)

        self.assertTrue(is_new)
        self.assertFalse(repeated_is_new)
        self.assertEqual(mail_id, repeated_id)
        self.assertEqual(
            mail_store.get_messages_by_fetch_task(second_batch, owner_employee_id="employee-a")[0]["is_new"],
            0,
        )

    def test_smtp_configuration_is_private_and_testable_without_sending(self) -> None:
        account_id = mail_store.create_or_update_account(
            "owner@example.com", owner_employee_id="employee-a", auth_code="imap-secret"
        )
        saved = mail_store.save_smtp_config(
            account_id,
            owner_employee_id="employee-a",
            host="smtp.example.com",
            port=465,
            security="ssl",
            username="owner@example.com",
            auth_code="smtp-secret",
            sender_name="订单中心",
            enabled=0,
        )
        self.assertTrue(saved["configured"])
        self.assertFalse(saved["enabled"])
        self.assertNotIn("auth_code", saved)
        self.assertIsNone(mail_store.get_smtp_config(account_id, owner_employee_id="employee-b"))

        client = _FakeSmtpClient()
        with patch.object(smtp_service, "_connect", return_value=client):
            result = smtp_service.test_smtp_connection(account_id, owner_employee_id="employee-a")
        self.assertIn("未发送任何邮件", result["message"])
        self.assertEqual(client.messages, [])
        self.assertTrue(client.closed)
        enabled = mail_store.save_smtp_config(
            account_id,
            owner_employee_id="employee-a",
            host="smtp.example.com",
            port=465,
            security="ssl",
            username="owner@example.com",
            auth_code="",
            sender_name="订单中心",
            enabled=1,
        )
        self.assertTrue(enabled["enabled"])

    def test_order_reply_send_uses_source_mailbox_and_records_audit(self) -> None:
        account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="imap-secret"
        )
        mail_store.save_smtp_config(
            account_id,
            owner_employee_id="employee-a",
            host="smtp.example.com",
            port=465,
            security="ssl",
            username="orders@example.com",
            auth_code="smtp-secret",
            sender_name="南亚订单中心",
            enabled=0,
        )
        mail_store.set_smtp_test_status(account_id, "success", owner_employee_id="employee-a")
        mail_store.save_smtp_config(
            account_id,
            owner_employee_id="employee-a",
            host="smtp.example.com",
            port=465,
            security="ssl",
            username="orders@example.com",
            auth_code="",
            sender_name="南亚订单中心",
            enabled=1,
        )
        mail_store.upsert_message(
            account_id,
            folder="INBOX",
            uid="reply-1",
            message_id="<reply-1@example.com>",
            subject="采购订单",
            sender="Customer <customer@example.com>",
            sent_at="2026-08-19 09:00:00",
            received_at="2026-08-19 09:00:00",
            body_html="",
            body_text="PO-001",
            eml_path="",
            is_order=1,
        )
        bootstrap_cases("employee-a", account_id)
        with db.db_cursor() as conn:
            case_id = int(conn.execute("SELECT id FROM order_intake_cases").fetchone()["id"])

        client = _FakeSmtpClient()
        with patch.object(smtp_service, "_connect", return_value=client):
            result = smtp_service.send_order_reply(
                case_id,
                employee_id="employee-a",
                to="customer@example.com",
                cc="copy@example.com",
                subject="Re: 采购订单",
                body="订单已确认。",
            )
        self.assertEqual(result["to"], ["customer@example.com"])
        self.assertEqual(len(client.messages), 1)
        message, sender, recipients = client.messages[0]
        self.assertEqual(sender, "orders@example.com")
        self.assertEqual(recipients, ["customer@example.com", "copy@example.com"])
        self.assertEqual(message["Subject"], "Re: 采购订单")
        with db.db_cursor() as conn:
            event = conn.execute(
                "SELECT event_type, detail_json FROM order_entry_detail_events WHERE case_id = ?",
                (case_id,),
            ).fetchone()
        self.assertEqual(event["event_type"], "order_reply_sent")
        self.assertNotIn("订单已确认", event["detail_json"])


if __name__ == "__main__":
    unittest.main()
