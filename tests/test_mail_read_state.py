from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_fetch_service, mail_store
from fangzheng_web_app.order_intake_service import bootstrap_cases, list_cases
from fangzheng_web_app.routes import _filter_order_cases_by_read_state


class _WritableImap:
    def __init__(self) -> None:
        self.selected = None
        self.store_args = None
        self.logged_out = False

    def select(self, mailbox, readonly=False):
        self.selected = (mailbox, readonly)
        return "OK", [b""]

    def uid(self, command, uid, mode, flags):
        self.store_args = (command, uid, mode, flags)
        return "OK", [b""]

    def logout(self):
        self.logged_out = True


class MailReadStateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "read-state.sqlite3")
        self.db_patch.start()
        db.init_db()
        self.account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def _message(self, *, is_seen: bool = False) -> int:
        mail_id, _ = mail_store.upsert_message(
            self.account_id,
            folder="INBOX",
            uid="9001",
            message_id="<9001@example.com>",
            subject="采购订单 PO-9001",
            sender="buyer@example.com",
            sent_at="2026-09-10 10:00:00",
            received_at="2026-09-10 10:00:00",
            body_html="",
            body_text="请处理采购订单。",
            eml_path="",
            is_order=1,
            is_seen=is_seen,
        )
        return mail_id

    def test_fetch_flag_parser_handles_seen_and_unseen(self) -> None:
        self.assertTrue(mail_fetch_service._is_seen_flag(b"1 (UID 9001 FLAGS (\\Seen) BODY[] {12}"))
        self.assertFalse(mail_fetch_service._is_seen_flag(b"1 (UID 9001 FLAGS (\\Answered) BODY[] {12}"))

    def test_read_state_filter_keeps_only_requested_rows(self) -> None:
        rows = [{"id": 1, "is_seen": 0}, {"id": 2, "is_seen": 1}]

        self.assertEqual([item["id"] for item in _filter_order_cases_by_read_state(rows, "unread")], [1])
        self.assertEqual([item["id"] for item in _filter_order_cases_by_read_state(rows, "read")], [2])

    def test_list_exposes_mailbox_read_state(self) -> None:
        self._message(is_seen=False)
        bootstrap_cases("employee-a", self.account_id)

        case = list_cases("employee-a", "2026-09-10", account_id=self.account_id)[0]

        self.assertEqual(case["is_seen"], 0)

    def test_opening_message_marks_imap_and_local_state_as_seen(self) -> None:
        mail_id = self._message(is_seen=False)
        client = _WritableImap()

        with patch.object(mail_fetch_service, "_connect", return_value=client):
            marked = mail_fetch_service.mark_message_seen_remotely(
                mail_id, owner_employee_id="employee-a"
            )

        self.assertTrue(marked)
        self.assertEqual(client.selected, ("INBOX", False))
        self.assertEqual(client.store_args, ("STORE", "9001", "+FLAGS.SILENT", r"(\Seen)"))
        self.assertTrue(client.logged_out)
        self.assertEqual(
            mail_store.get_message(mail_id, owner_employee_id="employee-a")["is_seen"], 1
        )
