from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.automation_migration.outbox import OUTBOX_TABLE, parse_primary_key
from fangzheng_web_app.automation_migration.preflight import BOOLEAN_REQUIREMENTS, evaluate_preflight
from fangzheng_web_app.automation_migration.shadow import _canonical
from fangzheng_web_app.automation_migration.sync import _apply_event, outbox_status, process_outbox


class _Result:
    def __init__(self, row=None):
        self.row = row

    def fetchone(self):
        return self.row


class _IdempotentTarget:
    def __init__(self) -> None:
        self.claimed: set[str] = set()
        self.business_writes = 0

    def execute(self, sql, params=()):
        if "automation_migration_inbox" in sql:
            event_id = params[0]
            if event_id in self.claimed:
                return _Result(None)
            self.claimed.add(event_id)
            return _Result({"event_id": event_id})
        self.business_writes += 1
        return _Result()


class AutomationShadowMigrationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "shadow.sqlite3"
        self.database_patch = patch.object(db, "DATABASE_PATH", self.database_path)
        self.database_patch.start()
        db.init_db()
        with db.db_cursor() as connection:
            connection.execute(f"DELETE FROM {OUTBOX_TABLE}")

    def tearDown(self) -> None:
        self.database_patch.stop()
        self.temp_dir.cleanup()

    def test_business_write_and_outbox_event_share_the_same_transaction(self) -> None:
        with db.db_cursor() as connection:
            connection.execute(
                """INSERT INTO mail_accounts
                   (email,owner_employee_id,auth_code_ciphertext,created_at,updated_at)
                   VALUES (?,?,?,?,?)""",
                ("shadow@example.com", "E001", "secret-ciphertext", "now", "now"),
            )
        with closing(sqlite3.connect(self.database_path)) as connection:
            connection.row_factory = sqlite3.Row
            event = connection.execute(f"SELECT * FROM {OUTBOX_TABLE}").fetchone()
        self.assertEqual("mail_accounts", event["source_table"])
        self.assertEqual("insert", event["operation"])
        self.assertEqual({"id": 1}, parse_primary_key(event["pk_json"], "mail_accounts"))
        self.assertNotIn("secret-ciphertext", event["pk_json"])

    def test_rolled_back_business_write_does_not_leave_an_outbox_event(self) -> None:
        with self.assertRaises(RuntimeError):
            with db.db_cursor() as connection:
                connection.execute(
                    "INSERT INTO order_change_tags(employee_id,name,created_at,updated_at) VALUES (?,?,?,?)",
                    ("E001", "rollback", "now", "now"),
                )
                raise RuntimeError("inject rollback")
        self.assertEqual(0, outbox_status(self.database_path)["pending"])

    def test_postgresql_failure_only_schedules_retry(self) -> None:
        with db.db_cursor() as connection:
            connection.execute(
                "INSERT INTO order_change_tags(employee_id,name,created_at,updated_at) VALUES (?,?,?,?)",
                ("E001", "retry", "now", "now"),
            )
        with patch(
            "fangzheng_web_app.automation_migration.sync.psycopg.connect",
            side_effect=ConnectionError("database unavailable"),
        ):
            result = process_outbox(self.database_path, "postgresql://not-logged")
        self.assertEqual(1, result["failed"])
        status = outbox_status(self.database_path)
        self.assertEqual(1, status["pending"])
        self.assertEqual(1, status["retrying"])
        with closing(sqlite3.connect(self.database_path)) as connection:
            self.assertEqual(1, connection.execute("SELECT COUNT(*) FROM order_change_tags WHERE name='retry'").fetchone()[0])

    def test_postgresql_inbox_makes_duplicate_delivery_idempotent(self) -> None:
        with db.db_cursor() as connection:
            connection.execute(
                "INSERT INTO order_change_tags(employee_id,name,created_at,updated_at) VALUES (?,?,?,?)",
                ("E001", "idempotent", "now", "now"),
            )
        with closing(sqlite3.connect(self.database_path)) as source:
            source.row_factory = sqlite3.Row
            event = source.execute(f"SELECT * FROM {OUTBOX_TABLE}").fetchone()
            target = _IdempotentTarget()
            self.assertTrue(_apply_event(source, target, event))
            self.assertFalse(_apply_event(source, target, event))
        self.assertEqual(1, target.business_writes)

    def test_metadata_trigger_records_only_automation_keys(self) -> None:
        with db.db_cursor() as connection:
            connection.execute("INSERT INTO settings(key,value) VALUES (?,?)", ("unrelated", "ignored"))
            connection.execute("INSERT INTO settings(key,value) VALUES (?,?)", ("order_mail_rule_seed_v2:E001", "now"))
        with closing(sqlite3.connect(self.database_path)) as connection:
            rows = connection.execute(f"SELECT source_table,pk_json FROM {OUTBOX_TABLE}").fetchall()
        self.assertEqual([("automation_metadata", json.dumps({"key": "order_mail_rule_seed_v2:E001"}, separators=(",", ":")))], rows)

    def test_shadow_hash_is_stable_and_does_not_expose_values(self) -> None:
        rows = [{"subject": "sensitive", "id": 1}]
        digest = _canonical(rows)
        self.assertEqual(64, len(digest))
        self.assertNotIn("sensitive", digest)
        self.assertEqual(digest, _canonical([{"id": 1, "subject": "sensitive"}]))

    def test_preflight_fails_closed_until_every_evidence_item_passes(self) -> None:
        failed = evaluate_preflight({})
        self.assertFalse(failed["passed"])
        evidence = {name: True for name in BOOLEAN_REQUIREMENTS}
        evidence.update({
            "shadow_observation_days": 7,
            "unexplained_shadow_differences": 0,
            "outbox_pending": 0,
            "performance_thresholds_approved": True,
            "performance_test_passed": True,
        })
        self.assertEqual({"passed": True, "blockers": []}, evaluate_preflight(evidence))


if __name__ == "__main__":
    unittest.main()
