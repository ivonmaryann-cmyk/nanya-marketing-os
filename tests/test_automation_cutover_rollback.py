from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.automation_migration.cutover import prepare_cutover
from fangzheng_web_app.automation_migration.observation import evaluate_observation_reports
from fangzheng_web_app.automation_migration.outbox import OUTBOX_TABLE
from fangzheng_web_app.automation_migration.preflight import BOOLEAN_REQUIREMENTS
from fangzheng_web_app.automation_migration.rollback import apply_changes_to_sqlite


class AutomationCutoverRollbackTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "rollback.sqlite3"
        self.database_patch = patch.object(db, "DATABASE_PATH", self.database_path)
        self.database_patch.start()
        db.init_db()
        with db.db_cursor() as connection:
            connection.execute(f"DELETE FROM {OUTBOX_TABLE}")

    def tearDown(self) -> None:
        self.database_patch.stop()
        self.temp_dir.cleanup()

    def _tag_change(self, *, name: str = "replayed", operation: str = "insert") -> dict:
        return {
            "source_table": "order_change_tags",
            "operation": operation,
            "pk_json": {"id": 101},
            "row_json": None if operation == "delete" else {
                "id": 101,
                "employee_id": "E001",
                "name": name,
                "enabled": 1,
                "created_at": "2026-08-20T00:00:00",
                "updated_at": "2026-08-20T00:00:00",
            },
        }

    def test_replay_is_idempotent_and_does_not_create_outbox_events(self) -> None:
        connection = db.get_connection()
        try:
            connection.execute("BEGIN IMMEDIATE")
            self.assertEqual(2, apply_changes_to_sqlite(connection, [self._tag_change(), self._tag_change()]))
            connection.commit()
            row = connection.execute("SELECT name FROM order_change_tags WHERE id=101").fetchone()
            self.assertEqual("replayed", row["name"])
            self.assertEqual(0, connection.execute(f"SELECT COUNT(*) FROM {OUTBOX_TABLE}").fetchone()[0])
            self.assertEqual("0", connection.execute(
                "SELECT value FROM automation_runtime_flags WHERE key='suppress_outbox'"
            ).fetchone()["value"])
        finally:
            connection.close()

    def test_replay_delete_is_idempotent(self) -> None:
        connection = db.get_connection()
        try:
            connection.execute("BEGIN IMMEDIATE")
            apply_changes_to_sqlite(connection, [self._tag_change()])
            connection.commit()
            connection.execute("BEGIN IMMEDIATE")
            apply_changes_to_sqlite(connection, [self._tag_change(operation="delete"), self._tag_change(operation="delete")])
            connection.commit()
            self.assertEqual(0, connection.execute("SELECT COUNT(*) FROM order_change_tags WHERE id=101").fetchone()[0])
        finally:
            connection.close()

    def test_replay_schema_error_rolls_back_the_entire_sqlite_transaction(self) -> None:
        invalid = self._tag_change()
        invalid["row_json"] = {"id": 101, "unexpected": "value"}
        connection = db.get_connection()
        try:
            connection.execute("BEGIN IMMEDIATE")
            with self.assertRaisesRegex(ValueError, "schema mismatch"):
                apply_changes_to_sqlite(connection, [self._tag_change(name="before-error"), invalid])
            connection.rollback()
            self.assertEqual(0, connection.execute("SELECT COUNT(*) FROM order_change_tags WHERE id=101").fetchone()[0])
            self.assertEqual("0", connection.execute(
                "SELECT value FROM automation_runtime_flags WHERE key='suppress_outbox'"
            ).fetchone()["value"])
        finally:
            connection.close()

    def test_cutover_preflight_blocks_before_any_database_connection(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "preflight"):
            prepare_cutover(self.database_path, "postgresql://must-not-connect", {}, Path(self.temp_dir.name))

    def test_cutover_requires_three_distinct_named_owners(self) -> None:
        evidence = {name: True for name in BOOLEAN_REQUIREMENTS}
        evidence.update({
            "shadow_observation_days": 7,
            "unexplained_shadow_differences": 0,
            "outbox_pending": 0,
            "performance_thresholds_approved": True,
            "performance_test_passed": True,
            "switch_owner": "A",
            "review_owner": "A",
            "rollback_owner": "C",
            "maintenance_window": "approved window",
        })
        with self.assertRaisesRegex(RuntimeError, "three distinct"):
            prepare_cutover(self.database_path, "postgresql://must-not-connect", evidence, Path(self.temp_dir.name))

    def test_observation_gate_requires_seven_reviewed_days(self) -> None:
        paths = []
        start = date(2026, 8, 1)
        for offset in range(7):
            path = Path(self.temp_dir.name) / f"observation-{offset}.json"
            path.write_text(json.dumps({
                "observed_at": f"{start + timedelta(days=offset)}T12:00:00+00:00",
                "health_ok": True,
                "outbox": {"pending": 0},
                "shadow_7d": {"differences": 0},
            }), encoding="utf-8")
            paths.append(path)
        self.assertEqual(
            {"passed": True, "observed_days": 7, "blockers": []},
            evaluate_observation_reports(paths),
        )
        paths[0].write_text(json.dumps({
            "observed_at": "2026-08-01T12:00:00+00:00",
            "health_ok": False,
            "outbox": {"pending": 0},
            "shadow_7d": {"differences": 0},
        }), encoding="utf-8")
        self.assertIn("all_health_reviews_passed", evaluate_observation_reports(paths)["blockers"])


if __name__ == "__main__":
    unittest.main()
