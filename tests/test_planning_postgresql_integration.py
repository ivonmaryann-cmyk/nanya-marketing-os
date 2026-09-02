from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

import psycopg
from psycopg.rows import dict_row

from fangzheng_web_app import db, task_backup
from fangzheng_web_app.database.planning import close_planning_pool
from fangzheng_web_app.planning_migration.copy import copy_snapshot
from fangzheng_web_app.planning_migration.rollback import (
    ENABLE_CONFIRMATION,
    ROLLBACK_CONFIRMATION,
    enable_change_capture,
    file_sha256,
    replay_change_log,
)
from fangzheng_web_app.planning_migration.schema import apply_migrations
from fangzheng_web_app.planning_migration.verify import verify_snapshot


@unittest.skipUnless(os.getenv("PLANNING_TEST_DATABASE_URL"), "PLANNING_TEST_DATABASE_URL is not configured")
class PlanningPostgresqlIntegrationTests(unittest.TestCase):
    def test_copy_runtime_crud_backup_restore_and_rollback_replay(self) -> None:
        database_url = os.environ["PLANNING_TEST_DATABASE_URL"]
        with tempfile.TemporaryDirectory() as directory:
            sqlite_path = Path(directory) / "source.db"
            backup_path = Path(directory) / "source.backup.db"
            with patch.object(db, "DATABASE_PATH", sqlite_path), patch.dict(
                os.environ,
                {"PLANNING_DATABASE_BACKEND": "sqlite", "PLANNING_POSTGRESQL_READ_WRITE_ENABLED": "false"},
                clear=False,
            ):
                db.init_db()
                category_id = db.create_task_category("E001", "Initial", "I")
                db.create_personal_task(
                    "E001", title="Initial task", category_id=category_id, due_date="2026-08-20"
                )
            shutil.copy2(sqlite_path, backup_path)
            backup_hash = file_sha256(backup_path)

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    apply_migrations(target)
                    target.execute("UPDATE planning_runtime_flags SET value='false'")
                    target.execute("TRUNCATE planning_change_log RESTART IDENTITY")
                    target.execute("TRUNCATE personal_tasks,task_categories,feedback RESTART IDENTITY CASCADE")
                    first = copy_snapshot(sqlite_path, target)
                    second = copy_snapshot(sqlite_path, target)
                    self.assertEqual(first, second)
                    verification = verify_snapshot(sqlite_path, target)
                    self.assertTrue(verification["ok"])
                    self.assertEqual(0, verification["task_category_orphans"])

            enable_change_capture(database_url, backup_path, backup_hash, ENABLE_CONFIRMATION)
            environment = {
                "PLANNING_DATABASE_BACKEND": "postgresql",
                "PLANNING_POSTGRESQL_READ_WRITE_ENABLED": "true",
                "PLANNING_DATABASE_URL": database_url,
            }
            close_planning_pool()
            try:
                with patch.dict(os.environ, environment, clear=False), patch.object(
                    task_backup, "BACKUP_ROOT", Path(directory) / "task_backups"
                ):
                    extra_id = db.create_task_category("E001", "Extra", "E")
                    task_id = db.create_personal_task(
                        "E001", title="PostgreSQL task", category_id=extra_id,
                        priority="high", progress="in_progress", due_date="2026-08-20",
                    )
                    self.assertEqual(2, len(db.list_personal_tasks("E001", due_before="2026-08-21")))
                    self.assertEqual(1, db.reorder_personal_tasks("E001", [task_id], extra_id))
                    feedback_id = db.create_feedback("E001", "suggestion", content="runtime probe")
                    db.update_feedback_status(feedback_id, "completed", "verified")
                    self.assertEqual("completed", db.get_feedback(feedback_id)["status"])
                    status = task_backup.save_task_backup("E001")
                    self.assertEqual(2, status["task_count"])
                    task_backup.restore_task_backup("E001")
                    self.assertEqual(2, len(db.list_personal_tasks("E001", archived=None)))
            finally:
                close_planning_pool()

            replay = replay_change_log(
                sqlite_path, database_url, backup_path, backup_hash, ROLLBACK_CONFIRMATION,
                batch_size=1000,
            )
            self.assertGreater(replay["applied"], 0)
            self.assertEqual(0, replay["remaining"])
            with closing(sqlite3.connect(sqlite_path)) as source:
                self.assertEqual(2, source.execute("SELECT COUNT(*) FROM personal_tasks").fetchone()[0])
                self.assertEqual(1, source.execute("SELECT COUNT(*) FROM feedback").fetchone()[0])

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    target.execute("TRUNCATE planning_change_log RESTART IDENTITY")
                    target.execute("TRUNCATE personal_tasks,task_categories,feedback RESTART IDENTITY CASCADE")
                    target.execute("UPDATE planning_runtime_flags SET value='false'")


if __name__ == "__main__":
    unittest.main()
