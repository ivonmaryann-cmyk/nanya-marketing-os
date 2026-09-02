from __future__ import annotations

import os
import sqlite3
import unittest
from unittest.mock import patch

from fangzheng_web_app.database.config import PlanningDatabaseConfig
from fangzheng_web_app.planning_migration.rollback import apply_changes_to_sqlite
from fangzheng_web_app.planning_migration.schema import MIGRATION_DIR


class PlanningMigrationTests(unittest.TestCase):
    def test_default_config_keeps_sqlite_and_redacts_url(self) -> None:
        with patch.dict(os.environ, {"PLANNING_DATABASE_URL": "postgresql://secret"}, clear=True):
            config = PlanningDatabaseConfig.from_env()
        self.assertEqual("sqlite", config.backend)
        self.assertNotIn("database_url", config.redacted_summary())
        self.assertTrue(config.redacted_summary()["database_url_configured"])

    def test_postgresql_requires_url_and_explicit_switch(self) -> None:
        with patch.dict(os.environ, {"PLANNING_DATABASE_BACKEND": "postgresql"}, clear=True):
            with self.assertRaisesRegex(ValueError, "PLANNING_DATABASE_URL"):
                PlanningDatabaseConfig.from_env()
        environment = {
            "PLANNING_DATABASE_BACKEND": "postgresql",
            "PLANNING_DATABASE_URL": "postgresql://configured-but-locked",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(RuntimeError, "formal switch approval"):
                PlanningDatabaseConfig.from_env()

    def test_ddl_is_limited_to_three_business_tables(self) -> None:
        sql = (MIGRATION_DIR / "0001_planning_schema.sql").read_text(encoding="utf-8").lower()
        declared = {
            line.split("(", 1)[0].split()[-1]
            for line in sql.splitlines()
            if line.startswith("create table if not exists ")
        }
        self.assertEqual(
            {
                "task_categories", "personal_tasks", "feedback",
                "planning_runtime_flags", "planning_change_log",
            },
            declared,
        )

    def test_change_replay_is_idempotent_and_rejects_out_of_scope_table(self) -> None:
        connection = sqlite3.connect(":memory:")
        try:
            connection.execute(
                "CREATE TABLE task_categories(id INTEGER PRIMARY KEY,employee_id TEXT,name TEXT,"
                "short_label TEXT,sort_order INTEGER,created_at TEXT,updated_at TEXT)"
            )
            row = {
                "id": 1, "employee_id": "E001", "name": "Work", "short_label": "W",
                "sort_order": 1, "created_at": "2026-08-21", "updated_at": "2026-08-21",
            }
            change = {
                "table_name": "task_categories", "operation": "insert",
                "pk_json": {"id": 1}, "row_json": row,
            }
            self.assertEqual(1, apply_changes_to_sqlite(connection, [change]))
            row["name"] = "Updated"
            self.assertEqual(1, apply_changes_to_sqlite(connection, [change]))
            self.assertEqual("Updated", connection.execute(
                "SELECT name FROM task_categories WHERE id=1"
            ).fetchone()[0])
            with self.assertRaisesRegex(ValueError, "outside the migration scope"):
                apply_changes_to_sqlite(connection, [{**change, "table_name": "users"}])
        finally:
            connection.close()


if __name__ == "__main__":
    unittest.main()
