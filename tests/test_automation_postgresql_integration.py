from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

import psycopg
from psycopg.rows import dict_row

from fangzheng_web_app.automation_migration.schema import apply_migrations
from fangzheng_web_app.automation_migration.copy import copy_snapshot
from fangzheng_web_app.automation_migration.cli import rollback_test_target
from fangzheng_web_app.automation_migration.shadow import run_shadow_comparison
from fangzheng_web_app.automation_migration.spec import TABLES
from fangzheng_web_app import db


@unittest.skipUnless(os.getenv("AUTOMATION_TEST_DATABASE_URL"), "AUTOMATION_TEST_DATABASE_URL is not configured")
class AutomationPostgresqlIntegrationTests(unittest.TestCase):
    def test_empty_database_migration_is_repeatable_and_scoped(self) -> None:
        with psycopg.connect(os.environ["AUTOMATION_TEST_DATABASE_URL"], row_factory=dict_row) as connection:
            first = apply_migrations(connection)
            second = apply_migrations(connection)
            self.assertTrue(first or second == [])
            self.assertEqual([], second)
            rows = connection.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_schema='public'"
            ).fetchall()
            names = {row["table_name"] for row in rows}
            self.assertTrue(set(TABLES).issubset(names))

    def test_local_snapshot_can_be_copied_with_psycopg_connection(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            sqlite_path = Path(directory) / "source.db"
            with patch.object(db, "DATABASE_PATH", sqlite_path):
                db.init_db()
            with closing(sqlite3.connect(sqlite_path)) as source:
                source.execute(
                    "INSERT INTO mail_accounts(email,created_at,updated_at) VALUES (?,?,?)",
                    ("migration-test@example.com", "2026-08-20T00:00:00", "2026-08-20T00:00:00"),
                )
                source.commit()
            with psycopg.connect(os.environ["AUTOMATION_TEST_DATABASE_URL"], row_factory=dict_row) as connection:
                with connection.transaction():
                    apply_migrations(connection)
                    counts = copy_snapshot(sqlite_path, connection)
                    self.assertEqual(1, counts["mail_accounts"])
                    self.assertEqual(set(TABLES) | {"automation_metadata"}, set(counts))
                    connection.execute("TRUNCATE automation_metadata")
                    for table in reversed(TABLES):
                        connection.execute(f'TRUNCATE TABLE "{table}" CASCADE')

    def test_shadow_comparison_runs_against_postgresql(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            sqlite_path = Path(directory) / "shadow.db"
            with patch.object(db, "DATABASE_PATH", sqlite_path):
                db.init_db()
            with psycopg.connect(os.environ["AUTOMATION_TEST_DATABASE_URL"], row_factory=dict_row) as connection:
                with connection.transaction():
                    apply_migrations(connection)
                    copy_snapshot(sqlite_path, connection)
            result = run_shadow_comparison(sqlite_path, os.environ["AUTOMATION_TEST_DATABASE_URL"])
            self.assertEqual(10, result["metric_count"])
            self.assertEqual(0, result["difference_count"])

    def test_z_rollback_removes_all_migration_owned_objects(self) -> None:
        with psycopg.connect(os.environ["AUTOMATION_TEST_DATABASE_URL"], row_factory=dict_row) as connection:
            with connection.transaction():
                apply_migrations(connection)
        rollback_test_target("DROP-AUTOMATION-TEST-DATA")
        with psycopg.connect(os.environ["AUTOMATION_TEST_DATABASE_URL"], row_factory=dict_row) as connection:
            tables = connection.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_schema='public'"
            ).fetchall()
            functions = connection.execute(
                "SELECT routine_name FROM information_schema.routines "
                "WHERE routine_schema='public' AND routine_name='automation_capture_change'"
            ).fetchall()
        self.assertEqual([], tables)
        self.assertEqual([], functions)


if __name__ == "__main__":
    unittest.main()
