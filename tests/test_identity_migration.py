from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app.database.config import IdentityDatabaseConfig
from fangzheng_web_app.identity_migration.copy import USER_COLUMNS
from fangzheng_web_app.identity_migration.rollback import apply_changes_to_sqlite
from fangzheng_web_app.identity_migration.schema import MIGRATION_DIR


def create_users_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE users (
            employee_id TEXT PRIMARY KEY,
            display_name TEXT,
            department TEXT,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            enabled INTEGER NOT NULL DEFAULT 1,
            must_change_password INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )


class IdentityMigrationTests(unittest.TestCase):
    def test_default_config_keeps_sqlite_and_redacts_url(self) -> None:
        with patch.dict(os.environ, {"IDENTITY_DATABASE_URL": "postgresql://secret"}, clear=True):
            config = IdentityDatabaseConfig.from_env()
        self.assertEqual("sqlite", config.backend)
        self.assertNotIn("database_url", config.redacted_summary())
        self.assertTrue(config.redacted_summary()["database_url_configured"])

    def test_postgresql_requires_url_and_explicit_switch(self) -> None:
        with patch.dict(os.environ, {"IDENTITY_DATABASE_BACKEND": "postgresql"}, clear=True):
            with self.assertRaisesRegex(ValueError, "IDENTITY_DATABASE_URL"):
                IdentityDatabaseConfig.from_env()
        environment = {
            "IDENTITY_DATABASE_BACKEND": "postgresql",
            "IDENTITY_DATABASE_URL": "postgresql://configured-but-locked",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(RuntimeError, "formal switch approval"):
                IdentityDatabaseConfig.from_env()

    def test_ddl_contains_only_users_as_the_business_table(self) -> None:
        sql = (MIGRATION_DIR / "0001_users.sql").read_text(encoding="utf-8").lower()
        declared = {
            line.split("(", 1)[0].split()[-1]
            for line in sql.splitlines()
            if line.startswith("create table if not exists ")
        }
        self.assertEqual(
            {"users", "identity_runtime_flags", "identity_change_log"},
            declared,
        )

    def test_change_replay_is_idempotent_and_rejects_wrong_schema(self) -> None:
        row = dict(zip(USER_COLUMNS, (
            "E001", "User", "Sales", "hash", "admin", 1, 0,
            "2026-08-21T00:00:00", "2026-08-21T00:00:00",
        )))
        with tempfile.TemporaryDirectory() as directory:
            database = Path(directory) / "users.db"
            connection = sqlite3.connect(database)
            try:
                create_users_table(connection)
                change = {"operation": "insert", "employee_id": "E001", "row_json": row}
                self.assertEqual(1, apply_changes_to_sqlite(connection, [change]))
                self.assertEqual(1, apply_changes_to_sqlite(connection, [change]))
                self.assertEqual(1, connection.execute("SELECT COUNT(*) FROM users").fetchone()[0])
                invalid = {**row, "unexpected": "value"}
                with self.assertRaisesRegex(ValueError, "schema mismatch"):
                    apply_changes_to_sqlite(
                        connection,
                        [{"operation": "update", "employee_id": "E001", "row_json": invalid}],
                    )
            finally:
                connection.close()


if __name__ == "__main__":
    unittest.main()
