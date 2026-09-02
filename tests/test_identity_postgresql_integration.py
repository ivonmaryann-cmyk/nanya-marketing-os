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
from werkzeug.security import generate_password_hash

from fangzheng_web_app import db
from fangzheng_web_app.database.identity import close_identity_pool
from fangzheng_web_app.identity_migration.copy import USER_COLUMNS, copy_users
from fangzheng_web_app.identity_migration.rollback import (
    ENABLE_CONFIRMATION,
    ROLLBACK_CONFIRMATION,
    enable_change_capture,
    file_sha256,
    replay_change_log,
)
from fangzheng_web_app.identity_migration.schema import apply_migrations
from fangzheng_web_app.identity_migration.verify import verify_users
from tests.test_identity_migration import create_users_table


@unittest.skipUnless(os.getenv("IDENTITY_TEST_DATABASE_URL"), "IDENTITY_TEST_DATABASE_URL is not configured")
class IdentityPostgresqlIntegrationTests(unittest.TestCase):
    def test_copy_runtime_access_and_rollback_replay(self) -> None:
        database_url = os.environ["IDENTITY_TEST_DATABASE_URL"]
        with tempfile.TemporaryDirectory() as directory:
            sqlite_path = Path(directory) / "source.db"
            backup_path = Path(directory) / "source.backup.db"
            admin_hash = generate_password_hash("admin-password")
            disabled_hash = generate_password_hash("disabled-password")
            with closing(sqlite3.connect(sqlite_path)) as source:
                create_users_table(source)
                source.executemany(
                    f"INSERT INTO users({','.join(USER_COLUMNS)}) VALUES ({','.join('?' for _ in USER_COLUMNS)})",
                    [
                        ("ADMIN", "Administrator", "IT", admin_hash, "admin", 1, 0, "2026-08-21", "2026-08-21"),
                        ("OFF", "Disabled", "Sales", disabled_hash, "user", 0, 1, "2026-08-21", "2026-08-21"),
                    ],
                )
                source.commit()
            shutil.copy2(sqlite_path, backup_path)
            backup_hash = file_sha256(backup_path)

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    apply_migrations(target)
                    target.execute("UPDATE identity_runtime_flags SET value='false'")
                    target.execute("TRUNCATE identity_change_log RESTART IDENTITY")
                    target.execute("TRUNCATE users")
                    self.assertEqual(2, copy_users(sqlite_path, target))
                    self.assertEqual(2, copy_users(sqlite_path, target))
                    verification = verify_users(sqlite_path, target)
                    self.assertTrue(verification["ok"])
                    self.assertTrue(verification["password_hashes_match"])

            enable_change_capture(database_url, backup_path, backup_hash, ENABLE_CONFIRMATION)
            environment = {
                "IDENTITY_DATABASE_BACKEND": "postgresql",
                "IDENTITY_POSTGRESQL_READ_WRITE_ENABLED": "true",
                "IDENTITY_DATABASE_URL": database_url,
            }
            close_identity_pool()
            try:
                with patch.dict(os.environ, environment, clear=False):
                    self.assertTrue(db.verify_user_password("ADMIN", "admin-password"))
                    self.assertTrue(db.is_admin_user("ADMIN"))
                    self.assertFalse(db.verify_user_password("OFF", "disabled-password"))
                    db.create_user(
                        "NEW",
                        display_name="New User",
                        department="Sales",
                        role="user",
                        enabled=True,
                    )
                    self.assertTrue(db.verify_user_password("NEW", "NEW"))
                    db.change_user_password("NEW", "changed-password")
                    self.assertTrue(db.verify_user_password("NEW", "changed-password"))
                    db.create_user("NEW", role="user", enabled=False)
                    self.assertFalse(db.verify_user_password("NEW", "changed-password"))
            finally:
                close_identity_pool()

            replay = replay_change_log(
                sqlite_path,
                database_url,
                backup_path,
                backup_hash,
                ROLLBACK_CONFIRMATION,
            )
            self.assertGreaterEqual(replay["applied"], 3)
            self.assertEqual(0, replay["remaining"])
            with closing(sqlite3.connect(sqlite_path)) as source:
                source.row_factory = sqlite3.Row
                new_user = source.execute("SELECT * FROM users WHERE employee_id='NEW'").fetchone()
                self.assertIsNotNone(new_user)
                self.assertEqual(0, new_user["enabled"])

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    target.execute("TRUNCATE identity_change_log RESTART IDENTITY")
                    target.execute("TRUNCATE users")
                    target.execute("UPDATE identity_runtime_flags SET value='false'")


if __name__ == "__main__":
    unittest.main()
