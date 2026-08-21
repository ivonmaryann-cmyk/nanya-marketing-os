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
from cryptography.fernet import Fernet
from psycopg.rows import dict_row
from werkzeug.security import generate_password_hash

from fangzheng_web_app import db
from fangzheng_web_app.ai_repair_config import (
    AiRepairConfig,
    get_ai_repair_config,
    list_ai_config_versions,
    save_ai_config_version,
)
from fangzheng_web_app.configuration_migration.copy import copy_snapshot
from fangzheng_web_app.configuration_migration.rollback import (
    ENABLE_CONFIRMATION,
    ROLLBACK_CONFIRMATION,
    enable_change_capture,
    file_sha256,
    replay_change_log,
)
from fangzheng_web_app.configuration_migration.schema import apply_migrations
from fangzheng_web_app.configuration_migration.verify import verify_snapshot
from fangzheng_web_app.database.configuration import close_configuration_pool


@unittest.skipUnless(os.getenv("CONFIG_TEST_DATABASE_URL"), "CONFIG_TEST_DATABASE_URL is not configured")
class ConfigurationPostgresqlIntegrationTests(unittest.TestCase):
    def test_copy_runtime_access_ai_ciphertext_and_rollback_replay(self) -> None:
        database_url = os.environ["CONFIG_TEST_DATABASE_URL"]
        with tempfile.TemporaryDirectory() as directory:
            sqlite_path = Path(directory) / "source.db"
            backup_path = Path(directory) / "source.backup.db"
            with patch.object(db, "DATABASE_PATH", sqlite_path), patch.dict(
                os.environ,
                {
                    "CONFIG_DATABASE_BACKEND": "sqlite",
                    "CONFIG_POSTGRESQL_READ_WRITE_ENABLED": "false",
                },
                clear=False,
            ):
                db.init_db()
            admin_hash = generate_password_hash("admin-password")
            with closing(sqlite3.connect(sqlite_path)) as source:
                source.execute(
                    "UPDATE settings SET value=? WHERE key='admin_password_hash'",
                    (admin_hash,),
                )
                source.commit()
            shutil.copy2(sqlite_path, backup_path)
            backup_hash = file_sha256(backup_path)

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    apply_migrations(target)
                    target.execute("UPDATE configuration_runtime_flags SET value='false'")
                    target.execute("TRUNCATE configuration_change_log RESTART IDENTITY")
                    target.execute("TRUNCATE pdf_excel_ai_config_versions RESTART IDENTITY CASCADE")
                    target.execute("TRUNCATE settings")
                    first_copy = copy_snapshot(sqlite_path, target)
                    second_copy = copy_snapshot(sqlite_path, target)
                    self.assertEqual(first_copy, second_copy)
                    verification = verify_snapshot(sqlite_path, target)
                    self.assertTrue(verification["ok"])
                    self.assertTrue(verification["admin_password_hash_match"])
                    self.assertTrue(verification["ai_ciphertexts_match"])

            enable_change_capture(database_url, backup_path, backup_hash, ENABLE_CONFIRMATION)
            master_key = Fernet.generate_key().decode("ascii")
            environment = {
                "CONFIG_DATABASE_BACKEND": "postgresql",
                "CONFIG_POSTGRESQL_READ_WRITE_ENABLED": "true",
                "CONFIG_DATABASE_URL": database_url,
                "PDF_EXCEL_AI_CONFIG_MASTER_KEY": master_key,
            }
            close_configuration_pool()
            try:
                with patch.dict(os.environ, environment, clear=False):
                    self.assertEqual(admin_hash, db.get_setting("admin_password_hash"))
                    db.set_setting("stage4_runtime_probe", "passed")
                    self.assertEqual("passed", db.get_setting("stage4_runtime_probe"))
                    config = AiRepairConfig(
                        enabled=False,
                        api_key="test-api-key",
                        base_url="https://api.example.test",
                        model="test-model",
                        timeout_seconds=30,
                        max_rows=5,
                    )
                    saved = save_ai_config_version(
                        config,
                        employee_id="ADMIN",
                        expected_active_version_id=None,
                        test_status="not_tested",
                        test_message="",
                    )
                    self.assertEqual("test-api-key", saved.api_key)
                    self.assertEqual(saved.version_id, get_ai_repair_config(strict=True).version_id)
                    self.assertEqual(1, len(list_ai_config_versions()))
            finally:
                close_configuration_pool()

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
                self.assertEqual(
                    "passed",
                    source.execute(
                        "SELECT value FROM settings WHERE key='stage4_runtime_probe'"
                    ).fetchone()["value"],
                )
                ai_row = source.execute(
                    "SELECT api_key_ciphertext FROM pdf_excel_ai_config_versions"
                ).fetchone()
                self.assertTrue(ai_row["api_key_ciphertext"])

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    target.execute("TRUNCATE configuration_change_log RESTART IDENTITY")
                    target.execute("TRUNCATE pdf_excel_ai_config_versions RESTART IDENTITY CASCADE")
                    target.execute("TRUNCATE settings")
                    target.execute("UPDATE configuration_runtime_flags SET value='false'")


if __name__ == "__main__":
    unittest.main()
