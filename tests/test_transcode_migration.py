from __future__ import annotations

import os
import unittest
from unittest.mock import patch

from fangzheng_web_app.database.config import TranscodeDatabaseConfig
from fangzheng_web_app.transcode_migration.schema import MIGRATION_DIR
from fangzheng_web_app.transcode_migration.spec import TABLES


class TranscodeMigrationTests(unittest.TestCase):
    def test_default_config_keeps_sqlite_and_postgresql_is_locked(self) -> None:
        with patch.dict(os.environ, {"TRANSCODE_DATABASE_URL": "postgresql://secret"}, clear=True):
            config = TranscodeDatabaseConfig.from_env()
        self.assertEqual("sqlite", config.backend)
        self.assertNotIn("database_url", config.redacted_summary())
        environment = {
            "TRANSCODE_DATABASE_BACKEND": "postgresql",
            "TRANSCODE_DATABASE_URL": "postgresql://configured-but-locked",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(RuntimeError, "formal switch approval"):
                TranscodeDatabaseConfig.from_env()

    def test_ddl_contains_exactly_the_eighteen_scoped_business_tables(self) -> None:
        sql = (MIGRATION_DIR / "0001_transcode_schema.sql").read_text(encoding="utf-8").lower()
        declared = {
            line.split("(", 1)[0].split()[-1]
            for line in sql.splitlines()
            if line.startswith("create table if not exists ")
        }
        auxiliary = {"transcode_runtime_flags", "transcode_change_log"}
        self.assertEqual(set(TABLES), declared - auxiliary)
        self.assertEqual(auxiliary, declared & auxiliary)
        self.assertNotIn("create table if not exists settings", sql)


if __name__ == "__main__":
    unittest.main()
