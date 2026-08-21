from __future__ import annotations

import os
import sqlite3
import unittest
from unittest.mock import patch

from fangzheng_web_app.configuration_migration.rollback import apply_changes_to_sqlite
from fangzheng_web_app.configuration_migration.schema import MIGRATION_DIR
from fangzheng_web_app.configuration_migration.spec import TABLE_COLUMNS
from fangzheng_web_app.database.config import ConfigurationDatabaseConfig


class ConfigurationMigrationTests(unittest.TestCase):
    def test_default_config_keeps_sqlite_and_redacts_url(self) -> None:
        with patch.dict(os.environ, {"CONFIG_DATABASE_URL": "postgresql://secret"}, clear=True):
            config = ConfigurationDatabaseConfig.from_env()
        self.assertEqual("sqlite", config.backend)
        self.assertNotIn("database_url", config.redacted_summary())
        self.assertTrue(config.redacted_summary()["database_url_configured"])

    def test_postgresql_requires_url_and_explicit_switch(self) -> None:
        with patch.dict(os.environ, {"CONFIG_DATABASE_BACKEND": "postgresql"}, clear=True):
            with self.assertRaisesRegex(ValueError, "CONFIG_DATABASE_URL"):
                ConfigurationDatabaseConfig.from_env()
        environment = {
            "CONFIG_DATABASE_BACKEND": "postgresql",
            "CONFIG_DATABASE_URL": "postgresql://configured-but-locked",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(RuntimeError, "formal switch approval"):
                ConfigurationDatabaseConfig.from_env()

    def test_ddl_is_limited_to_configuration_scope(self) -> None:
        sql = (MIGRATION_DIR / "0001_configuration_schema.sql").read_text(encoding="utf-8").lower()
        declared = {
            line.split("(", 1)[0].split()[-1]
            for line in sql.splitlines()
            if line.startswith("create table if not exists ")
        }
        self.assertEqual(
            {
                "settings",
                "pdf_excel_ai_config_versions",
                "configuration_runtime_flags",
                "configuration_change_log",
            },
            declared,
        )

    def test_change_replay_is_idempotent_and_rejects_out_of_scope_table(self) -> None:
        connection = sqlite3.connect(":memory:")
        try:
            connection.execute("CREATE TABLE settings(key TEXT PRIMARY KEY,value TEXT NOT NULL)")
            columns = TABLE_COLUMNS["pdf_excel_ai_config_versions"]
            definitions = ",".join(
                "id INTEGER PRIMARY KEY" if column == "id" else f'"{column}" TEXT'
                for column in columns
            )
            connection.execute(f"CREATE TABLE pdf_excel_ai_config_versions({definitions})")
            change = {
                "table_name": "settings",
                "operation": "insert",
                "pk_json": {"key": "example"},
                "row_json": {"key": "example", "value": "one"},
            }
            self.assertEqual(1, apply_changes_to_sqlite(connection, [change]))
            change["row_json"]["value"] = "two"
            self.assertEqual(1, apply_changes_to_sqlite(connection, [change]))
            self.assertEqual("two", connection.execute(
                "SELECT value FROM settings WHERE key='example'"
            ).fetchone()[0])
            with self.assertRaisesRegex(ValueError, "outside the migration scope"):
                apply_changes_to_sqlite(connection, [{**change, "table_name": "users"}])
        finally:
            connection.close()


if __name__ == "__main__":
    unittest.main()
