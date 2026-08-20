from __future__ import annotations

import os
import io
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app.automation_migration.copy import _upsert_sql
from fangzheng_web_app.automation_migration.schema import MIGRATION_DIR
from fangzheng_web_app.automation_migration.snapshot import sqlite_snapshot
from fangzheng_web_app.automation_migration.spec import TABLES
from fangzheng_web_app.database.config import AutomationDatabaseConfig
from fangzheng_web_app.database.sql import qmark_to_pyformat, sqlite_to_postgresql
from fangzheng_web_app.file_storage.local import FallbackFileStorage


class AutomationMigrationInfrastructureTests(unittest.TestCase):
    def test_default_config_keeps_sqlite_and_redacts_url(self) -> None:
        with patch.dict(os.environ, {"AUTOMATION_DATABASE_URL": "postgresql://secret"}, clear=True):
            config = AutomationDatabaseConfig.from_env()
        self.assertEqual("sqlite", config.backend)
        self.assertNotIn("database_url", config.redacted_summary())
        self.assertTrue(config.redacted_summary()["database_url_configured"])

    def test_postgresql_requires_url(self) -> None:
        with patch.dict(os.environ, {"AUTOMATION_DATABASE_BACKEND": "postgresql"}, clear=True):
            with self.assertRaisesRegex(ValueError, "AUTOMATION_DATABASE_URL"):
                AutomationDatabaseConfig.from_env()

    def test_postgresql_application_access_requires_formal_switch_flag(self) -> None:
        environment = {
            "AUTOMATION_DATABASE_BACKEND": "postgresql",
            "AUTOMATION_DATABASE_URL": "postgresql://configured-but-not-opened",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(RuntimeError, "formal switch approval"):
                AutomationDatabaseConfig.from_env()

    def test_qmark_conversion_ignores_literals_and_comments(self) -> None:
        sql = "SELECT '?' AS literal, value FROM sample WHERE id=? -- ?\nAND note='it''s ?' /* ? */"
        converted = qmark_to_pyformat(sql)
        self.assertEqual(1, converted.count("%s"))
        self.assertIn("id=%s", converted)
        self.assertIn("'?'", converted)

    def test_audited_sqlite_dialect_is_translated_explicitly(self) -> None:
        group_sql, _ = sqlite_to_postgresql("SELECT GROUP_CONCAT(a.filename, ' ') FROM mail_attachments a WHERE a.mail_id=?")
        self.assertIn("STRING_AGG(a.filename, ' ')", group_sql)
        self.assertIn("a.mail_id=%s", group_sql)
        replace_sql, returns_id = sqlite_to_postgresql(
            "INSERT OR REPLACE INTO mail_fetch_task_messages(fetch_task_id,mail_id,is_new,created_at) VALUES (?,?,?,?)"
        )
        self.assertIn("ON CONFLICT (fetch_task_id, mail_id) DO UPDATE", replace_sql)
        self.assertFalse(returns_id)
        metadata_sql, _ = sqlite_to_postgresql("SELECT 1 FROM settings WHERE key=?")
        self.assertIn("automation_metadata", metadata_sql)

    def test_identity_insert_adds_returning_id(self) -> None:
        sql, returns_id = sqlite_to_postgresql("INSERT INTO mail_accounts(email,created_at,updated_at) VALUES (?,?,?)")
        self.assertTrue(returns_id)
        self.assertTrue(sql.endswith("RETURNING id"))

    def test_ddl_contains_exactly_the_twenty_scoped_business_tables(self) -> None:
        sql = (MIGRATION_DIR / "0001_automation_schema.sql").read_text(encoding="utf-8").lower()
        declared = {
            line.split("(", 1)[0].split()[-1]
            for line in sql.splitlines()
            if line.startswith("create table if not exists ")
        }
        auxiliary = {"automation_schema_migrations", "automation_metadata"}
        self.assertEqual(set(TABLES), declared - auxiliary)
        self.assertEqual(auxiliary, declared & auxiliary)

    def test_shadow_migration_adds_only_auxiliary_tables(self) -> None:
        sql = (MIGRATION_DIR / "0002_shadow_sync.sql").read_text(encoding="utf-8").lower()
        declared = {
            line.split("(", 1)[0].split()[-1]
            for line in sql.splitlines()
            if line.startswith("create table if not exists ")
        }
        self.assertEqual({"automation_migration_inbox", "automation_shadow_runs"}, declared)

    def test_generated_copy_statement_is_idempotent(self) -> None:
        sql = _upsert_sql("mail_accounts", ["id", "email", "updated_at"])
        self.assertIn('ON CONFLICT ("id") DO UPDATE', sql)
        self.assertNotIn('"id"=EXCLUDED."id"', sql)

    def test_snapshot_is_independent_and_source_stays_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.db"
            with closing(sqlite3.connect(source)) as connection:
                connection.execute("CREATE TABLE sample(id INTEGER PRIMARY KEY, value TEXT)")
                connection.execute("INSERT INTO sample(value) VALUES ('original')")
                connection.commit()
            with sqlite_snapshot(source) as snapshot:
                with closing(sqlite3.connect(snapshot)) as connection:
                    connection.execute("UPDATE sample SET value='snapshot'")
                    connection.commit()
                with closing(sqlite3.connect(source)) as connection:
                    value = connection.execute("SELECT value FROM sample").fetchone()[0]
            self.assertEqual("original", value)

    def test_file_storage_prefers_object_then_falls_back_to_legacy(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "objects"
            root.mkdir()
            legacy = Path(directory) / "legacy.txt"
            legacy.write_text("legacy", encoding="utf-8")
            storage = FallbackFileStorage(root)
            self.assertEqual(legacy, storage.resolve(object_key="missing.txt", legacy_path=str(legacy)))
            managed = root / "mail" / "attachment.txt"
            managed.parent.mkdir()
            managed.write_text("managed", encoding="utf-8")
            self.assertEqual(managed.resolve(), storage.resolve(object_key="mail/attachment.txt", legacy_path=str(legacy)))

    def test_local_storage_save_is_atomic_and_object_keys_cannot_escape(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            storage = FallbackFileStorage(Path(directory))
            saved = storage.save(io.BytesIO(b"content"), "mail/file.bin")
            self.assertEqual(b"content", saved.read_bytes())
            self.assertEqual(64, len(storage.checksum(object_key="mail/file.bin")))
            self.assertIsNone(storage.temporary_download_url("mail/file.bin"))
            with self.assertRaises(ValueError):
                storage.save(io.BytesIO(b"bad"), "../escape.bin")


if __name__ == "__main__":
    unittest.main()
