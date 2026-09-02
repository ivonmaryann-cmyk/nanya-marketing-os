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

from fangzheng_web_app import db
from fangzheng_web_app.database.transcode import close_transcode_pool
from fangzheng_web_app.pp_transcode_rules import (
    ensure_pp_transcode_tables,
    get_base_rule,
    list_pp_confirmation_items,
    replace_pp_confirmation_items,
    save_base_rule,
    seed_pp_transcode_rules,
)
from fangzheng_web_app.transcode_customer_rule_admin import ensure_customer_rule_maintenance_tables
from fangzheng_web_app.transcode_migration.copy import copy_snapshot
from fangzheng_web_app.transcode_migration.rollback import (
    ENABLE_CONFIRMATION,
    ROLLBACK_CONFIRMATION,
    enable_change_capture,
    file_sha256,
    replay_change_log,
)
from fangzheng_web_app.transcode_migration.schema import apply_migrations
from fangzheng_web_app.transcode_migration.spec import TABLES
from fangzheng_web_app.transcode_migration.verify import verify_snapshot
from fangzheng_web_app.transcode_rule_center import (
    create_backup,
    ensure_rule_center_tables,
    restore_backup,
    save_lookup_override,
)


@unittest.skipUnless(os.getenv("TRANSCODE_TEST_DATABASE_URL"), "TRANSCODE_TEST_DATABASE_URL is not configured")
class TranscodePostgresqlIntegrationTests(unittest.TestCase):
    def test_copy_runtime_writes_and_rollback_replay(self) -> None:
        database_url = os.environ["TRANSCODE_TEST_DATABASE_URL"]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sqlite_path = root / "source.db"
            backup_path = root / "source.backup.db"
            with patch.object(db, "DATABASE_PATH", sqlite_path):
                db.init_db()
                ensure_pp_transcode_tables()
                ensure_customer_rule_maintenance_tables()
                with patch("fangzheng_web_app.transcode_rule_center.BACKUP_DIR", root / "rule-backups"):
                    ensure_rule_center_tables()
            with closing(sqlite3.connect(sqlite_path)) as source:
                source.execute(
                    "INSERT INTO users(employee_id,password_hash,role,enabled,must_change_password,created_at,updated_at) "
                    "VALUES (?,?,?,?,?,?,?)",
                    ("E001", "test-hash", "user", 1, 1, "2026-08-21", "2026-08-21"),
                )
                source.execute(
                    "INSERT INTO jobs(feature,employee_id,source_filename,stored_input_path,status,rule_version,created_at) "
                    "VALUES (?,?,?,?,?,?,?)",
                    ("pp_transcode_agent", "E001", "source.xlsx", "input.xlsx", "completed", "v1", "2026-08-21"),
                )
                source.commit()
            shutil.copy2(sqlite_path, backup_path)
            backup_hash = file_sha256(backup_path)

            with psycopg.connect(database_url, row_factory=dict_row) as target:
                with target.transaction():
                    apply_migrations(target)
                    target.execute('TRUNCATE TABLE ' + ','.join(f'"{table}"' for table in TABLES) + ' RESTART IDENTITY CASCADE')
                    target.execute("TRUNCATE transcode_change_log RESTART IDENTITY")
                    target.execute("UPDATE transcode_runtime_flags SET value='false'")
                    first = copy_snapshot(sqlite_path, target)
                    second = copy_snapshot(sqlite_path, target)
                    self.assertEqual(first, second)
                    self.assertTrue(verify_snapshot(sqlite_path, target)["ok"])
                    trigger_count = target.execute(
                        "SELECT COUNT(*) AS total FROM pg_trigger WHERE tgname='transcode_change_capture' AND NOT tgisinternal"
                    ).fetchone()["total"]
                    self.assertEqual(18, trigger_count)

            enable_change_capture(database_url, backup_path, backup_hash, ENABLE_CONFIRMATION)
            environment = {
                "TRANSCODE_DATABASE_BACKEND": "postgresql",
                "TRANSCODE_POSTGRESQL_READ_WRITE_ENABLED": "true",
                "TRANSCODE_DATABASE_URL": database_url,
            }
            close_transcode_pool()
            try:
                with patch.dict(os.environ, environment, clear=False), patch.object(db, "DATABASE_PATH", sqlite_path):
                    job_id = db.create_job("E001", "new.xlsx", "new-input.xlsx", "v2", "pp_transcode_agent")
                    db.update_job_status(job_id, status="completed", success_count=1, completed=True)
                    self.assertEqual("completed", db.get_job(job_id)["status"])
                    self.assertTrue(any(row["id"] == job_id for row in db.list_jobs("E001")))
                    self.assertTrue(db.list_jobs("E001", start_date="2020-01-01", end_date="2030-01-01"))
                    self.assertTrue(any(row["id"] == job_id for row in db.list_expired_terminal_jobs("2030-01-01")))
                    rule_id = save_base_rule(
                        {"field_key": "glass_style", "input_value": "TEST-STYLE", "output_value": "1080"},
                        "E001",
                    )
                    self.assertEqual("1080", get_base_rule(rule_id)["output_value"])
                    seed_pp_transcode_rules()
                    replace_pp_confirmation_items(
                        job_id,
                        "E001",
                        [{
                            "excel_row": 2, "customer_code": "C001", "customer_name": "Customer",
                            "spec": "1080", "order_remark": "", "pending_code": "TEST",
                            "confidence": 80, "summary": "review", "field_evidence": [],
                        }],
                    )
                    self.assertEqual(1, len(list_pp_confirmation_items(job_id, "E001")))
                    confirmation_ids = db.replace_transcode_agent_confirmation_items(
                        job_id,
                        "E001",
                        [{
                            "excel_row": 2, "field_key": "glue_code", "field_label": "Glue",
                            "pending_code": "T1", "score": 80,
                        }],
                    )
                    self.assertEqual(1, len(confirmation_ids))
                    self.assertEqual(1, db.transcode_agent_confirmation_counts(job_id)["pending"])
                    db.save_transcode_model_config(
                        "E001", enabled=True, base_url="https://example.invalid", api_key="test-key",
                        model="test-model", timeout_seconds=30, max_order_calls=2,
                    )
                    self.assertEqual("test-model", db.get_transcode_model_config("E001")["model"])
                    with patch("fangzheng_web_app.transcode_rule_center.BACKUP_DIR", root / "rule-backups"):
                        save_lookup_override(
                            {"lookup_group": "glue_code", "lookup_input": "TEST", "lookup_output": "T1"},
                            updated_by="E001",
                        )
                        rule_backup = create_backup(reason="integration-test")
                        restore_backup(rule_backup.name, updated_by="E001")
            finally:
                close_transcode_pool()

            applied = 0
            while True:
                replay = replay_change_log(
                    sqlite_path, database_url, backup_path, backup_hash,
                    ROLLBACK_CONFIRMATION, batch_size=100,
                )
                applied += replay["applied"]
                if replay["remaining"] == 0:
                    break
            self.assertGreater(applied, 0)
            with closing(sqlite3.connect(sqlite_path)) as source:
                source.row_factory = sqlite3.Row
                self.assertEqual("completed", source.execute("SELECT status FROM jobs WHERE id=?", (job_id,)).fetchone()["status"])
                self.assertEqual("1080", source.execute("SELECT output_value FROM pp_transcode_base_rules WHERE id=?", (rule_id,)).fetchone()["output_value"])
                self.assertEqual("test-model", source.execute("SELECT model FROM transcode_model_configs WHERE employee_id='E001'").fetchone()["model"])
                self.assertEqual("T1", source.execute("SELECT output_value FROM transcode_rule_center_lookup_overrides WHERE group_key='glue_code' AND input_value='TEST'").fetchone()["output_value"])


if __name__ == "__main__":
    unittest.main()
