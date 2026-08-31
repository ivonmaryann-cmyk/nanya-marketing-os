from __future__ import annotations

import gc
import sqlite3
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from fangzheng_web_app import adhesive_code_service as service


class AdhesiveCodeMasterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "adhesives.sqlite3"
        with sqlite3.connect(self.database_path) as connection:
            connection.executescript(
                """
                CREATE TABLE automation_adhesive_codes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    adhesive_code TEXT NOT NULL UNIQUE,
                    adhesive_name TEXT NOT NULL DEFAULT '',
                    usage_category TEXT NOT NULL DEFAULT '',
                    finance_category TEXT NOT NULL DEFAULT '',
                    legacy_adhesive_code TEXT NOT NULL DEFAULT '',
                    enabled INTEGER NOT NULL DEFAULT 1,
                    source_json TEXT NOT NULL DEFAULT '{}',
                    updated_by TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                """
            )

        @contextmanager
        def cursor():
            connection = sqlite3.connect(self.database_path)
            connection.row_factory = sqlite3.Row
            try:
                yield connection
                connection.commit()
            except Exception:
                connection.rollback()
                raise
            finally:
                connection.close()

        self.cursor_patch = patch.object(service, "db_cursor", cursor)
        self.cursor_patch.start()

    def tearDown(self) -> None:
        self.cursor_patch.stop()
        gc.collect()
        self.temp_dir.cleanup()

    def _workbook(self) -> Path:
        path = Path(self.temp_dir.name) / "adhesive-master.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.append(["胶系编号", "胶系名称", "用途类别", "财务类别", "旧胶系编号"])
        sheet.append(["1ann", "NY1140 N料", "其他", "普通", "1A"])
        sheet.append(["6CNL", "NY6300SP", "高速", "无卤", ""])
        sheet.append(["", "错误行", "其他", "普通", ""])
        book.save(path)
        book.close()
        return path

    def test_import_is_idempotent_and_supports_filters(self) -> None:
        first = service.import_adhesive_code_workbook(self._workbook(), operated_by="tester")
        second = service.import_adhesive_code_workbook(self._workbook(), operated_by="tester")

        self.assertEqual((2, 0, 1), (first["imported"], first["updated"], first["skipped"]))
        self.assertEqual((0, 2, 1), (second["imported"], second["updated"], second["skipped"]))
        self.assertEqual("1ANN", service.list_adhesive_codes(keyword="NY1140")[0]["adhesive_code"])
        self.assertEqual("6CNL", service.list_adhesive_codes(usage_category="高速")[0]["adhesive_code"])
        self.assertEqual(["其他", "高速"], service.adhesive_code_filter_values()["usage_categories"])

    def test_manual_save_validates_duplicate_and_status(self) -> None:
        record_id = service.save_adhesive_code({
            "adhesive_code": "6CNL", "adhesive_name": "NY6300SP", "usage_category": "高速",
            "finance_category": "普通", "legacy_adhesive_code": "6CN", "enabled": "1",
        }, operated_by="tester")
        self.assertEqual("6CNL", service.get_adhesive_code(record_id)["adhesive_code"])
        with self.assertRaisesRegex(ValueError, "已存在"):
            service.save_adhesive_code({"adhesive_code": "6cnl", "enabled": "1"})
        service.set_adhesive_code_enabled(record_id, False, operated_by="tester")
        self.assertEqual(0, service.get_adhesive_code(record_id)["enabled"])

    def test_candidate_lookup_returns_enabled_matches_in_code_order(self) -> None:
        service.save_adhesive_code({"adhesive_code": "6CNZ", "adhesive_name": "NY6300SP 专用", "enabled": "1"})
        service.save_adhesive_code({"adhesive_code": "6CNA", "adhesive_name": "NY6300SP 通用", "enabled": "1"})
        disabled_id = service.save_adhesive_code({"adhesive_code": "6CNN", "adhesive_name": "NY6300SP 停用", "enabled": "1"})
        service.set_adhesive_code_enabled(disabled_id, False)

        self.assertEqual(
            ["6CNA", "6CNZ"],
            [item["adhesive_code"] for item in service.find_adhesive_code_candidates("NY6300SP")],
        )

    def test_postgresql_migration_defines_adhesive_master(self) -> None:
        sql = (Path(__file__).resolve().parents[1] / "migrations" / "automation" / "postgresql" / "0010_adhesive_code_master.sql").read_text(encoding="utf-8")
        self.assertIn("CREATE TABLE IF NOT EXISTS automation_adhesive_codes", sql)
        self.assertIn("adhesive_code TEXT NOT NULL UNIQUE", sql)
