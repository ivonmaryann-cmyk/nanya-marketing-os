from __future__ import annotations

import gc
import sqlite3
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

from flask import Flask, session
from openpyxl import Workbook

from fangzheng_web_app import customer_spec_mapping_service as service
from fangzheng_web_app.routes import bp


class CustomerSpecMappingServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "spec-mappings.sqlite3"
        with sqlite3.connect(self.database_path) as connection:
            connection.executescript(
                """
                CREATE TABLE automation_customers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    customer_code TEXT NOT NULL UNIQUE,
                    customer_name TEXT NOT NULL DEFAULT '',
                    customer_short_name TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE automation_customer_spec_mappings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    customer_code TEXT NOT NULL,
                    customer_name TEXT NOT NULL DEFAULT '',
                    product_type TEXT NOT NULL,
                    delimiter TEXT NOT NULL DEFAULT '',
                    glue_system_position INTEGER,
                    thickness_position INTEGER,
                    core_thickness_position INTEGER,
                    dimension_position INTEGER,
                    copper_foil_type_position INTEGER,
                    copper_thickness_position INTEGER,
                    structure_position INTEGER,
                    watermark_position INTEGER,
                    halogen_position INTEGER,
                    rc_position INTEGER,
                    cloth_type_position INTEGER,
                    size_position INTEGER,
                    note TEXT NOT NULL DEFAULT '',
                    enabled INTEGER NOT NULL DEFAULT 1,
                    source_json TEXT NOT NULL DEFAULT '{}',
                    updated_by TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(customer_code, product_type)
                );
                INSERT INTO automation_customers(customer_code,customer_name,customer_short_name)
                VALUES ('100001','已建档客户','已建档');
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

    def _workbook(self, delimiter_header: str = "分隔符") -> Path:
        path = Path(self.temp_dir.name) / f"mapping-{delimiter_header}.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.append([
            "OCC02", "TC_FEM01", "TC_FEM02", delimiter_header,
            "TC_FEM04", "TC_FEM05", "TC_FEM06", "TC_FEM07", "TC_FEM08",
            "TC_FEM09", "TC_FEM10", "TC_FEM11", "TC_FEM12", "TC_FEM13",
            "TC_FEM14", "TC_FEM15", "TC_FEM16", "TC_FEM17",
        ])
        sheet.append(["已建档客户", "100001", 1, "_", 2, 4, None, 8, None, 5, 6, None, None, None, None, None, "FR4_NY2600", None])
        sheet.append(["未建档客户", "999999", 2, None, 1, None, None, None, None, None, None, None, None, 2, 1, 4, "7628 RC47", "N"])
        sheet.append(["错误客户", "888888", 1, None, "bad", None, None, None, None, None, None, None, None, None, None, None, "bad row", None])
        book.save(path)
        book.close()
        return path

    def test_import_is_idempotent_and_keeps_unmatched_customer(self) -> None:
        first = service.import_spec_mapping_workbook(self._workbook(), operated_by="tester")
        second = service.import_spec_mapping_workbook(self._workbook(), operated_by="tester")

        self.assertEqual((2, 0, 1), (first["imported"], first["updated"], first["skipped"]))
        self.assertEqual((0, 2, 1), (second["imported"], second["updated"], second["skipped"]))
        rows = service.list_spec_mappings()
        self.assertEqual(2, len(rows))
        matched = next(row for row in rows if row["customer_code"] == "100001")
        unmatched = next(row for row in rows if row["customer_code"] == "999999")
        self.assertEqual(1, matched["customer_matched"])
        self.assertEqual(1, matched["enabled"])
        self.assertEqual(0, unmatched["customer_matched"])
        self.assertEqual(0, unmatched["enabled"])

    def test_tc_fem03_alias_and_filters_are_supported(self) -> None:
        service.import_spec_mapping_workbook(self._workbook("TC_FEM03"))
        self.assertEqual(1, len(service.list_spec_mappings(product_type="base", match_status="matched")))
        self.assertEqual(1, len(service.list_spec_mappings(status="disabled", match_status="unmatched")))
        self.assertEqual("_", service.list_spec_mappings(keyword="FR4")[0]["delimiter"])

    def test_manual_save_validates_positions_and_duplicate_key(self) -> None:
        values = {
            "customer_code": "200001", "customer_name": "手工客户", "product_type": "base",
            "delimiter": "-", "glue_system_position": "1", "note": "sample", "enabled": "1",
        }
        mapping_id = service.save_spec_mapping(values, operated_by="tester")
        self.assertEqual("200001", service.get_spec_mapping(mapping_id)["customer_code"])
        duplicate_error = ""
        try:
            service.save_spec_mapping(values)
        except ValueError as exc:
            duplicate_error = str(exc)
        self.assertIn("已存在", duplicate_error)
        with self.assertRaisesRegex(ValueError, "必须是整数"):
            service.save_spec_mapping({**values, "customer_code": "200002", "glue_system_position": "1.5"})
        service.set_spec_mapping_enabled(mapping_id, False, operated_by="tester")
        self.assertEqual(0, service.get_spec_mapping(mapping_id)["enabled"])

    def test_customer_spec_match_keeps_source_and_inserts_target_position_gaps(self) -> None:
        service.save_spec_mapping({
            "customer_code": "200001", "customer_name": "匹配客户",
            "product_type": "base", "delimiter": "_",
            "glue_system_position": "2", "thickness_position": "5",
            "copper_thickness_position": "6", "dimension_position": "7",
            "watermark_position": "8", "core_thickness_position": "9",
            "enabled": "1",
        })

        result = service.build_customer_spec_match(
            "200001", "基板", "上海南亚_NY1140_(TG140)_0.8mm_H/H_86X49_无Logo_含铜"
        )

        self.assertEqual(
            "上海南亚_NY1140_(TG140)_*_0.8mm_H/H_86X49_无Logo_含铜",
            result,
        )

    def test_customer_spec_match_handles_whitespace_and_empty_code(self) -> None:
        service.save_spec_mapping({
            "customer_code": "200002", "product_type": "base", "delimiter": "",
            "glue_system_position": "2", "thickness_position": "5", "enabled": "1",
        })

        self.assertEqual(
            "上海南亚 NY1140 (TG140) * 0.8mm H/H 86X49 无Logo 含铜",
            service.build_customer_spec_match(
                "200002", "基板", "上海南亚 NY1140 (TG140) 0.8mm H/H 86X49 无Logo 含铜"
            ),
        )
        self.assertEqual("", service.build_customer_spec_match("", "基板", "NY1140 0.8mm"))
        self.assertEqual("", service.build_customer_spec_match("200002", "PP", "NY1140"))

    def test_pp_directional_dimensions_are_one_size_segment(self) -> None:
        service.save_spec_mapping({
            "customer_code": "104253", "customer_name": "江苏博敏",
            "product_type": "pp", "delimiter": "",
            "glue_system_position": "2", "cloth_type_position": "3",
            "rc_position": "4", "size_position": "5", "enabled": "1",
        })

        result = service.build_customer_spec_match(
            "104253",
            "PP",
            "南亚新材料 NY6300SP 1080 RC=67% 经542 mm 纬620.00 mm 无卤 TG200",
        )

        self.assertEqual(
            "南亚新材料 NY6300SP 1080 RC=67% 经542mm纬620.00mm 无卤 TG200",
            result,
        )

        detail = service.build_customer_spec_match_detail(
            "104253",
            "PP",
            "南亚新材料 NY6300SP 1080 RC=67% 经542 mm 纬620.00 mm 无卤 TG200",
        )
        values = {field["field"]: field["value"] for field in detail["fields"]}
        self.assertEqual("NY6300SP", values["glue_system_position"])
        self.assertEqual("1080", values["cloth_type_position"])
        self.assertEqual("RC=67%", values["rc_position"])
        self.assertEqual("经542mm纬620.00mm", values["size_position"])
        self.assertTrue(detail["mapping_found"])

        mixed_unit_detail = service.build_customer_spec_match_detail(
            "104253",
            "PP",
            "南亚新材料 NY6300SP 1080 RC=69% 经300.00 m 纬49.50 inch 无卤 TG200(DMA)",
        )
        mixed_values = {
            field["field"]: field["value"] for field in mixed_unit_detail["fields"]
        }
        self.assertEqual(
            "南亚新材料 NY6300SP 1080 RC=69% 经300.00m纬49.50inch 无卤 TG200(DMA)",
            mixed_unit_detail["customer_spec_match"],
        )
        self.assertEqual("经300.00m纬49.50inch", mixed_values["size_position"])

        service.save_spec_mapping({
            "customer_code": "104254", "product_type": "pp", "delimiter": "|",
            "glue_system_position": "2", "cloth_type_position": "3",
            "rc_position": "4", "size_position": "5", "enabled": "1",
        })
        delimited_detail = service.build_customer_spec_match_detail(
            "104254",
            "PP",
            "南亚新材料|NY6300SP|1080|RC=69%|经300.00 m 纬49.50 inch|无卤|TG200(DMA)",
        )
        self.assertEqual(
            "南亚新材料|NY6300SP|1080|RC=69%|经300.00m纬49.50inch|无卤|TG200(DMA)",
            delimited_detail["customer_spec_match"],
        )

    def test_compound_core_and_watermark_are_split_into_configured_fields(self) -> None:
        service.save_spec_mapping({
            "customer_code": "123036", "product_type": "base", "delimiter": "",
            "core_thickness_position": "6", "watermark_position": "7", "enabled": "1",
        })

        detail = service.build_customer_spec_match_detail(
            "123036",
            "基板",
            "82.3*49.3 南亚 NY2170H 0.15mm-1/1-TG170 CTI≥175 不含铜无水印 "
            "82.3*49.3 A级 HTE (1080 2张) 有卤素 FR-4 耐CAF",
        )
        values = {field["field"]: field["value"] for field in detail["fields"]}
        labels = {field["header"]: field["label"] for field in detail["fields"]}

        self.assertEqual("不含铜", values["core_thickness_position"])
        self.assertEqual("无水印", values["watermark_position"])
        self.assertIn("不含铜 无水印", detail["customer_spec_match"])
        self.assertEqual("基板尺寸", labels["TC_FEM07"])
        self.assertEqual("PP尺寸", labels["TC_FEM15"])

        positive_detail = service.build_customer_spec_match_detail(
            "123036", "基板", "含铜有水印",
        )
        positive_values = {
            field["field"]: field["value"] for field in positive_detail["fields"]
        }
        self.assertEqual("含铜", positive_values["core_thickness_position"])
        self.assertEqual("有水印", positive_values["watermark_position"])


class CustomerSpecMappingRouteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = Flask(__name__, template_folder=str(Path(__file__).parents[1] / "templates"), static_folder=str(Path(__file__).parents[1] / "static"))
        self.app.config.update(SECRET_KEY="test-secret", TESTING=True)
        self.app.register_blueprint(bp)
        self.client = self.app.test_client()
        with self.client.session_transaction() as user_session:
            user_session["employee_id"] = "employee-a"

    def test_specs_dataset_renders_maintenance_controls(self) -> None:
        row = {
            "id": 7, "customer_code": "100001", "customer_name": "测试客户",
            "product_type": "base", "delimiter": "_", "note": "FR4_NY2600",
            "customer_matched": 1, "enabled": 1, "updated_at": "2026-08-24T10:00:00",
            "master_customer_name": "测试客户", "master_customer_short_name": "测试",
        }
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.routes.is_admin_user", return_value=False
        ), patch("fangzheng_web_app.routes.list_spec_mappings", return_value=[row]):
            response = self.client.get("/customers?dataset=specs")

        self.assertEqual(200, response.status_code)
        html = response.get_data(as_text=True)
        self.assertIn("客户信息表", html)
        self.assertIn("客户规格与厂内规格对照表", html)
        self.assertIn("导入规格对照", html)
        self.assertIn("客户档案待匹配", html)

    def test_logged_in_user_can_save_and_change_status(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.routes.save_spec_mapping", return_value=9
        ) as save, patch("fangzheng_web_app.routes.set_spec_mapping_enabled") as set_enabled:
            save_response = self.client.post(
                "/customers/spec-mappings/save",
                data={"customer_code": "100001", "product_type": "base", "enabled": "1"},
            )
            status_response = self.client.post(
                "/customers/spec-mappings/9/status", data={"enabled": "0"}
            )

        self.assertEqual(302, save_response.status_code)
        self.assertEqual(302, status_response.status_code)
        save.assert_called_once()
        set_enabled.assert_called_once_with(9, False, operated_by="employee-a")

    def test_entry_template_spec_match_endpoint_uses_owned_order_case(self) -> None:
        with patch(
            "fangzheng_web_app.routes.get_order_intake_case",
            return_value={"id": 12, "action_type": "new_order"},
        ), patch(
            "fangzheng_web_app.routes.build_customer_spec_match_detail",
            return_value={
                "customer_spec_match": "NY2150_*_1.6MM",
                "delimiter": "_",
                "segments": ["NY2150", "*", "1.6MM"],
                "fields": [],
            },
        ) as matcher:
            response = self.client.post(
                "/order-automation/cases/12/entry-template/spec-match",
                json={
                    "customer_code": "C001",
                    "customer_spec": "NY2150_TG150_1.6MM",
                    "product_type": "基板",
                },
            )

        self.assertEqual(200, response.status_code)
        self.assertEqual("NY2150_*_1.6MM", response.get_json()["customer_spec_match"])
        self.assertEqual(["NY2150", "*", "1.6MM"], response.get_json()["segments"])
        matcher.assert_called_once_with("C001", "基板", "NY2150_TG150_1.6MM")


if __name__ == "__main__":
    unittest.main()
