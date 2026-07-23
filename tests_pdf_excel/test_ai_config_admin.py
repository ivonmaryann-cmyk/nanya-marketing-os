from __future__ import annotations

import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from cryptography.fernet import Fernet
from flask import Flask
from werkzeug.datastructures import FileStorage

from fangzheng_web_app import db, routes
from fangzheng_web_app.ai_repair_config import (
    AiRepairConfig,
    AiConfigConflictError,
    config_from_manifest_snapshot,
    get_active_ai_config_version_id,
    get_ai_repair_config,
    save_ai_config_version,
    validate_ai_config_input,
)
from fangzheng_web_app.pdf_excel_service import queue_pdf_excel_job
from fangzheng_web_app.deepseek_repair_client import DeepSeekRepairError
from fangzheng_web_app.purchase_ai_repair import _body_missing_payload, _build_repair_payload
from fangzheng_web_app.purchase_factory_mapper import _request_ai_header_mapping


class PdfExcelAiConfigAdminTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DATABASE_PATH
        db.DATABASE_PATH = Path(self.temp_dir.name) / "storage" / "app.db"
        db.DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
        self.env = patch.dict(
            os.environ,
            {
                "PDF_EXCEL_AI_CONFIG_MASTER_KEY": Fernet.generate_key().decode("ascii"),
                "PDF_EXCEL_AI_REPAIR_ENABLED": "1",
                "DEEPSEEK_API_KEY": "legacy-secret-key",
                "DEEPSEEK_BASE_URL": "https://api.deepseek.com",
                "DEEPSEEK_MODEL": "legacy-model",
                "PDF_EXCEL_AI_REPAIR_TIMEOUT_SECONDS": "45",
                "PDF_EXCEL_AI_REPAIR_MAX_ROWS": "12",
            },
        )
        self.env.start()
        db.init_db()
        db.create_user("admin1", role="admin")
        db.change_user_password("admin1", "admin-pass")
        db.create_user("user1", role="user")
        db.change_user_password("user1", "user-pass")

        self.app = Flask(
            __name__,
            template_folder=str(Path(__file__).resolve().parents[1] / "templates"),
            static_folder=str(Path(__file__).resolve().parents[1] / "static"),
        )
        self.app.config.update(TESTING=True, SECRET_KEY="test-secret")
        self.app.register_blueprint(routes.bp)
        self.client = self.app.test_client()

    def tearDown(self) -> None:
        self.env.stop()
        db.DATABASE_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def _login_session(self, employee_id: str) -> None:
        with self.client.session_transaction() as session:
            session["employee_id"] = employee_id

    def _csrf_token(self) -> str:
        self.client.get("/features/pdf-excel?ai_config=1")
        with self.client.session_transaction() as session:
            return str(session["pdf_ai_csrf_token"])

    def _form(self, *, token: str, expected: str = "", model: str = "model-v1", api_key: str = "web-secret-key") -> dict[str, str]:
        return {
            "csrf_token": token,
            "expected_active_version_id": expected,
            "action": "save",
            "enabled": "1",
            "api_key": api_key,
            "base_url": "https://api.deepseek.com",
            "model": model,
            "timeout_seconds": "60",
            "max_rows": "20",
            "repair_instruction": "字段修复业务指令",
            "rebuild_instruction": "正文重建业务指令",
            "header_mapping_instruction": "表头映射业务指令",
            "current_password": "admin-pass",
        }

    def test_non_admin_cannot_see_or_access_ai_config(self) -> None:
        self._login_session("user1")
        response = self.client.get("/features/pdf-excel")
        self.assertEqual(response.status_code, 200)
        self.assertNotIn("/admin/pdf-excel-ai", response.get_data(as_text=True))
        self.assertEqual(self.client.get("/admin/pdf-excel-ai").status_code, 403)
        self.assertEqual(self.client.post("/admin/pdf-excel-ai").status_code, 403)

    def test_admin_opens_ai_config_as_pdf_page_dialog(self) -> None:
        self._login_session("admin1")
        response = self.client.get("/features/pdf-excel")
        page = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn('id="open-pdf-ai-config"', page)
        self.assertIn('id="pdf-ai-config-dialog"', page)
        self.assertIn('data-auto-open="false"', page)

        popup_page = self.client.get("/features/pdf-excel?ai_config=1").get_data(as_text=True)
        self.assertIn('data-auto-open="true"', popup_page)
        legacy_response = self.client.get("/admin/pdf-excel-ai")
        self.assertEqual(legacy_response.status_code, 302)
        self.assertIn("/features/pdf-excel?ai_config=1", legacy_response.headers["Location"])

    def test_admin_save_encrypts_key_and_never_renders_secret(self) -> None:
        self._login_session("admin1")
        token = self._csrf_token()
        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            response = self.client.post("/admin/pdf-excel-ai", data=self._form(token=token))
        self.assertEqual(response.status_code, 302)
        version_id = get_active_ai_config_version_id()
        self.assertIsNotNone(version_id)
        loaded = get_ai_repair_config(version_id, strict=True)
        self.assertEqual(loaded.api_key, "web-secret-key")
        self.assertEqual(loaded.model, "model-v1")

        with db.db_cursor() as conn:
            row = conn.execute(
                "SELECT api_key_ciphertext FROM pdf_excel_ai_config_versions WHERE id = ?",
                (version_id,),
            ).fetchone()
        self.assertNotIn("web-secret-key", row["api_key_ciphertext"])
        page = self.client.get("/features/pdf-excel?ai_config=1").get_data(as_text=True)
        self.assertNotIn("web-secret-key", page)
        self.assertNotIn("legacy-secret-key", page)
        self.assertIn(f"PDF/Excel AI 配置 v{version_id} 已启用。", page)

    def test_bad_password_csrf_and_stale_version_do_not_change_active_config(self) -> None:
        self._login_session("admin1")
        token = self._csrf_token()
        form = self._form(token=token)
        form["current_password"] = "wrong"
        response = self.client.post("/admin/pdf-excel-ai", data=form)
        self.assertEqual(response.status_code, 200)
        self.assertIsNone(get_active_ai_config_version_id())

        bad_csrf = self._form(token="invalid")
        self.assertEqual(self.client.post("/admin/pdf-excel-ai", data=bad_csrf).status_code, 400)
        self.assertIsNone(get_active_ai_config_version_id())

        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            self.client.post("/admin/pdf-excel-ai", data=self._form(token=token))
        active_id = get_active_ai_config_version_id()
        stale = self._form(token=token, expected="", model="stale-model")
        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            response = self.client.post("/admin/pdf-excel-ai", data=stale)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(get_active_ai_config_version_id(), active_id)

    def test_version_conflict_is_enforced_in_storage_transaction(self) -> None:
        legacy = get_ai_repair_config()
        candidate = validate_ai_config_input(
            {
                "enabled": True,
                "model": "model-v1",
                "base_url": "https://api.deepseek.com",
                "timeout_seconds": "30",
                "max_rows": "10",
                "api_key": "first-key",
            },
            legacy,
        )
        first = save_ai_config_version(
            candidate,
            employee_id="admin1",
            expected_active_version_id=None,
            test_status="passed",
            test_message="ok",
        )
        self.assertEqual(first.version_id, 1)
        with self.assertRaises(AiConfigConflictError):
            save_ai_config_version(
                candidate,
                employee_id="admin1",
                expected_active_version_id=None,
                test_status="passed",
                test_message="ok",
            )

    def test_job_manifest_locks_ai_version_without_storing_key(self) -> None:
        first_candidate = validate_ai_config_input(
            {
                "enabled": True,
                "model": "model-v1",
                "base_url": "https://api.deepseek.com",
                "timeout_seconds": "30",
                "max_rows": "10",
                "api_key": "version-one-secret",
            },
            get_ai_repair_config(),
        )
        first = save_ai_config_version(
            first_candidate,
            employee_id="admin1",
            expected_active_version_id=None,
            test_status="passed",
            test_message="ok",
        )
        jobs_root = Path(self.temp_dir.name) / "jobs"
        upload = FileStorage(stream=io.BytesIO(b"fake-pdf"), filename="order.pdf")
        with patch("fangzheng_web_app.pdf_excel_service.JOBS_DIR", jobs_root), patch(
            "fangzheng_web_app.pdf_excel_service.launch_job_process"
        ):
            job_id = queue_pdf_excel_job("admin1", [upload])
        job = db.get_job(job_id)
        manifest_text = Path(job["stored_input_path"]).read_text(encoding="utf-8")
        manifest = json.loads(manifest_text)
        self.assertEqual(manifest["ai_config"]["version_id"], first.version_id)
        self.assertNotIn("version-one-secret", manifest_text)

        second_candidate = validate_ai_config_input(
            {
                "enabled": True,
                "model": "model-v2",
                "base_url": "https://api.deepseek.com",
                "timeout_seconds": "30",
                "max_rows": "10",
                "api_key": "version-two-secret",
            },
            first,
        )
        second = save_ai_config_version(
            second_candidate,
            employee_id="admin1",
            expected_active_version_id=first.version_id,
            test_status="passed",
            test_message="ok",
        )
        self.assertNotEqual(second.version_id, first.version_id)
        locked = config_from_manifest_snapshot(manifest["ai_config"])
        self.assertEqual(locked.version_id, first.version_id)
        self.assertEqual(locked.model, "model-v1")

    def test_connection_failure_keeps_current_version_active(self) -> None:
        self._login_session("admin1")
        token = self._csrf_token()
        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            self.client.post("/admin/pdf-excel-ai", data=self._form(token=token))
        active_id = get_active_ai_config_version_id()
        failed_form = self._form(token=token, expected=str(active_id), model="bad-model")
        with patch.object(routes, "test_repair_connection", side_effect=DeepSeekRepairError("认证失败")):
            response = self.client.post("/admin/pdf-excel-ai", data=failed_form)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(get_active_ai_config_version_id(), active_id)

    def test_missing_master_key_blocks_persistence(self) -> None:
        self._login_session("admin1")
        token = self._csrf_token()
        with patch.dict(os.environ, {"PDF_EXCEL_AI_CONFIG_MASTER_KEY": ""}), patch.object(
            routes, "test_repair_connection", return_value="连接测试通过"
        ):
            response = self.client.post("/admin/pdf-excel-ai", data=self._form(token=token))
        self.assertEqual(response.status_code, 200)
        self.assertIn("服务器未配置", response.get_data(as_text=True))
        self.assertIsNone(get_active_ai_config_version_id())

    def test_rollback_creates_a_new_immutable_version(self) -> None:
        self._login_session("admin1")
        token = self._csrf_token()
        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            self.client.post("/admin/pdf-excel-ai", data=self._form(token=token, model="model-v1"))
        first_id = get_active_ai_config_version_id()
        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            self.client.post(
                "/admin/pdf-excel-ai",
                data=self._form(token=token, expected=str(first_id), model="model-v2", api_key=""),
            )
        second_id = get_active_ai_config_version_id()
        rollback_form = {
            "csrf_token": token,
            "action": "rollback",
            "version_id": str(first_id),
            "expected_active_version_id": str(second_id),
            "current_password": "admin-pass",
        }
        with patch.object(routes, "test_repair_connection", return_value="连接测试通过"):
            response = self.client.post("/admin/pdf-excel-ai", data=rollback_form)
        self.assertEqual(response.status_code, 302)
        third_id = get_active_ai_config_version_id()
        self.assertGreater(third_id, second_id)
        rolled_back = get_ai_repair_config(third_id, strict=True)
        self.assertEqual(rolled_back.model, "model-v1")
        with db.db_cursor() as conn:
            row = conn.execute(
                "SELECT source_version_id FROM pdf_excel_ai_config_versions WHERE id = ?", (third_id,)
            ).fetchone()
        self.assertEqual(row["source_version_id"], first_id)

    def test_business_instructions_are_separate_from_fixed_protocol_rules(self) -> None:
        config = AiRepairConfig(
            enabled=True,
            api_key="test-key",
            base_url="https://api.deepseek.com",
            model="test-model",
            timeout_seconds=30,
            max_rows=10,
            repair_instruction="管理员字段修复指令",
            rebuild_instruction="管理员正文重建指令",
            header_mapping_instruction="管理员表头映射指令",
        )
        row = {
            "page_index": 0,
            "table_index": 0,
            "row_index": 1,
            "raw_text": "A-1 2卷",
            "original": {"物料编号": "A-1"},
            "standard": {"物料编码": "A-1", "数量": ""},
        }
        repair_payload = _build_repair_payload(
            {"source_file": "order.pdf", "raw_detail_tables": []},
            [{"row_key": 0, "row": row, "missing_fields": ["数量"], "suspect_fields": [], "issues": []}],
            config,
        )
        self.assertEqual(repair_payload["business_instruction"], "管理员字段修复指令")
        self.assertTrue(any("不要凭空" in rule for rule in repair_payload["rules"]))

        rebuild_payload = _body_missing_payload(
            {
                "source_file": "order.pdf",
                "raw_detail_tables": [{"rows": [["A-1", "2卷"]]}],
                "pages": [],
            },
            10,
            config,
        )
        self.assertEqual(rebuild_payload["business_instruction"], "管理员正文重建指令")
        self.assertIn("allowed_fields", rebuild_payload)

        document = {
            "source_file": "order.pdf",
            "mapped_detail_rows": [{"original": {"客户料号": "A-1"}}],
        }
        with patch(
            "fangzheng_web_app.purchase_factory_mapper.request_repair_json",
            return_value={"mappings": []},
        ) as mocked_request:
            _request_ai_header_mapping(document, config=config)
        header_payload = mocked_request.call_args.args[1]
        self.assertEqual(header_payload["business_instruction"], "管理员表头映射指令")
        self.assertIn("target_fields", header_payload)


if __name__ == "__main__":
    unittest.main()
