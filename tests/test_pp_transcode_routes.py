from __future__ import annotations

import io
import unittest
from unittest.mock import patch

from flask import Flask

from fangzheng_web_app.routes import bp


class PPTranscodeRouteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = Flask(
            __name__,
            template_folder="../templates",
            static_folder="../static",
        )
        self.app.config.update(SECRET_KEY="test-secret", TESTING=True)
        self.app.register_blueprint(bp)
        self.client = self.app.test_client()
        with self.client.session_transaction() as session:
            session["employee_id"] = "tester"

    def test_pp_page_uses_unified_single_and_batch_task_controls(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.routes.list_jobs", return_value=[]
        ):
            response = self.client.get("/features/pp-transcode-agent")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("开始单条PP转码", html)
        self.assertIn("上传 PP 转码需求 Excel", html)
        self.assertIn("最近 PP 转码任务", html)
        self.assertIn("处理日志", html)

    def test_confirmation_api_returns_pending_code_for_single_result(self) -> None:
        job = {"id": 88, "employee_id": "tester", "feature": "pp_transcode_agent", "status": "awaiting_confirmation"}
        item = {
            "id": 7,
            "excel_row": 2,
            "customer_code": "100001",
            "customer_name": "测试客户",
            "spec": "PP NY2150 1080 RC70% 300M/卷",
            "order_remark": "测试备注",
            "pending_code": "2B1080300MR700*************",
            "confirmed_pending_code": "",
            "confidence": 60,
            "summary": "请进入确认中心",
            "field_evidence": [],
            "confirmation_status": "pending",
        }
        self.assertEqual(len(item["pending_code"]), 27)
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.routes.get_job", return_value=job
        ), patch(
            "fangzheng_web_app.routes.list_pp_confirmation_items", return_value=[item]
        ), patch(
            "fangzheng_web_app.routes.pp_confirmation_counts",
            return_value={"total": 1, "pending": 1, "confirmed": 0, "skipped": 0},
        ):
            response = self.client.get("/api/pp-transcode-agent/jobs/88/confirmations")

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertEqual(payload["job_status"], "awaiting_confirmation")
        self.assertEqual(payload["records"][0]["pending_code"], item["pending_code"])

    def test_batch_upload_redirects_to_current_pp_task(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.routes.queue_pp_transcode_job", return_value=91
        ):
            response = self.client.post(
                "/features/pp-transcode-agent/jobs",
                data={"pp_file": (io.BytesIO(b"test"), "pp.xlsx")},
                content_type="multipart/form-data",
            )

        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers["Location"].endswith("/features/pp-transcode-agent?job_id=91"))


if __name__ == "__main__":
    unittest.main()
