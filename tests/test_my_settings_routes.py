from __future__ import annotations

import unittest
from unittest.mock import patch

from flask import Flask

from fangzheng_web_app.mail_transcode_agent.routes import bp as mail_transcode_bp
from fangzheng_web_app.routes import bp


class MySettingsRouteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = Flask(
            __name__,
            template_folder="../templates",
            static_folder="../static",
        )
        self.app.config.update(SECRET_KEY="test-secret", TESTING=True)
        self.app.register_blueprint(bp)
        self.app.register_blueprint(mail_transcode_bp, url_prefix="/mail-transcode")
        self.client = self.app.test_client()
        with self.client.session_transaction() as session:
            session["employee_id"] = "tester"

    def test_my_settings_is_a_first_level_page_with_mail_as_second_level(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=None):
            response = self.client.get("/my-settings")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("我的配置", html)
        self.assertIn("邮箱配置", html)
        self.assertIn('href="/mail-transcode/accounts"', html)

    def test_mailbox_page_returns_to_my_settings(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.mail_transcode_agent.routes.mail_store.list_accounts",
            return_value=[],
        ), patch(
            "fangzheng_web_app.mail_transcode_agent.routes.mail_store.list_fetch_logs",
            return_value=[],
        ):
            response = self.client.get("/mail-transcode/accounts")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("返回我的配置", html)
        self.assertIn('href="/my-settings"', html)


if __name__ == "__main__":
    unittest.main()
