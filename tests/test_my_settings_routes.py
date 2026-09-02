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
        self.assertIn("我的", html)
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
        self.assertIn("返回我的", html)
        self.assertIn('href="/my-settings"', html)
        self.assertIn("IMAP 收信配置", html)
        self.assertIn("SMTP 发信配置（可选）", html)
        self.assertIn("使用上方收信授权码", html)

    def test_existing_mailbox_exposes_matching_smtp_secret_controls(self) -> None:
        account = {
            "id": 7,
            "email": "tester@example.com",
            "imap_host": "imap.example.com",
            "imap_port": 993,
            "enabled": 1,
            "smtp_host": "smtp.example.com",
            "smtp_port": 465,
            "smtp_security": "ssl",
            "smtp_username": "tester@example.com",
            "smtp_auth_code_ciphertext": "configured",
            "smtp_enabled": 0,
            "smtp_last_test_status": "",
            "smtp_last_test_at": "",
        }
        with patch("fangzheng_web_app.routes.get_user", return_value=None), patch(
            "fangzheng_web_app.mail_transcode_agent.routes.mail_store.list_accounts",
            return_value=[account],
        ), patch(
            "fangzheng_web_app.mail_transcode_agent.routes.mail_store.get_account",
            return_value=account,
        ):
            response = self.client.get("/mail-transcode/accounts?edit=7")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("IMAP 收信配置", html)
        self.assertIn("SMTP 发信配置", html)
        self.assertIn("reveal-smtp-auth-code", html)
        self.assertIn("copy-smtp-auth-code", html)
        self.assertIn("使用收信授权码", html)

    def test_smtp_authorization_code_is_only_revealed_for_the_owner(self) -> None:
        with patch(
            "fangzheng_web_app.mail_transcode_agent.routes.mail_store.get_smtp_config",
            return_value={"auth_code": "smtp-secret"},
        ):
            response = self.client.post("/mail-transcode/accounts/7/reveal-smtp-auth-code")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"auth_code": "smtp-secret"})


if __name__ == "__main__":
    unittest.main()
