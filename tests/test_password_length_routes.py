from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from flask import Flask

from fangzheng_web_app.routes import bp


class PasswordLengthRouteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = Flask(
            __name__,
            template_folder=str(Path(__file__).parents[1] / "templates"),
        )
        self.app.config.update(SECRET_KEY="test-secret", TESTING=True)
        self.app.register_blueprint(bp)
        self.client = self.app.test_client()
        self.user = {"employee_id": "tester", "must_change_password": False}
        with self.client.session_transaction() as session:
            session["employee_id"] = "tester"

    def test_change_password_accepts_five_characters_and_rejects_four(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=self.user), patch(
            "fangzheng_web_app.routes.verify_user_password", return_value=True
        ), patch("fangzheng_web_app.routes.change_user_password") as change_password:
            accepted = self.client.post(
                "/change-password",
                data={"old_password": "old", "new_password": "abcde", "confirm_password": "abcde"},
            )
            rejected = self.client.post(
                "/change-password",
                data={"old_password": "old", "new_password": "abcd", "confirm_password": "abcd"},
            )

        self.assertEqual(accepted.status_code, 302)
        change_password.assert_called_once_with("tester", "abcde")
        self.assertEqual(rejected.status_code, 200)
        self.assertIn("新密码至少 5 位。", rejected.get_data(as_text=True))

    def test_account_page_password_actions_accept_five_characters(self) -> None:
        with patch("fangzheng_web_app.routes.get_user", return_value=self.user), patch(
            "fangzheng_web_app.routes.verify_user_password", return_value=True
        ), patch("fangzheng_web_app.routes.verify_admin_password", return_value=True), patch(
            "fangzheng_web_app.routes.change_user_password"
        ) as change_password, patch("fangzheng_web_app.routes.update_admin_password") as update_password:
            user_response = self.client.post(
                "/admin/password?mode=user",
                data={
                    "action": "user_password",
                    "old_password": "old",
                    "new_password": "abcde",
                    "confirm_password": "abcde",
                },
            )
            admin_response = self.client.post(
                "/admin/password?mode=admin",
                data={
                    "action": "password",
                    "current_password": "old-admin",
                    "new_password": "abcde",
                    "confirm_password": "abcde",
                },
            )

        self.assertEqual(user_response.status_code, 302)
        self.assertEqual(admin_response.status_code, 302)
        change_password.assert_called_once_with("tester", "abcde")
        update_password.assert_called_once_with("abcde")

    def test_password_forms_expose_five_character_minimum(self) -> None:
        change_template = (Path(__file__).parents[1] / "templates" / "change_password.html").read_text(encoding="utf-8")
        admin_template = (Path(__file__).parents[1] / "templates" / "admin_password.html").read_text(encoding="utf-8")

        self.assertNotIn('minlength="6"', change_template)
        self.assertEqual(change_template.count('minlength="5"'), 2)
        self.assertEqual(admin_template.count('minlength="5"'), 4)


if __name__ == "__main__":
    unittest.main()
