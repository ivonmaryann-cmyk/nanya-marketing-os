from __future__ import annotations

import unittest

from fangzheng_web_app.mail_transcode_agent.mail_html_parser import (
    decode_and_simplify_html,
    decode_mail_text,
)


class MailContentDecodingTests(unittest.TestCase):
    def test_falls_back_to_cp950_when_a_traditional_chinese_mail_is_marked_utf8(self) -> None:
        payload = "親愛的，請確認需求，謝謝。".encode("cp950")

        self.assertEqual(decode_mail_text(payload, "utf-8"), "親愛的，請確認需求，謝謝。")

    def test_html_is_simplified_after_charset_fallback(self) -> None:
        payload = "<p>親愛的，請確認需求，謝謝。</p>".encode("cp950")

        self.assertEqual(decode_and_simplify_html(payload, "utf-8"), "<p>亲爱的，请确认需求，谢谢。</p>")


if __name__ == "__main__":
    unittest.main()
