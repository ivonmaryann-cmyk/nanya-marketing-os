import unittest
from datetime import datetime
from pathlib import Path
from jinja2 import Environment, FileSystemLoader


class MailDatetimeDisplayTests(unittest.TestCase):
    def test_display_format(self):
        env = Environment(loader=FileSystemLoader(Path(__file__).resolve().parents[1] / 'templates'), autoescape=True)
        render = env.get_template('macros/mail_datetime.html').module.mail_datetime
        for value in ('2026-09-10T09:12:36', '2026-09-10T09:12:36.123456', '2026-09-10 09:12:36', datetime(2026,9,10,9,12,36)):
            self.assertEqual(str(render(value)), '2026-09-10 09:12:36')
        self.assertEqual(str(render('2026-09-10T09:12')), '2026-09-10 09:12:00')
        self.assertEqual(str(render('从未抓取')), '从未抓取')
        self.assertEqual(str(render(None)), '')
        self.assertEqual(str(render('2026-09-10')), '2026-09-10')
