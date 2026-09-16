from __future__ import annotations

import unittest

from fangzheng_web_app.routes import _order_change_reply_body


class OrderChangeReplyBodyTests(unittest.TestCase):
    def test_keeps_original_html_table(self) -> None:
        body = _order_change_reply_body(
            "<p>您好</p><table><tr><td>原邮件表格</td></tr></table>", "", [], {},
        )

        self.assertIn("<table>", body)
        self.assertIn("原邮件表格", body)

    def test_builds_change_table_from_template_and_erp_demand_date(self) -> None:
        body = _order_change_reply_body(
            "", "无表格正文",
            [{"line_no": 3, "values": {
                "customer_order_number": "PO-001", "line_no": "3",
                "customer_product_code": "CUST-001", "customer_spec": "FR-4 <1.0mm>",
                "delivery_date": "2026-09-25", "quantity": "100",
            }}],
            {3: {"selected_candidate": {"sctb16": "2026-09-22"}}},
        )

        self.assertIn("厂内交期回复", body)
        self.assertIn("2026-09-22", body)
        self.assertIn("FR-4 &lt;1.0mm&gt;", body)


if __name__ == "__main__":
    unittest.main()
