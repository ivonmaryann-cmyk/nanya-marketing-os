from __future__ import annotations

import unittest

from fangzheng_web_app.routes import _order_change_reply_body


class OrderChangeReplyBodyTests(unittest.TestCase):
    def test_places_generated_table_before_the_original_html_table(self) -> None:
        body = _order_change_reply_body(
            "<p>您好</p><table><tr><td>原邮件表格</td></tr></table>", "", [{"line_no": 1, "values": {
                "customer_order_number": "PO-001", "line_no": "1", "customer_product_code": "CUST-001",
                "customer_spec": "客户规格", "delivery_date": "2026-09-25", "quantity": "10",
            }}], {},
        )

        self.assertIn("客户订单号", body)
        self.assertIn("原邮件表格", body)
        self.assertIn("----- 原邮件 -----", body)
        self.assertLess(body.index("客户订单号"), body.index("原邮件表格"))

    def test_quotes_original_mail_metadata_below_the_new_table(self) -> None:
        body = _order_change_reply_body(
            "", "原邮件正文", [], {},
            sender="buyer@example.com", occurred_at="2026-09-16 10:00:00", subject="订单变更",
        )

        self.assertIn("----- 原邮件 -----", body)
        self.assertIn("发件人：buyer@example.com", body)
        self.assertIn("发送时间：2026-09-16 10:00:00", body)
        self.assertIn("主题：订单变更", body)
        self.assertLess(body.index("客户订单号"), body.index("----- 原邮件 -----"))

    def test_builds_change_table_from_template_and_erp_demand_date_plus_transit_days(self) -> None:
        body = _order_change_reply_body(
            "", "无表格正文",
            [{"line_no": 3, "values": {
                "customer_order_number": "PO-001", "line_no": "3",
                "customer_product_code": "CUST-001", "customer_spec": "FR-4 <1.0mm>",
                "delivery_date": "2026-09-25", "quantity": "100",
            }}],
            {3: {"selected_candidate": {"sctb16": "2026-09-22"}}}, transit_days="3",
        )

        self.assertIn("厂内交期回复", body)
        self.assertIn("2026-09-25", body)
        self.assertIn("FR-4 &lt;1.0mm&gt;", body)
        self.assertIn("无表格正文", body)


if __name__ == "__main__":
    unittest.main()
