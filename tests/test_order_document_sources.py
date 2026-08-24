from __future__ import annotations

import unittest

from fangzheng_web_app.order_document_sources import (
    build_mail_html_purchase_document,
    order_source_adapter_catalog,
)
from fangzheng_web_app.pdf_excel_domestic_export import build_domestic_template_data
from fangzheng_web_app.purchase_factory_mapper import project_factory_document
from fangzheng_web_app.purchase_field_rules import order_field_rule_catalog


ORDER_TABLE_HTML = """
<table>
  <tr>
    <th>PO单号</th><th>PO项目号</th><th>物料编码</th><th>物料名称</th>
    <th>物料规格</th><th>数量</th><th>单位</th><th>单价</th><th>交期</th>
  </tr>
  <tr>
    <td>MZPOM12608190027</td><td>2713183</td><td>10101008248</td>
    <td>FR-4（高速材料）</td>
    <td>南亚新材料 NY6180L 0.127 mm 1/1 RTF2/RTF2 经41*纬49 inch</td>
    <td>20</td><td>张</td><td>215.6814</td><td>2026-08-28</td>
  </tr>
  <tr>
    <td>MZPOM12608190027</td><td>2713184</td><td>10601001676</td>
    <td>半固化片</td>
    <td>南亚新材料 NY6180LP 106 RC=75% 经300.00 m 纬49.50 inch</td>
    <td>30</td><td>张</td><td>36.380531</td><td>2026-08-28</td>
  </tr>
</table>
"""


class MailHtmlOrderDocumentTests(unittest.TestCase):
    def test_html_table_reuses_canonical_document_and_domestic_mapping(self) -> None:
        document = build_mail_html_purchase_document(ORDER_TABLE_HTML)

        self.assertIsNotNone(document)
        assert document is not None
        self.assertEqual(document["source_adapter"]["kind"], "mail_html_table")
        self.assertEqual(document["header_info"]["订单号"], "MZPOM12608190027")
        self.assertEqual(document["mapped_detail_rows"][0]["standard"]["序号"], "2713183")
        self.assertEqual(document["mapped_detail_rows"][1]["standard"]["物料编码"], "10601001676")

        # The output boundary is exactly the same factory projection and
        # domestic-template builder used by PDF/image recognition.
        project_factory_document(document)
        domestic = build_domestic_template_data(document)
        self.assertEqual(domestic["header"]["customer_order_number"], "MZPOM12608190027")
        self.assertEqual(domestic["lines"][0]["customer_order_seq"], "2713183")
        self.assertEqual(domestic["lines"][0]["product_type"], "基板")
        self.assertEqual(domestic["lines"][1]["customer_product_code"], "10601001676")
        self.assertEqual(domestic["lines"][1]["product_type"], "PP")
        self.assertEqual(domestic["lines"][1]["quantity"], "30")

    def test_non_order_html_table_is_not_misclassified(self) -> None:
        self.assertIsNone(
            build_mail_html_purchase_document(
                "<table><tr><th>联系人</th><th>电话</th></tr><tr><td>张三</td><td>123</td></tr></table>"
            )
        )

    def test_rule_catalogues_are_data_only_and_include_mail_po_headings(self) -> None:
        field_rules = order_field_rule_catalog()
        source_rules = order_source_adapter_catalog()

        self.assertEqual(field_rules["version"], "purchase_order_fields_v1")
        self.assertIn("PO项目号", field_rules["detail_headers"]["序号"])
        self.assertIn("PO单号", field_rules["header_fields"]["订单号"])
        self.assertEqual(source_rules["adapters"]["mail_html_table"]["source_kind"], "mail_html_table")
