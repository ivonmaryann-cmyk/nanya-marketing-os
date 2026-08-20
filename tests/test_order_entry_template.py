from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_entry_service import (
    _line_entry,
    _line_from_pipeline_row,
    _merge_initial_rows,
    build_domestic_export,
    get_or_create_template,
    reextract_all_templates,
    save_template,
)
from fangzheng_web_app.order_intake_service import bootstrap_cases, list_cases


class OrderEntryTemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "entry.sqlite3")
        self.db_patch.start()
        db.init_db()
        self.account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            self.account_id, folder="INBOX", uid="entry-1", message_id="<entry-1@example.com>",
            subject="采购订单", sender="buyer@example.com", sent_at="2026-08-19 09:00:00",
            received_at="2026-08-19 09:00:00", body_html="", body_text=(
                "HJ20260818013\nA1A150224149YNNYZ002\n南亚\nNY2150\nFR-4 1.5\n±\n"
                "0.075MM 2/2 41\n英寸\n*49\n英寸含铜无水印\n(TG>150)A\n级\n300"
            ), eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", self.account_id)
        self.case_id = list_cases("employee-a", "2026-08-19", "new_order", self.account_id)[0]["id"]

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_saved_template_reopens_and_exports_same_values(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        self.assertEqual(template["lines"][0]["values"]["customer_order_number"], "HJ20260818013")
        saved = save_template(self.case_id, "employee-a", {
            "header": {"order_type": "SO", "bill_to_customer_code": "C001", "ledger": "151"},
            "lines": [{"values": {
                "line_no": "1", "product_code": "P001", "product_name": "南亚NY2150",
                "customer_product_code": "A1A150224149YNNYZ002", "quantity": "300",
                "customer_order_number": "HJ20260818013",
            }}],
        })
        self.assertEqual(saved["current_version"], 1)
        _case, reopened = get_or_create_template(self.case_id, "employee-a")
        self.assertEqual(reopened["header"]["ledger"], "151")
        output, _name = build_domestic_export(self.case_id, "employee-a")
        destination = Path(self.temp_dir.name) / "export.xlsx"
        destination.write_bytes(output.getvalue())
        book = load_workbook(destination, data_only=True)
        sheet = book["内销"]
        self.assertEqual(sheet["A2"].value, "SO")
        self.assertEqual(sheet["D2"].value, "C001")
        self.assertEqual(sheet["H2"].value, "151")
        self.assertEqual(sheet["D4"].value, "A1A150224149YNNYZ002")
        self.assertEqual(sheet["G4"].value, 300)
        # The business template is a populated example; exporting a different
        # mail must never carry its sample detail rows into the new file.
        self.assertIsNone(sheet["D5"].value)
        self.assertIsNone(sheet["G6"].value)
        book.close()

    def test_excel_attachment_is_merged_into_the_same_email_template(self) -> None:
        attachment_path = Path(self.temp_dir.name) / "客户订单.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "订单"
        sheet.append(["客户订单号", "客户产品编号", "客户规格", "数量", "含税单价", "备注"])
        sheet.append(["PO-20260819", "CUST-001", "FR-4 1.6MM", 500, 12.5, "加急"])
        workbook.save(attachment_path)
        workbook.close()
        with db.db_cursor() as conn:
            mail_id = conn.execute("SELECT id FROM mail_messages WHERE uid='entry-1'").fetchone()["id"]
            conn.execute(
                """INSERT INTO mail_attachments(mail_id,filename,content_type,size_bytes,sha256,stored_path,is_inline,parse_status,created_at)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (mail_id, "客户订单.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", attachment_path.stat().st_size, "", str(attachment_path), 0, "", db.utcnow()),
            )
        _case, template = get_or_create_template(self.case_id, "employee-a")
        matching = [line for line in template["lines"] if line["values"].get("customer_product_code") == "CUST-001"]
        self.assertEqual(len(matching), 1)
        self.assertEqual(matching[0]["values"]["quantity"], "500")
        self.assertEqual(matching[0]["values"]["unit_price"], "12.5")
        self.assertEqual(matching[0]["sources"]["quantity"]["label"], "附件：客户订单.xlsx")

    def test_ccl_and_manual_only_fields_follow_conservative_extraction_policy(self) -> None:
        line = _line_entry(
            {
                "customer_product_code": "CUST-001",
                "customer_spec": "FR-4 1.6MM",
                "quantity": "500",
                "product_code": "FACTORY-001",
                "product_name": "不应自动填入的品名",
                "origin": "中国",
                "one_to_many": "一对多关系",
            },
            label="附件：订单.xlsx", reference="订单 第 2 行",
        )
        values = line["values"]
        self.assertEqual(values["quantity"], "500")
        self.assertEqual(values["customer_product_code"], "CUST-001")
        for field in ("product_code", "product_name", "origin", "one_to_many"):
            self.assertEqual(values[field], "")
            self.assertNotIn(field, line["sources"])

    def test_pp_roll_converts_quantity_only_with_explicit_unit_and_metre_value(self) -> None:
        line = _line_entry(
            {
                "customer_product_code": "PP-01",
                "customer_spec": "PP 1080 300M/卷",
                "quantity": "2",
                "remark": "客户加急",
            },
            label="附件：PP订单.xlsx", reference="订单 第 2 行", quantity_unit="卷",
        )
        self.assertEqual(line["values"]["quantity"], "600")
        self.assertEqual(line["values"]["remark"], "客户加急；PP米数：300米")

    def test_pp_sheet_keeps_customer_quantity_and_unknown_pp_is_blank(self) -> None:
        small_piece = _line_entry(
            {"customer_spec": "PP 1080 300m", "quantity": "30"},
            label="附件：PP订单.xlsx", reference="订单 第 2 行", quantity_unit="张",
        )
        unknown = _line_entry(
            {"customer_spec": "PP 1080 300m", "quantity": "2"},
            label="附件：PP订单.xlsx", reference="订单 第 3 行", quantity_unit="",
        )
        self.assertEqual(small_piece["values"]["quantity"], "30")
        self.assertEqual(small_piece["values"]["remark"], "PP米数：300米")
        self.assertEqual(unknown["values"]["quantity"], "")
        self.assertEqual(unknown["values"]["remark"], "PP米数：300米")

    def test_pipeline_mapping_uses_raw_description_and_keeps_pp_detail_rows(self) -> None:
        first = _line_from_pipeline_row(
            {
                "original": {
                    "PO项目号 Project No": "2713186",
                    "物料编码 Material Code": "10601001676",
                    "物料品名 Material Name": "半固化片",
                    "物料描述 Description": "南亚新材料 NY6180LP 106 RC=75% 经300.00 m 纬49.50 inch",
                    "单位 Unit": "卷",
                    "数量 Quantity": "0.1",
                    "不含税单价 Not tax inclusive Unit Price": "10914.1593",
                },
                "standard": {"数量": "0.1", "单位": "卷"},
            },
            "MZPOM12608190027", "附件：采购订单.pdf", "识别明细第 2 行", 1,
        )
        second = _line_from_pipeline_row(
            {
                "original": {
                    "PO项目号 Project No": "2713185",
                    "物料编码 Material Code": "10601001674",
                    "物料品名 Material Name": "半固化片",
                    "物料描述 Description": "南亚新材料 NY6180LP 2116 RC=56% 经200.00 m 纬49.50 inch",
                    "单位 Unit": "卷",
                    "数量 Quantity": "0.1",
                },
                "standard": {"数量": "0.1", "单位": "卷"},
            },
            "MZPOM12608190027", "附件：采购订单.pdf", "识别明细第 3 行", 2,
        )
        self.assertEqual(first["values"]["customer_spec"], "南亚新材料 NY6180LP 106 RC=75% 经300.00 m 纬49.50 inch")
        self.assertEqual(first["values"]["quantity"], "30")
        self.assertEqual(first["values"]["price_before_tax"], "10914.1593")
        self.assertEqual(first["values"]["unit_price"], "")
        self.assertEqual(first["values"]["product_name"], "")
        self.assertEqual(second["values"]["quantity"], "20")
        self.assertEqual(len(_merge_initial_rows([first, second])), 2)

    def test_pipeline_mapping_keeps_a_real_tax_inclusive_unit_price(self) -> None:
        line = _line_from_pipeline_row(
            {
                "original": {
                    "物料描述 Description": "测试规格",
                    "数量 Quantity": "2",
                    "含税单价 Tax inclusive Unit Price": "12.50",
                },
                "standard": {"数量": "2"},
            },
            "PO123456", "附件：采购订单.pdf", "识别明细第 2 行", 1,
        )
        self.assertEqual(line["values"]["price_before_tax"], "")
        self.assertEqual(line["values"]["unit_price"], "12.50")

    def test_batch_reextract_keeps_header_and_backs_up_before_replacing_lines(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"order_type": "SO", "bill_to_customer_code": "C001", "ledger": "151"},
            "lines": [{"values": {
                "line_no": "1", "product_name": "历史错误品名",
                "customer_product_code": "OLD-CODE", "quantity": "999",
            }}],
        })

        summary = reextract_all_templates("employee-a")
        self.assertEqual(summary["template_count"], 1)
        self.assertEqual(summary["previous_line_count"], 1)
        self.assertGreaterEqual(summary["line_count"], 1)
        _case, template = get_or_create_template(self.case_id, "employee-a")
        self.assertEqual(template["header"]["ledger"], "151")
        self.assertEqual(template["current_version"], 3)
        self.assertTrue(all(line["values"]["product_name"] == "" for line in template["lines"]))
        with db.db_cursor() as conn:
            versions = conn.execute(
                "SELECT version_number,lines_json FROM order_entry_template_versions ORDER BY version_number"
            ).fetchall()
        self.assertEqual([row["version_number"] for row in versions], [1, 2, 3])
        self.assertIn("历史错误品名", versions[1]["lines_json"])
