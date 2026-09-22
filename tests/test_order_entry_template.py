from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_entry_service import (
    _apply_customer_extraction_mappings,
    _apply_customer_spec_matches,
    _apply_customer_transit_days,
    _attachment_rows,
    _line_entry,
    _line_from_pipeline_row,
    _merge_initial_rows,
    _remove_prefixed_line_no_from_customer_product,
    _initial_order_groups,
    _initial_template_data,
    _split_body_order_rows,
    _rows_from_pdf_or_image,
    build_domestic_export,
    get_or_create_template,
    reextract_template,
    reextract_all_templates,
    save_template,
    normalize_customer_order_number,
)
from fangzheng_web_app.purchase_factory_mapper import FACTORY_DETAIL_HEADERS
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

    def test_templates_expose_shared_column_filter_controls(self) -> None:
        root = Path(__file__).resolve().parents[1]
        entry = (root / "templates" / "order_automation_entry_template.html").read_text(encoding="utf-8")
        change = (root / "templates" / "order_automation_order_change_template.html").read_text(encoding="utf-8")
        editor = (root / "static" / "order_template_bulk_editor.js").read_text(encoding="utf-8")

        self.assertIn('data-template-filter-field="{{ field }}"', entry)
        self.assertIn('data-template-filter-field="{{ field }}"', change)
        self.assertIn("template-filter-dialog", editor)
        self.assertIn("筛选：显示", editor)
        self.assertIn("清除筛选", editor)
        self.assertIn("specialFilterMatches", editor)
        self.assertIn("需要新建料号", entry)
        self.assertIn("需要多选候选料号", entry)

    def test_pdf_column_merge_does_not_prefix_customer_product_with_item_number(self) -> None:
        self.assertEqual(
            _remove_prefixed_line_no_from_customer_product("1AAN31AW01520022", "1"),
            "AAN31AW01520022",
        )
        self.assertEqual(
            _remove_prefixed_line_no_from_customer_product("10LAN31AW16600D", "10"),
            "LAN31AW16600D",
        )
        self.assertEqual(
            _remove_prefixed_line_no_from_customer_product("1A23", "1"), "1A23",
        )

    def test_initial_template_dates_are_backdated_by_customer_transit_days(self) -> None:
        lines = [{
            "values": {"delivery_date": "2026-09-25", "customer_spec": "NY2150"},
            "sources": {"delivery_date": {"label": "邮件正文", "reference": "第 1 行"}},
        }]
        with patch("fangzheng_web_app.order_entry_service.get_customer", return_value={"transit_days": "3"}):
            result = _apply_customer_transit_days({"customer_id": 12}, lines)

        self.assertEqual("2026-09-22", result[0]["values"]["delivery_date"])
        self.assertEqual("客户运输天数倒推", result[0]["sources"]["delivery_date"]["label"])
        self.assertIn("客户需求日 2026-09-25 - 运输天数 3 天", result[0]["sources"]["delivery_date"]["reference"])

    def test_empty_configured_spec_field_shows_its_position_as_placeholder(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn("input.placeholder=`第${field.position}位`", template)
        self.assertIn(".oe-match-field input::placeholder{color:#9aa8ba;opacity:1}", template)
        self.assertIn(
            "shipToCode.value=code.value.trim()",
            template,
        )
        self.assertIn("<datalist id=\"orderTypeOptions\">", template)
        self.assertIn("<datalist id=\"customerCodeOptions\">", template)
        self.assertNotIn("customer.customer_name", template)
        self.assertIn("220:['1','1'],221:['3','2'],331:['3','1']", template)
        self.assertIn("normalizeChoice(orderType,3)", template)
        self.assertIn("normalizeCustomerCode();shipToCode.value=code.value.trim()", template)
        self.assertNotIn('aria-label="料号状态"', template)
        self.assertIn("{% for field in hidden_line_fields %}", template)
        self.assertIn('id="materialCreateDialog"', template)
        self.assertIn('id="createMaterialOpen"', template)
        self.assertIn("!row.querySelector('[data-field=\"product_code\"]')?.value.trim()", template)
        self.assertIn("materialCodeOptions", template)
        self.assertIn("syncPair", template)
        self.assertIn("NYEOS订单号：{{ nyeos_order_number }}", template)

    def test_generated_erp_order_number_is_displayed_with_nyeos_order_number(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn("NYEOS订单号：{{ nyeos_order_number }}", template)
        self.assertIn("ERP订单号：{{ erp_order_number }}", template)

    def test_material_create_dialog_reuses_customer_spec_match_editor(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn('oe-create-match-editor', template)
        self.assertIn('openCustomerSpecMatchEditor', template)
        self.assertIn("control.dispatchEvent(new Event('input'))", template)

    def test_material_create_dialog_keeps_transcoded_product_name_visible(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn("['layout_structure','textarea']", template)
        self.assertIn("['thickness_description','textarea']", template)
        self.assertIn("['special_requirements','textarea']", template)
        self.assertIn("['origin','select']", template)
        self.assertIn("['上海','上海'],['江西','江西'],['江苏','江苏']", template)
        self.assertIn("['customer_spec_match','textarea']", template)
        self.assertIn('id="materialNameValidationDialog"', template)
        self.assertIn("confirmation_required", template)
        self.assertIn("confirm_name_validation", template)

    def test_entry_interface_actions_keep_original_toolbar_layout(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn(".oe-head>.oe-actions,#entryInterfaceActions{display:none}", template)
        self.assertIn("group.className='oe-meta-actions'", template)
        self.assertIn("save.disabled=false;submit.disabled=false}}};", template)

    def test_multiple_material_candidates_use_compact_colored_count_badge(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn(".oe-candidate-name-cell>.oe-candidate-open", template)
        self.assertIn("width:20px!important", template)
        self.assertIn("height:20px!important", template)
        self.assertIn("button.textContent=String(count)", template)
        self.assertIn(".oe-candidate-open.is-selected{background:#43a66d;color:#fff}", template)

    def test_material_candidates_and_template_show_old_product_name(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn("旧品名", template)
        self.assertIn("old_product_name", template)
        self.assertIn("grid_column_count", template)
        self.assertIn("grid_column_letters", template)
        self.assertNotIn("'ABCDEFGHIJKLMNOP'", template)

    def test_entry_template_exposes_bulk_spreadsheet_controls(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")
        editor = (
            Path(__file__).resolve().parents[1]
            / "static"
            / "order_template_bulk_editor.js"
        ).read_text(encoding="utf-8")

        self.assertIn('id="entryBulkTools"', template)
        self.assertIn('data-bulk-fill-down', template)
        self.assertIn('data-bulk-edit', template)
        self.assertIn("excludedFields: ['line_no', 'old_product_name']", template)
        self.assertIn("matrixFromClipboard", editor)
        self.assertIn("event.key.toLowerCase() === 'd'", editor)
        self.assertIn("selectedBounds", editor)

    def test_create_material_dialog_resolves_adhesive_codes_from_spec_match(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")
        self.assertIn("order_automation_entry_template_adhesive_candidates", template)
        self.assertIn("candidate.adhesive_code}｜${candidate.adhesive_name", template)
        self.assertIn("customer_spec_match:match.value", template)

    def test_creation_task_id_is_displayed_without_changing_saved_placeholder(self) -> None:
        template = (
            Path(__file__).resolve().parents[1]
            / "templates"
            / "order_automation_entry_template.html"
        ).read_text(encoding="utf-8")

        self.assertIn("创建料号中（${item.external_task_id}）", template)
        self.assertIn("materialCreationPlaceholder", template)

    def test_saved_template_reopens_and_exports_same_values(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        self.assertEqual(template["header"]["customer_order_number"], "HJ20260818013")
        self.assertEqual(template["header"]["order_type"], "220")
        self.assertEqual(template["header"]["type_1"], "1")
        self.assertEqual(template["header"]["type_2"], "1")
        self.assertEqual(template["header"]["ledger"], "KL01")
        self.assertTrue(all(line["values"]["material_status"] == "查询" for line in template["lines"]))
        saved = save_template(self.case_id, "employee-a", {
            "header": {
                "order_type": "SO", "bill_to_customer_code": "C001", "ledger": "151",
                "tax_type": "VAT", "customer_invoice_number": "INV-001",
                "commission_rate": "2.5%",
            },
            "lines": [{"values": {
                "line_no": "1", "material_status": "新增", "product_code": "P001", "product_name": "南亚NY2150",
                "customer_product_code": "A1A150224149YNNYZ002", "quantity": "300",
                "customer_spec_match": "NY2150", "product_type": "基板",
            }}],
        })
        self.assertEqual(saved["current_version"], 1)
        _case, reopened = get_or_create_template(self.case_id, "employee-a")
        self.assertEqual(reopened["header"]["ledger"], "151")
        self.assertEqual(reopened["lines"][0]["values"]["material_status"], "新增")
        self.assertEqual(reopened["lines"][0]["values"]["adhesive_code"], "")
        output, _name = build_domestic_export(self.case_id, "employee-a")
        destination = Path(self.temp_dir.name) / "export.xlsx"
        destination.write_bytes(output.getvalue())
        book = load_workbook(destination, data_only=True)
        sheet = book["内销"]
        self.assertEqual(sheet["A2"].value, "SO")
        self.assertEqual(sheet["D2"].value, "C001")
        self.assertEqual(sheet["H2"].value, "151")
        self.assertEqual(sheet["I2"].value, "VAT")
        self.assertEqual(sheet["J2"].value, "INV-001")
        self.assertEqual(sheet["K2"].value, "2.5%")
        self.assertEqual(sheet["I1"].value, "税种（选填）")
        self.assertEqual(sheet["J1"].value, "客户发票号（选填）")
        self.assertEqual(sheet["K1"].value, "佣金比率（选填）")
        self.assertEqual(sheet["B3"].value, "料号状态（选填）")
        self.assertEqual(sheet["E3"].value, "胶系编码（选填）")
        self.assertIsNone(sheet["E4"].value)
        self.assertEqual(sheet["B4"].value, "新增")
        self.assertEqual(sheet["F4"].value, "A1A150224149YNNYZ002")
        self.assertIsNone(sheet["G4"].value)
        self.assertEqual(sheet["I4"].value, "基板")
        self.assertEqual(sheet["K4"].value, 300)
        # The business template is a populated example; exporting a different
        # mail must never carry its sample detail rows into the new file.
        self.assertIsNone(sheet["E5"].value)
        self.assertIsNone(sheet["J6"].value)
        book.close()

    def test_multiple_purchase_orders_are_normalized_and_grouped_in_first_seen_order(self) -> None:
        body_lines = ["序号", "PR号", "PO号", "物料编码", "物料描述", "需求数量", "需求交期", "供应商回复", "备注"]
        for index in range(1, 31):
            order_number = "建价PO-26F7-044522" if index <= 12 else "PO-26F7-044521"
            body_lines.extend([
                str(index), "PR-26F7-004045", order_number, f"AA141007612003{index:04d}",
                "NY3170LK 0.076mm 1/2 37x49 无卤 RTF2 1x1080",
                "6.0", "2026/9/22", f"第{index}项备注",
            ])
        lines = _split_body_order_rows("\n".join(body_lines), "2026-09-16")

        groups = _initial_order_groups({"customer_order_number": ""}, lines)

        self.assertEqual(len(lines), 30)
        self.assertEqual(normalize_customer_order_number("建价PO-26F7-044522"), "PO-26F7-044522")
        self.assertEqual(normalize_customer_order_number("暂无PO号-123"), "")
        self.assertEqual([group["order_number"] for group in groups], ["PO-26F7-044522", "PO-26F7-044521"])
        self.assertEqual([len(group["lines"]) for group in groups], [12, 18])

    def test_blank_po_lines_form_a_separate_group_when_other_rows_have_po(self) -> None:
        lines = [
            _line_entry({"customer_order_number": "PO-A", "customer_product_code": "A"}, label="测试", reference="测试", line_no=1),
            _line_entry({"customer_order_number": "", "customer_product_code": "B"}, label="测试", reference="测试", line_no=2),
            _line_entry({"customer_order_number": "PO-A", "customer_product_code": "C"}, label="测试", reference="测试", line_no=3),
        ]
        groups = _initial_order_groups({"customer_order_number": "PO-A"}, lines)

        self.assertEqual([group["order_number"] for group in groups], ["PO-A", ""])
        self.assertEqual([len(group["lines"]) for group in groups], [2, 1])
        self.assertTrue(groups[1]["group_key"])

    def test_blank_po_group_uses_a_stable_display_identifier(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        group_key = template["groups"][0]["group_key"]
        save_template(self.case_id, "employee-a", {"groups": [{
            "group_key": group_key,
            "order_number": "",
            "header": {"customer_order_number": ""},
            "lines": [{"values": {"line_no": "1", "customer_product_code": "CUST-001", "quantity": "1"}}],
        }]})

        _case, refreshed = get_or_create_template(self.case_id, "employee-a")
        group = refreshed["groups"][0]
        self.assertEqual(group["order_number"], "")
        self.assertEqual(group["display_order_number"], f"暂无PO号-{group_key[:8]}")

    def test_group_headers_and_current_po_export_are_independent(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        saved = save_template(self.case_id, "employee-a", {"groups": [
            {
                "group_key": template["groups"][0]["group_key"],
                "order_number": "PO-A",
                "header": {"order_type": "220", "bill_to_customer_code": "C001", "ledger": "KL01", "customer_order_number": "PO-A"},
                "lines": [{"values": {"line_no": "1", "customer_product_code": "A-001", "quantity": "1"}}],
            },
            {
                "group_key": "po-b",
                "order_number": "PO-B",
                "header": {"order_type": "331", "bill_to_customer_code": "C002", "ledger": "KL02", "customer_order_number": "PO-B"},
                "lines": [{"values": {"line_no": "1", "customer_product_code": "B-001", "quantity": "2"}}],
            },
        ]})

        self.assertEqual([group["header"]["ledger"] for group in saved["groups"]], ["KL01", "KL02"])
        self.assertEqual(len({line["line_no"] for line in saved["lines"]}), 2)
        output, filename = build_domestic_export(self.case_id, "employee-a", "po-b")
        destination = Path(self.temp_dir.name) / "po-b.xlsx"
        destination.write_bytes(output.getvalue())
        book = load_workbook(destination, data_only=True)
        self.assertEqual(book["内销"]["A2"].value, "331")
        self.assertEqual(book["内销"]["D2"].value, "C002")
        self.assertEqual(book["内销"]["F4"].value, "B-001")
        self.assertIn("PO-B", filename)
        book.close()

    def test_refresh_preserves_submitted_po_and_replaces_only_pending_groups(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        saved = save_template(self.case_id, "employee-a", {"groups": [
            {
                "group_key": template["groups"][0]["group_key"], "order_number": "PO-A",
                "header": {"customer_order_number": "PO-A"},
                "lines": [{"values": {"line_no": "1", "customer_product_code": "LOCKED", "quantity": "1"}}],
            },
            {
                "group_key": "po-b", "order_number": "PO-B",
                "header": {"customer_order_number": "PO-B"},
                "lines": [{"values": {"line_no": "2", "customer_product_code": "OLD-B", "quantity": "1"}}],
            },
        ]})
        with db.db_cursor() as conn:
            conn.execute(
                "UPDATE order_entry_template_groups SET status='submitted' WHERE group_key=?",
                (saved["groups"][0]["group_key"],),
            )
        refreshed_lines = [
            _line_entry({"customer_order_number": "PO-A", "customer_product_code": "NEW-A", "quantity": "1"}, label="测试", reference="刷新", line_no=1),
            _line_entry({"customer_order_number": "PO-B", "customer_product_code": "NEW-B", "quantity": "1"}, label="测试", reference="刷新", line_no=2),
        ]
        snapshot = {"tax_mode": "unknown", "target_field": "", "target_label": "", "by_line": {}, "mismatches": []}
        with patch(
            "fangzheng_web_app.order_entry_service._initial_template_data",
            return_value=({"customer_order_number": ""}, refreshed_lines),
        ), patch(
            "fangzheng_web_app.order_entry_service.review_case_template_prices", return_value=snapshot,
        ):
            result = reextract_template(self.case_id, "employee-a")

        groups = {group["order_number"]: group for group in result["template"]["groups"]}
        self.assertEqual(groups["PO-A"]["lines"][0]["values"]["customer_product_code"], "LOCKED")
        self.assertEqual(groups["PO-B"]["lines"][0]["values"]["customer_product_code"], "NEW-B")
        self.assertTrue(groups["PO-A"]["submitted"])

    def test_refresh_keeps_blank_po_tab_identifier_while_replacing_its_rows(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        saved = save_template(self.case_id, "employee-a", {"groups": [
            {
                "group_key": template["groups"][0]["group_key"],
                "order_number": "PO-A",
                "header": {"customer_order_number": "PO-A"},
                "lines": [{"values": {"line_no": "1", "customer_product_code": "OLD-A", "quantity": "1"}}],
            },
            {
                "group_key": "blank-po-group",
                "order_number": "",
                "header": {"customer_order_number": ""},
                "lines": [{"values": {"line_no": "2", "customer_product_code": "OLD-BLANK", "quantity": "1"}}],
            },
        ]})
        old_keys = {group["order_number"]: group["group_key"] for group in saved["groups"]}
        refreshed_lines = [
            _line_entry({"customer_order_number": "PO-A", "customer_product_code": "NEW-A", "quantity": "1"}, label="测试", reference="刷新", line_no=1),
            _line_entry({"customer_order_number": "", "customer_product_code": "NEW-BLANK", "quantity": "2"}, label="测试", reference="刷新", line_no=2),
        ]
        snapshot = {"tax_mode": "unknown", "target_field": "", "target_label": "", "by_line": {}, "mismatches": []}
        with patch(
            "fangzheng_web_app.order_entry_service._initial_template_data",
            return_value=({"customer_order_number": "PO-A"}, refreshed_lines),
        ), patch(
            "fangzheng_web_app.order_entry_service.review_case_template_prices", return_value=snapshot,
        ):
            result = reextract_template(self.case_id, "employee-a")

        groups = {group["order_number"]: group for group in result["template"]["groups"]}
        self.assertEqual(set(groups), {"PO-A", ""})
        self.assertEqual(groups["PO-A"]["group_key"], old_keys["PO-A"])
        self.assertEqual(groups[""]["group_key"], old_keys[""])
        self.assertEqual(groups[""]["display_order_number"], "暂无PO号-blank-po")
        self.assertEqual(groups[""]["lines"][0]["values"]["customer_product_code"], "NEW-BLANK")

    def test_price_review_is_calculated_only_on_initial_generation_and_refresh(self) -> None:
        snapshot = {"tax_mode": "unknown", "target_field": "", "target_label": "", "by_line": {}, "mismatches": []}
        with patch(
            "fangzheng_web_app.order_entry_service.review_case_template_prices", return_value=snapshot,
        ) as review:
            _case, template = get_or_create_template(self.case_id, "employee-a")
            get_or_create_template(self.case_id, "employee-a")
            save_template(self.case_id, "employee-a", {
                "header": template["header"],
                "lines": template["lines"],
            })
            reextract_template(self.case_id, "employee-a")

        self.assertEqual(review.call_count, 2)
        with db.db_cursor() as conn:
            header = json.loads(conn.execute(
                "SELECT header_json FROM order_entry_templates WHERE id=?", (template["id"],)
            ).fetchone()["header_json"])
        self.assertEqual(header["_price_review"], snapshot)

    def test_material_status_defaults_to_query_and_rejects_unknown_values(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        saved = save_template(self.case_id, "employee-a", {
            "header": {"order_type": "220", "bill_to_customer_code": "C001", "ledger": "KL01"},
            "lines": [{"values": {
                "line_no": "1", "customer_product_code": "CUST-1", "quantity": "20",
            }}],
        })
        self.assertEqual(saved["lines"][0]["values"]["material_status"], "查询")

        with self.assertRaisesRegex(ValueError, "料号状态只能选择"):
            save_template(self.case_id, "employee-a", {
                "header": {"order_type": "220", "bill_to_customer_code": "C001", "ledger": "KL01"},
                "lines": [{"values": {
                    "line_no": "1", "material_status": "删除",
                    "customer_product_code": "CUST-1", "quantity": "20",
                }}],
            })

    def test_line_number_tracks_customer_order_sequence_or_auto_increments(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        saved = save_template(self.case_id, "employee-a", {
            "header": {"order_type": "220", "bill_to_customer_code": "C001", "ledger": "KL01"},
            "lines": [
                {"values": {
                    "line_no": "1", "customer_order_seq": "10",
                    "customer_product_code": "CUST-10", "quantity": "20",
                }},
                {"values": {
                    "line_no": "2", "customer_order_seq": "",
                    "customer_product_code": "CUST-AUTO", "quantity": "30",
                }},
            ],
        })
        values_by_code = {
            line["values"]["customer_product_code"]: line["values"]
            for line in saved["lines"]
        }
        self.assertEqual(values_by_code["CUST-10"]["line_no"], "10")
        self.assertEqual(values_by_code["CUST-10"]["customer_order_seq"], "10")
        self.assertEqual(values_by_code["CUST-AUTO"]["line_no"], "2")
        self.assertEqual(values_by_code["CUST-AUTO"]["customer_order_seq"], "2")

    def test_customer_spec_match_recalculates_from_header_code_and_line_values(self) -> None:
        lines = [{
            "values": {
                "customer_spec": "NY2150_TG150_1.6MM",
                "customer_spec_match": "旧值",
                "product_type": "基板",
            },
            "sources": {"customer_spec_match": {"label": "旧来源", "reference": "旧规则"}},
        }]
        with patch(
            "fangzheng_web_app.order_entry_service.build_customer_spec_match",
            return_value="NY2150_*_1.6MM",
        ) as matcher:
            matched = _apply_customer_spec_matches(
                {"bill_to_customer_code": "C001"}, lines
            )

        self.assertEqual("NY2150_*_1.6MM", matched[0]["values"]["customer_spec_match"])
        self.assertEqual("查询", matched[0]["values"]["material_status"])
        self.assertEqual("客户规格对照表", matched[0]["sources"]["customer_spec_match"]["label"])
        matcher.assert_called_once_with("C001", "基板", "NY2150_TG150_1.6MM")

    def test_matched_customer_code_populates_header_and_spec_match_on_generation(self) -> None:
        with db.db_cursor() as conn:
            conn.execute(
                "UPDATE order_intake_cases SET customer_id=?,customer_match_status=? WHERE id=?",
                (42, "matched", self.case_id),
            )
        with patch(
            "fangzheng_web_app.order_entry_service.get_customer",
            return_value={"id": 42, "customer_code": "123036"},
        ), patch(
            "fangzheng_web_app.order_entry_service.build_customer_spec_match",
            return_value="自动客户规格匹配",
        ) as matcher:
            _case, template = get_or_create_template(self.case_id, "employee-a")

        self.assertEqual("123036", template["header"]["bill_to_customer_code"])
        self.assertEqual("123036", template["header"]["ship_to_customer_code"])
        self.assertTrue(template["lines"])
        self.assertTrue(all(
            line["values"]["customer_spec_match"] == "自动客户规格匹配"
            for line in template["lines"]
        ))
        self.assertTrue(all(
            line["sources"]["customer_spec_match"]["label"] == "客户规格对照表"
            for line in template["lines"]
        ))
        self.assertTrue(all(call.args[0] == "123036" for call in matcher.call_args_list))

    def test_existing_blank_template_is_backfilled_after_customer_match(self) -> None:
        _case, original = get_or_create_template(self.case_id, "employee-a")
        self.assertEqual("", original["header"]["bill_to_customer_code"])
        with db.db_cursor() as conn:
            conn.execute(
                "UPDATE order_intake_cases SET customer_id=?,customer_match_status=? WHERE id=?",
                (43, "matched", self.case_id),
            )
        with patch(
            "fangzheng_web_app.order_entry_service.get_customer",
            return_value={"id": 43, "customer_code": "104253"},
        ), patch(
            "fangzheng_web_app.order_entry_service.build_customer_spec_match",
            return_value="补齐后的规格匹配",
        ):
            _case, template = get_or_create_template(self.case_id, "employee-a")

        self.assertEqual("104253", template["header"]["bill_to_customer_code"])
        self.assertEqual("104253", template["header"]["ship_to_customer_code"])
        self.assertTrue(all(
            line["values"]["customer_spec_match"] == "补齐后的规格匹配"
            for line in template["lines"]
        ))

    def test_existing_bill_to_code_backfills_only_blank_ship_to_code(self) -> None:
        _case, original = get_or_create_template(self.case_id, "employee-a")
        with db.db_cursor() as conn:
            header = dict(original["header"])
            header["bill_to_customer_code"] = "123036"
            header["ship_to_customer_code"] = ""
            conn.execute(
                "UPDATE order_entry_templates SET header_json=? WHERE id=?",
                (json.dumps(header, ensure_ascii=False), original["id"]),
            )
            conn.execute(
                "UPDATE order_intake_cases SET customer_id=?,customer_match_status=? WHERE id=?",
                (44, "matched", self.case_id),
            )
        with patch(
            "fangzheng_web_app.order_entry_service.get_customer",
            return_value={"id": 44, "customer_code": "123036"},
        ):
            _case, template = get_or_create_template(self.case_id, "employee-a")

        self.assertEqual("123036", template["header"]["bill_to_customer_code"])
        self.assertEqual("123036", template["header"]["ship_to_customer_code"])

    def test_empty_customer_code_clears_customer_spec_match(self) -> None:
        lines = [{
            "values": {"customer_spec": "NY2150", "customer_spec_match": "旧值", "product_type": "基板"},
            "sources": {"customer_spec_match": {"label": "旧来源"}},
        }]
        with patch("fangzheng_web_app.order_entry_service.build_customer_spec_match") as matcher:
            matched = _apply_customer_spec_matches({"bill_to_customer_code": ""}, lines)

        self.assertEqual("", matched[0]["values"]["customer_spec_match"])
        self.assertNotIn("customer_spec_match", matched[0]["sources"])
        matcher.assert_not_called()

    def test_manual_customer_spec_match_is_preserved_for_same_context(self) -> None:
        lines = [{
            "values": {
                "customer_spec": "NY2150_TG150_1.6MM",
                "customer_spec_match": "人工调整后的结果",
                "product_type": "基板",
            },
            "sources": {"customer_spec_match": {
                "label": "人工修改",
                "reference": "录单模板手工修改",
                "context": '["C001","基板","NY2150_TG150_1.6MM"]',
            }},
        }]
        with patch("fangzheng_web_app.order_entry_service.build_customer_spec_match") as matcher:
            matched = _apply_customer_spec_matches(
                {"bill_to_customer_code": "C001"}, lines
            )

        self.assertEqual("人工调整后的结果", matched[0]["values"]["customer_spec_match"])
        self.assertEqual("人工修改", matched[0]["sources"]["customer_spec_match"]["label"])
        matcher.assert_not_called()

    def test_manual_customer_spec_match_recalculates_after_context_change(self) -> None:
        lines = [{
            "values": {
                "customer_spec": "NY2150_TG150_1.6MM",
                "customer_spec_match": "人工调整后的结果",
                "product_type": "PP",
            },
            "sources": {"customer_spec_match": {
                "label": "人工修改",
                "context": '["C001","基板","NY2150_TG150_1.6MM"]',
            }},
        }]
        with patch(
            "fangzheng_web_app.order_entry_service.build_customer_spec_match",
            return_value="重新自动匹配",
        ) as matcher:
            matched = _apply_customer_spec_matches(
                {"bill_to_customer_code": "C001"}, lines
            )

        self.assertEqual("重新自动匹配", matched[0]["values"]["customer_spec_match"])
        self.assertEqual("客户规格对照表", matched[0]["sources"]["customer_spec_match"]["label"])
        matcher.assert_called_once_with("C001", "PP", "NY2150_TG150_1.6MM")

    def test_save_template_ignores_manual_match_and_persists_generated_value(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        with patch(
            "fangzheng_web_app.order_entry_service.build_customer_spec_match",
            return_value="NY2150_*_1.6MM",
        ):
            saved = save_template(self.case_id, "employee-a", {
                "header": {"bill_to_customer_code": "C001"},
                "lines": [{"values": {
                    "customer_spec": "NY2150_TG150_1.6MM",
                    "customer_spec_match": "手工旧值",
                    "product_type": "基板",
                }}],
            })

        self.assertEqual(
            "NY2150_*_1.6MM",
            saved["lines"][0]["values"]["customer_spec_match"],
        )

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

    def test_excel_attachment_takes_priority_over_pdf_regardless_of_attachment_order(self) -> None:
        excel_path = Path(self.temp_dir.name) / "订单.xlsx"
        pdf_path = Path(self.temp_dir.name) / "订单.pdf"
        excel_path.touch()
        pdf_path.touch()
        excel_rows = [{"values": {"customer_product_code": "EXCEL-001"}}]
        case = {
            "customer_id": None,
            "attachments": [
                {"filename": "订单.pdf", "stored_path": str(pdf_path), "is_inline": 0},
                {"filename": "订单.xlsx", "stored_path": str(excel_path), "is_inline": 0},
            ],
        }

        with patch("fangzheng_web_app.order_entry_service._rows_from_excel", return_value=excel_rows) as excel_parser, patch(
            "fangzheng_web_app.order_entry_service._rows_from_pdf_or_image"
        ) as pdf_parser:
            rows = _attachment_rows(case)

        self.assertEqual(rows, excel_rows)
        excel_parser.assert_called_once()
        pdf_parser.assert_not_called()

    def test_pdf_attachment_is_used_when_excel_has_no_order_rows(self) -> None:
        excel_path = Path(self.temp_dir.name) / "空白.xlsx"
        pdf_path = Path(self.temp_dir.name) / "订单.pdf"
        excel_path.touch()
        pdf_path.touch()
        pdf_rows = [{"values": {"customer_product_code": "PDF-001"}}]
        case = {
            "customer_id": None,
            "attachments": [
                {"filename": "订单.pdf", "stored_path": str(pdf_path), "is_inline": 0},
                {"filename": "空白.xlsx", "stored_path": str(excel_path), "is_inline": 0},
            ],
        }

        with patch("fangzheng_web_app.order_entry_service._rows_from_excel", return_value=[]), patch(
            "fangzheng_web_app.order_entry_service._rows_from_pdf_or_image", return_value=pdf_rows
        ) as pdf_parser:
            rows = _attachment_rows(case)

        self.assertEqual(rows, pdf_rows)
        pdf_parser.assert_called_once()

    def test_customer_product_type_mapping_normalizes_board_material(self) -> None:
        mappings = [{
            "source_kind": "attachment_table", "target_field": "product_type",
            "source_label": "物料类别", "transform_type": "direct",
        }]
        values = _apply_customer_extraction_mappings(
            {"product_type": "PP"}, {"物料类别": "板材"}, mappings,
        )

        self.assertEqual(values["product_type"], "基板")

    def test_board_size_with_foil_type_is_classified_for_direct_order_rows(self) -> None:
        line = _line_entry(
            {"customer_spec": 'NY6200 0.089mm 1/1 37"*49"(1067*2)(RTF)(有卤素)'},
            label="邮件正文", reference="第 1 行",
        )

        self.assertEqual(line["values"]["product_type"], "基板")

    def test_resin_content_spec_is_classified_for_direct_order_rows(self) -> None:
        line = _line_entry(
            {"customer_spec": 'NY-A2P 1080 RC68% 20.6"x24.6"有卤 CAF（汽车板）'},
            label="邮件正文", reference="第 1 行",
        )

        self.assertEqual(line["values"]["product_type"], "PP")

    def test_html_mail_table_uses_shared_domestic_template_mapping(self) -> None:
        body_html = """
        <table><tr><th>PO单号</th><th>PO项目号</th><th>物料编码</th><th>物料名称</th>
        <th>物料规格</th><th>数量</th><th>单位</th><th>单价</th><th>交期</th></tr>
        <tr><td>MZPOM12608190027</td><td>2713183</td><td>10101008248</td><td>FR-4</td>
        <td>南亚新材料 NY6180L 0.127 mm</td><td>20</td><td>张</td><td>215.6814</td><td>2026-08-28</td></tr>
        </table>
        """
        with patch("fangzheng_web_app.order_entry_service.get_enabled_extraction_maps", return_value=[]):
            header, rows = _initial_template_data({
                "id": 999,
                "body_html": body_html,
                "body_text": "",
                "attachments": [],
                "detected_fields": {},
                "customer_id": None,
            })
        self.assertEqual(header["customer_order_number"], "MZPOM12608190027")
        self.assertEqual(rows[0]["values"]["customer_product_code"], "10101008248")
        self.assertEqual(rows[0]["values"]["quantity"], "20")
        self.assertEqual(rows[0]["values"]["product_type"], "基板")
        self.assertEqual(rows[0]["sources"]["quantity"]["label"], "邮件正文表格")

    def test_html_mail_pp_fractional_roll_quantity_is_converted_to_metres(self) -> None:
        body_html = """
        <table><tr><th>PO单号</th><th>PO项目号</th><th>物料编码</th><th>物料名称</th>
        <th>物料规格</th><th>数量</th><th>单位</th><th>单价</th><th>交期</th></tr>
        <tr><td>PO-20260910</td><td>1</td><td>LA-001</td><td></td>
        <td>NY-A2P 1080 RC66% 49.5\"有卤 CAF 300M/卷（汽车板）</td>
        <td>0.7</td><td></td><td>27.72</td><td>2026-09-15</td></tr>
        </table>
        """
        with patch("fangzheng_web_app.order_entry_service.get_enabled_extraction_maps", return_value=[]):
            _header, rows = _initial_template_data({
                "id": 1003, "body_html": body_html, "body_text": "", "attachments": [],
                "detected_fields": {}, "customer_id": None,
            })

        self.assertEqual(rows[0]["values"]["product_type"], "PP")
        self.assertEqual(rows[0]["values"]["quantity"], "210")
        self.assertEqual(rows[0]["values"]["remark"], "0.7卷&")

    def test_html_mail_table_uses_customer_order_table_mappings(self) -> None:
        body_html = """
        <table><tr><th>PO单号</th><th>项次</th><th>客户料号</th><th>数量</th>
        <th>要求交货期</th><th>订单类型</th></tr>
        <tr><td>PO-20260909</td><td>1</td><td>CUST-001</td><td>20</td>
        <td>2026/09/25</td><td>加急订单</td></tr>
        </table>
        """
        mappings = [
            {
                "source_kind": "attachment_table", "target_field": "delivery_date",
                "source_label": "要求交货日期", "transform_type": "direct",
            },
            {
                "source_kind": "attachment_table", "target_field": "remark",
                "source_label": "订单类型", "transform_type": "direct",
            },
        ]
        with patch("fangzheng_web_app.order_entry_service.get_enabled_extraction_maps", return_value=mappings):
            _header, rows = _initial_template_data({
                "id": 1002,
                "body_html": body_html,
                "body_text": "",
                "attachments": [],
                "detected_fields": {},
                "customer_id": 1,
            })

        self.assertEqual("2026-09-25", rows[0]["values"]["delivery_date"])
        self.assertEqual("&加急订单", rows[0]["values"]["remark"])

    def test_plain_text_demand_delivery_rows_extract_yearless_dates(self) -> None:
        body_text = """
        下单日期 类别 供应商 编码 描述 数量 单位 需求交期 供应商交期复期 料号
        8/29 板材 南亚新材 AA1130050110010002 NY6300S 0.050mm 1/1 43\"x49\" Halogen-free RTF2 1x1027 25 PIE 9月25日 R0O30A520298A
        8/29 PP 南亚新材 LA0911078650191001 NY6300SP 1078 RC65% 21.7\"x24.5\" Halogen-free CAF 390 PIE 9月25日 R0O30A520298A
        """
        header, rows = _initial_template_data({
            "id": 1001,
            "body_html": "",
            "body_text": body_text,
            "attachments": [],
            "detected_fields": {},
            "customer_id": None,
            "received_at": "2026-08-29 09:00:00",
        })

        self.assertTrue(header["customer_order_number"].startswith("暂无PO号-"))
        self.assertEqual(2, len(rows))
        self.assertEqual("AA1130050110010002", rows[0]["values"]["customer_product_code"])
        self.assertEqual("2026-09-25", rows[0]["values"]["delivery_date"])
        self.assertEqual("PP", rows[1]["values"]["product_type"])
        self.assertEqual("2026-09-25", rows[1]["values"]["delivery_date"])

    def test_missing_customer_order_number_receives_uuid_placeholder(self) -> None:
        with patch("fangzheng_web_app.order_entry_service.uuid.uuid4", return_value="12345678-1234-5678-9abc-def012345678"):
            header, _rows = _initial_template_data({
                "id": 1000,
                "body_html": "",
                "body_text": "没有可识别的订单号",
                "attachments": [],
                "detected_fields": {},
                "customer_id": None,
            })

        self.assertEqual(
            "暂无PO号-12345678-1234-5678-9abc-def012345678",
            header["customer_order_number"],
        )

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
        self.assertEqual(line["values"]["remark"], "客户加急；2卷")

    def test_fractional_pp_quantity_converts_from_rolls_without_unit(self) -> None:
        line = _line_entry(
            {"customer_spec": 'NY-A2P 1080 RC66% 49.5"有卤 CAF 300M/卷', "quantity": "0.7"},
            label="附件：PP订单.xlsx", reference="订单 第 2 行", quantity_unit="",
        )

        self.assertEqual(line["values"]["product_type"], "PP")
        self.assertEqual(line["values"]["quantity"], "210")
        self.assertEqual(line["values"]["remark"], "0.7卷")

    def test_pp_sheet_and_no_roll_length_keep_customer_quantity(self) -> None:
        small_piece = _line_entry(
            {"customer_spec": "PP 1080 300m", "quantity": "30"},
            label="附件：PP订单.xlsx", reference="订单 第 2 行", quantity_unit="张",
        )
        unknown = _line_entry(
            {"customer_spec": "PP 1080 300m", "quantity": "2"},
            label="附件：PP订单.xlsx", reference="订单 第 3 行", quantity_unit="",
        )
        self.assertEqual(small_piece["values"]["quantity"], "30")
        self.assertEqual(small_piece["values"]["remark"], "")
        self.assertEqual(unknown["values"]["quantity"], "2")
        self.assertEqual(unknown["values"]["remark"], "")

        no_roll_length = _line_entry(
            {"customer_spec": 'NY-A2P 2116 RC60% 18.62"x16.42"有卤 CAF', "quantity": "0.1"},
            label="附件：PP订单.xlsx", reference="订单 第 4 行", quantity_unit="",
        )
        self.assertEqual(no_roll_length["values"]["quantity"], "0.1")

    def test_template_remark_keeps_generated_meter_note_before_source_note(self) -> None:
        rows = _merge_initial_rows([
            _line_entry(
                {"customer_product_code": "PP-01", "customer_spec": "PP 1080", "remark": "客户加急；0.1卷"},
                label="附件：PP订单.xlsx", reference="订单第 2 行",
            ),
            _line_entry(
                {"customer_product_code": "CCL-01", "customer_spec": "FR-4", "remark": "订单说明"},
                label="附件：基板订单.xlsx", reference="订单第 3 行",
            ),
        ])

        self.assertEqual(rows[0]["values"]["remark"], "0.1卷&客户加急")
        self.assertEqual(rows[1]["values"]["remark"], "&订单说明")

    def test_template_remark_moves_roll_quantity_before_existing_source_separator(self) -> None:
        rows = _merge_initial_rows([_line_entry(
            {"customer_product_code": "PP-01", "remark": "&客户加急；0.7卷"},
            label="附件：PP订单.xlsx", reference="订单第 2 行",
        )])

        self.assertEqual(rows[0]["values"]["remark"], "0.7卷&客户加急")

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
        self.assertEqual(first["values"]["customer_product_code"], "10601001676")
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

    def test_pdf_attachment_reuses_pdf_domestic_mapping(self) -> None:
        document = {
            "mapped_detail_rows": [{
                "original": {"物料描述": "PP 1080 300M/卷"},
                "standard": {"物料编码": "CUST-PP", "物料名称": "半固化片"},
            }],
            "factory_import": {
                "main_values": ["", "", "", "", "", "", "", "PO-001"],
                "rows": [{
                    FACTORY_DETAIL_HEADERS[0]: "9",
                    FACTORY_DETAIL_HEADERS[3]: "CUST-PP",
                    FACTORY_DETAIL_HEADERS[4]: "2026-08-30",
                    FACTORY_DETAIL_HEADERS[5]: "600",
                    FACTORY_DETAIL_HEADERS[6]: "10",
                    FACTORY_DETAIL_HEADERS[7]: "11.3",
                    FACTORY_DETAIL_HEADERS[11]: "加急",
                }],
            },
        }
        with patch("fangzheng_web_app.order_entry_service.recognize_purchase_order_document", return_value=document), patch("fangzheng_web_app.order_entry_service.project_factory_document"):
            rows = _rows_from_pdf_or_image(Path("/tmp/PO-001.pdf"), "PO-001.pdf")
        self.assertEqual(rows[0]["values"]["customer_product_code"], "CUST-PP")
        self.assertEqual(rows[0]["values"]["quantity"], "600")
        self.assertEqual(rows[0]["values"]["price_before_tax"], "10")
        self.assertEqual(rows[0]["extracted_header"]["customer_order_number"], "PO-001")
        self.assertEqual(rows[0]["values"]["product_type"], "PP")
        self.assertEqual(rows[0]["sources"]["customer_product_code"]["label"], "附件：PO-001.pdf")

    def test_pdf_attachment_keeps_explicit_blank_po_for_its_own_group(self) -> None:
        document = {
            "mapped_detail_rows": [
                {
                    "original": {"PO号": "PO-GM260008046", "物料描述": "PP 1080 300M/卷"},
                    "standard": {"物料编码": "CUST-A", "物料名称": "半固化片"},
                },
                {
                    "original": {"PO号": "", "物料描述": "PP 2116 300M/卷"},
                    "standard": {"物料编码": "CUST-B", "物料名称": "半固化片"},
                },
            ],
            "factory_import": {
                "main_values": ["", "", "", "", "", "", "", "PO-GM260008046"],
                "rows": [
                    {
                        FACTORY_DETAIL_HEADERS[0]: "1", FACTORY_DETAIL_HEADERS[3]: "CUST-A",
                        FACTORY_DETAIL_HEADERS[4]: "2026-09-25", FACTORY_DETAIL_HEADERS[5]: "600",
                    },
                    {
                        FACTORY_DETAIL_HEADERS[0]: "2", FACTORY_DETAIL_HEADERS[3]: "CUST-B",
                        FACTORY_DETAIL_HEADERS[4]: "2026-09-25", FACTORY_DETAIL_HEADERS[5]: "300",
                    },
                ],
            },
        }
        with patch("fangzheng_web_app.order_entry_service.recognize_purchase_order_document", return_value=document), patch(
            "fangzheng_web_app.order_entry_service.project_factory_document"
        ):
            rows = _rows_from_pdf_or_image(Path("/tmp/PO-GM260008046.pdf"), "PO-GM260008046.pdf")

        groups = _initial_order_groups({"customer_order_number": "PO-GM260008046"}, rows)
        self.assertEqual([row["values"]["customer_order_number"] for row in rows], ["PO-GM260008046", ""])
        self.assertEqual([group["order_number"] for group in groups], ["PO-GM260008046", ""])
        self.assertEqual([len(group["lines"]) for group in groups], [1, 1])

    def test_shared_document_keeps_blank_english_po_cell_for_its_own_group(self) -> None:
        document = {
            "mapped_detail_rows": [
                {"original": {"PO": "PO-GM260008046"}, "standard": {}},
                {"original": {"PO": ""}, "standard": {}},
            ],
            "factory_import": {
                "main_values": ["", "", "", "", "", "", "", "PO-GM260008046"],
                "rows": [
                    {FACTORY_DETAIL_HEADERS[0]: "1", FACTORY_DETAIL_HEADERS[5]: "600"},
                    {FACTORY_DETAIL_HEADERS[0]: "2", FACTORY_DETAIL_HEADERS[5]: "300"},
                ],
            },
        }
        with patch("fangzheng_web_app.order_entry_service.project_factory_document"):
            from fangzheng_web_app.order_entry_service import _rows_from_shared_purchase_document
            rows = _rows_from_shared_purchase_document(document, label="邮件正文表格", reference_prefix="表格明细")

        groups = _initial_order_groups({"customer_order_number": "PO-GM260008046"}, rows)
        self.assertEqual([row["values"]["customer_order_number"] for row in rows], ["PO-GM260008046", ""])
        self.assertEqual([group["order_number"] for group in groups], ["PO-GM260008046", ""])

    def test_shared_document_does_not_convert_already_projected_roll_metres_twice(self) -> None:
        document = {
            "header_info": {"订单号": "PO-GM260008046"},
            "mapped_detail_rows": [{
                "original": {
                    "项目": "36", "物料编码": "LAN31AZP73003",
                    "物料描述": 'PP NY6300P(C) 1078 RC73% 49.5"*300M/Roll',
                    "数量": "4", "单位": "卷", "未税单价": "25078.69",
                    "未税金额": "100314.76", "含税单价": "28338.9197", "含税金额": "113355.68",
                },
                "standard": {
                    "序号": "36", "物料编码": "LAN31AZP73003",
                    "物料名称": 'PP NY6300P(C) 1078 RC73% 49.5"*300M/Roll',
                    "说明": 'PP NY6300P(C) 1078 RC73% 49.5"*300M/Roll',
                    "数量": "4", "单位": "卷", "未税单价": "25078.69",
                    "未税金额": "100314.76", "含税单价": "28338.9197", "含税金额": "113355.68",
                    "交货日期": "2026-10-10",
                },
            }],
        }
        from fangzheng_web_app.order_entry_service import _rows_from_shared_purchase_document

        rows = _rows_from_shared_purchase_document(document, label="附件", reference_prefix="附件")

        self.assertEqual(rows[0]["values"]["quantity"], "1200")
        self.assertEqual(rows[0]["values"]["price_before_tax"], "83.59563333333333333333333333")
        self.assertEqual(rows[0]["values"]["unit_price"], "94.46306566666666666666666667")
        self.assertEqual(rows[0]["values"]["amount_before_tax"], "100314.76")
        self.assertEqual(rows[0]["values"]["amount_with_tax"], "113355.68")
        self.assertIn("4卷", rows[0]["values"]["remark"])

    def test_entry_template_exposes_delete_selected_rows_control(self) -> None:
        template = (Path(__file__).resolve().parents[1] / "templates" / "order_automation_entry_template.html").read_text(encoding="utf-8")
        self.assertIn('id="deleteSelectedLines"', template)
        self.assertIn("删除所选", template)
        self.assertIn('<input class="oe-line-select"', template)
        self.assertIn('<span data-row-number></span>', template)

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
        self.assertEqual(template["current_version"], 1)
        self.assertTrue(all(line["values"]["product_name"] == "" for line in template["lines"]))
        with db.db_cursor() as conn:
            versions = conn.execute(
                "SELECT version_number,lines_json FROM order_entry_template_versions ORDER BY version_number"
            ).fetchall()
        self.assertEqual([row["version_number"] for row in versions], [1])
        self.assertIn("历史错误品名", versions[0]["lines_json"])
