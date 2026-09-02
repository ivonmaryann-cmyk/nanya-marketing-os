from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import db
from fangzheng_web_app.mail_transcode_agent import mail_store
from fangzheng_web_app.order_entry_service import get_or_create_template, reextract_template, save_template
from fangzheng_web_app.order_interface_service import (
    build_domestic_order_entry,
    build_domestic_order_entry_mock,
    build_material_creation,
    build_material_query,
    build_material_query_mock,
    get_interface_config,
    is_domestic_order_entry_completed,
    get_order_detail_records,
    get_material_resolution_states,
    list_nyeos_order_numbers,
    list_interface_configs,
    process_material_created_callback,
    _real_material_request_item,
    _transcode_product_name,
    select_material_candidate,
    save_interface_config,
    test_interface_config,
    validate_domestic_order_entry,
)
from fangzheng_web_app.order_intake_service import bootstrap_cases, list_cases


class OrderInterfaceMaintenanceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_patch = patch.object(db, "DATABASE_PATH", Path(self.temp_dir.name) / "interfaces.sqlite3")
        self.db_patch.start()
        db.init_db()
        self.account_id = mail_store.create_or_update_account(
            "orders@example.com", owner_employee_id="employee-a", auth_code="auth-code"
        )
        mail_store.upsert_message(
            self.account_id, folder="INBOX", uid="interface-1", message_id="<interface-1@example.com>",
            subject="采购订单", sender="buyer@example.com", sent_at="2026-08-19 09:00:00",
            received_at="2026-08-19 09:00:00", body_html="", body_text="PO-20260819",
            eml_path="", is_order=1,
        )
        bootstrap_cases("employee-a", self.account_id)
        self.case_id = list_cases("employee-a", "2026-08-19", "new_order", self.account_id)[0]["id"]

    def tearDown(self) -> None:
        self.db_patch.stop()
        self.temp_dir.cleanup()

    def test_default_configs_are_seeded_and_versioned(self) -> None:
        configs = list_interface_configs()
        self.assertEqual({item["interface_key"] for item in configs}, {"material_batch_query", "domestic_order_entry"})
        material = get_interface_config("material_batch_query")
        self.assertEqual(material["mode"], "mock")
        self.assertEqual(
            material["endpoint_url"],
            "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/pe/queryMaterial",
        )
        self.assertEqual(material["request_mapping"]["materialInfoList[].categoryCode"],
                         "模板明细.产品类型（PP=698，基板=718，必填）")
        self.assertEqual(material["request_mapping"]["materialInfoList[].customerSpec"],
                         "新建料号弹窗.客户规格匹配（仅点击新建料号时传）")
        self.assertEqual(material["request_mapping"]["materialInfoList[].adhesiveCode"],
                         "新建料号弹窗.胶系编码（仅点击新建料号时传）")
        self.assertEqual(material["request_mapping"]["materialInfoList[].customerSpecOld"],
                         "模板明细.客户规格（选填）")
        self.assertEqual(material["response_mapping"]["hitMaterialList[].peag01"],
                         "料号查询建议.产品编号")
        self.assertEqual(material["response_mapping"]["hitMaterialList[].peag09"],
                         "料号查询建议.品名")
        self.assertTrue(any("按运行模式执行" in note for note in material["maintenance_notes"]))
        domestic = get_interface_config("domestic_order_entry")
        self.assertEqual(
            domestic["request_mapping"]["sctoDataList[].materialCode"],
            "模板明细.客户产品编号（必填）",
        )
        self.assertEqual(
            domestic["request_mapping"]["sctoDataList[].custOrderId"],
            "模板表头.客户订单号（必填）",
        )
        self.assertEqual(
            domestic["request_mapping"]["sctoDataList[].spec"],
            "模板明细.客户规格（选填）",
        )
        saved = save_interface_config("material_batch_query", {
            "display_name": material["display_name"],
            "description": material["description"],
            "enabled": "1",
            "mode": "mock",
            "method": "POST",
            "endpoint_url": material["endpoint_url"],
            "request_mapping": "{}",
            "response_mapping": "{}",
            "mock_scenarios": "{}",
        }, "23582")
        self.assertEqual(saved["config_version"], 2)
        self.assertEqual(saved["endpoint_url"], material["endpoint_url"])

    def test_untouched_legacy_material_config_is_upgraded_without_enabling_real_calls(self) -> None:
        material = get_interface_config("material_batch_query")
        with db.db_cursor() as conn:
            conn.execute(
                """UPDATE order_interface_configs
                   SET description=?,base_url=?,port=?,timeout_seconds=?,request_mapping_json=?,
                       response_mapping_json=?,config_version=1
                   WHERE interface_key='material_batch_query'""",
                (
                    "按订单明细批量查询料号；返回结果仅作为建议，不覆盖人工填写内容。",
                    "https://mock.nouya.local/material/batch-query", 443, 8,
                    '{"items[].line_no":"模板明细.项次","items[].customer_part_no":"模板明细.客户产品编号","items[].customer_spec":"模板明细.客户规格"}',
                    '{"items[].factory_part_no":"料号查询建议.产品编号","items[].product_name":"料号查询建议.品名","items[].matched_spec":"料号查询建议.匹配规格","items[].status":"接口交互记录.状态","items[].message":"接口交互记录.提示"}',
                ),
            )
        upgraded = get_interface_config("material_batch_query")
        self.assertEqual(upgraded["config_version"], 2)
        self.assertEqual(upgraded["mode"], "mock")
        self.assertEqual(upgraded["endpoint_url"], material["endpoint_url"])
        with db.db_cursor() as conn:
            versions = conn.execute(
                "SELECT COUNT(*) AS count FROM order_interface_config_versions"
            ).fetchone()["count"]
        self.assertEqual(versions, 1)

    def test_legacy_domestic_config_keeps_selected_real_mode_when_upgraded(self) -> None:
        get_interface_config("domestic_order_entry")
        with db.db_cursor() as conn:
            conn.execute(
                """UPDATE order_interface_configs
                   SET display_name='内销录单',description=?,mode='real',base_url=?,port=443,
                       timeout_seconds=10,request_mapping_json=?,response_mapping_json=?,config_version=1
                   WHERE interface_key='domestic_order_entry'""",
                (
                    "人工确认订单内容后提交内销录单；当前先维护 Mock 配置。",
                    "https://mock.nouya.local/sales/internal-entry",
                    '{"header":"内销模板.表头","items":"内销模板.明细行"}',
                    '{"status":"接口交互记录.状态","message":"接口交互记录.提示"}',
                ),
            )
        upgraded = get_interface_config("domestic_order_entry")
        self.assertEqual(upgraded["display_name"], "生成订单")
        self.assertEqual(upgraded["mode"], "real")
        self.assertEqual(
            upgraded["endpoint_url"],
            "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/sc/saveSctoAndGenerateOrder",
        )
        self.assertEqual(upgraded["config_version"], 2)

    def test_mock_interface_test_uses_the_full_endpoint_url(self) -> None:
        result = test_interface_config({
            "interface_key": "material_batch_query",
            "mode": "mock",
            "method": "POST",
            "endpoint_url": "https://mock.nouya.local/material/batch-query",
        })
        self.assertTrue(result["ok"])
        self.assertEqual(result["status_code"], 200)
        self.assertEqual(result["request"]["materialInfoList"][0]["categoryCode"], "718")
        self.assertEqual(result["request"]["materialInfoList"][0]["customerMaterialNo"], "")
        self.assertEqual(result["response"]["code"], 200)

        domestic = test_interface_config({
            "interface_key": "domestic_order_entry", "mode": "mock", "method": "POST",
            "endpoint_url": "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/sc/saveSctoAndGenerateOrder",
        })
        self.assertIn("sctoDataList", domestic["request"])
        self.assertNotIn("header", domestic["request"])
        self.assertEqual(domestic["response"]["code"], 200)

    def test_template_events_and_mock_query_do_not_overwrite_manual_fields(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "product_code": "MANUAL-001",
                "customer_product_code": "CUST-002", "customer_spec": "FR-4 1.6",
                "quantity": "20",
            }}],
        })
        result = build_material_query_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual(len(result["items"]), 1)
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertTrue(any(item["event_type"] == "template_extracted" for item in details["events"]))
        self.assertTrue(any(item["event_type"] == "template_saved" for item in details["events"]))
        self.assertTrue(any(item["event_type"] == "material_query_mock" for item in details["events"]))
        self.assertEqual(len(details["calls"]), 1)
        self.assertTrue(details["events"][0]["trace_id"].startswith("E-"))
        self.assertRegex(details["events"][0]["occurred_at"], r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
        self.assertTrue(details["calls"][0]["trace_id"].startswith("I-"))
        self.assertEqual(details["calls"][0]["interface_label"], "批量料号查询")
        self.assertEqual(
            details["calls"][0]["endpoint_url"],
            "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/pe/queryMaterial",
        )
        self.assertEqual(details["calls"][0]["method"], "POST")
        with db.db_cursor() as conn:
            value = conn.execute(
                "SELECT values_json FROM order_entry_template_lines"
            ).fetchone()["values_json"]
        self.assertIn("MANUAL-001", value)

    def test_mock_error_and_timeout_are_preserved_as_failed_calls(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        failed = build_material_query_mock(
            self.case_id, "employee-a", "employee-a", scenario="business_error"
        )
        timeout = build_material_query_mock(
            self.case_id, "employee-a", "employee-a", scenario="timeout"
        )
        self.assertEqual(failed["status"], "failed")
        self.assertEqual(timeout["status"], "failed")
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertEqual([item["status"] for item in details["calls"]], ["failed", "failed"])

    def test_domestic_entry_mock_records_the_submission_without_changing_template(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "product_code": "MANUAL-001", "customer_product_code": "CUST-001",
                "customer_spec": "FR-4 1.6", "quantity": "20",
            }}],
        })
        result = build_domestic_order_entry_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual(result["status"], "success")
        self.assertTrue(result["entry_no"].startswith("MOCK-SO-"))
        self.assertTrue(is_domestic_order_entry_completed(self.case_id, "employee-a"))
        with db.db_cursor() as conn:
            case = conn.execute(
                "SELECT status,workflow_stage,erp_prepare_status,completed_at FROM order_intake_cases WHERE id=?",
                (self.case_id,),
            ).fetchone()
        self.assertEqual(
            (case["status"], case["workflow_stage"], case["erp_prepare_status"]),
            ("archived", "completed", "submitted"),
        )
        self.assertTrue(case["completed_at"])
        with self.assertRaisesRegex(ValueError, "不能重复提交"):
            build_domestic_order_entry_mock(self.case_id, "employee-a", "employee-a")
        with self.assertRaisesRegex(ValueError, "不能再次请求料号查询接口"):
            build_material_query_mock(self.case_id, "employee-a", "employee-a")
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertEqual(details["calls"][0]["interface_key"], "domestic_order_entry")
        self.assertTrue(any(item["event_type"] == "domestic_order_entry_mock" for item in details["events"]))

    def test_real_domestic_entry_posts_scto_payload_and_completes_only_on_success(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {
                "order_type": "220", "bill_to_customer_code": "103814",
                "customer_order_number": "PO20260824002", "ledger": "KL01", "tax_type": "1",
            },
            "lines": [{"values": {
                "line_no": "1", "product_code": "6900000008", "product_name": "厂内品名",
                "customer_product_code": "CUST-001", "customer_spec": "原始客户规格",
                "customer_spec_match": "客户匹配规格",
                "quantity": "100", "unit_price": "11.3", "delivery_date": "2026-09-01",
                "customer_order_seq": "10", "remark": "测试行",
            }}],
        })
        config = get_interface_config("domestic_order_entry")
        save_interface_config("domestic_order_entry", {
            "display_name": config["display_name"], "description": config["description"],
            "mode": "real", "method": "POST", "endpoint_url": config["endpoint_url"],
            "request_mapping": json.dumps(config["request_mapping"], ensure_ascii=False),
            "response_mapping": json.dumps(config["response_mapping"], ensure_ascii=False),
            "mock_scenarios": json.dumps(config["mock_scenarios"], ensure_ascii=False),
        }, "employee-a")
        response = {
            "msg": "订单生成完成", "code": 200,
            "data": {
                "data": [{
                    "orderNumber": "PO20260824002", "sctaCode": "SA2608250002",
                    "message": "", "lineCount": 1, "status": "success",
                }],
                "failCount": 0, "successCount": 1,
            },
        }
        with patch(
            "fangzheng_web_app.order_interface_service._post_json_endpoint",
            return_value=(200, response, 26),
        ) as request_mock:
            result = build_domestic_order_entry(self.case_id, "employee-a", "employee-a")
        payload = request_mock.call_args.args[1]["sctoDataList"][0]
        self.assertEqual(result["mode"], "real")
        self.assertEqual(result["entry_no"], "SA2608250002")
        self.assertEqual(payload["customerCode"], "103814")
        self.assertEqual(payload["orderType"], "220")
        self.assertEqual(payload["operator"], "employee-a")
        self.assertEqual(payload["materialCode"], "CUST-001")
        self.assertEqual(payload["orderNumber"], "PO20260824002")
        self.assertEqual(payload["custOrderId"], "PO20260824002")
        self.assertEqual(payload["lineNumber"], "10")
        self.assertEqual(payload["lineId"], "10")
        self.assertEqual(payload["spec"], "原始客户规格")
        self.assertEqual(payload["demandDate"], "2026-09-01")
        self.assertEqual(payload["taxPrice"], "11.3")
        with db.db_cursor() as conn:
            call = conn.execute(
                "SELECT is_mock,status,http_status FROM order_interface_call_logs ORDER BY id DESC LIMIT 1"
            ).fetchone()
            case = conn.execute(
                "SELECT status,workflow_stage,erp_prepare_status FROM order_intake_cases WHERE id=?",
                (self.case_id,),
            ).fetchone()
        self.assertEqual((call["is_mock"], call["status"], call["http_status"]), (0, "success", 200))
        self.assertEqual(
            (case["status"], case["workflow_stage"], case["erp_prepare_status"]),
            ("archived", "completed", "submitted"),
        )
        self.assertEqual(
            list_nyeos_order_numbers([self.case_id], "employee-a"),
            {self.case_id: "SA2608250002"},
        )

    def test_real_domestic_entry_business_failure_is_logged_without_completing_case(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {
                "order_type": "220", "bill_to_customer_code": "103814",
                "customer_order_number": "PO-FAIL", "ledger": "KL01",
            },
            "lines": [{"values": {
                "line_no": "1", "product_code": "6900000008", "customer_product_code": "CUST-001",
                "quantity": "100", "price_before_tax": "10", "delivery_date": "2026-09-01",
            }}],
        })
        config = get_interface_config("domestic_order_entry")
        save_interface_config("domestic_order_entry", {
            "display_name": config["display_name"], "description": config["description"],
            "mode": "real", "method": "POST", "endpoint_url": config["endpoint_url"],
            "request_mapping": json.dumps(config["request_mapping"], ensure_ascii=False),
            "response_mapping": json.dumps(config["response_mapping"], ensure_ascii=False),
            "mock_scenarios": json.dumps(config["mock_scenarios"], ensure_ascii=False),
        }, "employee-a")
        response = {
            "msg": "订单生成完成", "code": 200,
            "data": {
                "data": [{"orderNumber": "PO-FAIL", "sctaCode": "", "message": "料号不存在", "lineCount": 0, "status": "fail"}],
                "failCount": 1, "successCount": 0,
            },
        }
        with patch(
            "fangzheng_web_app.order_interface_service._post_json_endpoint",
            return_value=(200, response, 20),
        ):
            with self.assertRaisesRegex(ValueError, "料号不存在"):
                build_domestic_order_entry(self.case_id, "employee-a", "employee-a")
        with db.db_cursor() as conn:
            call = conn.execute(
                "SELECT is_mock,status,error_message FROM order_interface_call_logs ORDER BY id DESC LIMIT 1"
            ).fetchone()
            case = conn.execute("SELECT status FROM order_intake_cases WHERE id=?", (self.case_id,)).fetchone()
        self.assertEqual((call["is_mock"], call["status"]), (0, "failed"))
        self.assertIn("料号不存在", call["error_message"])
        self.assertNotEqual(case["status"], "archived")

    def test_material_mock_backfills_then_callback_completes_creation(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [
                {"values": {"line_no": "1", "customer_product_code": "CUST-001", "customer_spec": "规格一", "customer_spec_match": "匹配一", "quantity": "1"}},
                {"values": {"line_no": "2", "customer_product_code": "CUST-002", "customer_spec": "规格二", "customer_spec_match": "匹配二", "quantity": "1"}},
                {"values": {"line_no": "3", "customer_product_code": "CUST-003", "customer_spec": "规格三", "customer_spec_match": "匹配三", "quantity": "1"}},
            ],
        })
        result = build_material_query_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual([item["status"] for item in result["items"]], ["matched", "creating", "failed"])
        states = get_material_resolution_states(self.case_id, "employee-a")["items"]
        self.assertEqual([item["status"] for item in states], ["resolved", "waiting_callback", "failed"])
        self.assertEqual(len(states[0]["candidates"]), 2)
        self.assertEqual(states[0]["candidates"][0]["product_code"], result["items"][0]["factory_part_no"])
        self.assertEqual(states[0]["candidates"][1]["product_name"], result["items"][0]["candidates"][1]["product_name"])
        self.assertTrue(validate_domestic_order_entry(self.case_id, "employee-a"))
        waiting = next(item for item in states if item["status"] == "waiting_callback")
        process_material_created_callback(
            waiting["correlation_id"], product_code="CALLBACK-002", product_name="回调品名", source="test",
        )
        states = get_material_resolution_states(self.case_id, "employee-a")["items"]
        self.assertEqual(states[1]["status"], "resolved")
        with db.db_cursor() as conn:
            values = conn.execute(
                "SELECT values_json FROM order_entry_template_lines WHERE line_no=2"
            ).fetchone()["values_json"]
        self.assertIn("CALLBACK-002", values)

    def test_multiple_candidates_require_explicit_selection_and_persist_pair(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "customer_product_code": "CUST-001", "customer_spec": "规格一",
                "quantity": "1",
            }}],
        })
        build_material_query_mock(self.case_id, "employee-a", "employee-a")
        state = get_material_resolution_states(self.case_id, "employee-a")["items"][0]
        self.assertEqual(state["candidate_count"], 2)
        self.assertTrue(state["selection_required"])
        self.assertIsNone(state["selected_candidate"])
        with db.db_cursor() as conn:
            values = json.loads(
                conn.execute(
                    "SELECT values_json FROM order_entry_template_lines WHERE line_no=1"
                ).fetchone()["values_json"]
            )
        self.assertEqual(values["product_code"], state["candidates"][0]["product_code"])
        self.assertEqual(values["product_name"], state["candidates"][0]["product_name"])
        self.assertFalse(validate_domestic_order_entry(self.case_id, "employee-a"))
        with self.assertRaisesRegex(ValueError, "候选列表"):
            select_material_candidate(
                self.case_id, "employee-a", 1, product_code="INVALID", product_name="错误品名"
            )

        selected = select_material_candidate(
            self.case_id,
            "employee-a",
            1,
            product_code=state["candidates"][1]["product_code"],
            product_name=state["candidates"][1]["product_name"],
        )
        self.assertEqual(selected["selected_candidate"], state["candidates"][1])
        confirmed = get_material_resolution_states(self.case_id, "employee-a")["items"][0]
        self.assertFalse(confirmed["selection_required"])
        self.assertEqual(confirmed["selected_candidate"], state["candidates"][1])
        self.assertFalse(validate_domestic_order_entry(self.case_id, "employee-a"))
        details = get_order_detail_records(self.case_id, "employee-a")
        self.assertTrue(any(item["event_type"] == "material_candidate_selected" for item in details["events"]))

    def test_material_query_can_limit_processing_to_selected_line_numbers(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [
                {"values": {"line_no": "1", "customer_product_code": "CUST-001", "customer_spec": "规格一", "quantity": "1"}},
                {"values": {"line_no": "2", "customer_product_code": "CUST-002", "customer_spec": "规格二", "quantity": "1"}},
            ],
        })

        result = build_material_query_mock(self.case_id, "employee-a", "employee-a", line_nos={2})

        self.assertEqual([item["line_no"] for item in result["items"]], [2])
        self.assertEqual([item["line_no"] for item in get_material_resolution_states(self.case_id, "employee-a")["items"]], [2])

    def test_clearing_query_result_then_saving_allows_material_creation(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "customer_product_code": "CUST-001", "customer_spec": "规格一",
                "customer_spec_match": "匹配规格一", "quantity": "1",
            }}],
        })
        build_material_query_mock(self.case_id, "employee-a", "employee-a")
        self.assertTrue(get_material_resolution_states(self.case_id, "employee-a")["items"])

        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "product_code": "", "product_name": "",
                "customer_product_code": "CUST-001", "customer_spec": "规格一",
                "customer_spec_match": "匹配规格一", "quantity": "1",
            }}],
        })

        self.assertEqual(get_material_resolution_states(self.case_id, "employee-a")["items"], [])
        result = build_material_creation(self.case_id, "employee-a", "employee-a", [{
            "line_no": 1, "adhesive_code": "", "customer_product_code": "CUST-001",
            "customer_spec": "规格一", "customer_spec_match": "匹配规格一",
        }])
        self.assertEqual(result["items"][0]["status"], "creating")

    def test_reextract_clears_stale_material_candidates_but_keeps_query_history(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "C001"},
            "lines": [{"values": {
                "line_no": "1", "customer_product_code": "CUST-001", "customer_spec": "旧规格",
                "quantity": "1",
            }}],
        })
        build_material_query_mock(self.case_id, "employee-a", "employee-a")
        self.assertEqual(get_material_resolution_states(self.case_id, "employee-a")["items"][0]["candidate_count"], 2)

        regenerated = {
            "header": {},
            "lines": [{"values": {"line_no": "1", "customer_product_code": "CUST-NEW", "customer_spec": "新规格"}, "sources": {}}],
        }
        with patch(
            "fangzheng_web_app.order_entry_service._initial_template_data",
            return_value=(regenerated["header"], regenerated["lines"]),
        ):
            reextract_template(self.case_id, "employee-a")

        self.assertEqual(get_material_resolution_states(self.case_id, "employee-a")["items"], [])
        details = get_order_detail_records(self.case_id, "employee-a")
        reextract_event = next(item for item in details["events"] if item["event_type"] == "template_reextracted")
        self.assertEqual(len(reextract_event["detail"]["cleared_material_resolution_tasks"]), 1)
        self.assertTrue(details["calls"])

    def test_real_mode_posts_official_payload_and_keeps_multiple_candidates_linked(self) -> None:
        get_or_create_template(self.case_id, "employee-a")
        save_template(self.case_id, "employee-a", {
            "header": {"bill_to_customer_code": "103878"},
            "lines": [{"values": {
                "line_no": "1", "material_status": "查询",
                "customer_product_code": "A021000550",
                "customer_spec": "南亚新材料 NY6180L 原始客户规格",
                "customer_spec_match": (
                    "南亚新材料 NY6180L 0.127 mm 1/1 RTF2/RTF2 经41纬49inch "
                    "* 不含铜 有卤 1080x2 TG210(DMA) 橙红色 无水印 高速材料 "
                    "CTI≥175 公差±0.018mm"
                ),
                "product_name": "创建料号中", "quantity": "20",
            }}],
        })
        with db.db_cursor() as conn:
            row = conn.execute(
                "SELECT id,values_json FROM order_entry_template_lines ORDER BY id DESC LIMIT 1"
            ).fetchone()
            values = json.loads(row["values_json"])
            values.update({"product_type": "PP", "customer_spec_match": "客户匹配规格"})
            conn.execute(
                "UPDATE order_entry_template_lines SET values_json=? WHERE id=?",
                (json.dumps(values, ensure_ascii=False), row["id"]),
            )
        config = get_interface_config("material_batch_query")
        save_interface_config("material_batch_query", {
            "display_name": config["display_name"], "description": config["description"],
            "mode": "real", "method": "POST", "endpoint_url": config["endpoint_url"],
            "request_mapping": json.dumps(config["request_mapping"], ensure_ascii=False),
            "response_mapping": json.dumps(config["response_mapping"], ensure_ascii=False),
            "mock_scenarios": json.dumps(config["mock_scenarios"], ensure_ascii=False),
        }, "employee-a")
        response = {
            "msg": "处理成功", "code": 200,
            "hitMaterialList": [
                {"scca03": "A021000550", "scca05": "客户匹配规格", "peag01": "6900013796", "peag08": "不应使用", "peag09": "旧品名一"},
                {"scca03": "A021000550", "scca05": "客户匹配规格", "peag01": "6900013797", "peag08": "不应使用", "peag09": "旧品名二"},
            ],
        }
        with patch(
            "fangzheng_web_app.order_interface_service._post_json_endpoint",
            return_value=(200, response, 18),
        ) as request_mock:
            result = build_material_query(self.case_id, "employee-a", "employee-a")
        payload = request_mock.call_args.args[1]
        self.assertEqual(result["mode"], "real")
        self.assertEqual(payload["customerCode"], "103878")
        self.assertEqual(payload["operatorCode"], "employee-a")
        self.assertEqual(payload["materialInfoList"][0]["categoryCode"], "698")
        self.assertNotIn("category", payload["materialInfoList"][0])
        self.assertEqual(payload["materialInfoList"][0]["customerSpecOld"], "南亚新材料 NY6180L 原始客户规格")
        self.assertNotIn("customerSpec", payload["materialInfoList"][0])
        self.assertNotIn("adhesiveCode", payload["materialInfoList"][0])
        self.assertNotIn("newFlag", payload["materialInfoList"][0])
        self.assertNotIn("oldProductName", payload["materialInfoList"][0])
        states = get_material_resolution_states(self.case_id, "employee-a")["items"]
        self.assertEqual(
            states[0]["candidates"],
            [
                {"product_code": "6900013796", "product_name": "旧品名一"},
                {"product_code": "6900013797", "product_name": "旧品名二"},
            ],
        )
        with db.db_cursor() as conn:
            call = conn.execute(
                "SELECT is_mock,http_status,duration_ms FROM order_interface_call_logs ORDER BY id DESC LIMIT 1"
            ).fetchone()
        self.assertEqual((call["is_mock"], call["http_status"], call["duration_ms"]), (0, 200, 18))

    def test_real_material_request_sends_new_fields_only_for_create_action(self) -> None:
        query_item = _real_material_request_item({
            "material_status": "新增", "product_type": "PP", "customer_product_code": "CUST-001",
            "customer_spec": "原客户规格", "customer_spec_match": "客户规格匹配",
            "adhesive_code": "ADH-01", "product_name": "现有品名",
        })
        self.assertEqual(query_item["customerSpecOld"], "原客户规格")
        self.assertNotIn("oldProductName", query_item)
        self.assertNotIn("adhesiveCode", query_item)
        self.assertNotIn("customerSpec", query_item)
        self.assertNotIn("newFlag", query_item)

        new_item = _real_material_request_item({
            "material_status": "新增", "product_type": "基板", "customer_product_code": "CUST-002",
            "customer_spec": "原客户规格", "customer_spec_match": "客户规格匹配",
            "adhesive_code": "ADH-02", "product_name": "新品名",
        }, create=True)
        self.assertEqual(new_item["categoryCode"], "718")
        self.assertNotIn("category", new_item)
        self.assertEqual(new_item["customerSpecOld"], "原客户规格")
        self.assertEqual(new_item["adhesiveCode"], "ADH-02")
        self.assertEqual(new_item["customerSpec"], "客户规格匹配")
        self.assertEqual(new_item["newFlag"], "Y")
        self.assertEqual(new_item["newProductName"], "")
        self.assertNotIn("oldProductName", new_item)

    def test_unmatched_board_query_backfills_marketing_agent_result(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        values = dict(template["lines"][0]["values"])
        values.update({
            "line_no": "1", "product_type": "基板", "customer_product_code": "BOARD-001",
            "customer_spec": "基板客户规格", "remark": "订单备注", "product_code": "", "product_name": "",
        })
        with db.db_cursor() as conn:
            conn.execute(
                "UPDATE order_entry_templates SET header_json=? WHERE id=?",
                (json.dumps({"bill_to_customer_code": "C-BOARD"}, ensure_ascii=False), template["id"]),
            )
            conn.execute(
                "UPDATE order_entry_template_lines SET values_json=? WHERE template_id=? AND line_no=1",
                (json.dumps(values, ensure_ascii=False), template["id"]),
            )
            conn.execute(
                """INSERT INTO automation_customers(customer_code,customer_short_name,status,created_at,updated_at)
                   VALUES (?,?,?,?,?)""",
                ("C-BOARD", "基板客户", "active", "2026-09-01", "2026-09-01"),
            )
        config = get_interface_config("material_batch_query")
        save_interface_config("material_batch_query", {
            "display_name": config["display_name"], "description": config["description"],
            "mode": "real", "method": "POST", "endpoint_url": config["endpoint_url"],
            "request_mapping": json.dumps(config["request_mapping"], ensure_ascii=False),
            "response_mapping": json.dumps(config["response_mapping"], ensure_ascii=False),
            "mock_scenarios": json.dumps(config["mock_scenarios"], ensure_ascii=False),
        }, "employee-a")
        with patch(
            "fangzheng_web_app.order_interface_service._post_json_endpoint",
            return_value=(200, {"code": 200, "msg": "未找到料号", "hitMaterialList": []}, 10),
        ), patch(
            "fangzheng_web_app.order_interface_service.calculate_transcode_agent_quote",
            return_value={"result": "BOARD-CODE", "note": "营销转码成功"},
        ) as quote:
            build_material_query(self.case_id, "employee-a", "employee-a")
        self.assertEqual(quote.call_args.args[0], "基板客户规格")
        self.assertEqual(quote.call_args.kwargs["customer"], "基板客户")
        self.assertEqual(quote.call_args.kwargs["customer_code"], "C-BOARD")
        self.assertEqual(quote.call_args.kwargs["order_remark"], "订单备注")
        with db.db_cursor() as conn:
            line = conn.execute(
                "SELECT values_json,sources_json FROM order_entry_template_lines WHERE template_id=? AND line_no=1",
                (template["id"],),
            ).fetchone()
        self.assertEqual(json.loads(line["values_json"])["product_name"], "BOARD-CODE")
        self.assertEqual(json.loads(line["sources_json"])["product_name"]["label"], "营销转码Agent")

    def test_unmatched_pp_query_backfills_pending_transcode_code(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        values = dict(template["lines"][0]["values"])
        values.update({
            "line_no": "1", "product_type": "PP", "customer_product_code": "PP-001",
            "customer_spec": "PP客户规格", "remark": "PP备注", "product_code": "", "product_name": "",
        })
        with db.db_cursor() as conn:
            conn.execute(
                "UPDATE order_entry_templates SET header_json=? WHERE id=?",
                (json.dumps({"bill_to_customer_code": "C-PP"}, ensure_ascii=False), template["id"]),
            )
            conn.execute(
                "UPDATE order_entry_template_lines SET values_json=? WHERE template_id=? AND line_no=1",
                (json.dumps(values, ensure_ascii=False), template["id"]),
            )
        config = get_interface_config("material_batch_query")
        save_interface_config("material_batch_query", {
            "display_name": config["display_name"], "description": config["description"],
            "mode": "real", "method": "POST", "endpoint_url": config["endpoint_url"],
            "request_mapping": json.dumps(config["request_mapping"], ensure_ascii=False),
            "response_mapping": json.dumps(config["response_mapping"], ensure_ascii=False),
            "mock_scenarios": json.dumps(config["mock_scenarios"], ensure_ascii=False),
        }, "employee-a")
        with patch(
            "fangzheng_web_app.order_interface_service._post_json_endpoint",
            return_value=(200, {"code": 200, "msg": "未找到料号", "hitMaterialList": []}, 10),
        ), patch(
            "fangzheng_web_app.order_interface_service.calculate_pp_transcode_quote",
            return_value={"pending_code": "PP-PENDING", "note": "PP 转码待确认"},
        ) as quote:
            build_material_query(self.case_id, "employee-a", "employee-a")
        self.assertEqual(quote.call_args.args[0], "PP客户规格")
        self.assertEqual(quote.call_args.kwargs["customer_code"], "C-PP")
        self.assertEqual(quote.call_args.kwargs["order_remark"], "PP备注")
        with db.db_cursor() as conn:
            line = conn.execute(
                "SELECT values_json FROM order_entry_template_lines WHERE template_id=? AND line_no=1",
                (template["id"],),
            ).fetchone()
        self.assertEqual(json.loads(line["values_json"])["product_name"], "PP-PENDING")

    def test_pp_transcode_placeholder_is_not_backfilled_as_product_name(self) -> None:
        with patch(
            "fangzheng_web_app.order_interface_service.calculate_pp_transcode_quote",
            return_value={"pending_code": "***************************", "note": "存在未识别字段"},
        ):
            result = _transcode_product_name(
                {"product_type": "PP", "customer_spec": "PP 客户规格"}, "C-PP", "PP客户", "employee-a"
            )

        self.assertEqual(result["product_name"], "")
        self.assertIn("全为占位符", result["message"])

    def test_material_creation_sends_only_blank_rows_and_records_external_task(self) -> None:
        _case, template = get_or_create_template(self.case_id, "employee-a")
        header = dict(template["header"])
        header["bill_to_customer_code"] = "103878"
        values = dict(template["lines"][0]["values"])
        values.update({
            "line_no": "1", "product_code": "", "product_name": "转码品名",
            "product_type": "PP", "adhesive_code": "6CNL",
            "customer_product_code": "CUST-NEW-01", "customer_spec": "原客户规格",
            "customer_spec_match": "标准客户规格", "quantity": "10",
        })
        with db.db_cursor() as conn:
            conn.execute(
                "UPDATE order_entry_templates SET header_json=? WHERE id=?",
                (json.dumps(header, ensure_ascii=False), template["id"]),
            )
            conn.execute(
                "UPDATE order_entry_template_lines SET values_json=? WHERE template_id=? AND line_no=1",
                (json.dumps(values, ensure_ascii=False), template["id"]),
            )
        config = get_interface_config("material_batch_query")
        save_interface_config("material_batch_query", {
            "display_name": config["display_name"], "description": config["description"],
            "enabled": "1", "mode": "real", "method": "POST",
            "endpoint_url": config["endpoint_url"],
            "request_mapping": json.dumps(config["request_mapping"], ensure_ascii=False),
            "response_mapping": json.dumps(config["response_mapping"], ensure_ascii=False),
            "mock_scenarios": json.dumps(config["mock_scenarios"], ensure_ascii=False),
        }, "employee-a")
        with patch(
            "fangzheng_web_app.order_interface_service._post_json_endpoint",
            return_value=(200, {"code": 200, "msg": "已受理", "external_task_id": "TASK-ORDER-001"}, 12),
        ) as request_mock:
            result = build_material_creation(self.case_id, "employee-a", "employee-a", [{
                "line_no": 1, "adhesive_code": "6CNL", "customer_product_code": "CUST-NEW-01",
                "customer_spec": "原客户规格", "customer_spec_match": "标准客户规格", "product_name": "转码品名",
            }])
        item = request_mock.call_args.args[1]["materialInfoList"][0]
        self.assertEqual(item["newFlag"], "Y")
        self.assertNotIn("category", item)
        self.assertEqual(item["adhesiveCode"], "6CNL")
        self.assertEqual(item["customerSpec"], "标准客户规格")
        self.assertEqual(item["customerSpecOld"], "原客户规格")
        self.assertEqual(item["newProductName"], "")
        self.assertNotIn("oldProductName", item)
        self.assertEqual(result["items"][0]["external_task_id"], "TASK-ORDER-001")
        with db.db_cursor() as conn:
            task = conn.execute(
                "SELECT external_task_id FROM order_material_resolution_tasks WHERE case_id=? AND line_no=1",
                (self.case_id,),
            ).fetchone()
            line = conn.execute(
                "SELECT values_json FROM order_entry_template_lines WHERE template_id=? AND line_no=1",
                (template["id"],),
            ).fetchone()
        self.assertEqual(task["external_task_id"], "TASK-ORDER-001")
        state = get_material_resolution_states(self.case_id, "employee-a")["items"][0]
        self.assertEqual(state["external_task_id"], "TASK-ORDER-001")
        self.assertEqual(json.loads(line["values_json"])["material_status"], "新增")
        with self.assertRaisesRegex(ValueError, "已有产品编号"):
            build_material_creation(self.case_id, "employee-a", "employee-a", [{
                "line_no": 1, "adhesive_code": "6CNL", "customer_product_code": "CUST-NEW-01",
                "customer_spec": "原客户规格", "customer_spec_match": "标准客户规格",
            }])
