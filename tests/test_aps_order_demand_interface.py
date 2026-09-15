from __future__ import annotations

import unittest

from fangzheng_web_app.order_interface_service import (
    INTERFACE_DEFAULTS,
    _aps_order_demand_request_payload,
    test_interface_config,
)


class ApsOrderDemandInterfaceTests(unittest.TestCase):
    def test_default_config_and_mock_test_follow_aps_contract(self) -> None:
        config = INTERFACE_DEFAULTS["aps_order_demand_import"]
        self.assertEqual(
            config["base_url"],
            "http://aps.nouyatec.com:13000/forward/erp_order_change_blocking",
        )
        self.assertIn("data[].require_shipment_date", config["request_mapping"])
        self.assertIn("data[].Order_Item_Account_Set_outer_key", config["request_mapping"])
        self.assertIn("data[].alter_type", config["request_mapping"])
        self.assertIn("data[].creator_name", config["request_mapping"])
        self.assertIn("data[].created_at", config["request_mapping"])

        result = test_interface_config({
            "interface_key": "aps_order_demand_import",
            "mode": "mock",
            "method": "POST",
            "endpoint_url": config["base_url"],
        })

        self.assertTrue(result["ok"])
        self.assertEqual(result["mode"], "mock")
        self.assertEqual(len(result["request"]["data"]), 1)
        item = result["request"]["data"][0]
        self.assertEqual(item["require_shipment_date"], "2026-09-15")
        self.assertEqual(item["Order_Item_Account_Set_outer_key"], "220-260114007_2_KL01")
        self.assertEqual(item["alter_type"], "交期变更")
        self.assertEqual(item["creator_name"], "Mock 测试用户")
        self.assertEqual(item["created_at"], "2026-09-14 10:00:00")
        self.assertEqual(result["response"]["success_count"], 1)
        self.assertEqual(result["response"]["fail_count"], 0)
        self.assertEqual(result["response"]["failed_details"], [])

    def test_change_submission_uses_confirmed_erp_item_and_shared_modal_values(self) -> None:
        payload = _aps_order_demand_request_payload(
            [{"line_no": 2, "values_json": '{"delivery_date":"2026/09/25"}'}],
            {2: {"status": "matched", "selected_candidate": {
                "scta39": "220-260114007", "sctb35": "2", "acsn": "NY02",
            }}},
            "交期变更", "客户要求提前交货", "张三", "2026-09-14 10:00:00",
        )

        self.assertEqual(payload, {"data": [{
            "require_shipment_date": "2026-09-25",
            "Order_Item_Account_Set_outer_key": "220-260114007_2_KL02",
            "alter_type": "交期变更",
            "creator_name": "张三",
            "require_specification": "客户要求提前交货",
            "created_at": "2026-09-14 10:00:00",
        }]})

    def test_change_submission_rejects_missing_erp_item(self) -> None:
        with self.assertRaisesRegex(ValueError, "缺少 ERP 项次"):
            _aps_order_demand_request_payload(
                [{"line_no": 2, "values_json": '{"delivery_date":"2026-09-25"}'}],
                {2: {"status": "matched", "selected_candidate": {"scta39": "220-260114007"}}},
                "交期变更", "", "张三", "2026-09-14 10:00:00",
            )

    def test_change_submission_rejects_missing_account_set_code(self) -> None:
        with self.assertRaisesRegex(ValueError, "账套组织代码"):
            _aps_order_demand_request_payload(
                [{"line_no": 2, "values_json": '{"delivery_date":"2026-09-25"}'}],
                {2: {"status": "matched", "selected_candidate": {
                    "scta39": "220-260114007", "sctb35": "2",
                }}},
                "交期变更", "", "张三", "2026-09-14 10:00:00",
            )

    def test_change_submission_uses_kl55_for_ny03(self) -> None:
        payload = _aps_order_demand_request_payload(
            [{"line_no": 3, "values_json": '{"delivery_date":"2026-09-25"}'}],
            {3: {"status": "matched", "selected_candidate": {
                "scta39": "220-260114008", "sctb35": "3", "acsn": "ny03",
            }}},
            "交期变更", "", "张三", "2026-09-14 10:00:00",
        )

        self.assertEqual(
            payload["data"][0]["Order_Item_Account_Set_outer_key"],
            "220-260114008_3_KL55",
        )


if __name__ == "__main__":
    unittest.main()
