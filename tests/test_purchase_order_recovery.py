from __future__ import annotations

import unittest

from fangzheng_web_app.purchase_order_pipeline import _native_detail_rows_missing_from_docling
from fangzheng_web_app.purchase_result_normalizer import normalize_order_spec_spacing


class PurchaseOrderSpacingTests(unittest.TestCase):
    def test_confirmed_line_wrap_artifacts_are_repaired(self) -> None:
        self.assertEqual(
            normalize_order_spec_spacing("7628 48% 150M/卷耐 CAF"),
            "7628 48% 150M/卷 耐 CAF",
        )
        self.assertEqual(normalize_order_spec_spacing("526*626m m纬 向"), "526*626mm纬向")
        self.assertEqual(normalize_order_spec_spacing("FR -4 高 速材料"), "FR-4 高速材料")


class NativeDetailRecoveryTests(unittest.TestCase):
    @staticmethod
    def _cells(rows: list[list[str]]) -> list[dict]:
        return [
            {"row_index": row_index, "column_index": column_index, "text": value}
            for row_index, row in enumerate(rows)
            for column_index, value in enumerate(row)
        ]

    def test_only_valid_rows_missing_from_docling_are_recovered(self) -> None:
        headers = [
            "NO.", "物料编码", "材料名称", "规格", "数量", "单位",
            "单价（含税）", "金额", "要求到货日期", "备注",
        ]
        rows = [
            headers,
            ["25", "C0202NY09097", "PP 1067", "526*626mm纬向", "36", "片", "13.30", "478.80", "2026-08-20", ""],
            ["26", "C0202NY09098", "PP 2116", "526*626mm纬向", "58", "片", "15.00", "870.00", "2026-08-20", ""],
        ]
        native = {
            "pages": [
                {"page_index": 8, "tables": [{"table_index": 0, "cells": self._cells(rows)}]}
            ]
        }
        docling_rows = [{"standard": {"序号": "25", "物料编码": "C0202NY09097"}}]

        tables, recovered, issues = _native_detail_rows_missing_from_docling(native, docling_rows)

        self.assertEqual(len(tables), 1)
        self.assertEqual([row["standard"]["序号"] for row in recovered], ["26"])
        self.assertEqual(recovered[0]["method"], "pdf_native_table_recovery")
        self.assertIsInstance(issues, list)


if __name__ == "__main__":
    unittest.main()
