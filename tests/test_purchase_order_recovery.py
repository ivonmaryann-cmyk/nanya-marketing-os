from __future__ import annotations

import unittest

from fangzheng_web_app.purchase_order_pipeline import (
    _native_detail_rows_missing_from_docling,
    _native_detail_rows_for_merged_docling,
    _native_grid_rows_from_page,
)
from fangzheng_web_app.purchase_field_rules import header_score
from fangzheng_web_app.purchase_result_normalizer import normalize_order_spec_spacing


class PurchaseOrderSpacingTests(unittest.TestCase):
    def test_confirmed_line_wrap_artifacts_are_repaired(self) -> None:
        self.assertEqual(
            normalize_order_spec_spacing("7628 48% 150M/卷耐 CAF"),
            "7628 48% 150M/卷 耐 CAF",
        )
        self.assertEqual(normalize_order_spec_spacing("526*626m m纬 向"), "526*626mm纬向")
        self.assertEqual(normalize_order_spec_spacing("FR -4 高 速材料"), "FR-4 高速材料")

    def test_material_description_is_not_mapped_as_material_name(self) -> None:
        headers = [
            "序号 No",
            "物料编码 Material Code",
            "物料品名 Material Name",
            "物料描述 Description",
            "数量 Quantity",
            "单位 Unit",
            "单价 Unit Price",
            "金额 Total Amount",
        ]

        _score, mapping = header_score(headers)

        self.assertEqual(mapping[2], "物料名称")
        self.assertEqual(mapping[3], "说明")


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

    def test_headerless_continuation_row_is_rebuilt_from_pdf_grid(self) -> None:
        boundaries = [0, 40, 100, 180, 260]
        page = {
            "width": 300,
            "lines": [
                {"orientation": "v", "bbox": [x, 20, x, 80]}
                for x in boundaries
            ],
            "words": [
                {"text": "4", "bbox": [10, 45, 15, 55]},
                {"text": "MAT-004", "bbox": [45, 45, 90, 55]},
                {"text": "FR-4", "bbox": [110, 35, 135, 45]},
                {"text": "NY2150H 1.1 mm", "bbox": [185, 35, 250, 45]},
                {"text": "7628x5 TG150", "bbox": [185, 55, 245, 65]},
            ],
        }

        rows = _native_grid_rows_from_page(page, 4)

        self.assertEqual(rows[0][0], "4")
        self.assertEqual(rows[0][1], "MAT-004")
        self.assertEqual(rows[0][2], "FR-4")
        self.assertEqual(rows[0][3], "NY2150H 1.1 mm 7628x5 TG150")

    def test_native_columns_replace_docling_merged_material_fields(self) -> None:
        native_rows = [
            ["序号", "物料编码", "物料名称", "单位", "数量", "含税单价", "金额"],
            ["1", "MAT-001-TAIL", "NY 3170M2 0.8mm", "张", "200", "355", "71000"],
        ]
        native = {
            "pages": [
                {"page_index": 0, "tables": [{"table_index": 0, "cells": self._cells(native_rows)}]}
            ]
        }
        docling_rows = [
            {
                "standard": {
                    "序号": "1",
                    "物料编码": "MAT-001",
                    "物料名称": "MAT-001-TAIL NY 3170M2 0.8mm",
                    "单位": "张",
                    "数量": "200",
                    "含税单价": "355",
                    "金额": "71000",
                    "交货日期": "",
                    "备注": "",
                }
            }
        ]

        tables, rows, issues = _native_detail_rows_for_merged_docling(native, docling_rows)

        self.assertEqual(len(tables), 1)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["standard"]["物料编码"], "MAT-001-TAIL")
        self.assertEqual(rows[0]["standard"]["物料名称"], "NY 3170M2 0.8mm")
        self.assertEqual(rows[0]["method"], "pdf_native_table_reconciled")
        self.assertEqual(issues, [])


if __name__ == "__main__":
    unittest.main()
