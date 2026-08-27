from __future__ import annotations

import io
import unittest
import zipfile
from unittest.mock import patch

from flask import Flask

from fangzheng_web_app.routes import bp
from fangzheng_web_app.traditional_simplified_service import (
    WorkbookConversionError,
    convert_text_to_simplified,
    convert_workbook_to_simplified,
)


def _workbook_bytes(*, macro: bool = False) -> bytes:
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"""
    workbook = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="考試表" sheetId="1" state="hidden"/></sheets>
</workbook>"""
    shared = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="1">
  <si><r><rPr><b/></rPr><t>考試</t></r><r><t>板</t></r></si>
</sst>"""
    sheet = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">
  <c r="A1" t="s"><v>0</v></c><c r="A2" t="s"><v>0</v></c>
  <c r="B1" t="inlineStr"><is><t>無鹵素</t></is></c>
  <c r="C1" t="n" s="1"><v>123</v></c>
  <c r="D1" t="str"><f>IF(1=1,&quot;繁體&quot;,&quot;&quot;)</f><v>繁體</v></c>
</row></sheetData><mergeCells count="1"><mergeCell ref="C1:C2"/></mergeCells></worksheet>"""
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/sharedStrings.xml", shared)
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
        archive.writestr("xl/styles.xml", b"unchanged-styles")
        if macro:
            archive.writestr("xl/vbaProject.bin", b"unchanged-macro")
    return output.getvalue()


class TraditionalSimplifiedServiceTests(unittest.TestCase):
    def test_text_conversion_handles_sample_terms_and_keeps_simplified_text(self) -> None:
        converted, changed = convert_text_to_simplified("考試板 雙面反轉 無鹵素 樣")
        self.assertEqual(converted, "考试板 双面反转 无卤素 样")
        self.assertEqual(changed, 6)
        self.assertEqual(convert_text_to_simplified("已经是简体"), ("已经是简体", 0))

    def test_text_conversion_replaces_non_breaking_space_rejected_by_target_system(self) -> None:
        converted, changed = convert_text_to_simplified("HVLP0.8\u00a0無鹵素NON-DICY")

        self.assertEqual(converted, "HVLP0.8 无卤素NON-DICY")
        self.assertEqual(changed, 3)

    def test_workbook_conversion_preserves_structure_formula_and_macro(self) -> None:
        source = _workbook_bytes(macro=True)
        result = convert_workbook_to_simplified(source, "測試.xlsm")

        self.assertEqual(result.filename, "測試_简体版.xlsm")
        self.assertEqual(result.stats.sheet_count, 1)
        self.assertEqual(result.stats.text_cell_count, 3)
        self.assertEqual(result.stats.changed_cell_count, 3)
        self.assertEqual(result.stats.changed_character_count, 4)
        with zipfile.ZipFile(io.BytesIO(source)) as before, zipfile.ZipFile(result.content) as after:
            shared = after.read("xl/sharedStrings.xml").decode("utf-8")
            sheet = after.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("考试", shared)
            self.assertIn("无卤素", sheet)
            self.assertIn("考試表", after.read("xl/workbook.xml").decode("utf-8"))
            self.assertIn("繁體", sheet)
            self.assertEqual(before.read("xl/styles.xml"), after.read("xl/styles.xml"))
            self.assertEqual(before.read("xl/vbaProject.bin"), after.read("xl/vbaProject.bin"))

    def test_invalid_empty_and_oversized_workbooks_are_rejected(self) -> None:
        with self.assertRaisesRegex(WorkbookConversionError, "仅支持"):
            convert_workbook_to_simplified(b"content", "old.xls")
        with self.assertRaisesRegex(WorkbookConversionError, "为空"):
            convert_workbook_to_simplified(b"", "empty.xlsx")
        with self.assertRaisesRegex(WorkbookConversionError, "损坏"):
            convert_workbook_to_simplified(b"not-a-zip", "broken.xlsx")
        with patch(
            "fangzheng_web_app.traditional_simplified_service.MAX_UNCOMPRESSED_BYTES", 10
        ):
            with self.assertRaisesRegex(WorkbookConversionError, "体积过大"):
                convert_workbook_to_simplified(_workbook_bytes(), "large.xlsx")


class TraditionalSimplifiedRouteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = Flask(__name__, template_folder="../templates", static_folder="../static")
        self.app.config.update(SECRET_KEY="test-secret", TESTING=True)
        self.app.register_blueprint(bp)
        self.client = self.app.test_client()

    def _login(self) -> None:
        with self.client.session_transaction() as session:
            session["employee_id"] = "tester"

    def test_page_and_endpoints_require_login(self) -> None:
        for path in [
            "/features/traditional-to-simplified",
            "/features/traditional-to-simplified/excel",
            "/features/traditional-to-simplified/text",
        ]:
            response = self.client.get(path) if path.endswith("simplified") else self.client.post(path)
            self.assertEqual(response.status_code, 302)
            self.assertIn("/login", response.headers["Location"])

    def test_page_renders_both_conversion_modes(self) -> None:
        self._login()
        with patch("fangzheng_web_app.routes.get_user", return_value=None):
            response = self.client.get("/features/traditional-to-simplified")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Excel 转换", html)
        self.assertIn("文本转换", html)
        self.assertIn(".xlsx / .xlsm", html)

    def test_text_endpoint_converts_and_validates_input(self) -> None:
        self._login()
        with patch("fangzheng_web_app.routes.get_user", return_value=None):
            response = self.client.post(
                "/features/traditional-to-simplified/text", json={"text": "無鹵素考試板"}
            )
            empty = self.client.post(
                "/features/traditional-to-simplified/text", json={"text": "  "}
            )
            with patch("fangzheng_web_app.routes.TRADITIONAL_SIMPLIFIED_MAX_TEXT_LENGTH", 2):
                oversized = self.client.post(
                    "/features/traditional-to-simplified/text", json={"text": "繁體字"}
                )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["converted_text"], "无卤素考试板")
        self.assertEqual(empty.status_code, 400)
        self.assertEqual(oversized.status_code, 400)

    def test_excel_endpoint_downloads_converted_workbook_with_stats(self) -> None:
        self._login()
        with patch("fangzheng_web_app.routes.get_user", return_value=None):
            response = self.client.post(
                "/features/traditional-to-simplified/excel",
                data={"excel_file": (io.BytesIO(_workbook_bytes()), "測試.xlsx")},
                content_type="multipart/form-data",
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Conversion-Sheets"], "1")
        self.assertEqual(response.headers["X-Conversion-Changed-Cells"], "3")
        self.assertIn("attachment", response.headers["Content-Disposition"])
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            self.assertIn("考试", archive.read("xl/sharedStrings.xml").decode("utf-8"))


if __name__ == "__main__":
    unittest.main()
