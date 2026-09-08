from __future__ import annotations

import json
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from werkzeug.datastructures import FileStorage

from fangzheng_web_app import price_calculation_rules as rule_service


class GuangheRuleUploadTests(unittest.TestCase):
    def test_merged_nanya_quote_is_the_only_required_upload(self) -> None:
        source = Path("fangzheng_web_app/default_rules/price_calculation/guanghe/price_rules.xlsx")
        settings: dict[str, str] = {}

        def get_setting(name: str, default: str = "") -> str:
            return settings.get(name, default)

        def set_setting(name: str, value: str) -> None:
            settings[name] = value

        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            rule_service, "PRICE_CALCULATION_RULES_DIR", Path(temp_dir)
        ), patch.object(rule_service, "get_setting", side_effect=get_setting), patch.object(
            rule_service, "set_setting", side_effect=set_setting
        ):
            version = rule_service.save_new_guanghe_rule_version(
                FileStorage(BytesIO(source.read_bytes()), filename="广合南亚价格更新.xlsx"),
                updated_by="tester",
                remark="",
            )

            version_dir = Path(temp_dir) / "guanghe" / "versions" / version
            self.assertTrue((version_dir / rule_service.GUANGHE_NANYA_RULE_FILENAME).is_file())
            self.assertTrue((version_dir / rule_service.PRICE_RULE_FILENAME).is_file())
            self.assertFalse((version_dir / rule_service.GUANGHE_HUANGSHI_RULE_FILENAME).exists())
            self.assertEqual(settings["active_price_rule_version:guanghe"], version)

            history = json.loads(settings["price_rule_history:guanghe"])
            self.assertEqual(history[0]["rule_file"], "广合南亚价格更新.xlsx")
            self.assertIn("南亚新材价格更新表", history[0]["remark"])
