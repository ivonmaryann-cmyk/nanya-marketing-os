from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import price_calculation_rules as rule_service
from fangzheng_web_app.price_calculation_service import load_price_rules


class PriceCalculationRuleVersionTests(unittest.TestCase):
    def test_missing_active_rule_file_is_not_replaced_by_bootstrap(self) -> None:
        settings = {"active_price_rule_version:jingwang": "jingwang_new_rules_20260911_093818"}

        def get_setting(name: str, default: str = "") -> str:
            return settings.get(name, default)

        def set_setting(name: str, value: str) -> None:
            settings[name] = value

        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            rule_service, "PRICE_CALCULATION_RULES_DIR", Path(temp_dir)
        ), patch.object(rule_service, "get_setting", side_effect=get_setting), patch.object(
            rule_service, "set_setting", side_effect=set_setting
        ):
            self.assertEqual(
                "jingwang_new_rules_20260911_093818",
                rule_service.ensure_default_price_rule_version("jingwang"),
            )
            self.assertEqual("jingwang_new_rules_20260911_093818", settings["active_price_rule_version:jingwang"])
            self.assertEqual(
                "jingwang_new_rules_20260911_093818",
                rule_service.get_active_price_rule_version("jingwang"),
            )

    def test_missing_local_rule_file_reports_storage_problem(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            missing = Path(temp_dir) / "uploaded" / rule_service.PRICE_RULE_FILENAME
            with self.assertRaisesRegex(ValueError, "本机缺失"):
                load_price_rules("jingwang", missing)
