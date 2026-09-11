from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app import bomin_rules, hushi_rules, rules, shennan_rules


class DedicatedPriceRuleStorageSafetyTests(unittest.TestCase):
    def test_fangzheng_missing_uploaded_files_do_not_reset_active_version(self) -> None:
        settings = {"active_rule_version": "rules_uploaded"}
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            rules, "RULES_VERSIONS_DIR", Path(temp_dir)
        ), patch.object(rules, "get_setting", side_effect=lambda name, default="": settings.get(name, default)), patch.object(
            rules, "set_setting", side_effect=lambda name, value: settings.__setitem__(name, value)
        ):
            self.assertEqual("rules_uploaded", rules.ensure_default_rule_version())
            self.assertEqual("rules_uploaded", settings["active_rule_version"])
            self.assertEqual("rules_uploaded", rules.get_active_rule_version())
            with self.assertRaisesRegex(ValueError, "本机缺失"):
                rules.load_rule_dataframes()

    def test_bomin_missing_uploaded_file_does_not_reset_active_version(self) -> None:
        self._assert_missing_storage_is_preserved(
            bomin_rules,
            "BOMIN_RULES_VERSIONS_DIR",
            "active_bomin_rule_version",
            "bomin_uploaded",
            bomin_rules.ensure_default_bomin_rule_version,
            bomin_rules.get_active_bomin_rule_version,
        )

    def test_hushi_missing_uploaded_files_do_not_reset_active_version(self) -> None:
        self._assert_missing_storage_is_preserved(
            hushi_rules,
            "HUSHI_RULES_VERSIONS_DIR",
            "active_hushi_rule_version",
            "hushi_uploaded",
            hushi_rules.ensure_default_hushi_rule_version,
            hushi_rules.get_active_hushi_rule_version,
        )

    def test_shennan_missing_uploaded_file_does_not_reset_active_version(self) -> None:
        self._assert_missing_storage_is_preserved(
            shennan_rules,
            "SHENNAN_RULES_VERSIONS_DIR",
            "active_shennan_rule_version",
            "shennan_uploaded",
            shennan_rules.ensure_default_shennan_rule_version,
            shennan_rules.get_active_shennan_rule_version,
        )

    def _assert_missing_storage_is_preserved(
        self,
        module,
        root_name: str,
        active_key: str,
        active_version: str,
        ensure_default,
        get_active,
    ) -> None:
        settings = {active_key: active_version}
        with tempfile.TemporaryDirectory() as temp_dir, patch.object(
            module, root_name, Path(temp_dir)
        ), patch.object(module, "get_setting", side_effect=lambda name, default="": settings.get(name, default)), patch.object(
            module, "set_setting", side_effect=lambda name, value: settings.__setitem__(name, value)
        ):
            self.assertEqual(active_version, ensure_default())
            self.assertEqual(active_version, settings[active_key])
            self.assertEqual(active_version, get_active())
