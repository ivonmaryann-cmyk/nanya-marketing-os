from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from fangzheng_web_app import bomin_rules, hushi_rules, rules, shennan_rules


class DedicatedPriceRuleStorageSafetyTests(unittest.TestCase):
    def test_partial_fangzheng_upload_reuses_latest_available_counterpart_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            fallback_account = root / "bootstrap_old" / rules.ACCOUNT_FILENAME
            fallback_account.parent.mkdir()
            fallback_account.write_bytes(b"existing account rules")
            missing_price = root / "missing" / rules.PRICE_FILENAME
            missing_account = root / "missing" / rules.ACCOUNT_FILENAME
            uploaded_price = MagicMock(filename="price.xlsx")
            uploaded_price.save.side_effect = lambda path: Path(path).write_bytes(b"new price rules")
            with patch.object(rules, "RULES_VERSIONS_DIR", root), patch.object(
                rules, "get_rule_file_paths", return_value=(missing_price, missing_account)
            ), patch.object(rules, "validate_rule_files"), patch.object(
                rules, "set_setting"
            ), patch.object(rules, "append_rule_history"):
                version = rules.save_new_rule_version(
                    uploaded_price, None, updated_by="employee-a", remark="只更新价格"
                )
            self.assertEqual(
                b"existing account rules",
                (root / version / rules.ACCOUNT_FILENAME).read_bytes(),
            )

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

    def test_fangzheng_complete_historical_version_can_be_activated(self) -> None:
        settings = {"active_rule_version": "rules_current"}
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            version_dir = root / "rules_previous"
            version_dir.mkdir()
            price_path = version_dir / rules.PRICE_FILENAME
            account_path = version_dir / rules.ACCOUNT_FILENAME
            with patch.object(rules, "RULES_VERSIONS_DIR", root), patch.object(
                rules, "validate_rule_files", return_value=(MagicMock(), MagicMock())
            ) as validate, patch.object(
                rules, "get_setting", side_effect=lambda name, default="": settings.get(name, default)
            ), patch.object(
                rules, "set_setting", side_effect=lambda name, value: settings.__setitem__(name, value)
            ):
                price_path.write_bytes(b"price")
                account_path.write_bytes(b"account")
                self.assertEqual("rules_previous", rules.activate_rule_version("rules_previous"))
                self.assertEqual("rules_previous", settings["active_rule_version"])
                validate.assert_called_once_with(price_path, account_path)

    def test_fangzheng_incomplete_historical_version_cannot_be_activated(self) -> None:
        settings = {"active_rule_version": "rules_current"}
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            version_dir = root / "rules_incomplete"
            version_dir.mkdir()
            (version_dir / rules.PRICE_FILENAME).write_bytes(b"price")
            with patch.object(rules, "RULES_VERSIONS_DIR", root), patch.object(
                rules, "set_setting", side_effect=lambda name, value: settings.__setitem__(name, value)
            ):
                with self.assertRaisesRegex(ValueError, "本机不完整"):
                    rules.activate_rule_version("rules_incomplete")
            self.assertEqual("rules_current", settings["active_rule_version"])

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
