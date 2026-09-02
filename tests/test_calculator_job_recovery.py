from __future__ import annotations

import unittest
from unittest.mock import MagicMock, patch

from fangzheng_web_app import calculator_service


class CalculatorJobRecoveryTests(unittest.TestCase):
    def test_missing_task_rule_version_falls_back_to_active_version(self):
        job = {"stored_input_path": "unused.xlsx"}
        calculator = MagicMock()
        calculator.__file__ = __file__

        with (
            patch("fangzheng_web_app.db.get_job", return_value=job),
            patch.object(calculator_service, "load_calculator_module", return_value=calculator),
            patch.object(
                calculator_service,
                "load_rule_dataframes",
                side_effect=[FileNotFoundError("missing rule"), (MagicMock(), MagicMock())],
            ) as load_rules,
            patch.object(calculator_service, "get_active_rule_version", return_value="rules_current"),
            patch.object(calculator_service, "append_job_log") as append_log,
            patch.object(calculator_service, "update_job_status"),
            patch.object(calculator_service, "prune_jobs_for_employee", return_value=[]),
            patch.object(calculator_service, "load_workbook_compat", side_effect=AssertionError("stop after rules")),
        ):
            calculator_service.run_job(1, "employee", "missing_rules")

        self.assertEqual(load_rules.call_args_list[0].args, ("missing_rules",))
        self.assertEqual(load_rules.call_args_list[1].args, ("rules_current",))
        self.assertTrue(
            any("已回退至当前规则版本 rules_current" in str(call.args) for call in append_log.call_args_list)
        )
