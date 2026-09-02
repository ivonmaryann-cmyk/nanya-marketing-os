import os
import unittest
from unittest.mock import DEFAULT, patch

import fangzheng_web_app as application


class SharedDatabaseStartupTests(unittest.TestCase):
    def test_shared_client_skips_maintenance_but_checks_schema(self):
        self.check_startup('false', False)

    def test_default_startup_preserves_maintenance(self):
        self.check_startup(None, True)

    def check_startup(self, setting, expected):
        maintenance = (
            'ensure_rule_center_tables', 'ensure_pp_transcode_tables',
            'reconcile_interrupted_jobs', 'ensure_default_rule_version',
            'ensure_default_transcode_rule_version', 'ensure_default_transcode_agent_rule_version',
            'ensure_default_transcode_semantic_rule_version', 'ensure_daily_backup',
            'seed_pp_transcode_rules', 'ensure_pp_transcode_daily_backup',
            'ensure_default_shennan_rule_version', 'ensure_default_hushi_rule_version',
            'ensure_default_bomin_rule_version', 'ensure_default_price_rule_versions',
        )
        hooks = ('load_local_env', 'ensure_storage_dirs', 'init_db', *maintenance)
        with patch.dict(os.environ), patch.multiple(application, **{name: DEFAULT for name in hooks}) as mocks:
            if setting is None:
                os.environ.pop('APP_STARTUP_MAINTENANCE_ENABLED', None)
            else:
                os.environ['APP_STARTUP_MAINTENANCE_ENABLED'] = setting
            app = application.create_app()
            mocks['init_db'].assert_called_once()
            mocks['ensure_storage_dirs'].assert_called_once()
            for name in maintenance:
                self.assertEqual(mocks[name].call_count, int(expected), name)
            self.assertIn('main.order_automation_material_create', app.view_functions)
