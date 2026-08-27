from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app.database.config import (
    AutomationDatabaseConfig,
    ConfigurationDatabaseConfig,
    IdentityDatabaseConfig,
    PlanningDatabaseConfig,
    TranscodeDatabaseConfig,
)
from fangzheng_web_app.platform_rollback import configure_legacy_rollback_environment


CONFIG_TYPES = (
    AutomationDatabaseConfig,
    IdentityDatabaseConfig,
    TranscodeDatabaseConfig,
    ConfigurationDatabaseConfig,
    PlanningDatabaseConfig,
)


class PlatformDatabaseConfigTests(unittest.TestCase):
    def test_platform_postgresql_profile_is_shared_by_every_module(self) -> None:
        environment = {
            "PLATFORM_DATABASE_BACKEND": "postgresql",
            "PLATFORM_DATABASE_URL": "postgresql://unified-example",
            "PLATFORM_DATABASE_READ_WRITE_ENABLED": "true",
            # A conflicting legacy value must not split one platform across databases.
            "AUTOMATION_DATABASE_BACKEND": "sqlite",
            "AUTOMATION_DATABASE_URL": "postgresql://legacy-example",
        }
        with patch.dict(os.environ, environment, clear=True):
            configs = [config_type.from_env() for config_type in CONFIG_TYPES]

        self.assertTrue(all(config.backend == "postgresql" for config in configs))
        self.assertEqual({"postgresql://unified-example"}, {config.database_url for config in configs})

    def test_platform_profile_requires_its_own_url(self) -> None:
        with patch.dict(os.environ, {"PLATFORM_DATABASE_BACKEND": "postgresql"}, clear=True):
            for config_type in CONFIG_TYPES:
                with self.subTest(config=config_type.__name__):
                    with self.assertRaisesRegex(ValueError, "PLATFORM_DATABASE_URL"):
                        config_type.from_env()

    def test_platform_profile_requires_global_read_write_approval(self) -> None:
        environment = {
            "PLATFORM_DATABASE_BACKEND": "postgresql",
            "PLATFORM_DATABASE_URL": "postgresql://unified-example",
        }
        with patch.dict(os.environ, environment, clear=True):
            for config_type in CONFIG_TYPES:
                with self.subTest(config=config_type.__name__):
                    with self.assertRaisesRegex(RuntimeError, "formal switch approval"):
                        config_type.from_env()

    def test_platform_sqlite_profile_overrides_legacy_postgresql(self) -> None:
        environment = {
            "PLATFORM_DATABASE_BACKEND": "sqlite",
            "AUTOMATION_DATABASE_BACKEND": "postgresql",
            "AUTOMATION_DATABASE_URL": "postgresql://legacy-example",
            "AUTOMATION_POSTGRESQL_READ_WRITE_ENABLED": "true",
        }
        with patch.dict(os.environ, environment, clear=True):
            configs = [config_type.from_env() for config_type in CONFIG_TYPES]

        self.assertTrue(all(config.backend == "sqlite" for config in configs))
        self.assertTrue(all(config.database_url is None for config in configs))

    def test_emergency_legacy_profile_restores_module_sources(self) -> None:
        environment = {
            "PLATFORM_DATABASE_BACKEND": "legacy",
            "AUTOMATION_DATABASE_BACKEND": "postgresql",
            "AUTOMATION_DATABASE_URL": "postgresql://legacy-automation",
            "AUTOMATION_POSTGRESQL_READ_WRITE_ENABLED": "true",
        }
        with patch.dict(os.environ, environment, clear=True):
            automation = AutomationDatabaseConfig.from_env()
            identity = IdentityDatabaseConfig.from_env()
            transcode = TranscodeDatabaseConfig.from_env()
            configuration = ConfigurationDatabaseConfig.from_env()
            planning = PlanningDatabaseConfig.from_env()

        self.assertEqual("postgresql", automation.backend)
        self.assertEqual("postgresql://legacy-automation", automation.database_url)
        self.assertTrue(all(item.backend == "sqlite" for item in (identity, transcode, configuration, planning)))

    def test_rollback_helper_builds_local_docker_profile_without_persisting_it(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            env_path = Path(temporary_directory) / ".env.postgres.local"
            env_path.write_text(
                "POSTGRES_DB=legacy_db\nPOSTGRES_USER=legacy_user\n"
                "POSTGRES_PASSWORD=not-for-output\nPOSTGRES_PORT=55432\n",
                encoding="utf-8",
            )
            with patch.dict(os.environ, {}, clear=True):
                configure_legacy_rollback_environment(env_path=env_path)
                config = AutomationDatabaseConfig.from_env()

        self.assertEqual("postgresql", config.backend)
        self.assertTrue(str(config.database_url).startswith("postgresql://legacy_user:"))
        self.assertIn("@127.0.0.1:55432/legacy_db", str(config.database_url))


if __name__ == "__main__":
    unittest.main()
