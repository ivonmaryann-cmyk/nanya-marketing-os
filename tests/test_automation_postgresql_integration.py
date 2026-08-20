from __future__ import annotations

import os
import unittest

import psycopg
from psycopg.rows import dict_row

from fangzheng_web_app.automation_migration.schema import apply_migrations
from fangzheng_web_app.automation_migration.spec import TABLES


@unittest.skipUnless(os.getenv("AUTOMATION_TEST_DATABASE_URL"), "AUTOMATION_TEST_DATABASE_URL is not configured")
class AutomationPostgresqlIntegrationTests(unittest.TestCase):
    def test_empty_database_migration_is_repeatable_and_scoped(self) -> None:
        with psycopg.connect(os.environ["AUTOMATION_TEST_DATABASE_URL"], row_factory=dict_row) as connection:
            first = apply_migrations(connection)
            second = apply_migrations(connection)
            self.assertTrue(first or second == [])
            self.assertEqual([], second)
            rows = connection.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_schema='public'"
            ).fetchall()
            names = {row["table_name"] for row in rows}
            self.assertTrue(set(TABLES).issubset(names))


if __name__ == "__main__":
    unittest.main()
