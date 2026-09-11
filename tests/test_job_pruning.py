from __future__ import annotations

import unittest
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

from fangzheng_web_app import db


class JobPruningTests(unittest.TestCase):
    def test_prune_uses_non_negative_cross_database_limit(self):
        connection = MagicMock()
        connection.execute.return_value.fetchall.return_value = []

        @contextmanager
        def cursor():
            yield connection

        with patch.object(db, "transcode_db_cursor", cursor):
            self.assertEqual(db.prune_jobs_for_employee("employee", keep_limit=500), [])

        sql, params = connection.execute.call_args.args
        self.assertNotIn("LIMIT -1", sql)
        self.assertEqual(params, ("employee", 2_147_483_647, 500))
