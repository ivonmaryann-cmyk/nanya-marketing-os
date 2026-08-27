from __future__ import annotations

import sqlite3
import unittest

from fangzheng_web_app.transcode_migration.copy import _source_rows
from fangzheng_web_app.transcode_migration.spec import TABLE_COLUMNS


class TranscodeMigrationCopyTests(unittest.TestCase):
    def test_explicitly_excluded_job_ids_are_not_selected(self) -> None:
        with sqlite3.connect(":memory:") as source:
            columns = ", ".join(f'"{column}" TEXT' for column in TABLE_COLUMNS["jobs"])
            source.execute(f'CREATE TABLE jobs ({columns})')
            values = [None] * len(TABLE_COLUMNS["jobs"])
            values[0] = 33
            source.execute(
                f'INSERT INTO jobs VALUES ({",".join("?" for _ in values)})',
                values,
            )
            values[0] = 34
            source.execute(
                f'INSERT INTO jobs VALUES ({",".join("?" for _ in values)})',
                values,
            )
            source.row_factory = sqlite3.Row

            rows = _source_rows(source, "jobs", (33,))

        self.assertEqual(["34"], [row["id"] for row in rows])


if __name__ == "__main__":
    unittest.main()
