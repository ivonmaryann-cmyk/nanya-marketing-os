from __future__ import annotations

import sqlite3
from contextlib import closing
from pathlib import Path
from typing import Any

from .spec import IDENTITY_TABLES, PRIMARY_KEYS, TABLE_COLUMNS, TABLES


def _quote(identifier: str) -> str:
    if not identifier.replace("_", "").isalnum():
        raise ValueError(f"unsafe SQL identifier: {identifier}")
    return f'"{identifier}"'


def _upsert_sql(table: str) -> str:
    columns = TABLE_COLUMNS[table]
    keys = PRIMARY_KEYS[table]
    quoted = ",".join(_quote(column) for column in columns)
    placeholders = ",".join("%s" for _ in columns)
    conflict = ",".join(_quote(column) for column in keys)
    mutable = [column for column in columns if column not in keys]
    action = "DO NOTHING"
    if mutable:
        updates = ",".join(f'{_quote(column)}=EXCLUDED.{_quote(column)}' for column in mutable)
        action = f"DO UPDATE SET {updates}"
    return f"INSERT INTO {_quote(table)}({quoted}) VALUES ({placeholders}) ON CONFLICT ({conflict}) {action}"


def copy_snapshot(snapshot: Path, target: Any) -> dict[str, int]:
    counts: dict[str, int] = {}
    with closing(sqlite3.connect(snapshot)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only=ON")
        with target.cursor() as cursor:
            for table in TABLES:
                actual = {row["name"] for row in source.execute(f'PRAGMA table_info("{table}")')}
                if actual != set(TABLE_COLUMNS[table]):
                    raise RuntimeError(f"SQLite schema mismatch for scoped table: {table}")
                selected_columns = ",".join(_quote(column) for column in TABLE_COLUMNS[table])
                rows = source.execute(f'SELECT {selected_columns} FROM "{table}"').fetchall()
                if rows:
                    cursor.executemany(
                        _upsert_sql(table),
                        [tuple(row[column] for column in TABLE_COLUMNS[table]) for row in rows],
                    )
                counts[table] = len(rows)
    for table in IDENTITY_TABLES:
        target.execute(
            "SELECT setval(pg_get_serial_sequence(%s, 'id'), "
            f"GREATEST(COALESCE((SELECT MAX(id) FROM {_quote(table)}), 0), 1), "
            f"EXISTS(SELECT 1 FROM {_quote(table)}))",
            (table,),
        )
    return counts
