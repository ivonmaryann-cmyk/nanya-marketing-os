from __future__ import annotations

import re
from contextlib import contextmanager
from threading import Lock
from typing import Any, Iterator

from .automation import PostgresResultAdapter
from .config import PlanningDatabaseConfig
from .sql import qmark_to_pyformat


_IDENTITY_TABLES = {"task_categories", "personal_tasks", "feedback"}


class PlanningPostgresConnectionAdapter:
    dialect = "postgresql"

    def __init__(self, connection: Any):
        self._connection = connection

    def execute(self, sql: str, params: tuple[Any, ...] | list[Any] = ()):
        statement = qmark_to_pyformat(sql)
        ignore_match = re.match(
            r"\s*INSERT\s+OR\s+IGNORE\s+INTO\s+task_categories\b",
            statement,
            re.IGNORECASE,
        )
        if ignore_match:
            statement = re.sub(
                r"INSERT\s+OR\s+IGNORE", "INSERT", statement, count=1, flags=re.IGNORECASE
            )
            statement += " ON CONFLICT (employee_id, name) DO NOTHING"
        insert = re.match(
            r"\s*INSERT\s+INTO\s+([a-zA-Z_][a-zA-Z0-9_]*)\b",
            statement,
            re.IGNORECASE,
        )
        returns_identity = bool(
            insert
            and insert.group(1).lower() in _IDENTITY_TABLES
            and not re.search(r"\bRETURNING\b", statement, re.IGNORECASE)
        )
        if returns_identity:
            statement += " RETURNING id"
        return PostgresResultAdapter(self._connection.execute(statement, params), returns_identity)

    def executemany(self, sql: str, params_seq):
        statement = qmark_to_pyformat(sql)
        if re.match(r"\s*INSERT\s+OR\s+IGNORE\s+INTO\s+task_categories\b", statement, re.IGNORECASE):
            statement = re.sub(
                r"INSERT\s+OR\s+IGNORE", "INSERT", statement, count=1, flags=re.IGNORECASE
            )
            statement += " ON CONFLICT (employee_id, name) DO NOTHING"
        cursor = self._connection.cursor()
        cursor.executemany(statement, params_seq)
        return PostgresResultAdapter(cursor, False)


_pool = None
_pool_key: tuple[object, ...] | None = None
_pool_lock = Lock()


def _get_pool(config: PlanningDatabaseConfig):
    global _pool, _pool_key
    key = (config.database_url, config.pool_min_size, config.pool_max_size)
    with _pool_lock:
        if _pool is not None and _pool_key == key:
            return _pool
        if _pool is not None:
            _pool.close()
        from psycopg.rows import dict_row
        from psycopg_pool import ConnectionPool

        _pool = ConnectionPool(
            conninfo=config.database_url or "",
            min_size=config.pool_min_size,
            max_size=config.pool_max_size,
            timeout=config.connect_timeout_seconds,
            kwargs={
                "row_factory": dict_row,
                "options": f"-c statement_timeout={config.statement_timeout_ms}",
            },
            open=True,
        )
        _pool_key = key
        return _pool


@contextmanager
def planning_cursor(config: PlanningDatabaseConfig | None = None) -> Iterator[Any]:
    resolved = config or PlanningDatabaseConfig.from_env()
    if resolved.backend == "sqlite":
        from ..db import db_cursor

        with db_cursor() as connection:
            yield connection
        return
    pool = _get_pool(resolved)
    with pool.connection() as connection:
        try:
            yield PlanningPostgresConnectionAdapter(connection)
            connection.commit()
        except Exception:
            connection.rollback()
            raise


def close_planning_pool() -> None:
    global _pool, _pool_key
    with _pool_lock:
        if _pool is not None:
            _pool.close()
        _pool = None
        _pool_key = None
