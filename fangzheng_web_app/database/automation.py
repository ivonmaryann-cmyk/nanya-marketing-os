from __future__ import annotations

from contextlib import contextmanager
from threading import Lock
from typing import Any, Iterator

from ..db import _sqlite_db_cursor
from .config import AutomationDatabaseConfig
from .sql import qmark_to_pyformat, sqlite_to_postgresql


class PostgresResultAdapter:
    def __init__(self, cursor: Any, returns_identity: bool):
        self._cursor = cursor
        self._identity = None
        self._returns_identity = returns_identity

    @property
    def lastrowid(self):
        if not self._returns_identity:
            return None
        if self._identity is None:
            row = self._cursor.fetchone()
            self._identity = row["id"] if isinstance(row, dict) else row[0]
        return self._identity

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def fetchone(self):
        return self._cursor.fetchone()

    def fetchall(self):
        return self._cursor.fetchall()

    def __iter__(self):
        return iter(self._cursor)


class PostgresConnectionAdapter:
    dialect = "postgresql"

    def __init__(self, connection: Any, *, map_automation_metadata: bool = True):
        self._connection = connection
        self._map_automation_metadata = map_automation_metadata

    def execute(self, sql: str, params: tuple[Any, ...] | list[Any] = (), *, postgres_sql: str | None = None):
        if postgres_sql is not None:
            statement, returns_identity = qmark_to_pyformat(postgres_sql), False
        else:
            statement, returns_identity = sqlite_to_postgresql(
                sql, map_automation_metadata=self._map_automation_metadata
            )
        cursor = self._connection.execute(statement, params)
        return PostgresResultAdapter(cursor, returns_identity)

    def executemany(self, sql: str, params_seq, *, postgres_sql: str | None = None):
        statement = postgres_sql if postgres_sql is not None else qmark_to_pyformat(sql)
        cursor = self._connection.cursor()
        cursor.executemany(statement, params_seq)
        return PostgresResultAdapter(cursor, False)


_pool = None
_pool_key: tuple[object, ...] | None = None
_pool_lock = Lock()


def _get_pool(config: AutomationDatabaseConfig):
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
def automation_cursor(config: AutomationDatabaseConfig | None = None) -> Iterator[Any]:
    """Open an automation transaction without changing the core SQLite database."""
    resolved = config or AutomationDatabaseConfig.from_env()
    if resolved.backend == "sqlite":
        with _sqlite_db_cursor() as connection:
            yield connection
        return
    pool = _get_pool(resolved)
    with pool.connection() as connection:
        try:
            yield PostgresConnectionAdapter(connection)
            connection.commit()
        except Exception:
            connection.rollback()
            raise


@contextmanager
def platform_cursor(config: AutomationDatabaseConfig | None = None) -> Iterator[Any]:
    """Core platform cursor: preserve the global ``settings`` table name."""
    resolved = config or AutomationDatabaseConfig.from_env()
    if resolved.backend == "sqlite":
        with _sqlite_db_cursor() as connection:
            yield connection
        return
    pool = _get_pool(resolved)
    with pool.connection() as connection:
        try:
            yield PostgresConnectionAdapter(connection, map_automation_metadata=False)
            connection.commit()
        except Exception:
            connection.rollback()
            raise


def close_automation_pool() -> None:
    global _pool, _pool_key
    with _pool_lock:
        if _pool is not None:
            _pool.close()
        _pool = None
        _pool_key = None
