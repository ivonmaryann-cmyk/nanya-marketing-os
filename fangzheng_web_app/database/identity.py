from __future__ import annotations

from contextlib import contextmanager
from threading import Lock
from typing import Any, Iterator

from .automation import PostgresConnectionAdapter
from .config import IdentityDatabaseConfig


_pool = None
_pool_key: tuple[object, ...] | None = None
_pool_lock = Lock()


def _get_pool(config: IdentityDatabaseConfig):
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
def identity_cursor(config: IdentityDatabaseConfig | None = None) -> Iterator[Any]:
    resolved = config or IdentityDatabaseConfig.from_env()
    if resolved.backend == "sqlite":
        from ..db import db_cursor

        with db_cursor() as connection:
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


def close_identity_pool() -> None:
    global _pool, _pool_key
    with _pool_lock:
        if _pool is not None:
            _pool.close()
        _pool = None
        _pool_key = None
