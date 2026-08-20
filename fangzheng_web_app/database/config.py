from __future__ import annotations

import os
from dataclasses import dataclass


_BACKENDS = {"sqlite", "postgresql"}


def _positive_int(name: str, default: int) -> int:
    raw = os.getenv(name, str(default)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer") from exc
    if value <= 0:
        raise ValueError(f"{name} must be greater than zero")
    return value


@dataclass(frozen=True)
class AutomationDatabaseConfig:
    backend: str
    database_url: str | None
    pool_min_size: int
    pool_max_size: int
    connect_timeout_seconds: int
    statement_timeout_ms: int

    @classmethod
    def from_env(cls) -> "AutomationDatabaseConfig":
        backend = os.getenv("AUTOMATION_DATABASE_BACKEND", "sqlite").strip().lower()
        if backend == "postgres":
            backend = "postgresql"
        if backend not in _BACKENDS:
            raise ValueError("AUTOMATION_DATABASE_BACKEND must be sqlite or postgresql")
        database_url = os.getenv("AUTOMATION_DATABASE_URL", "").strip() or None
        if backend == "postgresql" and not database_url:
            raise ValueError("AUTOMATION_DATABASE_URL is required for the postgresql backend")
        read_write_enabled = os.getenv("AUTOMATION_POSTGRESQL_READ_WRITE_ENABLED", "false").strip().lower()
        if backend == "postgresql" and read_write_enabled not in {"1", "true", "yes"}:
            raise RuntimeError(
                "PostgreSQL application read/write is locked; formal switch approval is required"
            )
        pool_min = _positive_int("AUTOMATION_DB_POOL_MIN_SIZE", 1)
        pool_max = _positive_int("AUTOMATION_DB_POOL_MAX_SIZE", 8)
        if pool_min > pool_max:
            raise ValueError("AUTOMATION_DB_POOL_MIN_SIZE cannot exceed AUTOMATION_DB_POOL_MAX_SIZE")
        return cls(
            backend=backend,
            database_url=database_url,
            pool_min_size=pool_min,
            pool_max_size=pool_max,
            connect_timeout_seconds=_positive_int("AUTOMATION_DB_CONNECT_TIMEOUT_SECONDS", 10),
            statement_timeout_ms=_positive_int("AUTOMATION_DB_STATEMENT_TIMEOUT_MS", 30000),
        )

    def redacted_summary(self) -> dict[str, object]:
        return {
            "backend": self.backend,
            "database_url_configured": bool(self.database_url),
            "pool_min_size": self.pool_min_size,
            "pool_max_size": self.pool_max_size,
            "connect_timeout_seconds": self.connect_timeout_seconds,
            "statement_timeout_ms": self.statement_timeout_ms,
        }
