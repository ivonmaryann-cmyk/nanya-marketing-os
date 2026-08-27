from __future__ import annotations

import os
from dataclasses import dataclass


_BACKENDS = {"sqlite", "postgresql"}
_ENABLED_VALUES = {"1", "true", "yes"}


def _positive_int(name: str, default: int) -> int:
    raw = os.getenv(name, str(default)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer") from exc
    if value <= 0:
        raise ValueError(f"{name} must be greater than zero")
    return value


def _resolve_database_settings(
    backend_name: str,
    url_name: str,
    read_write_name: str,
) -> tuple[str, str | None, str, str]:
    """Resolve one platform-wide database profile before legacy module settings.

    ``PLATFORM_DATABASE_BACKEND`` is the explicit switch. When it is present,
    every module must use the same URL and read/write approval flag; a partial
    mix of platform and module connection settings is deliberately forbidden.
    ``legacy`` is an emergency-only exception used by the documented rollback
    command: it intentionally restores the pre-cutover module source profile.
    Omitting the platform switch preserves the existing module-level behavior.
    """
    platform_backend = os.getenv("PLATFORM_DATABASE_BACKEND", "").strip()
    if platform_backend:
        if platform_backend.lower() == "legacy":
            backend = os.getenv(backend_name, "sqlite").strip().lower()
            database_url = os.getenv(url_name, "").strip() or None
            read_write_enabled = os.getenv(read_write_name, "false").strip().lower()
            return backend, database_url, read_write_enabled, backend_name.removesuffix("_BACKEND")
        backend = platform_backend.lower()
        database_url = os.getenv("PLATFORM_DATABASE_URL", "").strip() or None
        read_write_enabled = os.getenv(
            "PLATFORM_DATABASE_READ_WRITE_ENABLED", "false"
        ).strip().lower()
        return backend, database_url, read_write_enabled, "PLATFORM_DATABASE"

    backend = os.getenv(backend_name, "sqlite").strip().lower()
    database_url = os.getenv(url_name, "").strip() or None
    read_write_enabled = os.getenv(read_write_name, "false").strip().lower()
    return backend, database_url, read_write_enabled, backend_name.removesuffix("_BACKEND")


def _validated_backend(value: str, setting_name: str) -> str:
    backend = "postgresql" if value == "postgres" else value
    if backend not in _BACKENDS:
        raise ValueError(f"{setting_name} must be sqlite or postgresql")
    return backend


def _require_postgresql_access(
    backend: str,
    database_url: str | None,
    read_write_enabled: str,
    setting_prefix: str,
    module_label: str,
) -> None:
    if backend != "postgresql":
        return
    url_name = f"{setting_prefix}_URL"
    if not database_url:
        raise ValueError(f"{url_name} is required for the postgresql backend")
    if read_write_enabled not in _ENABLED_VALUES:
        raise RuntimeError(
            f"PostgreSQL {module_label} read/write is locked; formal switch approval is required"
        )


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
        raw_backend, database_url, read_write_enabled, setting_prefix = _resolve_database_settings(
            "AUTOMATION_DATABASE_BACKEND",
            "AUTOMATION_DATABASE_URL",
            "AUTOMATION_POSTGRESQL_READ_WRITE_ENABLED",
        )
        backend = _validated_backend(raw_backend, f"{setting_prefix}_BACKEND")
        _require_postgresql_access(
            backend, database_url, read_write_enabled, setting_prefix, "application"
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


@dataclass(frozen=True)
class IdentityDatabaseConfig:
    backend: str
    database_url: str | None
    pool_min_size: int
    pool_max_size: int
    connect_timeout_seconds: int
    statement_timeout_ms: int

    @classmethod
    def from_env(cls) -> "IdentityDatabaseConfig":
        raw_backend, database_url, read_write_enabled, setting_prefix = _resolve_database_settings(
            "IDENTITY_DATABASE_BACKEND",
            "IDENTITY_DATABASE_URL",
            "IDENTITY_POSTGRESQL_READ_WRITE_ENABLED",
        )
        backend = _validated_backend(raw_backend, f"{setting_prefix}_BACKEND")
        _require_postgresql_access(
            backend, database_url, read_write_enabled, setting_prefix, "identity"
        )
        pool_min = _positive_int("IDENTITY_DB_POOL_MIN_SIZE", 1)
        pool_max = _positive_int("IDENTITY_DB_POOL_MAX_SIZE", 4)
        if pool_min > pool_max:
            raise ValueError("IDENTITY_DB_POOL_MIN_SIZE cannot exceed IDENTITY_DB_POOL_MAX_SIZE")
        return cls(
            backend=backend,
            database_url=database_url,
            pool_min_size=pool_min,
            pool_max_size=pool_max,
            connect_timeout_seconds=_positive_int("IDENTITY_DB_CONNECT_TIMEOUT_SECONDS", 10),
            statement_timeout_ms=_positive_int("IDENTITY_DB_STATEMENT_TIMEOUT_MS", 30000),
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


@dataclass(frozen=True)
class TranscodeDatabaseConfig:
    backend: str
    database_url: str | None
    pool_min_size: int
    pool_max_size: int
    connect_timeout_seconds: int
    statement_timeout_ms: int

    @classmethod
    def from_env(cls) -> "TranscodeDatabaseConfig":
        raw_backend, database_url, read_write_enabled, setting_prefix = _resolve_database_settings(
            "TRANSCODE_DATABASE_BACKEND",
            "TRANSCODE_DATABASE_URL",
            "TRANSCODE_POSTGRESQL_READ_WRITE_ENABLED",
        )
        backend = _validated_backend(raw_backend, f"{setting_prefix}_BACKEND")
        _require_postgresql_access(
            backend, database_url, read_write_enabled, setting_prefix, "transcode"
        )
        pool_min = _positive_int("TRANSCODE_DB_POOL_MIN_SIZE", 1)
        pool_max = _positive_int("TRANSCODE_DB_POOL_MAX_SIZE", 8)
        if pool_min > pool_max:
            raise ValueError("TRANSCODE_DB_POOL_MIN_SIZE cannot exceed TRANSCODE_DB_POOL_MAX_SIZE")
        return cls(
            backend=backend,
            database_url=database_url,
            pool_min_size=pool_min,
            pool_max_size=pool_max,
            connect_timeout_seconds=_positive_int("TRANSCODE_DB_CONNECT_TIMEOUT_SECONDS", 10),
            statement_timeout_ms=_positive_int("TRANSCODE_DB_STATEMENT_TIMEOUT_MS", 30000),
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


@dataclass(frozen=True)
class ConfigurationDatabaseConfig:
    backend: str
    database_url: str | None
    pool_min_size: int
    pool_max_size: int
    connect_timeout_seconds: int
    statement_timeout_ms: int

    @classmethod
    def from_env(cls) -> "ConfigurationDatabaseConfig":
        raw_backend, database_url, read_write_enabled, setting_prefix = _resolve_database_settings(
            "CONFIG_DATABASE_BACKEND",
            "CONFIG_DATABASE_URL",
            "CONFIG_POSTGRESQL_READ_WRITE_ENABLED",
        )
        backend = _validated_backend(raw_backend, f"{setting_prefix}_BACKEND")
        _require_postgresql_access(
            backend, database_url, read_write_enabled, setting_prefix, "configuration"
        )
        pool_min = _positive_int("CONFIG_DB_POOL_MIN_SIZE", 1)
        pool_max = _positive_int("CONFIG_DB_POOL_MAX_SIZE", 4)
        if pool_min > pool_max:
            raise ValueError("CONFIG_DB_POOL_MIN_SIZE cannot exceed CONFIG_DB_POOL_MAX_SIZE")
        return cls(
            backend=backend,
            database_url=database_url,
            pool_min_size=pool_min,
            pool_max_size=pool_max,
            connect_timeout_seconds=_positive_int("CONFIG_DB_CONNECT_TIMEOUT_SECONDS", 10),
            statement_timeout_ms=_positive_int("CONFIG_DB_STATEMENT_TIMEOUT_MS", 30000),
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


@dataclass(frozen=True)
class PlanningDatabaseConfig:
    backend: str
    database_url: str | None
    pool_min_size: int
    pool_max_size: int
    connect_timeout_seconds: int
    statement_timeout_ms: int

    @classmethod
    def from_env(cls) -> "PlanningDatabaseConfig":
        raw_backend, database_url, read_write_enabled, setting_prefix = _resolve_database_settings(
            "PLANNING_DATABASE_BACKEND",
            "PLANNING_DATABASE_URL",
            "PLANNING_POSTGRESQL_READ_WRITE_ENABLED",
        )
        backend = _validated_backend(raw_backend, f"{setting_prefix}_BACKEND")
        _require_postgresql_access(
            backend, database_url, read_write_enabled, setting_prefix, "planning"
        )
        pool_min = _positive_int("PLANNING_DB_POOL_MIN_SIZE", 1)
        pool_max = _positive_int("PLANNING_DB_POOL_MAX_SIZE", 4)
        if pool_min > pool_max:
            raise ValueError("PLANNING_DB_POOL_MIN_SIZE cannot exceed PLANNING_DB_POOL_MAX_SIZE")
        return cls(
            backend=backend,
            database_url=database_url,
            pool_min_size=pool_min,
            pool_max_size=pool_max,
            connect_timeout_seconds=_positive_int("PLANNING_DB_CONNECT_TIMEOUT_SECONDS", 10),
            statement_timeout_ms=_positive_int("PLANNING_DB_STATEMENT_TIMEOUT_MS", 30000),
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
