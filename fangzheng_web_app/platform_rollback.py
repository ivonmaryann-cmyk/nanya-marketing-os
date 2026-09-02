"""Emergency-only helpers for returning to the pre-cutover source profile.

The normal runtime has exactly one active ``PLATFORM_DATABASE_*`` profile.
This module is deliberately opt-in and is used only by the documented
rollback launcher.  It never persists credentials or edits ``local.env``.
"""

from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import quote

from .local_env import PROJECT_ROOT


LEGACY_AUTOMATION_ENV_PATH = PROJECT_ROOT / ".env.postgres.local"


def _read_env_file(path: Path) -> dict[str, str]:
    if not path.is_file():
        raise RuntimeError(f"回退所需的本机 Docker 配置不存在：{path.name}")
    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        if key:
            values[key] = value
    return values


def configure_legacy_rollback_environment(
    *,
    env_path: Path = LEGACY_AUTOMATION_ENV_PATH,
) -> None:
    """Configure only the current process for the verified legacy sources."""
    values = _read_env_file(env_path)
    required = ("POSTGRES_DB", "POSTGRES_USER", "POSTGRES_PASSWORD", "POSTGRES_PORT")
    missing = [key for key in required if not values.get(key)]
    if missing:
        raise RuntimeError("回退 Docker 配置缺少：" + ", ".join(missing))

    database_url = "postgresql://{user}:{password}@127.0.0.1:{port}/{database}".format(
        user=quote(values["POSTGRES_USER"], safe=""),
        password=quote(values["POSTGRES_PASSWORD"], safe=""),
        port=values["POSTGRES_PORT"],
        database=quote(values["POSTGRES_DB"], safe=""),
    )
    # ``legacy`` makes database.config resolve the historical source split:
    # SQLite for platform domains and Docker PostgreSQL for order automation.
    os.environ["PLATFORM_DATABASE_BACKEND"] = "legacy"
    os.environ["AUTOMATION_DATABASE_BACKEND"] = "postgresql"
    os.environ["AUTOMATION_POSTGRESQL_READ_WRITE_ENABLED"] = "true"
    os.environ["AUTOMATION_DATABASE_URL"] = database_url
