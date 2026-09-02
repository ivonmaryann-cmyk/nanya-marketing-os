from __future__ import annotations

import hashlib
from pathlib import Path


MIGRATION_DIR = Path(__file__).resolve().parents[2] / "migrations" / "transcode" / "postgresql"


def apply_migrations(connection) -> list[str]:
    connection.execute(
        "CREATE TABLE IF NOT EXISTS transcode_schema_migrations "
        "(version TEXT PRIMARY KEY, checksum TEXT NOT NULL, "
        "applied_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP)"
    )
    applied: list[str] = []
    for path in sorted(MIGRATION_DIR.glob("*.sql")):
        sql = path.read_text(encoding="utf-8")
        checksum = hashlib.sha256(sql.encode("utf-8")).hexdigest()
        row = connection.execute(
            "SELECT checksum FROM transcode_schema_migrations WHERE version=%s", (path.name,)
        ).fetchone()
        if row:
            if row["checksum"] != checksum:
                raise RuntimeError(f"transcode migration checksum changed: {path.name}")
            continue
        connection.execute(sql)
        connection.execute(
            "INSERT INTO transcode_schema_migrations(version,checksum) VALUES (%s,%s)",
            (path.name, checksum),
        )
        applied.append(path.name)
    return applied
