from __future__ import annotations

import sqlite3
from contextlib import closing
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import psycopg
from psycopg.rows import dict_row

from .copy import _columns, _quote, _upsert_sql
from .outbox import OUTBOX_TABLE, ensure_sqlite_outbox, parse_primary_key


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def outbox_status(sqlite_path: Path) -> dict[str, int | str]:
    with closing(sqlite3.connect(sqlite_path)) as source:
        source.row_factory = sqlite3.Row
        ensure_sqlite_outbox(source)
        source.commit()
        row = source.execute(
            f"""SELECT COUNT(*) pending, COALESCE(SUM(CASE WHEN attempts>0 THEN 1 ELSE 0 END),0) retrying,
                       COALESCE(MAX(attempts),0) max_attempts, COALESCE(MIN(created_at),'') oldest_created_at
                FROM {OUTBOX_TABLE} WHERE processed_at IS NULL"""
        ).fetchone()
        return dict(row)


def _source_row(source: sqlite3.Connection, table: str, primary_key: dict[str, object]):
    if table == "automation_metadata":
        return source.execute("SELECT key, value FROM settings WHERE key=?", (primary_key["key"],)).fetchone()
    where = " AND ".join(f'{_quote(key)}=?' for key in primary_key)
    return source.execute(f'SELECT * FROM {_quote(table)} WHERE {where}', tuple(primary_key.values())).fetchone()


def _apply_event(source: sqlite3.Connection, target: Any, event: sqlite3.Row) -> bool:
    table = event["source_table"]
    primary_key = parse_primary_key(event["pk_json"], table)
    claimed = target.execute(
        "INSERT INTO automation_migration_inbox(event_id,source_table,operation) VALUES (%s,%s,%s) "
        "ON CONFLICT(event_id) DO NOTHING RETURNING event_id",
        (event["event_id"], table, event["operation"]),
    ).fetchone()
    if not claimed:
        return False
    row = None if event["operation"] == "delete" else _source_row(source, table, primary_key)
    if row is None:
        where = " AND ".join(f'{_quote(key)}=%s' for key in primary_key)
        target.execute(f'DELETE FROM {_quote(table)} WHERE {where}', tuple(primary_key.values()))
    elif table == "automation_metadata":
        target.execute(
            "INSERT INTO automation_metadata(key,value,updated_at) VALUES (%s,%s,%s) "
            "ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value,updated_at=EXCLUDED.updated_at",
            (row["key"], row["value"], _now()),
        )
    else:
        columns = _columns(source, table)
        target.execute(_upsert_sql(table, columns), tuple(row[column] for column in columns))
    return True


def _schedule_retry(source: sqlite3.Connection, event: sqlite3.Row, exc: Exception) -> None:
    attempts = int(event["attempts"]) + 1
    delay = min(3600, 2 ** min(attempts, 10))
    retry_at = (datetime.now(timezone.utc) + timedelta(seconds=delay)).isoformat(timespec="milliseconds")
    source.execute(
        f"UPDATE {OUTBOX_TABLE} SET attempts=?,last_error_type=?,next_attempt_at=? WHERE id=?",
        (attempts, type(exc).__name__, retry_at, event["id"]),
    )
    source.commit()


def process_outbox(sqlite_path: Path, database_url: str, *, batch_size: int = 100) -> dict[str, int]:
    if batch_size <= 0 or batch_size > 1000:
        raise ValueError("batch_size must be between 1 and 1000")
    result = {"selected": 0, "applied": 0, "duplicates": 0, "failed": 0}
    with closing(sqlite3.connect(sqlite_path, timeout=5)) as source:
        source.row_factory = sqlite3.Row
        ensure_sqlite_outbox(source)
        source.commit()
        events = source.execute(
            f"""SELECT * FROM {OUTBOX_TABLE}
                WHERE processed_at IS NULL AND (next_attempt_at IS NULL OR next_attempt_at<=?)
                ORDER BY id LIMIT ?""",
            (_now(), batch_size),
        ).fetchall()
        result["selected"] = len(events)
        if not events:
            return result
        try:
            target = psycopg.connect(database_url, row_factory=dict_row)
        except Exception as exc:
            _schedule_retry(source, events[0], exc)
            result["failed"] = 1
            return result
        try:
            for event in events:
                try:
                    with target.transaction():
                        applied = _apply_event(source, target, event)
                    source.execute(
                        f"UPDATE {OUTBOX_TABLE} SET processed_at=?,last_error_type='',next_attempt_at=NULL WHERE id=?",
                        (_now(), event["id"]),
                    )
                    source.commit()
                    result["applied" if applied else "duplicates"] += 1
                except Exception as exc:
                    _schedule_retry(source, event, exc)
                    result["failed"] += 1
                    break
        finally:
            target.close()
    return result
