from __future__ import annotations

import math
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any

from psycopg_pool import ConnectionPool


READ_SCENARIOS = {
    "mail_list": "SELECT id,received_at FROM mail_messages ORDER BY received_at DESC,id DESC LIMIT 50",
    "case_counts": "SELECT status,action_type,COUNT(*) FROM order_intake_cases GROUP BY status,action_type",
    "template_versions": "SELECT template_id,version_number FROM order_entry_template_versions ORDER BY template_id,version_number DESC LIMIT 100",
    "outbox_inbox": "SELECT COUNT(*) FROM automation_migration_inbox",
}


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    return ordered[min(len(ordered) - 1, max(0, math.ceil(len(ordered) * percentile) - 1))]


def run_read_benchmark(database_url: str, *, concurrency: int = 30, iterations: int = 5) -> dict[str, Any]:
    if concurrency <= 0 or concurrency > 100:
        raise ValueError("concurrency must be between 1 and 100")
    if iterations <= 0 or iterations > 100:
        raise ValueError("iterations must be between 1 and 100")
    timings: list[float] = []
    errors: dict[str, int] = {}

    def execute(sql: str) -> float:
        started = time.perf_counter()
        with pool.connection() as connection:
            connection.execute(sql).fetchall()
        return (time.perf_counter() - started) * 1000

    pool = ConnectionPool(database_url, min_size=1, max_size=concurrency, timeout=10, open=True)
    try:
        work = [
            sql
            for _ in range(iterations)
            for _worker in range(concurrency)
            for sql in READ_SCENARIOS.values()
        ]
        with ThreadPoolExecutor(max_workers=concurrency) as executor:
            futures = [executor.submit(execute, sql) for sql in work]
            for future in as_completed(futures):
                try:
                    timings.append(future.result())
                except Exception as exc:
                    name = type(exc).__name__
                    errors[name] = errors.get(name, 0) + 1
    finally:
        pool.close()
    total = len(timings) + sum(errors.values())
    return {
        "concurrency": concurrency, "requests": total, "successes": len(timings),
        "errors": errors, "error_rate": (sum(errors.values()) / total if total else 0.0),
        "p50_ms": round(_percentile(timings, 0.50), 2),
        "p95_ms": round(_percentile(timings, 0.95), 2),
        "max_ms": round(max(timings, default=0.0), 2),
    }
