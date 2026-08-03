from __future__ import annotations

import sys
import time
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_glue_resolver import clear_agent_glue_index_cache
from fangzheng_web_app.transcode_agent_rules import load_transcode_agent_mapping_tables
from fangzheng_web_app.transcode_agent_service import _load_runtime, analyze_spec


BASELINE = Path(
    "/Users/ny/Desktop/南亚/产品包/转码采集规则源/CCL转码测试数据_业务确认基线_20260721.xlsx"
)
LEGACY_AGENT_VERSION = "transcode_agent_rules_20260721_111515"
SAMPLE_LIMIT = 200


def main() -> None:
    if not BASELINE.exists():
        print(f"agent glue performance skipped: missing {BASELINE}")
        return

    rows = _load_rows(BASELINE, SAMPLE_LIMIT)
    engine, tables, rules, current_mappings, _, _ = _load_runtime()
    legacy_mappings = load_transcode_agent_mapping_tables(LEGACY_AGENT_VERSION)

    legacy_seconds = _measure(engine, tables, rules, legacy_mappings, rows)
    current_seconds = _measure(engine, tables, rules, current_mappings, rows)
    milliseconds_per_row = current_seconds / len(rows) * 1000

    assert milliseconds_per_row <= 10, milliseconds_per_row
    assert current_seconds <= max(legacy_seconds * 2.5, legacy_seconds + 0.5), (
        legacy_seconds,
        current_seconds,
    )
    print(
        "agent glue performance passed "
        f"rows={len(rows)} legacy={legacy_seconds:.3f}s "
        f"current={current_seconds:.3f}s current_ms_per_row={milliseconds_per_row:.2f}"
    )


def _measure(engine, tables, rules, mappings, rows: list[tuple]) -> float:
    clear_agent_glue_index_cache()
    started = time.perf_counter()
    for row_number, customer_code, customer_name, spec in rows:
        analyze_spec(
            engine,
            tables,
            rules,
            spec,
            agent_mapping_tables=mappings,
            customer=customer_name,
            customer_code=customer_code,
            excel_row=row_number,
        )
    return time.perf_counter() - started


def _load_rows(path: Path, limit: int) -> list[tuple]:
    worksheet = openpyxl.load_workbook(path, read_only=True, data_only=True).active
    rows = []
    for row_number, values in enumerate(worksheet.iter_rows(min_row=2, values_only=True), 2):
        cells = list(values) + [None] * 4
        customer_code, customer_name, _, spec = cells[:4]
        if not str(spec or "").strip():
            continue
        rows.append(
            (
                row_number,
                str(customer_code or ""),
                str(customer_name or ""),
                str(spec),
            )
        )
        if len(rows) >= limit:
            break
    return rows


if __name__ == "__main__":
    main()
