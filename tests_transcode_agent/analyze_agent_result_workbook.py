import argparse
from collections import Counter
import json

import openpyxl

from fangzheng_web_app.transcode_agent_service import _load_runtime, analyze_spec


def _text(value) -> str:
    return str(value or "").strip()


def _code22(value) -> str:
    text = _text(value).replace(" ", "")
    if not text or text.startswith(("未识别", "待确认", "跳过")):
        return ""
    return text.split("*")[0][:22]


def main() -> None:
    parser = argparse.ArgumentParser(description="复算营销转码Agent结果文件并按前22码比对")
    parser.add_argument("workbook")
    args = parser.parse_args()

    workbook = openpyxl.load_workbook(args.workbook, read_only=True, data_only=True)
    sheet = workbook["Sheet1"]
    rows = sheet.iter_rows(values_only=True)
    headers = [_text(value) for value in next(rows)]
    source_rows = [dict(zip(headers, values)) for values in rows]
    workbook.close()

    engine, tables, rules, mappings, _, _ = _load_runtime()
    counts = Counter()
    errors = []
    for excel_row, row in enumerate(source_rows, start=2):
        spec = _text(row.get("客户规格"))
        expected = _code22(row.get("品    名"))
        if not spec or not expected:
            counts["skip"] += 1
            continue
        analysis = analyze_spec(
            engine,
            tables,
            rules,
            spec,
            agent_mapping_tables=mappings,
            customer=_text(row.get("客户简称")),
            customer_code=_text(row.get("客户编号")),
            parse_fallback_text=_text(row.get("规格")),
            excel_row=excel_row,
        )
        actual = _code22(analysis.get("candidate_code"))
        if not actual:
            counts["not_emitted"] += 1
        elif actual == expected:
            counts["correct"] += 1
        else:
            counts["wrong"] += 1
            errors.append(
                {
                    "row": excel_row,
                    "customer": _text(row.get("客户简称")),
                    "spec": spec,
                    "expected": expected,
                    "actual": actual,
                    "status": analysis.get("status"),
                    "reason": analysis.get("reason"),
                }
            )

    emitted = counts["correct"] + counts["wrong"]
    effective = emitted + counts["not_emitted"]
    result = {
        "effective": effective,
        "emitted": emitted,
        "correct": counts["correct"],
        "wrong": counts["wrong"],
        "not_emitted": counts["not_emitted"],
        "skip": counts["skip"],
        "emitted_accuracy": round(counts["correct"] / emitted * 100, 4) if emitted else 0,
        "overall_accuracy": round(counts["correct"] / effective * 100, 4) if effective else 0,
        "errors": errors,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
