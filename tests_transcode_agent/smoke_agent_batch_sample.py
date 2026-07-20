from __future__ import annotations

from collections import Counter
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.excel_utils import load_workbook_compat, normalized_xlsx_source
from fangzheng_web_app.transcode_agent_rules import parse_customer_special_master
from fangzheng_web_app.transcode_agent_service import analyze_spec, load_transcode_module
from fangzheng_web_app.transcode_rules import get_active_transcode_rule_version, get_transcode_rule_file_path


DRAFT_PATH = ROOT / "docs/develop0707/客户特殊规则结构化草稿_按原表行_20260708.xlsx"
SAMPLE_PATH = ROOT / "tests_transcode_agent/fixtures/transcode_agent_499_regression.xlsx"
EXPECTED_RULE_COUNT = 200
FORBIDDEN_RULE_TERMS = ("订单", "第5码", "第5位", "客户订单", "订单内容", "订单字段")


def main() -> None:
    assert DRAFT_PATH.exists(), DRAFT_PATH
    assert SAMPLE_PATH.exists(), SAMPLE_PATH

    engine = load_transcode_module()
    base_version = get_active_transcode_rule_version()
    rule_path = get_transcode_rule_file_path(base_version)
    tables = engine.build_lookup_tables(engine.load_rule_sheets(str(rule_path)))
    tables["structured_special_rules"] = []

    agent_rules, summary = parse_customer_special_master(DRAFT_PATH)
    assert summary["source_type"] == "confirmed_structured_draft"
    assert len(agent_rules) == EXPECTED_RULE_COUNT, len(agent_rules)

    workbook = load_workbook_compat(str(SAMPLE_PATH), data_only=True)
    source_for_result = normalized_xlsx_source(str(SAMPLE_PATH), workbook)
    sheets, _ = engine.load_transcode_inputs(str(source_for_result), str(rule_path))
    df_req = sheets["转码需求表"].copy()

    spec_col = engine.select_transcode_spec_column(df_req)
    customer_col = engine.detect_customer_column(df_req, spec_col)
    customer_code_col = engine.detect_customer_code_column(df_req)
    context_cols = engine.detect_transcode_context_columns(df_req, spec_col, customer_col, customer_code_col)
    data_indices = [i for i in range(1, len(df_req)) if _is_effective_spec(df_req.iloc[i, spec_col], engine)]

    analyses = []
    skipped_pp = 0
    for i in data_indices:
        row = df_req.iloc[i]
        customer_code = _cell(row, customer_code_col, engine) if customer_code_col is not None else ""
        customer = _cell(row, customer_col, engine) if customer_col is not None else ""
        spec = _cell(row, spec_col, engine)
        context = engine.build_context_text_from_row(row, context_cols)
        cust_spec = _cell(row, 6, engine)
        normalized_spec = _cell(row, 7, engine)
        pp_check_text = " ".join([spec, cust_spec, normalized_spec])

        if engine.is_pp_or_rc_spec(pp_check_text):
            skipped_pp += 1
            continue

        analyses.append(
            analyze_spec(
                engine,
                tables,
                agent_rules,
                spec,
                customer=customer,
                customer_code=customer_code,
                context_text=context,
                excel_row=i + 1,
            )
        )

    status_counter = Counter(item["status"] for item in analyses)
    applied_counter = Counter()
    applied_field_counter = Counter()
    applied_customer_counter = Counter()
    forbidden_hits = []
    struct_hits = []
    pp_hits = []
    conflict_rows = []

    for analysis in analyses:
        if analysis.get("conflicts"):
            conflict_rows.append((analysis.get("row"), analysis.get("customer"), analysis.get("conflicts")))
        for applied in analysis.get("applied_rules", []):
            applied_counter[applied.get("rule_id", "")] += 1
            applied_field_counter[applied.get("field", "")] += 1
            applied_customer_counter[analysis.get("customer", "")] += 1
            text = f"{applied.get('text', '')} {applied.get('source', '')}"
            if applied.get("field") == "struct_code":
                struct_hits.append((analysis.get("row"), applied))
            if any(term in text for term in FORBIDDEN_RULE_TERMS):
                forbidden_hits.append((analysis.get("row"), applied))
        if engine.is_pp_or_rc_spec(str(analysis.get("spec", ""))) and analysis.get("applied_rules"):
            pp_hits.append((analysis.get("row"), analysis.get("applied_rules")))

    assert data_indices, "no effective rows found in sample"
    assert analyses, "no CCL rows analyzed in sample"
    assert not struct_hits, struct_hits
    assert not forbidden_hits, forbidden_hits
    assert not pp_hits, pp_hits
    assert not conflict_rows, conflict_rows[:5]

    total_applied = sum(applied_counter.values())
    print(
        "batch sample smoke passed "
        f"rows={len(data_indices)} analyzed={len(analyses)} skipped_pp={skipped_pp} "
        f"status={dict(status_counter)} agent_rule_hits={total_applied} "
        f"fields={dict(applied_field_counter)} top_customers={applied_customer_counter.most_common(8)} "
        f"top_rules={applied_counter.most_common(8)}"
    )


def _cell(row, index: int | None, engine) -> str:
    if index is None or index >= len(row):
        return ""
    return engine._clean_cell(row.iloc[index])


def _is_effective_spec(value, engine) -> bool:
    text = engine._clean_cell(value)
    if not text:
        return False
    return text.lower() not in ("nan", "客户规格", "规格", "品名")


if __name__ == "__main__":
    main()
