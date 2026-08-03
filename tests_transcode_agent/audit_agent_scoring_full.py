from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import (  # noqa: E402
    _calculate_transcode_agent_analysis,
    _load_runtime,
)


FIELD_SLICES = {
    "glue": slice(0, 2),
    "thickness": slice(2, 7),
    "copper": slice(7, 9),
    "size": slice(9, 17),
    "glue_category": slice(17, 18),
    "copper_type": slice(18, 19),
    "grade": slice(19, 21),
    "total_core": slice(21, 22),
}


def main() -> None:
    parser = argparse.ArgumentParser(description="营销转码Agent全量评分与正确码审计")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument(
        "--summary-only",
        action="store_true",
        help="只输出行数、状态、准确率和评分异常等核心汇总",
    )
    args = parser.parse_args()
    result = audit_workbook(args.workbook, limit=args.limit)
    printable = _summary(result) if args.summary_only else result
    print(json.dumps(printable, ensure_ascii=False, indent=2))
    assert result["scoring_anomaly_count"] == 0, result["scoring_anomalies"][:10]


def audit_workbook(path: Path, *, limit: int = 0) -> dict:
    if not path.exists():
        raise FileNotFoundError(path)
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    headers = [_normalize_header(cell.value) for cell in sheet[1]]
    columns = {
        "customer_code": _find_column(headers, ("客户编号", "客户代码")),
        "customer": _find_column(headers, ("客户简称", "客户名称")),
        "correct": _find_column(headers, ("品名", "正确码", "正确码值")),
        "spec": _find_column(headers, ("客户规格", "单条客户规格")),
        "remark": _find_column(headers, ("订单备注", "备注整行上下文", "备注")),
    }
    if columns["spec"] is None:
        raise ValueError(f"未找到客户规格列: {headers}")

    runtime = _load_runtime()
    engine = runtime[0]
    active_machine_rule_ids = {
        str(rule.get("规则ID") or "").strip()
        for rule in runtime[2]
        if str(rule.get("启用") or "").strip() == "是"
        and str(rule.get("待确认") or "").strip() != "是"
        and str(rule.get("强制执行") or "").strip() == "是"
        and str(rule.get("物料类别") or "").strip() == "CCL"
        and str(rule.get("规则ID") or "").strip()
    }
    status_counts: Counter[str] = Counter()
    decision_counts: Counter[str] = Counter()
    wrong_fields: Counter[str] = Counter()
    wrong_customers: Counter[str] = Counter()
    wrong_transitions: Counter[str] = Counter()
    mismatch_samples: list[dict] = []
    failure_fields: Counter[str] = Counter()
    failure_combinations: Counter[str] = Counter()
    conflict_reasons: Counter[str] = Counter()
    applied_rule_hits: Counter[str] = Counter()
    field_score_distribution: Counter[str] = Counter()
    scoring_anomalies: list[dict] = []
    formal_correct = formal_wrong = 0
    pending_correct = pending_wrong = 0
    skipped = effective = 0

    for excel_row, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
        if limit and effective >= limit:
            break
        spec = _value(values, columns["spec"])
        if not spec or _normalize_header(spec) in {"客户规格", "规格", "品名"}:
            continue
        if engine.is_pp_or_rc_spec(spec):
            skipped += 1
            continue
        effective += 1
        customer = _value(values, columns["customer"])
        customer_code = _value(values, columns["customer_code"])
        remark = _value(values, columns["remark"])
        correct = _code22(_value(values, columns["correct"]))
        analysis, _base_version, _agent_version = _calculate_transcode_agent_analysis(
            spec,
            customer=customer,
            customer_code=customer_code,
            order_remark=remark,
            employee_id="scoring-audit",
            runtime=runtime,
        )
        status = str(analysis.get("status") or "")
        score = int(analysis.get("overall_score") or 0)
        formal = _code22(analysis.get("formal_code"))
        candidate = _code22(analysis.get("candidate_code"))
        conflicts = list(analysis.get("conflicts") or [])
        evidence = [item for item in analysis.get("field_evidence") or [] if item.get("gate")]
        for item in evidence:
            field_score_distribution[
                f"{item.get('field_key') or 'unknown'}:{int(item.get('score') or 0)}"
            ] += 1
        for rule in analysis.get("applied_rules") or []:
            rule_id = str(rule.get("rule_id") or "").strip()
            if rule_id:
                applied_rule_hits[rule_id] += 1
        low_fields = [
            str(item.get("field_key") or "unknown")
            for item in evidence
            if int(item.get("score") or 0) < 100
        ]
        status_counts[status] += 1
        decision_counts[_decision_name(status, conflicts)] += 1

        anomalies = []
        if status == "成功":
            if score != 100:
                anomalies.append("正式出码但整行分数不是100")
            if not formal:
                anomalies.append("正式出码状态缺少正式码")
            if conflicts:
                anomalies.append("正式出码仍存在规则冲突")
            if low_fields:
                anomalies.append(f"正式出码存在低分字段:{','.join(low_fields)}")
        else:
            if formal:
                anomalies.append("非成功状态仍写入正式码")
            if status == "待确认" and score == 100 and not conflicts:
                anomalies.append("100分且无冲突仍进入待确认")
        if anomalies:
            scoring_anomalies.append(
                {"row": excel_row, "customer": customer, "status": status, "issues": anomalies}
            )

        if status == "失败":
            failure_fields.update(low_fields)
            failure_combinations["+".join(sorted(low_fields)) or "unknown"] += 1
        if conflicts:
            for conflict in conflicts:
                conflict_reasons[_compact_conflict(conflict)] += 1

        if correct:
            compared = formal if status == "成功" else candidate if status == "待确认" else ""
            if status == "成功":
                if compared == correct:
                    formal_correct += 1
                else:
                    formal_wrong += 1
                    different_fields = _different_fields(compared, correct)
                    wrong_fields.update(different_fields)
                    wrong_customers[customer or "未提供客户"] += 1
                    for field in different_fields:
                        part = FIELD_SLICES[field]
                        wrong_transitions[
                            f"{field}:{compared[part]}->{correct[part]}"
                        ] += 1
                    if len(mismatch_samples) < 30:
                        mismatch_samples.append(
                            {
                                "row": excel_row,
                                "customer": customer,
                                "spec": spec,
                                "actual": compared,
                                "expected": correct,
                                "fields": different_fields,
                            }
                        )
            elif status == "待确认":
                if compared == correct:
                    pending_correct += 1
                else:
                    pending_wrong += 1
                    wrong_fields.update(_different_fields(compared, correct))

    workbook.close()
    formal_compared = formal_correct + formal_wrong
    pending_compared = pending_correct + pending_wrong
    hit_machine_rule_ids = active_machine_rule_ids.intersection(applied_rule_hits)
    return {
        "file": str(path),
        "rows_analyzed": effective,
        "rows_skipped_pp_rc": skipped,
        "status_counts": dict(status_counts),
        "decision_counts": dict(decision_counts),
        "formal_correct": formal_correct,
        "formal_wrong": formal_wrong,
        "formal_accuracy_pct": _percent(formal_correct, formal_compared),
        "pending_candidate_correct": pending_correct,
        "pending_candidate_wrong": pending_wrong,
        "pending_candidate_accuracy_pct": _percent(pending_correct, pending_compared),
        "wrong_fields": dict(wrong_fields.most_common()),
        "wrong_customers": dict(wrong_customers.most_common(20)),
        "wrong_transitions": dict(wrong_transitions.most_common(30)),
        "formal_mismatch_samples": mismatch_samples,
        "failure_fields": dict(failure_fields.most_common()),
        "failure_combinations": dict(failure_combinations.most_common(12)),
        "conflict_reasons": dict(conflict_reasons.most_common(12)),
        "active_machine_rule_count": len(active_machine_rule_ids),
        "hit_machine_rule_count": len(hit_machine_rule_ids),
        "unhit_machine_rule_count": len(active_machine_rule_ids - hit_machine_rule_ids),
        "unhit_machine_rule_ids": sorted(active_machine_rule_ids - hit_machine_rule_ids),
        "top_applied_rule_hits": dict(applied_rule_hits.most_common(30)),
        "field_score_distribution": dict(field_score_distribution),
        "scoring_anomaly_count": len(scoring_anomalies),
        "scoring_anomalies": scoring_anomalies[:100],
    }


def _summary(result: dict) -> dict:
    keys = (
        "file",
        "rows_analyzed",
        "rows_skipped_pp_rc",
        "status_counts",
        "decision_counts",
        "formal_correct",
        "formal_wrong",
        "formal_accuracy_pct",
        "pending_candidate_correct",
        "pending_candidate_wrong",
        "pending_candidate_accuracy_pct",
        "wrong_fields",
        "failure_fields",
        "active_machine_rule_count",
        "hit_machine_rule_count",
        "unhit_machine_rule_count",
        "scoring_anomaly_count",
        "scoring_anomalies",
    )
    return {key: result[key] for key in keys}


def _find_column(headers: list[str], aliases: tuple[str, ...]) -> int | None:
    normalized_aliases = {_normalize_header(alias) for alias in aliases}
    for index, header in enumerate(headers):
        if header in normalized_aliases:
            return index
    return None


def _normalize_header(value) -> str:
    return re.sub(r"\s+", "", str(value or "")).strip()


def _value(values, index: int | None) -> str:
    if index is None or index >= len(values):
        return ""
    return str(values[index] or "").strip()


def _code22(value) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()[:22]


def _different_fields(actual: str, expected: str) -> list[str]:
    return [name for name, part in FIELD_SLICES.items() if actual[part] != expected[part]]


def _decision_name(status: str, conflicts: list) -> str:
    if conflicts:
        return "规则冲突"
    if status == "失败":
        return "基础字段缺失"
    if status == "成功":
        return "确定"
    return status or "未知"


def _compact_conflict(conflict) -> str:
    text = str(conflict or "")
    match = re.search(r"(NY[-A-Z0-9()]+).*?候选([^；]+)", text, flags=re.I)
    return f"{match.group(1).upper()}候选{match.group(2)}" if match else text[:100]


def _percent(numerator: int, denominator: int) -> float | None:
    return round(numerator * 100 / denominator, 4) if denominator else None


if __name__ == "__main__":
    main()
