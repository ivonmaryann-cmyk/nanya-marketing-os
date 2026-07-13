from __future__ import annotations

import copy
from decimal import Decimal
from typing import Any, Callable

from .ai_repair_config import get_ai_repair_config
from .deepseek_repair_client import DeepSeekRepairError, request_repair_json
from .purchase_field_rules import STANDARD_HEADERS, clean_text, decimal_or_none, normalize_date, normalize_number


SEQ = "序号"
CODE = "物料编码"
NAME = "物料名称"
DESC = "说明"
QTY = "数量"
UNIT = "单位"
PRICE = "含税单价"
AMOUNT = "金额"
DATE = "交货日期"
REMARK = "备注"

ALLOWED_FIELDS = set(STANDARD_HEADERS)
CRITICAL_FIELDS = [CODE, NAME, QTY, UNIT, PRICE, AMOUNT, DATE]
NUMERIC_FIELDS = {QTY, PRICE, AMOUNT}


def _money(value: Any) -> Decimal | None:
    return decimal_or_none(value)


def _amount_mismatch(standard: dict[str, Any]) -> bool:
    qty = _money(standard.get(QTY))
    price = _money(standard.get(PRICE))
    amount = _money(standard.get(AMOUNT))
    return qty is not None and price is not None and amount is not None and abs(qty * price - amount) > Decimal("0.05")


def _missing_fields(standard: dict[str, Any]) -> list[str]:
    return [field for field in CRITICAL_FIELDS if not clean_text(standard.get(field))]


def _row_needs_repair(row: dict[str, Any], related_issues: list[dict[str, Any]]) -> bool:
    standard = row.get("standard") or {}
    missing = _missing_fields(standard)
    if len(missing) >= 2:
        return True
    if _amount_mismatch(standard):
        return True
    if related_issues and missing:
        return True
    return False


def _issue_matches_row(issue: dict[str, Any], row: dict[str, Any]) -> bool:
    issue_text = clean_text(issue.get("raw_value"))
    row_text = clean_text(row.get("raw_text"))
    if not issue_text or not row_text:
        return False
    return row_text[:40] in issue_text or issue_text[:40] in row_text


def _related_issues(document: dict[str, Any], row: dict[str, Any]) -> list[dict[str, Any]]:
    issues = []
    for issue in document.get("issues") or []:
        if issue.get("page_index") not in {None, "", row.get("page_index")}:
            continue
        if _issue_matches_row(issue, row):
            issues.append(issue)
    return issues[:3]


def _row_number(value: Any) -> int | None:
    try:
        return int(float(str(value).split(".")[0]))
    except (TypeError, ValueError):
        return None


def _nearby_raw_rows(document: dict[str, Any], row: dict[str, Any]) -> list[list[str]]:
    page_index = row.get("page_index")
    table_index = row.get("table_index")
    row_number = _row_number(row.get("row_index"))
    for table in document.get("raw_detail_tables") or []:
        if table.get("page_index") != page_index or table.get("table_index") != table_index:
            continue
        rows = table.get("rows") or []
        if row_number is None:
            return rows[:8]
        start = max(0, row_number - 2)
        end = min(len(rows), row_number + 3)
        return rows[start:end]
    return []


def _build_repair_payload(document: dict[str, Any], candidates: list[dict[str, Any]]) -> dict[str, Any]:
    rows = []
    for item in candidates:
        row = item["row"]
        rows.append(
            {
                "row_key": item["row_key"],
                "page_index": row.get("page_index"),
                "table_index": row.get("table_index"),
                "row_index": row.get("row_index"),
                "raw_text": clean_text(row.get("raw_text"))[:800],
                "original": row.get("original") or {},
                "standard": row.get("standard") or {},
                "missing_fields": item["missing_fields"],
                "suspect_fields": item["suspect_fields"],
                "nearby_raw_rows": _nearby_raw_rows(document, row),
                "related_issues": item["issues"],
            }
        )
    return {
        "task": "repair_purchase_order_body_only",
        "source_file": document.get("source_file", ""),
        "allowed_fields": list(STANDARD_HEADERS),
        "rules": [
            "只修复订单正文中的明细行，不处理订单头、付款、条款、签核区。",
            "只能根据 raw_text、nearby_raw_rows、original、standard 中能看到的证据补字段或替换明显错误字段。",
            "不要凭空新增不存在的物料、数量、价格、金额或日期。",
            "金额以原文金额优先；原文金额缺失但数量和单价明确时，可以给出金额建议。",
            "返回严格 JSON，格式为 {\"repairs\":[{\"row_key\":0,\"set_fields\":{\"字段\":\"值\"},\"confidence\":0.0,\"source_evidence\":\"证据\",\"reason\":\"原因\"}]}。",
        ],
        "rows": rows,
    }


def _candidate_rows(document: dict[str, Any], max_rows: int) -> list[dict[str, Any]]:
    candidates = []
    for index, row in enumerate(document.get("mapped_detail_rows") or []):
        issues = _related_issues(document, row)
        if not _row_needs_repair(row, issues):
            continue
        standard = row.get("standard") or {}
        suspect_fields = []
        if _amount_mismatch(standard):
            suspect_fields.append(AMOUNT)
        candidates.append(
            {
                "row_key": index,
                "row": row,
                "missing_fields": _missing_fields(standard),
                "suspect_fields": suspect_fields,
                "issues": issues,
            }
        )
        if len(candidates) >= max_rows:
            break
    return candidates


def _normalize_field_value(field: str, value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if field in NUMERIC_FIELDS:
        return normalize_number(text)
    if field == DATE:
        return normalize_date(text) or text
    return text


def _can_apply_field(field: str, current: Any, candidate: Any, suspect_fields: list[str]) -> bool:
    if field not in ALLOWED_FIELDS:
        return False
    candidate_text = clean_text(candidate)
    if not candidate_text:
        return False
    if field in NUMERIC_FIELDS and not normalize_number(candidate_text):
        return False
    if field == UNIT and normalize_number(candidate_text) == candidate_text:
        return False
    if not clean_text(current):
        return True
    return field in suspect_fields


def _append_ai_issue(document: dict[str, Any], row: dict[str, Any], field: str, raw_value: str, clean_value: str, message: str) -> None:
    document.setdefault("issues", []).append(
        {
            "page_index": row.get("page_index", ""),
            "region": "明细表",
            "field": field,
            "raw_value": raw_value[:500],
            "clean_value": clean_value,
            "confidence": "AI",
            "message": message,
        }
    )


def _apply_repairs(document: dict[str, Any], response: dict[str, Any]) -> int:
    rows = document.get("mapped_detail_rows") or []
    applied = 0
    for repair in response.get("repairs") or []:
        try:
            row_key = int(repair.get("row_key"))
        except (TypeError, ValueError):
            continue
        if row_key < 0 or row_key >= len(rows):
            continue
        confidence = repair.get("confidence", 0)
        try:
            confidence_value = float(confidence)
        except (TypeError, ValueError):
            confidence_value = 0
        if confidence_value < 0.65:
            continue

        row = rows[row_key]
        standard = row.setdefault("standard", {})
        suspect_fields = []
        if _amount_mismatch(standard):
            suspect_fields.append(AMOUNT)
        set_fields = repair.get("set_fields") or {}
        source_evidence = clean_text(repair.get("source_evidence"))
        reason = clean_text(repair.get("reason"))
        for field, raw_value in set_fields.items():
            current = standard.get(field, "")
            if not _can_apply_field(field, current, raw_value, suspect_fields):
                continue
            clean_value = _normalize_field_value(field, raw_value)
            if not clean_value:
                continue
            standard[field] = clean_value
            row.setdefault("ai_repair", []).append(
                {
                    "field": field,
                    "value": clean_value,
                    "source_evidence": source_evidence,
                    "reason": reason,
                    "confidence": confidence_value,
                }
            )
            _append_ai_issue(
                document,
                row,
                field,
                source_evidence or clean_text(row.get("raw_text")),
                clean_value,
                f"AI补缺已写入：{reason or '根据订单正文证据补足字段'}",
            )
            applied += 1
    return applied


def audit_and_repair_purchase_document(
    document: dict[str, Any],
    *,
    log: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    repaired = copy.deepcopy(document)
    config = get_ai_repair_config()
    if not config.available:
        return repaired

    candidates = _candidate_rows(repaired, config.max_rows)
    if not candidates:
        return repaired

    if log:
        log(f"AI补缺检查：发现 {len(candidates)} 行明细需要补缺/复核，开始调用 {config.model}。")
    payload = _build_repair_payload(repaired, candidates)
    try:
        response = request_repair_json(config, payload)
    except DeepSeekRepairError as exc:
        repaired.setdefault("issues", []).append(
            {
                "page_index": "",
                "region": "AI补缺",
                "field": "",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": f"AI补缺跳过：{exc}",
            }
        )
        if log:
            log(f"AI补缺跳过：{exc}")
        return repaired

    applied = _apply_repairs(repaired, response)
    repaired.setdefault("ai_repair_summary", {})["candidate_rows"] = len(candidates)
    repaired.setdefault("ai_repair_summary", {})["applied_fields"] = applied
    if log:
        log(f"AI补缺完成：候选行 {len(candidates)}，写入字段 {applied}。")
    return repaired
