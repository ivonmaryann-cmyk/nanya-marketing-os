from __future__ import annotations

import copy
from decimal import Decimal
from typing import Any, Callable

from .ai_repair_config import AiRepairConfig, get_ai_repair_config
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
FIELD_ALIASES = {
    SEQ: SEQ,
    CODE: CODE,
    NAME: NAME,
    DESC: DESC,
    QTY: QTY,
    UNIT: UNIT,
    PRICE: PRICE,
    AMOUNT: AMOUNT,
    DATE: DATE,
    REMARK: REMARK,
    "\u5e8f\u53f7": SEQ,
    "\u7269\u6599\u7f16\u7801": CODE,
    "\u6599\u4ef6\u7f16\u53f7": CODE,
    "\u4ea7\u54c1\u7f16\u7801": CODE,
    "\u7269\u6599\u540d\u79f0": NAME,
    "\u4ea7\u54c1\u540d\u79f0": NAME,
    "\u540d\u79f0\u89c4\u683c": NAME,
    "\u89c4\u683c\u578b\u53f7": DESC,
    "\u8bf4\u660e": DESC,
    "\u6570\u91cf": QTY,
    "\u8ba2\u8d2d\u6570\u91cf": QTY,
    "\u5355\u4f4d": UNIT,
    "\u542b\u7a0e\u5355\u4ef7": PRICE,
    "\u5355\u4ef7": PRICE,
    "\u91d1\u989d": AMOUNT,
    "\u542b\u7a0e\u91d1\u989d": AMOUNT,
    "\u4ea4\u8d27\u65e5\u671f": DATE,
    "\u4ea4\u671f": DATE,
    "\u5230\u8d27\u65e5\u671f": DATE,
    "\u5907\u6ce8": REMARK,
}


def _money(value: Any) -> Decimal | None:
    return decimal_or_none(value)


def _amount_mismatch(standard: dict[str, Any]) -> bool:
    qty = _money(standard.get(QTY))
    price = _money(standard.get(PRICE))
    amount = _money(standard.get(AMOUNT))
    return qty is not None and price is not None and amount is not None and abs(qty * price - amount) > Decimal("0.05")


def _missing_fields(standard: dict[str, Any]) -> list[str]:
    return [field for field in CRITICAL_FIELDS if not clean_text(standard.get(field))]


def _canonical_field(field: Any) -> str:
    text = clean_text(field)
    return FIELD_ALIASES.get(text, text)


def _row_needs_repair(row: dict[str, Any], related_issues: list[dict[str, Any]]) -> bool:
    standard = row.get("standard") or {}
    missing = _missing_fields(standard)
    if QTY in missing or UNIT in missing:
        return True
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


def _build_repair_payload(
    document: dict[str, Any],
    candidates: list[dict[str, Any]],
    config: AiRepairConfig,
) -> dict[str, Any]:
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
        "business_instruction": getattr(config, "repair_instruction", ""),
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


def _raw_table_rows_for_body_repair(document: dict[str, Any]) -> list[dict[str, Any]]:
    tables = []
    for table in document.get("raw_detail_tables") or []:
        rows = table.get("rows") or table.get("raw_rows") or []
        if not rows:
            continue
        tables.append(
            {
                "page_index": table.get("page_index", ""),
                "table_index": table.get("table_index", ""),
                "method": table.get("method", ""),
                "bbox": table.get("bbox") or [],
                "rows": rows[:30],
            }
        )
    return tables[:4]


def _page_lines_for_body_repair(document: dict[str, Any]) -> list[dict[str, Any]]:
    pages = []
    for page in document.get("pages") or []:
        lines = [clean_text(line) for line in page.get("text_lines") or [] if clean_text(line)]
        if not lines:
            continue
        pages.append({"page_index": page.get("page_index", ""), "text_lines": lines[:80]})
    return pages[:3]


def _body_missing_payload(
    document: dict[str, Any],
    max_rows: int,
    config: AiRepairConfig,
) -> dict[str, Any] | None:
    raw_tables = _raw_table_rows_for_body_repair(document)
    page_lines = _page_lines_for_body_repair(document)
    if not raw_tables and not page_lines:
        return None
    return {
        "task": "rebuild_missing_purchase_order_body_only",
        "source_file": document.get("source_file", ""),
        "allowed_fields": list(STANDARD_HEADERS),
        "field_aliases": {
            "\u5e8f\u53f7": SEQ,
            "\u7269\u6599\u7f16\u7801": CODE,
            "\u7269\u6599\u540d\u79f0": NAME,
            "\u8bf4\u660e": DESC,
            "\u6570\u91cf": QTY,
            "\u5355\u4f4d": UNIT,
            "\u542b\u7a0e\u5355\u4ef7": PRICE,
            "\u91d1\u989d": AMOUNT,
            "\u4ea4\u8d27\u65e5\u671f": DATE,
            "\u5907\u6ce8": REMARK,
        },
        "rules": [
            "\u53ea\u91cd\u5efa\u8ba2\u5355\u6b63\u6587\u660e\u7ec6\u884c\uff0c\u4e0d\u5904\u7406\u8ba2\u5355\u5934\u3001\u4ed8\u6b3e\u3001\u6761\u6b3e\u3001\u7b7e\u6838\u533a\u3002",
            "\u53ea\u80fd\u4f7f\u7528 raw_detail_tables \u6216 page_text_lines \u4e2d\u770b\u5f97\u5230\u7684\u8bc1\u636e\uff0c\u4e0d\u8981\u51ed\u7a7a\u7f16\u9020\u660e\u7ec6\u3002",
            "\u5982\u679c\u65e0\u6cd5\u786e\u8ba4\u660e\u7ec6\u884c\uff0c\u8fd4\u56de {\"rows\":[]} \u3002",
            "\u6700\u591a\u8fd4\u56de max_rows \u6761\u660e\u7ec6\uff0c\u6bcf\u6761\u9700\u8981 confidence>=0.65 \u548c source_evidence\u3002",
            "\u8fd4\u56de\u4e25\u683c JSON\uff1a{\"rows\":[{\"standard\":{\"\u5b57\u6bb5\":\"\u503c\"},\"original\":{\"raw\":\"\u539f\u6587\"},\"confidence\":0.0,\"source_evidence\":\"\u8bc1\u636e\",\"reason\":\"\u539f\u56e0\"}]}",
        ],
        "business_instruction": getattr(config, "rebuild_instruction", ""),
        "max_rows": max_rows,
        "raw_detail_tables": raw_tables,
        "page_text_lines": page_lines,
        "issues": (document.get("issues") or [])[:10],
    }


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
    field = _canonical_field(field)
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


def _apply_repairs(document: dict[str, Any], response: dict[str, Any]) -> dict[str, int]:
    rows = document.get("mapped_detail_rows") or []
    stats = {
        "returned_repairs": 0,
        "applied_fields": 0,
        "rejected_row_key": 0,
        "rejected_low_confidence": 0,
        "rejected_fields": 0,
    }
    repairs = response.get("repairs") or []
    if not isinstance(repairs, list):
        return stats
    for repair in repairs:
        stats["returned_repairs"] += 1
        try:
            row_key = int(repair.get("row_key"))
        except (TypeError, ValueError):
            stats["rejected_row_key"] += 1
            continue
        if row_key < 0 or row_key >= len(rows):
            stats["rejected_row_key"] += 1
            continue
        confidence = repair.get("confidence", 0)
        try:
            confidence_value = float(confidence)
        except (TypeError, ValueError):
            confidence_value = 0
        if confidence_value < 0.65:
            stats["rejected_low_confidence"] += 1
            continue

        row = rows[row_key]
        standard = row.setdefault("standard", {})
        suspect_fields = []
        if _amount_mismatch(standard):
            suspect_fields.append(AMOUNT)
        set_fields = repair.get("set_fields") or {}
        source_evidence = clean_text(repair.get("source_evidence"))
        reason = clean_text(repair.get("reason"))
        for raw_field, raw_value in set_fields.items():
            field = _canonical_field(raw_field)
            current = standard.get(field, "")
            if not _can_apply_field(field, current, raw_value, suspect_fields):
                stats["rejected_fields"] += 1
                continue
            clean_value = _normalize_field_value(field, raw_value)
            if not clean_value:
                stats["rejected_fields"] += 1
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
            stats["applied_fields"] += 1
    return stats


def _unresolved_required_fields(standard: dict[str, Any]) -> list[str]:
    fields: list[str] = []
    if not clean_text(standard.get(CODE)) and not clean_text(standard.get(NAME)):
        fields.append("物料编码/物料名称")
    for field in [QTY, UNIT]:
        if not clean_text(standard.get(field)):
            fields.append(field)
    if not clean_text(standard.get(PRICE)) and not clean_text(standard.get(AMOUNT)):
        fields.append("含税单价/金额")
    return fields


def _append_unresolved_issues(document: dict[str, Any]) -> int:
    issues = document.setdefault("issues", [])
    existing = {
        (issue.get("page_index"), issue.get("field"), issue.get("message"), clean_text(issue.get("raw_value")))
        for issue in issues
    }
    count = 0
    for row in document.get("mapped_detail_rows") or []:
        missing = _unresolved_required_fields(row.get("standard") or {})
        if not missing:
            continue
        field_text = "、".join(missing)
        message = f"本地识别与 AI 补缺后仍缺少关键字段：{field_text}，需人工复核。"
        raw_text = clean_text(row.get("raw_text"))[:500]
        key = (row.get("page_index", ""), field_text, message, raw_text)
        if key in existing:
            continue
        issues.append(
            {
                "page_index": row.get("page_index", ""),
                "region": "明细表",
                "field": field_text,
                "raw_value": raw_text,
                "clean_value": "",
                "confidence": row.get("confidence", ""),
                "message": message,
            }
        )
        existing.add(key)
        count += 1
    return count


def _row_has_minimum_body_fields(standard: dict[str, Any]) -> bool:
    has_name_or_code = bool(clean_text(standard.get(CODE)) or clean_text(standard.get(NAME)))
    has_qty = bool(clean_text(standard.get(QTY)))
    has_price_or_amount = bool(clean_text(standard.get(PRICE)) or clean_text(standard.get(AMOUNT)))
    return has_name_or_code and has_qty and has_price_or_amount


def _apply_missing_body_rows(document: dict[str, Any], response: dict[str, Any]) -> int:
    rows = []
    for index, item in enumerate(response.get("rows") or []):
        try:
            confidence = float(item.get("confidence", 0))
        except (TypeError, ValueError):
            confidence = 0
        if confidence < 0.65:
            continue

        standard: dict[str, str] = {}
        for raw_field, raw_value in (item.get("standard") or {}).items():
            field = _canonical_field(raw_field)
            if field not in ALLOWED_FIELDS:
                continue
            clean_value = _normalize_field_value(field, raw_value)
            if clean_value:
                standard[field] = clean_value
        if not _row_has_minimum_body_fields(standard):
            continue

        original = item.get("original") if isinstance(item.get("original"), dict) else {}
        source_evidence = clean_text(item.get("source_evidence"))
        reason = clean_text(item.get("reason"))
        raw_text = source_evidence or clean_text(original.get("raw"))
        row = {
            "original": original,
            "standard": standard,
            "cleaning_notes": [],
            "page_index": "",
            "table_index": "",
            "row_index": index,
            "raw_text": raw_text,
            "confidence": confidence,
            "method": "ai_missing_body_rebuild",
            "ai_repair": [
                {
                    "field": "\u660e\u7ec6\u884c",
                    "value": raw_text,
                    "source_evidence": source_evidence,
                    "reason": reason,
                    "confidence": confidence,
                }
            ],
        }
        rows.append(row)
        document.setdefault("issues", []).append(
            {
                "page_index": "",
                "region": "AI补缺",
                "field": "订单正文",
                "raw_value": raw_text[:500],
                "clean_value": " ".join(f"{field}:{value}" for field, value in standard.items())[:500],
                "confidence": confidence,
                "message": f"AI在原始表格/文本证据中重建缺失明细：{reason}",
            }
        )
    if not rows:
        return 0
    document["mapped_detail_rows"] = rows
    return len(rows)


def _try_repair_missing_body(
    document: dict[str, Any],
    config: Any,
    *,
    log: Callable[[str], None] | None = None,
) -> bool:
    if document.get("mapped_detail_rows"):
        return False
    payload = _body_missing_payload(document, config.max_rows, config)
    if not payload:
        return False
    if log:
        log(f"AI\u8865\u7f3a\u68c0\u67e5\uff1a\u8ba2\u5355\u6b63\u6587\u660e\u7ec6\u7f3a\u5931\uff0c\u5c1d\u8bd5\u8c03\u7528 {config.model} \u91cd\u5efa\u3002")
    try:
        response = request_repair_json(config, payload)
    except DeepSeekRepairError as exc:
        document.setdefault("issues", []).append(
            {
                "page_index": "",
                "region": "AI补缺",
                "field": "订单正文",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": f"AI正文重建跳过：{exc}",
            }
        )
        if log:
            log(f"AI\u6b63\u6587\u91cd\u5efa\u8df3\u8fc7\uff1a{exc}")
        return False
    applied_rows = _apply_missing_body_rows(document, response)
    document.setdefault("ai_repair_summary", {})["missing_body_rows"] = applied_rows
    if log:
        log(f"AI\u6b63\u6587\u91cd\u5efa\u5b8c\u6210\uff1a\u5199\u5165\u660e\u7ec6 {applied_rows} \u884c\u3002")
    return applied_rows > 0


def audit_and_repair_purchase_document(
    document: dict[str, Any],
    *,
    config: AiRepairConfig | None = None,
    log: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    repaired = copy.deepcopy(document)
    config = config or get_ai_repair_config()
    summary = repaired.setdefault("ai_repair_summary", {})
    summary.update(
        {
            "available": bool(config.available),
            "model": config.model if config.available else "",
            "config_version": getattr(config, "version_id", None),
            "config_fingerprint": getattr(config, "fingerprint", ""),
            "prompt_digest": getattr(config, "prompt_digest", ""),
            "candidate_rows": 0,
            "returned_repairs": 0,
            "applied_fields": 0,
            "rejected_row_key": 0,
            "rejected_low_confidence": 0,
            "rejected_fields": 0,
        }
    )
    if not config.available:
        summary["unresolved_rows"] = _append_unresolved_issues(repaired)
        return repaired

    if _try_repair_missing_body(repaired, config, log=log):
        summary["unresolved_rows"] = _append_unresolved_issues(repaired)
        return repaired

    candidates = _candidate_rows(repaired, config.max_rows)
    summary["candidate_rows"] = len(candidates)
    if not candidates:
        summary["unresolved_rows"] = _append_unresolved_issues(repaired)
        return repaired

    if log:
        log(f"AI补缺检查：发现 {len(candidates)} 行明细需要补缺/复核，开始调用 {config.model}。")
    payload = _build_repair_payload(repaired, candidates, config)
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
        summary["request_error"] = str(exc)
        summary["unresolved_rows"] = _append_unresolved_issues(repaired)
        return repaired

    stats = _apply_repairs(repaired, response)
    summary.update(stats)
    summary["unresolved_rows"] = _append_unresolved_issues(repaired)
    if log:
        log(
            f"AI补缺完成：候选行 {len(candidates)}，模型返回 {stats['returned_repairs']} 条，"
            f"写入字段 {stats['applied_fields']}，剩余需复核 {summary['unresolved_rows']} 行。"
        )
    return repaired
