from __future__ import annotations

import re
import unicodedata
from collections import Counter
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable

from .ai_repair_config import AiRepairConfig, get_ai_repair_config
from .deepseek_repair_client import request_repair_json
from .purchase_field_rules import clean_text, decimal_or_none, normalize_date, normalize_number


FACTORY_MAIN_HEADERS = [
    "单别（必填）",
    "类型1（必填）",
    "类型2（必填）",
    "账款客户编号（必填）",
    "送货客户编号（必填）",
    "到期付款日（必填）",
    "送货厂别（选填）",
    "客户订单号（选填）",
    "",
    "",
    "",
    "",
]

FACTORY_DETAIL_HEADERS = [
    "项次（必填）",
    "合约项次/估价项次/报价项次",
    "产品编号",
    "客户产品编号（必填）",
    "出货日期（选填）",
    "数量（必填）",
    "税前单价（选填）",
    "单价（选填）",
    "产地（根据账套默认产地）",
    "客户订单序号（选填）",
    "客户订单号",
    "备注（选填）",
]

ROLE_ITEM = "item"
ROLE_CONTRACT_ITEM = "contract_item"
ROLE_CUSTOMER_PRODUCT = "customer_product_code"
ROLE_DATE = "shipping_date"
ROLE_QUANTITY = "quantity"
ROLE_PRE_TAX_PRICE = "pre_tax_price"
ROLE_TAX_PRICE = "tax_price"
ROLE_REMARK = "remark"

AI_TARGET_ROLES = {
    ROLE_ITEM,
    ROLE_CONTRACT_ITEM,
    ROLE_CUSTOMER_PRODUCT,
    ROLE_DATE,
    ROLE_QUANTITY,
    ROLE_PRE_TAX_PRICE,
    ROLE_TAX_PRICE,
    ROLE_REMARK,
}

ROLE_ALIASES = {
    ROLE_ITEM: ["序号", "项次", "项目", "行号", "item", "itemno", "no"],
    ROLE_CONTRACT_ITEM: ["合约项次", "合同项次", "估价项次", "报价项次"],
    ROLE_CUSTOMER_PRODUCT: [
        "客户产品编号",
        "客户物料编号",
        "物料编号",
        "物料编码",
        "物料代码",
        "原料编码",
        "料件编号",
        "料号",
        "品号",
        "产品编号",
        "goodsno",
        "partno",
        "partnumber",
        "customerpartno",
        "pn",
    ],
    ROLE_DATE: ["出货日期", "发货日期", "交货日期", "到货日期", "交期", "deliverydate", "arrivaldate"],
    ROLE_QUANTITY: ["数量", "采购量", "订购数量", "采购数量", "quantity", "qty"],
    ROLE_REMARK: ["备注", "采购备注", "附注", "notes", "remark", "comments", "comment"],
}

TEMPLATE_ROLE_ALIASES = {
    "jingwang_purchase_order": {
        ROLE_ITEM: ["itemno"],
        ROLE_CUSTOMER_PRODUCT: ["partno"],
        ROLE_DATE: ["deliverydate"],
        ROLE_QUANTITY: ["quantity"],
    },
    "talian_purchase_order": {
        ROLE_ITEM: ["no"],
        ROLE_CUSTOMER_PRODUCT: ["原料编码"],
        ROLE_DATE: ["交期"],
        ROLE_QUANTITY: ["数量"],
    },
    "ganzhou_chaoyue_purchase_order": {
        ROLE_ITEM: ["项次"],
        ROLE_CUSTOMER_PRODUCT: ["料件编号"],
        ROLE_DATE: ["交货日期"],
        ROLE_QUANTITY: ["数量"],
    },
    "ganzhou_yihao_purchase_order": {
        ROLE_ITEM: ["项目"],
        ROLE_CUSTOMER_PRODUCT: ["物料编号", "物料编码"],
        ROLE_DATE: ["到货日期"],
        ROLE_QUANTITY: ["数量"],
    },
}

ORDER_CONTRACT_FALLBACK_TEMPLATES = {"talian_purchase_order"}


def _header_key(value: Any) -> str:
    text = unicodedata.normalize("NFKC", clean_text(value)).lower()
    text = re.sub(r"\([^)]*\)|（[^）]*）", "", text)
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", text)


def _decimal_text(value: Decimal) -> str:
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _roll_length_sources(detail: dict[str, Any]) -> list[str]:
    standard = detail.get("standard") or {}
    candidates = [clean_text(standard.get("说明")), clean_text(standard.get("物料名称"))]
    for header, value in (detail.get("original") or {}).items():
        compact_header = _header_key(header)
        if any(keyword in compact_header for keyword in ["规格", "型号", "描述", "description", "spec"]):
            candidates.append(clean_text(value))
    sources: list[str] = []
    seen: set[str] = set()
    for text in candidates:
        if text and text not in seen:
            seen.add(text)
            sources.append(text)
    return sources


def _is_model_meter_token(text: str, start: int) -> bool:
    prefix = text[:start].rstrip()
    return bool(re.search(r"(?:NY|PP)\s*$", prefix, flags=re.I))


def _extract_roll_length(detail: dict[str, Any]) -> tuple[Decimal | None, str, str]:
    sources = _roll_length_sources(detail)
    if not sources:
        return None, "", "规格为空，无法提取每卷米数"

    values: set[Decimal] = set()
    evidence: list[str] = []
    patterns = [
        re.compile(r"(?<![A-Za-z])(?P<value>\d+(?:\.\d+)?)\s*(?:米|[mM])\s*/\s*卷", flags=re.I),
        re.compile(r"[*×xX]\s*(?P<value>\d+(?:\.\d+)?)\s*(?:米|[mM])(?=$|[\s/，,;；)）])", flags=re.I),
        re.compile(r"(?<![A-Za-z0-9])(?P<value>\d+(?:\.\d+)?)\s*(?:米|[mM])(?=$|[\s/，,;；)）])", flags=re.I),
    ]
    for text in sources:
        for pattern_index, pattern in enumerate(patterns):
            for match in pattern.finditer(text):
                if pattern_index == 2 and _is_model_meter_token(text, match.start()):
                    continue
                value = Decimal(match.group("value"))
                if value <= 0 or value > Decimal("10000"):
                    continue
                values.add(value)
                evidence.append(match.group(0).strip())
    if not values:
        return None, "\n".join(sources), "未找到可靠的米制卷长"
    if len(values) > 1:
        found = "、".join(_decimal_text(value) for value in sorted(values))
        return None, "\n".join(sources), f"识别到多个不同卷长（{found}米），无法唯一确认"
    length = next(iter(values))
    source_evidence = next((item for item in evidence if item), _decimal_text(length))
    return length, source_evidence, ""


def _project_detail_with_meta(
    detail: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any] | None, Decimal | None, Decimal | None]:
    standard = dict(detail.get("standard") or {})
    if clean_text(standard.get("单位")) != "卷":
        return standard, None, None, None

    quantity = decimal_or_none(standard.get("数量"))
    if quantity is None or quantity <= 0:
        reason = "卷数为空、非数字或不大于零"
        length = None
        evidence = "\n".join(_roll_length_sources(detail))
    else:
        length, evidence, reason = _extract_roll_length(detail)
    if length is None:
        code = clean_text(standard.get("物料编码"))
        raw_value = f"物料编码：{code}\n规格：{evidence}" if code else f"规格：{evidence}"
        issue = {
            "page_index": detail.get("page_index", 0),
            "region": "明细数据",
            "field": "卷长",
            "raw_value": raw_value.strip(),
            "clean_value": "",
            "confidence": "规则",
            "message": f"单位为卷，但{reason}；厂内明细保留卷制口径（数量和价格）。",
        }
        return standard, issue, None, None

    standard["数量"] = _decimal_text(quantity * length)
    standard["单位"] = "米"
    price = decimal_or_none(standard.get("含税单价"))
    if price is not None:
        standard["含税单价"] = _decimal_text(price / length)
    return standard, None, length, quantity


def _append_original_roll_remark(remark: Any, roll_quantity: Decimal | None) -> str:
    current = clean_text(remark)
    if roll_quantity is None:
        return current

    roll_text = f"{_decimal_text(roll_quantity)}卷"
    existing_parts = {
        part.strip() for part in re.split(r"[；;\n]+", current) if part.strip()
    }
    if roll_text in existing_parts:
        return current
    return f"{current}；{roll_text}" if current else roll_text


def _factory_remark(remark: Any, roll_quantity: Decimal | None) -> str:
    current = clean_text(remark)
    while current.endswith("&"):
        current = current[:-1].rstrip()
    combined = _append_original_roll_remark(current, roll_quantity)
    return f"{combined}&" if combined else ""


def _project_detail_standard_fields(detail: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any] | None]:
    standard, issue, _length, _roll_quantity = _project_detail_with_meta(detail)
    return standard, issue


def _append_issue_once(document: dict[str, Any], issue: dict[str, Any]) -> None:
    issues = document.setdefault("issues", [])
    key_fields = ["page_index", "region", "field", "raw_value", "message"]
    key = tuple(issue.get(field) for field in key_fields)
    if any(tuple(existing.get(field) for field in key_fields) == key for existing in issues):
        return
    issues.append(issue)


def _aliases_for_role(role: str, template_id: str) -> list[str]:
    aliases = list((TEMPLATE_ROLE_ALIASES.get(template_id) or {}).get(role) or [])
    aliases.extend(ROLE_ALIASES.get(role) or [])
    return list(dict.fromkeys(_header_key(alias) for alias in aliases if _header_key(alias)))


def _header_match_score(header: str, alias: str) -> int:
    if not header or not alias:
        return 0
    if header == alias:
        return 100
    if len(alias) >= 2 and alias in header:
        return 80
    return 0


def _header_role(header: Any, template_id: str) -> tuple[str, int] | None:
    key = _header_key(header)
    if not key:
        return None
    if any(token in key for token in ["物料名称", "产品名称", "品名", "规格", "型号", "description", "desc"]):
        return None
    scores: list[tuple[int, str]] = []
    for role in [ROLE_ITEM, ROLE_CONTRACT_ITEM, ROLE_CUSTOMER_PRODUCT, ROLE_DATE, ROLE_QUANTITY, ROLE_REMARK]:
        score = max((_header_match_score(key, alias) for alias in _aliases_for_role(role, template_id)), default=0)
        if score:
            scores.append((score, role))
    if not scores:
        return None
    scores.sort(reverse=True)
    best_score, best_role = scores[0]
    if len(scores) > 1 and scores[1][0] == best_score and scores[1][1] != best_role:
        return None
    return best_role, best_score


def _price_role(header: Any) -> tuple[str, str] | None:
    key = _header_key(header)
    if not key or any(token in key for token in ["金额", "amount", "total"]):
        return None
    if not any(token in key for token in ["单价", "unitprice", "price"]):
        return None
    if any(token in key for token in ["税前", "未税", "不含税", "excltax", "withouttax"]):
        return ROLE_PRE_TAX_PRICE, "explicit_pre_tax"
    if "含税" in key or "incltax" in key:
        return ROLE_TAX_PRICE, "explicit_tax_included"
    return ROLE_TAX_PRICE, "generic_default_tax_included"


def _value_valid(role: str, value: Any) -> bool:
    text = clean_text(value)
    if not text:
        return False
    if role in {ROLE_QUANTITY, ROLE_PRE_TAX_PRICE, ROLE_TAX_PRICE}:
        number = decimal_or_none(text)
        return number is not None and number >= 0
    if role == ROLE_DATE:
        return bool(normalize_date(text))
    if role in {ROLE_ITEM, ROLE_CONTRACT_ITEM}:
        return bool(re.fullmatch(r"[A-Za-z0-9._/-]{1,30}", text))
    if role == ROLE_CUSTOMER_PRODUCT:
        return len(text) <= 100 and bool(re.search(r"[A-Za-z0-9]", text))
    return True


def _original_candidates(detail: dict[str, Any], role: str, template_id: str) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for header, value in (detail.get("original") or {}).items():
        match = _header_role(header, template_id)
        if not match or match[0] != role or not _value_valid(role, value):
            continue
        candidates.append(
            {
                "header": clean_text(header),
                "value": clean_text(value),
                "method": (
                    "template_alias"
                    if _header_key(header)
                    in {_header_key(alias) for alias in (TEMPLATE_ROLE_ALIASES.get(template_id) or {}).get(role, [])}
                    else "standard_alias"
                ),
                "confidence": match[1] / 100,
            }
        )
    return candidates


def _unique_candidate(candidates: list[dict[str, Any]]) -> tuple[dict[str, Any] | None, str]:
    if not candidates:
        return None, ""
    values = {clean_text(candidate.get("value")) for candidate in candidates if clean_text(candidate.get("value"))}
    if len(values) > 1:
        return None, "同一目标字段命中多个不同来源值"
    candidates.sort(key=lambda item: float(item.get("confidence") or 0), reverse=True)
    return candidates[0], ""


def _explicit_prices(detail: dict[str, Any]) -> tuple[dict[str, Decimal], list[dict[str, Any]], bool]:
    candidates: dict[str, list[tuple[Decimal, str, str]]] = {ROLE_PRE_TAX_PRICE: [], ROLE_TAX_PRICE: []}
    for header, value in (detail.get("original") or {}).items():
        role_info = _price_role(header)
        number = decimal_or_none(value)
        if not role_info or number is None:
            continue
        role, method = role_info
        candidates[role].append((number, clean_text(header), method))

    prices: dict[str, Decimal] = {}
    evidence: list[dict[str, Any]] = []
    for role, entries in candidates.items():
        explicit_entries = [entry for entry in entries if entry[2] != "generic_default_tax_included"]
        selected_entries = explicit_entries or entries
        values = {entry[0] for entry in selected_entries}
        if len(values) != 1:
            if len(values) > 1:
                evidence.append({"role": role, "status": "rejected", "reason": "同一价格口径存在多个不同值"})
            continue
        if selected_entries:
            prices[role] = selected_entries[0][0]
            evidence.append(
                {
                    "role": role,
                    "status": "accepted",
                    "source_header": selected_entries[0][1],
                    "method": selected_entries[0][2],
                }
            )
    generic_defaulted = any(
        item.get("status") == "accepted" and item.get("method") == "generic_default_tax_included"
        for item in evidence
    )
    return prices, evidence, generic_defaulted


def _clean_order_identifier(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{2,60}", text):
        return text
    candidates = re.findall(r"(?<![A-Za-z0-9])([A-Za-z]{1,8}[-_]?\d{3,}(?:[-_]\d+)*)(?![A-Za-z0-9])", text)
    if not candidates:
        return ""
    candidates.sort(key=lambda item: ("-" not in item and "_" not in item, -len(item)))
    return candidates[0]


def _order_number(document: dict[str, Any]) -> tuple[str, str]:
    header_info = document.get("header_info") or {}
    order_number = _clean_order_identifier(header_info.get("订单号"))
    if order_number:
        return order_number, "订单号"
    if clean_text(document.get("template_id")) in ORDER_CONTRACT_FALLBACK_TEMPLATES:
        contract = _clean_order_identifier(header_info.get("合同编号"))
        if contract:
            return contract, "客户模板合同编号"
    source_stem = Path(clean_text(document.get("source_file"))).stem
    source_stem = re.sub(r"^\d{3}_", "", source_stem)
    filename_order = _clean_order_identifier(source_stem)
    if filename_order and filename_order == source_stem:
        return filename_order, "来源文件名"
    return "", ""


def _customer_identity(document: dict[str, Any]) -> str:
    header_info = document.get("header_info") or {}
    customer = clean_text(header_info.get("客户"))
    if customer:
        return customer
    supplier = clean_text(header_info.get("供应商"))
    candidates: list[str] = []
    for page in document.get("pages") or []:
        for line in page.get("text_lines") or []:
            text = clean_text(line)
            for match in re.findall(r"[A-Za-z\u4e00-\u9fff（）()]{4,}(?:股份有限公司|有限责任公司|有限公司)", text):
                name = clean_text(match)
                if name and name != supplier and "南亚新材料" not in name and name not in candidates:
                    candidates.append(name)
    if candidates:
        return candidates[0]
    return clean_text(document.get("template_label")).removesuffix("采购订单")


def _ai_header_candidates(document: dict[str, Any]) -> list[dict[str, Any]]:
    samples: dict[str, list[str]] = {}
    for detail in document.get("mapped_detail_rows") or []:
        for header, value in (detail.get("original") or {}).items():
            text = clean_text(value)
            if not clean_text(header) or not text:
                continue
            bucket = samples.setdefault(clean_text(header), [])
            if text not in bucket and len(bucket) < 3:
                bucket.append(text[:160])
    return [{"source_header": header, "samples": values} for header, values in samples.items()]


def _request_ai_header_mapping(
    document: dict[str, Any],
    *,
    config: AiRepairConfig | None = None,
    log: Callable[[str], None] | None = None,
) -> tuple[dict[str, str], dict[str, Any]]:
    config = config or get_ai_repair_config()
    summary = {"available": bool(config.available), "requested": False, "returned": 0, "accepted": 0, "rejected": 0}
    if not config.available:
        return {}, summary
    source_headers = _ai_header_candidates(document)
    if not source_headers:
        return {}, summary
    payload = {
        "task": "map_purchase_order_headers_to_factory_fields",
        "source_file": document.get("source_file", ""),
        "template_id": document.get("template_id", ""),
        "customer": _customer_identity(document),
        "target_fields": sorted(AI_TARGET_ROLES),
        "rules": [
            "只判断来源表头的业务语义，不生成或修改任何单元格值。",
            "客户采购单中的物料编号、料件编号、Part No. 应映射为 customer_product_code，不是厂内产品编号。",
            "物料名称、品名、规格、型号不得映射为 customer_product_code。",
            "只有明确写税前、未税或不含税时才能映射为 pre_tax_price；普通单价映射为 tax_price。",
            "返回严格 JSON：{\"mappings\":[{\"source_header\":\"...\",\"target_field\":\"...\",\"confidence\":0.0,\"reason\":\"...\"}]}。",
        ],
        "business_instruction": getattr(config, "header_mapping_instruction", ""),
        "source_headers": source_headers,
    }
    summary["requested"] = True
    try:
        response = request_repair_json(config, payload)
    except Exception as exc:
        summary["error"] = str(exc)
        if log:
            log(f"厂内表头 AI 映射跳过：{exc}")
        return {}, summary

    available_headers = {clean_text(item["source_header"]): item["samples"] for item in source_headers}
    accepted: dict[str, str] = {}
    claimed_targets: set[str] = set()
    claimed_sources: set[str] = set()
    mappings = response.get("mappings") or []
    if not isinstance(mappings, list):
        return {}, summary
    for mapping in mappings:
        summary["returned"] += 1
        source_header = clean_text(mapping.get("source_header"))
        target = clean_text(mapping.get("target_field"))
        try:
            confidence = float(mapping.get("confidence") or 0)
        except (TypeError, ValueError):
            confidence = 0
        samples = available_headers.get(source_header) or []
        if (
            confidence < 0.9
            or target not in AI_TARGET_ROLES
            or source_header not in available_headers
            or target in claimed_targets
            or source_header in claimed_sources
            or not any(_value_valid(target, sample) for sample in samples)
        ):
            summary["rejected"] += 1
            continue
        accepted[source_header] = target
        claimed_targets.add(target)
        claimed_sources.add(source_header)
        summary["accepted"] += 1
    return accepted, summary


def _ai_candidate(detail: dict[str, Any], role: str, ai_mapping: dict[str, str]) -> dict[str, Any] | None:
    matches = []
    for header, value in (detail.get("original") or {}).items():
        if ai_mapping.get(clean_text(header)) != role or not _value_valid(role, value):
            continue
        matches.append({"header": clean_text(header), "value": clean_text(value), "method": "ai_header_mapping", "confidence": 0.9})
    candidate, _reason = _unique_candidate(matches)
    return candidate


def _resolve_value(
    detail: dict[str, Any],
    role: str,
    template_id: str,
    ai_mapping: dict[str, str],
) -> tuple[str, dict[str, Any] | None, str]:
    candidate, reason = _unique_candidate(_original_candidates(detail, role, template_id))
    if candidate:
        return clean_text(candidate["value"]), candidate, ""
    if reason:
        return "", None, reason
    ai_candidate = _ai_candidate(detail, role, ai_mapping)
    if ai_candidate:
        return clean_text(ai_candidate["value"]), ai_candidate, ""
    return "", None, ""


def _typed_identifier(value: Any) -> str:
    text = clean_text(value)
    number = decimal_or_none(text)
    if number is not None and number == number.to_integral_value():
        return str(int(number))
    return text


def project_factory_document(
    document: dict[str, Any],
    *,
    config: AiRepairConfig | None = None,
    log: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    config = config or get_ai_repair_config()
    template_id = clean_text(document.get("template_id"))
    order_number, order_source = _order_number(document)
    details = document.get("mapped_detail_rows") or []

    needs_ai = any(
        not clean_text((detail.get("standard") or {}).get("物料编码"))
        or not clean_text((detail.get("standard") or {}).get("数量"))
        for detail in details
    )
    ai_mapping: dict[str, str] = {}
    ai_summary: dict[str, Any] = {"available": False, "requested": False, "returned": 0, "accepted": 0, "rejected": 0}
    if needs_ai:
        ai_mapping, ai_summary = _request_ai_header_mapping(document, config=config, log=log)
    ai_summary.update(
        {
            "config_version": getattr(config, "version_id", None),
            "config_fingerprint": getattr(config, "fingerprint", ""),
            "prompt_digest": getattr(config, "prompt_digest", ""),
        }
    )

    factory_rows: list[dict[str, Any]] = []
    evidence_rows: list[dict[str, Any]] = []
    method_counts: Counter[str] = Counter()
    ready_rows = 0
    review_rows = 0
    generic_price_logged = False

    for index, detail in enumerate(details, start=1):
        source_standard = detail.get("standard") or {}
        projected, projection_issue, roll_length, original_roll_quantity = _project_detail_with_meta(detail)
        if projection_issue:
            _append_issue_once(document, projection_issue)

        row_evidence: list[dict[str, Any]] = []
        source_item = clean_text(source_standard.get("序号"))
        if source_item:
            item = _typed_identifier(source_item)
            row_evidence.append({"field": "项次", "method": "standard", "value": item})
            method_counts["standard"] += 1
        else:
            item_value, item_evidence, item_reason = _resolve_value(detail, ROLE_ITEM, template_id, ai_mapping)
            if item_value:
                item = _typed_identifier(item_value)
                row_evidence.append({"field": "项次", **(item_evidence or {}), "value": item})
                method_counts[(item_evidence or {}).get("method", "alias")] += 1
            else:
                item = str(index)
                method_counts["generated"] += 1
                _append_issue_once(
                    document,
                    {
                        "page_index": detail.get("page_index", 0),
                        "region": "厂内模板",
                        "field": "项次",
                        "raw_value": clean_text(detail.get("raw_text")),
                        "clean_value": item,
                        "confidence": "顺序生成",
                        "message": f"客户明细未识别到项次，厂内必填项次按输出顺序生成 {item}。{item_reason}",
                    },
                )

        contract_item, contract_evidence, contract_reason = _resolve_value(
            detail, ROLE_CONTRACT_ITEM, template_id, ai_mapping
        )
        if contract_evidence:
            row_evidence.append({"field": "合约项次/估价项次/报价项次", **contract_evidence})
            method_counts[contract_evidence.get("method", "alias")] += 1
        if contract_reason:
            _append_issue_once(
                document,
                {
                    "page_index": detail.get("page_index", 0),
                    "region": "厂内模板",
                    "field": "合约项次/估价项次/报价项次",
                    "raw_value": clean_text(detail.get("raw_text")),
                    "clean_value": "",
                    "confidence": "冲突",
                    "message": contract_reason,
                },
            )

        customer_product = clean_text(source_standard.get("物料编码"))
        if customer_product and not _value_valid(ROLE_CUSTOMER_PRODUCT, customer_product):
            customer_product = ""
        if customer_product:
            row_evidence.append({"field": "客户产品编号", "method": "standard", "value": customer_product})
            method_counts["standard"] += 1
        else:
            customer_product, product_evidence, product_reason = _resolve_value(
                detail, ROLE_CUSTOMER_PRODUCT, template_id, ai_mapping
            )
            if product_evidence:
                row_evidence.append({"field": "客户产品编号", **product_evidence})
                method_counts[product_evidence.get("method", "alias")] += 1
            if product_reason:
                _append_issue_once(
                    document,
                    {
                        "page_index": detail.get("page_index", 0),
                        "region": "厂内模板",
                        "field": "客户产品编号",
                        "raw_value": clean_text(detail.get("raw_text")),
                        "clean_value": "",
                        "confidence": "冲突",
                        "message": product_reason,
                    },
                )

        date_value = normalize_date(projected.get("交货日期"))
        if not date_value:
            raw_date, date_evidence, date_reason = _resolve_value(detail, ROLE_DATE, template_id, ai_mapping)
            date_value = normalize_date(raw_date)
            if date_evidence and date_value:
                row_evidence.append({"field": "出货日期", **date_evidence, "value": date_value})
                method_counts[date_evidence.get("method", "alias")] += 1
            if date_reason:
                _append_issue_once(
                    document,
                    {
                        "page_index": detail.get("page_index", 0),
                        "region": "厂内模板",
                        "field": "出货日期",
                        "raw_value": clean_text(detail.get("raw_text")),
                        "clean_value": "",
                        "confidence": "冲突",
                        "message": date_reason,
                    },
                )

        quantity = normalize_number(projected.get("数量"))
        quantity_number = decimal_or_none(quantity)
        if quantity_number is None or quantity_number <= 0:
            quantity = ""
        if not quantity:
            raw_quantity, quantity_evidence, quantity_reason = _resolve_value(detail, ROLE_QUANTITY, template_id, ai_mapping)
            quantity = normalize_number(raw_quantity)
            if quantity_evidence and quantity:
                row_evidence.append({"field": "数量", **quantity_evidence, "value": quantity})
                method_counts[quantity_evidence.get("method", "alias")] += 1
            if quantity_reason:
                _append_issue_once(
                    document,
                    {
                        "page_index": detail.get("page_index", 0),
                        "region": "厂内模板",
                        "field": "数量",
                        "raw_value": clean_text(detail.get("raw_text")),
                        "clean_value": "",
                        "confidence": "冲突",
                        "message": quantity_reason,
                    },
                )

        prices, price_evidence, generic_defaulted = _explicit_prices(detail)
        tax_price = prices.get(ROLE_TAX_PRICE)
        pre_tax_price = prices.get(ROLE_PRE_TAX_PRICE)
        tax_price_conflict = any(
            item.get("role") == ROLE_TAX_PRICE and item.get("status") == "rejected"
            for item in price_evidence
        )
        # Some recognizers also project an explicitly untaxed source column to
        # the generic standard ``含税单价`` field. Once the original heading has
        # established an untaxed price, that derived value must not populate the
        # tax-inclusive column as well.
        has_explicit_pre_tax = any(
            item.get("role") == ROLE_PRE_TAX_PRICE
            and item.get("status") == "accepted"
            and item.get("method") == "explicit_pre_tax"
            for item in price_evidence
        )
        if tax_price is None and not tax_price_conflict and not has_explicit_pre_tax:
            tax_price = decimal_or_none(source_standard.get("含税单价"))
            if tax_price is not None:
                price_evidence.append({"role": ROLE_TAX_PRICE, "status": "accepted", "method": "standard_tax_price"})
        if roll_length is not None:
            if tax_price is not None:
                tax_price /= roll_length
            if pre_tax_price is not None:
                pre_tax_price /= roll_length
        if generic_defaulted and not generic_price_logged:
            generic_price_logged = True
            _append_issue_once(
                document,
                {
                    "page_index": detail.get("page_index", 0),
                    "region": "厂内模板",
                    "field": "单价",
                    "raw_value": "来源表头仅写单价",
                    "clean_value": "按含税单价写入",
                    "confidence": "默认口径",
                    "message": "来源价格表头未注明含税或未税，按已确认规则默认写入厂内“单价”列。",
                },
            )
        for rejected_price in [item for item in price_evidence if item.get("status") == "rejected"]:
            target_label = "税前单价" if rejected_price.get("role") == ROLE_PRE_TAX_PRICE else "单价"
            _append_issue_once(
                document,
                {
                    "page_index": detail.get("page_index", 0),
                    "region": "厂内模板",
                    "field": target_label,
                    "raw_value": clean_text(detail.get("raw_text")),
                    "clean_value": "",
                    "confidence": "冲突",
                    "message": f"{target_label}{rejected_price.get('reason')}，厂内价格保持为空。",
                },
            )
        row_evidence.extend(price_evidence)

        row = {
            FACTORY_DETAIL_HEADERS[0]: item,
            FACTORY_DETAIL_HEADERS[1]: contract_item,
            FACTORY_DETAIL_HEADERS[2]: "",
            FACTORY_DETAIL_HEADERS[3]: customer_product,
            FACTORY_DETAIL_HEADERS[4]: date_value,
            FACTORY_DETAIL_HEADERS[5]: quantity,
            FACTORY_DETAIL_HEADERS[6]: _decimal_text(pre_tax_price) if pre_tax_price is not None else "",
            FACTORY_DETAIL_HEADERS[7]: _decimal_text(tax_price) if tax_price is not None else "",
            FACTORY_DETAIL_HEADERS[8]: "",
            FACTORY_DETAIL_HEADERS[9]: "",
            FACTORY_DETAIL_HEADERS[10]: order_number,
            FACTORY_DETAIL_HEADERS[11]: _factory_remark(
                projected.get("备注"), original_roll_quantity
            ),
        }
        missing_required = [
            header
            for header in [FACTORY_DETAIL_HEADERS[0], FACTORY_DETAIL_HEADERS[3], FACTORY_DETAIL_HEADERS[5]]
            if not clean_text(row.get(header))
        ]
        if missing_required:
            review_rows += 1
            _append_issue_once(
                document,
                {
                    "page_index": detail.get("page_index", 0),
                    "region": "厂内模板",
                    "field": "、".join(missing_required),
                    "raw_value": clean_text(detail.get("raw_text")),
                    "clean_value": "",
                    "confidence": "需复核",
                    "message": f"厂内模板必填字段缺失：{'、'.join(missing_required)}；该行已保留，未猜测填值。",
                },
            )
        else:
            ready_rows += 1
        factory_rows.append(row)
        evidence_rows.append({"row_index": index, "evidence": row_evidence, "missing_required": missing_required})

    if not order_number:
        _append_issue_once(
            document,
            {
                "page_index": "",
                "region": "厂内模板",
                "field": "客户订单号",
                "raw_value": clean_text(document.get("source_file")),
                "clean_value": "",
                "confidence": "需复核",
                "message": "未识别到客户订单号，主档和明细客户订单号保持为空。",
            },
        )

    main_values = ["220", "1", "1", "", "", "", "", order_number, "", "", "", ""]
    summary = {
        "template_id": template_id,
        "template_label": clean_text(document.get("template_label")),
        "customer": _customer_identity(document),
        "order_number": order_number,
        "order_number_source": order_source,
        "row_count": len(factory_rows),
        "ready_rows": ready_rows,
        "review_rows": review_rows,
        "mapping_methods": dict(method_counts),
        "ai_header_mapping": ai_summary,
    }
    document["factory_import"] = {
        "main_headers": list(FACTORY_MAIN_HEADERS),
        "main_values": main_values,
        "detail_headers": list(FACTORY_DETAIL_HEADERS),
        "rows": factory_rows,
        "mapping_evidence": evidence_rows,
    }
    document["factory_mapping_summary"] = summary
    return summary


def safe_result_stem(source_file: Any, fallback: str = "采购单") -> str:
    stem = Path(clean_text(source_file)).stem or fallback
    stem = re.sub(r"[<>:\"/\\|?*\x00-\x1f]+", "_", stem).strip(" ._")
    return (stem or fallback)[:80]
