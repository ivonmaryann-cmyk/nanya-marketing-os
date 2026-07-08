from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class TemplateMatch:
    template_id: str
    label: str
    score: int
    supplier: str = ""


TEMPLATES: list[dict[str, Any]] = [
    {
        "template_id": "ganzhou_chaoyue_purchase_order",
        "label": "赣州超跃采购订单",
        "filename_keywords": ["赣州超跃", "超跃"],
        "strong_text_keywords": ["赣州市超跃科技股份有限公司", "Ganzhou Beyond", "采购单号: GA", "采购单号：GA"],
        "field_keywords": ["交易条件", "付款条件", "送货地址", "料件编号"],
        "weak_text_keywords": ["采购订单", "供应商", "采购日期"],
        "supplier": "",
        "min_score": 60,
    },
    {
        "template_id": "ganzhou_yihao_purchase_order",
        "label": "赣州逸豪采购订单",
        "filename_keywords": ["赣州逸豪", "逸豪"],
        "strong_text_keywords": ["赣州逸豪新材料股份有限公司", "Ganzhou Yihao", "表格编号", "PC B事业部", "PCB事业部"],
        "field_keywords": ["物料编码", "物料名称规格", "到货日期", "订单号"],
        "weak_text_keywords": ["采购订单", "供应商名称"],
        "supplier": "南亚新材料科技（江西）有限公司",
        "min_score": 60,
    },
    {
        "template_id": "jingwang_purchase_order",
        "label": "景旺采购订单",
        "filename_keywords": ["景旺"],
        "strong_text_keywords": ["江西景旺精密电路有限公司", "P.ONO", "VendorCode", "PaymentTerms", "Shipto"],
        "field_keywords": ["Part No.", "Description", "Quantity", "Unit Price", "Delivery Date"],
        "weak_text_keywords": ["Purchase Order", "Vendor", "Comments"],
        "supplier": "",
        "min_score": 70,
    },
    {
        "template_id": "talian_purchase_order",
        "label": "塔联采购订单",
        "filename_keywords": ["塔联"],
        "strong_text_keywords": ["深圳市塔联科技有限公司", "合同编号：TL", "合同编号:TL", "合同编号： TL", "chinatalian.com"],
        "field_keywords": ["原料编码", "原料名称", "单价RMB", "税率%", "交期"],
        "weak_text_keywords": ["采购订单", "买方", "卖方"],
        "supplier": "南亚新材料科技（江西）有限公司",
        "min_score": 70,
    },
]


def _contains_any(text: str, keywords: list[str]) -> int:
    return sum(1 for keyword in keywords if keyword and keyword.lower() in text.lower())


def identify_template(source_filename: str, text: str = "") -> TemplateMatch | None:
    filename = Path(source_filename).name
    best: TemplateMatch | None = None
    for template in TEMPLATES:
        filename_hits = _contains_any(filename, template["filename_keywords"])
        strong_hits = _contains_any(text, template["strong_text_keywords"])
        field_hits = _contains_any(text, template["field_keywords"])
        weak_hits = _contains_any(text, template["weak_text_keywords"])

        if filename_hits == 0 and strong_hits == 0:
            continue

        score = filename_hits * 90 + strong_hits * 35 + field_hits * 10 + weak_hits * 2
        if score < int(template["min_score"]):
            continue

        match = TemplateMatch(
            template_id=template["template_id"],
            label=template["label"],
            score=score,
            supplier=template.get("supplier", ""),
        )
        if best is None or match.score > best.score:
            best = match
    return best


def likely_order_number(text: str) -> str:
    patterns = [
        r"(?:订单号|采购单号|采购订单号|订单编号)\s*(?:\([^)）]*\)|（[^)）]*）)?\s*[:：]\s*([A-Za-z0-9][A-Za-z0-9_-]{4,})",
        r"(?:P\.?\s*O\.?\s*NO|PONO)\s*(?:\([^)）]*\)|（[^)）]*）)?\s*[:：]\s*([A-Za-z0-9][A-Za-z0-9_-]{4,})",
        r"\b(GA\d{3,}-\d{6,})\b",
        r"\b(P\d{6,})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            value = match.group(1).strip()
            if re.search(r"\d{5,}", value):
                return value
    return ""


def likely_supplier(text: str, template: TemplateMatch | None = None) -> str:
    if template and template.supplier:
        return template.supplier

    label_match = re.search(
        r"(?:供应商名称|供应商|Vendor)\s*(?:\([^)）]*\)|（[^)）]*）)?\s*[:：]\s*(.+?)(?:\s+(?:供应商代码|VendorCode|订购日期|P\.ODate)|\n|$)",
        text,
        flags=re.I,
    )
    if label_match:
        return label_match.group(1).strip(" ：:")

    if "南亚新材料科技（江西）有限公司" in text:
        return "南亚新材料科技（江西）有限公司"
    if "南亚新材料科技股份有限公司" in text:
        return "南亚新材料科技股份有限公司"
    if "南亚新材料" in text:
        return "南亚新材料科技（江西）有限公司"
    return ""
