from __future__ import annotations

import re
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Any

import openpyxl


IN_SCOPE_FIELDS = (
    "胶系",
    "基板厚度",
    "铜箔规格",
    "基板尺寸",
    "胶水类别",
    "铜箔类型+印字/非印字",
    "基板级别",
    "总/芯厚",
)
OUT_OF_SCOPE_FIELDS = ("组合结构", "铜箔厂商", "玻布厂商", "配方代码")
ALL_RULE_FIELDS = IN_SCOPE_FIELDS + OUT_OF_SCOPE_FIELDS
SOURCE_COLUMNS = ("CCL特殊规则", "通用特殊规则", "非影响转码备注")
SOURCE_COLUMN_CODES = {
    "CCL特殊规则": "",
    "通用特殊规则": "G",
    "非影响转码备注": "N",
}

FIELD_PATTERN = re.compile(
    r"(铜箔类型\s*\+\s*印字\s*/\s*非印字|胶水类别|基板厚度|铜箔规格|基板尺寸|基板级别|总\s*/\s*芯厚|组合结构|铜箔厂商|玻布厂商|配方代码|胶系)\s*[:：]",
    flags=re.IGNORECASE,
)
EXTERNAL_LOOKUP_PATTERN = re.compile(r"参考.+表|换算表|详情见.+表", flags=re.IGNORECASE)
SEMANTIC_CONTEXT_PATTERN = re.compile(
    r"订单|备注|客户料品|料品名称|客户物料|物料编码|订单规格|下单有|下单无|有双面|无双面",
    flags=re.IGNORECASE,
)
SEMANTIC_EXCLUSION_PATTERN = re.compile(r"以外|除.+外|其余|除开|不属于", flags=re.IGNORECASE)
SEMANTIC_NEGATION_PATTERN = re.compile(
    r"没有\s*[A-Za-z0-9\u4e00-\u9fff]+|未备注|未写|未标|无该字样|没有写|不含铜会备注",
    flags=re.IGNORECASE,
)
SEMANTIC_POSITION_PATTERN = re.compile(r"第\s*\d+\s*码|最后一位|末位", flags=re.IGNORECASE)
CUSTOMER_SPEC_SEMANTIC_PATTERN = re.compile(
    r"客户规格.{0,16}(没有|未|字样|描述|包含|带有|有[^卤铜厚])",
    flags=re.IGNORECASE,
)
UNSTRUCTURED_TARGET_PATTERNS = (
    ("基板级别", re.compile(r"汽车板|电源板|能源板|MINI\s*LED|基板级别|\b(?:A[1-9C-Z]|F1|PG)\s*等级", re.IGNORECASE)),
    ("铜箔类型+印字/非印字", re.compile(r"水印|印字|非印字|(?<![A-Z0-9])(?:HTE|RTF\d*|HVLP\d*|VLP|IGAV)(?![A-Z0-9])", re.IGNORECASE)),
    ("总/芯厚", re.compile(r"总厚|芯厚|含铜|不含铜|不连铜", re.IGNORECASE)),
    ("基板尺寸", re.compile(r"尺寸.{0,12}(?:放大|映射|录入)|(?:放大|录入).{0,12}尺寸", re.IGNORECASE)),
    ("铜箔规格", re.compile(r"(?:R\s*/\s*R|H\s*/\s*H|1\s*/\s*1).{0,18}(?:对应|=)", re.IGNORECASE)),
)


@dataclass(frozen=True)
class SemanticRuleCandidate:
    candidate_id: str
    customer_code: str
    customer_name: str
    source_row: int
    source_column: str
    business_field: str
    source_text: str
    semantic_type: str
    required_input_fields: str
    execution_path: str
    status: str = "待模型JSON化"
    model_json: str = ""
    validation_result: str = ""
    business_confirmation: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def normalize_rule_field_name(value: str) -> str:
    name = re.sub(r"\s+", "", str(value or ""))
    if name == "铜箔类型+印字/非印字":
        return name
    if name == "总/芯厚":
        return name
    return name


def split_ccl_rule_fields(rule_text: str) -> dict[str, str]:
    text = str(rule_text or "").strip()
    if not text or text == "21":
        return {}
    matches = list(FIELD_PATTERN.finditer(text))
    fields: dict[str, str] = {}
    for index, match in enumerate(matches):
        field = normalize_rule_field_name(match.group(1))
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        content = text[start:end].strip(" \t\r\n;；,，")
        if field and is_meaningful_rule_text(content):
            fields[field] = content
    return fields


def is_meaningful_rule_text(value: Any) -> bool:
    text = re.sub(r"[\s;；,，:：]", "", str(value or ""))
    return bool(text and text.lower() not in {"21", "无", "无规则", "none", "nan"})


def split_atomic_clauses(value: str) -> list[str]:
    text = str(value or "").strip()
    text = re.sub(r"[,，](?=\s*当)", "；", text)
    text = re.sub(r"[,，](?=\s*客户订单)", "；", text)
    clauses = [item.strip(" \t\r\n;；") for item in re.split(r"[;；\n]+", text)]
    return [item for item in clauses if is_meaningful_rule_text(item)]


def infer_unstructured_rule_clauses(value: str) -> list[tuple[str, str]]:
    clauses: list[tuple[str, str]] = []
    for clause in split_atomic_clauses(value):
        for field, pattern in UNSTRUCTURED_TARGET_PATTERNS:
            if pattern.search(clause):
                clauses.append((field, clause))
    return clauses


def requires_semantic_model(clause: str) -> bool:
    text = str(clause or "").strip()
    if not text:
        return False
    if EXTERNAL_LOOKUP_PATTERN.search(text) and not (
        SEMANTIC_CONTEXT_PATTERN.search(text)
        or SEMANTIC_EXCLUSION_PATTERN.search(text)
        or SEMANTIC_NEGATION_PATTERN.search(text)
    ):
        return False
    return bool(
        SEMANTIC_CONTEXT_PATTERN.search(text)
        or SEMANTIC_EXCLUSION_PATTERN.search(text)
        or SEMANTIC_NEGATION_PATTERN.search(text)
        or SEMANTIC_POSITION_PATTERN.search(text)
        or CUSTOMER_SPEC_SEMANTIC_PATTERN.search(text)
    )


def infer_semantic_type(clause: str) -> str:
    text = str(clause or "")
    if SEMANTIC_POSITION_PATTERN.search(text):
        return "字符位置条件"
    if SEMANTIC_EXCLUSION_PATTERN.search(text):
        return "排除集合"
    if SEMANTIC_NEGATION_PATTERN.search(text):
        if "默认" in text:
            return "缺失时默认"
        return "否定/缺失条件"
    if SEMANTIC_CONTEXT_PATTERN.search(text) or CUSTOMER_SPEC_SEMANTIC_PATTERN.search(text):
        return "上下文关键词条件"
    return "自然语言条件"


def infer_required_input_fields(clause: str, business_field: str) -> list[str]:
    text = str(clause or "")
    fields: list[str] = []

    def add(value: str) -> None:
        if value not in fields:
            fields.append(value)

    if "备注" in text:
        add("订单备注")
    if "客户规格" in text or "订单规格" in text:
        add("客户规格")
    if "客户料品" in text or "料品名称" in text:
        add("客户料品名称")
    if "物料编码" in text or SEMANTIC_POSITION_PATTERN.search(text):
        add("客户物料编码/订单备注")
    if re.search(r"(?:品号|料号).{0,8}(?:中|第\s*\d+\s*码|最后一位|末位)", text):
        add("品号/物料编号")
    if ("订单" in text or "下单" in text) and not any(value.startswith("订单") for value in fields):
        add("订单规格/订单备注")
    if "胶系" in text or SEMANTIC_EXCLUSION_PATTERN.search(text):
        add("胶系")
    if "厚度" in text:
        add("基板厚度")
    if "铜箔" in text or "铜厚" in text:
        add("铜箔规格")
    if not fields:
        add(business_field)
    return fields


def build_semantic_candidates(
    *,
    customer_code: str,
    customer_name: str,
    source_row: int,
    ccl_rule: str,
    source_column: str = "CCL特殊规则",
) -> list[SemanticRuleCandidate]:
    fields = split_ccl_rule_fields(ccl_rule)
    raw_candidates: list[tuple[str, str]] = []
    for field in IN_SCOPE_FIELDS:
        for clause in split_atomic_clauses(fields.get(field, "")):
            if requires_semantic_model(clause):
                raw_candidates.append((field, clause))

    if source_column != "CCL特殊规则":
        for field, clause in infer_unstructured_rule_clauses(ccl_rule):
            item = (field, clause)
            if item not in raw_candidates:
                raw_candidates.append(item)

    candidates: list[SemanticRuleCandidate] = []
    for index, (field, clause) in enumerate(raw_candidates, start=1):
        source_code = SOURCE_COLUMN_CODES.get(source_column, "X")
        suffix = f"-{source_code}" if source_code else ""
        candidate_id = f"MSR-{int(source_row):04d}{suffix}-{index:02d}"
        candidates.append(
            SemanticRuleCandidate(
                candidate_id=candidate_id,
                customer_code=str(customer_code or "").strip(),
                customer_name=str(customer_name or "").strip(),
                source_row=int(source_row),
                source_column=source_column,
                business_field=field,
                source_text=clause,
                semantic_type=infer_semantic_type(clause),
                required_input_fields="；".join(infer_required_input_fields(clause, field)),
                execution_path="模型语义",
            )
        )
    return candidates


def classify_draft_row(row: dict[str, Any]) -> dict[str, Any]:
    source_row = _to_int(row.get("来源行号"))
    candidates: list[SemanticRuleCandidate] = []
    for source_column in SOURCE_COLUMNS:
        candidates.extend(
            build_semantic_candidates(
                customer_code=str(row.get("客户代码") or ""),
                customer_name=str(row.get("客户简称") or ""),
                source_row=source_row,
                ccl_rule=str(row.get(source_column) or ""),
                source_column=source_column,
            )
        )
    has_standard = any(
        is_meaningful_rule_text(row.get(column))
        for column in ("CCL特殊规则_结构化", "通用特殊规则_结构化")
    )
    current_status = str(row.get("结构化处理状态") or "").strip()
    if candidates and has_standard:
        execution_path = "标准规则+模型语义"
    elif candidates:
        execution_path = "模型语义"
    elif has_standard:
        execution_path = "标准规则"
    elif current_status == "不纳入":
        execution_path = "不纳入"
    else:
        execution_path = "无规则"

    candidates = [replace(candidate, execution_path=execution_path) for candidate in candidates]

    candidate_text = "\n".join(
        f"{item.candidate_id} | 字段={item.business_field} | 类型={item.semantic_type} | 原文={item.source_text}"
        for item in candidates
    )
    required_fields = []
    for candidate in candidates:
        for field in candidate.required_input_fields.split("；"):
            if field and field not in required_fields:
                required_fields.append(field)
    return {
        "execution_path": execution_path,
        "candidate_text": candidate_text,
        "required_input_fields": "；".join(required_fields),
        "model_status": "待模型JSON化" if candidates else "不需要模型",
        "candidate_ids": "；".join(item.candidate_id for item in candidates),
        "candidates": [item.to_dict() for item in candidates],
    }


def classify_draft_workbook(path: str | Path) -> dict[str, Any]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook["结构化草稿"]
    headers = [str(cell.value or "").strip() for cell in worksheet[1]]
    rows: list[dict[str, Any]] = []
    candidates: list[dict[str, Any]] = []
    path_counts: dict[str, int] = {}
    for values in worksheet.iter_rows(min_row=2, values_only=True):
        row = {headers[index]: values[index] for index in range(min(len(headers), len(values)))}
        if not (str(row.get("客户代码") or "").strip() or str(row.get("客户简称") or "").strip()):
            continue
        result = classify_draft_row(row)
        item = {
            "excel_row": len(rows) + 2,
            "customer_code": str(row.get("客户代码") or "").strip(),
            "customer_name": str(row.get("客户简称") or "").strip(),
            "source_row": _to_int(row.get("来源行号")),
            **{key: value for key, value in result.items() if key != "candidates"},
        }
        rows.append(item)
        candidates.extend(result["candidates"])
        path_counts[item["execution_path"]] = path_counts.get(item["execution_path"], 0) + 1
    return {
        "source_file": str(Path(path).resolve()),
        "row_count": len(rows),
        "candidate_count": len(candidates),
        "candidate_customer_count": len(
            {(item["customer_code"], item["customer_name"]) for item in candidates}
        ),
        "path_counts": path_counts,
        "rows": rows,
        "candidates": candidates,
    }


def _to_int(value: Any) -> int:
    try:
        return int(float(str(value or "0").strip()))
    except ValueError:
        return 0
