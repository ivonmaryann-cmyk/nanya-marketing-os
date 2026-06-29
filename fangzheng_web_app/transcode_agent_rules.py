from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import get_setting, set_setting
from .paths import TRANSCODE_AGENT_RULES_DIR, TRANSCODE_AGENT_RULES_VERSIONS_DIR


FEATURE_KEY = "transcode_agent"
RULE_FILENAME = "transcode_agent_rules.xlsx"
ORIGINAL_FILENAME = "customer_special_master.xlsx"

MACHINE_RULE_HEADERS = [
    "规则ID",
    "启用",
    "客户代码",
    "客户简称",
    "物料类别",
    "来源字段",
    "原始字段",
    "规则文本",
    "条件文本",
    "条件胶系",
    "条件关键词",
    "条件铜厚",
    "覆盖字段",
    "覆盖值",
    "命中来源",
    "优先级",
    "强制执行",
    "待确认",
    "来源行号",
]

FIELD_TO_OVERRIDE = {
    "胶系": "glue_code",
    "基板厚度": "thickness_code",
    "铜箔规格": "copper_code",
    "基板尺寸": "size_code",
    "胶水类别": "glue_category_code",
    "铜箔类型+印字/非印字": "copper_type_code",
    "基板级别": "grade_code",
    "总/芯厚": "tc_code",
    "组合结构": "struct_code",
    "铜箔厂商": "copper_vendor",
    "玻布厂商": "cloth_vendor",
    "配方代码": "formula_code",
    "PP长度": "pp_length",
    "PP级别": "pp_grade",
    "PP窄幅宽": "pp_width",
    "GT长短秒": "pp_gt",
    "树脂含量": "pp_rc",
    "小片尺寸": "pp_piece_size",
}

EXECUTABLE_FIELDS = {
    "glue_code",
    "thickness_code",
    "copper_code",
    "size_code",
    "glue_category_code",
    "copper_type_code",
    "grade_code",
    "tc_code",
    "struct_code",
}

COPPER_TYPE_VALUE_MAP = {
    "HVLP5": "J",
    "HVLP4": "Z",
    "HVLP3": "K",
    "HVLP2": "P",
    "HVLP1": "O",
    "HVLP": "O",
    "RTF4": "G",
    "RTF3": "A",
    "RTF2": "B",
    "RTF1": "R",
    "RTF": "R",
    "VLP": "L",
    "有水印": "Q",
}


def _history_key() -> str:
    return "transcode_agent_rule_history"


def _active_key() -> str:
    return "active_transcode_agent_rule_version"


def _read_history() -> list[dict]:
    raw = get_setting(_history_key(), "[]") or "[]"
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _write_history(history: list[dict]) -> None:
    set_setting(_history_key(), json.dumps(history[:50], ensure_ascii=False))


def get_transcode_agent_rule_history() -> list[dict]:
    return _read_history()


def get_active_transcode_agent_rule_version() -> str:
    return get_setting(_active_key(), "") or ""


def get_transcode_agent_rule_dir(version: str | None = None) -> Path:
    rule_version = version or get_active_transcode_agent_rule_version()
    return TRANSCODE_AGENT_RULES_VERSIONS_DIR / rule_version if rule_version else TRANSCODE_AGENT_RULES_DIR


def get_transcode_agent_rule_file_path(version: str | None = None) -> Path:
    return get_transcode_agent_rule_dir(version) / RULE_FILENAME


def get_transcode_agent_original_file_path(version: str | None = None) -> Path:
    return get_transcode_agent_rule_dir(version) / ORIGINAL_FILENAME


def save_new_transcode_agent_rule_version(
    rule_file: FileStorage,
    *,
    updated_by: str,
    remark: str = "",
) -> str:
    if not rule_file or not rule_file.filename:
        raise ValueError("请上传客户特殊清单结构化 Excel 文件")
    if not rule_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("客户特殊清单仅支持 .xlsx / .xlsm 文件")

    version = datetime.now().strftime("transcode_agent_rules_%Y%m%d_%H%M%S")
    version_dir = TRANSCODE_AGENT_RULES_VERSIONS_DIR / version
    version_dir.mkdir(parents=True, exist_ok=True)
    original_path = version_dir / ORIGINAL_FILENAME
    rule_path = version_dir / RULE_FILENAME
    rule_file.save(original_path)

    rules, summary = parse_customer_special_master(original_path)
    build_machine_rule_workbook(rule_path, rules, summary)
    set_setting(_active_key(), version)

    history = _read_history()
    history.insert(
        0,
        {
            "version": version,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": updated_by,
            "remark": remark or "上传客户特殊清单并转换为营销转码Agent规则",
            "source_file": secure_filename(rule_file.filename) or ORIGINAL_FILENAME,
            "rule_count": len(rules),
        },
    )
    _write_history(history)
    return version


def parse_customer_special_master(path: Path) -> tuple[list[dict], dict]:
    workbook = openpyxl.load_workbook(path, data_only=True)
    worksheet = workbook["客户特殊清单"] if "客户特殊清单" in workbook.sheetnames else workbook.worksheets[0]
    header_map = {
        _clean(worksheet.cell(1, col).value): col
        for col in range(1, worksheet.max_column + 1)
        if _clean(worksheet.cell(1, col).value)
    }
    required = ["客户代码", "客户简称", "CCL特殊规则", "PP特殊规则", "PP小片特殊规则"]
    missing = [name for name in required if name not in header_map]
    if missing:
        raise ValueError(f"客户特殊清单缺少列：{', '.join(missing)}")

    rows: list[dict] = []
    for row_idx in range(2, worksheet.max_row + 1):
        customer_code = _clean(worksheet.cell(row_idx, header_map["客户代码"]).value)
        customer_name = _clean(worksheet.cell(row_idx, header_map["客户简称"]).value)
        if not customer_code and not customer_name:
            continue
        if _is_template_row(customer_code, customer_name):
            continue
        for source_col, material_type in [
            ("CCL特殊规则", "CCL"),
            ("PP特殊规则", "PP"),
            ("PP小片特殊规则", "PP小片"),
        ]:
            source_text = _clean_multiline(worksheet.cell(row_idx, header_map[source_col]).value)
            rows.extend(_parse_structured_cell(customer_code, customer_name, material_type, source_col, source_text, row_idx))
        common_text = _clean_multiline(worksheet.cell(row_idx, header_map.get("通用特殊规则", 0)).value) if "通用特殊规则" in header_map else ""
        if common_text:
            rows.extend(_parse_free_text_rule(customer_code, customer_name, "通用", "通用特殊规则", common_text, row_idx))
        non_transcode_text = _clean_multiline(worksheet.cell(row_idx, header_map.get("非影响转码备注", 0)).value) if "非影响转码备注" in header_map else ""
        if non_transcode_text:
            rows.append(_make_rule(customer_code, customer_name, "通用", "非影响转码备注", "非影响转码备注", non_transcode_text, row_idx, "non_transcode_note", "", pending="否"))

    for idx, rule in enumerate(rows, 1):
        rule["规则ID"] = f"TAR-{idx:05d}"

    summary = {
        "source_path": str(path),
        "rule_count": len(rows),
        "customer_count": len({(_clean(rule["客户代码"]), _clean(rule["客户简称"])) for rule in rows}),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return rows, summary


def build_machine_rule_workbook(path: Path, rules: list[dict], summary: dict) -> None:
    workbook = Workbook()
    ws = workbook.active
    ws.title = "机器规则"
    ws.append(MACHINE_RULE_HEADERS)
    for rule in rules:
        ws.append([rule.get(header, "") for header in MACHINE_RULE_HEADERS])

    ws_summary = workbook.create_sheet("转换说明")
    ws_summary.append(["项目", "值"])
    for key, value in summary.items():
        ws_summary.append([key, value])
    ws_summary.append(["说明", "该文件由客户级特殊清单转换生成，供营销转码Agent读取；旧转码功能不读取该规则。"])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col, width in enumerate([14, 8, 14, 18, 10, 18, 18, 48, 34, 22, 28, 20, 18, 18, 16, 8, 10, 10, 10], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    ws_summary.column_dimensions["A"].width = 24
    ws_summary.column_dimensions["B"].width = 90
    workbook.save(path)


def load_transcode_agent_rules(version: str | None = None) -> list[dict]:
    path = get_transcode_agent_rule_file_path(version)
    if not path.exists():
        return []
    workbook = openpyxl.load_workbook(path, data_only=True)
    if "机器规则" not in workbook.sheetnames:
        return []
    worksheet = workbook["机器规则"]
    headers = [_clean(worksheet.cell(1, col).value) for col in range(1, worksheet.max_column + 1)]
    rules = []
    for row_idx in range(2, worksheet.max_row + 1):
        item = {headers[col - 1]: _clean(worksheet.cell(row_idx, col).value) for col in range(1, worksheet.max_column + 1) if headers[col - 1]}
        if item.get("规则ID"):
            rules.append(item)
    return rules


def get_transcode_agent_rule_count() -> int:
    return len(load_transcode_agent_rules())


def export_transcode_agent_rules(export_type: str) -> Path:
    rule_path = get_transcode_agent_rule_file_path()
    original_path = get_transcode_agent_original_file_path()
    if export_type == "machine":
        if not rule_path.exists():
            raise FileNotFoundError("还没有上传并转换过营销转码Agent规则")
        return rule_path
    if export_type == "original":
        if not original_path.exists():
            raise FileNotFoundError("还没有上传客户特殊清单原文件")
        return original_path
    if export_type == "full":
        if not rule_path.exists():
            raise FileNotFoundError("还没有上传并转换过营销转码Agent规则")
        export_path = TRANSCODE_AGENT_RULES_DIR / "transcode_agent_rule_package.xlsx"
        shutil.copy2(rule_path, export_path)
        return export_path
    raise ValueError("未知导出类型")


def _parse_structured_cell(customer_code: str, customer_name: str, material_type: str, source_col: str, text: str, source_row: int) -> list[dict]:
    if not text:
        return []
    rules = []
    for line in text.splitlines():
        line = _clean(line)
        if not line or "：" not in line:
            continue
        field_name, raw_value = line.split("：", 1)
        field_name = _clean(field_name)
        raw_value = raw_value.strip("；; \t")
        if not raw_value:
            continue
        override_field = FIELD_TO_OVERRIDE.get(field_name, "")
        for clause in _split_clauses(raw_value):
            if not clause:
                continue
            condition_text, override_value = _split_condition_value(clause)
            normalized_value = _normalize_override_value(override_field, override_value)
            pending = "否" if override_field in EXECUTABLE_FIELDS and normalized_value else "是"
            if _looks_uncertain(clause):
                pending = "是"
            rules.append(
                _make_rule(
                    customer_code,
                    customer_name,
                    material_type,
                    source_col,
                    field_name,
                    clause,
                    source_row,
                    override_field,
                    normalized_value,
                    condition_text=condition_text,
                    pending=pending,
                )
            )
    return rules


def _parse_free_text_rule(customer_code: str, customer_name: str, material_type: str, source_col: str, text: str, source_row: int) -> list[dict]:
    rules = []
    for clause in _split_clauses(text):
        condition_text, override_value = _split_condition_value(clause)
        override_field = _guess_free_text_override_field(clause)
        normalized_value = _normalize_override_value(override_field, override_value)
        rules.append(
            _make_rule(
                customer_code,
                customer_name,
                material_type,
                source_col,
                "通用",
                clause,
                source_row,
                override_field,
                normalized_value,
                condition_text=condition_text,
                pending="否" if override_field in EXECUTABLE_FIELDS and normalized_value else "是",
            )
        )
    return rules


def _make_rule(
    customer_code: str,
    customer_name: str,
    material_type: str,
    source_col: str,
    original_field: str,
    rule_text: str,
    source_row: int,
    override_field: str,
    override_value: str,
    *,
    condition_text: str = "",
    pending: str = "否",
) -> dict:
    condition_glue = _extract_condition_glue(condition_text or rule_text)
    condition_keywords = _extract_condition_keywords(condition_text or rule_text)
    condition_copper = _extract_condition_copper(condition_text or rule_text)
    priority = 100
    if condition_keywords:
        priority += 30
    if condition_copper:
        priority += 20
    if condition_glue:
        priority += 10
    return {
        "规则ID": "",
        "启用": "是",
        "客户代码": customer_code,
        "客户简称": customer_name,
        "物料类别": material_type,
        "来源字段": source_col,
        "原始字段": original_field,
        "规则文本": rule_text,
        "条件文本": condition_text,
        "条件胶系": condition_glue,
        "条件关键词": condition_keywords,
        "条件铜厚": condition_copper,
        "覆盖字段": override_field,
        "覆盖值": override_value,
        "命中来源": "客户特殊清单结构化母表",
        "优先级": str(priority),
        "强制执行": "是" if override_field in EXECUTABLE_FIELDS and override_value and pending != "是" else "否",
        "待确认": pending,
        "来源行号": str(source_row),
    }


def _split_clauses(value: str) -> list[str]:
    normalized = _clean(value).replace("；", ";")
    pieces = re.split(r";+|(?=当[^;；]*?=)|(?=如果[^;；]*?=)|[,，](?=[^,，;；]{0,30}=)", normalized)
    return [piece.strip(" ,，;；") for piece in pieces if piece.strip(" ,，;；")]


def _split_condition_value(clause: str) -> tuple[str, str]:
    text = _clean(clause)
    if "=" not in text:
        return text, ""
    left, right = text.rsplit("=", 1)
    return left.strip(" ,，"), right.strip(" ,，;；")


def _normalize_override_value(override_field: str, value: str) -> str:
    raw = _clean(value).upper()
    raw = raw.replace("（", "(").replace("）", ")")
    if not raw:
        return ""
    if override_field == "copper_code":
        compact = raw.replace(" ", "")
        if "1.5/1.5" in compact or "F/F" in compact:
            return "FF"
        if "J/J" in compact:
            return "JJ"
        if "K/K" in compact:
            return "KK"
        if "H/H" in compact:
            return "HH"
        if re.fullmatch(r"[A-Z0-9]{2}", compact):
            return compact
    if override_field == "copper_type_code":
        for keyword, code in COPPER_TYPE_VALUE_MAP.items():
            if keyword in raw:
                return code
        if re.fullmatch(r"[A-Z0-9]", raw):
            return raw
    if override_field == "tc_code":
        if "芯" in raw or raw == "C":
            return "C"
        if "总" in raw or raw == "T":
            return "T"
    if override_field == "struct_code":
        if re.fullmatch(r"[A-Z0-9*]", raw):
            return raw
    if override_field == "grade_code":
        if "汽车" in raw:
            return "AC"
        if re.fullmatch(r"[A-Z0-9]{2,4}", raw):
            return raw
    if override_field == "glue_category_code":
        if "普通" in raw or raw == "Y":
            return "Y"
        if "特殊" in raw or raw == "R":
            return "R"
    if override_field == "size_code" and re.fullmatch(r"\d{7,8}", raw):
        return raw.zfill(8)
    if override_field == "glue_code" and re.fullmatch(r"[A-Z0-9]{2,4}", raw):
        return raw
    return raw if override_field not in EXECUTABLE_FIELDS else ""


def _extract_condition_glue(text: str) -> str:
    candidates = re.findall(r"NY[-]?[A-Z0-9]+(?:HF|HC|M2|M|H|L)?", _clean(text).upper())
    return "/".join(dict.fromkeys(candidates))


def _extract_condition_copper(text: str) -> str:
    compact = _clean(text).upper().replace(" ", "")
    values = []
    for token in ("R/R", "F/F", "J/J", "K/K", "H/H", "1/1", "2/2", "0.5/0.5"):
        if token in compact:
            values.append(token)
    return "/".join(values)


def _extract_condition_keywords(text: str) -> str:
    source = _clean(text)
    match = re.search(r"(?:备注|订单|客户规格|物料描述)[^有]{0,8}有(.+?)(?:字样|时|$)", source)
    if match:
        return match.group(1).replace("或", "/").replace("、", "/").strip(" ，,")
    if "汽车板" in source:
        return "汽车板"
    if "MINILED" in source.upper():
        return "MINILED"
    return ""


def _guess_free_text_override_field(text: str) -> str:
    upper = _clean(text).upper()
    if "汽车板" in upper and ("AC" in upper or "料号" in upper):
        return "grade_code"
    if "HVLP" in upper or "RTF" in upper or "VLP" in upper:
        return "copper_type_code"
    if "R/R" in upper or "F/F" in upper or "J/J" in upper or "K/K" in upper:
        return "copper_code"
    return ""


def _looks_uncertain(text: str) -> bool:
    return bool(re.search(r"待确认|需确认|参考|详见|注意|特殊留意|可能|或", _clean(text)))


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def _clean_multiline(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    if text.strip().lower() == "nan":
        return ""
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]
    return "\n".join(line for line in lines if line)


def _is_template_row(customer_code: str, customer_name: str) -> bool:
    return not customer_code and not customer_name
