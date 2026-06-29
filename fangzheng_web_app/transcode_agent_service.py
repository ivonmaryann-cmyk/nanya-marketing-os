from __future__ import annotations

import importlib
import re
import sys
import traceback
from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import append_job_log, create_job, get_job, prune_jobs_for_employee, update_job_status
from .excel_utils import load_workbook_compat, normalized_xlsx_source
from .file_utils import safe_unlink
from .job_control import launch_job_process
from .paths import JOBS_DIR
from .transcode_agent_rules import (
    FEATURE_KEY,
    get_active_transcode_agent_rule_version,
    load_transcode_agent_rules,
)
from .transcode_rules import get_active_transcode_rule_version, get_transcode_rule_file_path


TRANSCODE_MODULE_NAME = "fangzheng_web_app.transcode_engine"
FIELD_GATE_THRESHOLD = 85
FORMAL_RESULT_HEADER = "Agent转码结果"


FIELD_DEFS = [
    ("glue", "胶系", "step1_glue_code", True),
    ("thickness", "厚度", "step2_thick_code", True),
    ("copper", "铜厚", "step3_copper_code", True),
    ("size", "尺寸", "step4_size_code", True),
    ("glue_category", "胶水类别", "step5_glue_cat_code", True),
    ("copper_type", "铜箔类型", "step6_copper_type_code", True),
    ("grade", "基板级别", "step7_grade_code", True),
    ("total_core", "总/芯厚", "step8_tc_code", True),
    ("structure", "结构码", "step9_struct_code", False),
]

OVERRIDE_STEP_MAP = {
    "glue_code": ("step1_glue_code", "无法识别胶系型号"),
    "thickness_code": ("step2_thick_code", "无法识别厚度"),
    "copper_code": ("step3_copper_code", "无法识别铜箔规格"),
    "size_code": ("step4_size_code", "无法识别尺寸"),
    "glue_category_code": ("step5_glue_cat_code", ""),
    "copper_type_code": ("step6_copper_type_code", ""),
    "grade_code": ("step7_grade_code", ""),
    "tc_code": ("step8_tc_code", ""),
    "struct_code": ("step9_struct_code", ""),
}


def load_transcode_module():
    if TRANSCODE_MODULE_NAME in sys.modules:
        return importlib.reload(sys.modules[TRANSCODE_MODULE_NAME])
    return importlib.import_module(TRANSCODE_MODULE_NAME)


def queue_transcode_agent_job(employee_id: str, uploaded_file: FileStorage, source_filename: str) -> int:
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    employee_dir = JOBS_DIR / employee_id
    employee_dir.mkdir(parents=True, exist_ok=True)
    safe_filename = secure_filename(source_filename) or f"transcode_agent_{timestamp}.xlsx"
    input_path = employee_dir / f"{timestamp}_transcode_agent_{safe_filename}"
    uploaded_file.save(input_path)
    base_version = get_active_transcode_rule_version()
    agent_version = get_active_transcode_agent_rule_version() or "未上传"
    job_id = create_job(
        employee_id,
        source_filename,
        str(input_path),
        f"base:{base_version};agent:{agent_version}",
        feature=FEATURE_KEY,
    )
    launch_job_process(job_id, FEATURE_KEY, employee_id)
    return job_id


def calculate_transcode_agent_quote(
    spec: str,
    *,
    customer: str = "",
    customer_code: str = "",
    order_text: str = "",
) -> dict:
    spec = str(spec or "").strip()
    if not spec:
        return {"status": "失败", "result": None, "error": "请输入客户规格"}
    engine, tables, agent_rules, base_version, agent_version = _load_runtime()
    analysis = analyze_spec(
        engine,
        tables,
        agent_rules,
        spec,
        customer=customer,
        customer_code=customer_code,
        context_text=order_text,
    )
    return {
        "status": analysis["status"],
        "result": analysis["formal_code"],
        "candidate_code": analysis["candidate_code"],
        "note": analysis["summary"],
        "error": analysis["reason"] if analysis["status"] != "成功" else "",
        "confidence": analysis["overall_score"],
        "field_evidence": analysis["field_evidence"],
        "rule_version": base_version,
        "agent_rule_version": agent_version or "未上传",
    }


def run_transcode_agent_job(job_id: int, employee_id: str) -> None:
    update_job_status(job_id, status="running", log_text="")
    job = get_job(job_id)
    if not job:
        return
    append_job_log(job_id, f"开始营销转码Agent任务，规则版本：{job['rule_version']}")
    try:
        engine, tables, agent_rules, base_version, agent_version = _load_runtime()
        append_job_log(job_id, f"基础转码规则：{base_version}；Agent规则：{agent_version or '未上传'}")

        workbook = load_workbook_compat(job["stored_input_path"], data_only=True)
        source_for_result = normalized_xlsx_source(job["stored_input_path"], workbook)
        sheets, _ = engine.load_transcode_inputs(str(source_for_result), str(get_transcode_rule_file_path(base_version)))
        df_req = sheets["转码需求表"].copy()
        spec_col = engine.select_transcode_spec_column(df_req)
        customer_col = engine.detect_customer_column(df_req, spec_col)
        customer_code_col = engine.detect_customer_code_column(df_req)
        context_cols = engine.detect_transcode_context_columns(df_req, spec_col, customer_col, customer_code_col)
        result_col = _ensure_agent_result_column(df_req)
        data_indices = [i for i in range(1, len(df_req)) if _is_effective_spec(df_req.iloc[i, spec_col], engine)]
        total_rows = len(data_indices)
        update_job_status(job_id, status="running", total_rows=total_rows)
        append_job_log(job_id, f"识别规格列：第 {spec_col + 1} 列；Agent结果写入：第 {result_col + 1} 列")
        append_job_log(job_id, f"检测到 {total_rows} 行有效规格数据", total_rows=total_rows)

        analyses: list[dict] = []
        success_count = fail_count = skip_count = confirm_count = 0
        for processed, i in enumerate(data_indices, start=1):
            row = df_req.iloc[i]
            customer_code = engine._clean_cell(row.iloc[customer_code_col]) if customer_code_col is not None and len(row) > customer_code_col else ""
            customer = str(row.iloc[customer_col]).strip() if customer_col is not None and len(row) > customer_col and pd.notna(row.iloc[customer_col]) else ""
            spec = engine._clean_cell(row.iloc[spec_col])
            context = engine.build_context_text_from_row(row, context_cols)
            cust_spec = engine._clean_cell(row.iloc[6]) if len(row) > 6 else ""
            normalized_spec = engine._clean_cell(row.iloc[7]) if len(row) > 7 else ""
            pp_check_text = " ".join([spec, cust_spec, normalized_spec])
            excel_row = i + 1

            if engine.is_pp_or_rc_spec(pp_check_text):
                skip_count += 1
                result_text = "跳过：PP/RC/% 暂不输出CCL制造编码"
                df_req.iloc[i, result_col] = result_text
                analyses.append(_skip_analysis(excel_row, customer_code, customer, spec, result_text))
                append_job_log(job_id, f"第 {excel_row} 行跳过：PP/RC/%", skip_count=skip_count, current_row=processed, total_rows=total_rows)
                continue

            analysis = analyze_spec(
                engine,
                tables,
                agent_rules,
                spec,
                customer=customer,
                customer_code=customer_code,
                context_text=context,
                excel_row=excel_row,
            )
            analyses.append(analysis)
            if analysis["status"] == "成功":
                success_count += 1
                df_req.iloc[i, result_col] = analysis["formal_code"]
                log_text = f"第 {excel_row} 行高置信出码：{analysis['formal_code']}"
            elif analysis["status"] == "待确认":
                confirm_count += 1
                fail_count += 1
                df_req.iloc[i, result_col] = f"待确认：{analysis['reason']}"
                log_text = f"第 {excel_row} 行待确认：{analysis['reason']}"
            else:
                fail_count += 1
                df_req.iloc[i, result_col] = f"未识别：{analysis['reason']}"
                log_text = f"第 {excel_row} 行未识别：{analysis['reason']}"
            append_job_log(
                job_id,
                log_text,
                success_count=success_count,
                fail_count=fail_count,
                skip_count=skip_count,
                current_row=processed,
                total_rows=total_rows,
            )

        input_path = Path(job["stored_input_path"])
        output_path = input_path.with_name(f"{input_path.stem}_Agent转码结果.xlsx")
        _save_agent_result(str(source_for_result), str(output_path), df_req, result_col, analyses, agent_rules, confirm_count)
        append_job_log(job_id, "Agent结果文件已生成，任务完成", current_row=total_rows, total_rows=total_rows)
        update_job_status(
            job_id,
            status="completed",
            stored_result_path=str(output_path),
            success_count=success_count,
            fail_count=fail_count,
            skip_count=skip_count,
            current_row=total_rows,
            total_rows=total_rows,
            completed=True,
        )
    except Exception as exc:
        append_job_log(job_id, f"任务失败：{exc}")
        update_job_status(
            job_id,
            status="failed",
            error_message=f"{exc}\n{traceback.format_exc(limit=8)}",
            completed=True,
        )
    finally:
        stale_jobs = prune_jobs_for_employee(employee_id, keep_limit=500)
        for stale in stale_jobs:
            for key in ["stored_input_path", "stored_result_path"]:
                safe_unlink(stale[key])


def analyze_spec(
    engine,
    tables: dict,
    agent_rules: list[dict],
    spec: str,
    *,
    customer: str = "",
    customer_code: str = "",
    context_text: str = "",
    excel_row: int | None = None,
) -> dict:
    runtime_tables = dict(tables)
    runtime_tables["structured_special_rules"] = []
    code, steps, err = engine.transcode_row(
        engine._clean_cell(spec),
        "",
        str(customer or "").strip(),
        str(customer_code or "").strip(),
        runtime_tables,
        str(context_text or "").strip(),
    )
    steps = dict(steps or {})
    errors = list(steps.get("errors") or [])
    applied_rules, conflicts = _apply_agent_rules(agent_rules, customer_code, customer, spec, context_text, steps, errors)
    candidate_code = _build_code_from_steps(steps, errors)
    field_evidence = _build_field_evidence(steps, errors, applied_rules, conflicts)
    gate_scores = [item["score"] for item in field_evidence if item["gate"]]
    overall_score = min(gate_scores) if gate_scores else 0
    low_fields = [item["field"] for item in field_evidence if item["gate"] and item["score"] < FIELD_GATE_THRESHOLD]

    if errors:
        status = "失败"
        formal_code = ""
        reason = "; ".join(errors)
    elif conflicts:
        status = "待确认"
        formal_code = ""
        reason = "Agent规则冲突：" + "; ".join(conflicts)
    elif not candidate_code:
        status = "失败"
        formal_code = ""
        reason = "无法生成候选编码"
    elif low_fields:
        status = "待确认"
        formal_code = ""
        reason = f"低置信字段：{', '.join(low_fields)}"
    else:
        status = "成功"
        formal_code = candidate_code
        reason = ""

    summary = _format_agent_summary(status, formal_code or candidate_code, overall_score, reason, applied_rules)
    return {
        "row": excel_row,
        "customer_code": customer_code,
        "customer": customer,
        "spec": spec,
        "status": status,
        "formal_code": formal_code,
        "candidate_code": candidate_code,
        "overall_score": overall_score,
        "reason": reason,
        "summary": summary,
        "field_evidence": field_evidence,
        "applied_rules": applied_rules,
        "conflicts": conflicts,
        "engine_steps": steps,
    }


def _load_runtime():
    engine = load_transcode_module()
    base_version = get_active_transcode_rule_version()
    rule_path = get_transcode_rule_file_path(base_version)
    tables = engine.build_lookup_tables(engine.load_rule_sheets(str(rule_path)))
    tables["structured_special_rules"] = []
    agent_version = get_active_transcode_agent_rule_version()
    agent_rules = load_transcode_agent_rules(agent_version) if agent_version else []
    return engine, tables, agent_rules, base_version, agent_version


def _apply_agent_rules(
    rules: list[dict],
    customer_code: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
    errors: list[str],
) -> tuple[list[dict], list[str]]:
    values: dict[str, str] = {}
    priorities: dict[str, int] = {}
    applied: list[dict] = []
    conflicts: list[str] = []
    for rule in sorted(rules, key=lambda item: _rule_priority(item), reverse=True):
        if not _rule_executable(rule):
            continue
        if not _rule_matches(rule, customer_code, customer_name, spec, context, steps):
            continue
        field = rule.get("覆盖字段", "")
        value = str(rule.get("覆盖值", "") or "").strip().upper()
        if field in values and values[field] != value:
            if priorities.get(field, 0) > _rule_priority(rule):
                continue
            conflicts.append(f"{field}: {values[field]} vs {value} ({rule.get('规则ID')})")
            continue
        step_key, error_text = OVERRIDE_STEP_MAP.get(field, ("", ""))
        if not step_key:
            continue
        old_value = str(steps.get(step_key, "") or "")
        values[field] = value
        priorities[field] = _rule_priority(rule)
        steps[step_key] = value
        if field == "tc_code":
            steps["order_type"] = "芯厚" if value == "C" else "总厚"
        if error_text:
            errors[:] = [item for item in errors if item != error_text and not item.startswith(error_text)]
            steps["errors"] = errors
        applied.append(
            {
                "rule_id": rule.get("规则ID", ""),
                "field": field,
                "old": old_value,
                "new": value,
                "text": rule.get("规则文本", ""),
                "source": rule.get("命中来源", "Agent规则包"),
            }
        )
    if applied:
        steps["agent_rules"] = [f"{item['rule_id']} {item['field']}:{item['old']}->{item['new']}" for item in applied]
    if conflicts:
        steps["agent_rule_conflicts"] = conflicts
    return applied, conflicts


def _rule_priority(rule: dict) -> int:
    try:
        return int(float(rule.get("优先级") or 0))
    except (TypeError, ValueError):
        return 0


def _rule_executable(rule: dict) -> bool:
    return (
        str(rule.get("启用", "")).strip() == "是"
        and str(rule.get("待确认", "")).strip() != "是"
        and str(rule.get("强制执行", "")).strip() == "是"
        and str(rule.get("物料类别", "")).strip() == "CCL"
        and rule.get("覆盖字段") in OVERRIDE_STEP_MAP
        and bool(str(rule.get("覆盖值", "") or "").strip())
    )


def _rule_matches(rule: dict, customer_code: str, customer_name: str, spec: str, context: str, steps: dict) -> bool:
    if not _customer_matches(rule, customer_code, customer_name):
        return False
    combined = _norm_match(f"{spec} {context}")
    spec_norm = _norm_match(spec)
    glue_norm = _norm_match(steps.get("glue_model", ""))
    copper_norm = _norm_match(steps.get("copper_spec_raw", ""))
    raw_condition_text = str(rule.get("条件文本", "") or rule.get("规则文本", "") or "")
    has_positive_tolerance = bool(re.search(r"\+\s*\d|±|偏正", str(spec or "") + " " + str(context or "")))
    if "有偏正公差" in raw_condition_text and not has_positive_tolerance:
        return False
    if "无偏正公差" in raw_condition_text and has_positive_tolerance:
        return False

    glue_condition = _norm_match(rule.get("条件胶系", ""))
    if glue_condition:
        glue_tokens = [item for item in re.split(r"[/,，;；]+", glue_condition) if item]
        if glue_tokens and not any(token in glue_norm or token in spec_norm for token in glue_tokens):
            return False

    copper_condition = _norm_match(rule.get("条件铜厚", ""))
    if copper_condition:
        copper_tokens = [item for item in re.split(r"[/,，;；]+", copper_condition) if item]
        if copper_tokens and not any(token in copper_norm or token in spec_norm for token in copper_tokens):
            return False

    keyword_condition = str(rule.get("条件关键词", "") or "").strip()
    if keyword_condition:
        keywords = [_norm_match(item) for item in re.split(r"[/,，、;；]+", keyword_condition) if item.strip()]
        if keywords and not any(keyword in combined for keyword in keywords):
            return False

    condition_text = _norm_match(rule.get("条件文本", ""))
    if condition_text and not (glue_condition or copper_condition or keyword_condition):
        field_label = _norm_match(rule.get("原始字段", ""))
        if field_label and field_label in condition_text:
            return True
        relaxed = re.sub(r"^(当|如果)|时$", "", condition_text)
        if relaxed and relaxed not in combined:
            return False
    return True


def _customer_matches(rule: dict, customer_code: str, customer_name: str) -> bool:
    rule_code = _norm_customer_code(rule.get("客户代码", ""))
    current_code = _norm_customer_code(customer_code)
    if rule_code:
        if not current_code or rule_code != current_code:
            return False
    rule_name = _norm_match(rule.get("客户简称", ""))
    current_name = _norm_match(customer_name)
    if rule_name:
        if not current_name or (rule_name not in current_name and current_name not in rule_name):
            return False
    return bool(rule_code or rule_name)


def _build_field_evidence(steps: dict, errors: list[str], applied_rules: list[dict], conflicts: list[str]) -> list[dict]:
    by_field = {item["field"]: item for item in applied_rules}
    evidence = []
    for key, label, step_key, gate in FIELD_DEFS:
        code_value = str(steps.get(step_key, "") or "")
        score, hit_type, source, note = _score_field(key, code_value, steps, errors, by_field, conflicts)
        evidence.append(
            {
                "field_key": key,
                "field": label,
                "value": _field_raw_value(key, steps),
                "code": code_value,
                "score": score,
                "gate": gate,
                "hit_type": hit_type,
                "source": source,
                "evidence": note,
            }
        )
    return evidence


def _score_field(key: str, code_value: str, steps: dict, errors: list[str], by_field: dict, conflicts: list[str]) -> tuple[int, str, str, str]:
    override_field = {
        "glue": "glue_code",
        "thickness": "thickness_code",
        "copper": "copper_code",
        "size": "size_code",
        "glue_category": "glue_category_code",
        "copper_type": "copper_type_code",
        "grade": "grade_code",
        "total_core": "tc_code",
        "structure": "struct_code",
    }[key]
    if override_field in by_field:
        rule = by_field[override_field]
        return 99, "Agent规则覆盖", rule["rule_id"], rule["text"]
    if conflicts:
        return 60, "规则冲突", "Agent规则包", "; ".join(conflicts)
    if _field_has_error(key, errors) or _is_placeholder(code_value):
        return 0, "未识别", "基础解析", "; ".join(errors)
    if key == "thickness":
        source = str(steps.get("thickness_mode_source", "") or steps.get("thickness_unit", "") or "基础解析")
        if "客户特殊" in source or "特殊板厚" in source or "健鼎" in source or "超颖" in source:
            return 97, "客户特殊厚度规则", source, str(steps.get("thickness_mode_note", "") or steps.get("thickness_raw", ""))
        if "通用阈值" in source:
            return 88, "通用阈值", source, str(steps.get("thickness_raw", ""))
        return 94, "标准解析", source, str(steps.get("thickness_raw", ""))
    if key == "size":
        note = str(steps.get("size_note", ""))
        return (96 if "特殊" in note else 94), "特殊尺寸" if "特殊" in note else "标准尺寸", note or "基础解析", f"{steps.get('size_w', '')}x{steps.get('size_h', '')}"
    if key == "grade":
        if steps.get("grade_note"):
            return 97, "客户下单转换表", "客户下单与胶系基板转换", str(steps.get("grade_note", ""))
        if code_value == "A1":
            return 86, "默认规则", "基础规则", "未命中特殊基板等级，按默认A1"
        return 94, "标准/关键词规则", "基础规则", code_value
    if key == "total_core":
        source = str(steps.get("thickness_mode_source", "") or "基础规则")
        return (86 if "通用阈值" in source else 94), "总芯厚判断", source, str(steps.get("order_type", ""))
    if key == "copper_type":
        if code_value == "W":
            return 88, "默认常规铜", "基础规则", "未命中特殊铜箔关键词"
        return 94, "关键词命中", "基础规则", code_value
    if key == "structure" and code_value == "*":
        return 80, "阶段性占位", "首版策略", "结构码首版允许*占位，不参与正式出码拦截"
    return 94, "标准解析", "基础规则", code_value


def _field_raw_value(key: str, steps: dict) -> str:
    if key == "glue":
        return str(steps.get("glue_model", ""))
    if key == "thickness":
        return str(steps.get("thickness_raw", ""))
    if key == "copper":
        return str(steps.get("copper_spec_raw", ""))
    if key == "size":
        return f"{steps.get('size_w', '')}x{steps.get('size_h', '')}".strip("x")
    if key == "glue_category":
        return str(steps.get("glue_category", ""))
    if key == "total_core":
        return str(steps.get("order_type", ""))
    return ""


def _field_has_error(key: str, errors: list[str]) -> bool:
    text = ";".join(errors)
    return {
        "glue": "胶系" in text,
        "thickness": "厚度" in text,
        "copper": "铜箔规格" in text,
        "size": "尺寸" in text,
    }.get(key, False)


def _build_code_from_steps(steps: dict, errors: list[str]) -> str:
    if errors:
        return ""
    parts = [
        str(steps.get("step1_glue_code", "") or ""),
        str(steps.get("step2_thick_code", "") or ""),
        str(steps.get("step3_copper_code", "") or ""),
        str(steps.get("step4_size_code", "") or ""),
        str(steps.get("step5_glue_cat_code", "") or ""),
        str(steps.get("step6_copper_type_code", "") or ""),
        str(steps.get("step7_grade_code", "") or ""),
        str(steps.get("step8_tc_code", "") or ""),
        str(steps.get("step9_struct_code", "") or ""),
    ]
    if any(_is_placeholder(part) for part in parts[:-1]):
        return ""
    suffix = "XXXXXX" if str(steps.get("step5_glue_cat_code", "")) == "R" else ""
    return "".join(parts) + suffix


def _is_placeholder(value: str) -> bool:
    return "?" in str(value or "")


def _format_agent_summary(status: str, code: str, score: int, reason: str, applied_rules: list[dict]) -> str:
    applied_text = "；".join(f"{item['rule_id']}:{item['field']}->{item['new']}" for item in applied_rules) or "无"
    if status == "成功":
        return f"高置信出码：{code}；最低关键字段分={score}；Agent规则={applied_text}"
    return f"{status}：{reason}；候选码={code or '无'}；最低关键字段分={score}；Agent规则={applied_text}"


def _save_agent_result(
    source_path: str,
    output_path: str,
    df_req,
    result_col: int,
    analyses: list[dict],
    agent_rules: list[dict],
    confirm_count: int,
) -> None:
    wb = openpyxl.load_workbook(source_path)
    ws = wb["转码需求表"] if "转码需求表" in wb.sheetnames else wb[wb.sheetnames[0]]
    green_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    red_fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")

    ws.cell(row=1, column=result_col + 1, value=FORMAL_RESULT_HEADER)
    for i in range(1, len(df_req)):
        value = df_req.iloc[i, result_col]
        if pd.isna(value) or str(value).strip().lower() == "nan":
            continue
        cell = ws.cell(row=i + 1, column=result_col + 1, value=str(value))
        if str(value).startswith("待确认"):
            cell.fill = yellow_fill
        elif str(value).startswith("未识别"):
            cell.fill = red_fill
        elif str(value).startswith("跳过"):
            cell.fill = yellow_fill
        else:
            cell.fill = green_fill

    for sheet_name in ["字段证据链", "待确认清单", "规则命中汇总"]:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
    _append_evidence_sheet(wb, analyses)
    _append_confirm_sheet(wb, analyses)
    _append_summary_sheet(wb, analyses, agent_rules, confirm_count)
    wb.save(output_path)


def _append_evidence_sheet(wb, analyses: list[dict]) -> None:
    ws = wb.create_sheet("字段证据链")
    headers = ["行号", "客户代码", "客户", "规格", "状态", "字段", "原始值", "编码片段", "置信度", "命中方式", "规则来源", "证据"]
    ws.append(headers)
    for analysis in analyses:
        for item in analysis.get("field_evidence", []):
            ws.append([
                analysis.get("row"),
                analysis.get("customer_code", ""),
                analysis.get("customer", ""),
                analysis.get("spec", ""),
                analysis.get("status", ""),
                item.get("field", ""),
                item.get("value", ""),
                item.get("code", ""),
                item.get("score", ""),
                item.get("hit_type", ""),
                item.get("source", ""),
                item.get("evidence", ""),
            ])
    _format_sheet(ws, [10, 14, 18, 46, 10, 12, 18, 12, 10, 18, 24, 54])


def _append_confirm_sheet(wb, analyses: list[dict]) -> None:
    ws = wb.create_sheet("待确认清单")
    headers = ["行号", "客户代码", "客户", "规格", "状态", "正式结果", "候选编码", "最低分", "原因", "低分字段"]
    ws.append(headers)
    for analysis in analyses:
        if analysis.get("status") == "成功" or analysis.get("status") == "跳过":
            continue
        low_fields = [
            f"{item['field']}({item['score']})"
            for item in analysis.get("field_evidence", [])
            if item.get("gate") and int(item.get("score", 0)) < FIELD_GATE_THRESHOLD
        ]
        ws.append([
            analysis.get("row"),
            analysis.get("customer_code", ""),
            analysis.get("customer", ""),
            analysis.get("spec", ""),
            analysis.get("status", ""),
            analysis.get("formal_code", ""),
            analysis.get("candidate_code", ""),
            analysis.get("overall_score", ""),
            analysis.get("reason", ""),
            "；".join(low_fields),
        ])
    _format_sheet(ws, [10, 14, 18, 54, 10, 24, 24, 10, 44, 38])


def _append_summary_sheet(wb, analyses: list[dict], agent_rules: list[dict], confirm_count: int) -> None:
    ws = wb.create_sheet("规则命中汇总")
    status_counter = Counter(analysis.get("status") for analysis in analyses)
    applied_counter = Counter()
    for analysis in analyses:
        for item in analysis.get("applied_rules", []):
            applied_counter[item.get("rule_id", "")] += 1
    ws.append(["指标", "数量/说明"])
    ws.append(["总行数", len(analyses)])
    ws.append(["高置信出码", status_counter.get("成功", 0)])
    ws.append(["待确认", confirm_count])
    ws.append(["未识别", status_counter.get("失败", 0)])
    ws.append(["跳过PP/RC/%", status_counter.get("跳过", 0)])
    ws.append(["当前Agent机器规则数", len(agent_rules)])
    ws.append([])
    ws.append(["规则ID", "命中次数"])
    for rule_id, count in applied_counter.most_common():
        ws.append([rule_id, count])
    _format_sheet(ws, [24, 38])


def _format_sheet(ws, widths: list[int]) -> None:
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(index)].width = width


def _skip_analysis(row: int, customer_code: str, customer: str, spec: str, reason: str) -> dict:
    return {
        "row": row,
        "customer_code": customer_code,
        "customer": customer,
        "spec": spec,
        "status": "跳过",
        "formal_code": "",
        "candidate_code": "",
        "overall_score": 0,
        "reason": reason,
        "summary": reason,
        "field_evidence": [],
        "applied_rules": [],
        "conflicts": [],
    }


def _ensure_agent_result_column(df_req) -> int:
    for index, col in enumerate(df_req.columns):
        if str(col).strip() == FORMAL_RESULT_HEADER:
            df_req.iloc[:, index] = ""
            return index
    result_col = len(df_req.columns)
    df_req[FORMAL_RESULT_HEADER] = ""
    return result_col


def _is_effective_spec(value: Any, engine) -> bool:
    text = engine._clean_cell(value)
    if not text:
        return False
    lower = text.lower()
    return lower not in ("nan", "客户规格", "规格", "品名")


def _norm_match(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").upper())


def _norm_customer_code(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))
