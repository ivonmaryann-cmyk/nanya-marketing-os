from __future__ import annotations

import importlib
import json
import math
import re
import sys
import traceback
from collections import Counter
from functools import lru_cache
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zipfile import ZipFile

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from .db import append_job_log, create_job, get_job, prune_jobs_for_employee, update_job_status
from .excel_utils import load_workbook_compat, normalized_xlsx_source
from .transcode_evidence_scoring import (
    empty_evidence_score_shadow,
    evidence_gate_decision,
    evaluate_evidence_score_shadow,
    format_source_fields,
    get_evidence_score_runtime_mode,
    get_evidence_gate_mode,
    load_evidence_score_matrix,
)
from .transcode_evidence_model import (
    get_evidence_model_max_calls,
    load_evidence_model_runtime,
    review_evidence_shadow,
)
from .file_utils import safe_unlink
from .job_control import launch_job_process
from .paths import JOBS_DIR, PROJECT_DIR
from .transcode_agent_rules import (
    FEATURE_KEY,
    get_active_transcode_agent_rule_version,
    load_transcode_agent_mapping_tables,
    load_transcode_agent_rules,
)
from .transcode_rules import get_active_transcode_rule_version, get_transcode_rule_file_path
from .transcode_semantic_rules import (
    get_active_transcode_semantic_rule_version,
    load_transcode_semantic_rules,
)
from .transcode_semantic_shadow import (
    SHADOW_STATUS_ERROR,
    SHADOW_STATUS_MATCHED,
    SHADOW_STATUS_MISSING_INPUT,
    SHADOW_STATUS_NOT_MATCHED,
    evaluate_semantic_shadow,
    format_condition_results,
    format_observed_inputs,
    get_semantic_rule_runtime_mode,
)
from .transcode_semantic_overrides import (
    apply_confirmed_semantic_overrides,
    get_semantic_override_mode,
)
from .transcode_pending_risks import match_pending_formal_risk
from .transcode_order_semantic_model import (
    build_order_semantic_cache_key,
    load_order_semantic_runtime,
    normalize_order_shadow,
    should_normalize_order,
    source_fields_from_observations,
)


TRANSCODE_MODULE_NAME = "fangzheng_web_app.transcode_agent_engine"
FIELD_GATE_THRESHOLD = 90
FORMAL_RESULT_HEADER = "Agent转码结果"
OUTPUT_STATUS_HEADER = "结果对比"
TRANSCODE_STATUS_HEADER = "转码状态"
CONFIRMATION_HEADER = "人工确认提示"


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

OVERRIDE_FIELD_LABELS = {
    "glue_code": "胶系",
    "thickness_code": "厚度",
    "copper_code": "铜厚",
    "size_code": "尺寸",
    "glue_category_code": "胶水类别",
    "copper_type_code": "铜箔类型",
    "grade_code": "基板级别",
    "tc_code": "总/芯厚",
    "struct_code": "结构码",
}

FIELD_KEY_TO_OVERRIDE = {
    "glue": "glue_code",
    "thickness": "thickness_code",
    "copper": "copper_code",
    "size": "size_code",
    "glue_category": "glue_category_code",
    "copper_type": "copper_type_code",
    "grade": "grade_code",
    "total_core": "tc_code",
    "structure": "struct_code",
}

AGENT_EXECUTABLE_OVERRIDE_FIELDS = {
    "glue_code",
    "thickness_code",
    "copper_code",
    "size_code",
    "glue_category_code",
    "copper_type_code",
    "grade_code",
    "tc_code",
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
    semantic_version = get_active_transcode_semantic_rule_version() or "未发布"
    job_id = create_job(
        employee_id,
        source_filename,
        str(input_path),
        f"base:{base_version};agent:{agent_version};semantic:{semantic_version}",
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
    engine, tables, agent_rules, agent_mapping_tables, base_version, agent_version = _load_runtime()
    analysis = analyze_spec(
        engine,
        tables,
        agent_rules,
        spec,
        agent_mapping_tables=agent_mapping_tables,
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
        engine, tables, agent_rules, agent_mapping_tables, base_version, agent_version = _load_runtime()
        append_job_log(job_id, f"基础转码规则：{base_version}；Agent规则：{agent_version or '未上传'}")

        semantic_mode = get_semantic_rule_runtime_mode()
        semantic_override_mode = get_semantic_override_mode()
        semantic_version = get_active_transcode_semantic_rule_version()
        semantic_rules: list[dict] = []
        semantic_load_error = ""
        if semantic_mode == "shadow" and semantic_version:
            try:
                semantic_rules = load_transcode_semantic_rules(semantic_version)
            except Exception as exc:
                semantic_load_error = str(exc)
        append_job_log(
            job_id,
            "语义规则影子模式："
            f"{semantic_mode}；正式覆盖：{semantic_override_mode}；版本：{semantic_version or '未发布'}；"
            f"规则数：{len(semantic_rules)}"
            + (f"；加载失败：{semantic_load_error}" if semantic_load_error else ""),
        )
        evidence_score_mode = get_evidence_score_runtime_mode()
        evidence_gate_mode = get_evidence_gate_mode()
        evidence_score_matrix: dict[str, Any] = {}
        evidence_score_load_error = ""
        if evidence_score_mode == "shadow":
            try:
                evidence_score_matrix = load_evidence_score_matrix()
            except Exception as exc:
                evidence_score_load_error = str(exc)
        append_job_log(
            job_id,
            "证据评分影子模式："
            f"{evidence_score_mode}；正式证据门禁：{evidence_gate_mode}"
            + (f"；矩阵加载失败：{evidence_score_load_error}" if evidence_score_load_error else ""),
        )
        evidence_model_runtime = load_evidence_model_runtime()
        evidence_model_max_calls = get_evidence_model_max_calls()
        append_job_log(
            job_id,
            "模型证据审查模式："
            f"{evidence_model_runtime.mode}；模型：{evidence_model_runtime.model or '未启用'}；"
            f"单任务调用上限：{evidence_model_max_calls}"
            + (f"；加载失败：{evidence_model_runtime.load_error}" if evidence_model_runtime.load_error else ""),
        )
        order_semantic_runtime = load_order_semantic_runtime()
        append_job_log(
            job_id,
            "订单实时语义标准化："
            f"{order_semantic_runtime.mode}；模型：{order_semantic_runtime.model or '未启用'}；"
            f"单任务调用上限：{order_semantic_runtime.max_calls}"
            + (f"；加载失败：{order_semantic_runtime.load_error}" if order_semantic_runtime.load_error else ""),
        )

        workbook = load_workbook_compat(job["stored_input_path"], data_only=True)
        source_for_result = normalized_xlsx_source(job["stored_input_path"], workbook)
        sheets, _ = engine.load_transcode_inputs(str(source_for_result), str(get_transcode_rule_file_path(base_version)))
        df_req = sheets["转码需求表"].copy()
        spec_col = engine.select_transcode_spec_column(df_req)
        customer_col = engine.detect_customer_column(df_req, spec_col)
        customer_code_col = engine.detect_customer_code_column(df_req)
        context_cols = engine.detect_transcode_context_columns(df_req, spec_col, customer_col, customer_code_col)
        parse_fallback_cols = _detect_parse_fallback_columns(df_req, spec_col)
        pp_context_cols = _detect_pp_context_columns(df_req, spec_col)
        semantic_input_columns = _detect_semantic_input_columns(df_req, spec_col)
        result_col = _ensure_agent_result_column(df_req)
        data_indices = [i for i in range(1, len(df_req)) if _is_effective_spec(df_req.iloc[i, spec_col], engine)]
        total_rows = len(data_indices)
        update_job_status(job_id, status="running", total_rows=total_rows)
        append_job_log(job_id, f"识别规格列：第 {spec_col + 1} 列；Agent结果写入：第 {result_col + 1} 列")
        append_job_log(job_id, f"检测到 {total_rows} 行有效规格数据", total_rows=total_rows)

        analyses: list[dict] = []
        success_count = fail_count = skip_count = confirm_count = 0
        evidence_model_call_count = 0
        order_semantic_call_count = 0
        order_semantic_cache: dict[str, dict[str, Any]] = {}
        for processed, i in enumerate(data_indices, start=1):
            row = df_req.iloc[i]
            customer_code = engine._clean_cell(row.iloc[customer_code_col]) if customer_code_col is not None and len(row) > customer_code_col else ""
            customer = str(row.iloc[customer_col]).strip() if customer_col is not None and len(row) > customer_col and pd.notna(row.iloc[customer_col]) else ""
            spec = engine._clean_cell(row.iloc[spec_col])
            context = engine.build_context_text_from_row(row, context_cols)
            parse_fallback_text = engine.build_context_text_from_row(row, parse_fallback_cols)
            pp_context_text = engine.build_context_text_from_row(row, pp_context_cols)
            cust_spec = engine._clean_cell(row.iloc[6]) if len(row) > 6 else ""
            normalized_spec = engine._clean_cell(row.iloc[7]) if len(row) > 7 else ""
            pp_check_text = " ".join([spec, cust_spec, normalized_spec, pp_context_text])
            excel_row = i + 1

            if engine.is_pp_or_rc_spec(pp_check_text):
                skip_count += 1
                result_text = "跳过：PP/RC/% 暂不输出CCL制造编码"
                df_req.iloc[i, result_col] = result_text
                skip_analysis = _skip_analysis(excel_row, customer_code, customer, spec, result_text)
                _attach_semantic_shadow_metadata(
                    skip_analysis,
                    semantic_mode,
                    semantic_version,
                    len(semantic_rules),
                    semantic_load_error,
                    [],
                )
                skip_analysis["evidence_score_shadow"] = empty_evidence_score_shadow(
                    reason="PP/RC/%跳过，不参与CCL证据评分",
                )
                skip_analysis["order_semantic_model"] = {
                    "mode": order_semantic_runtime.mode,
                    "status": "跳过",
                    "model": order_semantic_runtime.model,
                    "reason": "PP/RC/%不属于本轮CCL语义标准化范围",
                }
                analyses.append(skip_analysis)
                append_job_log(job_id, f"第 {excel_row} 行跳过：PP/RC/%", skip_count=skip_count, current_row=processed, total_rows=total_rows)
                continue

            analysis = analyze_spec(
                engine,
                tables,
                agent_rules,
                spec,
                agent_mapping_tables=agent_mapping_tables,
                customer=customer,
                customer_code=customer_code,
                context_text=context,
                parse_fallback_text=parse_fallback_text,
                excel_row=excel_row,
            )
            shadow_results: list[dict] = []
            observations: dict[str, dict[str, Any]] = {}
            if semantic_mode == "shadow" or evidence_score_mode == "shadow":
                observations = _build_semantic_observations(
                    row,
                    semantic_input_columns,
                    analysis.get("engine_steps") or {},
                )
            if semantic_mode == "shadow" and semantic_rules and not semantic_load_error:
                try:
                    shadow_results = evaluate_semantic_shadow(
                        semantic_rules,
                        customer_code=customer_code,
                        customer_name=customer,
                        observations=observations,
                        excel_row=excel_row,
                        spec=spec,
                    )
                except Exception as exc:
                    shadow_results = [_semantic_shadow_error_result(excel_row, customer_code, customer, spec, exc)]
            if semantic_override_mode == "enforce" and shadow_results:
                semantic_applied, semantic_conflicts = apply_confirmed_semantic_overrides(
                    engine,
                    tables,
                    analysis,
                    shadow_results,
                )
                if semantic_applied or semantic_conflicts:
                    _refresh_analysis_after_semantic_overrides(
                        analysis,
                        semantic_applied,
                        semantic_conflicts,
                    )
            _attach_semantic_shadow_metadata(
                analysis,
                semantic_mode,
                semantic_version,
                len(semantic_rules),
                semantic_load_error,
                shadow_results,
            )
            semantic_source_fields = source_fields_from_observations(observations)
            analysis["order_semantic_model"] = {
                "mode": order_semantic_runtime.mode,
                "status": "未调用",
                "model": order_semantic_runtime.model,
                "source_fields": semantic_source_fields,
                "rule_ids": [
                    str(item.get("rule_id") or "")
                    for item in shadow_results
                    if item.get("rule_id")
                ],
            }
            if (
                order_semantic_runtime.mode in {"shadow", "active"}
                and order_semantic_runtime.client is not None
                and should_normalize_order(shadow_results, semantic_source_fields)
            ):
                cache_key = build_order_semantic_cache_key(
                    customer_code,
                    customer,
                    semantic_source_fields,
                    shadow_results,
                )
                cached = order_semantic_cache.get(cache_key)
                if cached is not None:
                    analysis["order_semantic_model"] = dict(cached, cached=True)
                elif order_semantic_call_count >= order_semantic_runtime.max_calls:
                    analysis["order_semantic_model"].update(
                        {"status": "限流跳过", "reason": "达到单任务DeepSeek调用上限"}
                    )
                else:
                    order_semantic_call_count += 1
                    try:
                        normalized = normalize_order_shadow(
                            order_semantic_runtime,
                            customer_code=customer_code,
                            customer_name=customer,
                            source_fields=semantic_source_fields,
                            semantic_evaluations=shadow_results,
                        )
                        model_record = {
                            "mode": order_semantic_runtime.mode,
                            "status": "成功",
                            "model": order_semantic_runtime.model,
                            "cached": False,
                            "source_fields": semantic_source_fields,
                            "rule_ids": analysis["order_semantic_model"]["rule_ids"],
                            "result": normalized,
                        }
                    except Exception as exc:
                        model_record = {
                            "mode": order_semantic_runtime.mode,
                            "status": "失败",
                            "model": order_semantic_runtime.model,
                            "cached": False,
                            "source_fields": semantic_source_fields,
                            "rule_ids": analysis["order_semantic_model"]["rule_ids"],
                            "error": str(exc),
                        }
                    order_semantic_cache[cache_key] = model_record
                    analysis["order_semantic_model"] = model_record
            pending_risk = match_pending_formal_risk(
                customer,
                spec,
                context,
                observations.get("客户规格", {}).get("value", ""),
                observations.get("客户料品名称", {}).get("value", ""),
            )
            if pending_risk and analysis.get("status") == "成功":
                analysis["status"] = "待确认"
                analysis["formal_code"] = analysis.get("candidate_code", "")
                analysis["reason"] = (
                    f"待业务确认风险[{pending_risk.get('risk_id')}]: "
                    f"{pending_risk.get('field')}；{pending_risk.get('reason')}"
                )
                analysis["summary"] = _format_agent_summary(
                    analysis["status"],
                    analysis.get("candidate_code", ""),
                    analysis.get("overall_score", 0),
                    analysis["reason"],
                    analysis.get("applied_rules") or [],
                )
                analysis["pending_formal_risk"] = pending_risk
            if evidence_score_mode == "shadow" and evidence_score_matrix and not evidence_score_load_error:
                try:
                    analysis["evidence_score_shadow"] = evaluate_evidence_score_shadow(
                        analysis,
                        semantic_evaluations=shadow_results,
                        observations=observations,
                        matrix=evidence_score_matrix,
                    )
                except Exception as exc:
                    analysis["evidence_score_shadow"] = empty_evidence_score_shadow(
                        current_score=analysis.get("overall_score", 0),
                        reason=f"证据影子评分异常，不影响正式转码：{exc}",
                    )
            else:
                analysis["evidence_score_shadow"] = empty_evidence_score_shadow(
                    current_score=analysis.get("overall_score", 0),
                    reason=evidence_score_load_error or "证据影子评分未启用",
                )
            if (
                evidence_model_runtime.mode == "shadow"
                and evidence_model_runtime.client is not None
                and evidence_score_matrix
                and analysis["evidence_score_shadow"].get("field_reviews")
                and evidence_model_call_count < evidence_model_max_calls
            ):
                analysis["evidence_score_shadow"] = review_evidence_shadow(
                    analysis,
                    semantic_evaluations=shadow_results,
                    matrix=evidence_score_matrix,
                    client=evidence_model_runtime.client,
                )
                evidence_model_call_count += int(
                    analysis["evidence_score_shadow"].get("model_call_count") or 0
                )
            gate_result = evidence_gate_decision(analysis, mode=evidence_gate_mode)
            analysis["evidence_gate"] = gate_result
            if analysis.get("status") == "成功" and gate_result["blocked"]:
                analysis["status"] = "待确认"
                analysis["formal_code"] = analysis.get("candidate_code", "")
                blocked_fields = "、".join(gate_result["blockers"]) or "证据有效分低于90"
                analysis["reason"] = f"90分证据门禁拦截：{blocked_fields}"
                analysis["summary"] = _format_agent_summary(
                    analysis["status"],
                    analysis.get("candidate_code", ""),
                    gate_result["effective_score"],
                    analysis["reason"],
                    analysis.get("applied_rules") or [],
                )
            analyses.append(analysis)
            if analysis["status"] == "成功":
                success_count += 1
                df_req.iloc[i, result_col] = analysis["formal_code"]
                log_text = f"第 {excel_row} 行高置信出码：{analysis['formal_code']}"
            elif analysis["status"] == "待确认":
                confirm_count += 1
                analysis["formal_code"] = analysis.get("formal_code") or analysis.get("candidate_code", "")
                if analysis["formal_code"]:
                    success_count += 1
                    df_req.iloc[i, result_col] = analysis["formal_code"]
                else:
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
        _save_agent_result(
            str(source_for_result),
            str(output_path),
            df_req,
            result_col,
            analyses,
            agent_rules,
            agent_mapping_tables,
            confirm_count,
        )
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
    agent_mapping_tables: dict[str, list[dict]] | None = None,
    customer: str = "",
    customer_code: str = "",
    context_text: str = "",
    parse_fallback_text: str = "",
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
    initial_errors = list((steps or {}).get("errors") or [])
    if initial_errors and str(parse_fallback_text or "").strip():
        fallback_spec = str(parse_fallback_text).strip()
        fallback_code, fallback_steps, fallback_err = engine.transcode_row(
            fallback_spec,
            "",
            str(customer or "").strip(),
            str(customer_code or "").strip(),
            runtime_tables,
            str(context_text or "").strip(),
        )
        fallback_errors = list((fallback_steps or {}).get("errors") or [])
        if len(fallback_errors) < len(initial_errors):
            code, steps, err = fallback_code, fallback_steps, fallback_err
            steps["context_fallback_used"] = True
            steps["context_fallback_note"] = "客户规格解析失败后，使用同一订单行的标准规格列补全缺失字段"
    steps = dict(steps or {})
    errors = list(steps.get("errors") or [])
    applied_rules, conflicts = _apply_agent_rules(agent_rules, customer_code, customer, spec, context_text, steps, errors)
    applied_field_mappings = _apply_agent_field_mappings(
        engine,
        agent_mapping_tables or {},
        customer_code,
        customer,
        spec,
        context_text,
        steps,
        errors,
        applied_rules,
    )
    if applied_field_mappings:
        applied_rules.extend(applied_field_mappings)
    applied_mappings = _apply_agent_size_mappings(
        engine,
        agent_mapping_tables or {},
        customer_code,
        customer,
        spec,
        context_text,
        steps,
        errors,
        applied_rules,
    )
    if applied_mappings:
        applied_rules.extend(applied_mappings)
    applied_thickness_mappings = _apply_agent_thickness_mappings(
        engine,
        agent_mapping_tables or {},
        customer_code,
        customer,
        spec,
        context_text,
        steps,
        errors,
        applied_rules,
    )
    if applied_thickness_mappings:
        applied_rules.extend(applied_thickness_mappings)
    applied_material_mappings = _apply_agent_material_code_mappings(
        agent_mapping_tables or {},
        customer_code,
        customer,
        spec,
        context_text,
        steps,
        applied_rules,
    )
    if applied_material_mappings:
        applied_rules.extend(applied_material_mappings)
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
        formal_code = candidate_code
        reason = "Agent规则冲突：" + "; ".join(conflicts)
    elif not candidate_code:
        status = "失败"
        formal_code = ""
        reason = "无法生成候选编码"
    elif low_fields:
        status = "待确认"
        formal_code = candidate_code
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
    agent_mapping_tables = load_transcode_agent_mapping_tables(agent_version) if agent_version else {}
    return engine, tables, agent_rules, agent_mapping_tables, base_version, agent_version


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
        if _should_skip_agent_rule_override(rule, field, value, customer_name, spec, context, steps):
            continue
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
                "source_row": rule.get("来源行号", ""),
                "source_field": rule.get("来源字段", ""),
                "rule_type": "确认草稿机器规则",
            }
        )
        linked_glue_category = _linked_glue_category_override(field, value)
        if linked_glue_category and "glue_category_code" not in values:
            old_category = str(steps.get("step5_glue_cat_code", "") or "")
            values["glue_category_code"] = linked_glue_category
            priorities["glue_category_code"] = _rule_priority(rule)
            steps["step5_glue_cat_code"] = linked_glue_category
            steps["glue_category"] = "普通"
            applied.append(
                {
                    "rule_id": rule.get("规则ID", ""),
                    "field": "glue_category_code",
                    "old": old_category,
                    "new": linked_glue_category,
                    "text": rule.get("规则文本", ""),
                    "source": rule.get("命中来源", "Agent规则包"),
                    "source_row": rule.get("来源行号", ""),
                    "source_field": rule.get("来源字段", ""),
                    "rule_type": "确认草稿机器规则/胶系联动",
                }
            )
    if applied:
        steps["agent_rules"] = [f"{item['rule_id']} {item['field']}:{item['old']}->{item['new']}" for item in applied]
    if conflicts:
        steps["agent_rule_conflicts"] = conflicts
    return applied, conflicts


def _should_skip_agent_rule_override(
    rule: dict,
    field: str,
    value: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
) -> bool:
    if field == "grade_code" and value == "F1" and _is_shenwan_core_ny2140(customer_name, spec, context, steps):
        return True
    return False


def _is_shenwan_core_ny2140(customer_name: str, spec: str, context: str, steps: dict) -> bool:
    combined = f"{customer_name or ''} {spec or ''} {context or ''}"
    if "深万基隆" not in combined or "芯板" not in combined:
        return False
    glue_text = " ".join(
        str(steps.get(key, "") or "")
        for key in ("glue_model", "step1_glue_code", "glue_code", "raw_glue")
    ).upper()
    return "NY2140" in combined.upper() or "NY2140" in glue_text or str(steps.get("step1_glue_code", "")).upper() == "2A"


def _linked_glue_category_override(field: str, value: str) -> str:
    if field == "glue_code" and value in {"2A", "2B", "2C"}:
        return "Y"
    return ""


def _apply_agent_field_mappings(
    engine,
    mapping_tables: dict[str, list[dict]],
    customer_code: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
    errors: list[str],
    applied_rules: list[dict],
) -> list[dict]:
    """Apply exact customer+glue mappings derived from validated correct-code samples."""
    already_overridden = {str(item.get("field") or "") for item in applied_rules}
    combined = _norm_match(f"{spec or ''} {context or ''}")
    current_glue = _norm_match(steps.get("glue_model", ""))
    applied: list[dict] = []

    for row in mapping_tables.get("客户字段映射", []):
        if not _mapping_row_enabled(row) or not _mapping_customer_matches(row, customer_code, customer_name):
            continue
        field = str(row.get("覆盖字段", "") or "").strip()
        value = str(row.get("覆盖值", "") or "").strip().upper()
        if field not in AGENT_EXECUTABLE_OVERRIDE_FIELDS or not value or field in already_overridden:
            continue
        glue_condition = _norm_match(row.get("条件胶系", ""))
        if glue_condition and glue_condition != current_glue:
            continue
        keyword_condition = str(row.get("条件关键词", "") or "").strip()
        if keyword_condition:
            keywords = [item.strip() for item in re.split(r"[/,，、;；]+", keyword_condition) if item.strip()]
            if keywords and not any(_norm_match(item) in combined for item in keywords):
                continue

        step_key, error_text = OVERRIDE_STEP_MAP.get(field, ("", ""))
        if not step_key:
            continue
        old_value = str(steps.get(step_key, "") or "")
        steps[step_key] = value
        if field == "glue_category_code":
            steps["glue_category"] = "普通" if value == "Y" else "特殊"
        elif field == "tc_code":
            steps["order_type"] = "芯厚" if value == "C" else "总厚"
        if error_text:
            errors[:] = [item for item in errors if item != error_text and not item.startswith(error_text)]
            steps["errors"] = errors
        semantic_mapping = "模型语义" in str(row.get("来源批次", "") or "")
        applied.append(
            {
                "rule_id": row.get("映射ID", ""),
                "field": field,
                "old": old_value,
                "new": value,
                "text": row.get("规则文本", ""),
                "source": "已批准模型语义映射" if semantic_mapping else "正确码回归客户字段映射",
                "source_row": row.get("来源行号", ""),
                "source_field": "客户字段映射",
                "rule_type": "已批准模型语义规则/受控映射" if semantic_mapping else "辅助客户字段映射",
            }
        )
        already_overridden.add(field)
    return applied


def _apply_agent_size_mappings(
    engine,
    mapping_tables: dict[str, list[dict]],
    customer_code: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
    errors: list[str],
    applied_rules: list[dict],
) -> list[dict]:
    if any(item.get("field") == "size_code" for item in applied_rules):
        return []
    combined_text = f"{spec or ''} {context or ''}"

    for row in mapping_tables.get("客户尺寸映射", []):
        if not _mapping_row_enabled(row) or not _mapping_customer_matches(row, customer_code, customer_name):
            continue
        customer_w = _to_float(row.get("客户尺寸W"))
        customer_h = _to_float(row.get("客户尺寸H"))
        factory_w = _to_float(row.get("厂内尺寸W"))
        factory_h = _to_float(row.get("厂内尺寸H"))
        if None in (customer_w, customer_h, factory_w, factory_h):
            continue
        if not (_steps_size_matches(steps, customer_w, customer_h) or _text_size_pair_matches(combined_text, customer_w, customer_h)):
            continue
        return [
            _apply_size_mapping_override(
                engine,
                steps,
                errors,
                row,
                factory_w,
                factory_h,
                "Agent尺寸映射/完整尺寸",
            )
        ]

    single_side_rows = [
        row for row in mapping_tables.get("客户单边尺寸映射", [])
        if _mapping_row_enabled(row) and _mapping_customer_matches(row, customer_code, customer_name)
    ]
    if single_side_rows:
        side_map = {
            _norm_size_token(row.get("客户单边尺寸")): _to_float(row.get("厂内单边尺寸"))
            for row in single_side_rows
            if _norm_size_token(row.get("客户单边尺寸")) and _to_float(row.get("厂内单边尺寸")) is not None
        }
        raw_pair = _extract_raw_size_pair(combined_text)
        if raw_pair:
            left_raw, right_raw = raw_pair
            left_mapped = side_map.get(_norm_size_token(left_raw))
            right_mapped = side_map.get(_norm_size_token(right_raw))
            if left_mapped is not None or right_mapped is not None:
                factory_w = left_mapped if left_mapped is not None else _to_float(steps.get("size_w"))
                factory_h = right_mapped if right_mapped is not None else _to_float(steps.get("size_h"))
                if factory_w is not None and factory_h is not None:
                    source_row = next(
                        row for row in single_side_rows
                        if _norm_size_token(row.get("客户单边尺寸")) in {_norm_size_token(left_raw), _norm_size_token(right_raw)}
                    )
                    return [
                        _apply_size_mapping_override(
                            engine,
                            steps,
                            errors,
                            source_row,
                            factory_w,
                            factory_h,
                            "Agent尺寸映射/单边尺寸",
                        )
                    ]

    for row in mapping_tables.get("客户尺寸算法", []):
        if not _mapping_row_enabled(row) or not _mapping_customer_matches(row, customer_code, customer_name):
            continue
        if row.get("算法类型") != "尺寸加大":
            continue
        current_w = _to_float(steps.get("size_w"))
        current_h = _to_float(steps.get("size_h"))
        delta_w = _to_float(row.get("加大W"))
        delta_h = _to_float(row.get("加大H"))
        if None in (current_w, current_h, delta_w, delta_h):
            continue
        # A size already ending in the configured enlargement must not be
        # enlarged a second time (e.g. 41.3*49.3 with a +0.3 rule).
        if _size_already_enlarged(current_w, delta_w) and _size_already_enlarged(current_h, delta_h):
            continue
        return [
            _apply_size_mapping_override(
                engine,
                steps,
                errors,
                row,
                current_w + delta_w,
                current_h + delta_h,
                "Agent尺寸映射/尺寸算法",
            )
        ]

    for row in mapping_tables.get("外部尺寸表引用", []):
        if not _mapping_row_enabled(row) or not _mapping_customer_matches(row, customer_code, customer_name):
            continue
        applied = _apply_external_size_reference(engine, row, combined_text, steps, errors)
        if applied:
            return [applied]
    return []


def _apply_size_mapping_override(engine, steps: dict, errors: list[str], row: dict, width: float, height: float, source: str) -> dict:
    old_value = str(steps.get("step4_size_code", "") or "")
    new_value = engine.size_to_code(width, height)
    steps["size_w"] = round(float(width), 2)
    steps["size_h"] = round(float(height), 2)
    steps["step4_size_code"] = new_value
    steps["size_note"] = f"{source}：{_format_size(width)}x{_format_size(height)} -> {new_value}"
    errors[:] = [item for item in errors if item != "无法识别尺寸" and not item.startswith("无法识别尺寸")]
    steps["errors"] = errors
    return {
        "rule_id": row.get("映射ID", ""),
        "field": "size_code",
        "old": old_value,
        "new": new_value,
        "text": row.get("规则文本", ""),
        "source": source,
        "source_row": row.get("来源行号", ""),
        "source_field": row.get("来源字段", ""),
        "rule_type": "辅助尺寸映射",
    }


def _apply_external_size_reference(engine, row: dict, combined_text: str, steps: dict, errors: list[str]) -> dict | None:
    if "新美亚" not in str(row.get("规则文本", "") + row.get("引用文件", "")):
        return None
    raw_pair = _extract_raw_size_pair(combined_text)
    if not raw_pair:
        return None
    raw_w = _to_float(raw_pair[0])
    raw_h = _to_float(raw_pair[1])
    if raw_w is None or raw_h is None:
        return None

    if _material_code_matches(combined_text, "631"):
        converted = _lookup_xinmeiya_631_size(row, raw_w, raw_h)
        source_label = "Agent尺寸映射/外部尺寸表-新美亚631"
    elif _material_code_matches(combined_text, "632"):
        converted = _convert_xinmeiya_632_size(raw_w, raw_h)
        source_label = "Agent尺寸映射/外部尺寸表-新美亚632"
    else:
        return None
    if not converted:
        return None

    mm_w, mm_h = converted
    width = mm_w / 25.4
    height = mm_h / 25.4
    applied = _apply_size_mapping_override(engine, steps, errors, row, width, height, source_label)
    applied["text"] = f"{raw_pair[0]}*{raw_pair[1]} -> {int(mm_w)}*{int(mm_h)}mm -> {_format_size(width)}*{_format_size(height)}"
    applied["rule_type"] = "辅助外部尺寸表映射"
    steps["size_note"] = f"{source_label}：{applied['text']} -> {applied['new']}"
    return applied


def _lookup_xinmeiya_631_size(row: dict, width: float, height: float) -> tuple[float, float] | None:
    reference_path = _resolve_external_reference_path(row.get("引用文件", ""))
    if not reference_path:
        return None
    for entry in _load_xinmeiya_size_table(str(reference_path)):
        if _near_size(entry["inch_w"], width) and _near_size(entry["inch_h"], height):
            return entry["mm_w"], entry["mm_h"]
    return None


def _convert_xinmeiya_632_size(width: float, height: float) -> tuple[float, float]:
    return _round_half_up(width * 25.4), _round_half_up(height * 25.4)


def _round_half_up(value: float) -> int:
    return int(math.floor(float(value) + 0.5))


def _resolve_external_reference_path(value: Any) -> Path | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    path = Path(raw)
    if not path.is_absolute():
        path = PROJECT_DIR / path
    return path if path.exists() else None


@lru_cache(maxsize=8)
def _load_xinmeiya_size_table(path: str) -> tuple[dict[str, float], ...]:
    workbook_rows = _read_xlsx_first_sheet_rows(Path(path))
    entries: list[dict[str, float]] = []
    in_631_section = False
    for row in workbook_rows:
        row_text = " ".join(str(value or "") for value in row)
        if "631." in row_text and "Thin Core" in row_text:
            in_631_section = True
            continue
        if in_631_section and ("620." in row_text or "Double Side" in row_text or "632." in row_text):
            break
        if not in_631_section or len(row) < 5:
            continue
        inch_w = _to_float(row[1])
        inch_h = _to_float(row[2])
        mm_w = _to_float(row[3])
        mm_h = _to_float(row[4])
        if None in (inch_w, inch_h, mm_w, mm_h):
            continue
        entries.append({"inch_w": inch_w, "inch_h": inch_h, "mm_w": mm_w, "mm_h": mm_h})
    return tuple(entries)


def _read_xlsx_first_sheet_rows(path: Path) -> list[list[str]]:
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(path) as archive:
        shared_strings = _read_xlsx_shared_strings(archive, namespace)
        sheet_xml = _read_xlsx_member(archive, "xl/worksheets/sheet1.xml")
        root = ElementTree.fromstring(sheet_xml)
        rows: list[list[str]] = []
        for row in root.findall(".//x:row", namespace):
            values: dict[int, str] = {}
            for cell in row.findall("x:c", namespace):
                ref = cell.attrib.get("r", "")
                value_node = cell.find("x:v", namespace)
                value = ""
                if value_node is not None:
                    raw_value = value_node.text or ""
                    if cell.attrib.get("t") == "s" and raw_value.isdigit():
                        value = shared_strings[int(raw_value)] if int(raw_value) < len(shared_strings) else ""
                    else:
                        value = raw_value
                values[_column_index_from_cell_ref(ref)] = value
            if values:
                rows.append([values.get(idx, "") for idx in range(max(values) + 1)])
    return rows


def _read_xlsx_shared_strings(archive: ZipFile, namespace: dict[str, str]) -> list[str]:
    try:
        root = ElementTree.fromstring(_read_xlsx_member(archive, "xl/sharedStrings.xml"))
    except KeyError:
        return []
    values: list[str] = []
    for item in root.findall("x:si", namespace):
        values.append("".join(text_node.text or "" for text_node in item.findall(".//x:t", namespace)))
    return values


def _read_xlsx_member(archive: ZipFile, normalized_name: str) -> bytes:
    by_normalized_name = {name.replace("\\", "/"): name for name in archive.namelist()}
    actual_name = by_normalized_name.get(normalized_name)
    if not actual_name:
        raise KeyError(normalized_name)
    return archive.read(actual_name)


def _column_index_from_cell_ref(cell_ref: str) -> int:
    letters = "".join(ch for ch in str(cell_ref or "") if ch.isalpha())
    value = 0
    for letter in letters:
        value = value * 26 + ord(letter.upper()) - ord("A") + 1
    return max(value - 1, 0)


def _apply_agent_thickness_mappings(
    engine,
    mapping_tables: dict[str, list[dict]],
    customer_code: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
    errors: list[str],
    applied_rules: list[dict],
) -> list[dict]:
    combined_text = f"{spec or ''} {context or ''}"
    applied: list[dict] = []
    has_thickness_override = any(item.get("field") == "thickness_code" for item in applied_rules)
    has_tc_override = any(item.get("field") == "tc_code" for item in applied_rules)

    for row in mapping_tables.get("客户厚度映射", []):
        if not _mapping_row_enabled(row) or not _mapping_customer_matches(row, customer_code, customer_name):
            continue
        alias = str(row.get("客户厚度写法", "") or "").strip()
        mm_value = _to_float(row.get("厚度mm"))
        mil_value = _to_float(row.get("厚度mil"))
        tc_code = str(row.get("总芯厚口径", "") or "").strip().upper()

        if mm_value is not None and alias and not has_thickness_override and _thickness_alias_matches(combined_text, alias):
            applied.append(_apply_thickness_mapping_override(engine, steps, errors, row, mm_value, mil_value))
            has_thickness_override = True
            if mil_value is not None and mil_value < 31 and not has_tc_override:
                core_row = dict(row)
                core_row["规则文本"] = f"{row.get('规则文本', '')}；{_format_size(mil_value)}mil<31mil按芯厚"
                applied.append(_apply_tc_mapping_override(steps, core_row, "C"))
                has_tc_override = True

        if tc_code in {"T", "C"} and not has_tc_override and _thickness_tc_condition_matches(combined_text, steps, mil_value, alias):
            applied.append(_apply_tc_mapping_override(steps, row, tc_code))
            has_tc_override = True
    return applied


def _apply_agent_material_code_mappings(
    mapping_tables: dict[str, list[dict]],
    customer_code: str,
    customer_name: str,
    spec: str,
    context: str,
    steps: dict,
    applied_rules: list[dict],
) -> list[dict]:
    if any(item.get("field") == "tc_code" for item in applied_rules):
        return []
    combined_text = f"{spec or ''} {context or ''}"
    for row in mapping_tables.get("客户物料编码口径", []):
        if not _mapping_row_enabled(row) or not _mapping_customer_matches(row, customer_code, customer_name):
            continue
        tc_code = str(row.get("总芯厚口径", "") or "").strip().upper()
        match_value = str(row.get("命中值", "") or "").strip()
        if tc_code not in {"T", "C"} or not match_value:
            continue
        if not _material_code_matches(combined_text, match_value):
            continue
        return [
            _apply_tc_mapping_override(
                steps,
                row,
                tc_code,
                source="Agent物料编码口径",
                rule_type="辅助物料编码口径",
            )
        ]
    return []


def _apply_thickness_mapping_override(engine, steps: dict, errors: list[str], row: dict, mm_value: float, mil_value: float | None) -> dict:
    old_value = str(steps.get("step2_thick_code", "") or "")
    new_value = engine.thickness_to_code(mm_value)
    steps["thickness_raw"] = row.get("客户厚度写法", "")
    steps["thickness_mm"] = mm_value
    steps["thickness_unit"] = "Agent客户厚度映射"
    steps["thickness_mode_source"] = "Agent客户厚度映射"
    note_parts = [str(row.get("规则文本", "") or "").strip()]
    if mil_value is not None:
        note_parts.append(f"{_format_size(mil_value)}mil")
    note_parts.append(f"{_format_size(mm_value)}mm")
    steps["thickness_mode_note"] = " -> ".join(item for item in note_parts if item)
    steps["order_mm"] = mm_value
    steps["step2_thick_code"] = new_value
    errors[:] = [item for item in errors if item != "无法识别厚度" and not item.startswith("无法识别厚度")]
    steps["errors"] = errors
    return {
        "rule_id": row.get("映射ID", ""),
        "field": "thickness_code",
        "old": old_value,
        "new": new_value,
        "text": row.get("规则文本", ""),
        "source": "Agent厚度映射",
        "source_row": row.get("来源行号", ""),
        "source_field": row.get("来源字段", ""),
        "rule_type": "辅助厚度映射",
    }


def _apply_tc_mapping_override(
    steps: dict,
    row: dict,
    tc_code: str,
    *,
    source: str = "Agent总芯厚映射",
    rule_type: str = "辅助厚度映射",
) -> dict:
    old_value = str(steps.get("step8_tc_code", "") or "")
    steps["step8_tc_code"] = tc_code
    steps["order_type"] = "总厚" if tc_code == "T" else "芯厚"
    steps["thickness_mode"] = steps["order_type"]
    steps["thickness_mode_source"] = "Agent客户厚度映射"
    if row.get("规则文本"):
        steps["thickness_mode_note"] = row.get("规则文本", "")
    return {
        "rule_id": row.get("映射ID", ""),
        "field": "tc_code",
        "old": old_value,
        "new": tc_code,
        "text": row.get("规则文本", ""),
        "source": source,
        "source_row": row.get("来源行号", ""),
        "source_field": row.get("来源字段", ""),
        "rule_type": rule_type,
    }


def _thickness_alias_matches(text: str, alias: str) -> bool:
    raw_alias = str(alias or "").strip()
    if not raw_alias:
        return False
    source = str(text or "")
    mil_alias = re.fullmatch(r"(\d+(?:\.\d+)?)\s*mil", raw_alias, flags=re.IGNORECASE)
    if mil_alias:
        value = re.escape(mil_alias.group(1))
        return bool(re.search(
            rf"(?<![\d.]){value}\s*(?:±\s*\d+(?:\.\d+)?\s*)?mil(?![A-Z0-9])",
            source,
            flags=re.IGNORECASE,
        ))
    if re.fullmatch(r"\d+(?:\.\d+)?", raw_alias):
        return bool(re.search(rf"(?<!\d){re.escape(raw_alias)}(?!\d)", source))
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(raw_alias)}(?![A-Z0-9])", source, flags=re.IGNORECASE))


def _size_already_enlarged(value: float, delta: float) -> bool:
    if delta <= 0:
        return False
    fractional = round(float(value) - int(float(value)), 2)
    return abs(fractional - round(float(delta), 2)) <= 0.01


def _material_code_matches(text: str, value: str) -> bool:
    raw_value = str(value or "").strip()
    if not raw_value:
        return False
    source = str(text or "")
    if re.fullmatch(r"\d+", raw_value):
        return bool(re.search(rf"(?<!\d){re.escape(raw_value)}(?!\d)", source))
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(raw_value)}(?![A-Z0-9])", source, flags=re.IGNORECASE))


def _thickness_tc_condition_matches(text: str, steps: dict, mil_value: float | None, alias: str) -> bool:
    actual_mil = _extract_mil_from_text(text)
    if actual_mil is None:
        current_mm = _to_float(steps.get("thickness_mm"))
        actual_mil = current_mm / 0.0254 if current_mm is not None else None
    if mil_value is not None and "含以上" in str(alias):
        return actual_mil is not None and actual_mil >= mil_value - 1e-9
    if mil_value is not None:
        return actual_mil is not None and abs(actual_mil - mil_value) <= 0.01
    return _thickness_alias_matches(text, alias)


def _extract_mil_from_text(text: str) -> float | None:
    matches = re.findall(r"(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*MIL(?![A-Z0-9])", str(text or ""), flags=re.IGNORECASE)
    if not matches:
        return None
    try:
        return float(matches[-1])
    except ValueError:
        return None


def _mapping_row_enabled(row: dict) -> bool:
    value = str(row.get("启用", "")).strip().upper()
    return value in {"是", "Y", "YES", "TRUE", "1"}


def _mapping_customer_matches(row: dict, customer_code: str, customer_name: str) -> bool:
    rule = {"客户代码": row.get("客户代码", ""), "客户简称": row.get("客户简称", "")}
    return _customer_matches(rule, customer_code, customer_name)


def _steps_size_matches(steps: dict, width: float, height: float) -> bool:
    current_w = _to_float(steps.get("size_w"))
    current_h = _to_float(steps.get("size_h"))
    return current_w is not None and current_h is not None and _near_size(current_w, width) and _near_size(current_h, height)


def _text_size_pair_matches(text: str, width: float, height: float) -> bool:
    for left, right in _extract_all_raw_size_pairs(text):
        left_num = _to_float(left)
        right_num = _to_float(right)
        if left_num is not None and right_num is not None and _near_size(left_num, width) and _near_size(right_num, height):
            return True
    return False


def _extract_raw_size_pair(text: str) -> tuple[str, str] | None:
    pairs = _extract_all_raw_size_pairs(text)
    return pairs[-1] if pairs else None


def _extract_all_raw_size_pairs(text: str) -> list[tuple[str, str]]:
    source = str(text or "").replace("×", "*").replace("x", "*").replace("X", "*")
    explicit = re.findall(r"(?<!\d)(\d{2,4}(?:\.\d+)?)\s*\*\s*(\d{2,4}(?:\.\d+)?)(?!\d)", source)
    loose = re.findall(r"(?<!\d)(\d{2,4}(?:\.\d+)?)\s+(\d{2,4}(?:\.\d+)?)(?!\d)", source)
    loose = [
        pair for pair in loose
        if pair not in explicit
        and not (len(pair[0].split(".")[0]) == 4 and len(pair[1].split(".")[0]) == 4)
    ]
    return explicit + loose


def _norm_size_token(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    if re.fullmatch(r"\d+(?:\.0+)?", raw):
        return str(int(float(raw)))
    return raw


def _to_float(value: Any) -> float | None:
    try:
        raw = str(value or "").strip()
        if not raw:
            return None
        return float(raw)
    except (TypeError, ValueError):
        return None


def _near_size(left: float, right: float) -> bool:
    return abs(float(left) - float(right)) <= 0.01


def _format_size(value: float) -> str:
    return f"{float(value):.2f}".rstrip("0").rstrip(".")


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
        and rule.get("覆盖字段") in AGENT_EXECUTABLE_OVERRIDE_FIELDS
        and bool(str(rule.get("覆盖值", "") or "").strip())
    )


def _rule_matches(rule: dict, customer_code: str, customer_name: str, spec: str, context: str, steps: dict) -> bool:
    if not _customer_matches(rule, customer_code, customer_name):
        return False
    combined = _norm_match(f"{spec} {context}")
    spec_norm = _norm_match(spec)
    glue_norm = _norm_match(steps.get("glue_model", ""))
    copper_norm = _norm_match(steps.get("copper_spec_raw", ""))
    thickness_norm = _norm_match(steps.get("thickness_raw", ""))
    size_norm = _norm_match(f"{steps.get('size_w', '')}*{steps.get('size_h', '')}")
    raw_condition_text = str(rule.get("条件文本", "") or rule.get("规则文本", "") or "")
    has_positive_tolerance = bool(re.search(r"\+\s*\d|±|偏正", str(spec or "") + " " + str(context or "")))
    if "有偏正公差" in raw_condition_text and not has_positive_tolerance:
        return False
    if "无偏正公差" in raw_condition_text and has_positive_tolerance:
        return False

    glue_condition = _norm_match(rule.get("条件胶系", ""))
    if glue_condition:
        glue_tokens = [item for item in re.split(r"[/,，;；]+", glue_condition) if item]
        if glue_tokens and not any(_glue_condition_matches(token, glue_norm, spec_norm) for token in glue_tokens):
            return False

    copper_condition = _norm_match(rule.get("条件铜厚", ""))
    if copper_condition:
        copper_tokens = _copper_condition_tokens(copper_condition)
        if copper_tokens and not any(_copper_condition_matches(token, copper_norm, spec_norm) for token in copper_tokens):
            return False

    thickness_condition = str(rule.get("条件厚度", "") or "").strip()
    if thickness_condition:
        thickness_tokens = [item for item in re.split(r"[/,，;；]+", thickness_condition) if item.strip()]
        if thickness_tokens and not any(_thickness_condition_matches(token, steps, spec_norm, thickness_norm) for token in thickness_tokens):
            return False

    size_condition = _norm_match(rule.get("条件尺寸", ""))
    if size_condition:
        size_tokens = [item for item in re.split(r"[/,，;；]+", size_condition) if item]
        if size_tokens and not any(token in size_norm or token in spec_norm or token in combined for token in size_tokens):
            return False

    keyword_condition = str(rule.get("条件关键词", "") or "").strip()
    if keyword_condition:
        if _is_legacy_copper_ff_keyword_condition(rule, keyword_condition):
            if not _legacy_copper_ff_keyword_matches(combined, steps):
                return False
        else:
            keywords = [item for item in re.split(r"[/,，、;；]+", keyword_condition) if item.strip()]
            if keywords and not any(_keyword_condition_matches(keyword, combined) for keyword in keywords):
                return False

    condition_text = _norm_match(rule.get("条件文本", ""))
    if condition_text and not (glue_condition or copper_condition or thickness_condition or size_condition or keyword_condition):
        field_label = _norm_match(rule.get("原始字段", ""))
        if field_label and field_label in condition_text:
            return True
        relaxed = re.sub(r"^(当|如果)|时$", "", condition_text)
        if relaxed and relaxed not in combined:
            return False
    return True


def _customer_matches(rule: dict, customer_code: str, customer_name: str) -> bool:
    rule_codes = set(_customer_code_tokens(rule.get("客户代码", "")))
    current_codes = set(_customer_code_tokens(customer_code))
    if rule_codes:
        if not current_codes or rule_codes.isdisjoint(current_codes):
            return False
    rule_name = _norm_match(rule.get("客户简称", ""))
    current_name = _norm_match(customer_name)
    if rule_name:
        if not current_name or (rule_name not in current_name and current_name not in rule_name):
            return False
    return bool(rule_codes or rule_name)


def _build_field_evidence(steps: dict, errors: list[str], applied_rules: list[dict], conflicts: list[str]) -> list[dict]:
    by_field = {item["field"]: item for item in applied_rules}
    evidence = []
    for key, label, step_key, gate in FIELD_DEFS:
        code_value = str(steps.get(step_key, "") or "")
        override_field = FIELD_KEY_TO_OVERRIDE[key]
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
                "rule_id": by_field.get(override_field, {}).get("rule_id", ""),
                "rule_type": by_field.get(override_field, {}).get("rule_type", ""),
                "source_row": by_field.get(override_field, {}).get("source_row", ""),
            }
        )
    return evidence


def _score_field(key: str, code_value: str, steps: dict, errors: list[str], by_field: dict, conflicts: list[str]) -> tuple[int, str, str, str]:
    override_field = FIELD_KEY_TO_OVERRIDE[key]
    if override_field in by_field:
        rule = by_field[override_field]
        if str(rule.get("source", "")).startswith(("已批准语义规则", "已批准模型语义映射")):
            return 98, "已批准模型语义映射", rule["rule_id"], rule["text"]
        if str(rule.get("source", "")).startswith("Agent尺寸映射"):
            return 98, "Agent尺寸映射", rule["rule_id"], rule["text"]
        if str(rule.get("source", "")).startswith("Agent厚度映射"):
            return 98, "Agent厚度映射", rule["rule_id"], rule["text"]
        if str(rule.get("source", "")).startswith("Agent总芯厚映射"):
            return 98, "Agent总芯厚映射", rule["rule_id"], rule["text"]
        if str(rule.get("source", "")).startswith("Agent物料编码口径"):
            return 98, "Agent物料编码口径", rule["rule_id"], rule["text"]
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
            return 90, "通用阈值", source, str(steps.get("thickness_raw", ""))
        return 94, "标准解析", source, str(steps.get("thickness_raw", ""))
    if key == "size":
        note = str(steps.get("size_note", ""))
        return (96 if "特殊" in note else 94), "特殊尺寸" if "特殊" in note else "标准尺寸", note or "基础解析", f"{steps.get('size_w', '')}x{steps.get('size_h', '')}"
    if key == "grade":
        if steps.get("grade_note"):
            return 97, "客户下单转换表", "客户下单与胶系基板转换", str(steps.get("grade_note", ""))
        if code_value == "A1":
            return 90, "默认规则", "基础规则", "未命中特殊基板等级，按已确认默认A1"
        return 94, "标准/关键词规则", "基础规则", code_value
    if key == "total_core":
        source = str(steps.get("thickness_mode_source", "") or "基础规则")
        return (90 if "通用阈值" in source else 94), "总芯厚判断", source, str(steps.get("order_type", ""))
    if key == "copper_type":
        if code_value == "W":
            return 90, "默认常规铜", "基础规则", "按已确认常规铜HTE/W默认"
        return 94, "关键词命中", "基础规则", code_value
    if key == "structure" and code_value == "*":
        return 80, "阶段性占位", "首版策略", "结构码首版允许*占位，不参与正式出码拦截"
    return 94, "标准解析", "基础规则", code_value


def _refresh_analysis_after_semantic_overrides(
    analysis: dict[str, Any],
    semantic_applied: list[dict[str, Any]],
    semantic_conflicts: list[str],
) -> None:
    steps = analysis.get("engine_steps") or {}
    errors = list(steps.get("errors") or [])
    applied_rules = list(analysis.get("applied_rules") or []) + semantic_applied
    conflicts = list(analysis.get("conflicts") or []) + semantic_conflicts
    candidate_code = _build_code_from_steps(steps, errors)
    field_evidence = _build_field_evidence(steps, errors, applied_rules, conflicts)
    gate_scores = [item["score"] for item in field_evidence if item["gate"]]
    overall_score = min(gate_scores) if gate_scores else 0
    low_fields = [
        item["field"]
        for item in field_evidence
        if item["gate"] and item["score"] < FIELD_GATE_THRESHOLD
    ]
    if errors:
        status, formal_code, reason = "失败", "", "; ".join(errors)
    elif conflicts:
        status, formal_code, reason = "待确认", candidate_code, "语义规则冲突：" + "; ".join(conflicts)
    elif not candidate_code:
        status, formal_code, reason = "失败", "", "无法生成候选编码"
    elif low_fields:
        status, formal_code, reason = "待确认", candidate_code, f"低置信字段：{', '.join(low_fields)}"
    else:
        status, formal_code, reason = "成功", candidate_code, ""
    analysis.update(
        {
            "status": status,
            "formal_code": formal_code,
            "candidate_code": candidate_code,
            "overall_score": overall_score,
            "reason": reason,
            "field_evidence": field_evidence,
            "applied_rules": applied_rules,
            "conflicts": conflicts,
            "summary": _format_agent_summary(status, formal_code or candidate_code, overall_score, reason, applied_rules),
        }
    )


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


def _detect_semantic_input_columns(df_req, spec_col: int) -> dict[str, dict[str, Any]]:
    headers = [str(value or "").strip() for value in df_req.iloc[0].tolist()]
    normalized = [_normalize_semantic_header(value) for value in headers]
    aliases = {
        "订单备注": {"订单备注", "备注", "订单说明", "订单附注"},
        "客户规格": {"客户规格"},
        "客户料品名称": {"客户料品名称", "客户料品", "料品名称", "品名"},
        "客户物料编码": {"客户物料编码", "客户产品编号", "客户料号"},
        "品号/物料编号": {"品号", "物料编号", "料号"},
    }
    result: dict[str, dict[str, Any]] = {}
    for field, names in aliases.items():
        normalized_names = {_normalize_semantic_header(name) for name in names}
        indices = [index for index, name in enumerate(normalized) if name in normalized_names]
        result[field] = {"indices": indices, "headers": [headers[index] for index in indices]}
    result["订单规格"] = {
        "indices": [spec_col],
        "headers": [headers[spec_col] if spec_col < len(headers) else "规格"],
    }
    return result


def _detect_parse_fallback_columns(df_req, spec_col: int) -> list[int]:
    """Find normalized spec columns used only after customer-spec parsing fails.

    A plain 品名 column is excluded because historical test files use it for the
    correct manufacturing code, which is not a parse input.
    """
    if df_req.empty:
        return []
    result: list[int] = []
    for row_idx in range(min(3, len(df_req))):
        for col_idx in range(len(df_req.columns)):
            if col_idx == spec_col:
                continue
            header = _normalize_semantic_header(df_req.iloc[row_idx, col_idx])
            if header in {"品名规格", "规格", "标准规格"} and col_idx not in result:
                result.append(col_idx)
    return result


def _detect_pp_context_columns(df_req, spec_col: int) -> list[int]:
    """Find normalized specification columns suitable for PP/RC classification.

    Correct-code/product-code columns are intentionally excluded because values such as
    RC can also be valid CCL glue codes and must not cause a false PP skip.
    """
    if df_req.empty:
        return []
    result: list[int] = []
    for row_idx in range(min(3, len(df_req))):
        for col_idx in range(len(df_req.columns)):
            if col_idx == spec_col:
                continue
            header = _normalize_semantic_header(df_req.iloc[row_idx, col_idx])
            if header in {"规格", "标准规格"} and col_idx not in result:
                result.append(col_idx)
    return result


def _build_semantic_observations(
    row,
    column_map: dict[str, dict[str, Any]],
    engine_steps: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    observations = {
        field: _column_observation(row, metadata)
        for field, metadata in column_map.items()
    }
    observations["胶系"] = _derived_observation(
        engine_steps.get("glue_model"),
        "规格解析/胶系",
    )
    observations["基板厚度"] = _derived_observation(
        engine_steps.get("thickness_mm"),
        "规格解析/基板厚度",
    )
    observations["铜箔规格"] = _derived_observation(
        engine_steps.get("copper_spec_raw"),
        "规格解析/铜箔规格",
    )
    order_spec = observations.get("订单规格") or {"available": False, "value": "", "sources": []}
    order_remark = observations.get("订单备注") or {"available": False, "value": "", "sources": []}
    combined_values = [
        str(item.get("value") or "").strip()
        for item in [order_spec, order_remark]
        if item.get("available") and str(item.get("value") or "").strip()
    ]
    combined_sources = [
        source
        for item in [order_spec, order_remark]
        for source in item.get("sources") or []
    ]
    observations["订单规格/订单备注"] = {
        "available": bool(order_spec.get("available") or order_remark.get("available")),
        "value": " ".join(dict.fromkeys(combined_values)),
        "sources": list(dict.fromkeys(combined_sources)),
    }
    return observations


def _column_observation(row, metadata: dict[str, Any]) -> dict[str, Any]:
    indices = list(metadata.get("indices") or [])
    values: list[str] = []
    for index in indices:
        if index >= len(row):
            continue
        value = row.iloc[index]
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text and text.lower() != "nan":
            values.append(text)
    return {
        "available": bool(indices),
        "value": " ".join(dict.fromkeys(values)),
        "sources": list(metadata.get("headers") or []),
    }


def _derived_observation(value: Any, source: str) -> dict[str, Any]:
    available = value is not None and str(value).strip() not in {"", "nan", "None"}
    return {
        "available": available,
        "value": value if available else "",
        "sources": [source] if available else [],
    }


def _normalize_semantic_header(value: Any) -> str:
    return re.sub(r"[\s　]+", "", str(value or "")).strip().lower()


def _attach_semantic_shadow_metadata(
    analysis: dict,
    mode: str,
    version: str,
    rule_count: int,
    load_error: str,
    evaluations: list[dict],
) -> None:
    analysis["semantic_shadow_mode"] = mode
    analysis["semantic_rule_version"] = version
    analysis["semantic_rule_count"] = rule_count
    analysis["semantic_load_error"] = load_error
    analysis["semantic_shadow"] = evaluations


def _semantic_shadow_error_result(
    excel_row: int,
    customer_code: str,
    customer: str,
    spec: str,
    error: Exception,
) -> dict[str, Any]:
    return {
        "row": excel_row,
        "customer_code": customer_code,
        "customer": customer,
        "spec": spec,
        "rule_id": "",
        "source_candidate_id": "",
        "business_field": "",
        "target_fields": [],
        "normalized_values": [],
        "stated_target_values": [],
        "status": SHADOW_STATUS_ERROR,
        "missing_fields": [],
        "condition_results": [],
        "observed_inputs": {},
        "source_text": "",
        "evidence_texts": [],
        "model": "",
        "note": f"影子评估异常，不影响正式转码：{error}",
    }


def _save_agent_result(
    source_path: str,
    output_path: str,
    df_req,
    result_col: int,
    analyses: list[dict],
    agent_rules: list[dict],
    agent_mapping_tables: dict[str, list[dict]],
    confirm_count: int,
) -> None:
    wb = openpyxl.load_workbook(source_path)
    ws = wb["转码需求表"] if "转码需求表" in wb.sheetnames else wb[wb.sheetnames[0]]
    green_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    red_fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")

    ws.cell(row=1, column=result_col + 1, value=FORMAL_RESULT_HEADER)
    ws.cell(row=1, column=result_col + 2, value=OUTPUT_STATUS_HEADER)
    ws.cell(row=1, column=result_col + 3, value=TRANSCODE_STATUS_HEADER)
    ws.cell(row=1, column=result_col + 4, value=CONFIRMATION_HEADER)
    product_name_col = _find_product_name_column(ws)
    analyses_by_row = {
        int(analysis["row"]): analysis
        for analysis in analyses
        if analysis.get("row") is not None
    }
    for i in range(1, len(df_req)):
        value = df_req.iloc[i, result_col]
        analysis = analyses_by_row.get(i + 1, {})
        ws.cell(
            row=i + 1,
            column=result_col + 2,
            value=_export_result_comparison(
                analysis,
                ws.cell(row=i + 1, column=product_name_col).value if product_name_col else "",
            ),
        )
        ws.cell(
            row=i + 1,
            column=result_col + 3,
            value=_export_transcode_status(analysis, value),
        )
        if analysis.get("status") == "待确认":
            ws.cell(
                row=i + 1,
                column=result_col + 4,
                value=f"待确认：{analysis.get('reason', '')}",
            ).fill = red_fill
        if pd.isna(value) or str(value).strip().lower() == "nan":
            continue
        cell = ws.cell(row=i + 1, column=result_col + 1, value=str(value))
        if analysis.get("status") == "待确认":
            cell.fill = red_fill
        elif str(value).startswith("未识别"):
            cell.fill = red_fill
        elif str(value).startswith("跳过"):
            cell.fill = yellow_fill
        else:
            cell.fill = green_fill

    for sheet_name in [
        "字段证据链",
        "待确认清单",
        "规则命中汇总",
        "技术待支持清单",
        "模型语义影子证据",
        "模型实时语义标准化",
        "证据评分影子对比",
    ]:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
    _append_evidence_sheet(wb, analyses)
    _append_confirm_sheet(wb, analyses)
    _append_summary_sheet(wb, analyses, agent_rules, agent_mapping_tables, confirm_count)
    _append_technical_pending_sheet(wb, agent_mapping_tables)
    _append_semantic_shadow_sheet(wb, analyses)
    _append_order_semantic_model_sheet(wb, analyses)
    _append_evidence_score_shadow_sheet(wb, analyses)
    _remove_empty_sheets(wb, protected_sheet=ws.title)
    wb.save(output_path)


def _export_transcode_status(analysis: dict[str, Any], value: Any) -> str:
    status = str(analysis.get("status") or "").strip()
    result_text = str(value or "").strip()
    if status == "成功":
        return "可直接采用"
    if status == "待确认":
        return "待人工确认"
    if status == "跳过" or result_text.startswith("跳过"):
        return "跳过"
    if status == "失败" or result_text.startswith("未识别"):
        return "未出码"
    return ""


def _find_product_name_column(ws) -> int | None:
    for cell in ws[1]:
        if _normalize_semantic_header(cell.value) in {"品名", "产品名称", "料品名称"}:
            return int(cell.column)
    return None


def _normalize_comparison_code(value: Any) -> str:
    text = re.sub(r"\s+", "", str(value or "").upper()).split("*", 1)[0]
    match = re.search(r"[A-Z0-9]{22,}", text)
    return match.group(0)[:22] if match else ""


def _export_result_comparison(analysis: dict[str, Any], product_name: Any = "") -> bool | str:
    if str(analysis.get("status") or "").strip() == "跳过":
        return "跳过"
    actual_code = _normalize_comparison_code(analysis.get("formal_code"))
    if not actual_code:
        return "未出码"
    expected_code = _normalize_comparison_code(product_name)
    if not expected_code:
        return "无法对比"
    return actual_code == expected_code


def _remove_empty_sheets(wb, *, protected_sheet: str) -> None:
    """Remove blank and header-only supporting sheets from the download."""
    for sheet_name in list(wb.sheetnames):
        if sheet_name == protected_sheet:
            continue
        ws = wb[sheet_name]
        populated_rows = [
            row
            for row in ws.iter_rows(values_only=True)
            if any(value is not None and str(value).strip() for value in row)
        ]
        if len(populated_rows) <= 1:
            del wb[sheet_name]


def _append_evidence_sheet(wb, analyses: list[dict]) -> None:
    ws = wb.create_sheet("字段证据链")
    headers = [
        "行号",
        "客户代码",
        "客户",
        "规格",
        "状态",
        "字段",
        "原始值",
        "编码片段",
        "置信度",
        "命中方式",
        "规则来源",
        "证据",
        "规则ID",
        "规则类型",
        "来源行号",
    ]
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
                item.get("rule_id", ""),
                item.get("rule_type", ""),
                item.get("source_row", ""),
            ])
    _format_sheet(ws, [10, 14, 18, 46, 10, 12, 18, 12, 10, 18, 24, 54, 18, 18, 12])


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


def _append_summary_sheet(
    wb,
    analyses: list[dict],
    agent_rules: list[dict],
    agent_mapping_tables: dict[str, list[dict]],
    confirm_count: int,
) -> None:
    ws = wb.create_sheet("规则命中汇总")
    status_counter = Counter(analysis.get("status") for analysis in analyses)
    applied_counter = Counter()
    hit_type_counter = Counter()
    field_hit_counter = Counter()
    customer_agent_counter = Counter()
    rule_details: dict[str, dict] = {}
    for analysis in analyses:
        for item in analysis.get("field_evidence", []):
            hit_type = item.get("hit_type", "")
            field = item.get("field", "")
            if hit_type:
                hit_type_counter[hit_type] += 1
            if field and hit_type:
                field_hit_counter[(field, hit_type)] += 1
        for item in analysis.get("applied_rules", []):
            rule_id = item.get("rule_id", "")
            if not rule_id:
                continue
            hit_type = _applied_rule_hit_type(item)
            field_label = OVERRIDE_FIELD_LABELS.get(item.get("field", ""), item.get("field", ""))
            customer = analysis.get("customer", "") or analysis.get("customer_code", "")
            applied_counter[rule_id] += 1
            customer_agent_counter[(customer, hit_type)] += 1
            rule_details.setdefault(
                rule_id,
                {
                    "hit_type": hit_type,
                    "field": field_label,
                    "source_row": item.get("source_row", ""),
                    "rule_type": item.get("rule_type", ""),
                    "text": item.get("text", ""),
                },
            )
    ws.append(["指标", "数量/说明"])
    ws.append(["总行数", len(analyses)])
    ws.append(["高置信出码", status_counter.get("成功", 0)])
    ws.append(["待确认", confirm_count])
    ws.append(["未识别", status_counter.get("失败", 0)])
    ws.append(["跳过PP/RC/%", status_counter.get("跳过", 0)])
    ws.append(["当前Agent机器规则数", len(agent_rules)])
    ws.append(["当前Agent辅助映射总数", _mapping_total_count(agent_mapping_tables)])
    ws.append(["当前已接入辅助映射数", _mapping_enabled_count(agent_mapping_tables)])
    ws.append(["当前待确认/未接入技术项数", len(_technical_pending_rows(agent_mapping_tables))])
    semantic_evaluations = [
        item
        for analysis in analyses
        for item in analysis.get("semantic_shadow", [])
    ]
    semantic_counter = Counter(item.get("status") for item in semantic_evaluations)
    semantic_version = next(
        (analysis.get("semantic_rule_version") for analysis in analyses if analysis.get("semantic_rule_version")),
        "",
    )
    semantic_rule_count = max(
        [int(analysis.get("semantic_rule_count") or 0) for analysis in analyses] or [0]
    )
    semantic_load_error = next(
        (analysis.get("semantic_load_error") for analysis in analyses if analysis.get("semantic_load_error")),
        "",
    )
    ws.append(["模型语义影子版本", semantic_version or "未发布"])
    ws.append(["模型语义正式规则数", semantic_rule_count])
    ws.append(["模型语义影子评估数", len(semantic_evaluations)])
    ws.append(["模型语义影子命中", semantic_counter.get(SHADOW_STATUS_MATCHED, 0)])
    ws.append(["模型语义缺少输入", semantic_counter.get(SHADOW_STATUS_MISSING_INPUT, 0)])
    ws.append(["模型语义未命中", semantic_counter.get(SHADOW_STATUS_NOT_MATCHED, 0)])
    ws.append(["模型语义条件错误", semantic_counter.get(SHADOW_STATUS_ERROR, 0)])
    ws.append(["模型语义加载错误", semantic_load_error])
    ws.append(["模型语义运行时影响", "影子观察；不覆盖编码和评分"])
    order_model_records = [analysis.get("order_semantic_model") or {} for analysis in analyses]
    order_model_counter = Counter(item.get("status") for item in order_model_records)
    ws.append(["DeepSeek实时语义模式", next((item.get("mode") for item in order_model_records if item.get("mode")), "off")])
    ws.append(["DeepSeek实时语义模型", next((item.get("model") for item in order_model_records if item.get("model")), "未启用")])
    ws.append(["DeepSeek实时语义成功", order_model_counter.get("成功", 0)])
    ws.append(["DeepSeek实时语义失败", order_model_counter.get("失败", 0)])
    ws.append(["DeepSeek实时语义缓存命中", sum(1 for item in order_model_records if item.get("cached"))])
    ws.append(["DeepSeek实时语义限流跳过", order_model_counter.get("限流跳过", 0)])
    ws.append(["DeepSeek实时语义运行时影响", "影子观察；不覆盖制造码和正式评分"])
    score_shadows = [
        analysis.get("evidence_score_shadow") or {}
        for analysis in analyses
        if (analysis.get("evidence_score_shadow") or {}).get("field_reviews")
    ]
    score_decisions = Counter(item.get("shadow_decision") for item in score_shadows)
    ws.append(["证据影子评分行数", len(score_shadows)])
    ws.append(["证据影子评分通过", score_decisions.get("通过", 0)])
    ws.append(["证据影子评分需标注", score_decisions.get("需标注", 0)])
    ws.append(["证据影子模型调用", sum(int(item.get("model_call_count") or 0) for item in score_shadows)])
    ws.append(["证据影子模型成功", sum(1 for item in score_shadows if item.get("model_called") and not item.get("model_error"))])
    ws.append(["证据影子模型失败", sum(1 for item in score_shadows if item.get("model_error"))])
    evidence_gates = [analysis.get("evidence_gate") or {} for analysis in analyses if analysis.get("evidence_gate")]
    ws.append(["正式证据门禁模式", next((item.get("mode") for item in evidence_gates), "shadow")])
    ws.append(["正式证据门禁拦截", sum(1 for item in evidence_gates if item.get("blocked"))])
    ws.append(["证据影子运行时影响", "只对比；不覆盖当前分数和90分门禁"])
    ws.append([])
    ws.append(["命中方式汇总", "命中次数"])
    for hit_type, count in hit_type_counter.most_common():
        ws.append([hit_type, count])
    ws.append([])
    ws.append(["字段命中汇总", "命中方式", "命中次数"])
    for (field, hit_type), count in field_hit_counter.most_common():
        ws.append([field, hit_type, count])
    ws.append([])
    ws.append(["客户Agent规则命中汇总", "命中方式", "命中次数"])
    for (customer, hit_type), count in customer_agent_counter.most_common():
        ws.append([customer, hit_type, count])
    ws.append([])
    ws.append(["规则ID", "命中次数", "命中方式", "字段", "规则类型", "来源行号", "规则文本"])
    for rule_id, count in applied_counter.most_common():
        detail = rule_details.get(rule_id, {})
        ws.append([
            rule_id,
            count,
            detail.get("hit_type", ""),
            detail.get("field", ""),
            detail.get("rule_type", ""),
            detail.get("source_row", ""),
            detail.get("text", ""),
        ])
    _format_sheet(ws, [26, 16, 18, 14, 18, 12, 60])


def _applied_rule_hit_type(item: dict) -> str:
    source = str(item.get("source", "") or "")
    if source.startswith("已批准模型语义映射"):
        return "已批准模型语义映射"
    if source.startswith("Agent尺寸映射"):
        return "Agent尺寸映射"
    if source.startswith("Agent厚度映射"):
        return "Agent厚度映射"
    if source.startswith("Agent总芯厚映射"):
        return "Agent总芯厚映射"
    if source.startswith("Agent物料编码口径"):
        return "Agent物料编码口径"
    return "Agent规则覆盖"


def _append_technical_pending_sheet(wb, agent_mapping_tables: dict[str, list[dict]]) -> None:
    ws = wb.create_sheet("技术待支持清单")
    headers = [
        "处理状态",
        "技术类型",
        "客户代码",
        "客户",
        "来源行号",
        "规则ID",
        "规则文本",
        "未执行原因",
        "建议确认/处理",
    ]
    ws.append(headers)
    rows = _technical_pending_rows(agent_mapping_tables)
    for row in rows:
        ws.append([
            row.get("处理状态", ""),
            row.get("技术类型", ""),
            row.get("客户代码", ""),
            row.get("客户简称", ""),
            row.get("来源行号", ""),
            row.get("规则ID", ""),
            row.get("规则文本", ""),
            row.get("未执行原因", ""),
            row.get("建议确认/处理", ""),
        ])
    _format_sheet(ws, [14, 18, 14, 18, 12, 18, 56, 40, 46])


def _append_semantic_shadow_sheet(wb, analyses: list[dict]) -> None:
    ws = wb.create_sheet("模型语义影子证据")
    headers = [
        "行号",
        "客户代码",
        "客户",
        "规格",
        "运行模式",
        "语义规则版本",
        "规则ID",
        "来源候选ID",
        "业务字段",
        "目标字段",
        "标准语义值",
        "原文目标值",
        "影子状态",
        "缺少输入字段",
        "条件校验JSON",
        "观察输入JSON",
        "规则原文",
        "业务证据",
        "模型",
        "说明",
    ]
    ws.append(headers)
    load_error_written = False
    for analysis in analyses:
        mode = analysis.get("semantic_shadow_mode", "")
        version = analysis.get("semantic_rule_version", "")
        load_error = str(analysis.get("semantic_load_error") or "")
        if load_error and not load_error_written:
            ws.append([
                "",
                "",
                "",
                "",
                mode,
                version,
                "",
                "",
                "",
                "",
                "",
                "",
                SHADOW_STATUS_ERROR,
                "",
                "[]",
                "{}",
                "",
                "",
                "",
                f"语义规则加载失败，不影响正式转码：{load_error}",
            ])
            load_error_written = True
        for item in analysis.get("semantic_shadow", []):
            ws.append([
                item.get("row", analysis.get("row")),
                item.get("customer_code", analysis.get("customer_code", "")),
                item.get("customer", analysis.get("customer", "")),
                item.get("spec", analysis.get("spec", "")),
                mode,
                version,
                item.get("rule_id", ""),
                item.get("source_candidate_id", ""),
                item.get("business_field", ""),
                "；".join(str(value) for value in item.get("target_fields", []) if value),
                "；".join(str(value) for value in item.get("normalized_values", []) if value),
                "；".join(str(value) for value in item.get("stated_target_values", []) if value),
                item.get("status", ""),
                "；".join(str(value) for value in item.get("missing_fields", []) if value),
                format_condition_results(item.get("condition_results", [])),
                format_observed_inputs(item.get("observed_inputs", {})),
                item.get("source_text", ""),
                "；".join(str(value) for value in item.get("evidence_texts", []) if value),
                item.get("model", ""),
                item.get("note", ""),
            ])
    _format_sheet(
        ws,
        [10, 14, 18, 42, 12, 30, 24, 24, 14, 20, 24, 24, 12, 22, 56, 56, 64, 52, 20, 42],
    )


def _append_order_semantic_model_sheet(wb, analyses: list[dict]) -> None:
    ws = wb.create_sheet("模型实时语义标准化")
    headers = [
        "行号",
        "客户代码",
        "客户",
        "规格",
        "运行模式",
        "调用状态",
        "模型",
        "缓存命中",
        "相关正式语义规则ID",
        "原始输入JSON",
        "模型置信度",
        "标准语义项JSON",
        "歧义JSON",
        "缺少输入JSON",
        "错误/说明",
        "运行时影响",
    ]
    ws.append(headers)
    for analysis in analyses:
        record = analysis.get("order_semantic_model") or {}
        if record.get("status") in {"未调用", "跳过", ""}:
            continue
        result = record.get("result") or {}
        ws.append([
            analysis.get("row"),
            analysis.get("customer_code", ""),
            analysis.get("customer", ""),
            analysis.get("spec", ""),
            record.get("mode", ""),
            record.get("status", ""),
            record.get("model", ""),
            "是" if record.get("cached") else "否",
            "；".join(record.get("rule_ids") or []),
            json.dumps(record.get("source_fields") or {}, ensure_ascii=False, separators=(",", ":")),
            result.get("model_confidence", ""),
            json.dumps(result.get("semantic_items") or [], ensure_ascii=False, separators=(",", ":")),
            json.dumps(result.get("ambiguities") or [], ensure_ascii=False, separators=(",", ":")),
            json.dumps(result.get("missing_inputs") or [], ensure_ascii=False, separators=(",", ":")),
            record.get("error", "") or record.get("reason", ""),
            "仅观察，不覆盖制造码和正式评分",
        ])
    _format_sheet(ws, [10, 14, 18, 42, 12, 12, 20, 10, 32, 60, 12, 72, 52, 52, 52, 28])


def _append_evidence_score_shadow_sheet(wb, analyses: list[dict]) -> None:
    ws = wb.create_sheet("证据评分影子对比")
    headers = [
        "行号",
        "客户代码",
        "客户",
        "规格",
        "当前状态",
        "正式结果",
        "候选码",
        "当前总分",
        "影子总分",
        "总分差",
        "影子决策",
        "字段",
        "候选原始值",
        "候选编码片段",
        "当前字段分",
        "影子字段分",
        "字段分差",
        "证据结论",
        "命中方式",
        "证据来源",
        "证据文本",
        "评分理由",
        "现有规则ID",
        "语义规则ID",
        "语义证据",
        "原始输入JSON",
        "模型调用",
        "程序结论",
        "程序影子分",
        "模型结论",
        "模型证据来源",
        "模型证据文本",
        "模型理由",
        "模型置信度",
        "模型结论采纳",
        "模型错误",
        "运行时影响",
        "正式证据门禁模式",
        "程序证据分",
        "正式有效分",
        "正式门禁拦截",
        "正式门禁原因",
    ]
    ws.append(headers)
    for analysis in analyses:
        shadow = analysis.get("evidence_score_shadow") or {}
        gate = analysis.get("evidence_gate") or {}
        source_fields_json = format_source_fields(shadow.get("source_fields") or {})
        for review in shadow.get("field_reviews") or []:
            ws.append([
                analysis.get("row"),
                analysis.get("customer_code", ""),
                analysis.get("customer", ""),
                analysis.get("spec", ""),
                analysis.get("status", ""),
                analysis.get("formal_code", ""),
                analysis.get("candidate_code", ""),
                shadow.get("current_score", ""),
                shadow.get("shadow_score", ""),
                shadow.get("score_delta", ""),
                shadow.get("shadow_decision", ""),
                review.get("field", ""),
                review.get("candidate_value", ""),
                review.get("candidate_code", ""),
                review.get("current_score", ""),
                review.get("shadow_score", ""),
                review.get("score_delta", ""),
                review.get("verdict", ""),
                review.get("hit_type", ""),
                review.get("source_field", ""),
                review.get("evidence_text", ""),
                review.get("reason", ""),
                review.get("rule_id", ""),
                "；".join(review.get("semantic_rule_ids") or []),
                "；".join(review.get("semantic_evidence") or []),
                source_fields_json,
                "否" if not review.get("model_called") else "是",
                review.get("program_verdict", ""),
                review.get("program_shadow_score", ""),
                review.get("model_verdict", ""),
                review.get("model_source_field", ""),
                review.get("model_evidence_text", ""),
                review.get("model_reason", ""),
                review.get("model_confidence", ""),
                "是" if review.get("model_accepted") else ("否" if review.get("model_called") else ""),
                shadow.get("model_error", ""),
                shadow.get("runtime_effect", ""),
                gate.get("mode", ""),
                gate.get("program_evidence_score", ""),
                gate.get("effective_score", ""),
                "是" if gate.get("blocked") else "否",
                "；".join(gate.get("blockers") or []) or ("证据有效分低于90" if gate.get("blocked") else ""),
            ])
    _format_sheet(
        ws,
        [10, 14, 18, 42, 12, 24, 24, 12, 12, 10, 12, 12, 18, 18, 12, 12, 10, 18, 18, 24, 44, 52, 20, 34, 52, 64, 12, 18, 14, 18, 24, 44, 52, 14, 14, 52, 48, 18, 14, 14, 14, 42],
    )


def _technical_pending_rows(agent_mapping_tables: dict[str, list[dict]]) -> list[dict]:
    pending: list[dict] = []
    for row in agent_mapping_tables.get("客户物料编码口径", []):
        if _mapping_row_enabled(row):
            continue
        pending.append(
            {
                "处理状态": "待确认后接入",
                "技术类型": "客户物料编码口径",
                "客户代码": row.get("客户代码", ""),
                "客户简称": row.get("客户简称", ""),
                "来源行号": row.get("来源行号", ""),
                "规则ID": row.get("映射ID", ""),
                "规则文本": row.get("规则文本", ""),
                "未执行原因": "物料编码来源字段未确认，避免把规格文本片段误判为订单物料编码",
                "建议确认/处理": "确认字段来源：品号、客户产品编号或其他订单字段；确认命中值是完整编码还是片段",
            }
        )
    for row in agent_mapping_tables.get("外部尺寸表引用", []):
        if _mapping_row_enabled(row):
            continue
        pending.append(
            {
                "处理状态": "待确认后接入",
                "技术类型": "外部尺寸表引用",
                "客户代码": row.get("客户代码", ""),
                "客户简称": row.get("客户简称", ""),
                "来源行号": row.get("来源行号", ""),
                "规则ID": row.get("映射ID", ""),
                "规则文本": row.get("规则文本", ""),
                "未执行原因": "外部尺寸表读取方式和命中字段未确认",
                "建议确认/处理": "确认是否允许运行时读取引用文件，以及命中后是否统一用基础 size_to_code 生成 size_code",
            }
        )
    for row in agent_mapping_tables.get("待接入规则", []):
        technical_type = str(row.get("技术类型", "") or "")
        if "物料" in technical_type:
            continue
        if "物料" in technical_type or "外部" in technical_type or "订单" in technical_type:
            pending.append(
                {
                    "处理状态": "待确认后接入",
                    "技术类型": technical_type,
                    "客户代码": row.get("客户代码", ""),
                    "客户简称": row.get("客户简称", ""),
                    "来源行号": row.get("来源行号", ""),
                    "规则ID": row.get("映射ID", ""),
                    "规则文本": row.get("原始规则", ""),
                    "未执行原因": "待接入规则总览项，本阶段不执行不确定语义",
                    "建议确认/处理": row.get("建议处理", ""),
                }
            )
    return pending


def _mapping_total_count(agent_mapping_tables: dict[str, list[dict]]) -> int:
    return sum(len(rows) for rows in agent_mapping_tables.values())


def _mapping_enabled_count(agent_mapping_tables: dict[str, list[dict]]) -> int:
    executable_sheets = {"客户尺寸映射", "客户单边尺寸映射", "客户尺寸算法", "客户厚度映射", "客户物料编码口径", "外部尺寸表引用"}
    return sum(
        1
        for sheet_name, rows in agent_mapping_tables.items()
        if sheet_name in executable_sheets
        for row in rows
        if _mapping_row_enabled(row)
    )


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
    if len(df_req) > 0:
        header_matches = [
            index
            for index in range(len(df_req.columns))
            if str(df_req.iloc[0, index]).strip() == FORMAL_RESULT_HEADER
        ]
        if header_matches:
            index = header_matches[-1]
            df_req.iloc[:, index] = ""
            return index
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


def _keyword_condition_matches(keyword: str, combined: str) -> bool:
    norm = _norm_match(keyword)
    if not norm:
        return False
    if norm == "__NO_TG__":
        return not re.search(r"TG\d+", combined)
    parts = [part for part in re.split(r"&|\+|和|且", norm) if part]
    if len(parts) > 1:
        return all(_keyword_condition_matches(part, combined) for part in parts)
    if re.fullmatch(r"\d+(?:\.\d+)?", norm):
        return bool(re.search(rf"(?<![A-Z0-9.]){re.escape(norm)}(?![A-Z0-9.])", combined))
    if _is_model_like_token(norm):
        return _model_token_matches(norm, combined)
    return norm in combined


def _glue_condition_matches(token: str, glue_norm: str, spec_norm: str) -> bool:
    token_norm = _norm_match(token)
    if not token_norm:
        return False
    if token_norm == glue_norm:
        return True
    if _is_model_like_token(token_norm):
        return _model_token_matches(token_norm, spec_norm)
    return _token_with_alnum_boundary_matches(token_norm, spec_norm)


def _copper_condition_matches(token: str, copper_norm: str, spec_norm: str) -> bool:
    token_norm = _norm_match(token)
    if not token_norm:
        return False
    if token_norm == copper_norm:
        return True
    return _token_with_alnum_boundary_matches(token_norm, spec_norm)


def _is_model_like_token(token: str) -> bool:
    return bool(re.fullmatch(r"(?:NY|DS|TG|FR)[A-Z0-9().-]*\d[A-Z0-9().-]*", token))


def _model_token_matches(token: str, text: str) -> bool:
    if not token or not text:
        return False
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(token)}(?![A-Z-])", text))


def _token_with_alnum_boundary_matches(token: str, text: str) -> bool:
    if not token or not text:
        return False
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(token)}(?![A-Z0-9])", text))


def _is_legacy_copper_ff_keyword_condition(rule: dict, keyword_condition: str) -> bool:
    keywords = {_norm_match(item) for item in re.split(r"[/,，、;；]+", keyword_condition) if item.strip()}
    return (
        str(rule.get("覆盖字段", "") or "") == "copper_code"
        and str(rule.get("覆盖值", "") or "").strip().upper() == "FF"
        and keywords == {"5"}
    )


def _legacy_copper_ff_keyword_matches(combined: str, steps: dict) -> bool:
    source = f"{combined} {_norm_match(steps.get('copper_spec_raw', ''))}"
    return any(token in source for token in ("1.5/1.5", "R/R", "F/F"))


def _copper_condition_tokens(value: str) -> list[str]:
    source = _norm_match(value)
    known_tokens = ["1.5/1.5", "0.5/0.5", "H/H", "R/R", "F/F", "J/J", "K/K", "W/W", "1/1", "2/2"]
    tokens = [token for token in known_tokens if token in source]
    if tokens:
        return tokens
    return [item for item in re.split(r"[,，;；]+", source) if item]


def _thickness_condition_matches(token: str, steps: dict, spec_norm: str, thickness_norm: str) -> bool:
    raw = str(token or "").strip()
    if not raw:
        return False
    norm = _norm_match(raw)
    if "MIL" in norm:
        return norm in spec_norm or norm in thickness_norm
    match = re.fullmatch(r"(>=|<=|>|<|≥|≤)?(\d+(?:\.\d+)?)", norm)
    if not match:
        return norm in spec_norm or norm in thickness_norm
    operator, number_text = match.groups()
    try:
        expected = float(number_text)
    except ValueError:
        return False
    actual = steps.get("thickness_mm")
    try:
        actual_value = float(actual)
    except (TypeError, ValueError):
        return number_text in thickness_norm
    if operator in (">=", "≥"):
        return actual_value >= expected
    if operator in ("<=", "≤"):
        return actual_value <= expected
    if operator == ">":
        return actual_value > expected
    if operator == "<":
        return actual_value < expected
    return abs(actual_value - expected) < 0.0001


def _norm_customer_code(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _customer_code_tokens(value: Any) -> list[str]:
    return re.findall(r"\d+", str(value or ""))
