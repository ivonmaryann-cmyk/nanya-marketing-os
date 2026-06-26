from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from werkzeug.utils import secure_filename

from .db import append_job_log, create_job, prune_jobs_for_employee, update_job_status
from .file_utils import safe_unlink
from .job_control import launch_job_process
from .paths import JOBS_DIR
from .transcode_special_rules import (
    build_export_workbook,
    parse_bulk_special_requirement_workbook_object,
    save_latest_original_import,
    save_structured_special_rules,
)


FEATURE_NAME = "transcode_special_import"


def queue_transcode_special_import_job(employee_id: str, uploaded_file, source_filename: str) -> int:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    employee_dir = JOBS_DIR / employee_id
    employee_dir.mkdir(parents=True, exist_ok=True)

    safe_filename = secure_filename(source_filename) or f"special_rules_{timestamp}.xlsx"
    input_path = employee_dir / f"{timestamp}_special_rules_{safe_filename}"
    uploaded_file.save(input_path)

    job_id = create_job(
        employee_id,
        source_filename,
        str(input_path),
        "客户特殊规则结构化导入",
        feature=FEATURE_NAME,
    )
    launch_job_process(job_id, FEATURE_NAME, employee_id)
    return job_id


def run_transcode_special_import_job(job_id: int, employee_id: str) -> None:
    from .db import get_job

    update_job_status(job_id, status="running", log_text="")
    job = get_job(job_id)
    if not job:
        return

    append_job_log(job_id, "开始批量导入客户特殊要求结构化规则")

    try:
        input_path = Path(job["stored_input_path"])
        file_bytes = input_path.read_bytes()
        save_latest_original_import(file_bytes, job["source_filename"])
        append_job_log(job_id, f"原文件已保存：{job['source_filename']}")

        workbook = load_workbook(input_path, read_only=True, data_only=True)

        last_logged_bucket = {"value": -1}

        def progress(processed: int, total: int, stats: dict[str, int]) -> None:
            update_job_status(
                job_id,
                status="running",
                current_row=processed,
                total_rows=total,
                success_count=stats.get("rules", 0),
                skip_count=stats.get("skipped_rows", 0),
            )
            bucket = processed // 100
            if bucket != last_logged_bucket["value"] or processed == total:
                last_logged_bucket["value"] = bucket
                append_job_log(
                    job_id,
                    f"解析进度：{processed}/{total}，已生成 {stats.get('rules', 0)} 条规则，跳过 {stats.get('skipped_rows', 0)} 行",
                    current_row=processed,
                    total_rows=total,
                    success_count=stats.get("rules", 0),
                    skip_count=stats.get("skipped_rows", 0),
                )

        rules, stats = parse_bulk_special_requirement_workbook_object(workbook, progress_callback=progress)
        if not rules:
            append_job_log(job_id, "未解析出可保存的结构化规则")
            update_job_status(
                job_id,
                status="failed",
                error_message="没有从文件中解析出可保存的结构化规则，请检查客户代码列和特殊要求列。",
                completed=True,
            )
            return

        append_job_log(job_id, f"开始保存结构化规则：{stats['customers']} 个客户，{len(rules)} 条规则")
        _, saved_ids = save_structured_special_rules(
            rules,
            saved_by=employee_id,
            import_mode="替换",
        )

        output_path = input_path.with_name(f"{input_path.stem}_解析后规则包.xlsx")
        export = build_export_workbook("full")
        output_path.write_bytes(export.getvalue())

        append_job_log(
            job_id,
            f"批量导入完成：读取 {stats['source_rows']} 行，覆盖 {stats['customers']} 个客户，保存 {len(saved_ids)} 条规则",
            current_row=stats.get("total_rows", stats["source_rows"] + stats["skipped_rows"]),
            total_rows=stats.get("total_rows", stats["source_rows"] + stats["skipped_rows"]),
            success_count=len(saved_ids),
            skip_count=stats["skipped_rows"],
        )
        update_job_status(
            job_id,
            status="completed",
            stored_result_path=str(output_path),
            success_count=len(saved_ids),
            fail_count=0,
            skip_count=stats["skipped_rows"],
            current_row=stats.get("total_rows", stats["source_rows"] + stats["skipped_rows"]),
            total_rows=stats.get("total_rows", stats["source_rows"] + stats["skipped_rows"]),
            completed=True,
        )
    except Exception as exc:
        append_job_log(job_id, f"批量导入失败：{exc}")
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
