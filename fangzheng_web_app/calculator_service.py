from __future__ import annotations

import importlib
import sys
import traceback
from datetime import datetime
from pathlib import Path

from werkzeug.utils import secure_filename

from .db import append_job_log, create_job, prune_jobs_for_employee, update_job_status
from .excel_utils import load_workbook_compat, normalized_xlsx_source
from .file_utils import safe_unlink
from .job_control import launch_job_process
from .paths import JOBS_DIR
from .rules import get_active_rule_version, load_rule_dataframes

CALCULATOR_MODULE_NAME = "fangzheng_web_app.price_calculator_v3"
NON_DATA_DESCRIPTIONS = {"物料描述", "规格", "客户规格", "材料描述", "物料规格"}


def load_calculator_module():
    """Reload the packaged calculator so web jobs follow its latest logic."""
    if CALCULATOR_MODULE_NAME in sys.modules:
        return importlib.reload(sys.modules[CALCULATOR_MODULE_NAME])
    return importlib.import_module(CALCULATOR_MODULE_NAME)


def is_effective_description(value) -> bool:
    text = str(value).strip() if value is not None else ""
    return bool(text and text.lower() not in {"nan", "none"} and text not in NON_DATA_DESCRIPTIONS)


def calculate_fangzheng_quote(spec: str) -> dict:
    spec = str(spec or "").strip()
    if not spec:
        return {"status": "失败", "price": None, "error": "请输入客户规格"}

    rule_version = get_active_rule_version()
    calculator = load_calculator_module()
    price_df, account_df = load_rule_dataframes(rule_version)
    price, note, err = calculator.calculate_price(spec, price_df, account_df)
    if err:
        return {
            "status": "失败",
            "price": None,
            "note": note or "",
            "material_type": "方正价格",
            "rule_version": rule_version,
            "error": err,
        }
    return {
        "status": "成功",
        "price": calculator.round_price(price) if price is not None else None,
        "note": note or "计算成功",
        "material_type": "方正价格",
        "rule_version": rule_version,
        "error": "",
    }


def queue_job(employee_id: str, uploaded_file, source_filename: str, rule_version: str) -> int:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    employee_dir = JOBS_DIR / employee_id
    employee_dir.mkdir(parents=True, exist_ok=True)

    safe_filename = secure_filename(source_filename) or f"upload_{timestamp}.xlsx"
    input_path = employee_dir / f"{timestamp}_{safe_filename}"
    uploaded_file.save(input_path)

    job_id = create_job(employee_id, source_filename, str(input_path), rule_version)
    launch_job_process(job_id, "fangzheng", employee_id)
    return job_id


def run_job(job_id: int, employee_id: str, rule_version: str) -> None:
    from .db import get_job

    update_job_status(job_id, status="running", log_text="")
    append_job_log(job_id, f"开始处理任务，规则版本：{rule_version}")
    job = get_job(job_id)
    if not job:
        return

    try:
        calculator = load_calculator_module()
        calculator_path = Path(calculator.__file__).resolve()
        price_df, account_df = load_rule_dataframes(rule_version)
        append_job_log(job_id, f"规则加载完成：价格表 {len(price_df)} 行，基板表 {len(account_df)} 行")
        append_job_log(job_id, f"计算引擎已加载：{calculator_path.name}（{datetime.fromtimestamp(calculator_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')}）")

        workbook = load_workbook_compat(job["stored_input_path"], data_only=True)
        source_for_result = normalized_xlsx_source(job["stored_input_path"], workbook)
        sheet_name = calculator.select_calculation_sheet_name(workbook.sheetnames)
        sheet = workbook[sheet_name]
        desc_col = calculator.detect_description_column_openpyxl(sheet)
        desc_index = desc_col - 1
        data_rows = []
        for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            desc = str(row[desc_index]).strip() if len(row) > desc_index and row[desc_index] is not None else ""
            if is_effective_description(desc):
                data_rows.append((row_index, row, desc))
        total_rows = len(data_rows)
        update_job_status(job_id, status="running", total_rows=total_rows)
        append_job_log(job_id, f"使用 Sheet：{sheet_name}，物料描述列：第 {desc_col} 列")
        append_job_log(job_id, f"检测到 {total_rows} 行有效数据", total_rows=total_rows)

        results = []
        success_count = 0
        fail_count = 0
        skip_count = 0

        for processed, (index, row, desc) in enumerate(data_rows, start=1):
            price, note, err = calculator.calculate_price(desc, price_df, account_df)
            pp_roll_price = calculator.calculate_pp_roll_price(desc, price_df)
            pp_roll_price_value = calculator.round_price(pp_roll_price) if pp_roll_price is not None else ""
            if err:
                fail_count += 1
                results.append({"行号": index, "物料描述": desc, "价格": "", "输出价格": "", "说明": err, "状态": "失败"})
                append_job_log(
                    job_id,
                    f"第 {index} 行失败：{err}",
                    fail_count=fail_count,
                    current_row=processed,
                    total_rows=total_rows,
                )
            else:
                success_count += 1
                output_price = calculator.output_price_for_desc(desc, price, pp_roll_price_value)
                results.append({"行号": index, "物料描述": desc, "价格": price, "输出价格": output_price, "说明": note, "状态": "成功"})
                append_job_log(
                    job_id,
                    f"第 {index} 行成功：{output_price}",
                    success_count=success_count,
                    current_row=processed,
                    total_rows=total_rows,
                )

        input_path = Path(job["stored_input_path"])
        output_path = input_path.with_name(f"{input_path.stem}_计算结果.xlsx")
        calculator.save_result_v3(str(source_for_result), results, str(output_path), sheet_name=sheet_name)
        append_job_log(job_id, "结果文件已生成，任务完成", current_row=total_rows, total_rows=total_rows)

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
            error_message=f"{exc}\n{traceback.format_exc(limit=5)}",
            completed=True,
        )
    finally:
        stale_jobs = prune_jobs_for_employee(employee_id, keep_limit=500)
        for stale in stale_jobs:
            for key in ["stored_input_path", "stored_result_path"]:
                safe_unlink(stale[key])
