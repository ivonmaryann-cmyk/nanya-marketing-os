from __future__ import annotations

import os
import sys
import traceback

from .db import append_job_log, get_job, set_job_worker, update_job_status


def main(argv: list[str] | None = None) -> int:
    args = argv or sys.argv[1:]
    if len(args) != 3:
        print("Usage: python -m fangzheng_web_app.job_worker <job_id> <feature> <employee_id>")
        return 2

    job_id = int(args[0])
    feature = args[1]
    employee_id = args[2]
    set_job_worker(job_id, os.getpid())

    job = get_job(job_id)
    if not job:
        return 1
    if job["status"] == "canceled":
        return 0

    try:
        if feature == "fangzheng":
            from .calculator_service import run_job

            run_job(job_id, employee_id, job["rule_version"])
        elif feature == "transcode":
            from .transcode_service import run_transcode_job

            run_transcode_job(job_id, employee_id)
        elif feature == "transcode_agent":
            from .transcode_agent_service import run_transcode_agent_job

            run_transcode_agent_job(job_id, employee_id)
        elif feature == "shennan":
            from .shennan_service import run_shennan_job

            run_shennan_job(job_id, employee_id)
        elif feature == "hushi":
            from .hushi_service import run_hushi_job

            run_hushi_job(job_id, employee_id)
        elif feature == "bomin":
            from .bomin_service import run_bomin_job

            run_bomin_job(job_id, employee_id)
        elif feature == "in_transit":
            from .in_transit_service import run_in_transit_job

            run_in_transit_job(job_id, employee_id)
        elif feature == "order_reprice":
            from .order_reprice_service import run_order_reprice_job

            run_order_reprice_job(job_id, employee_id)
        elif feature == "price_calculation":
            from .price_calculation_service import run_price_calculation_job

            run_price_calculation_job(job_id, employee_id)
        elif feature == "transcode_special_import":
            from .transcode_special_import_service import run_transcode_special_import_job

            run_transcode_special_import_job(job_id, employee_id)
        else:
            raise ValueError(f"未知任务类型：{feature}")
        return 0
    except Exception as exc:
        append_job_log(job_id, f"计算子进程异常：{exc}")
        update_job_status(
            job_id,
            status="failed",
            error_message=f"{exc}\n{traceback.format_exc(limit=8)}",
            completed=True,
        )
        return 1
    finally:
        set_job_worker(job_id, None)


if __name__ == "__main__":
    raise SystemExit(main())
