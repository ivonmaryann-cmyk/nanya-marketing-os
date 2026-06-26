from __future__ import annotations

import os
import signal
import subprocess
import sys
import time
from pathlib import Path

from .db import append_job_log, db_cursor, get_job, set_job_worker, update_job_status
from .paths import JOBS_DIR, PROJECT_DIR


RUNNING_STATUSES = {"queued", "running"}


def launch_job_process(job_id: int, feature: str, employee_id: str) -> None:
    job_log_dir = JOBS_DIR / employee_id
    job_log_dir.mkdir(parents=True, exist_ok=True)
    process_log = job_log_dir / f"job_{job_id}_worker.log"
    command = [
        sys.executable,
        "-m",
        "fangzheng_web_app.job_worker",
        str(job_id),
        feature,
        employee_id,
    ]
    try:
        log_handle = process_log.open("ab")
        popen_kwargs = {
            "cwd": PROJECT_DIR,
            "stdout": log_handle,
            "stderr": subprocess.STDOUT,
            "close_fds": os.name != "nt",
        }
        if os.name == "nt":
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
            if hasattr(subprocess, "CREATE_NO_WINDOW"):
                creationflags |= subprocess.CREATE_NO_WINDOW
            popen_kwargs["creationflags"] = creationflags
        process = subprocess.Popen(
            command,
            **popen_kwargs,
        )
        log_handle.close()
        set_job_worker(job_id, process.pid)
        append_job_log(job_id, f"计算子进程已启动：PID {process.pid}")
    except Exception as exc:
        update_job_status(
            job_id,
            status="failed",
            error_message=f"计算子进程启动失败：{exc}",
            completed=True,
        )


def is_process_alive(pid: int | None) -> bool:
    if not pid:
        return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def cancel_job_process(job_id: int, employee_id: str) -> tuple[bool, str]:
    job = get_job(job_id)
    if not job or job["employee_id"] != employee_id:
        return False, "未找到该任务。"
    if job["status"] not in RUNNING_STATUSES:
        return False, "只有排队中或运行中的任务可以停止。"

    pid = job["worker_pid"] if "worker_pid" in job.keys() else None
    append_job_log(job_id, "用户请求停止任务。")
    if pid and is_process_alive(int(pid)):
        _terminate_pid(int(pid))
    set_job_worker(job_id, None)
    update_job_status(
        job_id,
        status="canceled",
        error_message="任务已由用户停止。",
        completed=True,
    )
    return True, "任务已停止。"


def _terminate_pid(pid: int) -> None:
    try:
        os.kill(pid, signal.SIGTERM)
    except OSError:
        return
    time.sleep(0.5)
    if os.name != "nt" and is_process_alive(pid):
        try:
            os.kill(pid, signal.SIGKILL)
        except OSError:
            pass


def clear_worker_if_current(job_id: int, pid: int) -> None:
    job = get_job(job_id)
    if job and "worker_pid" in job.keys() and job["worker_pid"] == pid:
        set_job_worker(job_id, None)


def reconcile_interrupted_jobs() -> None:
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT id, worker_pid FROM jobs WHERE status IN ('queued', 'running')"
        ).fetchall()
    for row in rows:
        pid = row["worker_pid"]
        if pid and is_process_alive(int(pid)):
            continue
        append_job_log(row["id"], "服务启动时检测到任务未在运行，已标记为失败。")
        set_job_worker(row["id"], None)
        update_job_status(
            row["id"],
            status="failed",
            error_message="服务重启或计算子进程异常退出，任务未完成。",
            completed=True,
        )
