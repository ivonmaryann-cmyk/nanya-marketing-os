from __future__ import annotations

from datetime import date

from flask import Blueprint, abort, flash, jsonify, redirect, render_template, request, session, url_for

from ..db import is_admin_user
from ..routes import require_login
from . import mail_order_service, mail_store
from .mail_fetch_service import fetch_latest_order_mails, test_imap_connection
from .smtp_service import test_smtp_connection


bp = Blueprint("mail_transcode", __name__)


def _owner() -> str:
    return str(session.get("employee_id") or "").strip()


def _build_fetch_task_summary(fetch_task: dict, owner_employee_id: str) -> dict:
    """Build the business-facing status for one mail fetch task."""
    messages = mail_store.get_messages_by_fetch_task(
        int(fetch_task["id"]), owner_employee_id=owner_employee_id
    )
    orders = [
        order
        for message in messages
        for order in mail_store.get_tasks_by_mail(
            int(message["id"]), owner_employee_id=owner_employee_id
        )
    ]
    account = mail_store.get_account(
        int(fetch_task["account_id"]), owner_employee_id=owner_employee_id
    ) or {}
    pending_count = sum(
        1
        for order in orders
        if order["review_status"] == "pending_review"
        or order["transcode_status"] in {"待核实", "待确认", "失败"}
    )
    return {
        **fetch_task,
        "mailbox": str(account.get("email") or "当前测试邮箱"),
        "message_count": len(messages),
        "actual_order_count": len(orders),
        "pending_count": pending_count,
        "passed_count": sum(1 for order in orders if order["transcode_status"] == "已通过"),
    }


@bp.before_request
def _require_login():
    return require_login()


@bp.get("/")
def index():
    owner_employee_id = _owner()
    accounts = mail_store.list_accounts(owner_employee_id=owner_employee_id)
    fetch_tasks = mail_store.list_fetch_tasks(limit=10, owner_employee_id=owner_employee_id)
    all_fetch_tasks = mail_store.list_fetch_tasks(limit=1000, owner_employee_id=owner_employee_id)
    order_tasks = mail_store.list_order_tasks(limit=1000, owner_employee_id=owner_employee_id)
    today_prefix = date.today().isoformat()
    task_summaries = [_build_fetch_task_summary(task, owner_employee_id) for task in fetch_tasks]
    return render_template(
        "mail_transcode_agent/index.html",
        accounts=accounts,
        tasks=task_summaries,
        task_count=len(all_fetch_tasks),
        pending_task_count=sum(task["pending_count"] for task in task_summaries),
        enabled_account_count=sum(1 for account in accounts if account["enabled"]),
        today_order_count=sum(
            int(task["order_count"] or 0)
            for task in all_fetch_tasks
            if str(task["started_at"] or "").startswith(today_prefix)
        ),
        completed_transcode_count=sum(
            1 for task in order_tasks if task["transcode_status"] == "已通过"
        ),
    )


@bp.get("/accounts")
def accounts_page():
    owner_employee_id = _owner()
    edit_id = request.args.get("edit", type=int)
    edit_account = (
        mail_store.get_account(edit_id, owner_employee_id=owner_employee_id) if edit_id else None
    )
    if edit_id and not edit_account:
        abort(404)
    return render_template(
        "mail_transcode_agent/accounts.html",
        accounts=mail_store.list_accounts(owner_employee_id=owner_employee_id),
        edit_account=edit_account,
        smtp_config=mail_store.smtp_public_config(edit_account),
    )


@bp.post("/accounts")
def save_account():
    owner_employee_id = _owner()
    email = str(request.form.get("email") or "").strip()
    account_id = int(request.form.get("account_id") or 0)
    if not email:
        flash("请输入邮箱地址", "error")
        return redirect(url_for("mail_transcode.accounts_page"))
    try:
        imap_auth_code = str(request.form.get("auth_code") or "")
        saved_account_id = mail_store.create_or_update_account(
            email,
            owner_employee_id=owner_employee_id,
            account_id=account_id,
            imap_host=str(request.form.get("imap_host") or "imaphz.qiye.163.com"),
            imap_port=int(request.form.get("imap_port") or 993),
            auth_code=imap_auth_code,
            enabled=1 if request.form.get("enabled") else 0,
            allow_duplicate_email=is_admin_user(owner_employee_id),
        )
        if request.form.get("smtp_config_present"):
            mail_store.save_smtp_config(
                saved_account_id,
                owner_employee_id=owner_employee_id,
                host=str(request.form.get("smtp_host") or ""),
                port=request.form.get("smtp_port") or 465,
                security=str(request.form.get("smtp_security") or "ssl"),
                username=str(request.form.get("smtp_username") or ""),
                auth_code=(
                    imap_auth_code
                    if request.form.get("use_imap_auth_code")
                    else str(request.form.get("smtp_auth_code") or "")
                ),
                sender_name=str(request.form.get("smtp_sender_name") or ""),
                enabled=0,
            )
        flash("邮箱账号已保存", "success")
        if not account_id:
            flash("可在编辑页测试 SMTP 连接；测试成功后才能启用真实发信。", "success")
            return redirect(url_for("mail_transcode.accounts_page", edit=saved_account_id))
    except Exception as exc:
        flash(f"保存失败：{exc}", "error")
    return redirect(url_for("mail_transcode.accounts_page"))


@bp.post("/accounts/<int:account_id>/delete")
def delete_account(account_id: int):
    try:
        email = mail_store.delete_account(account_id, owner_employee_id=_owner())
        flash(f"已删除 {email} 的邮箱配置。原邮箱和已导入的本地历史记录不会受影响。", "success")
    except ValueError as exc:
        flash(f"删除失败：{exc}", "error")
    return redirect(url_for("mail_transcode.accounts_page"))


@bp.post("/accounts/<int:account_id>/reveal-auth-code")
def reveal_auth_code(account_id: int):
    auth_code = mail_store.get_account_auth_code(account_id, owner_employee_id=_owner())
    if auth_code is None:
        abort(404)
    return jsonify({"auth_code": auth_code})


@bp.post("/accounts/<int:account_id>/reveal-smtp-auth-code")
def reveal_smtp_auth_code(account_id: int):
    config = mail_store.get_smtp_config(account_id, owner_employee_id=_owner())
    if config is None:
        abort(404)
    return jsonify({"auth_code": str(config.get("auth_code") or "")})


@bp.post("/accounts/<int:account_id>/test-connection")
def test_account_connection(account_id: int):
    try:
        result = test_imap_connection(account_id, owner_employee_id=_owner())
        return jsonify({"ok": True, **result})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception:
        return jsonify({"ok": False, "error": "连接失败，请核对企业邮箱 IMAP 服务器、端口和客户端授权码。"}), 502


@bp.post("/accounts/<int:account_id>/smtp")
def save_account_smtp(account_id: int):
    owner_employee_id = _owner()
    try:
        smtp_auth_code = str(request.form.get("smtp_auth_code") or "")
        if request.form.get("use_imap_auth_code"):
            smtp_auth_code = mail_store.get_account_auth_code(
                account_id, owner_employee_id=owner_employee_id
            ) or ""
        mail_store.save_smtp_config(
            account_id,
            owner_employee_id=owner_employee_id,
            host=str(request.form.get("smtp_host") or ""),
            port=int(request.form.get("smtp_port") or 465),
            security=str(request.form.get("smtp_security") or "ssl"),
            username=str(request.form.get("smtp_username") or ""),
            auth_code=smtp_auth_code,
            sender_name=str(request.form.get("smtp_sender_name") or ""),
            enabled=1 if request.form.get("smtp_enabled") else 0,
        )
        flash("SMTP 发信配置已保存。授权码不会在页面中回显。", "success")
    except ValueError as exc:
        flash(f"SMTP 保存失败：{exc}", "error")
    return redirect(url_for("mail_transcode.accounts_page", edit=account_id))


@bp.post("/accounts/<int:account_id>/test-smtp-connection")
def test_account_smtp_connection(account_id: int):
    try:
        result = test_smtp_connection(account_id, owner_employee_id=_owner())
        return jsonify({"ok": True, **result})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@bp.post("/fetch")
def fetch_mails():
    owner_employee_id = _owner()
    accounts = mail_store.list_accounts(owner_employee_id=owner_employee_id)
    enabled = [account for account in accounts if account["enabled"]]
    if not enabled:
        flash("请先配置并启用邮箱账号", "error")
        return redirect(url_for("mail_transcode.index"))
    account_id = int(request.form.get("account_id") or enabled[0]["id"])
    try:
        result = fetch_latest_order_mails(
            account_id,
            created_by=owner_employee_id,
            owner_employee_id=owner_employee_id,
        )
        flash(result["message"], "success")
    except Exception as exc:
        flash(f"抓取失败：{exc}", "error")
    return redirect(url_for("mail_transcode.index"))


@bp.get("/orders")
def orders_page():
    page = max(1, request.args.get("page", 1, type=int))
    per_page = 20
    fetch_tasks = mail_store.list_fetch_tasks(limit=1000, owner_employee_id=_owner())
    total_pages = max(1, (len(fetch_tasks) + per_page - 1) // per_page)
    page = min(page, total_pages)
    start = (page - 1) * per_page
    page_tasks = fetch_tasks[start : start + per_page]
    return render_template(
        "mail_transcode_agent/orders.html",
        fetch_tasks=page_tasks,
        page=page,
        total_pages=total_pages,
    )


@bp.get("/fetch-tasks/<int:fetch_task_id>")
def fetch_task_detail(fetch_task_id: int):
    owner_employee_id = _owner()
    fetch_task = mail_store.get_fetch_task(fetch_task_id, owner_employee_id=owner_employee_id)
    if not fetch_task:
        abort(404)
    messages = mail_store.get_messages_by_fetch_task(
        fetch_task_id, owner_employee_id=owner_employee_id
    )
    rows = []
    for message in messages:
        tasks = mail_store.get_tasks_by_mail(
            int(message["id"]), owner_employee_id=owner_employee_id
        )
        rows.append(
            {
                "message": message,
                "order_count": len(tasks),
                "ready_count": sum(
                    1
                    for task in tasks
                    if task["field_status"] == "complete" and task["review_status"] == "reviewed"
                ),
                "pending_count": sum(
                    1 for task in tasks if task["review_status"] == "pending_review"
                ),
            }
        )
    return render_template(
        "mail_transcode_agent/fetch_task_detail.html",
        fetch_task=fetch_task,
        rows=rows,
    )


@bp.get("/tasks/<int:mail_id>")
def task_detail(mail_id: int):
    owner_employee_id = _owner()
    message = mail_store.get_message(mail_id, owner_employee_id=owner_employee_id)
    if not message:
        abort(404)
    tasks = mail_store.get_tasks_by_mail(mail_id, owner_employee_id=owner_employee_id)
    return render_template(
        "mail_transcode_agent/task_detail.html",
        message=message,
        tasks=tasks,
    )


@bp.get("/orders/<int:task_id>")
def order_detail(task_id: int):
    task = mail_store.get_order_task(task_id, owner_employee_id=_owner())
    if not task:
        abort(404)
    attachments = mail_store.list_attachments(int(task["mail_id"]))
    return render_template(
        "mail_transcode_agent/order_detail.html",
        task=task,
        attachments=attachments,
    )


@bp.post("/orders/<int:task_id>/review")
def review_order(task_id: int):
    task = mail_store.get_order_task(task_id, owner_employee_id=_owner())
    if not task:
        abort(404)
    mail_store.update_order_task(
        task_id,
        customer_code=request.form.get("customer_code") or "",
        customer_name=request.form.get("customer_name") or "",
        spec=request.form.get("spec") or "",
        remark=request.form.get("remark") or "",
        order_number=request.form.get("order_number") or "",
    )
    flash("订单字段已保存", "success")
    return redirect(url_for("mail_transcode.order_detail", task_id=task_id))


@bp.post("/orders/<int:task_id>/parse-attachments")
def parse_attachments(task_id: int):
    task = mail_store.get_order_task(task_id, owner_employee_id=_owner())
    if not task:
        abort(404)
    try:
        result = mail_order_service.parse_attachments_for_task(task_id)
        flash(f"附件解析完成：识别 {result['rows']} 行；{len(result['errors'])} 个错误", "success")
    except Exception as exc:
        flash(f"附件解析失败：{exc}", "error")
    return redirect(url_for("mail_transcode.order_detail", task_id=task_id))


@bp.post("/run")
def run_transcode():
    raw_ids = request.form.get("task_ids") or ""
    task_ids = [int(value) for value in raw_ids.split(",") if value.strip().isdigit()]
    if not task_ids:
        flash("请选择要执行的订单任务", "error")
        return redirect(url_for("mail_transcode.orders_page"))
    owner_employee_id = _owner()
    owned_task_ids = [
        task_id
        for task_id in task_ids
        if mail_store.get_order_task(task_id, owner_employee_id=owner_employee_id)
    ]
    if len(owned_task_ids) != len(task_ids):
        flash("不能执行其他用户的订单任务", "error")
        return redirect(url_for("mail_transcode.orders_page"))
    try:
        job_id = mail_order_service.run_transcode_for_tasks(owned_task_ids, owner_employee_id)
        return redirect(url_for("mail_transcode.job_detail", job_id=job_id))
    except Exception as exc:
        flash(f"转码执行失败：{exc}", "error")
        return redirect(url_for("mail_transcode.orders_page"))


@bp.get("/jobs")
def jobs_page():
    page = max(1, request.args.get("page", 1, type=int))
    per_page = 20
    jobs = mail_store.list_transcode_jobs(limit=1000, owner_employee_id=_owner())
    total_pages = max(1, (len(jobs) + per_page - 1) // per_page)
    page = min(page, total_pages)
    start = (page - 1) * per_page
    page_jobs = jobs[start : start + per_page]
    return render_template(
        "mail_transcode_agent/jobs.html",
        jobs=page_jobs,
        page=page,
        total_pages=total_pages,
    )


@bp.get("/jobs/<int:job_id>")
def job_detail(job_id: int):
    owner_employee_id = _owner()
    job = mail_store.get_transcode_job(job_id, owner_employee_id=owner_employee_id)
    if not job:
        abort(404)
    task_ids = [int(value) for value in (job["task_ids"] or "").split(",") if value.strip().isdigit()]
    tasks = [
        mail_store.get_order_task(task_id, owner_employee_id=owner_employee_id)
        for task_id in task_ids
    ]
    tasks = [task for task in tasks if task]
    return render_template(
        "mail_transcode_agent/job_detail.html",
        job=job,
        tasks=tasks,
    )


@bp.post("/jobs/<int:job_id>/verify")
def verify_job_task(job_id: int):
    owner_employee_id = _owner()
    job = mail_store.get_transcode_job(job_id, owner_employee_id=owner_employee_id)
    if not job:
        abort(404)
    task_id = int(request.form.get("task_id") or 0)
    action = str(request.form.get("action") or "")
    task = (
        mail_store.get_order_task(task_id, owner_employee_id=owner_employee_id) if task_id else None
    )
    if not task or action not in {"approve", "reject"}:
        flash("参数错误", "error")
        return redirect(url_for("mail_transcode.job_detail", job_id=job_id))
    mail_store.verify_order_task(task_id, action, owner_employee_id)
    flash("核实结果已保存", "success")
    return redirect(url_for("mail_transcode.job_detail", job_id=job_id))
