from __future__ import annotations

from datetime import date, datetime, timedelta
from flask import Blueprint, abort, render_template, request, url_for

from .db import is_admin_user
from .routes import current_employee, require_login
from . import task_board_service as service

bp=Blueprint("task_boards",__name__,url_prefix="/task-boards")


@bp.before_request
def authenticate():
    return require_login()


@bp.app_context_processor
def board_navigation():
    return {"can_view_team_board":is_admin_user(current_employee())}


@bp.cli.command("init-db")
def init_database():
    """Create the additive work-sampling table on the configured platform database."""
    service.init_work_schema()
    print("Task board work-sampling schema ready.")


@bp.get("/<scope>")
def overview(scope):
    if scope not in {"mine","team"}:
        abort(404)
    if scope=="team" and not is_admin_user(current_employee()):
        abort(403)
    today=datetime.now(service.SHANGHAI).date()
    try:
        start=date.fromisoformat(request.args.get("start",today.isoformat()))
        end=date.fromisoformat(request.args.get("end",today.isoformat()))
        if end<start or (end-start).days>92:
            raise ValueError()
    except ValueError:
        abort(400,description="请选择开始日期不晚于结束日期、且不超过 93 天的统计范围")
    action=request.args.get("type","all")
    state=request.args.get("state","all")
    tab=request.args.get("tab","tracking")
    if action not in {"all",*service.TYPES} or state not in {"all","pending","running","completed","finished"} or tab not in {"tracking","efficiency"}:
        abort(400)
    if scope=="mine":
        tab="tracking"
    employee=request.args.get("employee","") if scope=="team" else ""
    result=service.build_board(current_employee() if scope=="mine" else None,start,end,
                               action_type=action,employee=employee,state=state,page=request.args.get("page",1,type=int) or 1)
    def board_url(**changes):
        args=dict(start=start.isoformat(),end=end.isoformat(),type=action,state=state,tab=tab,employee=employee)
        args.update(changes)
        return url_for("task_boards.overview",scope=scope,**args)
    return render_template("task_boards/overview.html",scope=scope,start=start,end=end,action=action,state=state,tab=tab,
        employee=employee,board_url=board_url,today=today,yesterday=today-timedelta(days=1),week_start=today-timedelta(days=6),**result)


@bp.get("/task/<int:case_id>")
def detail(case_id):
    actor=current_employee()
    cases=service.load_cases(None if is_admin_user(actor) else actor,case_id)
    if not cases:
        abort(404)
    task=cases[0]
    return render_template("task_boards/detail.html",task=task,events=service.task_timeline(case_id),
                           can_operate=task["employee_id"]==actor)
