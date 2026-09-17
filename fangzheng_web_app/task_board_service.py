"""Read-only task projections and explicitly reported work samples.

No mailbox bootstrap/reclassification or business-state mutation on dashboard reads.
The source of truth remains the existing case and integration records.
"""
from __future__ import annotations

import json
from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo

from .db import db_cursor, list_users

SHANGHAI = ZoneInfo("Asia/Shanghai")
TYPES = {"new_order": "录单", "order_change": "修改订单"}
WORK_SCHEMA = """CREATE TABLE IF NOT EXISTS order_task_work_samples (
    id TEXT PRIMARY KEY,
    case_id BIGINT NOT NULL REFERENCES order_intake_cases(id),
    employee_id TEXT NOT NULL,
    work_date TEXT NOT NULL,
    minutes INTEGER NOT NULL CHECK(minutes BETWEEN 1 AND 480),
    kind TEXT NOT NULL CHECK(kind IN ('operation','rework')),
    note TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL
)"""


def init_work_schema():
    with db_cursor() as conn:
        conn.execute(WORK_SCHEMA)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_task_work_owner_date ON order_task_work_samples(employee_id,work_date,case_id)")


def work_schema_ready(conn):
    if getattr(conn, "dialect", "") == "postgresql":
        return bool(conn.execute("SELECT to_regclass('public.order_task_work_samples') AS name").fetchone()["name"])
    return bool(conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='order_task_work_samples'").fetchone())


def as_time(value, *, source_mail=False):
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=SHANGHAI if source_mail else timezone.utc)
        return parsed.astimezone(SHANGHAI)
    except (ValueError, TypeError):
        return None


def display_time(value, *, source_mail=False):
    parsed = as_time(value, source_mail=source_mail)
    return parsed.strftime("%Y-%m-%d %H:%M:%S") if parsed else "—"


def load_cases(owner=None, case_id=None):
    clauses = ["c.action_type IN ('new_order','order_change')"]
    params = []
    if owner is not None:
        clauses.append("c.employee_id=?")
        params.append(owner)
    if case_id is not None:
        clauses.append("c.id=?")
        params.append(case_id)
    with db_cursor() as conn:
        records = conn.execute(f"""SELECT c.*,m.subject,m.sent_at,m.received_at,m.account_id,
            m.created_at AS mail_imported_at,m.message_id,m.folder,m.uid,a.email AS mailbox,
            (SELECT MIN(e.created_at) FROM order_intake_case_events e
             WHERE e.case_id=c.id AND e.action='processing_started') AS processing_started_at,
            (SELECT e.detail_json FROM order_entry_detail_events e WHERE e.case_id=c.id
             AND e.event_type IN ('domestic_order_entry_real','domestic_order_entry_mock')
             ORDER BY e.id DESC LIMIT 1) AS completion_detail,
            t.id AS template_id,t.current_version,
            (SELECT x.status FROM order_entry_template_tasks x WHERE x.case_id=c.id
             AND x.employee_id=c.employee_id ORDER BY x.id DESC LIMIT 1) AS extraction_status,
            (SELECT COUNT(*) FROM order_entry_template_lines l WHERE l.template_id=t.id) AS line_count,
            (SELECT MAX(e.created_at) FROM order_entry_detail_events e
             WHERE e.case_id=c.id AND e.event_type='order_reply_sent') AS replied_at,
            (SELECT MAX(l.created_at) FROM order_interface_call_logs l
             WHERE l.case_id=c.id AND l.interface_key='aps_order_demand_import' AND l.status='success') AS aps_at,
            (SELECT l.is_mock FROM order_interface_call_logs l WHERE l.case_id=c.id
             AND l.interface_key='domestic_order_entry' AND l.status='success' ORDER BY l.id DESC LIMIT 1) AS completion_mock
            FROM order_intake_cases c JOIN mail_messages m ON m.id=c.mail_id
            LEFT JOIN mail_accounts a ON a.id=m.account_id
            LEFT JOIN order_entry_templates t ON t.case_id=c.id
            WHERE {' AND '.join(clauses)} ORDER BY c.created_at DESC,c.id DESC""", params).fetchall()
    result = []
    now = datetime.now(SHANGHAI)
    for row in records:
        item = dict(row)
        item["created"] = as_time(item.get("created_at"))
        closed = item.get("status") == "archived" and item.get("workflow_stage") == "completed"
        item["completed"] = as_time(item.get("completed_at")) if closed else None
        item["is_closed"] = closed
        item["type_label"] = TYPES[item["action_type"]]
        item["mock"] = item.get("completion_mock") == 1
        if closed:
            item["stage"] = "已完成已回复" if item["replied_at"] else "已完成待回复"
        elif item["action_type"] == "order_change" and item["aps_at"]:
            item["stage"] = "已提交 APS"
        elif item.get("status") == "on_hold":
            item["stage"] = "暂挂"
        elif item["template_id"]:
            item["stage"] = "待录单" if item["action_type"] == "new_order" else "待核对修改"
        elif item["extraction_status"] in {"queued", "running", "error"}:
            item["stage"] = {"queued": "等待提取", "running": "提取中", "error": "提取失败"}[item["extraction_status"]]
        else:
            item["stage"] = "待提取"
        item["age_days"] = max(0, (now.date()-item["created"].date()).days) if item["created"] else None
        item["created_display"] = display_time(item.get("created_at"))
        item["completed_display"] = display_time(item.get("completed_at")) if closed else "—"
        item["mail_display"] = display_time(item.get("sent_at") or item.get("received_at"), source_mail=True)
        item["reply_display"] = display_time(item.get("replied_at"))
        item['mail_date'] = as_time(item.get('received_at') or item.get('sent_at'), source_mail=True) or as_time(item.get('mail_imported_at'))
        item['started'] = as_time(item.get('processing_started_at'))
        item['started_display'] = display_time(item.get('processing_started_at'))
        item['state'] = 'completed' if closed else ('running' if item['started'] or item['template_id'] or item['extraction_status'] else 'pending')
        item['state_label'] = {'completed':'已完成','running':'处理中','pending':'待处理'}[item['state']]
        end = item['completed'] if closed else now
        item['duration_minutes'] = round((end-item['started']).total_seconds()/60, 1) if end and item['started'] and end >= item['started'] else None
        item['duration_display'] = (f"{item['duration_minutes']:g} 分钟" if item['duration_minutes'] is not None else '未记录')
        try:
            detail = json.loads(item.get('completion_detail') or '{}')
            count = detail.get('line_count')
            item['completed_line_count'] = int(count) if count is not None and int(count) >= 0 else None
        except (ValueError, TypeError, AttributeError):
            item['completed_line_count'] = None
        try:
            tags = json.loads(item.get('change_tags_json') or '[]')
            item['subtype'] = '、'.join(t for t in tags if isinstance(t,str)) if isinstance(tags,list) else ''
        except (ValueError, TypeError):
            item['subtype'] = ''
        result.append(item)
    return result


def load_samples(owner=None):
    with db_cursor() as conn:
        if not work_schema_ready(conn):
            return [], False
        sql = "SELECT w.* FROM order_task_work_samples w JOIN order_intake_cases c ON c.id=w.case_id WHERE c.action_type IN ('new_order','order_change')"
        params = []
        if owner is not None:
            sql += " AND c.employee_id=? AND w.employee_id=?"
            params = [owner, owner]
        return [dict(x) for x in conn.execute(sql, params).fetchall()], True


def mail_key(row):
    """Deduplicate only with stable mail identifiers, never subject/date guesses."""
    mailbox = (row.get('mailbox') or '').strip().lower() or str(row['account_id'])
    if row.get('message_id'):
        return (mailbox, row['message_id'].strip())
    return (mailbox, row.get('folder') or 'INBOX', row.get('uid') or str(row.get('mail_id') or row['id']))


def build_board(owner, start, end, *, action_type='all', employee='', state='all', page=1):
    # Select the same canonical case before owner filters so team and personal totals agree.
    all_cases = load_cases()
    canonical = {}
    for case in sorted(all_cases, key=lambda c: (not c['is_closed'], not bool(c['started']), c['id'])):
        canonical.setdefault(mail_key(case), case)
    cases = list(canonical.values())
    names = {str(u['employee_id']): str(u['display_name'] or u['employee_id']) for u in list_users()}
    target = owner if owner is not None else employee
    owners = sorted({*names, *(c['employee_id'] for c in cases)})
    if target:
        cases = [c for c in cases if c['employee_id'] == target]
    if action_type in TYPES:
        cases = [c for c in cases if c['action_type'] == action_type]
    in_period = lambda d: bool(d and start <= d.date() <= end)
    intake = [c for c in cases if in_period(c['mail_date'])]
    completed = [c for c in cases if c['is_closed'] and in_period(c['completed'])]
    # Count all received mail, including mail that has not yet been routed.
    with db_cursor() as conn:
        mails = [dict(r) for r in conn.execute('''SELECT m.*,a.email AS mailbox,
            a.owner_employee_id FROM mail_messages m LEFT JOIN mail_accounts a ON a.id=m.account_id
            ORDER BY m.id''').fetchall()]
    unique_mails = {}
    for mail in mails:
        unique_mails.setdefault(mail_key(mail), mail)
    received = []
    for key, mail in unique_mails.items():
        case = canonical.get(key)
        emp = case['employee_id'] if case else mail.get('owner_employee_id') or ''
        when = as_time(mail.get('received_at') or mail.get('sent_at'), source_mail=True) or as_time(mail.get('created_at'))
        if not in_period(when) or (target and emp != target):
            continue
        if action_type in TYPES and (not case or case['action_type'] != action_type):
            continue
        received.append(emp)
    def metrics(items, done, mail_count):
        timed = [c['duration_minutes'] for c in done if c['duration_minutes'] is not None]
        return dict(received=mail_count, pending=sum(c['state']=='pending' for c in items),
                    running=sum(c['state']=='running' for c in items), completed=sum(c['is_closed'] for c in items),
                    entry_completed=sum(c['action_type']=='new_order' for c in done),
                    change_completed=sum(c['action_type']=='order_change' for c in done),
                    completed_lines=sum(c['completed_line_count'] for c in done if c['completed_line_count'] is not None),
                    missing_lines=sum(c['completed_line_count'] is None for c in done),
                    avg_minutes=round(sum(timed)/len(timed),1) if timed else None,
                    timed_count=len(timed), completed_count=len(done), mock_completed=sum(c['mock'] for c in done))
    summary = metrics(intake, completed, len(received))
    people = []
    for emp in ([target] if target else owners):
        people.append(dict(id=emp, name=names.get(emp,emp), **metrics(
            [c for c in intake if c['employee_id']==emp],
            [c for c in completed if c['employee_id']==emp], received.count(emp))))
    selected = completed if state=='finished' else [c for c in intake if state=='all' or c['state']==state]
    if state == 'finished':
        summary = metrics(completed, completed, len(completed))
    selected.sort(key=lambda c: ({'pending':0,'running':1,'completed':2}[c['state']], c['mail_date'] or datetime.min.replace(tzinfo=SHANGHAI),c['id']))
    page=max(1,min(page,max(1,(len(selected)+19)//20)))
    return dict(summary=summary, people=people, tasks=selected[(page-1)*20:page*20],total=len(selected),page=page,
                pages=max(1,(len(selected)+19)//20),names=names,owners=[(o,names.get(o,o)) for o in owners])


def task_timeline(case_id):
    with db_cursor() as conn:
        records=conn.execute("SELECT event_type,title,operated_by,created_at FROM order_entry_detail_events WHERE case_id=? ORDER BY created_at,id", (case_id,)).fetchall()
        routing=conn.execute("SELECT action,employee_id,created_at FROM order_intake_case_events WHERE case_id=? ORDER BY created_at,id", (case_id,)).fetchall()
    events=[dict(title=r["title"],actor=r["operated_by"],time=display_time(r["created_at"]),sort=as_time(r["created_at"])) for r in records]
    labels={"processing_started":"开始处理", "manual_route":"人工分流", "manual_routing":"人工分流", "aps_order_change_submitted":"提交 APS", "auto_routing":"自动分流"}
    events.extend(dict(title=labels.get(r["action"],"任务状态更新"),actor=r["employee_id"],time=display_time(r["created_at"]),sort=as_time(r["created_at"])) for r in routing)
    return sorted(events,key=lambda e:e["sort"] or datetime.min.replace(tzinfo=SHANGHAI),reverse=True)


def record_work(case_id, employee_id, *, token, work_date, minutes, kind, note):
    import uuid
    uuid.UUID(token)
    day=date.fromisoformat(work_date)
    now=datetime.now(SHANGHAI)
    if day>now.date() or minutes<1 or minutes>480 or kind not in {"operation","rework"}:
        raise ValueError("请填写有效日期、1～480 分钟及工作类型")
    cases=load_cases(employee_id,case_id)
    if not cases:
        raise PermissionError("无权登记此任务")
    if cases[0]["created"] and day<cases[0]["created"].date():
        raise ValueError("工作日期不能早于任务创建日期")
    if len(note)>500:
        raise ValueError("说明最多 500 字")
    with db_cursor() as conn:
        if not work_schema_ready(conn):
            raise ValueError("工时登记尚未启用，请联系管理员完成看板数据库初始化")
        conn.execute("""INSERT INTO order_task_work_samples(id,case_id,employee_id,work_date,minutes,kind,note,created_at)
                     VALUES (?,?,?,?,?,?,?,?) ON CONFLICT(id) DO NOTHING""",
                     (token,case_id,employee_id,day.isoformat(),minutes,kind,note,datetime.now(timezone.utc).isoformat(timespec="seconds")))
