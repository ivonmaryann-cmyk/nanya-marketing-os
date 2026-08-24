from __future__ import annotations

from datetime import datetime
from typing import Any

from ..database import automation_cursor as db_cursor
from ..transcode_customer_rule_admin import resolve_customer_code_by_name
from .mail_crypto import decrypt_text, encrypt_text


def now_iso() -> str:
    return datetime.now().replace(microsecond=0).isoformat()


def _rows(rows) -> list[dict[str, Any]]:
    return [dict(row) for row in rows]


def _one(row) -> dict[str, Any] | None:
    return dict(row) if row else None


def list_accounts(*, owner_employee_id: str | None = None) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        if owner_employee_id:
            rows = conn.execute(
                "SELECT * FROM mail_accounts WHERE owner_employee_id = ? ORDER BY id DESC",
                (owner_employee_id,),
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM mail_accounts ORDER BY id DESC").fetchall()
    return _rows(rows)


def get_account(account_id: int, *, owner_employee_id: str | None = None) -> dict[str, Any] | None:
    with db_cursor() as conn:
        sql = "SELECT * FROM mail_accounts WHERE id = ?"
        params: list[Any] = [account_id]
        if owner_employee_id:
            sql += " AND owner_employee_id = ?"
            params.append(owner_employee_id)
        row = conn.execute(sql, params).fetchone()
    return _one(row)


def get_account_auth_code(account_id: int, *, owner_employee_id: str) -> str | None:
    account = get_account(account_id, owner_employee_id=owner_employee_id)
    if not account:
        return None
    return decrypt_text(str(account.get("auth_code_ciphertext") or ""))


def create_or_update_account(
    email: str,
    *,
    owner_employee_id: str,
    account_id: int = 0,
    imap_host: str = "imaphz.qiye.163.com",
    imap_port: int = 993,
    auth_code: str = "",
    enabled: int = 1,
) -> int:
    owner_employee_id = str(owner_employee_id or "").strip()
    if not owner_employee_id:
        raise ValueError("未识别当前登录用户")
    now = now_iso()
    email = email.strip()
    auth_code = str(auth_code or "").strip()
    with db_cursor() as conn:
        if account_id:
            existing = conn.execute(
                "SELECT * FROM mail_accounts WHERE id = ? AND owner_employee_id = ?",
                (int(account_id), owner_employee_id),
            ).fetchone()
            if not existing:
                raise ValueError("邮箱账号不存在或无权编辑")
            duplicate = conn.execute(
                "SELECT id FROM mail_accounts WHERE email = ? AND id != ?",
                (email, int(existing["id"])),
            ).fetchone()
            if duplicate:
                raise ValueError("该邮箱已被其他账号配置")
        else:
            existing = conn.execute("SELECT * FROM mail_accounts WHERE email = ?", (email,)).fetchone()
            if existing and str(existing["owner_employee_id"] or "") != owner_employee_id:
                raise ValueError("该邮箱已由其他用户配置")
        ciphertext = encrypt_text(auth_code) if auth_code else str(existing["auth_code_ciphertext"] or "") if existing else ""
        if not ciphertext:
            raise ValueError("请填写客户端授权码")
        if existing:
            conn.execute(
                """
                UPDATE mail_accounts
                SET email = ?, imap_host = ?, imap_port = ?, auth_code_ciphertext = ?,
                    enabled = ?, updated_at = ?
                WHERE id = ? AND owner_employee_id = ?
                """,
                (
                    email,
                    imap_host.strip(),
                    int(imap_port),
                    ciphertext,
                    int(enabled),
                    now,
                    existing["id"],
                    owner_employee_id,
                ),
            )
            return int(existing["id"])
        cursor = conn.execute(
            """
            INSERT INTO mail_accounts (
                email, owner_employee_id, imap_host, imap_port, auth_code_ciphertext, enabled,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                email.strip(),
                owner_employee_id,
                imap_host.strip(),
                int(imap_port),
                ciphertext,
                int(enabled),
                now,
                now,
            ),
        )
        return int(cursor.lastrowid)


def delete_account(account_id: int, *, owner_employee_id: str) -> str:
    """Remove only the current user's mailbox configuration.

    Imported local mail records are deliberately retained; deleting a configured
    account must not remove business history or touch the original mailbox.
    """
    owner_employee_id = str(owner_employee_id or "").strip()
    if not owner_employee_id:
        raise ValueError("未识别当前登录用户")
    with db_cursor() as conn:
        account = conn.execute(
            "SELECT id, email FROM mail_accounts WHERE id = ? AND owner_employee_id = ?",
            (int(account_id), owner_employee_id),
        ).fetchone()
        if not account:
            raise ValueError("邮箱账号不存在或无权删除")
        conn.execute(
            "DELETE FROM mail_accounts WHERE id = ? AND owner_employee_id = ?",
            (int(account_id), owner_employee_id),
        )
    return str(account["email"])


def set_account_fetch_status(account_id: int, status: str, *, at: str | None = None) -> None:
    now = at or now_iso()
    with db_cursor() as conn:
        conn.execute(
            "UPDATE mail_accounts SET last_fetch_at = ?, last_fetch_status = ?, updated_at = ? WHERE id = ?",
            (now, status, now, account_id),
        )


def upsert_message(
    account_id: int,
    *,
    folder: str,
    uid: str,
    message_id: str,
    subject: str,
    sender: str,
    sent_at: str,
    received_at: str,
    body_html: str,
    body_text: str,
    eml_path: str,
    is_order: int,
    fetch_task_id: int = 0,
) -> tuple[int, bool]:
    now = now_iso()
    with db_cursor() as conn:
        existing = conn.execute(
            "SELECT id FROM mail_messages WHERE account_id = ? AND folder = ? AND uid = ?",
            (account_id, folder, uid),
        ).fetchone()
        values = (
            account_id,
            folder,
            uid,
            message_id,
            subject,
            sender,
            sent_at,
            received_at,
            body_html,
            body_text,
            eml_path,
            int(is_order),
            int(fetch_task_id),
            now,
        )
        if existing:
            conn.execute(
                """
                UPDATE mail_messages SET
                    message_id = ?, subject = ?, sender = ?, sent_at = ?, received_at = ?,
                    body_html = ?, body_text = ?, eml_path = ?, is_order = ?,
                    created_at = created_at
                WHERE id = ?
                """,
                (
                    values[3],
                    values[4],
                    values[5],
                    values[6],
                    values[7],
                    values[8],
                    values[9],
                    values[10],
                    values[11],
                    existing["id"],
                ),
            )
            return int(existing["id"]), False
        cursor = conn.execute(
            """
            INSERT INTO mail_messages (
                account_id, folder, uid, message_id, subject, sender, sent_at, received_at,
                body_html, body_text, eml_path, is_order, fetch_task_id, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            values,
        )
        return int(cursor.lastrowid), True


def record_fetch_task_message(fetch_task_id: int, mail_id: int, *, is_new: bool) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO mail_fetch_task_messages
                (fetch_task_id, mail_id, is_new, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (fetch_task_id, mail_id, 1 if is_new else 0, now_iso()),
        )


def get_message(
    mail_id: int, *, owner_employee_id: str | None = None
) -> dict[str, Any] | None:
    with db_cursor() as conn:
        sql = """
            SELECT m.* FROM mail_messages m
            JOIN mail_accounts a ON a.id = m.account_id
            WHERE m.id = ?
        """
        params: list[Any] = [mail_id]
        if owner_employee_id:
            sql += " AND a.owner_employee_id = ?"
            params.append(owner_employee_id)
        row = conn.execute(sql, params).fetchone()
    return _one(row)


def list_messages(
    account_id: int | None = None,
    *,
    limit: int = 50,
    owner_employee_id: str | None = None,
) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        sql = """
            SELECT m.* FROM mail_messages m
            JOIN mail_accounts a ON a.id = m.account_id
        """
        clauses: list[str] = []
        params: list[Any] = []
        if account_id:
            clauses.append("m.account_id = ?")
            params.append(account_id)
        if owner_employee_id:
            clauses.append("a.owner_employee_id = ?")
            params.append(owner_employee_id)
        if clauses:
            sql += " WHERE " + " AND ".join(clauses)
        sql += " ORDER BY m.id DESC LIMIT ?"
        params.append(limit)
        rows = conn.execute(sql, params).fetchall()
    return _rows(rows)


def replace_attachments(mail_id: int, attachments: list[dict[str, Any]]) -> None:
    with db_cursor() as conn:
        conn.execute("DELETE FROM mail_attachments WHERE mail_id = ?", (mail_id,))
        for item in attachments:
            conn.execute(
                """
                INSERT INTO mail_attachments (
                    mail_id, filename, content_type, size_bytes, sha256, stored_path,
                    is_inline, parse_status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    mail_id,
                    item.get("filename", ""),
                    item.get("content_type", ""),
                    int(item.get("size_bytes") or 0),
                    item.get("sha256", ""),
                    item.get("stored_path", ""),
                    int(item.get("is_inline") or 0),
                    item.get("parse_status", ""),
                    now_iso(),
                ),
            )


def list_attachments(mail_id: int) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT * FROM mail_attachments WHERE mail_id = ? ORDER BY is_inline, id",
            (mail_id,),
        ).fetchall()
    return _rows(rows)


def upsert_order_task(
    mail_id: int,
    *,
    customer_code: str = "",
    customer_name: str = "",
    spec: str = "",
    remark: str = "",
    order_number: str = "",
    source_type: str = "",
) -> int:
    now = now_iso()
    spec = str(spec or "").strip()
    order_number = str(order_number or "").strip()
    with db_cursor() as conn:
        existing = conn.execute(
            """
            SELECT * FROM mail_order_tasks
            WHERE mail_id = ? AND spec = ? AND order_number = ?
            """,
            (mail_id, spec, order_number),
        ).fetchone()
        if existing:
            effective_code = str(customer_code or "").strip() or str(existing["customer_code"] or "")
            pending = existing["review_status"] == "pending_review"
            effective_name = str(customer_name or "").strip() if pending else str(existing["customer_name"] or "")
            effective_spec = spec if pending else str(existing["spec"] or "")
            effective_remark = str(remark or "").strip() if pending else str(existing["remark"] or "")
            effective_order = order_number if pending else str(existing["order_number"] or "")
            if not effective_code and effective_name:
                effective_code = resolve_customer_code_by_name(effective_name)
            complete = bool(effective_name and effective_spec)
            field_status = "complete" if complete else "missing"
            review_status = "reviewed" if complete else "pending_review"
            conn.execute(
                """
                UPDATE mail_order_tasks SET
                    customer_code = ?, customer_name = ?, spec = ?, remark = ?,
                    order_number = ?, field_status = ?, review_status = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    effective_code,
                    effective_name,
                    effective_spec,
                    effective_remark,
                    effective_order,
                    field_status,
                    review_status,
                    now,
                    existing["id"],
                ),
            )
            return int(existing["id"])
        effective_code = str(customer_code or "").strip()
        if not effective_code and customer_name:
            effective_code = resolve_customer_code_by_name(customer_name)
        complete = bool(customer_name and spec)
        field_status = "complete" if complete else "missing"
        review_status = "reviewed" if complete else "pending_review"
        cursor = conn.execute(
            """
            INSERT INTO mail_order_tasks (
                mail_id, customer_code, customer_name, spec, remark, order_number,
                source_type, field_status, review_status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                mail_id,
                effective_code,
                customer_name,
                spec,
                remark,
                order_number,
                source_type,
                field_status,
                review_status,
                now,
                now,
            ),
        )
        return int(cursor.lastrowid)


def prune_empty_order_items(mail_id: int) -> None:
    with db_cursor() as conn:
        non_empty = conn.execute(
            "SELECT COUNT(*) AS total FROM mail_order_tasks WHERE mail_id = ? AND spec != ''",
            (mail_id,),
        ).fetchone()["total"]
        if non_empty:
            conn.execute(
                """
                DELETE FROM mail_order_tasks
                WHERE mail_id = ? AND spec = ''
                  AND transcode_status = 'not_started' AND review_status = 'pending_review'
                """,
                (mail_id,),
            )


def get_tasks_by_mail(
    mail_id: int, *, owner_employee_id: str | None = None
) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        sql = """
            SELECT t.*, m.subject, m.sender, m.sent_at, m.received_at
            FROM mail_order_tasks t
            JOIN mail_messages m ON m.id = t.mail_id
            JOIN mail_accounts a ON a.id = m.account_id
            WHERE t.mail_id = ?
        """
        params: list[Any] = [mail_id]
        if owner_employee_id:
            sql += " AND a.owner_employee_id = ?"
            params.append(owner_employee_id)
        sql += " ORDER BY t.id"
        rows = conn.execute(sql, params).fetchall()
    return _rows(rows)


def set_mail_attachment_status(mail_id: int, status: str) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_order_tasks
            SET attachment_parse_status = ?, updated_at = ?
            WHERE mail_id = ?
            """,
            (status, now_iso(), mail_id),
        )


def _refresh_task_status(task_id: int) -> None:
    task = get_order_task(task_id)
    if not task:
        return
    complete = bool(task["customer_name"] and task["spec"])
    field_status = "complete" if complete else "missing"
    review_status = "reviewed" if complete else "pending_review"
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_order_tasks
            SET field_status = ?, review_status = ?, updated_at = ?
            WHERE id = ?
            """,
            (field_status, review_status, now_iso(), task_id),
        )


def get_order_task(
    task_id: int, *, owner_employee_id: str | None = None
) -> dict[str, Any] | None:
    with db_cursor() as conn:
        sql = """
            SELECT t.*, m.subject, m.sender, m.sent_at, m.received_at
            FROM mail_order_tasks t
            JOIN mail_messages m ON m.id = t.mail_id
            JOIN mail_accounts a ON a.id = m.account_id
            WHERE t.id = ?
        """
        params: list[Any] = [task_id]
        if owner_employee_id:
            sql += " AND a.owner_employee_id = ?"
            params.append(owner_employee_id)
        row = conn.execute(sql, params).fetchone()
    return _one(row)


def list_order_tasks(
    *,
    review_status: str | None = None,
    transcode_status: str | None = None,
    limit: int = 100,
    owner_employee_id: str | None = None,
) -> list[dict[str, Any]]:
    sql = """
        SELECT t.*, m.subject, m.sender, m.sent_at, m.received_at
        FROM mail_order_tasks t
        JOIN mail_messages m ON m.id = t.mail_id
        JOIN mail_accounts a ON a.id = m.account_id
    """
    clauses: list[str] = []
    params: list[Any] = []
    if review_status:
        clauses.append("t.review_status = ?")
        params.append(review_status)
    if transcode_status:
        clauses.append("t.transcode_status = ?")
        params.append(transcode_status)
    if owner_employee_id:
        clauses.append("a.owner_employee_id = ?")
        params.append(owner_employee_id)
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " ORDER BY t.id DESC LIMIT ?"
    params.append(limit)
    with db_cursor() as conn:
        rows = conn.execute(sql, params).fetchall()
    return _rows(rows)


def update_order_task(
    task_id: int,
    *,
    customer_code: str | None = None,
    customer_name: str | None = None,
    spec: str | None = None,
    remark: str | None = None,
    order_number: str | None = None,
) -> None:
    current = get_order_task(task_id)
    if not current:
        return
    updates = {
        "customer_code": current["customer_code"] if customer_code is None else customer_code,
        "customer_name": current["customer_name"] if customer_name is None else customer_name,
        "spec": current["spec"] if spec is None else spec,
        "remark": current["remark"] if remark is None else remark,
        "order_number": current["order_number"] if order_number is None else order_number,
    }
    if not str(updates["customer_code"]).strip() and str(updates["customer_name"]).strip():
        updates["customer_code"] = resolve_customer_code_by_name(str(updates["customer_name"]).strip())
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_order_tasks
            SET customer_code = ?, customer_name = ?, spec = ?, remark = ?,
                order_number = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                str(updates["customer_code"]).strip(),
                str(updates["customer_name"]).strip(),
                str(updates["spec"]).strip(),
                str(updates["remark"]).strip(),
                str(updates["order_number"]).strip(),
                now_iso(),
                task_id,
            ),
        )
    _refresh_task_status(task_id)


def update_task_attachment_status(task_id: int, status: str) -> None:
    with db_cursor() as conn:
        conn.execute(
            "UPDATE mail_order_tasks SET attachment_parse_status = ?, updated_at = ? WHERE id = ?",
            (status, now_iso(), task_id),
        )


def update_task_transcode(
    task_id: int,
    *,
    status: str,
    code: str,
    note: str,
    confidence: int,
) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_order_tasks
            SET transcode_status = ?, transcode_code = ?, transcode_note = ?,
                transcode_confidence = ?, updated_at = ?
            WHERE id = ?
            """,
            (status, code, note, confidence, now_iso(), task_id),
        )


def verify_order_task(task_id: int, action: str, employee_id: str) -> None:
    status = "已通过" if action == "approve" else "已驳回"
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_order_tasks
            SET transcode_status = ?, reviewed_by = ?, reviewed_at = ?, updated_at = ?
            WHERE id = ?
            """,
            (status, employee_id, now_iso(), now_iso(), task_id),
        )


def create_transcode_job(task_ids: list[int], input_path: str, created_by: str) -> int:
    now = now_iso()
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            INSERT INTO mail_transcode_jobs (
                task_ids, input_path, status, created_by, created_at
            ) VALUES (?, ?, 'running', ?, ?)
            """,
            (",".join(str(item) for item in task_ids), input_path, created_by, now),
        )
        return int(cursor.lastrowid)


def get_transcode_job(
    job_id: int, *, owner_employee_id: str | None = None
) -> dict[str, Any] | None:
    with db_cursor() as conn:
        sql = "SELECT * FROM mail_transcode_jobs WHERE id = ?"
        params: list[Any] = [job_id]
        if owner_employee_id:
            sql += " AND created_by = ?"
            params.append(owner_employee_id)
        row = conn.execute(sql, params).fetchone()
    return _one(row)


def list_transcode_jobs(
    limit: int = 20, *, owner_employee_id: str | None = None
) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        if owner_employee_id:
            rows = conn.execute(
                "SELECT * FROM mail_transcode_jobs WHERE created_by = ? ORDER BY id DESC LIMIT ?",
                (owner_employee_id, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM mail_transcode_jobs ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
    return _rows(rows)


def complete_transcode_job(job_id: int, *, status: str, success_count: int, fail_count: int) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_transcode_jobs
            SET status = ?, success_count = ?, fail_count = ?, completed_at = ?
            WHERE id = ?
            """,
            (status, success_count, fail_count, now_iso(), job_id),
        )


def create_fetch_log(account_id: int, started_at: str, status: str, message: str) -> int:
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            INSERT INTO mail_fetch_logs (account_id, started_at, status, message)
            VALUES (?, ?, ?, ?)
            """,
            (account_id, started_at, status, message),
        )
        return int(cursor.lastrowid)


def update_fetch_log(log_id: int, *, status: str, message: str, finished_at: str | None = None) -> None:
    with db_cursor() as conn:
        conn.execute(
            "UPDATE mail_fetch_logs SET status = ?, message = ?, finished_at = ? WHERE id = ?",
            (status, message, finished_at or now_iso(), log_id),
        )


def list_fetch_logs(
    limit: int = 20, *, owner_employee_id: str | None = None
) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        if owner_employee_id:
            rows = conn.execute(
                """
                SELECT l.*, a.email FROM mail_fetch_logs l
                JOIN mail_accounts a ON a.id = l.account_id
                WHERE a.owner_employee_id = ? ORDER BY l.id DESC LIMIT ?
                """,
                (owner_employee_id, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM mail_fetch_logs ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
    return _rows(rows)


def create_fetch_task(
    account_id: int,
    *,
    created_by: str = "",
    status: str = "queued",
    message: str = "",
    started_at: str | None = None,
) -> int:
    now = started_at or now_iso()
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            INSERT INTO mail_fetch_tasks (account_id, created_by, status, message, started_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (account_id, created_by, status, message, now),
        )
        return int(cursor.lastrowid)


def start_fetch_task(fetch_task_id: int, *, message: str = "开始抓取") -> None:
    with db_cursor() as conn:
        conn.execute(
            "UPDATE mail_fetch_tasks SET status=?, message=?, started_at=? WHERE id=?",
            ("running", message, now_iso(), fetch_task_id),
        )


def has_active_fetch_task(account_id: int) -> bool:
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT 1 FROM mail_fetch_tasks WHERE account_id=? AND status IN ('queued','running') LIMIT 1",
            (account_id,),
        ).fetchone()
    return bool(row)


def complete_fetch_task(
    fetch_task_id: int,
    *,
    status: str,
    email_count: int,
    new_count: int = 0,
    duplicate_count: int = 0,
    order_count: int,
    message: str,
    completed_at: str | None = None,
) -> None:
    now = completed_at or now_iso()
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE mail_fetch_tasks
            SET status = ?, email_count = ?, new_count = ?, duplicate_count = ?,
                order_count = ?, message = ?, completed_at = ?
            WHERE id = ?
            """,
            (status, email_count, new_count, duplicate_count, order_count, message, now, fetch_task_id),
        )


def get_fetch_task(
    fetch_task_id: int, *, owner_employee_id: str | None = None
) -> dict[str, Any] | None:
    with db_cursor() as conn:
        sql = """
            SELECT f.* FROM mail_fetch_tasks f
            JOIN mail_accounts a ON a.id = f.account_id
            WHERE f.id = ?
        """
        params: list[Any] = [fetch_task_id]
        if owner_employee_id:
            sql += " AND a.owner_employee_id = ?"
            params.append(owner_employee_id)
        row = conn.execute(sql, params).fetchone()
    return _one(row)


def list_fetch_tasks(
    limit: int = 100, *, owner_employee_id: str | None = None, account_id: int | None = None
) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        if owner_employee_id:
            sql = """
                SELECT f.* FROM mail_fetch_tasks f
                JOIN mail_accounts a ON a.id = f.account_id
                WHERE a.owner_employee_id = ?
            """
            params: list[Any] = [owner_employee_id]
            if account_id:
                sql += " AND f.account_id = ?"
                params.append(account_id)
            sql += " ORDER BY f.id DESC LIMIT ?"
            params.append(limit)
            rows = conn.execute(sql, params).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM mail_fetch_tasks ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
    return _rows(rows)


def get_messages_by_fetch_task(
    fetch_task_id: int, *, owner_employee_id: str | None = None
) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        sql = """
            SELECT m.*, x.is_new FROM mail_fetch_task_messages x
            JOIN mail_messages m ON m.id = x.mail_id
            JOIN mail_accounts a ON a.id = m.account_id
            WHERE x.fetch_task_id = ?
        """
        params: list[Any] = [fetch_task_id]
        if owner_employee_id:
            sql += " AND a.owner_employee_id = ?"
            params.append(owner_employee_id)
        sql += " ORDER BY m.id"
        rows = conn.execute(sql, params).fetchall()
    return _rows(rows)


def count_order_tasks_for_mail_ids(mail_ids: list[int]) -> int:
    if not mail_ids:
        return 0
    placeholders = ",".join("?" for _ in mail_ids)
    with db_cursor() as conn:
        row = conn.execute(
            f"SELECT COUNT(*) AS total FROM mail_order_tasks WHERE mail_id IN ({placeholders})",
            mail_ids,
        ).fetchone()
    return int(row["total"])
