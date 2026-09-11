from __future__ import annotations

import smtplib
import ssl
import time
from html import escape
from email.message import EmailMessage
from email.utils import formataddr, getaddresses, make_msgid, parseaddr
from typing import Any

from ..database import automation_cursor as db_cursor
from ..order_interface_service import record_order_detail_event
from . import mail_store


SMTP_TIMEOUT_SECONDS = 20


def _safe_header(value: str, label: str) -> str:
    value = str(value or "").strip()
    if "\r" in value or "\n" in value:
        raise ValueError(f"{label}不能包含换行符")
    return value


def _addresses(value: str, label: str) -> list[str]:
    value = _safe_header(value, label)
    result: list[str] = []
    for _display_name, address in getaddresses([value]):
        address = str(address or "").strip()
        if not address:
            continue
        if "@" not in address or address.startswith("@") or address.endswith("@"):
            raise ValueError(f"{label}包含无效邮箱地址")
        result.append(address)
    if value and not result:
        raise ValueError(f"{label}包含无效邮箱地址")
    return result


def _require_config(config: dict[str, Any] | None, *, require_enabled: bool) -> dict[str, Any]:
    if not config:
        raise ValueError("未找到该邮箱的 SMTP 配置")
    if not config.get("configured"):
        raise ValueError("请先在“我的 → 邮箱配置”中填写 SMTP 服务器、用户名和客户端授权码")
    if require_enabled and not config.get("enabled"):
        raise ValueError("该邮箱的 SMTP 发信尚未启用")
    return config


def _connect(config: dict[str, Any]):
    host = str(config["host"])
    port = int(config["port"])
    context = ssl.create_default_context()
    if config["security"] == "ssl":
        client = smtplib.SMTP_SSL(host, port, timeout=SMTP_TIMEOUT_SECONDS, context=context)
    else:
        client = smtplib.SMTP(host, port, timeout=SMTP_TIMEOUT_SECONDS)
        client.ehlo()
        client.starttls(context=context)
        client.ehlo()
    client.login(str(config["username"]), str(config["auth_code"]))
    return client


def _close(client: Any) -> None:
    try:
        client.quit()
    except Exception:
        try:
            client.close()
        except Exception:
            pass


def test_smtp_connection(account_id: int, *, owner_employee_id: str) -> dict[str, str]:
    """Authenticate to the configured SMTP server without sending any message."""
    config = _require_config(
        mail_store.get_smtp_config(account_id, owner_employee_id=owner_employee_id),
        require_enabled=False,
    )
    client = None
    try:
        client = _connect(config)
        try:
            client.noop()
        except Exception:
            # Successful TLS/login is enough for providers that do not implement NOOP.
            pass
    except Exception as exc:
        mail_store.set_smtp_test_status(
            account_id, "failed", owner_employee_id=owner_employee_id
        )
        raise ValueError("SMTP 连接失败，请核对服务器、端口、加密方式和客户端授权码") from exc
    finally:
        if client is not None:
            _close(client)
    mail_store.set_smtp_test_status(account_id, "success", owner_employee_id=owner_employee_id)
    return {"message": "SMTP 连接成功：已验证 TLS/登录，未发送任何邮件。"}


def smtp_ready_for_case(case_id: int, *, employee_id: str) -> bool:
    """Whether this mail case has a fully configured, enabled sender account."""
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT a.id
            FROM order_intake_cases c
            JOIN mail_messages m ON m.id = c.mail_id
            JOIN mail_accounts a ON a.id = m.account_id
            WHERE c.id = ? AND c.employee_id = ? AND a.owner_employee_id = ?
            """,
            (int(case_id), employee_id, employee_id),
        ).fetchone()
    if not row:
        return False
    config = mail_store.get_smtp_config(int(row["id"]), owner_employee_id=employee_id)
    return bool(config and config.get("enabled"))


def build_order_reply_draft(case_id: int, *, employee_id: str) -> dict[str, Any]:
    """Return the web-equivalent editable reply draft without sending mail."""
    from ..order_entry_service import get_saved_template
    from ..order_intake_service import get_case

    case = get_case(case_id, employee_id)
    if not case:
        raise ValueError("订单邮件不存在或无权读取")
    _case, template = get_saved_template(case_id, employee_id)
    sender_name, sender_email = parseaddr(str(case.get("sender") or ""))
    subject = str(case.get("subject") or "订单回复").strip()
    if not subject.lower().startswith("re:"):
        subject = f"Re: {subject}"
    header = (template or {}).get("header") or {}
    lines = (template or {}).get("lines") or []
    customer_name = str(case.get("customer_name") or "客户").strip()
    order_no = str(header.get("customer_order_number") or "").strip()
    body_lines = [f"尊敬的{customer_name}：", "", "您好！", "您的订单我司已收到并完成内部处理。"]
    if order_no:
        body_lines.append(f"客户订单号：{order_no}")
    if lines:
        body_lines.append(f"订单明细：共 {len(lines)} 项")
    body_lines.extend(["", "如需补充交期或其他信息，请直接回复本邮件。", "", "此致", "南亚营销自动化平台"])
    return {
        "case_id": int(case_id),
        "to": sender_email or str(case.get("sender") or "").strip(),
        "recipient_name": sender_name,
        "cc": "",
        "subject": subject,
        "body": "\n".join(body_lines),
        "template_available": bool(template),
        "smtp_ready": smtp_ready_for_case(case_id, employee_id=employee_id),
        "send_supported": True,
        "sent": False,
    }


def _case_sender_config(
    case_id: int, *, employee_id: str
) -> tuple[dict[str, Any], int | None, dict[str, Any]]:
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT a.id AS account_id, t.id AS template_id,
                   m.message_id, m.sender, m.subject, m.sent_at, m.received_at,
                   m.body_text, m.body_html
            FROM order_intake_cases c
            JOIN mail_messages m ON m.id = c.mail_id
            JOIN mail_accounts a ON a.id = m.account_id
            LEFT JOIN order_entry_templates t ON t.case_id = c.id AND t.employee_id = c.employee_id
            WHERE c.id = ? AND c.employee_id = ? AND a.owner_employee_id = ?
            """,
            (int(case_id), employee_id, employee_id),
        ).fetchone()
    if not row:
        raise ValueError("订单邮件不存在或无权发送回复")
    config = _require_config(
        mail_store.get_smtp_config(int(row["account_id"]), owner_employee_id=employee_id),
        require_enabled=True,
    )
    source_mail = {
        key: row[key]
        for key in (
            "message_id", "sender", "subject", "sent_at", "received_at", "body_text", "body_html"
        )
    }
    return (
        config,
        int(row["template_id"]) if row["template_id"] is not None else None,
        source_mail,
    )


def _quoted_original_text(source_mail: dict[str, Any]) -> str:
    original = str(source_mail.get("body_text") or "").strip()
    if not original and source_mail.get("body_html"):
        from .mail_html_parser import html_to_text

        original = html_to_text(str(source_mail["body_html"]))
    return "\n".join(
        [
            "----- 原邮件 -----",
            f"发件人：{source_mail.get('sender') or '未提供'}",
            f"发送时间：{source_mail.get('sent_at') or source_mail.get('received_at') or '未提供'}",
            f"主题：{source_mail.get('subject') or '（无主题）'}",
            "",
            original or "（原邮件正文为空）",
        ]
    )


def _reply_html(body: str, source_mail: dict[str, Any]) -> str:
    from .mail_html_parser import safe_display_html

    reply = escape(body).replace("\n", "<br>")
    original = safe_display_html(
        str(source_mail.get("body_html") or ""), str(source_mail.get("body_text") or "")
    )
    sender = escape(str(source_mail.get("sender") or "未提供"))
    sent_at = escape(
        str(source_mail.get("sent_at") or source_mail.get("received_at") or "未提供")
    )
    subject = escape(str(source_mail.get("subject") or "（无主题）"))
    return (
        f"<div>{reply}</div><br><div style=\"color:#666\">----- 原邮件 -----<br>"
        f"发件人：{sender}<br>发送时间：{sent_at}<br>主题：{subject}</div>"
        f"<blockquote style=\"margin:12px 0 0;padding-left:12px;border-left:2px solid #ccc\">"
        f"{original}</blockquote>"
    )


def _record_send_event(
    *,
    case_id: int,
    template_id: int | None,
    employee_id: str,
    event_type: str,
    title: str,
    detail: dict[str, Any],
) -> None:
    with db_cursor() as conn:
        record_order_detail_event(
            conn,
            case_id=case_id,
            template_id=template_id,
            employee_id=employee_id,
            event_type=event_type,
            title=title,
            detail=detail,
            operated_by=employee_id,
        )


def send_order_reply(
    case_id: int,
    *,
    employee_id: str,
    to: str,
    cc: str,
    subject: str,
    body: str,
    body_html: str | None = None,
) -> dict[str, Any]:
    """Send an operator-confirmed order reply through the source mailbox SMTP."""
    config, template_id, source_mail = _case_sender_config(case_id, employee_id=employee_id)
    recipients = _addresses(to, "收件人")
    cc_recipients = _addresses(cc, "抄送")
    subject = _safe_header(subject, "主题")
    body = str(body or "").strip()
    if not recipients or not subject or not body:
        raise ValueError("请填写收件人、主题和邮件正文后再发送")

    message = EmailMessage()
    sender_email = str(config.get("email") or config["username"])
    sender_name = str(config.get("sender_name") or "").strip()
    message["From"] = formataddr((sender_name, sender_email)) if sender_name else sender_email
    message["To"] = ", ".join(recipients)
    if cc_recipients:
        message["Cc"] = ", ".join(cc_recipients)
    message["Subject"] = subject
    message["Message-ID"] = make_msgid(domain=sender_email.split("@")[-1])
    source_message_id = _safe_header(str(source_mail.get("message_id") or ""), "原邮件标识")
    if source_message_id:
        message["In-Reply-To"] = source_message_id
        message["References"] = source_message_id
    if body_html is not None:
        from .mail_html_parser import safe_display_html, html_to_text
        cleaned = safe_display_html(body_html)
        message.set_content(html_to_text(cleaned))
        message.add_alternative(cleaned, subtype="html")
    else:
        message.set_content(f"{body}\n\n{_quoted_original_text(source_mail)}")
        message.add_alternative(_reply_html(body, source_mail), subtype="html")

    started = time.monotonic()
    client = None
    try:
        client = _connect(config)
        client.send_message(message, from_addr=sender_email, to_addrs=recipients + cc_recipients)
    except Exception as exc:
        _record_send_event(
            case_id=case_id,
            template_id=template_id,
            employee_id=employee_id,
            event_type="order_reply_send_failed",
            title="订单回复邮件发送失败",
            detail={
                "from": sender_email,
                "to": recipients,
                "cc": cc_recipients,
                "subject": subject,
                "duration_ms": int((time.monotonic() - started) * 1000),
                "error_type": type(exc).__name__,
            },
        )
        raise ValueError("邮件发送失败，请检查 SMTP 配置或稍后重试") from exc
    finally:
        if client is not None:
            _close(client)

    result = {
        "message_id": str(message["Message-ID"]),
        "from": sender_email,
        "to": recipients,
        "cc": cc_recipients,
        "subject": subject,
        "duration_ms": int((time.monotonic() - started) * 1000),
        "body_html": message.get_body(preferencelist=("html",)).get_content(),
    }
    _record_send_event(
        case_id=case_id,
        template_id=template_id,
        employee_id=employee_id,
        event_type="order_reply_sent",
        title="订单回复邮件已发送",
        detail=result,
    )
    return result
