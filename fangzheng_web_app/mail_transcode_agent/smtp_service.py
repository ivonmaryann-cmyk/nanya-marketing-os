from __future__ import annotations

import smtplib
import ssl
import time
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


def _case_sender_config(case_id: int, *, employee_id: str) -> tuple[dict[str, Any], int | None]:
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT a.id AS account_id, t.id AS template_id
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
    return config, (int(row["template_id"]) if row["template_id"] is not None else None)


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
) -> dict[str, Any]:
    """Send an operator-confirmed order reply through the source mailbox SMTP."""
    config, template_id = _case_sender_config(case_id, employee_id=employee_id)
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
    message.set_content(body)

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
