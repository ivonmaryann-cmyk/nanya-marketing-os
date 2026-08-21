from __future__ import annotations

import email
import hashlib
import imaplib
import re
import ssl
from datetime import date, datetime, timedelta
from email.header import decode_header
from pathlib import Path
from typing import Any

from ..file_storage import save_automation_file
from ..paths import STORAGE_DIR
from . import mail_store
from .mail_crypto import decrypt_text
from .mail_html_parser import decode_and_simplify_html, extract_order_fields, html_to_text


ORDER_SUBJECT_KEYWORDS = ("采购订单", "订单", "樣品需求", "样品需求", "po", "ga")
ORDER_SENDER_DOMAINS = ("nouyatec.com",)
ATTACHMENT_PARSE_SUFFIXES = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".xlsx", ".xlsm", ".xls"}


def _decode_header_value(value: str | None) -> str:
    if not value:
        return ""
    parts = decode_header(value)
    out: list[str] = []
    for text, charset in parts:
        if isinstance(text, bytes):
            out.append(text.decode(charset or "utf-8", errors="replace"))
        else:
            out.append(str(text))
    return "".join(out)


def _safe_filename(name: str, fallback: str) -> str:
    name = (name or "").strip().replace("/", "_").replace("\\", "_")
    name = re.sub(r'[:*?"<>|]', "_", name)
    return name or fallback


def _sha256(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def is_order_email(subject: str, sender: str) -> bool:
    joined = f"{subject} {sender}".lower()
    if any(keyword in joined for keyword in ORDER_SUBJECT_KEYWORDS):
        return True
    if any(sender.lower().endswith(domain) for domain in ORDER_SENDER_DOMAINS):
        return True
    return False


def _connect(account: dict[str, Any]) -> imaplib.IMAP4_SSL:
    # 邮箱授权码和订单内容均为敏感信息，始终校验 IMAP 服务端证书。
    context = ssl.create_default_context()
    client = imaplib.IMAP4_SSL(
        str(account["imap_host"]),
        int(account["imap_port"]),
        ssl_context=context,
    )
    auth_code = decrypt_text(str(account["auth_code_ciphertext"] or ""))
    client.login(str(account["email"]), auth_code)
    imaplib.Commands["ID"] = ("AUTH", "SELECTED")
    client._simple_command(
        "ID",
        '("name" "nanya-mail-transcode-agent" "version" "0.1" "vendor" "Nanya")',
    )
    return client


def test_imap_connection(account_id: int, *, owner_employee_id: str) -> dict[str, str]:
    """Authenticate and open INBOX read-only without reading or persisting mail."""
    account = mail_store.get_account(account_id, owner_employee_id=owner_employee_id)
    if not account:
        raise ValueError("邮箱账号不存在")
    if not account.get("auth_code_ciphertext"):
        raise ValueError("邮箱账号未配置授权码")

    client = _connect(account)
    try:
        typ, _ = client.select("INBOX", readonly=True)
        if str(typ).upper() != "OK":
            raise ValueError("无法以只读方式打开收件箱")
        return {"message": "连接成功：已使用只读模式验证 IMAP 登录和收件箱访问，未抓取或保存任何邮件。"}
    finally:
        try:
            client.logout()
        except Exception:
            pass


def _body_parts(message: email.message.Message) -> tuple[str, str]:
    html_payload = b""
    html_charset = ""
    plain_payload = b""
    plain_charset = ""
    for part in message.walk():
        if part.get_content_maintype() != "text" or part.get_filename():
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        charset = str(part.get_content_charset() or "utf-8")
        if part.get_content_subtype() == "html" and not html_payload:
            html_payload, html_charset = payload, charset
        elif part.get_content_subtype() == "plain" and not plain_payload:
            plain_payload, plain_charset = payload, charset
    if html_payload:
        html = decode_and_simplify_html(html_payload, html_charset)
        return html, html_to_text(html)
    if plain_payload:
        text = plain_payload.decode(plain_charset or "utf-8", errors="replace")
        return "", text
    return "", ""


def _collect_attachments(message: email.message.Message, base_dir: Path) -> list[dict[str, Any]]:
    attachments: list[dict[str, Any]] = []
    for part in message.walk():
        filename = part.get_filename()
        if not filename:
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        disposition = part.get_content_disposition()
        content_id = part.get("Content-ID")
        is_inline = disposition == "inline" or (disposition is None and bool(content_id))
        sub_dir = base_dir / ("inline_images" if is_inline else "attachments")
        sub_dir.mkdir(parents=True, exist_ok=True)
        clean_name = _safe_filename(_decode_header_value(filename), "attachment")
        target = sub_dir / clean_name
        counter = 1
        while target.exists():
            target = sub_dir / f"{target.stem}_{counter}{target.suffix}"
            counter += 1
        save_automation_file(payload, target)
        suffix = target.suffix.lower()
        parse_status = "inline" if is_inline else ("pending" if suffix in ATTACHMENT_PARSE_SUFFIXES else "ignored")
        attachments.append(
            {
                "filename": clean_name,
                "content_type": part.get_content_type(),
                "size_bytes": len(payload),
                "sha256": _sha256(payload),
                "stored_path": str(target),
                "is_inline": 1 if is_inline else 0,
                "parse_status": parse_status,
            }
        )
    return attachments


def _mail_datetime(message: email.message.Message, header: str) -> str:
    raw = message.get(header)
    if not raw:
        return ""
    try:
        parsed = email.utils.parsedate_to_datetime(raw)
        return parsed.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(raw)


def fetch_latest_order_mails(
    account_id: int,
    limit: int | None = None,
    *,
    created_by: str = "",
    owner_employee_id: str | None = None,
    lookback_days: int = 2,
) -> dict[str, Any]:
    account = mail_store.get_account(account_id, owner_employee_id=owner_employee_id)
    if not account:
        raise ValueError("邮箱账号不存在")
    if not account.get("auth_code_ciphertext"):
        raise ValueError("邮箱账号未配置授权码")

    started_at = datetime.now().isoformat(timespec="seconds")
    log_id = mail_store.create_fetch_log(account_id, started_at, "running", "开始抓取")
    fetch_task_id = mail_store.create_fetch_task(
        account_id,
        created_by=created_by,
        status="running",
        message="开始抓取",
        started_at=started_at,
    )
    client = None
    try:
        client = _connect(account)
        client.select("INBOX", readonly=True)
        # IMAP SINCE includes its day and BEFORE excludes its day.  BODY.PEEK
        # plus readonly=True keeps the mail server's seen/unseen flags unchanged.
        lookback_days = max(1, min(int(lookback_days), 30))
        start_date = (date.today() - timedelta(days=lookback_days - 1)).strftime("%d-%b-%Y")
        end_date = (date.today() + timedelta(days=1)).strftime("%d-%b-%Y")
        typ, data = client.search(None, "SINCE", start_date, "BEFORE", end_date)
        ids = data[0].split()
        matched: list[bytes] = list(reversed(ids))
        if limit is not None:
            matched = matched[:limit]

        fetched = 0
        new_count = 0
        duplicate_count = 0
        mail_ids: list[int] = []
        for num in matched:
            typ, body_data = client.fetch(num, "(UID BODY.PEEK[])")
            if not body_data or not isinstance(body_data[0], tuple):
                continue
            raw = body_data[0][1]
            uid_part = body_data[0][0].decode("utf-8", errors="replace")
            uid_match = re.search(r"UID\s+(\d+)", uid_part, flags=re.IGNORECASE)
            uid = uid_match.group(1) if uid_match else str(num.decode())
            message = email.message_from_bytes(raw)
            subject = _decode_header_value(message.get("Subject"))
            sender = _decode_header_value(message.get("From"))
            order_candidate = is_order_email(subject, sender)
            mail_dir = STORAGE_DIR / "mail_transcode" / str(account_id) / uid
            mail_dir.mkdir(parents=True, exist_ok=True)
            eml_path = mail_dir / "original.eml"
            save_automation_file(raw, eml_path)
            html, text = _body_parts(message)
            fields = extract_order_fields(text, sender) if order_candidate else {}
            mail_id, is_new = mail_store.upsert_message(
                account_id,
                folder="INBOX",
                uid=uid,
                message_id=str(message.get("Message-ID") or ""),
                subject=subject,
                sender=sender,
                sent_at=_mail_datetime(message, "Date"),
                received_at=_mail_datetime(message, "Date"),
                body_html=html,
                body_text=text,
                eml_path=str(eml_path),
                is_order=1,
                fetch_task_id=fetch_task_id,
            )
            mail_store.record_fetch_task_message(fetch_task_id, mail_id, is_new=is_new)
            fetched += 1
            if not is_new:
                duplicate_count += 1
                continue
            new_count += 1
            mail_ids.append(mail_id)
            attachments = _collect_attachments(message, mail_dir)
            mail_store.replace_attachments(mail_id, attachments)
            if order_candidate:
                specs = fields.get("specs") or ([fields.get("spec")] if fields.get("spec") else [""])
                for spec in specs:
                    mail_store.upsert_order_task(
                        mail_id,
                        customer_code="",
                        customer_name=fields.get("customer_name", ""),
                        spec=spec,
                        remark=fields.get("remark", ""),
                        order_number=fields.get("order_number", ""),
                        source_type="html",
                    )
                if specs and specs != [""]:
                    mail_store.prune_empty_order_items(mail_id)
        scope_label = "今天与昨天" if lookback_days == 2 else f"近 {lookback_days} 天"
        message = f"{scope_label}同步完成，本次新增 {new_count} 封邮件"
        order_count = mail_store.count_order_tasks_for_mail_ids(mail_ids)
        mail_store.complete_fetch_task(
            fetch_task_id,
            status="completed",
            email_count=fetched,
            new_count=new_count,
            duplicate_count=duplicate_count,
            order_count=order_count,
            message=message,
        )
        mail_store.update_fetch_log(log_id, status="success", message=message)
        mail_store.set_account_fetch_status(account_id, "success")
        return {
            "fetch_task_id": fetch_task_id,
            "fetched": fetched,
            "new_count": new_count,
            "duplicate_count": duplicate_count,
            "order_count": order_count,
            "message": message,
        }
    except Exception as exc:
        mail_store.complete_fetch_task(
            fetch_task_id,
            status="error",
            email_count=0,
            order_count=0,
            message=str(exc),
        )
        mail_store.update_fetch_log(log_id, status="error", message=str(exc))
        mail_store.set_account_fetch_status(account_id, "error")
        raise
    finally:
        try:
            if client is not None:
                client.logout()
        except Exception:
            pass
