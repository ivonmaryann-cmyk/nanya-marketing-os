from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from ..paths import STORAGE_DIR
from ..transcode_agent_service import calculate_transcode_agent_quote
from . import mail_store
from .mail_attachment_parser import parse_attachment


INPUT_ROOT = STORAGE_DIR / "mail_transcode" / "inputs"


def parse_attachments_for_task(task_id: int) -> dict[str, Any]:
    task = mail_store.get_order_task(task_id)
    if not task:
        raise ValueError("订单任务不存在")
    mail_id = int(task["mail_id"])
    attachments = mail_store.list_attachments(mail_id)
    rows: list[dict[str, Any]] = []
    errors: list[str] = []
    for attachment in attachments:
        if attachment["is_inline"] or attachment["parse_status"] == "ignored":
            continue
        path = Path(attachment["stored_path"])
        if not path.exists():
            errors.append(f"{attachment['filename']} 文件不存在")
            continue
        try:
            parsed = parse_attachment(path)
            rows.extend(parsed)
        except Exception as exc:
            errors.append(f"{attachment['filename']}: {exc}")
    base_order = str(task["order_number"] or "").strip()
    base_remark = str(task["remark"] or "").strip()
    base_customer = str(task["customer_name"] or "").strip()
    base_code = str(task["customer_code"] or "").strip()
    for row in rows:
        spec = str(row.get("spec") or "").strip()
        if not spec:
            continue
        mail_store.upsert_order_task(
            mail_id,
            customer_code=base_code,
            customer_name=base_customer or str(row.get("customer") or "").strip(),
            spec=spec,
            remark=str(row.get("remark") or "").strip() or base_remark,
            order_number=str(row.get("order_number") or "").strip() or base_order,
            source_type="attachment",
        )
    if rows:
        mail_store.prune_empty_order_items(mail_id)
    status = "parsed" if rows and not errors else ("error" if errors else "empty")
    mail_store.set_mail_attachment_status(mail_id, status)
    return {"rows": len(rows), "errors": errors, "status": status}


def build_transcode_input(task_ids: list[int], output_path: Path) -> int:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "转码需求表"
    sheet.append(["客户代码", "客户简称", "客户规格", "订单备注"])
    count = 0
    for task_id in task_ids:
        task = mail_store.get_order_task(task_id)
        if not task or task["review_status"] != "reviewed":
            continue
        sheet.append(
            [
                task["customer_code"],
                task["customer_name"],
                task["spec"],
                task["remark"],
            ]
        )
        count += 1
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return count


def run_transcode_for_tasks(task_ids: list[int], employee_id: str) -> int:
    valid_ids = [
        task_id
        for task_id in task_ids
        if mail_store.get_order_task(task_id) is not None
    ]
    if not valid_ids:
        raise ValueError("没有可执行的订单任务")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_path = INPUT_ROOT / f"转码输入_{timestamp}.xlsx"
    row_count = build_transcode_input(valid_ids, input_path)
    if row_count <= 0:
        raise ValueError("没有已核实且字段完整的订单，无法生成转码输入")
    job_id = mail_store.create_transcode_job(valid_ids, str(input_path), employee_id)
    success = 0
    failed = 0
    for task_id in valid_ids:
        task = mail_store.get_order_task(task_id)
        if not task or task["review_status"] != "reviewed":
            failed += 1
            continue
        result = calculate_transcode_agent_quote(
            str(task["spec"] or ""),
            customer=str(task["customer_name"] or ""),
            customer_code=str(task["customer_code"] or ""),
            order_remark=str(task["remark"] or ""),
            employee_id=employee_id,
        )
        quote_status = str(result.get("status") or "失败")
        code = str(result.get("result") or result.get("candidate_code") or "").strip()
        if quote_status == "成功" and code:
            transcode_status = "待核实"
            success += 1
        elif quote_status == "待确认" and code:
            transcode_status = "待核实"
            success += 1
        else:
            transcode_status = "失败"
            failed += 1
        mail_store.update_task_transcode(
            task_id,
            status=transcode_status,
            code=code,
            note=str(result.get("note") or ""),
            confidence=int(result.get("confidence") or 0),
        )
    mail_store.complete_transcode_job(job_id, status="completed", success_count=success, fail_count=failed)
    return job_id
