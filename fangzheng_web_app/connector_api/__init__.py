"""Machine-facing JSON and MCP facade for Marketing Automation.

The facade calls existing service functions, never the HTML/session routes.
"""

from __future__ import annotations

import hmac
import json
import os
import re
import uuid
from typing import Any

from flask import Blueprint, Response, jsonify, request

from ..mail_transcode_agent import mail_fetch_service, mail_store
from ..order_interface_service import (
    build_domestic_order_entry,
    build_material_query,
    get_interface_config,
    prepare_domestic_order_entry,
)
from ..order_intake_service import get_case, list_cases
from ..mail_transcode_agent.smtp_service import build_order_reply_draft


bp = Blueprint("connector_api", __name__, url_prefix="/api/connector/v1")
DEFAULT_EMPLOYEE_ID = "23582"
MCP_PROTOCOL_VERSION = "2025-03-26"
_SECRET_KEY = re.compile(r"(auth|password|secret|cipher|cookie|token)", re.I)


def _authenticate() -> None:
    """Require a connector-specific token; never reuse mail/Nyeos credentials."""
    configured = os.getenv("CONNECTOR_API_TOKEN", "")
    supplied = request.headers.get("Authorization", "")
    token = supplied[7:].strip() if supplied.lower().startswith("bearer ") else ""
    if not configured:
        raise RuntimeError("连接器尚未配置 CONNECTOR_API_TOKEN")
    if not token or not hmac.compare_digest(token, configured):
        raise PermissionError("连接器认证失败")


def _context() -> tuple[str, str]:
    employee_id = (request.headers.get("X-Nanya-Employee-Id") or os.getenv(
        "CONNECTOR_DEFAULT_EMPLOYEE_ID", DEFAULT_EMPLOYEE_ID
    )).strip()
    # The desktop supplies the signed-in employee number as business context.
    # It is not an MCP admission whitelist: individual tools decide whether
    # there are matching records or an applicable business permission.
    if not employee_id:
        raise PermissionError("当前连接器缺少员工工号")
    correlation_id = (request.headers.get("X-Correlation-Id") or uuid.uuid4().hex).strip()[:128]
    return employee_id, correlation_id


def _safe(value: Any) -> Any:
    if isinstance(value, dict):
        result = {}
        for key, item in value.items():
            text = str(key).lower()
            if _SECRET_KEY.search(text) or text.endswith("_path") or text in {"eml_path", "body_html", "display_html"}:
                continue
            result[str(key)] = _safe(item)
        return result
    if isinstance(value, list):
        return [_safe(item) for item in value]
    return value


def _as_int(value: Any, field: str, *, maximum: int = 1000) -> int:
    if isinstance(value, bool):
        raise ValueError(f"{field} 必须是整数")
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field} 必须是整数") from exc
    if result < 1 or result > maximum:
        raise ValueError(f"{field} 必须在 1 到 {maximum} 之间")
    return result


def _strict_object(value: Any, allowed: set[str]) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, dict) or any(key not in allowed for key in value):
        raise ValueError("工具参数不符合 JSON Schema")
    return value


def _business_account(employee_id: str, account_id: int | None = None) -> dict[str, Any] | None:
    accounts = [item for item in mail_store.list_accounts(owner_employee_id=employee_id) if item.get("enabled")]
    if account_id is not None:
        return next((item for item in accounts if int(item.get("id") or 0) == account_id), None)
    return accounts[0] if accounts else None


def _interface_status(key: str) -> dict[str, Any]:
    config = get_interface_config(key) or {}
    enabled, mode = bool(config.get("enabled")), str(config.get("mode") or "disabled")
    return {
        "enabled": enabled, "mode": mode,
        # A configured endpoint is not evidence that NYEOS has accepted a
        # request.  Health must never promote it to a verified real service.
        "readiness": "disabled" if not enabled else ("mock_ready" if mode == "mock" else "real_configured_unverified"),
        "last_test": "not_checked_by_connector",
    }


def connector_health(employee_id: str, correlation_id: str) -> dict[str, Any]:
    accounts = mail_store.list_accounts(owner_employee_id=employee_id)
    account = _business_account(employee_id)
    mail_ready = bool(account and account.get("imap_host") and account.get("auth_code_ciphertext"))
    smtp = mail_store.smtp_public_config(account) if account else {}
    return {
        "connector": "nanya-marketing-automation", "version": "v1",
        "employee_id": employee_id, "correlation_id": correlation_id,
        "mail": {"enabled": mail_ready, "mode": "configured" if mail_ready else "not_configured",
                 "readiness": "ready" if mail_ready else "no_mailbox", "account_count": len(accounts)},
        "material_query": _interface_status("material_batch_query"),
        "domestic_entry": _interface_status("domestic_order_entry"),
        "smtp": {"enabled": bool(smtp.get("enabled")),
                 "mode": "configured" if smtp.get("configured") else "not_configured",
                 "readiness": "ready" if smtp.get("enabled") else ("configured_not_verified" if smtp.get("configured") else "not_configured")},
        "capabilities": {
            "mail_sync": "available" if mail_ready else "unavailable",
            "material_query": "available" if _interface_status("material_batch_query")["enabled"] else "unavailable",
            "prepare_entry": "available",
            "submit_entry": "available" if _interface_status("domestic_order_entry")["enabled"] else "unavailable",
            "reply_draft": "available",
            # Existing SMTP service can send, but it does not expose an
            # idempotency key. The connector intentionally omits send_reply.
            "send_reply": "not_exposed_no_idempotency_key",
        },
    }


def _public_job(job: dict[str, Any] | None, correlation_id: str) -> dict[str, Any] | None:
    if not job:
        return None
    keys = {"id", "status", "email_count", "new_count", "duplicate_count", "order_count", "message", "created_at", "completed_at", "account_id"}
    return {**_safe({key: job.get(key) for key in keys if key in job}), "correlation_id": correlation_id}


def _public_case(case: dict[str, Any]) -> dict[str, Any]:
    result = _safe(case)
    if isinstance(result.get("body_text"), str):
        result["body_text"] = result["body_text"][:12000]
    return result


def _public_case_summary(case: dict[str, Any]) -> dict[str, Any]:
    """Return the decision-relevant fields for a list view.

    A case can contain full mail bodies, attachment extraction output and rule
    traces.  Returning those for every row makes a normal task list too large
    for an agent turn.  Details remain available through get_case(case_id).
    """
    fields = {
        "id", "employee_id", "mail_id", "action_type", "status",
        "customer_code", "customer_name", "order_number", "order_version",
        "parent_order_number", "workflow_stage", "customer_match_status",
        "source_document_status", "mapping_status", "erp_prepare_status",
        "routing_source", "routing_reason", "routing_state", "handling_note",
        "created_at", "updated_at", "received_at", "subject",
    }
    return _safe({key: value for key, value in case.items() if key in fields})


def _list_cases(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"limit", "action_type", "fetch_task_id", "target_date"})
    limit = _as_int(args.get("limit", 50), "limit", maximum=200)
    fetch_task_id = _as_int(args["fetch_task_id"], "fetch_task_id") if args.get("fetch_task_id") is not None else None
    values = list_cases(employee_id, target_date=str(args.get("target_date") or "").strip() or None,
                        action_type=str(args.get("action_type") or "all"), fetch_task_id=fetch_task_id, prepare=True)[:limit]
    return {"correlation_id": correlation_id, "count": len(values), "cases": [_public_case_summary(item) for item in values]}


def _get_case(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"case_id"})
    # Case IDs are database primary keys and are commonly six digits.  The
    # generic 1–1000 helper limit is for page size / line-number inputs only.
    case = get_case(_as_int(args.get("case_id"), "case_id", maximum=2_147_483_647), employee_id)
    if not case:
        raise LookupError("订单案件不存在或不属于当前工号")
    return {"correlation_id": correlation_id, "case": _public_case(case)}


def _sync_orders(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"account_id", "lookback_days", "limit"})
    account_id = _as_int(args["account_id"], "account_id") if args.get("account_id") is not None else None
    account = _business_account(employee_id, account_id)
    if not account:
        raise LookupError("当前工号没有可同步的业务邮箱")
    result = mail_fetch_service.queue_latest_order_mails(
        int(account["id"]), created_by=employee_id, owner_employee_id=employee_id,
        lookback_days=_as_int(args.get("lookback_days", 2), "lookback_days", maximum=30),
        limit=_as_int(args["limit"], "limit", maximum=500) if args.get("limit") is not None else None,
    )
    return {"correlation_id": correlation_id, "job": {"id": int(result["fetch_task_id"]), "status": "queued"},
            "message": str(result.get("message") or "已创建邮件同步任务")}


def _material_query(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"case_id", "line_nos"})
    case_id = _as_int(args.get("case_id"), "case_id")
    line_nos = args.get("line_nos")
    if line_nos is not None:
        if not isinstance(line_nos, list) or len(line_nos) > 500:
            raise ValueError("line_nos 必须是最多 500 项的整数数组")
        line_nos = {_as_int(item, "line_nos") for item in line_nos}
    # Existing service chooses its configured Mock/Real path, applies all
    # template/state validation, and records its own call/audit records.
    result = build_material_query(case_id, employee_id, employee_id, line_nos=line_nos)
    return {"correlation_id": correlation_id, "result": _safe(result)}


def _prepare_entry(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"case_id"})
    # This new service-layer method shares the exact payload builder with
    # build_domestic_order_entry but makes no HTTP call or order-state change.
    result = prepare_domestic_order_entry(_as_int(args.get("case_id"), "case_id"), employee_id)
    return {"correlation_id": correlation_id, "preview": _safe(result)}


def _reply_draft(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"case_id"})
    return {"correlation_id": correlation_id, "draft": _safe(build_order_reply_draft(_as_int(args.get("case_id"), "case_id"), employee_id=employee_id))}


def _submit_entry(employee_id: str, arguments: Any, correlation_id: str) -> dict[str, Any]:
    args = _strict_object(arguments, {"case_id", "confirm"})
    if args.get("confirm") is not True:
        raise ValueError("提交录单必须显式传 confirm=true")
    # Existing build service owns validation, Mock/Real choice, durable call
    # logs and the successful-submit idempotency check.
    result = build_domestic_order_entry(_as_int(args.get("case_id"), "case_id"), employee_id, employee_id)
    return {"correlation_id": correlation_id, "result": _safe(result)}


TOOLS = [
    {"name": "marketing.mail.sync_orders", "description": "异步同步当前员工的订单邮件。",
     "inputSchema": {"type": "object", "additionalProperties": False, "properties": {"account_id": {"type": "integer", "minimum": 1}, "lookback_days": {"type": "integer", "minimum": 1, "maximum": 30, "default": 2}, "limit": {"type": "integer", "minimum": 1, "maximum": 500}}}},
    {"name": "marketing.job.get", "description": "读取邮件同步任务状态。",
     "inputSchema": {"type": "object", "additionalProperties": False, "required": ["job_id"], "properties": {"job_id": {"type": "integer", "minimum": 1}}}},
    {"name": "marketing.order.list_cases", "description": "列出当前员工可见的订单案件。",
     "inputSchema": {"type": "object", "additionalProperties": False, "properties": {"limit": {"type": "integer", "minimum": 1, "maximum": 200, "default": 50}, "action_type": {"type": "string"}, "fetch_task_id": {"type": "integer", "minimum": 1}, "target_date": {"type": "string"}}}},
    {"name": "marketing.order.get_case", "description": "读取订单案件和提取的订单明细。",
     "inputSchema": {"type": "object", "additionalProperties": False, "required": ["case_id"], "properties": {"case_id": {"type": "integer", "minimum": 1}}}},
    {"name": "marketing.material.query", "description": "对已生成录单模板的订单执行既有 NYEOS 料号查询（按配置走 Mock 或 Real）。",
     "inputSchema": {"type": "object", "additionalProperties": False, "required": ["case_id"], "properties": {"case_id": {"type": "integer", "minimum": 1}, "line_nos": {"type": "array", "items": {"type": "integer", "minimum": 1}, "maxItems": 500}}}},
    {"name": "marketing.order.prepare_entry", "description": "校验订单并生成内销录单请求载荷预览；不提交订单、不写入接口日志。",
     "inputSchema": {"type": "object", "additionalProperties": False, "required": ["case_id"], "properties": {"case_id": {"type": "integer", "minimum": 1}}}},
    {"name": "marketing.order.reply_draft", "description": "根据订单案件生成可编辑的客户回复邮件草稿；不会发送邮件。",
     "inputSchema": {"type": "object", "additionalProperties": False, "required": ["case_id"], "properties": {"case_id": {"type": "integer", "minimum": 1}}}},
    {"name": "marketing.order.submit_entry", "description": "显式确认后提交内销录单；复用已有 Mock/Real、幂等校验和审计。",
     "inputSchema": {"type": "object", "additionalProperties": False, "required": ["case_id", "confirm"], "properties": {"case_id": {"type": "integer", "minimum": 1}, "confirm": {"type": "boolean", "const": True}}}},
]


def _call_tool(employee_id: str, correlation_id: str, name: str, arguments: Any) -> dict[str, Any]:
    if name == "marketing.mail.sync_orders":
        return _sync_orders(employee_id, arguments, correlation_id)
    if name == "marketing.job.get":
        args = _strict_object(arguments, {"job_id"})
        job = mail_store.get_fetch_task(_as_int(args.get("job_id"), "job_id"), owner_employee_id=employee_id)
        if not job:
            raise LookupError("邮件同步任务不存在或不属于当前工号")
        return {"correlation_id": correlation_id, "job": _public_job(job, correlation_id)}
    if name == "marketing.order.list_cases":
        return _list_cases(employee_id, arguments, correlation_id)
    if name == "marketing.order.get_case":
        return _get_case(employee_id, arguments, correlation_id)
    if name == "marketing.material.query":
        return _material_query(employee_id, arguments, correlation_id)
    if name == "marketing.order.prepare_entry":
        return _prepare_entry(employee_id, arguments, correlation_id)
    if name == "marketing.order.reply_draft":
        return _reply_draft(employee_id, arguments, correlation_id)
    if name == "marketing.order.submit_entry":
        return _submit_entry(employee_id, arguments, correlation_id)
    raise LookupError("未找到连接器工具")


def _error(message: str, status: int) -> tuple[Response, int]:
    return jsonify({"error": {"message": message}}), status


@bp.get("/health")
def health() -> Response | tuple[Response, int]:
    try:
        _authenticate(); employee_id, correlation_id = _context()
        return jsonify(connector_health(employee_id, correlation_id))
    except PermissionError as exc:
        return _error(str(exc), 403)
    except RuntimeError as exc:
        return _error(str(exc), 503)


@bp.get("/jobs/<int:job_id>")
def job(job_id: int) -> Response | tuple[Response, int]:
    try:
        _authenticate(); employee_id, correlation_id = _context()
        result = mail_store.get_fetch_task(job_id, owner_employee_id=employee_id)
        return jsonify({"correlation_id": correlation_id, "job": _public_job(result, correlation_id)}) if result else _error("任务不存在", 404)
    except PermissionError as exc:
        return _error(str(exc), 403)
    except RuntimeError as exc:
        return _error(str(exc), 503)


@bp.get("/orders/cases")
def cases() -> Response | tuple[Response, int]:
    try:
        _authenticate(); employee_id, correlation_id = _context()
        return jsonify(_list_cases(employee_id, request.args.to_dict(), correlation_id))
    except PermissionError as exc:
        return _error(str(exc), 403)
    except RuntimeError as exc:
        return _error(str(exc), 503)
    except ValueError as exc:
        return _error(str(exc), 400)


@bp.get("/orders/cases/<int:case_id>")
def case(case_id: int) -> Response | tuple[Response, int]:
    try:
        _authenticate(); employee_id, correlation_id = _context()
        return jsonify(_get_case(employee_id, {"case_id": case_id}, correlation_id))
    except PermissionError as exc:
        return _error(str(exc), 403)
    except RuntimeError as exc:
        return _error(str(exc), 503)
    except LookupError as exc:
        return _error(str(exc), 404)


@bp.post("/mcp")
def mcp() -> Response:
    payload = request.get_json(silent=True)
    request_id = payload.get("id") if isinstance(payload, dict) else None
    if not isinstance(payload, dict) or payload.get("jsonrpc") != "2.0" or not isinstance(payload.get("method"), str):
        return jsonify({"jsonrpc": "2.0", "id": request_id, "error": {"code": -32600, "message": "Invalid Request"}})
    try:
        _authenticate(); employee_id, correlation_id = _context(); method = payload["method"]
        if method == "initialize":
            result: dict[str, Any] = {"protocolVersion": MCP_PROTOCOL_VERSION, "capabilities": {"tools": {}}, "serverInfo": {"name": "nanya-marketing-automation", "version": "1.0.0"}}
        elif method == "tools/list":
            result = {"tools": TOOLS}
        elif method == "tools/call":
            params = _strict_object(payload.get("params"), {"name", "arguments"})
            if not isinstance(params.get("name"), str):
                raise ValueError("tools/call 需要 name")
            outcome = _call_tool(employee_id, correlation_id, params["name"], params.get("arguments"))
            result = {"content": [{"type": "text", "text": json.dumps(_safe(outcome), ensure_ascii=False)}], "isError": False}
        elif method == "notifications/initialized":
            return Response(status=202)
        else:
            return jsonify({"jsonrpc": "2.0", "id": request_id, "error": {"code": -32601, "message": "Method not found"}})
        response = jsonify({"jsonrpc": "2.0", "id": request_id, "result": result})
        response.headers["Mcp-Session-Id"] = correlation_id
        return response
    except (LookupError, ValueError) as exc:
        return jsonify({"jsonrpc": "2.0", "id": request_id, "result": {"content": [{"type": "text", "text": json.dumps({"error": str(exc)}, ensure_ascii=False)}], "isError": True}})
    except PermissionError as exc:
        return jsonify({"jsonrpc": "2.0", "id": request_id, "error": {"code": -32001, "message": str(exc)}})
    except RuntimeError as exc:
        return jsonify({"jsonrpc": "2.0", "id": request_id, "error": {"code": -32002, "message": str(exc)}})
    except Exception:
        return jsonify({"jsonrpc": "2.0", "id": request_id, "result": {"content": [{"type": "text", "text": json.dumps({"error": "连接器调用失败，请查看服务端日志"}, ensure_ascii=False)}], "isError": True}})
