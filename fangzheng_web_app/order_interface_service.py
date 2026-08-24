from __future__ import annotations

import json
import uuid
from datetime import datetime, timedelta
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen
from typing import Any

from .database import automation_cursor as db_cursor
from .db import utcnow


INTERFACE_DEFAULTS = {
    "material_batch_query": {
        "display_name": "批量料号查询",
        "description": "按订单明细批量查询料号；返回结果仅作为建议，不覆盖人工填写内容。",
        "method": "POST",
        "base_url": "https://mock.nouya.local/material/batch-query",
        "port": 443,
        "path": "",
        "timeout_seconds": 8,
        "request_mapping": {
            "items[].line_no": "模板明细.项次",
            "items[].customer_part_no": "模板明细.客户产品编号",
            "items[].customer_spec": "模板明细.客户规格",
        },
        "response_mapping": {
            "items[].factory_part_no": "料号查询建议.产品编号",
            "items[].product_name": "料号查询建议.品名",
            "items[].matched_spec": "料号查询建议.匹配规格",
            "items[].status": "接口交互记录.状态",
            "items[].message": "接口交互记录.提示",
        },
    },
    "domestic_order_entry": {
        "display_name": "内销录单",
        "description": "人工确认订单内容后提交内销录单；当前先维护 Mock 配置。",
        "method": "POST",
        "base_url": "https://mock.nouya.local/sales/internal-entry",
        "port": 443,
        "path": "",
        "timeout_seconds": 10,
        "request_mapping": {"header": "内销模板.表头", "items": "内销模板.明细行"},
        "response_mapping": {"status": "接口交互记录.状态", "message": "接口交互记录.提示"},
    },
}

MOCK_SCENARIOS = {
    "success": {"label": "成功", "description": "返回可匹配的 Mock 结果。"},
    "not_found": {"label": "未找到", "description": "指定明细未找到对应料号。"},
    "business_error": {"label": "业务错误", "description": "模拟接口业务校验失败。"},
    "timeout": {"label": "超时", "description": "模拟请求超时并记录失败原因。"},
}


def _json(value: Any, fallback: Any) -> Any:
    try:
        return json.loads(value or "")
    except (TypeError, ValueError):
        return fallback


def _row(row: Any) -> dict[str, Any]:
    return dict(row) if row else {}


def _endpoint_url(config: dict[str, Any]) -> str:
    base_url = str(config.get("base_url") or "").strip().rstrip("/")
    path = str(config.get("path") or "").strip()
    return base_url if not path else f"{base_url}/{path.lstrip('/')}"


def _valid_endpoint_url(value: str) -> str:
    endpoint_url = str(value or "").strip()
    parsed = urlparse(endpoint_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("请求地址必须是完整的 http 或 https URL")
    return endpoint_url


def ensure_interface_configs(operated_by: str = "system") -> None:
    """Seed the two approved profiles exactly once without replacing edits."""
    now = utcnow()
    with db_cursor() as conn:
        for interface_key, defaults in INTERFACE_DEFAULTS.items():
            existing = conn.execute(
                "SELECT id FROM order_interface_configs WHERE interface_key=?",
                (interface_key,),
            ).fetchone()
            if existing:
                continue
            conn.execute(
                """INSERT INTO order_interface_configs
                   (interface_key,display_name,description,enabled,mode,method,base_url,port,path,
                    timeout_seconds,request_mapping_json,response_mapping_json,mock_scenarios_json,
                    config_version,created_by,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    interface_key, defaults["display_name"], defaults["description"], 1, "mock",
                    defaults["method"], defaults["base_url"], defaults["port"], defaults["path"],
                    defaults["timeout_seconds"], json.dumps(defaults["request_mapping"], ensure_ascii=False),
                    json.dumps(defaults["response_mapping"], ensure_ascii=False),
                    json.dumps(MOCK_SCENARIOS, ensure_ascii=False), 1, operated_by, now, now,
                ),
            )


def list_interface_configs() -> list[dict[str, Any]]:
    ensure_interface_configs()
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT * FROM order_interface_configs ORDER BY interface_key"
        ).fetchall()
    result = []
    for row in rows:
        item = _row(row)
        item["request_mapping"] = _json(item.pop("request_mapping_json", ""), {})
        item["response_mapping"] = _json(item.pop("response_mapping_json", ""), {})
        item["mock_scenarios"] = _json(item.pop("mock_scenarios_json", ""), {})
        item["endpoint_url"] = _endpoint_url(item)
        result.append(item)
    return result


def get_interface_config(interface_key: str) -> dict[str, Any] | None:
    ensure_interface_configs()
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT * FROM order_interface_configs WHERE interface_key=?",
            (interface_key,),
        ).fetchone()
    if not row:
        return None
    item = _row(row)
    item["request_mapping"] = _json(item.pop("request_mapping_json", ""), {})
    item["response_mapping"] = _json(item.pop("response_mapping_json", ""), {})
    item["mock_scenarios"] = _json(item.pop("mock_scenarios_json", ""), {})
    item["endpoint_url"] = _endpoint_url(item)
    return item


def _mapping_from_text(value: str, label: str) -> dict[str, str]:
    try:
        payload = json.loads(value or "{}")
    except ValueError as exc:
        raise ValueError(f"{label}必须是合法 JSON 对象") from exc
    if not isinstance(payload, dict):
        raise ValueError(f"{label}必须是 JSON 对象")
    return {str(key): str(item) for key, item in payload.items()}


def save_interface_config(interface_key: str, payload: dict[str, Any], operated_by: str) -> dict[str, Any]:
    if interface_key not in INTERFACE_DEFAULTS:
        raise ValueError("不支持的接口配置")
    ensure_interface_configs(operated_by)
    method = str(payload.get("method") or "POST").upper()
    if method not in {"GET", "POST", "PUT", "PATCH"}:
        raise ValueError("请求方法不支持")
    mode = str(payload.get("mode") or "mock")
    if mode not in {"mock", "real"}:
        raise ValueError("运行模式不支持")
    request_mapping = _mapping_from_text(str(payload.get("request_mapping") or "{}"), "入参映射")
    response_mapping = _mapping_from_text(str(payload.get("response_mapping") or "{}"), "出参映射")
    mock_scenarios = _mapping_from_text(str(payload.get("mock_scenarios") or "{}"), "Mock 场景")
    now = utcnow()
    with db_cursor() as conn:
        previous = conn.execute(
            "SELECT * FROM order_interface_configs WHERE interface_key=?", (interface_key,)
        ).fetchone()
        if not previous:
            raise ValueError("接口配置不存在")
        before = _row(previous)
        endpoint_url = _valid_endpoint_url(payload.get("endpoint_url") or _endpoint_url(before))
        next_version = int(before["config_version"] or 0) + 1
        conn.execute(
            """UPDATE order_interface_configs
               SET display_name=?,description=?,enabled=?,mode=?,method=?,base_url=?,port=?,path=?,
                   timeout_seconds=?,request_mapping_json=?,response_mapping_json=?,mock_scenarios_json=?,
                   config_version=?,updated_at=?
               WHERE interface_key=?""",
            (
                str(payload.get("display_name") or "").strip() or INTERFACE_DEFAULTS[interface_key]["display_name"],
                str(payload.get("description") or "").strip(),
                1, mode, method, endpoint_url, int(before["port"] or 443), "",
                int(before["timeout_seconds"] or 8), json.dumps(request_mapping, ensure_ascii=False),
                json.dumps(response_mapping, ensure_ascii=False), json.dumps(mock_scenarios, ensure_ascii=False),
                next_version, now, interface_key,
            ),
        )
        current = conn.execute(
            "SELECT * FROM order_interface_configs WHERE interface_key=?", (interface_key,)
        ).fetchone()
        conn.execute(
            """INSERT INTO order_interface_config_versions
               (interface_config_id,config_version,before_json,after_json,operated_by,created_at)
               VALUES (?,?,?,?,?,?)""",
            (int(current["id"]), next_version, json.dumps(before, ensure_ascii=False),
             json.dumps(_row(current), ensure_ascii=False), operated_by, now),
        )
    return get_interface_config(interface_key) or {}


def test_interface_config(payload: dict[str, Any]) -> dict[str, Any]:
    """Test the current form values without saving them.

    The maintenance UI no longer exposes a timeout setting. A fixed 15-second
    server-side protection prevents a bad third-party endpoint from holding a
    web request indefinitely.
    """
    mode = str(payload.get("mode") or "mock")
    interface_key = str(payload.get("interface_key") or "")
    endpoint_url = _valid_endpoint_url(str(payload.get("endpoint_url") or ""))
    method = str(payload.get("method") or "POST").upper()
    if method not in {"GET", "POST", "PUT", "PATCH"}:
        raise ValueError("请求方法不支持")
    request_body = {
        "items": [{"line_no": 1, "customer_part_no": "TEST-001", "customer_spec": "Mock 测试规格"}]
    } if interface_key == "material_batch_query" else {
        "header": {"customer_order_number": "TEST-ORDER-001"},
        "items": [{"line_no": 1, "customer_part_no": "TEST-001"}],
    }
    if mode == "mock":
        return {
            "ok": True, "mode": "mock", "status_code": 200, "duration_ms": 0,
            "endpoint_url": endpoint_url,
            "request": request_body,
            "response": {"items": [{"line_no": 1, "status": "matched", "message": "Mock 测试成功"}]},
        }
    body = json.dumps(request_body, ensure_ascii=False).encode("utf-8")
    request = Request(
        endpoint_url,
        data=body if method != "GET" else None,
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        method=method,
    )
    try:
        import time
        started = time.monotonic()
        with urlopen(request, timeout=15) as response:
            raw = response.read(256 * 1024).decode("utf-8", errors="replace")
            status_code = int(response.status)
        try:
            response_body: Any = json.loads(raw) if raw else {}
        except ValueError:
            response_body = raw
        return {
            "ok": 200 <= status_code < 300, "mode": "real", "status_code": status_code,
            "duration_ms": int((time.monotonic() - started) * 1000),
            "endpoint_url": endpoint_url, "request": request_body, "response": response_body,
        }
    except HTTPError as exc:
        return {"ok": False, "mode": "real", "status_code": exc.code, "duration_ms": 0,
                "endpoint_url": endpoint_url, "request": request_body,
                "response": {}, "error": f"接口返回 HTTP {exc.code}"}
    except (URLError, TimeoutError, OSError) as exc:
        return {"ok": False, "mode": "real", "status_code": None, "duration_ms": 0,
                "endpoint_url": endpoint_url, "request": request_body,
                "response": {}, "error": f"接口请求失败：{str(exc)[:160]}"}


def record_order_detail_event(
    conn: Any,
    *,
    case_id: int,
    template_id: int | None,
    employee_id: str,
    event_type: str,
    title: str,
    detail: dict[str, Any] | None = None,
    operated_by: str = "",
) -> None:
    conn.execute(
        """INSERT INTO order_entry_detail_events
           (case_id,template_id,employee_id,event_type,title,detail_json,operated_by,created_at)
           VALUES (?,?,?,?,?,?,?,?)""",
        (
            case_id, template_id, employee_id, event_type, title,
            json.dumps(detail or {}, ensure_ascii=False), operated_by or employee_id, utcnow(),
        ),
    )


def _case_template_id(case_id: int, employee_id: str) -> int | None:
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT id FROM order_entry_templates WHERE case_id=? AND employee_id=?",
            (case_id, employee_id),
        ).fetchone()
    return int(row["id"]) if row else None


_AUDIT_FIELD_LABELS = {
    "product_code": "产品编号", "product_name": "品名", "customer_product_code": "客户产品编号",
    "customer_spec": "客户规格", "customer_spec_match": "客户规格匹配", "quantity": "数量",
}
_AUDIT_SOURCE_LABELS = {
    "template_saved": "人工保存", "template_extracted": "首次提取", "template_reextracted": "重新提取",
    "material_query_mock": "料号查询接口", "material_created_callback": "料号创建回调",
    "domestic_order_entry_mock": "内销录单接口",
}


def _audit_time(value: Any) -> str:
    """Stored audit values are UTC; render the user-facing China time to seconds."""
    raw = str(value or "")
    try:
        timestamp = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        if timestamp.tzinfo is not None:
            timestamp = timestamp.replace(tzinfo=None)
        return (timestamp + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return raw.replace("T", " ")


def _redact_audit_payload(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            str(key): "***" if any(token in str(key).lower() for token in ("password", "token", "secret", "authorization", "auth_code"))
            else _redact_audit_payload(item)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [_redact_audit_payload(item) for item in value]
    return value


def _call_summary(call: dict[str, Any]) -> str:
    response = call.get("response") or {}
    items = response.get("items") if isinstance(response, dict) else []
    if isinstance(items, list) and items:
        labels = {"matched": "已回填", "creating": "创建中", "failed": "失败", "not_found": "未找到"}
        counts: dict[str, int] = {}
        for item in items:
            status = labels.get(str((item or {}).get("status") or ""), "已处理")
            counts[status] = counts.get(status, 0) + 1
        return f"处理 {len(items)} 行：" + "、".join(f"{count} 行{label}" for label, count in counts.items())
    if isinstance(response, dict) and response.get("message"):
        return str(response["message"])
    return str(call.get("error_message") or "已记录本次接口交互")


def _event_summary(item: dict[str, Any]) -> str:
    detail = item.get("detail") or {}
    if detail.get("items"):
        statuses: dict[str, int] = {}
        for result in detail["items"]:
            status = str((result or {}).get("status") or "已处理")
            statuses[status] = statuses.get(status, 0) + 1
        status_labels = {"matched": "已回填", "creating": "创建中", "failed": "失败"}
        return "；".join(f"{count} 行{status_labels.get(status, status)}" for status, count in statuses.items())
    if detail.get("changes"):
        return f"涉及 {len(detail['changes'])} 个字段变更"
    if detail.get("line_count") is not None:
        return f"涉及 {detail['line_count']} 条订单明细"
    if detail.get("error_message"):
        return str(detail["error_message"])
    return "已生成可追溯记录"


def get_order_detail_records(case_id: int, employee_id: str) -> dict[str, list[dict[str, Any]]]:
    template_id = _case_template_id(case_id, employee_id)
    with db_cursor() as conn:
        events = conn.execute(
            """SELECT * FROM order_entry_detail_events
               WHERE case_id=? AND employee_id=? ORDER BY id DESC LIMIT 100""",
            (case_id, employee_id),
        ).fetchall()
        calls = conn.execute(
            """SELECT * FROM order_interface_call_logs
               WHERE case_id=? AND employee_id=? ORDER BY id DESC LIMIT 100""",
            (case_id, employee_id),
        ).fetchall()
        configs = conn.execute("SELECT * FROM order_interface_configs").fetchall()
        config_versions = conn.execute("SELECT * FROM order_interface_config_versions").fetchall()
    event_rows = []
    changes = []
    for row in events:
        item = _row(row)
        item["detail"] = _json(item.pop("detail_json", ""), {})
        item["trace_id"] = f"E-{item['id']}"
        item["occurred_at"] = _audit_time(item.get("created_at"))
        item["source_label"] = _AUDIT_SOURCE_LABELS.get(item.get("event_type"), "系统操作")
        item["summary"] = _event_summary(item)
        event_rows.append(item)
        for raw_change in item["detail"].get("changes") or []:
            change = dict(raw_change or {})
            field = str(change.get("field") or "")
            scope = str(change.get("scope") or (f"第 {change['line_no']} 行" if change.get("line_no") else "订单模板"))
            changes.append({
                **change,
                "field": _AUDIT_FIELD_LABELS.get(field, field or "订单数据"),
                "scope": scope,
                "occurred_at": item["occurred_at"],
                "operated_by": item.get("operated_by") or "系统",
                "source_label": item["source_label"],
                "event_id": int(item["id"]),
                "event_trace_id": item["trace_id"],
                "call_id": item["detail"].get("call_id"),
            })
    call_rows = []
    config_by_id = {int(row["id"]): _row(row) for row in configs}
    config_snapshots: dict[tuple[int, int], dict[str, Any]] = {}
    for row in config_versions:
        version = _row(row)
        config_id, version_no = int(version["interface_config_id"]), int(version["config_version"])
        config_snapshots[(config_id, version_no)] = _json(version.get("after_json"), {})
        config_snapshots.setdefault((config_id, version_no - 1), _json(version.get("before_json"), {}))
    for row in calls:
        item = _row(row)
        for key in ("request_json", "response_json"):
            item[key[:-5]] = _json(item.pop(key, ""), {})
        item["trace_id"] = f"I-{item['id']}"
        item["occurred_at"] = _audit_time(item.get("created_at"))
        item["interface_label"] = "批量料号查询" if item.get("interface_key") == "material_batch_query" else "内销录单"
        item["mode_label"] = "Mock" if item.get("is_mock") else "真实接口"
        item["outcome_label"] = "成功" if item.get("status") == "success" else "失败"
        item["summary"] = _call_summary(item)
        item["request"] = _redact_audit_payload(item["request"])
        item["response"] = _redact_audit_payload(item["response"])
        config_id = item.get("interface_config_id")
        snapshot = config_snapshots.get((int(config_id), int(item.get("config_version") or 1)), {}) if config_id else {}
        config = snapshot or config_by_id.get(int(config_id), {}) if config_id else {}
        item["endpoint_url"] = _endpoint_url(config) if config else "—"
        item["method"] = str(config.get("method") or "POST") if config else "—"
        call_rows.append(item)
    return {"template_id": template_id, "events": event_rows, "changes": changes[:100], "calls": call_rows}


MATERIAL_STATUS_LABELS = {
    "pending": "待查询", "waiting_callback": "创建料号中", "requerying": "正在获取新料号",
    "resolved": "已回填", "manual_resolved": "人工已填写", "failed": "查询异常",
}


def _material_request_item(line_no: int, values: dict[str, Any]) -> dict[str, Any]:
    """The three approved keys are deliberately the only material-query input."""
    return {
        "line_no": line_no,
        "customer_product_code": str(values.get("customer_product_code") or ""),
        "customer_spec": str(values.get("customer_spec") or ""),
        "customer_spec_match": str(values.get("customer_spec_match") or ""),
    }


def _insert_call_log(conn: Any, *, case_id: int, template_id: int, employee_id: str,
                     config: dict[str, Any], status: str, request_payload: dict[str, Any],
                     response_payload: dict[str, Any], triggered_by: str, error_message: str = "") -> int:
    cursor = conn.execute(
        """INSERT INTO order_interface_call_logs
           (case_id,template_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
            status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (case_id, template_id, employee_id, int(config["id"]), "material_batch_query",
         int(config["config_version"]), 1, status, 200 if status == "success" else 422, 1,
         json.dumps(request_payload, ensure_ascii=False), json.dumps(response_payload, ensure_ascii=False),
         error_message, triggered_by, utcnow()),
    )
    return int(cursor.lastrowid)


def _upsert_resolution_task(conn: Any, *, case_id: int, template_id: int, employee_id: str,
                            line_no: int, status: str, input_item: dict[str, Any], call_id: int,
                            result: dict[str, Any], correlation_id: str | None = None) -> dict[str, Any]:
    existing = conn.execute(
        "SELECT * FROM order_material_resolution_tasks WHERE template_id=? AND line_no=?",
        (template_id, line_no),
    ).fetchone()
    now = utcnow()
    correlation_id = correlation_id or (str(existing["correlation_id"]) if existing else uuid.uuid4().hex)
    if existing:
        conn.execute(
            """UPDATE order_material_resolution_tasks
               SET status=?,input_json=?,result_json=?,last_call_log_id=?,updated_at=? WHERE id=?""",
            (status, json.dumps(input_item, ensure_ascii=False), json.dumps(result, ensure_ascii=False),
             call_id, now, int(existing["id"])),
        )
        task_id = int(existing["id"])
    else:
        cursor = conn.execute(
            """INSERT INTO order_material_resolution_tasks
               (case_id,template_id,employee_id,line_no,status,correlation_id,input_json,result_json,last_call_log_id,created_at,updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (case_id, template_id, employee_id, line_no, status, correlation_id,
             json.dumps(input_item, ensure_ascii=False), json.dumps(result, ensure_ascii=False), call_id, now, now),
        )
        task_id = int(cursor.lastrowid)
    return {"id": task_id, "correlation_id": correlation_id, "status": status}


def _backfill_material_line(conn: Any, *, template_id: int, line_no: int, values: dict[str, Any],
                            factory_part_no: str, product_name: str, correlation_id: str) -> list[dict[str, Any]]:
    """Fill blanks and replace the interface's temporary creation message."""
    before = dict(values)
    sources_row = conn.execute(
        "SELECT sources_json FROM order_entry_template_lines WHERE template_id=? AND line_no=?", (template_id, line_no)
    ).fetchone()
    sources = _json(sources_row["sources_json"] if sources_row else "", {})
    changes: list[dict[str, Any]] = []
    for field, value in (("product_code", factory_part_no), ("product_name", product_name)):
        if value and str(values.get(field) or "").strip() in {"", "创建料号中"}:
            values[field] = value
            sources[field] = {"label": "料号查询接口（Mock）", "reference": f"关联号 {correlation_id}"}
            changes.append({"field": field, "before": before.get(field, ""), "after": value, "line_no": line_no})
    if changes:
        conn.execute(
            """UPDATE order_entry_template_lines SET values_json=?,sources_json=?,updated_at=?
               WHERE template_id=? AND line_no=?""",
            (json.dumps(values, ensure_ascii=False), json.dumps(sources, ensure_ascii=False), utcnow(), template_id, line_no),
        )
        conn.execute("UPDATE order_entry_templates SET updated_at=? WHERE id=?", (utcnow(), template_id))
    return changes


def _mock_material_response(item: dict[str, Any], *, callback_requery: bool = False) -> dict[str, Any]:
    line_no = int(item["line_no"])
    code = item["customer_product_code"]
    if not code:
        return {"line_no": line_no, "status": "failed", "message": "缺少客户产品编号，无法查询料号。"}
    if callback_requery or line_no % 3 == 1:
        return {"line_no": line_no, "status": "matched", "factory_part_no": f"MOCK-{code[-6:]}",
                "product_name": f"Mock 品名 {code[-4:]}", "matched_spec": item["customer_spec_match"] or item["customer_spec"],
                "message": "Mock 已命中并回填产品编号、品名。"}
    if line_no % 3 == 2:
        return {"line_no": line_no, "status": "creating", "factory_part_no": "创建料号中", "product_name": "创建料号中",
                "external_task_id": f"MOCK-CREATE-{line_no}", "message": "Mock 未找到料号，已发起创建；等待对方回调。"}
    return {"line_no": line_no, "status": "failed", "message": "Mock 查询异常：请检查客户产品编号、客户规格和规格匹配。"}


def build_material_query_mock(case_id: int, employee_id: str, triggered_by: str, scenario: str = "success",
                              *, line_nos: set[int] | None = None,
                              callback_requery: bool = False) -> dict[str, Any]:
    """Run the production-shaped Mock: hit, creation callback, then automatic backfill."""
    if scenario not in MOCK_SCENARIOS:
        raise ValueError("不支持的 Mock 场景")
    config = get_interface_config("material_batch_query")
    if not config or not config.get("enabled"):
        raise ValueError("批量料号查询接口未启用")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成录单模板")
    with db_cursor() as conn:
        rows = conn.execute("SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no", (template_id,)).fetchall()
        selected = [row for row in rows if line_nos is None or int(row["line_no"]) in line_nos]
        request_items = [_material_request_item(int(row["line_no"]), _json(row["values_json"], {})) for row in selected]
        if scenario in {"business_error", "timeout"}:
            message = "Mock 业务错误：订单明细校验未通过。" if scenario == "business_error" else "Mock 超时：请求未返回。"
            call_id = _insert_call_log(conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                                        status="failed", request_payload={"items": request_items}, response_payload={"items": [], "message": message},
                                        triggered_by=triggered_by, error_message=message)
            record_order_detail_event(conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                                      event_type="material_query_mock", title="批量料号查询（Mock）失败",
                                      detail={"call_id": call_id, "error_message": message}, operated_by=triggered_by)
            return {"call_id": call_id, "items": [], "status": "failed", "scheduled_task_ids": []}
        response_items = [_mock_material_response(item, callback_requery=callback_requery) for item in request_items]
        call_id = _insert_call_log(conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                                    status="success", request_payload={"items": request_items}, response_payload={"items": response_items}, triggered_by=triggered_by)
        all_changes: list[dict[str, Any]] = []
        for row, item, response in zip(selected, request_items, response_items):
            values = _json(row["values_json"], {})
            status = response["status"]
            task_status = "resolved" if status == "matched" else ("waiting_callback" if status == "creating" else "failed")
            task = _upsert_resolution_task(conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                                            line_no=item["line_no"], status=task_status, input_item=item,
                                            call_id=call_id, result=response)
            if status == "matched":
                changes = _backfill_material_line(conn, template_id=template_id, line_no=item["line_no"], values=values,
                                                   factory_part_no=response["factory_part_no"], product_name=response["product_name"],
                                                   correlation_id=task["correlation_id"])
                all_changes.extend(changes)
                conn.execute("UPDATE order_material_resolution_tasks SET resolved_at=?,updated_at=? WHERE id=?", (utcnow(), utcnow(), task["id"]))
            elif status == "creating":
                # “创建料号中” is an interface-returned value, so it belongs in
                # the two editable cells rather than as a second visual status.
                changes = _backfill_material_line(conn, template_id=template_id, line_no=item["line_no"], values=values,
                                                   factory_part_no=response["factory_part_no"], product_name=response["product_name"],
                                                   correlation_id=task["correlation_id"])
                all_changes.extend(changes)
                conn.execute("UPDATE order_material_resolution_tasks SET external_task_id=?,updated_at=? WHERE id=?",
                             (response["external_task_id"], utcnow(), task["id"]))
            conn.execute(
                """INSERT INTO order_material_query_suggestions
                   (call_log_id,template_id,line_no,factory_part_no,product_name,matched_spec,status,message,created_at)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (call_id, template_id, item["line_no"], response.get("factory_part_no", ""), response.get("product_name", ""),
                 response.get("matched_spec", ""), status, response["message"], utcnow()),
            )
        record_order_detail_event(conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                                  event_type="material_query_mock", title="批量料号查询（Mock）完成",
                                  detail={"call_id": call_id, "changes": all_changes, "items": response_items}, operated_by=triggered_by)
    return {"call_id": call_id, "items": response_items, "status": "success"}


def process_material_created_callback(correlation_id: str, *, product_code: str, product_name: str, source: str = "external") -> dict[str, Any]:
    """Apply the partner callback directly; it already contains the completed material values."""
    with db_cursor() as conn:
        task = conn.execute("SELECT * FROM order_material_resolution_tasks WHERE correlation_id=?", (correlation_id,)).fetchone()
        if not task:
            raise ValueError("未找到对应的料号创建任务")
        task = _row(task)
        if task["status"] in {"resolved", "manual_resolved"}:
            return {"status": task["status"], "message": "该料号任务已完成。"}
        config = get_interface_config("material_batch_query")
        if not config:
            raise ValueError("批量料号查询接口配置不存在")
        input_item = _json(task.get("input_json"), {})
        response_item = {
            "line_no": int(task["line_no"]), "status": "matched", "factory_part_no": product_code,
            "product_name": product_name, "message": "外部系统已回调料号和品名。",
        }
        call_id = _insert_call_log(conn, case_id=int(task["case_id"]), template_id=int(task["template_id"]),
                                   employee_id=str(task["employee_id"]), config=config, status="success",
                                   request_payload={"callback": {"correlation_id": correlation_id}},
                                   response_payload={"items": [response_item]}, triggered_by=source)
        line = conn.execute("SELECT values_json FROM order_entry_template_lines WHERE template_id=? AND line_no=?", (int(task["template_id"]), int(task["line_no"]))).fetchone()
        if not line:
            raise ValueError("回调对应的订单明细不存在")
        changes = _backfill_material_line(conn, template_id=int(task["template_id"]), line_no=int(task["line_no"]),
                                           values=_json(line["values_json"], {}), factory_part_no=product_code,
                                           product_name=product_name, correlation_id=correlation_id)
        conn.execute(
            """UPDATE order_material_resolution_tasks
               SET status='resolved',result_json=?,last_call_log_id=?,callback_received_at=?,resolved_at=?,updated_at=? WHERE id=?""",
            (json.dumps(response_item, ensure_ascii=False), call_id, utcnow(), utcnow(), utcnow(), int(task["id"])),
        )
        record_order_detail_event(conn, case_id=int(task["case_id"]), template_id=int(task["template_id"]), employee_id=str(task["employee_id"]),
                                  event_type="material_created_callback", title="收到料号创建完成回调，已回填料号和品名",
                                  detail={"correlation_id": correlation_id, "line_no": int(task["line_no"]), "source": source, "changes": changes}, operated_by=source)
    return {"status": "resolved", "call_id": call_id, "changes": changes}


def get_material_resolution_states(case_id: int, employee_id: str) -> dict[str, Any]:
    with db_cursor() as conn:
        rows = conn.execute("SELECT line_no,status,correlation_id,result_json,updated_at FROM order_material_resolution_tasks WHERE case_id=? AND employee_id=? ORDER BY line_no", (case_id, employee_id)).fetchall()
    items = []
    for row in rows:
        item = _row(row)
        item["result"] = _json(item.pop("result_json", ""), {})
        item["label"] = MATERIAL_STATUS_LABELS.get(item["status"], item["status"])
        items.append(item)
    return {"items": items, "pending": any(item["status"] in {"waiting_callback", "requerying"} for item in items)}


def validate_domestic_order_entry(case_id: int, employee_id: str) -> list[str]:
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        return ["请先生成并保存录单模板"]
    with db_cursor() as conn:
        lines = conn.execute("SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no", (template_id,)).fetchall()
        tasks = conn.execute("SELECT line_no,status FROM order_material_resolution_tasks WHERE template_id=?", (template_id,)).fetchall()
    states = {int(row["line_no"]): str(row["status"]) for row in tasks}
    issues = []
    for row in lines:
        line_no, values = int(row["line_no"]), _json(row["values_json"], {})
        if states.get(line_no) in {"waiting_callback", "requerying", "failed"}:
            issues.append(f"第 {line_no} 行料号{MATERIAL_STATUS_LABELS[states[line_no]]}")
        elif not str(values.get("product_code") or "").strip():
            issues.append(f"第 {line_no} 行未填写产品编号")
    return issues


def build_domestic_order_entry_mock(case_id: int, employee_id: str, triggered_by: str) -> dict[str, Any]:
    """Record a manual domestic-order submission Mock without changing the template."""
    issues = validate_domestic_order_entry(case_id, employee_id)
    if issues:
        raise ValueError("暂不能提交录单：" + "；".join(issues))
    config = get_interface_config("domestic_order_entry")
    if not config:
        raise ValueError("内销录单接口配置不存在")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成并保存录单模板")
    with db_cursor() as conn:
        template = conn.execute(
            "SELECT header_json,current_version FROM order_entry_templates WHERE id=?", (template_id,)
        ).fetchone()
        lines = conn.execute(
            "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
            (template_id,),
        ).fetchall()
        request_payload = {
            "header": _json(template["header_json"], {}),
            "items": [{"line_no": int(row["line_no"]), **_json(row["values_json"], {})} for row in lines],
        }
        response_payload = {
            "status": "accepted", "entry_no": f"MOCK-SO-{case_id}",
            "message": "Mock 录单成功；真实接口接入后将返回实际单号。",
        }
        now = utcnow()
        cursor = conn.execute(
            """INSERT INTO order_interface_call_logs
               (case_id,template_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
                status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                case_id, template_id, employee_id, int(config["id"]), "domestic_order_entry",
                int(config["config_version"]), 1, "success", 200, 1,
                json.dumps(request_payload, ensure_ascii=False), json.dumps(response_payload, ensure_ascii=False),
                "", triggered_by, now,
            ),
        )
        call_id = int(cursor.lastrowid)
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="domestic_order_entry_mock", title="提交内销录单（Mock）完成",
            detail={"call_id": call_id, "line_count": len(lines), "entry_no": response_payload["entry_no"]},
            operated_by=triggered_by,
        )
    return {"call_id": call_id, "entry_no": response_payload["entry_no"], "status": "success"}
