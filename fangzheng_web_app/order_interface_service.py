from __future__ import annotations

import json
import re
import time
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
        "description": "调用 NYEOS 料号查询/编制接口；先查客户料号，未命中时由对方系统解析客户规格。",
        "method": "POST",
        "base_url": "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/pe/queryMaterial",
        "port": 7030,
        "path": "",
        "timeout_seconds": 15,
        "request_mapping": {
            "customerCode": "模板表头.账款客户编号（必填）",
            "operatorCode": "当前登录账号（员工工号，必填）",
            "materialInfoList[].categoryCode": "模板明细.产品类型（PP=698，基板=718，必填）",
            "materialInfoList[].customerMaterialNo": "模板明细.客户产品编号（必填）",
            "materialInfoList[].customerSpec": "模板明细.客户规格（必填）",
            "materialInfoList[].newProductName": "暂不映射，发送空字符串（选填）",
            "materialInfoList[].oldProductName": "模板明细.品名（选填）",
        },
        "response_mapping": {
            "code": "接口交互记录.业务状态码（200=处理完成，999=失败）",
            "msg": "接口交互记录.提示",
            "pera01": "料号查询建议.客户料号编制作业单号",
            "newFlag": "料号查询建议.是否新建（Y=新建，N=未新建）",
            "source": "料号查询建议.来源（PARSE=规格解析）",
            "errors[]": "接口交互记录.逐行失败信息",
            "hitMaterialList[].peag01": "料号查询建议.产品编号",
            "hitMaterialList[].peag08": "料号查询建议.品名",
            "hitMaterialList[].peag06": "料号查询建议.销售品名规格",
            "hitMaterialList[].peag09": "料号查询建议.旧品名",
            "hitMaterialList[].scca05": "料号查询建议.客户规格",
            "hitMaterialList[].scca03": "料号查询建议.客户产品编号",
        },
    },
    "domestic_order_entry": {
        "display_name": "生成订单",
        "description": "将已确认的录单模板写入 SCTO 中间表并生成订单。",
        "method": "POST",
        "base_url": "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/sc/saveSctoAndGenerateOrder",
        "port": 7030,
        "path": "",
        "timeout_seconds": 15,
        "request_mapping": {
            "sctoDataList[].customerCode": "模板表头.账款客户编号（必填）",
            "sctoDataList[].orderType": "模板表头.单别（必填）",
            "sctoDataList[].operator": "当前登录账号（员工工号，必填）",
            "sctoDataList[].quantity": "模板明细.数量（必填）",
            "sctoDataList[].taxPrice": "模板明细.单价（与税前单价至少一项必填）",
            "sctoDataList[].untaxedPrice": "模板明细.税前单价（与单价至少一项必填）",
            "sctoDataList[].materialCode": "模板明细.产品编号（料号查询结果，必填）",
            "sctoDataList[].lineNumber": "模板明细.项次（必填）",
            "sctoDataList[].demandDate": "模板明细.出货日期（必填）",
            "sctoDataList[].orderNumber": "模板表头.客户订单号（必填）",
            "sctoDataList[].lineId": "模板明细.客户订单序号（选填）",
            "sctoDataList[].lineRemark": "模板明细.备注（选填）",
            "sctoDataList[].taxType": "模板表头.税种（选填）",
            "sctoDataList[].materialName": "模板明细.品名（选填）",
            "sctoDataList[].spec": "模板明细.客户规格匹配；为空时取客户规格（选填）",
        },
        "response_mapping": {
            "code": "接口交互记录.业务状态码",
            "msg": "接口交互记录.提示",
            "data.successCount": "接口交互记录.成功数量",
            "data.failCount": "接口交互记录.失败数量",
            "data.data[].orderNumber": "接口交互记录.客户订单号",
            "data.data[].sctaCode": "接口交互记录.生成订单号",
            "data.data[].status": "接口交互记录.生成状态（success/fail）",
            "data.data[].message": "接口交互记录.失败原因",
        },
    },
}

LEGACY_MATERIAL_DEFAULT = {
    "description": "按订单明细批量查询料号；返回结果仅作为建议，不覆盖人工填写内容。",
    "base_url": "https://mock.nouya.local/material/batch-query",
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
}

LEGACY_DOMESTIC_DEFAULT = {
    "description": "人工确认订单内容后提交内销录单；当前先维护 Mock 配置。",
    "base_url": "https://mock.nouya.local/sales/internal-entry",
    "request_mapping": {"header": "内销模板.表头", "items": "内销模板.明细行"},
    "response_mapping": {"status": "接口交互记录.状态", "message": "接口交互记录.提示"},
}

INTERFACE_MAINTENANCE_NOTES = {
    "material_batch_query": [
        "当前地址是 NYEOS 测试环境；正式环境地址确认后只需修改“请求地址”。",
        "产品类型必须转换为接口编码：PP 使用 698，基板使用 718。",
        "customerSpec 只发送模板中的“客户规格”，不发送“客户规格匹配”。",
        "接口未命中料号时可能返回 pera01，并在对方系统创建客户料号编制作业，请勿用随意数据测试。",
        "保存后业务页会按运行模式执行：Mock 走模拟流程，真实接口会请求当前地址。",
    ],
    "domestic_order_entry": [
        "当前地址是 NYEOS 测试环境；正式环境地址确认后只需修改“请求地址”。",
        "materialCode 使用料号查询后选定的产品编号，不发送原始客户产品编号。",
        "保存后业务页会按运行模式执行：Mock 走模拟流程，真实接口会生成订单。",
    ],
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
                "SELECT * FROM order_interface_configs WHERE interface_key=?",
                (interface_key,),
            ).fetchone()
            if existing:
                if interface_key == "material_batch_query" and _is_untouched_legacy_material_config(existing):
                    _upgrade_legacy_material_config(conn, existing, operated_by, now)
                    existing = conn.execute(
                        "SELECT * FROM order_interface_configs WHERE interface_key=?", (interface_key,)
                    ).fetchone()
                if interface_key == "material_batch_query":
                    _upgrade_material_customer_spec_source(conn, existing, operated_by, now)
                elif interface_key == "domestic_order_entry" and _is_untouched_legacy_domestic_config(existing):
                    _upgrade_legacy_domestic_config(conn, existing, operated_by, now)
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


def _is_untouched_legacy_material_config(row: Any) -> bool:
    item = _row(row)
    return (
        str(item.get("description") or "") == LEGACY_MATERIAL_DEFAULT["description"]
        and str(item.get("base_url") or "") == LEGACY_MATERIAL_DEFAULT["base_url"]
        and _json(item.get("request_mapping_json"), {}) == LEGACY_MATERIAL_DEFAULT["request_mapping"]
        and _json(item.get("response_mapping_json"), {}) == LEGACY_MATERIAL_DEFAULT["response_mapping"]
    )


def _upgrade_legacy_material_config(conn: Any, row: Any, operated_by: str, now: str) -> None:
    before = _row(row)
    defaults = INTERFACE_DEFAULTS["material_batch_query"]
    next_version = int(before.get("config_version") or 0) + 1
    conn.execute(
        """UPDATE order_interface_configs
           SET description=?,method=?,base_url=?,port=?,path=?,timeout_seconds=?,
               request_mapping_json=?,response_mapping_json=?,config_version=?,updated_at=?
           WHERE id=?""",
        (
            defaults["description"], defaults["method"], defaults["base_url"], defaults["port"],
            defaults["path"], defaults["timeout_seconds"],
            json.dumps(defaults["request_mapping"], ensure_ascii=False),
            json.dumps(defaults["response_mapping"], ensure_ascii=False),
            next_version, now, int(before["id"]),
        ),
    )
    current = conn.execute("SELECT * FROM order_interface_configs WHERE id=?", (int(before["id"]),)).fetchone()
    conn.execute(
        """INSERT INTO order_interface_config_versions
           (interface_config_id,config_version,before_json,after_json,operated_by,created_at)
           VALUES (?,?,?,?,?,?)""",
        (
            int(before["id"]), next_version, json.dumps(before, ensure_ascii=False),
            json.dumps(_row(current), ensure_ascii=False), operated_by, now,
        ),
    )


def _upgrade_material_customer_spec_source(conn: Any, row: Any, operated_by: str, now: str) -> None:
    before = _row(row)
    mapping = _json(before.get("request_mapping_json"), {})
    key = "materialInfoList[].customerSpec"
    if mapping.get(key) != "模板明细.客户规格匹配；为空时取客户规格（必填）":
        return
    mapping[key] = INTERFACE_DEFAULTS["material_batch_query"]["request_mapping"][key]
    next_version = int(before.get("config_version") or 0) + 1
    conn.execute(
        """UPDATE order_interface_configs
           SET request_mapping_json=?,config_version=?,updated_at=? WHERE id=?""",
        (json.dumps(mapping, ensure_ascii=False), next_version, now, int(before["id"])),
    )
    current = conn.execute("SELECT * FROM order_interface_configs WHERE id=?", (int(before["id"]),)).fetchone()
    conn.execute(
        """INSERT INTO order_interface_config_versions
           (interface_config_id,config_version,before_json,after_json,operated_by,created_at)
           VALUES (?,?,?,?,?,?)""",
        (
            int(before["id"]), next_version, json.dumps(before, ensure_ascii=False),
            json.dumps(_row(current), ensure_ascii=False), operated_by, now,
        ),
    )


def _is_untouched_legacy_domestic_config(row: Any) -> bool:
    item = _row(row)
    return (
        str(item.get("description") or "") == LEGACY_DOMESTIC_DEFAULT["description"]
        and str(item.get("base_url") or "") == LEGACY_DOMESTIC_DEFAULT["base_url"]
        and _json(item.get("request_mapping_json"), {}) == LEGACY_DOMESTIC_DEFAULT["request_mapping"]
        and _json(item.get("response_mapping_json"), {}) == LEGACY_DOMESTIC_DEFAULT["response_mapping"]
    )


def _upgrade_legacy_domestic_config(conn: Any, row: Any, operated_by: str, now: str) -> None:
    before = _row(row)
    defaults = INTERFACE_DEFAULTS["domestic_order_entry"]
    next_version = int(before.get("config_version") or 0) + 1
    conn.execute(
        """UPDATE order_interface_configs
           SET display_name=?,description=?,method=?,base_url=?,port=?,path=?,timeout_seconds=?,
               request_mapping_json=?,response_mapping_json=?,config_version=?,updated_at=?
           WHERE id=?""",
        (
            defaults["display_name"], defaults["description"], defaults["method"], defaults["base_url"],
            defaults["port"], defaults["path"], defaults["timeout_seconds"],
            json.dumps(defaults["request_mapping"], ensure_ascii=False),
            json.dumps(defaults["response_mapping"], ensure_ascii=False),
            next_version, now, int(before["id"]),
        ),
    )
    current = conn.execute("SELECT * FROM order_interface_configs WHERE id=?", (int(before["id"]),)).fetchone()
    conn.execute(
        """INSERT INTO order_interface_config_versions
           (interface_config_id,config_version,before_json,after_json,operated_by,created_at)
           VALUES (?,?,?,?,?,?)""",
        (
            int(before["id"]), next_version, json.dumps(before, ensure_ascii=False),
            json.dumps(_row(current), ensure_ascii=False), operated_by, now,
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
        item["maintenance_notes"] = INTERFACE_MAINTENANCE_NOTES.get(item["interface_key"], [])
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
    item["maintenance_notes"] = INTERFACE_MAINTENANCE_NOTES.get(item["interface_key"], [])
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
        "customerCode": "",
        "operatorCode": "",
        "materialInfoList": [{
            "categoryCode": "718",
            "customerMaterialNo": "",
            "customerSpec": "",
            "newProductName": "",
            "oldProductName": "",
        }],
    } if interface_key == "material_batch_query" else {
        "sctoDataList": [{
            "customerCode": "", "orderType": "", "operator": "", "quantity": "",
            "taxPrice": "", "untaxedPrice": "", "materialCode": "", "lineNumber": "1",
            "demandDate": "", "orderNumber": "",
        }],
    }
    if mode == "mock":
        mock_response = (
            {"msg": "Mock 测试成功", "code": 200, "reqParams": request_body, "hitMaterialList": []}
            if interface_key == "material_batch_query"
            else {"msg": "Mock 测试成功", "code": 200, "data": {"data": [], "failCount": 0, "successCount": 0}}
        )
        return {
            "ok": True, "mode": "mock", "status_code": 200, "duration_ms": 0,
            "endpoint_url": endpoint_url,
            "request": request_body,
            "response": mock_response,
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
    "material_query_real": "料号查询接口",
    "domestic_order_entry_mock": "生成订单接口", "domestic_order_entry_real": "生成订单接口",
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
        item["interface_label"] = "批量料号查询" if item.get("interface_key") == "material_batch_query" else "生成订单"
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
    return {
        "line_no": line_no,
        "material_status": str(values.get("material_status") or "查询"),
        "product_type": str(values.get("product_type") or ""),
        "product_name": str(values.get("product_name") or ""),
        "customer_product_code": str(values.get("customer_product_code") or ""),
        "customer_spec": str(values.get("customer_spec") or ""),
        "customer_spec_match": str(values.get("customer_spec_match") or ""),
    }


def _insert_call_log(conn: Any, *, case_id: int, template_id: int, employee_id: str,
                     config: dict[str, Any], status: str, request_payload: dict[str, Any],
                     response_payload: dict[str, Any], triggered_by: str, error_message: str = "",
                     is_mock: bool = True, http_status: int | None = None,
                     duration_ms: int | None = 1) -> int:
    if is_mock and http_status is None:
        http_status = 200 if status == "success" else 422
    cursor = conn.execute(
        """INSERT INTO order_interface_call_logs
           (case_id,template_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
            status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (case_id, template_id, employee_id, int(config["id"]), "material_batch_query",
         int(config["config_version"]), int(is_mock), status, http_status, duration_ms,
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
                            factory_part_no: str, product_name: str, correlation_id: str,
                            source_label: str = "料号查询接口（Mock）") -> list[dict[str, Any]]:
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
            sources[field] = {"label": source_label, "reference": f"关联号 {correlation_id}"}
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
        candidates = [
            {"factory_part_no": f"MOCK-{code[-6:]}", "product_name": f"Mock 品名 {code[-4:]}"},
            {"factory_part_no": f"MOCK-{code[-6:]}-ALT", "product_name": f"Mock 品名 {code[-4:]} 备选"},
        ]
        return {"line_no": line_no, "status": "matched", **candidates[0], "candidates": candidates,
                "matched_spec": item["customer_spec_match"] or item["customer_spec"],
                "message": "Mock 已命中多个料号，可在产品编号或品名中联动选择。"}
    if line_no % 3 == 2:
        return {"line_no": line_no, "status": "creating", "factory_part_no": "创建料号中", "product_name": "创建料号中",
                "external_task_id": f"MOCK-CREATE-{line_no}", "message": "Mock 未找到料号，已发起创建；等待对方回调。"}
    return {"line_no": line_no, "status": "failed", "message": "Mock 查询异常：请检查客户产品编号、客户规格和规格匹配。"}


def build_material_query(case_id: int, employee_id: str, triggered_by: str) -> dict[str, Any]:
    config = get_interface_config("material_batch_query")
    if not config or not config.get("enabled"):
        raise ValueError("批量料号查询接口未启用")
    if str(config.get("mode") or "mock") == "real":
        return build_material_query_real(case_id, employee_id, triggered_by, config=config)
    return build_material_query_mock(case_id, employee_id, triggered_by)


def _material_category_code(product_type: str) -> str:
    value = str(product_type or "").strip()
    normalized = value.upper()
    if normalized in {"PP", "698", "1"}:
        return "698"
    if value == "基板" or normalized in {"718", "2"}:
        return "718"
    return value


def _real_material_request_item(item: dict[str, Any]) -> dict[str, str]:
    product_name = str(item.get("product_name") or "").strip()
    if product_name in {"创建料号中", "创建品名中"}:
        product_name = ""
    is_new = str(item.get("material_status") or "查询").strip() == "新增"
    return {
        "categoryCode": _material_category_code(str(item.get("product_type") or "")),
        "customerMaterialNo": str(item.get("customer_product_code") or "").strip(),
        "customerSpec": str(item.get("customer_spec") or "").strip(),
        "newProductName": product_name if is_new else "",
        "oldProductName": "" if is_new else product_name,
    }


def _post_json_endpoint(
    config: dict[str, Any], payload: dict[str, Any], interface_label: str,
) -> tuple[int, dict[str, Any], int]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(
        _endpoint_url(config), data=body,
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        method=str(config.get("method") or "POST").upper(),
    )
    started = time.monotonic()
    try:
        with urlopen(request, timeout=int(config.get("timeout_seconds") or 15)) as response:
            raw = response.read(1024 * 1024).decode("utf-8", errors="replace")
            status_code = int(response.status)
    except HTTPError as exc:
        raw = exc.read(1024 * 1024).decode("utf-8", errors="replace")
        status_code = int(exc.code)
    except (URLError, TimeoutError, OSError) as exc:
        raise ValueError(f"{interface_label}请求失败：{str(exc)[:160]}") from exc
    try:
        response_body = json.loads(raw) if raw else {}
    except ValueError as exc:
        raise ValueError(f"{interface_label}返回的不是合法 JSON") from exc
    if not isinstance(response_body, dict):
        raise ValueError(f"{interface_label}返回格式错误")
    return status_code, response_body, int((time.monotonic() - started) * 1000)


def _real_material_response_items(
    request_items: list[dict[str, Any]], response_body: dict[str, Any], http_status: int,
) -> list[dict[str, Any]]:
    errors_by_line: dict[int, str] = {}
    for message in response_body.get("errors") or []:
        matched = re.search(r"第\s*(\d+)\s*行", str(message))
        if matched:
            errors_by_line[int(matched.group(1))] = str(message).strip()
    hits = [item for item in (response_body.get("hitMaterialList") or []) if isinstance(item, dict)]
    hits_by_customer_part: dict[str, list[dict[str, Any]]] = {}
    for hit in hits:
        hits_by_customer_part.setdefault(str(hit.get("scca03") or "").strip(), []).append(hit)
    business_ok = http_status == 200 and int(response_body.get("code") or 0) == 200
    results = []
    for index, item in enumerate(request_items, start=1):
        customer_part = str(item.get("customer_product_code") or "").strip()
        matched_hits = hits_by_customer_part.get(customer_part, [])
        if not matched_hits and len(request_items) == 1:
            matched_hits = hits
        candidates = [
            {
                "factory_part_no": str(hit.get("peag01") or "").strip(),
                "product_name": str(hit.get("peag08") or "").strip(),
            }
            for hit in matched_hits
            if str(hit.get("peag01") or hit.get("peag08") or "").strip()
        ]
        if candidates:
            results.append({
                "line_no": item["line_no"], "status": "matched", **candidates[0],
                "candidates": candidates, "matched_spec": str(matched_hits[0].get("scca05") or ""),
                "message": f"真实接口命中 {len(candidates)} 个候选料号。",
            })
        elif index in errors_by_line:
            results.append({"line_no": item["line_no"], "status": "failed", "message": errors_by_line[index]})
        elif business_ok and response_body.get("pera01"):
            results.append({
                "line_no": item["line_no"], "status": "creating",
                "factory_part_no": "创建料号中", "product_name": "创建料号中",
                "external_task_id": str(response_body.get("pera01") or ""),
                "message": f"未命中现有料号，已返回编制作业单 {response_body['pera01']}。",
            })
        else:
            results.append({
                "line_no": item["line_no"], "status": "failed",
                "message": str(response_body.get("msg") or f"接口返回 HTTP {http_status}"),
            })
    return results


def build_material_query_real(
    case_id: int, employee_id: str, triggered_by: str, *, config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if is_domestic_order_entry_completed(case_id, employee_id):
        raise ValueError("内销录单已完成，不能再次请求料号查询接口")
    config = config or get_interface_config("material_batch_query")
    if not config or not config.get("enabled"):
        raise ValueError("批量料号查询接口未启用")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成录单模板")
    with db_cursor() as conn:
        template = conn.execute("SELECT header_json FROM order_entry_templates WHERE id=?", (template_id,)).fetchone()
        rows = conn.execute(
            "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
            (template_id,),
        ).fetchall()
    header = _json(template["header_json"] if template else "", {})
    customer_code = str(header.get("bill_to_customer_code") or "").strip()
    if not customer_code:
        raise ValueError("请先填写并保存账款客户编号")
    request_items = [_material_request_item(int(row["line_no"]), _json(row["values_json"], {})) for row in rows]
    if not request_items:
        raise ValueError("当前没有可查询的订单明细")
    request_payload = {
        "customerCode": customer_code,
        "operatorCode": employee_id,
        "materialInfoList": [_real_material_request_item(item) for item in request_items],
    }
    try:
        http_status, response_body, duration_ms = _post_json_endpoint(config, request_payload, "真实料号查询")
    except ValueError as exc:
        with db_cursor() as conn:
            call_id = _insert_call_log(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                status="failed", request_payload=request_payload, response_payload={}, triggered_by=triggered_by,
                error_message=str(exc), is_mock=False, http_status=None, duration_ms=None,
            )
            record_order_detail_event(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                event_type="material_query_real", title="批量料号查询（真实接口）失败",
                detail={"call_id": call_id, "error_message": str(exc)}, operated_by=triggered_by,
            )
        raise
    response_items = _real_material_response_items(request_items, response_body, http_status)
    call_status = "success" if any(item["status"] in {"matched", "creating"} for item in response_items) else "failed"
    all_changes: list[dict[str, Any]] = []
    with db_cursor() as conn:
        call_id = _insert_call_log(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
            status=call_status, request_payload=request_payload, response_payload=response_body,
            triggered_by=triggered_by, error_message="" if call_status == "success" else str(response_body.get("msg") or "查询失败"),
            is_mock=False, http_status=http_status, duration_ms=duration_ms,
        )
        rows_by_line = {int(row["line_no"]): row for row in rows}
        for item, response in zip(request_items, response_items):
            status = response["status"]
            task_status = "resolved" if status == "matched" else ("waiting_callback" if status == "creating" else "failed")
            task = _upsert_resolution_task(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                line_no=item["line_no"], status=task_status, input_item=item, call_id=call_id, result=response,
            )
            values = _json(rows_by_line[item["line_no"]]["values_json"], {})
            if status in {"matched", "creating"}:
                changes = _backfill_material_line(
                    conn, template_id=template_id, line_no=item["line_no"], values=values,
                    factory_part_no=str(response.get("factory_part_no") or ""),
                    product_name=str(response.get("product_name") or ""), correlation_id=task["correlation_id"],
                    source_label="料号查询接口（真实）",
                )
                all_changes.extend(changes)
            if status == "matched":
                conn.execute("UPDATE order_material_resolution_tasks SET resolved_at=?,updated_at=? WHERE id=?", (utcnow(), utcnow(), task["id"]))
            elif status == "creating":
                conn.execute(
                    "UPDATE order_material_resolution_tasks SET external_task_id=?,updated_at=? WHERE id=?",
                    (response.get("external_task_id", ""), utcnow(), task["id"]),
                )
            suggestions = response.get("candidates") or [response]
            for suggestion in suggestions:
                conn.execute(
                    """INSERT INTO order_material_query_suggestions
                       (call_log_id,template_id,line_no,factory_part_no,product_name,matched_spec,status,message,created_at)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (
                        call_id, template_id, item["line_no"], suggestion.get("factory_part_no", ""),
                        suggestion.get("product_name", ""), response.get("matched_spec", ""), status,
                        response.get("message", ""), utcnow(),
                    ),
                )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="material_query_real", title="批量料号查询（真实接口）完成",
            detail={"call_id": call_id, "changes": all_changes, "items": response_items}, operated_by=triggered_by,
        )
    return {"call_id": call_id, "items": response_items, "status": call_status, "mode": "real"}


def build_material_query_mock(case_id: int, employee_id: str, triggered_by: str, scenario: str = "success",
                              *, line_nos: set[int] | None = None,
                              callback_requery: bool = False) -> dict[str, Any]:
    """Run the production-shaped Mock: hit, creation callback, then automatic backfill."""
    if is_domestic_order_entry_completed(case_id, employee_id):
        raise ValueError("内销录单已完成，不能再次请求料号查询接口")
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
        item["candidates"] = _material_candidates(item["result"])
        item["label"] = MATERIAL_STATUS_LABELS.get(item["status"], item["status"])
        items.append(item)
    return {"items": items, "pending": any(item["status"] in {"waiting_callback", "requerying"} for item in items)}


def _material_candidates(result: dict[str, Any]) -> list[dict[str, str]]:
    raw_candidates = result.get("candidates") or result.get("hitMaterialList") or []
    if not raw_candidates and (result.get("factory_part_no") or result.get("product_name")):
        raw_candidates = [result]
    candidates: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for raw in raw_candidates:
        if not isinstance(raw, dict):
            continue
        candidate = {
            "product_code": str(raw.get("factory_part_no") or raw.get("peag01") or "").strip(),
            "product_name": str(raw.get("product_name") or raw.get("peag08") or "").strip(),
        }
        key = (candidate["product_code"], candidate["product_name"])
        if key == ("", "") or key in seen:
            continue
        seen.add(key)
        candidates.append(candidate)
    return candidates


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


def is_domestic_order_entry_completed(case_id: int, employee_id: str) -> bool:
    """Whether a successful domestic-entry request has already been accepted.

    The interface log is the source of truth here: a case can remain visible for
    review after entry, but it must not submit the same order a second time.
    """
    with db_cursor() as conn:
        row = conn.execute(
            """SELECT 1 FROM order_interface_call_logs
               WHERE case_id=? AND employee_id=? AND interface_key='domestic_order_entry'
                 AND status='success'
               LIMIT 1""",
            (case_id, employee_id),
        ).fetchone()
    return bool(row)


def build_domestic_order_entry_mock(case_id: int, employee_id: str, triggered_by: str) -> dict[str, Any]:
    """Record a manual domestic-order submission Mock without changing the template."""
    if is_domestic_order_entry_completed(case_id, employee_id):
        raise ValueError("内销录单已完成，不能重复提交")
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
        conn.execute(
            """UPDATE order_intake_cases
               SET status='archived', workflow_stage='completed', erp_prepare_status='submitted',
                   completed_at=?, updated_at=?
               WHERE id=? AND employee_id=?""",
            (now, now, case_id, employee_id),
        )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="domestic_order_entry_mock", title="提交内销录单（Mock）完成",
            detail={"call_id": call_id, "line_count": len(lines), "entry_no": response_payload["entry_no"]},
            operated_by=triggered_by,
        )
    return {"call_id": call_id, "entry_no": response_payload["entry_no"], "status": "success", "mode": "mock"}


def build_domestic_order_entry(case_id: int, employee_id: str, triggered_by: str) -> dict[str, Any]:
    config = get_interface_config("domestic_order_entry")
    if not config or not config.get("enabled"):
        raise ValueError("生成订单接口未启用")
    if str(config.get("mode") or "mock") == "real":
        return build_domestic_order_entry_real(case_id, employee_id, triggered_by, config=config)
    return build_domestic_order_entry_mock(case_id, employee_id, triggered_by)


def _domestic_order_request_payload(
    header: dict[str, Any], rows: list[Any], employee_id: str,
) -> dict[str, Any]:
    customer_code = str(header.get("bill_to_customer_code") or "").strip()
    order_type = str(header.get("order_type") or "").strip()
    order_number = str(header.get("customer_order_number") or "").strip()
    header_issues = []
    if not customer_code:
        header_issues.append("未填写账款客户编号")
    if not order_type:
        header_issues.append("未填写单别")
    if not order_number:
        header_issues.append("未填写客户订单号")
    if header_issues:
        raise ValueError("暂不能提交录单：" + "；".join(header_issues))
    items = []
    line_issues = []
    for row in rows:
        line_no = int(row["line_no"])
        values = _json(row["values_json"], {})
        quantity = str(values.get("quantity") or "").strip()
        material_code = str(values.get("product_code") or "").strip()
        demand_date = str(values.get("delivery_date") or "").strip()
        tax_price = str(values.get("unit_price") or "").strip()
        untaxed_price = str(values.get("price_before_tax") or "").strip()
        if not quantity:
            line_issues.append(f"第 {line_no} 行未填写数量")
        if not material_code:
            line_issues.append(f"第 {line_no} 行未填写产品编号")
        if not demand_date:
            line_issues.append(f"第 {line_no} 行未填写出货日期")
        if not tax_price and not untaxed_price:
            line_issues.append(f"第 {line_no} 行单价和税前单价至少填写一项")
        items.append({
            "customerCode": customer_code,
            "orderType": order_type,
            "operator": employee_id,
            "quantity": quantity,
            "taxPrice": tax_price,
            "untaxedPrice": untaxed_price,
            "materialCode": material_code,
            "lineNumber": str(line_no),
            "demandDate": demand_date,
            "orderNumber": order_number,
            "custOrderId": order_number,
            "lineId": str(values.get("customer_order_seq") or "").strip(),
            "lineRemark": str(values.get("remark") or "").strip(),
            "taxType": str(header.get("tax_type") or "").strip(),
            "materialName": str(values.get("product_name") or "").strip(),
            "spec": str(values.get("customer_spec_match") or values.get("customer_spec") or "").strip(),
        })
    if line_issues:
        raise ValueError("暂不能提交录单：" + "；".join(line_issues))
    return {"sctoDataList": items}


def _insert_domestic_call_log(
    conn: Any, *, case_id: int, template_id: int, employee_id: str, config: dict[str, Any],
    status: str, request_payload: dict[str, Any], response_payload: dict[str, Any],
    triggered_by: str, http_status: int | None, duration_ms: int | None,
    error_message: str = "",
) -> int:
    cursor = conn.execute(
        """INSERT INTO order_interface_call_logs
           (case_id,template_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
            status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            case_id, template_id, employee_id, int(config["id"]), "domestic_order_entry",
            int(config["config_version"]), 0, status, http_status, duration_ms,
            json.dumps(request_payload, ensure_ascii=False), json.dumps(response_payload, ensure_ascii=False),
            error_message, triggered_by, utcnow(),
        ),
    )
    return int(cursor.lastrowid)


def _domestic_response_result(response_body: dict[str, Any], http_status: int) -> tuple[bool, str, str]:
    data = response_body.get("data") if isinstance(response_body.get("data"), dict) else {}
    records = [item for item in (data.get("data") or []) if isinstance(item, dict)]
    try:
        success_count = int(data.get("successCount") or 0)
        fail_count = int(data.get("failCount") or 0)
    except (TypeError, ValueError):
        success_count, fail_count = 0, len(records)
    record_failed = any(str(item.get("status") or "").lower() != "success" for item in records)
    ok = (
        http_status == 200
        and int(response_body.get("code") or 0) == 200
        and fail_count == 0
        and not record_failed
        and (success_count > 0 or bool(records))
    )
    entry_numbers = [str(item.get("sctaCode") or "").strip() for item in records if item.get("sctaCode")]
    messages = [str(item.get("message") or "").strip() for item in records if item.get("message")]
    message = "；".join(messages) or str(response_body.get("msg") or ("订单生成成功" if ok else "订单生成失败"))
    return ok, "、".join(entry_numbers), message


def build_domestic_order_entry_real(
    case_id: int, employee_id: str, triggered_by: str, *, config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if is_domestic_order_entry_completed(case_id, employee_id):
        raise ValueError("录单已完成，不能重复提交")
    issues = validate_domestic_order_entry(case_id, employee_id)
    if issues:
        raise ValueError("暂不能提交录单：" + "；".join(issues))
    config = config or get_interface_config("domestic_order_entry")
    if not config or not config.get("enabled"):
        raise ValueError("生成订单接口未启用")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成并保存录单模板")
    with db_cursor() as conn:
        template = conn.execute("SELECT header_json FROM order_entry_templates WHERE id=?", (template_id,)).fetchone()
        rows = conn.execute(
            "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
            (template_id,),
        ).fetchall()
    header = _json(template["header_json"] if template else "", {})
    request_payload = _domestic_order_request_payload(header, list(rows), employee_id)
    try:
        http_status, response_body, duration_ms = _post_json_endpoint(config, request_payload, "真实生成订单")
    except ValueError as exc:
        with db_cursor() as conn:
            call_id = _insert_domestic_call_log(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                status="failed", request_payload=request_payload, response_payload={}, triggered_by=triggered_by,
                http_status=None, duration_ms=None, error_message=str(exc),
            )
            record_order_detail_event(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                event_type="domestic_order_entry_real", title="提交录单（真实接口）失败",
                detail={"call_id": call_id, "error_message": str(exc)}, operated_by=triggered_by,
            )
        raise
    ok, entry_no, message = _domestic_response_result(response_body, http_status)
    now = utcnow()
    with db_cursor() as conn:
        call_id = _insert_domestic_call_log(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
            status="success" if ok else "failed", request_payload=request_payload,
            response_payload=response_body, triggered_by=triggered_by, http_status=http_status,
            duration_ms=duration_ms, error_message="" if ok else message,
        )
        if ok:
            conn.execute(
                """UPDATE order_intake_cases
                   SET status='archived',workflow_stage='completed',erp_prepare_status='submitted',
                       completed_at=?,updated_at=? WHERE id=? AND employee_id=?""",
                (now, now, case_id, employee_id),
            )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="domestic_order_entry_real",
            title="提交录单（真实接口）完成" if ok else "提交录单（真实接口）失败",
            detail={"call_id": call_id, "line_count": len(rows), "entry_no": entry_no, "error_message": "" if ok else message},
            operated_by=triggered_by,
        )
    if not ok:
        raise ValueError(message)
    return {"call_id": call_id, "entry_no": entry_no, "status": "success", "mode": "real", "message": message}
