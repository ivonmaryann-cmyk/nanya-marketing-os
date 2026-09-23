from __future__ import annotations

import json
import os
import re
import ssl
import time
import uuid
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen
from typing import Any

from .database import automation_cursor as db_cursor
from .customer_spec_mapping_service import extract_structure_from_customer_spec
from .db import get_user, utcnow
from .order_price_validation_service import (
    PRICE_REVIEW_SNAPSHOT_KEY,
    PriceMismatchConfirmationRequired,
    review_cached_template_prices,
)
from .paths import PACKAGE_DIR
from .purchase_field_rules import normalize_date
from .product_name_extract import extract as extract_new_product_name
from . import product_name_service
from zoneinfo import ZoneInfo


ORDER_ACCOUNT_SET_BY_ACSN = {
    "NY01": "KL01",
    "NY02": "KL02",
    "NY03": "KL55",
}

ORDER_INFO_QUERY_BATCH_SIZE = 50


def _order_account_set(acsn: Any) -> str:
    return ORDER_ACCOUNT_SET_BY_ACSN.get(str(acsn or "").strip().upper(), "")


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
            "acsn": "组织代码（上海=NY01，江西=NY02；默认传 NY01）",
            "operatorCode": "当前登录账号（员工工号，必填）",
            "materialInfoList[].categoryCode": "模板明细.产品类型（PP=698，基板=718，必填）",
            "materialInfoList[].customerMaterialNo": "模板明细.客户产品编号（必填）",
            "materialInfoList[].origin": "模板明细.产地（上海=1，江西=2，江苏=3；未填写为空）",
            "materialInfoList[].customerSpec": "新建料号弹窗.客户规格匹配（仅点击新建料号时传）",
            "materialInfoList[].customerSpecOld": "原始客户规格（选填，默认不传）",
            "materialInfoList[].oriCustomerSpec": "新建料号弹窗.客户规格（仅点击新建料号时传）",
            "materialInfoList[].newProductName": "新建料号弹窗.品名（仅点击新建料号时传）",
            "materialInfoList[].newFlag": "点击新建料号时固定发送 Y；批量查询不发送",
            "materialInfoList[].layoutStructure": "客户排版结构（选填，默认不传）",
            "materialInfoList[].thicknessDescription": "客户厚度描述（选填，默认不传）",
            "materialInfoList[].specialRequirements": "客户特殊要求（选填，默认不传）",
        },
        "response_mapping": {
            "code": "接口交互记录.业务状态码（200=处理完成，999=失败）",
            "msg": "接口交互记录.提示",
            "pera01": "料号查询建议.客户料号编制作业单号",
            "newFlag": "料号查询建议.是否新建（Y=新建，N=未新建）",
            "source": "料号查询建议.来源（PARSE=规格解析）",
            "errors[]": "接口交互记录.逐行失败信息",
            "hitMaterialList[].peag01": "料号查询建议.产品编号",
            "hitMaterialList[].peag08": "料号查询建议.新品名（回填模板）",
            "hitMaterialList[].peag09": "料号查询建议.旧品名（不回填模板）",
            "hitMaterialList[].peag06": "料号查询建议.销售品名规格",
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
            "sctoDataList[].factoryPartCode": "模板明细.产品编号（必填）",
            "sctoDataList[].materialCode": "模板明细.客户产品编号（必填）",
            "sctoDataList[].lineNumber": "模板明细.项次（与客户订单序号一致，必填）",
            "sctoDataList[].demandDate": "模板明细.出货日期（必填）",
            "sctoDataList[].orderNumber": "模板表头.客户订单号（必填）",
            "sctoDataList[].custOrderId": "模板表头.客户订单号（必填）",
            "sctoDataList[].lineId": "模板明细.客户订单序号（与项次一致，必填）",
            "sctoDataList[].lineRemark": "模板明细.备注（选填）",
            "sctoDataList[].taxType": "模板表头.税种（选填）",
            "sctoDataList[].materialName": "模板明细.品名（选填）",
            "sctoDataList[].spec": "模板明细.客户规格（选填）",
        },
        "response_mapping": {
            "code": "接口交互记录.业务状态码",
            "msg": "接口交互记录.提示",
            "data.successCount": "接口交互记录.成功数量",
            "data.failCount": "接口交互记录.失败数量",
            "data.data[].orderNumber": "接口交互记录.客户订单号",
            "data.data[].sctaCode": "接口交互记录.生成订单号",
            "data.data[].scta39": "接口交互记录.ERP订单号",
            "data.data[].status": "接口交互记录.生成状态（success/fail）",
            "data.data[].message": "接口交互记录.失败原因",
            "data.erpOrderMap": "接口交互记录.NYEOS订单号与ERP订单号映射",
        },
    },
    "order_info_query": {
        "display_name": "查询订单信息",
        "description": "按客户单号批量查询订单单头、明细及未匹配的客户单号。",
        "method": "POST",
        "base_url": "http://nyeos2.nouyatec.com:7030/NY01-APP/nyeos/api/sc/queryOrderInfo",
        "port": 7030,
        "path": "",
        "timeout_seconds": 15,
        "request_mapping": {
            "orderNumberList[]": "客户单号列表（必填，支持批量）",
            "customerMaterialCode": "客户料号（选填）",
            "customerOrderNo": "客户订单号（选填）",
            "itemNo": "项次（选填）",
            "customerSpec": "客户规格（选填）",
        },
        "response_mapping": {
            "code": "接口交互记录.业务状态码（200=查询成功）",
            "msg": "接口交互记录.提示",
            "data.orderList[].scta01": "订单信息.订单号",
            "data.orderList[].scta11": "订单信息.送货客户 ID",
            "data.orderList[].scta38": "订单信息.客户订单号",
            "data.orderList[].scta39": "订单信息.ERP 订单号",
            "data.orderList[].acsn": "订单信息.组织代码（NY01/NY02/NY03）",
            "data.orderList[].sctbList[].sctb02": "订单明细.料号",
            "data.orderList[].sctbList[].sctb03": "订单明细.品名规格",
            "data.orderList[].sctbList[].sctb05": "订单明细.数量",
            "data.orderList[].sctbList[].sctb23": "订单明细.未出数量",
            "data.orderList[].sctbList[].sctb06": "订单明细.含税单价",
            "data.orderList[].sctbList[].sctb07": "订单明细.未税单价",
            "data.orderList[].sctbList[].sctb14": "订单明细.客户料号",
            "data.orderList[].sctbList[].sctb15": "订单明细.客户单号",
            "data.orderList[].sctbList[].sctb16": "订单明细.客户需求日",
            "data.orderList[].sctbList[].sctb17": "订单明细.预计出货日",
            "data.orderList[].sctbList[].sctb30": "订单明细.结案码",
            "data.orderList[].sctbList[].sctb35": "订单明细.项次",
            "data.orderList[].sctbList[].sctb36": "订单明细.客户规格",
            "data.orderList[].sctbList[].sctb43": "订单明细.厂别",
            "data.orderCount": "查询结果.匹配订单数",
            "data.notFoundList[]": "查询结果.未匹配客户单号",
        },
    },
    "aps_order_demand_import": {
        "display_name": "APS订单需求导入",
        "description": "将已确认的订单需求变更批量导入 APS；创建人取当前账号用户名，单次建议不超过 500 条。",
        "method": "POST",
        "base_url": "http://aps.nouyatec.com:13000/forward/erp_order_change_blocking",
        "port": 13000,
        "path": "",
        "timeout_seconds": 15,
        "request_mapping": {
            "data[].require_shipment_date": "需求出货日期（必填，YYYY-MM-DD）",
            "data[].Order_Item_Account_Set_outer_key": "订单项次账套外键（必填，唯一标识订单项次）",
            "data[].alter_type": "变更类型（必填，仅限接口约定的10种值）",
            "data[].creator_name": "当前登录账号.用户名（必填，APS业务人员昵称）",
            "data[].require_specification": "需求说明（选填，默认空字符串）",
            "data[].emergency_score": "紧急度分值（选填，默认0）",
            "data[].NPI_Commitment_Statement": "NPI承诺说明（选填，默认空字符串）",
            "data[].Estimated_launch_month": "预估起量月份（选填，YYYY-MM或YYYY-MM-DD）",
            "data[].Estimated_volume": "预估量（选填，默认0.0）",
            "data[].created_at": "提交时间（选填，YYYY-MM-DD HH:MM:SS）",
        },
        "response_mapping": {
            "code": "接口交互记录.状态码",
            "message": "接口交互记录.处理结果",
            "success_count": "接口交互记录.成功条数",
            "fail_count": "接口交互记录.失败条数（HTTP 200时仍必须检查）",
            "failed_details[].row": "失败明细.请求序号（从1开始）",
            "failed_details[].order_key": "失败明细.订单项次账套外键",
            "failed_details[].reason": "失败明细.失败原因",
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
        "点击“新建料号”时，customerSpec 发送客户规格匹配，oriCustomerSpec 发送客户规格；customerSpecOld 默认不传。",
        "点击“新建料号”时，newProductName 发送当前品名，newFlag 固定发送 Y；不发送胶系编码和客户规格匹配。",
        "接口未命中料号时可能返回 pera01，并在对方系统创建客户料号编制作业，请勿用随意数据测试。",
        "保存后业务页会按运行模式执行：Mock 走模拟流程，真实接口会请求当前地址。",
    ],
    "domestic_order_entry": [
        "当前地址是 NYEOS 测试环境；正式环境地址确认后只需修改“请求地址”。",
        "factoryPartCode 使用模板中的产品编号（必填）；materialCode 使用模板中的客户产品编号。",
        "orderNumber 与 custOrderId 都使用客户订单号；lineNumber 与 lineId 都使用客户订单序号。",
        "保存后业务页会按运行模式执行：Mock 走模拟流程，真实接口会生成订单。",
    ],
    "order_info_query": [
        "按客户单号批量查询订单；可用客户料号、客户订单号、项次或客户规格进一步筛选。",
        "Mock 测试会返回一笔订单、一个订单明细及一个未匹配客户单号，便于核对字段映射。",
        "保存后可通过“接口测试”验证 Mock 或当前真实地址，查询本身不会修改订单数据。",
    ],
    "aps_order_demand_import": [
        "接口请求必须使用 data 数组包裹；即使只有一条记录也不能发送裸对象。",
        "creator_name 必填，取当前登录账号在“账户与密码”维护的用户名，且必须是 APS 业务人员昵称。",
        "created_at 由系统按中国标准时间生成，格式为 YYYY-MM-DD HH:MM:SS。",
        "HTTP 200 不代表全部成功，必须同时检查 fail_count 和 failed_details。",
        "接口自身不去重，正式调用前必须避免重复推送同一订单项次账套外键。",
        "当前默认使用 Mock 模式；切换真实接口后会向 APS 写入数据，测试数据需联系 APS 删除。",
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


def _ssl_context_for_endpoint(endpoint_url: str) -> ssl.SSLContext | None:
    """Return the default-verifying TLS context, extended for NYEOS if configured.

    ``api.crt`` is a self-signed certificate for the internal NYEOS host.  It
    must be added as a *single extra trust anchor* instead of disabling TLS
    verification for the whole application.  The setting deliberately applies
    only to the NYEOS host/IP covered by that certificate.
    """
    parsed = urlparse(endpoint_url)
    if parsed.scheme.lower() != "https":
        return None
    if (parsed.hostname or "").lower() not in {"nyeos.nouyatec.com", "10.30.12.117"}:
        return None
    ca_file = os.getenv("NYEOS_CA_CERT_FILE", "").strip()
    if not ca_file:
        return None
    try:
        # create_default_context keeps the OS trust store, then adds the
        # internal NYEOS certificate. Hostname validation remains enabled.
        return ssl.create_default_context(cafile=ca_file)
    except (FileNotFoundError, OSError, ssl.SSLError) as exc:
        raise ValueError(f"Nyeos 接口证书配置无效：{str(exc)[:160]}") from exc


def _valid_endpoint_url(value: str) -> str:
    endpoint_url = str(value or "").strip()
    parsed = urlparse(endpoint_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("请求地址必须是完整的 http 或 https URL")
    return endpoint_url


NYEOS_TLS_HOST = "nyeos.nouyatec.com"
NYEOS_TLS_CERT_FILE = PACKAGE_DIR / "certs" / "nyeos.nouyatec.com.crt"


def _interface_ssl_context(endpoint_url: str) -> ssl.SSLContext | None:
    """Trust the supplied NYEOS certificate without disabling TLS verification."""
    configured = _ssl_context_for_endpoint(endpoint_url)
    if configured is not None:
        return configured
    parsed = urlparse(endpoint_url)
    if parsed.scheme != "https" or (parsed.hostname or "").lower() != NYEOS_TLS_HOST:
        return None
    if not NYEOS_TLS_CERT_FILE.is_file():
        raise ValueError(f"未找到 NYEOS HTTPS 证书文件：{NYEOS_TLS_CERT_FILE}")
    return ssl.create_default_context(cafile=str(NYEOS_TLS_CERT_FILE))


def _open_interface_request(request: Request, *, timeout: int):
    return urlopen(request, timeout=timeout, context=_interface_ssl_context(request.full_url))


def _decode_interface_response(raw: bytes, headers: Any) -> str:
    declared = getattr(headers, "get_content_charset", lambda: None)() or "utf-8"
    text = raw.decode(declared, errors="replace")
    # The NYEOS gateway declares UTF-8 but returns GB18030 on some business errors.
    if "\ufffd" in text:
        fallback = raw.decode("gb18030", errors="replace")
        if fallback.count("\ufffd") < text.count("\ufffd"):
            return fallback
    return text


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
                    existing = conn.execute(
                        "SELECT * FROM order_interface_configs WHERE interface_key=?", (interface_key,)
                    ).fetchone()
                    _upgrade_material_optional_request_fields(conn, existing, operated_by, now)
                elif interface_key == "domestic_order_entry":
                    if _is_untouched_legacy_domestic_config(existing):
                        _upgrade_legacy_domestic_config(conn, existing, operated_by, now)
                        existing = conn.execute(
                            "SELECT * FROM order_interface_configs WHERE interface_key=?", (interface_key,)
                        ).fetchone()
                    _upgrade_domestic_request_mapping(conn, existing, operated_by, now)
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
    defaults = INTERFACE_DEFAULTS["material_batch_query"]["request_mapping"]
    key = "materialInfoList[].customerSpec"
    legacy_customer_specs = {
        "模板明细.客户规格（必填）",
        "模板明细.客户规格匹配；为空时取客户规格（必填）",
    }
    changed = False
    if mapping.get(key) in legacy_customer_specs:
        mapping[key] = defaults[key]
        changed = True
        for mapping_key in (
            "materialInfoList[].customerSpecOld",
            "materialInfoList[].newProductName",
            "materialInfoList[].newFlag",
        ):
            if mapping_key not in mapping:
                mapping[mapping_key] = defaults[mapping_key]
    if not changed:
        return
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


def _upgrade_material_optional_request_fields(conn: Any, row: Any, operated_by: str, now: str) -> None:
    """Expose new optional NYEOS fields without changing request behaviour."""
    before = _row(row)
    mapping = _json(before.get("request_mapping_json"), {})
    defaults = INTERFACE_DEFAULTS["material_batch_query"]["request_mapping"]
    optional_keys = (
        "materialInfoList[].layoutStructure",
        "materialInfoList[].thicknessDescription",
        "materialInfoList[].specialRequirements",
    )
    changed = False
    for key in optional_keys:
        if key not in mapping:
            mapping[key] = defaults[key]
            changed = True
    if mapping.get("materialInfoList[].customerSpecOld") != defaults["materialInfoList[].customerSpecOld"]:
        mapping["materialInfoList[].customerSpecOld"] = defaults["materialInfoList[].customerSpecOld"]
        changed = True
    for key in (
        "materialInfoList[].customerSpec",
        "materialInfoList[].oriCustomerSpec",
    ):
        if mapping.get(key) != defaults[key]:
            mapping[key] = defaults[key]
            changed = True
    if not changed:
        return
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


def _upgrade_domestic_request_mapping(conn: Any, row: Any, operated_by: str, now: str) -> None:
    before = _row(row)
    mapping = _json(before.get("request_mapping_json"), {})
    defaults = INTERFACE_DEFAULTS["domestic_order_entry"]["request_mapping"]
    changed = False
    legacy_values = {
        "sctoDataList[].materialCode": "模板明细.产品编号（料号查询结果，必填）",
        "sctoDataList[].lineNumber": "模板明细.项次（必填）",
        "sctoDataList[].lineId": "模板明细.客户订单序号（选填）",
        "sctoDataList[].spec": "模板明细.客户规格匹配；为空时取客户规格（选填）",
    }
    for key, legacy_value in legacy_values.items():
        if mapping.get(key) == legacy_value:
            mapping[key] = defaults[key]
            changed = True
    if "sctoDataList[].custOrderId" not in mapping:
        mapping["sctoDataList[].custOrderId"] = defaults["sctoDataList[].custOrderId"]
        changed = True
    if "sctoDataList[].factoryPartCode" not in mapping:
        mapping["sctoDataList[].factoryPartCode"] = defaults["sctoDataList[].factoryPartCode"]
        changed = True
    # Remove only the two mapping rows introduced by the reverted customer-
    # specification payload expansion; user-customized mappings remain intact.
    for key in ("sctoDataList[].customerSpec", "sctoDataList[].oriCustomerSpec"):
        if mapping.get(key) == "模板明细.客户规格（选填）":
            mapping.pop(key)
            changed = True
    if not changed:
        return
    next_version = int(before.get("config_version") or 0) + 1
    conn.execute(
        "UPDATE order_interface_configs SET request_mapping_json=?,config_version=?,updated_at=? WHERE id=?",
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
    if interface_key == "material_batch_query":
        request_body = {
            "customerCode": "",
            "acsn": "NY01",
            "operatorCode": "",
            "materialInfoList": [{
                "categoryCode": "718",
                "customerMaterialNo": "",
                "newProductName": "",
            }],
        }
    elif interface_key == "domestic_order_entry":
        request_body = {
            "sctoDataList": [{
                "customerCode": "", "orderType": "", "operator": "", "quantity": "",
                "taxPrice": "", "untaxedPrice": "", "materialCode": "", "lineNumber": "1",
                "demandDate": "", "orderNumber": "",
            }],
        }
    elif interface_key == "order_info_query":
        request_body = {
            "orderNumberList": ["MOCK-ORDER-001", "MOCK-ORDER-NOT-FOUND"],
            "customerMaterialCode": "",
            "customerOrderNo": "",
            "itemNo": "",
            "customerSpec": "",
        }
    elif interface_key == "aps_order_demand_import":
        request_body = {
            "data": [{
                "require_shipment_date": "2026-09-15",
                "Order_Item_Account_Set_outer_key": "220-260114007_2_KL01",
                "alter_type": "交期变更",
                "creator_name": "Mock 测试用户",
                "require_specification": "Mock 接口测试",
                "emergency_score": 5,
                "NPI_Commitment_Statement": "无",
                "Estimated_launch_month": "2026-10",
                "Estimated_volume": 1500.5,
                "created_at": "2026-09-14 10:00:00",
            }],
        }
    else:
        raise ValueError("未知的接口配置")
    if mode == "mock":
        if interface_key == "material_batch_query":
            mock_response = {
                "msg": "Mock 测试成功", "code": 200,
                "reqParams": request_body, "hitMaterialList": [],
            }
        elif interface_key == "domestic_order_entry":
            mock_response = {
                "msg": "Mock 测试成功", "code": 200,
                "data": {"data": [], "failCount": 0, "successCount": 0},
            }
        elif interface_key == "order_info_query":
            mock_response = {
                "msg": "查询成功",
                "code": 200,
                "data": {
                    "orderList": [{
                        "scta01": "MOCK-ORDER-001",
                        "scta11": "MOCK-CUSTOMER-001",
                        "scta38": "MOCK-ORDER-001",
                        "scta39": "MOCK-ERP-001", "acsn": "NY01",
                        "sctbList": [{
                            "sctb02": "MOCK-PART-001",
                            "sctb03": "Mock 品名规格",
                            "sctb04": "2",
                            "peag04": "Mock 品名规格",
                            "peag08": "MOCK-NEW-PRODUCT",
                            "peag09": "MOCK-OLD-PRODUCT",
                            "szaa01": "米",
                            "sctb05": 100,
                            "sctb23": 100,
                            "sctb06": 12.50,
                            "sctb07": 11.06,
                            "sctb14": "MOCK-CUSTOMER-PART",
                            "sctb15": "MOCK-ORDER-001",
                            "sctb16": "2026-09-14",
                            "sctb17": "2026-09-15 00:00:00",
                            "sctb30": "N",
                            "sctb35": 10, "sctb43": "S1",
                            "sctb36": "Mock 客户规格",
                        }],
                    }],
                    "orderCount": 1,
                    "notFoundList": ["MOCK-ORDER-NOT-FOUND"],
                },
            }
        else:
            mock_response = {
                "code": 200,
                "message": "处理完成。成功接收 1 条，失败 0 条",
                "success_count": 1,
                "fail_count": 0,
                "failed_details": [],
            }
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
        with _open_interface_request(request, timeout=15) as response:
            raw = _decode_interface_response(response.read(256 * 1024), response.headers)
            status_code = int(response.status)
        try:
            response_body: Any = json.loads(raw) if raw else {}
        except ValueError:
            response_body = raw
        business_code = response_body.get("code") if isinstance(response_body, dict) else None
        business_ok = business_code in {None, "", 200, "200"}
        error = ""
        if not business_ok:
            message = str(response_body.get("msg") or "接口未返回失败原因")
            error = f"接口已连通，但业务状态码为 {business_code}：{message}"
        return {
            "ok": 200 <= status_code < 300 and business_ok,
            "connection_ok": True,
            "business_ok": business_ok,
            "mode": "real", "status_code": status_code,
            "duration_ms": int((time.monotonic() - started) * 1000),
            "endpoint_url": endpoint_url, "request": request_body, "response": response_body, "error": error,
        }
    except HTTPError as exc:
        return {"ok": False, "mode": "real", "status_code": exc.code, "duration_ms": 0,
                "endpoint_url": endpoint_url, "request": request_body,
                "response": {}, "connection_ok": True, "business_ok": False,
                "error": f"接口已连通，但测试报文被业务接口拒绝（HTTP {exc.code}）。请使用实际订单数据验证业务处理。"}
    except (URLError, TimeoutError, OSError) as exc:
        return {"ok": False, "mode": "real", "status_code": None, "duration_ms": 0,
                "endpoint_url": endpoint_url, "request": request_body,
                "response": {}, "connection_ok": False, "business_ok": False,
                "error": f"接口请求失败：{str(exc)[:160]}"}


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
    "material_candidate_selected": "人工选择候选料号",
    "domestic_order_entry_mock": "生成订单接口", "domestic_order_entry_real": "生成订单接口",
    "order_info_query_mock": "查询订单信息接口", "order_info_query_real": "查询订单信息接口",
    "order_change_candidate_selected": "人工选择订单明细",
    "aps_order_demand_import_mock": "APS订单需求导入接口",
    "aps_order_demand_import_real": "APS订单需求导入接口",
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
        order_groups = conn.execute(
            """SELECT groups.id,groups.order_number
               FROM order_entry_template_groups groups
               JOIN order_entry_templates template ON template.id=groups.template_id
               WHERE template.case_id=? AND template.employee_id=?""",
            (case_id, employee_id),
        ).fetchall()
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
    order_number_by_group_id = {int(row["id"]): str(row["order_number"] or "待分配") for row in order_groups}
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
        item["interface_label"] = {
            "material_batch_query": "批量料号查询",
            "domestic_order_entry": "生成订单",
            "order_info_query": "查询订单信息",
        }.get(str(item.get("interface_key") or ""), "接口调用")
        item["mode_label"] = "Mock" if item.get("is_mock") else "真实接口"
        item["outcome_label"] = "成功" if item.get("status") == "success" else "失败"
        item["order_group_number"] = order_number_by_group_id.get(int(item.get("order_group_id") or 0), "")
        item["summary"] = _call_summary(item)
        if item["order_group_number"]:
            item["summary"] = f"{item['order_group_number']}：{item['summary']}"
        item["request"] = _redact_audit_payload(item["request"])
        item["response"] = _redact_audit_payload(item["response"])
        config_id = item.get("interface_config_id")
        snapshot = config_snapshots.get((int(config_id), int(item.get("config_version") or 1)), {}) if config_id else {}
        config = snapshot or config_by_id.get(int(config_id), {}) if config_id else {}
        item["endpoint_url"] = _endpoint_url(config) if config else "—"
        item["method"] = str(config.get("method") or "POST") if config else "—"
        call_rows.append(item)
    return {"template_id": template_id, "events": event_rows, "changes": changes[:100], "calls": call_rows}


def list_nyeos_order_numbers(case_ids: list[int], employee_id: str) -> dict[int, str]:
    ids = sorted({int(case_id) for case_id in case_ids if int(case_id) > 0})
    if not ids:
        return {}
    placeholders = ",".join("?" for _ in ids)
    with db_cursor() as conn:
        successful_call_rows = conn.execute(
            f"""SELECT id FROM order_interface_call_logs
                WHERE employee_id=? AND interface_key='domestic_order_entry'
                  AND status='success' AND case_id IN ({placeholders})""",
            (employee_id, *ids),
        ).fetchall()
        rows = conn.execute(
            f"""SELECT case_id,detail_json FROM order_entry_detail_events
                WHERE employee_id=? AND event_type='domestic_order_entry_real'
                  AND case_id IN ({placeholders})
                ORDER BY id DESC""",
            (employee_id, *ids),
        ).fetchall()
    successful_call_ids = {int(row["id"]) for row in successful_call_rows}
    result: dict[int, str] = {}
    for row in rows:
        case_id = int(row["case_id"])
        detail = _json(row["detail_json"], {})
        if int(detail.get("call_id") or 0) not in successful_call_ids:
            continue
        entry_no = str(detail.get("entry_no") or "").strip()
        if entry_no:
            result.setdefault(case_id, entry_no)
    return result


def list_erp_order_numbers(case_ids: list[int], employee_id: str) -> dict[int, str]:
    """Return ERP order numbers from the latest successful generation response."""
    ids = sorted({int(case_id) for case_id in case_ids if int(case_id) > 0})
    if not ids:
        return {}
    placeholders = ",".join("?" for _ in ids)
    with db_cursor() as conn:
        rows = conn.execute(
            f"""SELECT case_id,response_json FROM order_interface_call_logs
                WHERE employee_id=? AND interface_key='domestic_order_entry'
                  AND status='success' AND case_id IN ({placeholders})
                ORDER BY id DESC""",
            (employee_id, *ids),
        ).fetchall()
    result: dict[int, str] = {}
    for row in rows:
        case_id = int(row["case_id"])
        if case_id in result:
            continue
        response = _json(row["response_json"], {})
        data = response.get("data") if isinstance(response.get("data"), dict) else {}
        records = [item for item in data.get("data") or [] if isinstance(item, dict)]
        erp_map = data.get("erpOrderMap") if isinstance(data.get("erpOrderMap"), dict) else {}
        numbers = []
        for item in records:
            value = str(item.get("scta39") or erp_map.get(str(item.get("sctaCode") or "")) or "").strip()
            if value and value not in numbers:
                numbers.append(value)
        if numbers:
            result[case_id] = "、".join(numbers)
    return result


MATERIAL_STATUS_LABELS = {
    "pending": "待查询", "waiting_callback": "创建料号中", "requerying": "正在获取新料号",
    "resolved": "已回填", "manual_resolved": "人工已填写", "failed": "查询异常",
}

_LAYOUT_STRUCTURE_PATTERN = re.compile(
    r"(?:\(\s*(?P<parenthesized>\d{3,4}\s*[xX*×]\s*\d+"
    r"(?:\s*\+\s*\d{3,4}\s*[xX*×]\s*\d+)*)\s*\)"
    r"|(?P<compound>\d{3,4}\s*[xX*×]\s*\d+"
    r"(?:\s*\+\s*\d{3,4}\s*[xX*×]\s*\d+)+))"
)


def _extract_layout_structure(customer_code: Any, product_type: Any, customer_spec: Any) -> str:
    """Prefer a customer's configured structure position, then a conservative stackup pattern."""
    if _material_category_code(str(product_type or "")) != "718":
        return ""
    try:
        configured = extract_structure_from_customer_spec(customer_code, product_type, customer_spec)
    except Exception:
        # The optional customer-spec mapping data may not exist in historical
        # databases; new-material creation must still be able to proceed.
        configured = ""
    if configured:
        return configured
    matched = _LAYOUT_STRUCTURE_PATTERN.search(str(customer_spec or ""))
    return re.sub(r"\s+", "", (matched.group("parenthesized") or matched.group("compound"))) if matched else ""


def _material_request_item(line_no: int, values: dict[str, Any]) -> dict[str, Any]:
    return {
        "line_no": line_no,
        "material_status": str(values.get("material_status") or "查询"),
        "product_type": str(values.get("product_type") or ""),
        "product_name": str(values.get("product_name") or ""),
        "origin": str(values.get("origin") or ""),
        "adhesive_code": str(values.get("adhesive_code") or ""),
        "customer_product_code": str(values.get("customer_product_code") or ""),
        "customer_spec": str(values.get("customer_spec") or ""),
        "customer_spec_match": str(values.get("customer_spec_match") or ""),
        "layout_structure": str(values.get("layout_structure") or ""),
        "thickness_description": str(values.get("thickness_description") or ""),
        "special_requirements": str(values.get("special_requirements") or ""),
    }


def _material_query_is_already_filled(values: dict[str, Any]) -> bool:
    """Skip only lines with a resolved material code and product name."""
    placeholders = {"创建料号中", "创建品名中"}
    product_code = str(values.get("product_code") or "").strip()
    product_name = str(values.get("product_name") or "").strip()
    return bool(product_code and product_name and product_code not in placeholders and product_name not in placeholders)


def _insert_call_log(conn: Any, *, case_id: int, template_id: int, employee_id: str,
                     config: dict[str, Any], status: str, request_payload: dict[str, Any],
                     response_payload: dict[str, Any], triggered_by: str, error_message: str = "",
                     is_mock: bool = True, http_status: int | None = None,
                     duration_ms: int | None = 1, interface_key: str = "material_batch_query",
                     order_group_id: int | None = None) -> int:
    if is_mock and http_status is None:
        http_status = 200 if status == "success" else 422
    cursor = conn.execute(
        """INSERT INTO order_interface_call_logs
           (case_id,template_id,order_group_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
            status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (case_id, template_id, order_group_id, employee_id, int(config["id"]), interface_key,
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
                            factory_part_no: str, product_name: str, correlation_id: str, old_product_name: str = "",
                            source_label: str = "料号查询接口（Mock）") -> list[dict[str, Any]]:
    """Fill blanks and replace the interface's temporary creation message."""
    before = dict(values)
    sources_row = conn.execute(
        "SELECT sources_json FROM order_entry_template_lines WHERE template_id=? AND line_no=?", (template_id, line_no)
    ).fetchone()
    sources = _json(sources_row["sources_json"] if sources_row else "", {})
    changes: list[dict[str, Any]] = []
    for field, value in (
        ("product_code", factory_part_no),
        ("product_name", product_name),
        ("old_product_name", old_product_name),
    ):
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
            {"factory_part_no": f"MOCK-{code[-6:]}", "product_name": f"Mock 品名 {code[-4:]}", "old_product_name": f"Mock 旧品名 {code[-4:]}"},
            {"factory_part_no": f"MOCK-{code[-6:]}-ALT", "product_name": f"Mock 品名 {code[-4:]} 备选", "old_product_name": f"Mock 旧品名 {code[-4:]} 备选"},
        ]
        return {"line_no": line_no, "status": "matched", **candidates[0], "candidates": candidates,
                "matched_spec": item["customer_spec_match"] or item["customer_spec"],
                "message": "Mock 已命中多个料号，可在产品编号或品名中联动选择。"}
    if line_no % 3 == 2:
        return {"line_no": line_no, "status": "creating", "factory_part_no": "创建料号中", "product_name": "创建料号中",
                "external_task_id": f"MOCK-CREATE-{line_no}", "message": "Mock 未找到料号，已发起创建；等待对方回调。"}
    return {"line_no": line_no, "status": "failed", "message": "Mock 查询异常：请检查客户产品编号、客户规格和规格匹配。"}


def build_material_query(
    case_id: int, employee_id: str, triggered_by: str, *, line_nos: set[int] | None = None,
    group_key: str = "",
) -> dict[str, Any]:
    config = get_interface_config("material_batch_query")
    if not config or not config.get("enabled"):
        raise ValueError("批量料号查询接口未启用")
    if str(config.get("mode") or "mock") == "real":
        return build_material_query_real(
            case_id, employee_id, triggered_by, config=config, line_nos=line_nos, group_key=group_key,
        )
    return build_material_query_mock(
        case_id, employee_id, triggered_by, line_nos=line_nos, group_key=group_key,
    )


class MaterialNameValidationRequired(ValueError):
    def __init__(self, results: list[dict[str, str]]):
        super().__init__("新品名校验存在差异，请确认后再提交新建料号。")
        self.results = results


MATERIAL_NAME_VALIDATION_IGNORED_POSITIONS = {
    "pp": {13, 14, 24, 25},
    "board": {9, 10, 15, 25, 26},
}


def _material_name_validation_key(product: str, value: str) -> str:
    ignored = MATERIAL_NAME_VALIDATION_IGNORED_POSITIONS.get(product, set())
    return "".join(character for position, character in enumerate(value.strip(), start=1) if position not in ignored)


def validate_material_creation_names(lines: list[dict[str, Any]]) -> list[dict[str, str]]:
    """Validate dialog names with the maintained deterministic new-name rules."""
    try:
        mappings = product_name_service.mappings(enabled_only=True)
        mapping_error = ""
    except Exception as exc:
        # A missing or unavailable conversion rule still needs an explicit user decision.
        mappings = []
        mapping_error = f"新品名转换规则不可用：{exc}"
    results: list[dict[str, str]] = []
    for raw in lines:
        line_no = str(raw.get("line_no") or "").strip()
        customer_spec = str(raw.get("customer_spec") or "").strip()
        current_name = str(raw.get("product_name") or "").strip()
        category_code = _material_category_code(str(raw.get("product_type") or ""))
        product = "pp" if category_code == "698" else ("board" if category_code == "718" else "")
        result = {
            "line_no": line_no,
            "customer_spec": customer_spec,
            "current_product_name": current_name,
            "converted_product_name": "",
            "status": "failed",
            "message": "",
        }
        if not product:
            result["message"] = "产品类型无法判断，无法转换新品名。"
        elif not customer_spec:
            result["message"] = "缺少客户规格，无法转换新品名。"
        elif mapping_error:
            result["message"] = mapping_error
        else:
            try:
                recognized = extract_new_product_name(product, customer_spec, "", mappings)
                converted = str(product_name_service.build(product, recognized["values"], mappings)["code"] or "").strip()
                result["converted_product_name"] = converted
                if _material_name_validation_key(product, converted) == _material_name_validation_key(product, current_name):
                    result["status"] = "matched"
                    result["message"] = "新品名校验一致。"
                else:
                    result["status"] = "mismatch"
                    result["message"] = "转换新品名与当前新品名不一致。"
            except Exception as exc:
                result["message"] = str(exc)
        results.append(result)
    return results


def _save_material_creation_lines(
    case_id: int, employee_id: str, triggered_by: str, lines: list[dict[str, Any]],
    group_key: str = "", name_validation: list[dict[str, str]] | None = None,
    name_validation_confirmed: bool = False,
) -> tuple[int, set[int]]:
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成录单模板")
    if not isinstance(lines, list) or not lines:
        raise ValueError("当前没有需要新建料号的明细")
    submitted: dict[int, dict[str, Any]] = {}
    for raw in lines:
        try:
            line_no = int(raw.get("line_no") or 0)
        except (TypeError, ValueError):
            line_no = 0
        if line_no <= 0 or line_no in submitted:
            raise ValueError("新建料号明细的项次无效或重复")
        submitted[line_no] = raw
    changes: list[dict[str, Any]] = []
    now = utcnow()
    with db_cursor() as conn:
        group, group_rows = _domestic_group_data(conn, template_id, group_key)
        if str(group.get("status") or "") == "submitted":
            raise ValueError("当前 PO 已提交，不能新建料号")
        header = group["header"]
        customer_code = str(header.get("bill_to_customer_code") or "").strip()
        allowed_line_nos = {int(row["line_no"]) for row in group_rows}
        rows = conn.execute(
            """SELECT line_no,values_json,sources_json FROM order_entry_template_lines
               WHERE template_id=? AND line_no IN ({}) ORDER BY line_no""".format(
                ",".join("?" for _ in allowed_line_nos) or "NULL"
            ),
            (template_id, *sorted(allowed_line_nos)),
        ).fetchall()
        stored = {int(row["line_no"]): row for row in rows}
        for line_no, raw in submitted.items():
            row = stored.get(line_no)
            if not row:
                raise ValueError(f"第 {line_no} 项不存在")
            values = _json(row["values_json"], {})
            if str(values.get("product_code") or "").strip():
                raise ValueError(f"第 {line_no} 项已有产品编号，不能提交新建料号")
            task = conn.execute(
                "SELECT result_json FROM order_material_resolution_tasks WHERE template_id=? AND line_no=?",
                (template_id, line_no),
            ).fetchone()
            if task and _material_candidates(_json(task["result_json"], {})):
                raise ValueError(f"第 {line_no} 项已有候选料号，请先选择候选，不能重复新建")
            product_type = str(raw.get("product_type", values.get("product_type")) or "").strip()
            customer_spec = str(raw.get("customer_spec", values.get("customer_spec")) or "").strip()
            manual_layout = str(raw.get("layout_structure") or "").strip()
            automatic_layout = _extract_layout_structure(customer_code, product_type, customer_spec)
            updates = {
                "material_status": "新增",
                "product_name": (
                    str(raw.get("product_name") or "").strip()
                    if "product_name" in raw
                    else str(values.get("product_name") or "").strip()
                ),
                "customer_product_code": str(raw.get("customer_product_code", values.get("customer_product_code")) or "").strip(),
                "customer_spec": customer_spec,
                "product_type": product_type,
                "origin": str(raw.get("origin", values.get("origin")) or "").strip(),
                # 胶系编码仍由其他流程维护；客户规格匹配在新建料号弹窗中可修改。
                "adhesive_code": str(values.get("adhesive_code") or "").strip(),
                "customer_spec_match": str(raw.get(
                    "customer_spec_match", values.get("customer_spec_match"),
                ) or "").strip(),
                "layout_structure": manual_layout or automatic_layout,
                "thickness_description": str(raw.get("thickness_description") or "").strip(),
                "special_requirements": str(raw.get("special_requirements") or "").strip(),
            }
            if not updates["customer_product_code"]:
                raise ValueError(f"第 {line_no} 项缺少客户产品编号")
            if not updates["customer_spec"]:
                raise ValueError(f"第 {line_no} 项缺少客户规格")
            sources = _json(row["sources_json"], {})
            for field, value in updates.items():
                before = str(values.get(field) or "")
                values[field] = value
                if before != value:
                    sources[field] = {
                        "label": "客户规格自动提取" if field == "layout_structure" and not manual_layout else "人工修改",
                        "reference": "客户规格" if field == "layout_structure" and not manual_layout else "新建料号弹窗",
                    }
                    changes.append({"field": field, "before": before, "after": value, "line_no": line_no})
            conn.execute(
                """UPDATE order_entry_template_lines SET values_json=?,sources_json=?,updated_at=?
                   WHERE template_id=? AND line_no=?""",
                (json.dumps(values, ensure_ascii=False), json.dumps(sources, ensure_ascii=False), now, template_id, line_no),
            )
        conn.execute("UPDATE order_entry_templates SET updated_at=? WHERE id=?", (now, template_id))
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="material_create_prepare", title="确认新建料号数据",
            detail={
                "changes": changes,
                "line_nos": sorted(submitted),
                "name_validation": name_validation or [],
                "name_validation_confirmed": name_validation_confirmed,
            }, operated_by=triggered_by,
        )
    return template_id, set(submitted)


def build_material_creation(
    case_id: int, employee_id: str, triggered_by: str, lines: list[dict[str, Any]],
    group_key: str = "", confirm_name_validation: bool = False,
) -> dict[str, Any]:
    """Persist the creation dialog and submit only blank material rows."""
    if is_domestic_order_entry_completed(case_id, employee_id):
        raise ValueError("内销录单已完成，不能再次提交新建料号")
    config = get_interface_config("material_batch_query")
    if not config or not config.get("enabled"):
        raise ValueError("料号查询接口未启用")
    name_validation = validate_material_creation_names(lines)
    requires_confirmation = any(item["status"] != "matched" for item in name_validation)
    if requires_confirmation and not confirm_name_validation:
        raise MaterialNameValidationRequired(name_validation)
    _template_id, line_nos = _save_material_creation_lines(
        case_id, employee_id, triggered_by, lines, group_key,
        name_validation=name_validation,
        name_validation_confirmed=requires_confirmation,
    )
    if str(config.get("mode") or "mock") == "real":
        return build_material_query_real(
            case_id, employee_id, triggered_by, config=config, line_nos=line_nos,
            create_mode=True, group_key=group_key,
        )
    result = build_material_query_mock(
        case_id, employee_id, triggered_by, line_nos=line_nos, force_create=True,
        group_key=group_key,
    )
    result["mode"] = "mock"
    return result


def save_material_creation(
    case_id: int, employee_id: str, triggered_by: str, lines: list[dict[str, Any]],
    group_key: str = "",
) -> dict[str, Any]:
    """Save new-material form values without sending a creation request."""
    if is_domestic_order_entry_completed(case_id, employee_id):
        raise ValueError("内销录单已完成，不能再保存新建料号数据")
    template_id, line_nos = _save_material_creation_lines(
        case_id, employee_id, triggered_by, lines, group_key,
    )
    return {"template_id": template_id, "line_nos": sorted(line_nos)}


def _material_category_code(product_type: str) -> str:
    value = str(product_type or "").strip()
    normalized = value.upper()
    if value == "半固化片" or normalized in {"PP", "PREPREG", "698", "1"}:
        return "698"
    if value in {"基板", "板材", "覆铜板", "铜箔基板"} or normalized in {"CCL", "FR4", "718", "2"}:
        return "718"
    return value


MATERIAL_ORIGIN_CODES = {"上海": "1", "江西": "2", "江苏": "3"}


def _material_origin_code(value: Any) -> str:
    normalized = str(value or "").strip()
    return MATERIAL_ORIGIN_CODES.get(normalized, normalized if normalized in {"1", "2", "3"} else "")


def _real_material_request_item(item: dict[str, Any], *, create: bool = False) -> dict[str, str]:
    result = {
        "categoryCode": _material_category_code(str(item.get("product_type") or "")),
        "customerMaterialNo": str(item.get("customer_product_code") or "").strip(),
        "origin": _material_origin_code(item.get("origin")),
    }
    if create:
        result.update({
            "customerSpec": str(item.get("customer_spec_match") or "").strip(),
            "oriCustomerSpec": str(item.get("customer_spec") or "").strip(),
            "newProductName": str(item.get("product_name") or "").strip(),
            "newFlag": "Y",
        })
        for source, target in (
            ("layout_structure", "layoutStructure"),
            ("thickness_description", "thicknessDescription"),
            ("special_requirements", "specialRequirements"),
        ):
            value = str(item.get(source) or "").strip()
            if value:
                result[target] = value
    return result


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
        with _open_interface_request(request, timeout=int(config.get("timeout_seconds") or 15)) as response:
            raw = _decode_interface_response(response.read(1024 * 1024), response.headers)
            status_code = int(response.status)
    except HTTPError as exc:
        raw = _decode_interface_response(exc.read(1024 * 1024), exc.headers)
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


ORDER_CHANGE_MATCH_LABELS = {
    "unqueried": "未查询", "matched": "已匹配", "multiple": "多个匹配", "unmatched": "未匹配",
}

APS_ORDER_CHANGE_TYPES = {
    "交期变更", "品名变更", "大板料号变更", "数量变更", "备注修改",
    "订单取消", "订单不取消", "客户信息更改", "欠数补料", "特采交期调整",
}


def _match_text(value: Any) -> str:
    return str(value or "").strip().casefold()


def _match_spec(value: Any) -> str:
    """Compare customer specifications without PDF line-wrap whitespace."""
    return re.sub(r"\s+", "", _match_text(value))


def _match_decimal(value: Any) -> Decimal | None:
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return None


def _order_change_match_input(values: dict[str, Any]) -> dict[str, str]:
    return {
        "customer_order_number": str(values.get("customer_order_number") or "").strip(),
        "customer_product_code": str(values.get("customer_product_code") or "").strip(),
        "customer_spec": str(values.get("customer_spec") or "").strip(),
        "line_no": str(values.get("line_no") or "").strip(),
        "quantity": str(values.get("quantity") or "").strip(),
    }


def _flatten_order_info_candidates(response_payload: dict[str, Any]) -> list[dict[str, Any]]:
    data = response_payload.get("data") if isinstance(response_payload.get("data"), dict) else {}
    candidates: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str]] = set()
    for order in data.get("orderList") or []:
        if not isinstance(order, dict):
            continue
        header = {key: order.get(key) for key in ("scta01", "scta11", "scta38", "scta39", "acsn")}
        for detail in order.get("sctbList") or []:
            if not isinstance(detail, dict):
                continue
            candidate = {**header, **detail}
            candidate["account_set"] = _order_account_set(candidate.get("acsn"))
            candidate["customer_order_number"] = str(detail.get("sctb15") or order.get("scta38") or "").strip()
            key = tuple(_match_text(candidate.get(field)) for field in ("scta39", "sctb02", "sctb35", "customer_order_number"))
            if key in seen:
                continue
            seen.add(key)
            candidate["candidate_key"] = "|".join(key)
            candidates.append(candidate)
    return candidates


def _match_order_change_line(values: dict[str, Any], candidates: list[dict[str, Any]]) -> dict[str, Any]:
    match_input = _order_change_match_input(values)
    order_no = _match_text(match_input["customer_order_number"])
    customer_part = _match_text(match_input["customer_product_code"])
    customer_spec = _match_spec(match_input["customer_spec"])
    candidate_order_no = lambda item: _match_text(
        item.get("customer_order_number") or item.get("sctb15") or item.get("scta38")
    )
    base = [
        item for item in candidates
        if candidate_order_no(item) == order_no
        and _match_text(item.get("sctb14")) == customer_part
    ] if order_no and customer_part else []
    levels = (
        ("order_part_item", [item for item in base if _match_text(item.get("sctb35")) == _match_text(match_input["line_no"])]),
        ("order_part_quantity", [item for item in base if _match_decimal(item.get("sctb05")) is not None and _match_decimal(item.get("sctb05")) == _match_decimal(match_input["quantity"])]),
        ("order_part", base),
    )
    for level, matches in levels:
        if len(matches) == 1:
            return {"status": "matched", "match_level": level, "candidates": matches, "selected": matches[0]}
    final = levels[-1][1]
    if not final and order_no and customer_spec and _match_decimal(match_input["quantity"]) is not None:
        spec_quantity = [
            item for item in candidates
            if candidate_order_no(item) == order_no
            and _match_spec(item.get("sctb36")) == customer_spec
            and _match_decimal(item.get("sctb05")) == _match_decimal(match_input["quantity"])
        ]
        if len(spec_quantity) == 1:
            return {
                "status": "matched", "match_level": "order_spec_quantity",
                "candidates": spec_quantity, "selected": spec_quantity[0],
            }
        if spec_quantity:
            return {
                "status": "multiple", "match_level": "order_spec_quantity",
                "candidates": spec_quantity, "selected": {},
            }
    return {
        "status": "multiple" if len(final) > 1 else "unmatched",
        "match_level": "order_part" if final else "",
        "candidates": final,
        "selected": {},
    }


def _store_order_change_matches(
    conn: Any, *, case_id: int, template_id: int, employee_id: str,
    call_id: int, response_payload: dict[str, Any],
) -> list[dict[str, Any]]:
    rows = conn.execute(
        "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
        (template_id,),
    ).fetchall()
    candidates = _flatten_order_info_candidates(response_payload)
    now = utcnow()
    results = []
    for row in rows:
        values = _json(row["values_json"], {})
        matched = _match_order_change_line(values, candidates)
        match_input = _order_change_match_input(values)
        conn.execute(
            """INSERT INTO order_change_line_matches
               (case_id,template_id,employee_id,line_no,status,match_level,input_json,candidates_json,
                selected_candidate_json,query_call_log_id,selected_by,selected_at,created_at,updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(template_id,line_no) DO UPDATE SET
                 status=excluded.status,match_level=excluded.match_level,input_json=excluded.input_json,
                 candidates_json=excluded.candidates_json,selected_candidate_json=excluded.selected_candidate_json,
                 query_call_log_id=excluded.query_call_log_id,selected_by=excluded.selected_by,
                 selected_at=excluded.selected_at,updated_at=excluded.updated_at""",
            (
                case_id, template_id, employee_id, int(row["line_no"]), matched["status"], matched["match_level"],
                json.dumps(match_input, ensure_ascii=False), json.dumps(matched["candidates"], ensure_ascii=False),
                json.dumps(matched["selected"], ensure_ascii=False), call_id,
                "system" if matched["status"] == "matched" else "", now if matched["status"] == "matched" else None,
                now, now,
            ),
        )
        results.append({"line_no": int(row["line_no"]), **matched, "input": match_input})
    return results


def invalidate_changed_order_matches(conn: Any, template_id: int, lines: list[dict[str, Any]]) -> None:
    current = {int(item["values"]["line_no"]): _order_change_match_input(item["values"]) for item in lines}
    rows = conn.execute(
        "SELECT id,line_no,input_json FROM order_change_line_matches WHERE template_id=?", (template_id,),
    ).fetchall()
    for row in rows:
        if current.get(int(row["line_no"])) != _json(row["input_json"], {}):
            conn.execute("DELETE FROM order_change_line_matches WHERE id=?", (int(row["id"]),))


def get_order_change_matches(case_id: int, template_id: int, employee_id: str) -> dict[int, dict[str, Any]]:
    with db_cursor() as conn:
        rows = conn.execute(
            """SELECT * FROM order_change_line_matches
               WHERE case_id=? AND template_id=? AND employee_id=? ORDER BY line_no""",
            (case_id, template_id, employee_id),
        ).fetchall()
    result = {}
    for row in rows:
        item = _row(row)
        item["input"] = _json(item.pop("input_json"), {})
        item["candidates"] = _json(item.pop("candidates_json"), [])
        item["selected_candidate"] = _json(item.pop("selected_candidate_json"), {})
        item["label"] = ORDER_CHANGE_MATCH_LABELS.get(item["status"], item["status"])
        result[int(item["line_no"])] = item
    return result


def select_order_change_candidate(
    case_id: int, template_id: int, employee_id: str, line_no: int, candidate_key: str, selected_by: str,
) -> dict[str, Any]:
    with db_cursor() as conn:
        row = conn.execute(
            """SELECT * FROM order_change_line_matches
               WHERE case_id=? AND template_id=? AND employee_id=? AND line_no=?""",
            (case_id, template_id, employee_id, line_no),
        ).fetchone()
        if not row:
            raise ValueError("该明细尚未查询到可选订单")
        candidates = _json(row["candidates_json"], [])
        selected = next((item for item in candidates if item.get("candidate_key") == candidate_key), None)
        if not selected:
            raise ValueError("所选订单明细已失效，请重新查询")
        now = utcnow()
        conn.execute(
            """UPDATE order_change_line_matches SET status='matched',selected_candidate_json=?,
               selected_by=?,selected_at=?,updated_at=? WHERE id=?""",
            (json.dumps(selected, ensure_ascii=False), selected_by, now, now, int(row["id"])),
        )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="order_change_candidate_selected", title="已选择修改订单对应明细",
            detail={"line_no": line_no, "erp_order_number": selected.get("scta39"), "candidate": selected},
            operated_by=selected_by,
        )
    return selected


def _query_order_info(
    case_id: int, template_id: int, employee_id: str, triggered_by: str, order_numbers: list[Any],
    *, persist_matches: bool, event_context: str = "reply",
) -> dict[str, Any]:
    """Query existing orders and optionally persist modification-line matches.

    The modification-template page deliberately does not send its other
    columns as filters. They remain editable business data, while this lookup
    is only intended to retrieve the current ERP order by customer order no.
    """
    config = get_interface_config("order_info_query")
    if not config or not config.get("enabled"):
        raise ValueError("查询订单信息接口未启用")
    numbers = list(dict.fromkeys(str(value or "").strip() for value in order_numbers if str(value or "").strip()))
    if not numbers:
        raise ValueError("请先填写客户订单号后再查询")
    request_payload = {"orderNumberList": numbers}
    mode = str(config.get("mode") or "mock")

    if mode == "mock":
        with db_cursor() as conn:
            template_row = conn.execute(
                "SELECT header_json FROM order_entry_templates WHERE id=?", (template_id,),
            ).fetchone()
            template_rows = conn.execute(
                "SELECT values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
                (template_id,),
            ).fetchall()
        template_header = _json(template_row["header_json"], {}) if template_row else {}
        header_order_number = _match_text(template_header.get("customer_order_number"))
        template_values = [_json(row["values_json"], {}) for row in template_rows]
        order_list = [
            {
                "scta01": f"MOCK-{index:04d}",
                "scta38": number,
                "scta39": f"220-MOCK-{index:04d}",
                "sctbList": [
                    {
                        "sctb02": f"MOCK-PART-{line_index:04d}", "sctb03": "Mock 品名规格",
                        "sctb04": values.get("line_no", ""),
                        "sctb14": values.get("customer_product_code", ""), "sctb15": number,
                        "sctb16": values.get("delivery_date", ""), "sctb17": values.get("delivery_date", ""),
                        "sctb35": values.get("line_no", ""), "sctb36": values.get("customer_spec", ""),
                        "sctb05": values.get("quantity", ""), "sctb23": values.get("quantity", ""),
                    }
                    for line_index, values in enumerate(template_values, start=1)
                    if (
                        _match_text(values.get("customer_order_number")) == _match_text(number)
                        or (
                            not _match_text(values.get("customer_order_number"))
                            and header_order_number == _match_text(number)
                        )
                    )
                ],
            }
            for index, number in enumerate(numbers, start=1)
        ]
        response_payload = {
            "code": 200,
            "msg": "Mock 查询成功",
            "data": {"orderList": order_list, "orderCount": len(order_list), "notFoundList": []},
        }
        with db_cursor() as conn:
            call_id = _insert_call_log(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                status="success", request_payload=request_payload, response_payload=response_payload,
                triggered_by=triggered_by, interface_key="order_info_query",
            )
            matches = _store_order_change_matches(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                call_id=call_id, response_payload=response_payload,
            ) if persist_matches else []
            is_recovery = event_context == "domestic_recovery"
            is_batch = event_context == "batch"
            record_order_detail_event(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                event_type=(
                    "order_info_query_mock" if persist_matches else
                    "order_info_query_batch_mock" if is_batch else
                    "domestic_order_entry_recovery_query_mock" if is_recovery else
                    "order_reply_info_query_mock"
                ),
                title=(
                    "查询订单信息（Mock）完成" if persist_matches else
                    "查询订单信息（Mock，分批）完成" if is_batch else
                    "提交录单超时后查询订单信息（Mock）完成" if is_recovery else
                    "回复邮件查询订单信息（Mock）完成"
                ),
                detail={"call_id": call_id, "order_count": len(order_list), "order_numbers": numbers,
                        "matched_count": sum(item["status"] == "matched" for item in matches)},
                operated_by=triggered_by,
            )
        return {"call_id": call_id, "status": "success", "mode": "mock", "response": response_payload,
                "matches": matches}

    try:
        http_status, response_payload, duration_ms = _post_json_endpoint(config, request_payload, "查询订单信息")
        status = "success" if http_status == 200 and int(response_payload.get("code") or 0) == 200 else "failed"
        error_message = "" if status == "success" else str(response_payload.get("msg") or f"接口返回 HTTP {http_status}")
    except ValueError as exc:
        http_status, response_payload, duration_ms = None, {}, None
        status, error_message = "failed", str(exc)

    with db_cursor() as conn:
        call_id = _insert_call_log(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
            status=status, request_payload=request_payload, response_payload=response_payload,
            triggered_by=triggered_by, error_message=error_message, is_mock=False,
            http_status=http_status, duration_ms=duration_ms, interface_key="order_info_query",
        )
        matches = _store_order_change_matches(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            call_id=call_id, response_payload=response_payload,
        ) if status == "success" and persist_matches else []
        is_recovery = event_context == "domestic_recovery"
        is_batch = event_context == "batch"
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type=(
                "order_info_query_real" if persist_matches else
                "order_info_query_batch_real" if is_batch else
                "domestic_order_entry_recovery_query_real" if is_recovery else
                "order_reply_info_query_real"
            ),
            title=(
                f"查询订单信息（真实接口）{'完成' if status == 'success' else '失败'}"
                if persist_matches else
                f"查询订单信息（真实接口，分批）{'完成' if status == 'success' else '失败'}"
                if is_batch else
                f"提交录单超时后查询订单信息（真实接口）{'完成' if status == 'success' else '失败'}"
                if is_recovery else
                f"回复邮件查询订单信息（真实接口）{'完成' if status == 'success' else '失败'}"
            ),
            detail={"call_id": call_id, "order_numbers": numbers, "error_message": error_message},
            operated_by=triggered_by,
        )
    if status != "success":
        raise ValueError(error_message)
    return {"call_id": call_id, "status": status, "mode": "real", "response": response_payload,
            "matches": matches}


def query_order_info(
    case_id: int, template_id: int, employee_id: str, triggered_by: str, order_numbers: list[Any],
) -> dict[str, Any]:
    """Query orders for the modification workflow and persist line matches."""
    numbers = list(dict.fromkeys(str(value or "").strip() for value in order_numbers if str(value or "").strip()))
    if len(numbers) <= ORDER_INFO_QUERY_BATCH_SIZE:
        return _query_order_info(
            case_id, template_id, employee_id, triggered_by, numbers,
            persist_matches=True,
        )

    batch_results = []
    failures = []
    for start in range(0, len(numbers), ORDER_INFO_QUERY_BATCH_SIZE):
        batch = numbers[start:start + ORDER_INFO_QUERY_BATCH_SIZE]
        try:
            batch_results.append(_query_order_info(
                case_id, template_id, employee_id, triggered_by, batch,
                persist_matches=False, event_context="batch",
            ))
        except ValueError as exc:
            failures.append({"batch": start // ORDER_INFO_QUERY_BATCH_SIZE + 1, "orders": batch, "message": str(exc)})
    if failures:
        failed_batches = "；".join(
            f"第 {item['batch']} 批（{len(item['orders'])} 个PO）：{item['message']}" for item in failures
        )
        raise ValueError(f"订单查询部分失败，未更新匹配结果：{failed_batches}")

    order_list: list[dict[str, Any]] = []
    not_found: list[Any] = []
    for item in batch_results:
        data = item.get("response", {}).get("data") if isinstance(item.get("response"), dict) else {}
        data = data if isinstance(data, dict) else {}
        order_list.extend(item for item in data.get("orderList") or [] if isinstance(item, dict))
        not_found.extend(data.get("notFoundList") or [])
    response_payload = {
        "code": 200,
        "msg": "分批查询成功",
        "data": {"orderList": order_list, "orderCount": len(order_list), "notFoundList": not_found},
    }
    call_ids = [int(item["call_id"]) for item in batch_results]
    with db_cursor() as conn:
        matches = _store_order_change_matches(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            call_id=call_ids[-1], response_payload=response_payload,
        )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="order_info_query_batched",
            title="查询订单信息（分批）完成",
            detail={"batch_count": len(batch_results), "order_count": len(numbers), "call_ids": call_ids,
                    "matched_count": sum(item["status"] == "matched" for item in matches)},
            operated_by=triggered_by,
        )
    return {
        "call_id": call_ids[-1], "call_ids": call_ids, "status": "success",
        "mode": batch_results[0].get("mode", "real"), "response": response_payload, "matches": matches,
    }


def _expected_arrival_date(customer_demand_date: Any, transit_days: Any) -> str:
    demand_date = normalize_date(customer_demand_date)
    days_text = str(transit_days or "").strip()
    if not demand_date or not re.fullmatch(r"\d+", days_text):
        return ""
    try:
        return (datetime.fromisoformat(demand_date) + timedelta(days=int(days_text))).date().isoformat()
    except (OverflowError, ValueError):
        return ""


def _order_info_display_result(
    response_payload: dict[str, Any], *, transit_days: Any = None,
) -> dict[str, Any]:
    data = response_payload.get("data") if isinstance(response_payload.get("data"), dict) else {}
    orders: list[dict[str, Any]] = []
    for order in data.get("orderList") or []:
        if not isinstance(order, dict):
            continue
        details = []
        for detail in order.get("sctbList") or []:
            if not isinstance(detail, dict):
                continue
            details.append({
                "item_no": detail.get("sctb35", ""),
                "factory_part_code": detail.get("sctb02", ""),
                "product_spec": detail.get("sctb03", ""),
                "customer_part_code": detail.get("sctb14", ""),
                "customer_order_number": detail.get("sctb15", ""),
                "customer_spec": detail.get("sctb36", ""),
                "quantity": detail.get("sctb05", ""),
                "outstanding_quantity": detail.get("sctb23", ""),
                "tax_price": detail.get("sctb06", ""),
                "untaxed_price": detail.get("sctb07", ""),
                "demand_date": detail.get("sctb16", ""),
                "expected_ship_date": detail.get("sctb17", ""),
                "expected_arrival_date": _expected_arrival_date(detail.get("sctb16"), transit_days),
                "closing_code": detail.get("sctb30", ""),
                "factory": detail.get("sctb43", ""),
            })
        orders.append({
            "order_number": order.get("scta01", ""),
            "erp_order_number": order.get("scta39", ""),
            "customer_order_number": order.get("scta38", ""),
            "ship_to_customer_id": order.get("scta11", ""),
            "organization": order.get("acsn", ""),
            "details": details,
        })
    not_found = [str(value) for value in data.get("notFoundList") or [] if str(value).strip()]
    try:
        order_count = int(data.get("orderCount"))
    except (TypeError, ValueError):
        order_count = len(orders)
    return {
        "message": str(response_payload.get("msg") or "查询完成"),
        "order_count": order_count,
        "orders": orders,
        "not_found": not_found,
    }


def query_order_info_readonly(
    case_id: int, template_id: int, employee_id: str, triggered_by: str, order_numbers: list[Any],
    *, transit_days: Any = None,
) -> dict[str, Any]:
    """Query orders for reply review without changing modification matches."""
    result = _query_order_info(
        case_id, template_id, employee_id, triggered_by, order_numbers,
        persist_matches=False,
    )
    return {
        "call_id": result["call_id"],
        "status": result["status"],
        "mode": result["mode"],
        **_order_info_display_result(result["response"], transit_days=transit_days),
    }


def query_order_info_reply_rows(
    case_id: int, template_id: int, employee_id: str, triggered_by: str, rows: list[dict[str, Any]],
    *, transit_days: Any = None,
) -> dict[str, Any]:
    """Look up reply-table rows without changing templates or change-order matches."""
    normalized_rows = []
    for index, raw in enumerate(rows):
        if not isinstance(raw, dict):
            continue
        values = {
            "row_id": str(raw.get("row_id") or index),
            "customer_order_number": _normalize_reply_order_number(raw.get("customer_order_number")),
            "customer_product_code": str(raw.get("customer_product_code") or "").strip(),
            "line_no": str(raw.get("line_no") or "").strip(),
            "quantity": str(raw.get("quantity") or "").strip(),
        }
        if values["customer_order_number"]:
            normalized_rows.append(values)
    if not normalized_rows:
        raise ValueError("回复表格中未找到客户 PO 号")

    result = _query_order_info(
        case_id, template_id, employee_id, triggered_by,
        [item["customer_order_number"] for item in normalized_rows],
        persist_matches=False,
    )
    candidates = _flatten_order_info_candidates(result["response"])
    row_matches = []
    counts = {"matched": 0, "multiple": 0, "unmatched": 0, "invalid_date": 0}
    for values in normalized_rows:
        matched = _match_order_change_line(values, candidates)
        selected = matched["selected"]
        delivery_reply = _expected_arrival_date(selected.get("sctb16"), transit_days) if selected else ""
        status = matched["status"]
        if status == "matched" and not delivery_reply:
            status = "invalid_date"
            reason = "未能根据客户需求日和运输天数计算交期"
        elif status == "matched":
            reason = ""
        elif status == "multiple":
            reason = f"匹配到 {len(matched['candidates'])} 条订单明细，无法自动回填"
        else:
            reason = "未找到匹配的订单明细"
        counts[status] += 1
        row_matches.append({
            "row_id": values["row_id"], "status": status,
            "match_level": matched["match_level"], "delivery_reply": delivery_reply,
            "reason": reason,
        })
    return {
        "call_id": result["call_id"], "status": result["status"], "mode": result["mode"],
        **_order_info_display_result(result["response"], transit_days=transit_days),
        "row_matches": row_matches, "match_counts": counts,
    }


def _normalize_reply_order_number(value: Any) -> str:
    """Match the PO cleanup used by entry extraction without importing its service."""
    text = str(value or "").strip()
    if not text or text.startswith("暂无PO号-"):
        return ""
    text = re.sub(r"^(?:建价|估价|报价|询价)\s*[:：-]?\s*", "", text, flags=re.I)
    match = re.search(r"(?i)(PO(?:[-_][A-Z0-9]+)+)", text)
    return match.group(1).upper().replace("_", "-") if match else text


def _aps_order_demand_request_payload(
    lines: list[dict[str, Any]], matches: dict[int, dict[str, Any]],
    alter_type: str, require_specification: str, creator_name: str, created_at: str,
) -> dict[str, list[dict[str, Any]]]:
    """Build APS demand records from the saved template and confirmed ERP matches."""
    alter_type = str(alter_type or "").strip()
    if alter_type not in APS_ORDER_CHANGE_TYPES:
        raise ValueError("请选择有效的变更类型")
    specification = str(require_specification or "").strip()
    creator_name = str(creator_name or "").strip()
    if not creator_name:
        raise ValueError("请先在账户与密码中维护用户名")
    items: list[dict[str, Any]] = []
    issues: list[str] = []
    for line in lines:
        line_no = int(line["line_no"])
        values = _json(line["values_json"], {})
        match = matches.get(line_no) or {}
        selected = match.get("selected_candidate") or {}
        erp_order_number = str(selected.get("scta39") or "").strip()
        item_no = str(selected.get("sctb35") or "").strip()
        account_set = _order_account_set(selected.get("acsn"))
        shipment_date = normalize_date(values.get("delivery_date"))
        if match.get("status") != "matched":
            issues.append(f"第 {line_no} 项尚未确认 ERP 匹配")
        if not erp_order_number:
            issues.append(f"第 {line_no} 项缺少 ERP订单号")
        if not item_no:
            issues.append(f"第 {line_no} 项缺少 ERP 项次")
        if not account_set:
            issues.append(f"第 {line_no} 项缺少或不支持 ERP 账套组织代码")
        if not shipment_date:
            issues.append(f"第 {line_no} 项客户需求日期格式无效")
        if match.get("status") == "matched" and erp_order_number and item_no and account_set and shipment_date:
            items.append({
                "require_shipment_date": shipment_date,
                "Order_Item_Account_Set_outer_key": f"{erp_order_number}_{item_no}_{account_set}",
                "alter_type": alter_type,
                "creator_name": creator_name,
                "require_specification": specification,
                "created_at": created_at,
            })
    if issues:
        raise ValueError("暂不能提交修改订单：" + "；".join(issues))
    if not items:
        raise ValueError("当前没有可提交的修改订单明细")
    return {"data": items}


def _aps_order_demand_result(response_payload: dict[str, Any], http_status: int, expected_count: int) -> tuple[bool, str]:
    try:
        code = int(response_payload.get("code") or 0)
        success_count = int(response_payload.get("success_count") or 0)
        fail_count = int(response_payload.get("fail_count") or 0)
    except (TypeError, ValueError):
        return False, "APS 返回的成功/失败条数格式无效"
    failed_details = response_payload.get("failed_details") or []
    if http_status == 200 and code == 200 and success_count == expected_count and fail_count == 0 and not failed_details:
        return True, str(response_payload.get("message") or f"成功导入 {success_count} 条订单需求")
    reasons = [str(item.get("reason") or "") for item in failed_details if isinstance(item, dict)]
    message = str(response_payload.get("message") or f"成功 {success_count} 条，失败 {fail_count} 条")
    return False, "；".join([message, *[reason for reason in reasons if reason]])


def submit_aps_order_demand_import(
    case_id: int, template_id: int, employee_id: str, triggered_by: str,
    alter_type: str, require_specification: str = "",
) -> dict[str, Any]:
    """Submit all confirmed order-change lines to APS and keep a complete audit trail."""
    with db_cursor() as conn:
        case = conn.execute(
            "SELECT status,workflow_stage FROM order_intake_cases WHERE id=? AND employee_id=? AND action_type='order_change'",
            (case_id, employee_id),
        ).fetchone()
        if not case:
            raise ValueError("修改订单案件不存在或无权操作")
        if str(case["status"] or "") == "pending_reply":
            raise ValueError("修改订单已提交 APS，不能重复提交")
        original_stage = str(case["workflow_stage"] or "")
        claimed = conn.execute(
            """UPDATE order_intake_cases SET workflow_stage='aps_submitting',updated_at=?
               WHERE id=? AND employee_id=? AND action_type='order_change'
                 AND status<>'pending_reply' AND workflow_stage<> 'aps_submitting'""",
            (utcnow(), case_id, employee_id),
        )
        if not claimed.rowcount:
            raise ValueError("APS 正在提交，请勿重复操作")

    def release_submission_claim() -> None:
        with db_cursor() as conn:
            conn.execute(
                """UPDATE order_intake_cases SET workflow_stage=?,updated_at=?
                   WHERE id=? AND employee_id=? AND action_type='order_change' AND status<>'pending_reply'""",
                (original_stage, utcnow(), case_id, employee_id),
            )

    config = get_interface_config("aps_order_demand_import")
    if not config or not config.get("enabled"):
        release_submission_claim()
        raise ValueError("APS订单需求导入接口未启用")
    try:
        with db_cursor() as conn:
            rows = conn.execute(
                "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
                (template_id,),
            ).fetchall()
        matches = get_order_change_matches(case_id, template_id, employee_id)
        account = get_user(employee_id)
        creator_name = str(account["display_name"] or "").strip() if account else ""
        created_at = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M:%S")
        request_payload = _aps_order_demand_request_payload(
            list(rows), matches, alter_type, require_specification, creator_name, created_at,
        )
        mode = str(config.get("mode") or "mock")

        if mode == "mock":
            response_payload = {
                "code": 200,
                "message": f"处理完成。成功接收 {len(request_payload['data'])} 条，失败 0 条",
                "success_count": len(request_payload["data"]),
                "fail_count": 0,
                "failed_details": [],
            }
            http_status, duration_ms = 200, 0
        else:
            http_status, response_payload, duration_ms = _post_json_endpoint(
                config, request_payload, "APS订单需求导入",
            )
    except ValueError as exc:
        with db_cursor() as conn:
            call_id = _insert_call_log(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                status="failed", request_payload=locals().get("request_payload", {}), response_payload={}, triggered_by=triggered_by,
                error_message=str(exc), is_mock=False, http_status=None, duration_ms=None,
                interface_key="aps_order_demand_import",
            )
            record_order_detail_event(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                event_type="aps_order_demand_import_real", title="APS订单需求导入（真实接口）失败",
                detail={"call_id": call_id, "error_message": str(exc)}, operated_by=triggered_by,
            )
        release_submission_claim()
        raise

    ok, message = _aps_order_demand_result(response_payload, http_status, len(request_payload["data"]))
    with db_cursor() as conn:
        call_id = _insert_call_log(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
            status="success" if ok else "failed", request_payload=request_payload, response_payload=response_payload,
            triggered_by=triggered_by, error_message="" if ok else message, is_mock=mode == "mock",
            http_status=http_status, duration_ms=duration_ms, interface_key="aps_order_demand_import",
        )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type=f"aps_order_demand_import_{mode}",
            title=f"APS订单需求导入（{'Mock' if mode == 'mock' else '真实接口'}）{'完成' if ok else '失败'}",
            detail={
                "call_id": call_id, "line_count": len(request_payload["data"]), "alter_type": alter_type,
                "require_specification": str(require_specification or "").strip(), "message": message,
            },
            operated_by=triggered_by,
        )
        if ok:
            case = conn.execute(
                "SELECT status FROM order_intake_cases WHERE id=? AND employee_id=? AND action_type='order_change'",
                (case_id, employee_id),
            ).fetchone()
            if case and str(case["status"] or "") != "pending_reply":
                now = utcnow()
                conn.execute(
                    "UPDATE order_intake_cases SET status='pending_reply',workflow_stage=?,updated_at=? WHERE id=? AND employee_id=?",
                    (original_stage, now, case_id, employee_id),
                )
                conn.execute(
                    """INSERT INTO order_intake_case_events
                       (case_id,employee_id,action,before_json,after_json,created_at)
                       VALUES (?,?,?,?,?,?)""",
                    (
                        case_id, employee_id, "aps_order_change_submitted",
                        json.dumps({"status": case["status"]}, ensure_ascii=False),
                        json.dumps({"status": "pending_reply"}, ensure_ascii=False), now,
                    ),
                )
        else:
            conn.execute(
                """UPDATE order_intake_cases SET workflow_stage=?,updated_at=?
                   WHERE id=? AND employee_id=? AND action_type='order_change' AND status<>'pending_reply'""",
                (original_stage, utcnow(), case_id, employee_id),
            )
    return {
        "call_id": call_id, "status": "success" if ok else "failed", "mode": mode,
        "message": message, "request": request_payload, "response": response_payload,
    }


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
                "old_product_name": str(hit.get("peag09") or "").strip(),
            }
            for hit in matched_hits
            if str(hit.get("peag01") or "").strip()
        ]
        if candidates:
            results.append({
                "line_no": item["line_no"], "status": "matched", **candidates[0],
                "candidates": candidates, "matched_spec": str(matched_hits[0].get("scca05") or ""),
                "message": f"真实接口命中 {len(candidates)} 个候选料号。",
            })
        elif index in errors_by_line:
            results.append({"line_no": item["line_no"], "status": "failed", "message": errors_by_line[index]})
        elif business_ok and (response_body.get("external_task_id") or response_body.get("pera01")):
            external_task_id = str(response_body.get("external_task_id") or response_body.get("pera01") or "")
            results.append({
                "line_no": item["line_no"], "status": "creating",
                "factory_part_no": "创建料号中", "product_name": "创建料号中",
                "external_task_id": external_task_id,
                "message": f"已提交新建料号，外部任务号 {external_task_id}。",
            })
        else:
            results.append({
                "line_no": item["line_no"], "status": "failed",
                "message": str(response_body.get("msg") or f"接口返回 HTTP {http_status}"),
            })
    return results


def build_material_query_real(
    case_id: int, employee_id: str, triggered_by: str, *, config: dict[str, Any] | None = None,
    line_nos: set[int] | None = None, create_mode: bool = False, group_key: str = "",
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
        group, rows = _domestic_group_data(conn, template_id, group_key)
    if str(group.get("status") or "") == "submitted":
        raise ValueError("当前 PO 已提交，不能再次请求料号查询接口")
    header = group["header"]
    customer_code = str(header.get("bill_to_customer_code") or "").strip()
    if not customer_code:
        raise ValueError("请先填写并保存账款客户编号")
    rows = [row for row in rows if line_nos is None or int(row["line_no"]) in line_nos]
    skipped_line_nos: list[int] = []
    if not create_mode:
        pending_rows = []
        for row in rows:
            if _material_query_is_already_filled(_json(row["values_json"], {})):
                skipped_line_nos.append(int(row["line_no"]))
            else:
                pending_rows.append(row)
        rows = pending_rows
    request_items = [_material_request_item(int(row["line_no"]), _json(row["values_json"], {})) for row in rows]
    values_by_line = {int(row["line_no"]): _json(row["values_json"], {}) for row in rows}
    for item in request_items:
        item["remark"] = str(values_by_line[item["line_no"]].get("remark") or "")
    if not request_items:
        return {"items": [], "status": "success", "mode": "real", "skipped_line_nos": skipped_line_nos}
    request_payload = {
        "customerCode": customer_code,
        "acsn": "NY01",
        "operatorCode": employee_id,
        "materialInfoList": [_real_material_request_item(item, create=create_mode) for item in request_items],
    }
    operation_label = "新建料号" if create_mode else "批量料号查询"
    try:
        http_status, response_body, duration_ms = _post_json_endpoint(config, request_payload, f"真实{operation_label}")
    except ValueError as exc:
        with db_cursor() as conn:
            call_id = _insert_call_log(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                status="failed", request_payload=request_payload, response_payload={}, triggered_by=triggered_by,
                error_message=str(exc), is_mock=False, http_status=None, duration_ms=None,
                order_group_id=int(group["id"]) if group.get("id") else None,
            )
            record_order_detail_event(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                event_type="material_create_real" if create_mode else "material_query_real",
                title=f"{operation_label}（真实接口）失败",
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
            order_group_id=int(group["id"]) if group.get("id") else None,
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
            if status in {"matched", "creating"} or response.get("product_name"):
                candidates = _material_candidates(response)
                changes = _backfill_material_line(
                    conn, template_id=template_id, line_no=item["line_no"], values=values,
                    factory_part_no=str(response.get("factory_part_no") or ""),
                    product_name=str(response.get("product_name") or ""),
                    old_product_name=str(response.get("old_product_name") or ""),
                    correlation_id=task["correlation_id"],
                    source_label=str(response.get("product_name_source_label") or "料号查询接口（真实）"),
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
            event_type="material_create_real" if create_mode else "material_query_real",
            title=f"{operation_label}（真实接口）完成",
            detail={"call_id": call_id, "changes": all_changes, "items": response_items}, operated_by=triggered_by,
        )
    return {
        "call_id": call_id, "items": response_items, "status": call_status, "mode": "real",
        "skipped_line_nos": skipped_line_nos,
    }


def build_material_query_mock(case_id: int, employee_id: str, triggered_by: str, scenario: str = "success",
                              *, line_nos: set[int] | None = None,
                              callback_requery: bool = False, force_create: bool = False,
                              group_key: str = "") -> dict[str, Any]:
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
        group, rows = _domestic_group_data(conn, template_id, group_key)
        if str(group.get("status") or "") == "submitted":
            raise ValueError("当前 PO 已提交，不能再次请求料号查询接口")
        selected = [row for row in rows if line_nos is None or int(row["line_no"]) in line_nos]
        skipped_line_nos: list[int] = []
        if not force_create:
            pending_rows = []
            for row in selected:
                if _material_query_is_already_filled(_json(row["values_json"], {})):
                    skipped_line_nos.append(int(row["line_no"]))
                else:
                    pending_rows.append(row)
            selected = pending_rows
        request_items = [_material_request_item(int(row["line_no"]), _json(row["values_json"], {})) for row in selected]
        if not request_items:
            return {"items": [], "status": "success", "mode": "mock", "skipped_line_nos": skipped_line_nos}
        if scenario in {"business_error", "timeout"}:
            message = "Mock 业务错误：订单明细校验未通过。" if scenario == "business_error" else "Mock 超时：请求未返回。"
            call_id = _insert_call_log(conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                                        status="failed", request_payload={"items": request_items}, response_payload={"items": [], "message": message},
                                        triggered_by=triggered_by, error_message=message,
                                        order_group_id=int(group["id"]) if group.get("id") else None)
            record_order_detail_event(conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                                      event_type="material_query_mock", title="批量料号查询（Mock）失败",
                                      detail={"call_id": call_id, "error_message": message}, operated_by=triggered_by)
            return {"call_id": call_id, "items": [], "status": "failed", "scheduled_task_ids": []}
        response_items = [
            {
                "line_no": item["line_no"], "status": "creating",
                "factory_part_no": "创建料号中", "product_name": "创建料号中",
                "external_task_id": f"MOCK-CREATE-{template_id}",
                "message": "Mock 新建料号已提交；等待对方回调。",
            }
            if force_create else _mock_material_response(item, callback_requery=callback_requery)
            for item in request_items
        ]
        call_id = _insert_call_log(conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                                    status="success", request_payload={"items": request_items}, response_payload={"items": response_items},
                                    triggered_by=triggered_by,
                                    order_group_id=int(group["id"]) if group.get("id") else None)
        all_changes: list[dict[str, Any]] = []
        for row, item, response in zip(selected, request_items, response_items):
            values = _json(row["values_json"], {})
            status = response["status"]
            task_status = "resolved" if status == "matched" else ("waiting_callback" if status == "creating" else "failed")
            task = _upsert_resolution_task(conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                                            line_no=item["line_no"], status=task_status, input_item=item,
                                            call_id=call_id, result=response)
            if status == "matched":
                candidates = _material_candidates(response)
                changes = _backfill_material_line(conn, template_id=template_id, line_no=item["line_no"], values=values,
                                                   factory_part_no=response["factory_part_no"],
                                                   product_name=response["product_name"],
                                                   correlation_id=task["correlation_id"],
                                                   old_product_name=str(response.get("old_product_name") or ""))
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
    return {
        "call_id": call_id, "items": response_items, "status": "success", "mode": "mock",
        "skipped_line_nos": skipped_line_nos,
    }


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
        rows = conn.execute("SELECT line_no,status,correlation_id,external_task_id,result_json,updated_at FROM order_material_resolution_tasks WHERE case_id=? AND employee_id=? ORDER BY line_no", (case_id, employee_id)).fetchall()
    items = []
    for row in rows:
        item = _row(row)
        item["result"] = _json(item.pop("result_json", ""), {})
        item["candidates"] = _material_candidates(item["result"])
        item["candidate_count"] = len(item["candidates"])
        item["selected_candidate"] = _selected_material_candidate(item["result"], item["candidates"])
        item["selection_required"] = item["candidate_count"] > 1 and not item["selected_candidate"]
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
            "old_product_name": str(raw.get("old_product_name") or raw.get("peag09") or "").strip(),
        }
        if not candidate["product_code"] or candidate["product_code"] in {"创建料号中", "创建品名中"}:
            continue
        key = (candidate["product_code"], candidate["product_name"], candidate["old_product_name"])
        if key == ("", "", "") or key in seen:
            continue
        seen.add(key)
        candidates.append(candidate)
    return candidates


def _selected_material_candidate(
    result: dict[str, Any], candidates: list[dict[str, str]] | None = None,
) -> dict[str, str] | None:
    raw = result.get("selected_candidate")
    if not isinstance(raw, dict):
        return None
    selected = {
        "product_code": str(raw.get("product_code") or "").strip(),
        "product_name": str(raw.get("product_name") or "").strip(),
        "old_product_name": str(raw.get("old_product_name") or "").strip(),
    }
    if not selected["product_code"] and not selected["product_name"]:
        return None
    candidates = candidates if candidates is not None else _material_candidates(result)
    return selected if selected in candidates else None


def select_material_candidate(
    case_id: int,
    employee_id: str,
    line_no: int,
    *,
    product_code: str,
    product_name: str,
) -> dict[str, Any]:
    """Persist one user-confirmed material candidate and update its template line."""
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成录单模板")
    with db_cursor() as conn:
        task = conn.execute(
            """SELECT * FROM order_material_resolution_tasks
               WHERE case_id=? AND template_id=? AND employee_id=? AND line_no=?""",
            (case_id, template_id, employee_id, int(line_no)),
        ).fetchone()
        if not task:
            raise ValueError("未找到本行的料号查询结果")
        task = _row(task)
        result = _json(task.get("result_json"), {})
        candidates = _material_candidates(result)
        if len(candidates) < 2:
            raise ValueError("本行没有需要确认的多个候选料号")
        selected = next(
            (
                candidate for candidate in candidates
                if candidate["product_code"] == str(product_code or "").strip()
                and candidate["product_name"] == str(product_name or "").strip()
            ),
            None,
        )
        if selected is None:
            raise ValueError("所选料号不在本次接口返回的候选列表中")
        line = conn.execute(
            "SELECT values_json,sources_json FROM order_entry_template_lines WHERE template_id=? AND line_no=?",
            (template_id, int(line_no)),
        ).fetchone()
        if not line:
            raise ValueError("录单模板明细不存在")
        values = _json(line["values_json"], {})
        sources = _json(line["sources_json"], {})
        changes = []
        for field, value in selected.items():
            before = str(values.get(field) or "")
            values[field] = value
            sources[field] = {
                "label": "人工选择候选料号",
                "reference": f"关联号 {task['correlation_id']}",
            }
            if before != value:
                changes.append({"field": field, "before": before, "after": value, "line_no": int(line_no)})
        result["selected_candidate"] = selected
        now = utcnow()
        conn.execute(
            """UPDATE order_material_resolution_tasks
               SET result_json=?,status='resolved',resolved_at=?,updated_at=? WHERE id=?""",
            (json.dumps(result, ensure_ascii=False), now, now, int(task["id"])),
        )
        conn.execute(
            """UPDATE order_entry_template_lines SET values_json=?,sources_json=?,updated_at=?
               WHERE template_id=? AND line_no=?""",
            (json.dumps(values, ensure_ascii=False), json.dumps(sources, ensure_ascii=False), now,
             template_id, int(line_no)),
        )
        conn.execute("UPDATE order_entry_templates SET updated_at=? WHERE id=?", (now, template_id))
        record_order_detail_event(
            conn,
            case_id=case_id,
            template_id=template_id,
            employee_id=employee_id,
            event_type="material_candidate_selected",
            title=f"第 {int(line_no)} 行已确认候选料号",
            detail={"line_no": int(line_no), "selected_candidate": selected, "changes": changes},
            operated_by=employee_id,
        )
    return {"line_no": int(line_no), "selected_candidate": selected, "changes": changes}


def _domestic_group_data(conn: Any, template_id: int, group_key: str = "") -> tuple[dict[str, Any], list[Any]]:
    groups = conn.execute(
        "SELECT * FROM order_entry_template_groups WHERE template_id=? ORDER BY sort_order,id",
        (template_id,),
    ).fetchall()
    if not groups:
        template = conn.execute("SELECT header_json FROM order_entry_templates WHERE id=?", (template_id,)).fetchone()
        rows = conn.execute(
            "SELECT line_no,values_json FROM order_entry_template_lines WHERE template_id=? ORDER BY line_no",
            (template_id,),
        ).fetchall()
        return {
            "id": None, "group_key": "", "status": "pending",
            "header": _json(template["header_json"] if template else "", {}),
        }, list(rows)
    selected = next((row for row in groups if str(row["group_key"]) == str(group_key or "")), None)
    if selected is None and len(groups) == 1 and not group_key:
        selected = groups[0]
    if selected is None:
        raise ValueError("请选择需要处理的 PO 分组")
    rows = conn.execute(
        """SELECT line_no,values_json FROM order_entry_template_lines
           WHERE template_id=? AND group_id=? ORDER BY line_no""",
        (template_id, int(selected["id"])),
    ).fetchall()
    group_header = _json(selected["header_json"], {})
    if len(groups) == 1 and not group_key:
        legacy_template = conn.execute(
            "SELECT header_json FROM order_entry_templates WHERE id=?", (template_id,)
        ).fetchone()
        group_header = {**group_header, **_json(legacy_template["header_json"] if legacy_template else "", {})}
    return {
        **_row(selected),
        "header": group_header,
    }, list(rows)


def validate_domestic_order_entry(case_id: int, employee_id: str, group_key: str = "") -> list[str]:
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        return ["请先生成并保存录单模板"]
    with db_cursor() as conn:
        group, lines = _domestic_group_data(conn, template_id, group_key)
        tasks = conn.execute(
            "SELECT line_no,status,result_json FROM order_material_resolution_tasks WHERE template_id=?",
            (template_id,),
        ).fetchall()
    states = {int(row["line_no"]): _row(row) for row in tasks}
    issues = []
    if str(group.get("status") or "") == "submitted":
        issues.append("当前 PO 已提交，不能重复提交")
    if not str((group.get("header") or {}).get("customer_order_number") or "").strip():
        issues.append("当前分组尚未填写客户订单号")
    if not lines:
        issues.append("当前 PO 没有订单明细")
    for row in lines:
        line_no, values = int(row["line_no"]), _json(row["values_json"], {})
        task = states.get(line_no, {})
        status = str(task.get("status") or "")
        result = _json(task.get("result_json"), {})
        candidates = _material_candidates(result)
        if status in {"waiting_callback", "requerying", "failed"}:
            issues.append(f"第 {line_no} 行料号{MATERIAL_STATUS_LABELS[status]}")
        elif not str(values.get("customer_product_code") or "").strip():
            issues.append(f"第 {line_no} 行未填写客户产品编号")
    return issues


def prepare_domestic_order_entry(case_id: int, employee_id: str, group_key: str = "") -> dict[str, Any]:
    """Build the existing NYEOS domestic-entry payload without submitting it.

    This is deliberately a service-layer operation, so Connector/API clients
    use the same template, validation and payload mapping as the web workflow.
    Unlike ``build_domestic_order_entry`` it performs no HTTP request, writes
    no interface log and does not change an order state.
    """
    config = get_interface_config("domestic_order_entry")
    if not config:
        raise ValueError("内销录单接口配置不存在")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成并保存录单模板")
    issues = validate_domestic_order_entry(case_id, employee_id, group_key)
    with db_cursor() as conn:
        group, rows = _domestic_group_data(conn, template_id, group_key)
    try:
        payload = _domestic_order_request_payload(
            group["header"], list(rows), employee_id
        )
    except ValueError as exc:
        # A preview must show all business blockers rather than pretending the
        # payload exists.  The submit service will run the same validation.
        issues.append(str(exc))
        payload = None
    completed = is_domestic_order_entry_completed(case_id, employee_id)
    return {
        "case_id": case_id,
        "template_id": template_id,
        "validation_issues": issues,
        "can_submit": bool(config.get("enabled")) and not issues and not completed,
        "already_submitted": completed,
        "interface": {
            "enabled": bool(config.get("enabled")),
            "mode": str(config.get("mode") or "disabled"),
            "config_version": int(config.get("config_version") or 0),
        },
        "payload": payload,
    }


def is_domestic_order_entry_completed(case_id: int, employee_id: str) -> bool:
    """Whether a successful domestic-entry request has already been accepted.

    The interface log is the source of truth here: a case can remain visible for
    review after entry, but it must not submit the same order a second time.
    """
    with db_cursor() as conn:
        groups = conn.execute(
            """SELECT groups.status,groups.order_number
               FROM order_entry_template_groups groups
               JOIN order_entry_templates template ON template.id=groups.template_id
               WHERE template.case_id=? AND template.employee_id=?""",
            (case_id, employee_id),
        ).fetchall()
        if groups:
            return all(
                str(row["status"] or "") == "submitted" and str(row["order_number"] or "").strip()
                for row in groups
            )
        row = conn.execute(
            """SELECT 1 FROM order_interface_call_logs
               WHERE case_id=? AND employee_id=? AND interface_key='domestic_order_entry'
                 AND status='success' LIMIT 1""",
            (case_id, employee_id),
        ).fetchone()
        return bool(row)


def _complete_case_after_group_submissions(case_id: int, employee_id: str) -> None:
    if not is_domestic_order_entry_completed(case_id, employee_id):
        return
    now = utcnow()
    with db_cursor() as conn:
        conn.execute(
            """UPDATE order_intake_cases
               SET status='archived',workflow_stage='completed',erp_prepare_status='submitted',
                   completed_at=?,updated_at=? WHERE id=? AND employee_id=?""",
            (now, now, case_id, employee_id),
        )


def build_domestic_order_entry_mock(
    case_id: int, employee_id: str, triggered_by: str, group_key: str = "",
) -> dict[str, Any]:
    """Record a manual domestic-order submission Mock without changing the template."""
    issues = validate_domestic_order_entry(case_id, employee_id, group_key)
    if issues:
        raise ValueError("暂不能提交录单：" + "；".join(issues))
    config = get_interface_config("domestic_order_entry")
    if not config:
        raise ValueError("内销录单接口配置不存在")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成并保存录单模板")
    with db_cursor() as conn:
        group, lines = _domestic_group_data(conn, template_id, group_key)
        request_payload = {
            "header": group["header"],
            "items": [{"line_no": int(row["line_no"]), **_json(row["values_json"], {})} for row in lines],
        }
        response_payload = {
            "status": "accepted", "entry_no": f"MOCK-SO-{case_id}",
            "message": "Mock 录单成功；真实接口接入后将返回实际单号。",
        }
        now = utcnow()
        cursor = conn.execute(
            """INSERT INTO order_interface_call_logs
               (case_id,template_id,order_group_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
                status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                case_id, template_id, group.get("id"), employee_id, int(config["id"]), "domestic_order_entry",
                int(config["config_version"]), 1, "success", 200, 1,
                json.dumps(request_payload, ensure_ascii=False), json.dumps(response_payload, ensure_ascii=False),
                "", triggered_by, now,
            ),
        )
        call_id = int(cursor.lastrowid)
        if group.get("id"):
            conn.execute(
                """UPDATE order_entry_template_groups
                   SET status='submitted',nyeos_order_number=?,submitted_at=?,updated_at=? WHERE id=?""",
                (response_payload["entry_no"], now, now, int(group["id"])),
            )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="domestic_order_entry_mock", title="提交内销录单（Mock）完成",
            detail={"call_id": call_id, "group_key": group.get("group_key"), "line_count": len(lines), "entry_no": response_payload["entry_no"]},
            operated_by=triggered_by,
        )
    _complete_case_after_group_submissions(case_id, employee_id)
    return {"call_id": call_id, "entry_no": response_payload["entry_no"], "status": "success", "mode": "mock", "group_key": group.get("group_key")}


def review_domestic_order_entry_prices(
    case_id: int, employee_id: str, group_key: str = "",
) -> dict[str, Any]:
    """Check current template values against the saved price-review snapshot."""
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        return {"association": {}, "tax_mode": "unknown", "by_line": {}, "mismatches": []}
    with db_cursor() as conn:
        group, rows = _domestic_group_data(conn, template_id, group_key)
    lines = [
        {"line_no": int(row["line_no"]), "values": _json(row["values_json"], {})}
        for row in rows
    ]
    header = group["header"]
    return review_cached_template_prices(header.get(PRICE_REVIEW_SNAPSHOT_KEY), lines)


def build_domestic_order_entry(
    case_id: int, employee_id: str, triggered_by: str, *, group_key: str = "",
    allow_price_mismatch: bool = False,
) -> dict[str, Any]:
    price_review = review_domestic_order_entry_prices(case_id, employee_id, group_key)
    if price_review["mismatches"] and not allow_price_mismatch:
        raise PriceMismatchConfirmationRequired(price_review)
    config = get_interface_config("domestic_order_entry")
    if not config or not config.get("enabled"):
        raise ValueError("生成订单接口未启用")
    if str(config.get("mode") or "mock") == "real":
        return build_domestic_order_entry_real(case_id, employee_id, triggered_by, group_key=group_key, config=config)
    return build_domestic_order_entry_mock(case_id, employee_id, triggered_by, group_key)


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
        factory_part_code = str(values.get("product_code") or "").strip()
        material_code = str(values.get("customer_product_code") or "").strip()
        demand_date = str(values.get("delivery_date") or "").strip()
        tax_price = str(values.get("unit_price") or "").strip()
        untaxed_price = str(values.get("price_before_tax") or "").strip()
        if not quantity:
            line_issues.append(f"第 {line_no} 行未填写数量")
        if not material_code:
            line_issues.append(f"第 {line_no} 行未填写客户产品编号")
        if not factory_part_code:
            line_issues.append(f"第 {line_no} 行未填写产品编号")
        if not demand_date:
            line_issues.append(f"第 {line_no} 行未填写出货日期")
        if not tax_price and not untaxed_price:
            line_issues.append(f"第 {line_no} 行单价和税前单价至少填写一项")
        sequence = str(values.get("customer_order_seq") or line_no).strip()
        items.append({
            "customerCode": customer_code,
            "orderType": order_type,
            "operator": employee_id,
            "quantity": quantity,
            "taxPrice": tax_price,
            "untaxedPrice": untaxed_price,
            "factoryPartCode": factory_part_code,
            "materialCode": material_code,
            "lineNumber": sequence,
            "demandDate": demand_date,
            "orderNumber": order_number,
            "custOrderId": order_number,
            "lineId": sequence,
            "lineRemark": str(values.get("remark") or "").strip(),
            "taxType": str(header.get("tax_type") or "").strip(),
            "materialName": str(values.get("product_name") or "").strip(),
            "spec": str(values.get("customer_spec") or "").strip(),
        })
    if line_issues:
        raise ValueError("暂不能提交录单：" + "；".join(line_issues))
    return {"sctoDataList": items}


def _insert_domestic_call_log(
    conn: Any, *, case_id: int, template_id: int, employee_id: str, config: dict[str, Any],
    order_group_id: int | None,
    status: str, request_payload: dict[str, Any], response_payload: dict[str, Any],
    triggered_by: str, http_status: int | None, duration_ms: int | None,
    error_message: str = "",
) -> int:
    cursor = conn.execute(
        """INSERT INTO order_interface_call_logs
           (case_id,template_id,order_group_id,employee_id,interface_config_id,interface_key,config_version,is_mock,
            status,http_status,duration_ms,request_json,response_json,error_message,triggered_by,created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            case_id, template_id, order_group_id, employee_id, int(config["id"]), "domestic_order_entry",
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


def _is_uncertain_domestic_entry_error(error: Exception) -> bool:
    """Return whether the remote request might have been accepted before its response was lost."""
    text = str(error).casefold()
    return any(marker in text for marker in (
        "timed out", "timeout", "connection", "连接", "network", "reset by peer",
        "remote end closed", "远程主机强迫关闭",
    ))


def _matching_submitted_order(
    response_payload: dict[str, Any], customer_order_number: str, rows: list[Any],
) -> tuple[dict[str, Any] | None, str]:
    expected: list[tuple[str, Decimal]] = []
    for row in rows:
        values = _json(row["values_json"], {})
        sequence = _match_text(values.get("customer_order_seq") or row["line_no"])
        quantity = _match_decimal(values.get("quantity"))
        if not sequence or quantity is None:
            return None, "当前模板的项次或数量无效，无法自动核对"
        expected.append((sequence, quantity))

    order_number = _match_text(customer_order_number)
    data = response_payload.get("data") if isinstance(response_payload.get("data"), dict) else {}
    matches: list[dict[str, Any]] = []
    for order in data.get("orderList") or []:
        if not isinstance(order, dict):
            continue
        details = [item for item in order.get("sctbList") or [] if isinstance(item, dict)]
        belongs_to_po = _match_text(order.get("scta38")) == order_number or any(
            _match_text(item.get("sctb15")) == order_number for item in details
        )
        if not belongs_to_po:
            continue
        actual: list[tuple[str, Decimal]] = []
        for detail in details:
            item_no = _match_text(detail.get("sctb35"))
            quantity = _match_decimal(detail.get("sctb05"))
            if not item_no or quantity is None:
                actual = []
                break
            actual.append((item_no, quantity))
        if sorted(actual) == sorted(expected) and str(order.get("scta01") or "").strip() and str(order.get("scta39") or "").strip():
            matches.append(order)

    if len(matches) == 1:
        return matches[0], ""
    if len(matches) > 1:
        return None, "查到多笔项次和数量均一致的订单，无法自动确认"
    return None, "查单结果与当前模板的项次或数量不一致"


def _recover_domestic_entry_after_uncertain_failure(
    case_id: int, template_id: int, employee_id: str, triggered_by: str,
    group: dict[str, Any], rows: list[Any], failed_call_id: int,
) -> tuple[dict[str, Any] | None, str]:
    order_number = str((group.get("header") or {}).get("customer_order_number") or "").strip()
    try:
        query = _query_order_info(
            case_id, template_id, employee_id, triggered_by, [order_number],
            persist_matches=False, event_context="domestic_recovery",
        )
    except ValueError as exc:
        reason = f"自动查单失败：{exc}"
    else:
        matched, reason = _matching_submitted_order(query["response"], order_number, rows)
        if matched:
            nyeos_order_number = str(matched.get("scta01") or "").strip()
            erp_order_number = str(matched.get("scta39") or "").strip()
            now = utcnow()
            with db_cursor() as conn:
                conn.execute(
                    """UPDATE order_entry_template_groups
                       SET status='submitted',nyeos_order_number=?,erp_order_number=?,submitted_at=?,updated_at=?
                       WHERE id=?""",
                    (nyeos_order_number, erp_order_number, now, now, int(group["id"])),
                )
                record_order_detail_event(
                    conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                    event_type="domestic_order_entry_recovered",
                    title="提交录单超时后已查单确认成功",
                    detail={
                        "failed_call_id": failed_call_id, "query_call_id": query["call_id"],
                        "group_key": group.get("group_key"), "entry_no": nyeos_order_number,
                        "erp_order_number": erp_order_number,
                    },
                    operated_by=triggered_by,
                )
            _complete_case_after_group_submissions(case_id, employee_id)
            return {
                "call_id": failed_call_id,
                "entry_no": nyeos_order_number,
                "status": "success",
                "mode": "real",
                "message": "提交接口超时，已查单确认订单生成成功。",
                "group_key": group.get("group_key"),
                "recovered_from_timeout": True,
            }, ""

    with db_cursor() as conn:
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="domestic_order_entry_recovery_unconfirmed",
            title="提交录单超时后未能查单确认",
            detail={"failed_call_id": failed_call_id, "group_key": group.get("group_key"), "reason": reason},
            operated_by=triggered_by,
        )
    return None, reason


def build_domestic_order_entry_real(
    case_id: int, employee_id: str, triggered_by: str, *, group_key: str = "",
    config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    issues = validate_domestic_order_entry(case_id, employee_id, group_key)
    if issues:
        raise ValueError("暂不能提交录单：" + "；".join(issues))
    config = config or get_interface_config("domestic_order_entry")
    if not config or not config.get("enabled"):
        raise ValueError("生成订单接口未启用")
    template_id = _case_template_id(case_id, employee_id)
    if not template_id:
        raise ValueError("请先生成并保存录单模板")
    with db_cursor() as conn:
        group, rows = _domestic_group_data(conn, template_id, group_key)
    header = group["header"]
    request_payload = _domestic_order_request_payload(header, list(rows), employee_id)
    try:
        http_status, response_body, duration_ms = _post_json_endpoint(config, request_payload, "真实生成订单")
    except ValueError as exc:
        with db_cursor() as conn:
            call_id = _insert_domestic_call_log(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
                order_group_id=int(group["id"]) if group.get("id") else None,
                status="failed", request_payload=request_payload, response_payload={}, triggered_by=triggered_by,
                http_status=None, duration_ms=None, error_message=str(exc),
            )
            record_order_detail_event(
                conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
                event_type="domestic_order_entry_real", title="提交录单（真实接口）失败",
                detail={"call_id": call_id, "error_message": str(exc)}, operated_by=triggered_by,
            )
        if _is_uncertain_domestic_entry_error(exc):
            recovered, reason = _recover_domestic_entry_after_uncertain_failure(
                case_id, template_id, employee_id, triggered_by, group, list(rows), call_id,
            )
            if recovered:
                return recovered
            raise ValueError(f"{exc}；自动查单未能确认订单生成：{reason}") from exc
        raise
    ok, entry_no, message = _domestic_response_result(response_body, http_status)
    now = utcnow()
    with db_cursor() as conn:
        call_id = _insert_domestic_call_log(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id, config=config,
            order_group_id=int(group["id"]) if group.get("id") else None,
            status="success" if ok else "failed", request_payload=request_payload,
            response_payload=response_body, triggered_by=triggered_by, http_status=http_status,
            duration_ms=duration_ms, error_message="" if ok else message,
        )
        if ok:
            response_data = response_body.get("data") if isinstance(response_body.get("data"), dict) else {}
            records = [item for item in response_data.get("data") or [] if isinstance(item, dict)]
            erp_map = response_data.get("erpOrderMap") if isinstance(response_data.get("erpOrderMap"), dict) else {}
            erp_numbers = []
            for item in records:
                erp_number = str(item.get("scta39") or erp_map.get(str(item.get("sctaCode") or "")) or "").strip()
                if erp_number and erp_number not in erp_numbers:
                    erp_numbers.append(erp_number)
            if group.get("id"):
                conn.execute(
                    """UPDATE order_entry_template_groups
                       SET status='submitted',nyeos_order_number=?,erp_order_number=?,submitted_at=?,updated_at=?
                       WHERE id=?""",
                    (entry_no, "、".join(erp_numbers), now, now, int(group["id"])),
                )
        record_order_detail_event(
            conn, case_id=case_id, template_id=template_id, employee_id=employee_id,
            event_type="domestic_order_entry_real",
            title="提交录单（真实接口）完成" if ok else "提交录单（真实接口）失败",
            detail={"call_id": call_id, "group_key": group.get("group_key"), "line_count": len(rows), "entry_no": entry_no, "error_message": "" if ok else message},
            operated_by=triggered_by,
        )
    if not ok:
        raise ValueError(message)
    _complete_case_after_group_submissions(case_id, employee_id)
    return {"call_id": call_id, "entry_no": entry_no, "status": "success", "mode": "real", "message": message, "group_key": group.get("group_key")}
