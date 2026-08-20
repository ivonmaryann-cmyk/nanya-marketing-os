from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from .db import db_cursor, utcnow


ACTION_LABELS = {
    "unclassified": "暂不分流",
    "new_order": "录单",
    "order_change": "修改订单",
    "quotation": "报价",
}
ROUTABLE_ACTION_TYPES = {"new_order", "order_change", "quotation"}
STATUS_LABELS = {
    "pending_triage": "待处理",
    "pending_review": "处理中",
    "ready_for_erp": "待确认",
    "on_hold": "待补充",
    "archived": "已完成",
}
WORKFLOW_STAGE_LABELS = {
    "mail_triage": "1. 邮件分流",
    "order_identification": "2. 订单识别",
    "data_mapping": "3. 主数据与151映射",
    "erp_preparation": "4. 410录单准备",
    "completed": "5. 已提交，待回查",
}
MATCH_STATUS_LABELS = {
    "unmatched": "待匹配",
    "matched": "已匹配负责客户",
    "needs_assignment": "需调整归属",
}
DOCUMENT_STATUS_LABELS = {
    "pending": "待核对",
    "complete": "订单资料齐全",
    "missing": "缺少订单资料",
}
MAPPING_STATUS_LABELS = {
    "not_started": "尚未开始",
    "ready": "映射资料待核",
    "exception": "存在映射异常",
    "confirmed": "映射已确认",
}
ERP_PREPARE_STATUS_LABELS = {
    "not_started": "未准备",
    "waiting_interface": "等待接口/模板",
    "ready_for_entry": "可录入410",
    "submitted": "已提交待回查",
}

BUSINESS_TIMEZONE = ZoneInfo("Asia/Shanghai")


def business_today() -> date:
    """Calendar day used by the Shanghai-based operations team."""
    return datetime.now(BUSINESS_TIMEZONE).date()


def _business_date(timestamp: str) -> str:
    """Convert legacy UTC timestamps (stored without an offset) to business date."""
    if not timestamp:
        return ""
    try:
        value = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.astimezone(BUSINESS_TIMEZONE).date().isoformat()
    except ValueError:
        return timestamp[:10]


def _as_dict(row) -> dict[str, Any] | None:
    return dict(row) if row else None


def _rule_matches(rule: dict[str, Any], *, subject: str, sender: str, attachment_text: str = "") -> bool:
    """A business rule is an AND of the conditions it actually fills in."""
    conditions = (
        (str(rule.get("sender_contains") or ""), sender),
        (str(rule.get("subject_contains") or ""), subject),
        (str(rule.get("attachment_contains") or ""), attachment_text),
    )
    filled = [(needle.strip().lower(), haystack.lower()) for needle, haystack in conditions if needle.strip()]
    return bool(filled) and all(needle in haystack for needle, haystack in filled)


def _match_rule(rules: list[dict[str, Any]], *, subject: str, sender: str, attachment_text: str = "") -> dict[str, Any] | None:
    return next((rule for rule in rules if _rule_matches(rule, subject=subject, sender=sender, attachment_text=attachment_text)), None)


def classify_mail(subject: str, sender: str = "", attachment_text: str = "", rules: list[dict[str, Any]] | None = None) -> tuple[str, str, int | None, dict[str, Any] | None]:
    """Generate an explainable route that business users can override."""
    matched_rule = _match_rule(rules or [], subject=subject, sender=sender, attachment_text=attachment_text)
    if matched_rule:
        return (
            str(matched_rule["action_type"]),
            f"命中业务规则：{matched_rule['name']}",
            int(matched_rule["id"]),
            matched_rule,
        )
    text = f"{subject} {sender}".lower()
    rules = (
        ("order_change", r"修改|变更|改单|改期|取消|暂停|撤销|作废|revision|revise|amend|change|cancel", "主题包含订单变更信号"),
        ("quotation", r"rfq|rfx|询价|报价|竞价|议价|投标|中标|bidding", "主题包含询报价信号"),
        ("delivery", r"交期|交货|发货|送货|出货|催交|排产|shipment|delivery", "主题包含交期或发货信号"),
        ("new_order", r"采购订单|新增订单|订单发布|订单需求|新订单|purchase\s*order|\bpo\b|下单", "主题包含新订单信号"),
    )
    for action_type, pattern, reason in rules:
        if re.search(pattern, text, flags=re.IGNORECASE):
            return action_type, reason, None, None
    return "other", "暂未匹配到明确业务类型", None, None


def list_routing_rules(employee_id: str, *, include_disabled: bool = True) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        sql = "SELECT * FROM order_mail_routing_rules WHERE employee_id = ?"
        if not include_disabled:
            sql += " AND enabled = 1"
        sql += " ORDER BY enabled DESC, id ASC"
        rows = conn.execute(sql, (employee_id,)).fetchall()
    return [dict(row) for row in rows]


def get_routing_rule(rule_id: int, employee_id: str) -> dict[str, Any] | None:
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT * FROM order_mail_routing_rules WHERE id = ? AND employee_id = ?",
            (rule_id, employee_id),
        ).fetchone()
    return _as_dict(row)


def _rule_payload(payload: dict[str, Any]) -> dict[str, Any]:
    data = {
        "name": str(payload.get("name") or "").strip(),
        "enabled": 1 if str(payload.get("enabled") or "") in {"1", "on", "true"} else 0,
        "priority": 100,
        "sender_contains": str(payload.get("sender_contains") or "").strip(),
        "subject_contains": str(payload.get("subject_contains") or "").strip(),
        "attachment_contains": str(payload.get("attachment_contains") or "").strip(),
        "action_type": str(payload.get("action_type") or "").strip(),
        "customer_code": str(payload.get("customer_code") or "").strip(),
        "customer_name": str(payload.get("customer_name") or "").strip(),
        "note": str(payload.get("note") or "").strip(),
    }
    if not data["name"]:
        raise ValueError("请填写规则名称")
    if data["action_type"] not in ACTION_LABELS:
        raise ValueError("请选择规则要分流到的业务类型")
    if not (data["sender_contains"] or data["subject_contains"] or data["attachment_contains"]):
        raise ValueError("请至少填写一个匹配条件")
    return data


def save_routing_rule(employee_id: str, payload: dict[str, Any], *, rule_id: int | None = None) -> dict[str, Any]:
    data = _rule_payload(payload)
    now = utcnow()
    with db_cursor() as conn:
        before: dict[str, Any] = {}
        if rule_id:
            row = conn.execute("SELECT * FROM order_mail_routing_rules WHERE id = ? AND employee_id = ?", (rule_id, employee_id)).fetchone()
            if not row:
                raise ValueError("分流规则不存在或无权修改")
            before = dict(row)
            conn.execute(
                """UPDATE order_mail_routing_rules SET name=?, enabled=?, priority=?, sender_contains=?, subject_contains=?,
                   attachment_contains=?, action_type=?, customer_code=?, customer_name=?, note=?, updated_at=? WHERE id=?""",
                (*data.values(), now, rule_id),
            )
            saved_id = rule_id
            event = "updated"
        else:
            cursor = conn.execute(
                """INSERT INTO order_mail_routing_rules (employee_id, name, enabled, priority, sender_contains, subject_contains,
                   attachment_contains, action_type, customer_code, customer_name, note, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (employee_id, *data.values(), now, now),
            )
            saved_id = int(cursor.lastrowid)
            event = "created"
        conn.execute(
            """INSERT INTO order_mail_routing_rule_events (rule_id, employee_id, action, before_json, after_json, created_at)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (saved_id, employee_id, event, json.dumps(before, ensure_ascii=False), json.dumps(data, ensure_ascii=False), now),
        )
    return get_routing_rule(saved_id, employee_id) or {}


EXTRACTED_RULE_PATTERNS = (
    {
        "key": "bomin-po", "name": "博敏采购订单", "sender_contains": "bominelec.com",
        "subject_contains": "博敏电子采购订单", "attachment_contains": "", "action_type": "new_order",
        "customer_name": "博敏电子", "note": "从历史邮件中提炼：博敏采购订单主题。",
    },
    {
        "key": "bomin-new-order", "name": "博敏新增订单需求", "sender_contains": "bominelec.com",
        "subject_contains": "新增订单需求", "attachment_contains": "", "action_type": "new_order",
        "customer_name": "博敏电子", "note": "从历史邮件中提炼：博敏新增订单需求通知。",
    },
    {
        "key": "chinafast-order-release", "name": "方正订单发布", "sender_contains": "chinafastprint.com",
        "subject_contains": "订单发布消息通知", "attachment_contains": "", "action_type": "new_order",
        "customer_name": "方正", "note": "从历史邮件中提炼：SRM 订单发布通知。",
    },
    {
        "key": "chinafast-rfq", "name": "方正询报价通知", "sender_contains": "chinafastprint.com",
        "subject_contains": "RFQ", "attachment_contains": "", "action_type": "quotation",
        "customer_name": "方正", "note": "从历史邮件中提炼：SRM RFQ 通知。",
    },
    {
        "key": "chinafast-award", "name": "方正中标通知", "sender_contains": "chinafastprint.com",
        "subject_contains": "中标通知", "attachment_contains": "", "action_type": "quotation",
        "customer_name": "方正", "note": "从历史邮件中提炼：SRM 中标通知。",
    },
    {
        "key": "chinafast-rfx-adjust", "name": "方正 RFX 时间调整", "sender_contains": "chinafastprint.com",
        "subject_contains": "RFX时间调整通知", "attachment_contains": "", "action_type": "unclassified",
        "customer_name": "方正", "note": "从历史邮件中提炼：此类通知暂不进入订单处理。",
    },
    {
        "key": "bomin-price-change", "name": "博敏单价修改申请", "sender_contains": "bominelec.com",
        "subject_contains": "单价修改申请", "attachment_contains": "", "action_type": "order_change",
        "customer_name": "博敏电子", "note": "从历史邮件中提炼：单价修改申请。",
    },
)


def list_extracted_routing_suggestions(employee_id: str, account_id: int | None) -> list[dict[str, Any]]:
    if not account_id:
        return []
    active_conditions = {
        (rule["sender_contains"].lower(), rule["subject_contains"].lower(), rule["attachment_contains"].lower())
        for rule in list_routing_rules(employee_id)
    }
    suggestions: list[dict[str, Any]] = []
    with db_cursor() as conn:
        for pattern in EXTRACTED_RULE_PATTERNS:
            conditions = (pattern["sender_contains"].lower(), pattern["subject_contains"].lower(), pattern["attachment_contains"].lower())
            if conditions in active_conditions:
                continue
            where = ["m.account_id = ?"]
            params: list[Any] = [account_id]
            for column, key in (("m.sender", "sender_contains"), ("m.subject", "subject_contains")):
                if pattern[key]:
                    where.append(f"LOWER({column}) LIKE ?")
                    params.append(f"%{pattern[key].lower()}%")
            row = conn.execute(
                f"SELECT COUNT(*) AS total, MAX(COALESCE(m.sent_at, m.received_at, m.created_at)) AS latest_at "
                f"FROM mail_messages m WHERE {' AND '.join(where)}",
                params,
            ).fetchone()
            if row and int(row["total"]):
                item = dict(pattern)
                item["match_count"] = int(row["total"])
                item["latest_at"] = str(row["latest_at"] or "")
                suggestions.append(item)
    return sorted(suggestions, key=lambda item: item["match_count"], reverse=True)


def save_extracted_routing_suggestion(employee_id: str, suggestion_key: str, *, decision: str) -> dict[str, Any]:
    pattern = next((item for item in EXTRACTED_RULE_PATTERNS if item["key"] == suggestion_key), None)
    if not pattern:
        raise ValueError("分流建议不存在")
    if decision not in {"route", "ignore"}:
        raise ValueError("请选择明确分流或暂不分流")
    payload = dict(pattern)
    payload["enabled"] = "1"
    payload["action_type"] = pattern["action_type"] if decision == "route" else "unclassified"
    payload["name"] = pattern["name"] if decision == "route" else f"{pattern['name']}（暂不分流）"
    payload["note"] = f"{pattern['note']} 业务已确认：{'明确分流' if decision == 'route' else '暂不分流'}。"
    return save_routing_rule(employee_id, payload)


def toggle_routing_rule(rule_id: int, employee_id: str) -> dict[str, Any]:
    current = get_routing_rule(rule_id, employee_id)
    if not current:
        raise ValueError("分流规则不存在或无权操作")
    now = utcnow()
    enabled = 0 if current["enabled"] else 1
    with db_cursor() as conn:
        conn.execute("UPDATE order_mail_routing_rules SET enabled=?, updated_at=? WHERE id=?", (enabled, now, rule_id))
        conn.execute(
            "INSERT INTO order_mail_routing_rule_events (rule_id, employee_id, action, before_json, after_json, created_at) VALUES (?, ?, 'toggled', ?, ?, ?)",
            (rule_id, employee_id, json.dumps({"enabled": current["enabled"]}), json.dumps({"enabled": enabled}), now),
        )
    return get_routing_rule(rule_id, employee_id) or {}


def simulate_routing_rule(rule: dict[str, Any], employee_id: str, account_id: int, *, days: int = 30) -> dict[str, Any]:
    start_date = (business_today() - timedelta(days=max(1, days) - 1)).isoformat()
    with db_cursor() as conn:
        rows = conn.execute(
            """SELECT m.id, m.subject, m.sender, m.sent_at, m.received_at,
               COALESCE(GROUP_CONCAT(a.filename, ' '), '') AS attachment_text
               FROM mail_messages m LEFT JOIN mail_attachments a ON a.mail_id=m.id AND a.is_inline=0
               JOIN mail_accounts ma ON ma.id=m.account_id
               WHERE ma.owner_employee_id=? AND m.account_id=?
                 AND substr(COALESCE(NULLIF(m.sent_at, ''), NULLIF(m.received_at, ''), m.created_at), 1, 10) >= ?
               GROUP BY m.id ORDER BY COALESCE(m.sent_at, m.received_at, m.created_at) DESC""",
            (employee_id, account_id, start_date),
        ).fetchall()
    matches = [dict(row) for row in rows if _rule_matches(rule, subject=row["subject"], sender=row["sender"], attachment_text=row["attachment_text"])]
    return {"count": len(matches), "samples": matches[:5]}


SCOPE_LABELS = {"subject": "邮件主题", "body": "邮件正文", "attachment_name": "附件名称", "attachment_content": "附件内容"}
ROUTING_STATE_LABELS = {"routed": "已明确分流", "needs_business_routing": "待业务分流", "unrouted": "暂不分流", "business_routed": "业务已分流"}
CLEAR_ROUTING_SCOPES = {"subject", "attachment_name"}

DEFAULT_GROUPS = (
    ("录单", "new_order", (
        ("subject", "采购订单"), ("subject", "新增订单需求"), ("subject", "订单发布"), ("subject", "订单需求"),
        ("body", "采购订单"), ("body", "订单号"), ("body", "下单"),
        ("attachment_name", "PO"), ("attachment_name", "采购订单"), ("attachment_name", "订单"),
        ("attachment_content", "采购订单"), ("attachment_content", "订单号"),
    )),
    ("报价", "quotation", (
        ("subject", "RFQ"), ("subject", "询价"), ("subject", "报价"), ("subject", "中标"),
        ("body", "RFQ"), ("body", "询价"), ("body", "报价"), ("body", "请报价"), ("body", "议价"),
        ("attachment_name", "RFQ"), ("attachment_name", "报价单"),
        ("attachment_content", "RFQ"), ("attachment_content", "报价"),
    )),
    ("修改订单", "order_change", (
        ("subject", "变更"), ("subject", "改期"), ("subject", "取消订单"), ("subject", "单价修改"),
        ("subject", "交期提前"), ("subject", "交期协同提前"), ("subject", "交期加急"), ("subject", "提前交货"), ("subject", "提前交期"),
        ("body", "变更"), ("body", "改期"), ("body", "价格调整"), ("body", "单价修改"), ("body", "版本更新"), ("body", "取消"),
        ("body", "交期提前"), ("body", "交期协同提前"), ("body", "交期加急"), ("body", "提前交货"), ("body", "提前交期"),
        ("attachment_name", "变更"), ("attachment_name", "改单"), ("attachment_name", "改期"),
        ("attachment_name", "交期提前"), ("attachment_name", "交期协同提前"), ("attachment_name", "交期加急"), ("attachment_name", "提前交货"), ("attachment_name", "提前交期"),
        ("attachment_content", "变更"), ("attachment_content", "价格调整"), ("attachment_content", "交期修改"),
        ("attachment_content", "交期提前"), ("attachment_content", "交期协同提前"), ("attachment_content", "交期加急"), ("attachment_content", "提前交货"), ("attachment_content", "提前交期"),
    )),
)
DEFAULT_CHANGE_TAGS = (
    ("交期 / 发货日期", (
        ("subject", "交期修改"), ("subject", "交期变更"), ("subject", "改期"),
        ("subject", "交期提前"), ("subject", "交期协同提前"), ("subject", "交期加急"), ("subject", "提前交货"), ("subject", "提前交期"),
        ("body", "交期修改"), ("body", "交期变更"), ("body", "改期"),
        ("body", "发货日期修改"), ("body", "发货日期变更"),
        ("body", "交期提前"), ("body", "交期协同提前"), ("body", "交期加急"), ("body", "提前交货"), ("body", "提前交期"),
        ("attachment_name", "交期提前"), ("attachment_name", "交期协同提前"), ("attachment_name", "交期加急"), ("attachment_name", "提前交货"), ("attachment_name", "提前交期"),
        ("attachment_content", "交期提前"), ("attachment_content", "交期协同提前"), ("attachment_content", "交期加急"), ("attachment_content", "提前交货"), ("attachment_content", "提前交期"),
    )),
    ("价格调整", (("subject", "单价修改"), ("subject", "价格调整"), ("body", "价格调整"), ("body", "单价修改"), ("body", "单价调整"), ("body", "调价"))),
    ("数量调整", (("body", "数量调整"), ("body", "数量变更"))),
    ("规格 / 型号调整", (("subject", "规格调整"), ("subject", "规格变更"), ("subject", "型号调整"), ("subject", "型号变更"), ("body", "规格调整"), ("body", "规格变更"), ("body", "型号调整"), ("body", "型号变更"))),
    ("订单版本 / PO", (("body", "版本更新"), ("body", "修订"), ("body", "revision"))),
    ("取消 / 暂停", (("subject", "取消"), ("body", "暂停"), ("body", "撤销"))),
)


def ensure_universal_rules(employee_id: str) -> None:
    now = utcnow()
    with db_cursor() as conn:
        # “暂不分流”是未命中或冲突时的系统状态，不能作为业务规则维护。
        obsolete_ids = [row["id"] for row in conn.execute(
            "SELECT id FROM order_mail_rule_groups WHERE employee_id=? AND action_type='unclassified'", (employee_id,)
        ).fetchall()]
        for group_id in obsolete_ids:
            conn.execute("DELETE FROM order_mail_rule_keywords WHERE group_id=?", (group_id,))
            conn.execute("DELETE FROM order_mail_rule_groups WHERE id=?", (group_id,))
        seed_key = f"order_mail_rule_seed_v2:{employee_id}"
        seeded = conn.execute("SELECT 1 FROM settings WHERE key=?", (seed_key,)).fetchone()
        for name, action_type, keywords in DEFAULT_GROUPS:
            group = conn.execute("SELECT id FROM order_mail_rule_groups WHERE employee_id=? AND action_type=? ORDER BY id LIMIT 1", (employee_id, action_type)).fetchone()
            if group is None:
                cursor = conn.execute("INSERT INTO order_mail_rule_groups (employee_id,name,action_type,enabled,created_at,updated_at) VALUES (?,?,?,?,?,?)", (employee_id, name, action_type, 1, now, now))
                group_id = int(cursor.lastrowid)
            else:
                group_id = int(group["id"])
            # 仅在首次升级到四个检索区域时补齐示例词；之后业务删除的词不会再被加回。
            if not seeded:
                existing_pairs = {(row["scope"], row["keyword"]) for row in conn.execute("SELECT scope,keyword FROM order_mail_rule_keywords WHERE group_id=?", (group_id,)).fetchall()}
                for scope, keyword in keywords:
                    if (scope, keyword) not in existing_pairs:
                        conn.execute("INSERT INTO order_mail_rule_keywords (group_id,scope,keyword,created_at) VALUES (?,?,?,?)", (group_id, scope, keyword, now))
        if not seeded:
            conn.execute("INSERT INTO settings(key,value) VALUES (?,?)", (seed_key, now))

        # “交期”本身是新订单的正常字段，只有明确表达修改/变更时才可识别为订单修改。
        cleanup_key = f"order_mail_rule_cleanup_v3:{employee_id}"
        cleaned = conn.execute("SELECT 1 FROM settings WHERE key=?", (cleanup_key,)).fetchone()
        if not cleaned:
            change_group = conn.execute(
                "SELECT id FROM order_mail_rule_groups WHERE employee_id=? AND action_type='order_change' ORDER BY id LIMIT 1",
                (employee_id,),
            ).fetchone()
            if change_group:
                conn.execute("DELETE FROM order_mail_rule_keywords WHERE group_id=? AND keyword='交期'", (change_group["id"],))
                existing_pairs = {(row["scope"], row["keyword"]) for row in conn.execute("SELECT scope,keyword FROM order_mail_rule_keywords WHERE group_id=?", (change_group["id"],)).fetchall()}
                for scope, keyword in (("attachment_content", "交期修改"), ("attachment_content", "交期变更"), ("attachment_content", "发货日期修改")):
                    if (scope, keyword) not in existing_pairs:
                        conn.execute("INSERT INTO order_mail_rule_keywords (group_id,scope,keyword,created_at) VALUES (?,?,?,?)", (change_group["id"], scope, keyword, now))
            delivery_tag = conn.execute("SELECT id FROM order_change_tags WHERE employee_id=? AND name='交期 / 发货日期'", (employee_id,)).fetchone()
            if delivery_tag:
                conn.execute("DELETE FROM order_change_tag_keywords WHERE tag_id=? AND keyword='交期'", (delivery_tag["id"],))
                existing_pairs = {(row["scope"], row["keyword"]) for row in conn.execute("SELECT scope,keyword FROM order_change_tag_keywords WHERE tag_id=?", (delivery_tag["id"],)).fetchall()}
                for scope, keyword in DEFAULT_CHANGE_TAGS[0][1]:
                    if (scope, keyword) not in existing_pairs:
                        conn.execute("INSERT INTO order_change_tag_keywords (tag_id,scope,keyword,created_at) VALUES (?,?,?,?)", (delivery_tag["id"], scope, keyword, now))
            conn.execute("INSERT INTO settings(key,value) VALUES (?,?)", (cleanup_key, now))
        tags_exist = conn.execute("SELECT 1 FROM order_change_tags WHERE employee_id=?", (employee_id,)).fetchone()
        if not tags_exist:
            for name, keywords in DEFAULT_CHANGE_TAGS:
                cursor = conn.execute("INSERT INTO order_change_tags (employee_id,name,enabled,created_at,updated_at) VALUES (?,?,?,?,?)", (employee_id, name, 1, now, now))
                for scope, keyword in keywords:
                    conn.execute("INSERT INTO order_change_tag_keywords (tag_id,scope,keyword,created_at) VALUES (?,?,?,?)", (cursor.lastrowid, scope, keyword, now))

        # 收紧过于宽泛的二级事项词，避免普通订单中的规格、型号、单价字段被误判为变更。
        refinement_key = f"order_change_item_refinement_v4:{employee_id}"
        refined = conn.execute("SELECT 1 FROM settings WHERE key=?", (refinement_key,)).fetchone()
        if not refined:
            replacements = {
                "价格调整": {"remove": {"单价"}, "add": DEFAULT_CHANGE_TAGS[1][1]},
                "规格 / 型号调整": {"remove": {"规格", "型号"}, "add": DEFAULT_CHANGE_TAGS[3][1]},
            }
            for tag_name, change in replacements.items():
                tag = conn.execute("SELECT id FROM order_change_tags WHERE employee_id=? AND name=?", (employee_id, tag_name)).fetchone()
                if not tag:
                    continue
                for word in change["remove"]:
                    conn.execute("DELETE FROM order_change_tag_keywords WHERE tag_id=? AND keyword=?", (tag["id"], word))
                existing_pairs = {(row["scope"], row["keyword"]) for row in conn.execute("SELECT scope,keyword FROM order_change_tag_keywords WHERE tag_id=?", (tag["id"],)).fetchall()}
                for scope, keyword in change["add"]:
                    if (scope, keyword) not in existing_pairs:
                        conn.execute("INSERT INTO order_change_tag_keywords (tag_id,scope,keyword,created_at) VALUES (?,?,?,?)", (tag["id"], scope, keyword, now))
            conn.execute("INSERT INTO settings(key,value) VALUES (?,?)", (refinement_key, now))

        # 补齐明确的交期提前／加急表达；不使用孤立的“交期”或“最短交期”，
        # 避免把正常新订单的交期说明误判为订单变更。
        delivery_acceleration_key = f"order_change_delivery_acceleration_v6:{employee_id}"
        accelerated = conn.execute("SELECT 1 FROM settings WHERE key=?", (delivery_acceleration_key,)).fetchone()
        if not accelerated:
            delivery_keywords = (
                "交期提前", "交期协同提前", "交期加急", "提前交货", "提前交期",
            )
            scopes = ("subject", "body", "attachment_name", "attachment_content")
            change_group = conn.execute(
                "SELECT id FROM order_mail_rule_groups WHERE employee_id=? AND action_type='order_change' ORDER BY id LIMIT 1",
                (employee_id,),
            ).fetchone()
            if change_group:
                existing_pairs = {(row["scope"], row["keyword"]) for row in conn.execute(
                    "SELECT scope,keyword FROM order_mail_rule_keywords WHERE group_id=?", (change_group["id"],)
                ).fetchall()}
                for scope in scopes:
                    for keyword in delivery_keywords:
                        if (scope, keyword) not in existing_pairs:
                            conn.execute(
                                "INSERT INTO order_mail_rule_keywords (group_id,scope,keyword,created_at) VALUES (?,?,?,?)",
                                (change_group["id"], scope, keyword, now),
                            )
            delivery_tag = conn.execute(
                "SELECT id FROM order_change_tags WHERE employee_id=? AND name='交期 / 发货日期'", (employee_id,)
            ).fetchone()
            if delivery_tag:
                existing_pairs = {(row["scope"], row["keyword"]) for row in conn.execute(
                    "SELECT scope,keyword FROM order_change_tag_keywords WHERE tag_id=?", (delivery_tag["id"],)
                ).fetchall()}
                for scope in scopes:
                    for keyword in delivery_keywords:
                        if (scope, keyword) not in existing_pairs:
                            conn.execute(
                                "INSERT INTO order_change_tag_keywords (tag_id,scope,keyword,created_at) VALUES (?,?,?,?)",
                                (delivery_tag["id"], scope, keyword, now),
                            )
            conn.execute("INSERT INTO settings(key,value) VALUES (?,?)", (delivery_acceleration_key, now))


def list_universal_rules(employee_id: str) -> list[dict[str, Any]]:
    ensure_universal_rules(employee_id)
    with db_cursor() as conn:
        groups = [dict(row) for row in conn.execute("SELECT * FROM order_mail_rule_groups WHERE employee_id=? ORDER BY id", (employee_id,)).fetchall()]
        for group in groups:
            group["keywords"] = [dict(row) for row in conn.execute("SELECT * FROM order_mail_rule_keywords WHERE group_id=? ORDER BY id", (group["id"],)).fetchall()]
    return groups


def list_change_tags(employee_id: str) -> list[dict[str, Any]]:
    ensure_universal_rules(employee_id)
    with db_cursor() as conn:
        tags = [dict(row) for row in conn.execute("SELECT * FROM order_change_tags WHERE employee_id=? ORDER BY id", (employee_id,)).fetchall()]
        for tag in tags:
            tag["keywords"] = [dict(row) for row in conn.execute("SELECT * FROM order_change_tag_keywords WHERE tag_id=? ORDER BY id", (tag["id"],)).fetchall()]
    return tags


def save_universal_rule(employee_id: str, payload: dict[str, Any], group_id: int | None = None) -> dict[str, Any]:
    name, action_type = str(payload.get("name") or "").strip(), str(payload.get("action_type") or "").strip()
    keywords = [(str(x.get("scope") or "").strip(), str(x.get("keyword") or "").strip()) for x in payload.get("keywords", []) if str(x.get("keyword") or "").strip()]
    if not name or action_type not in ROUTABLE_ACTION_TYPES or not keywords or any(scope not in SCOPE_LABELS for scope, _ in keywords):
        raise ValueError("请填写规则名称、明确分流结果和至少一个有效关键词")
    now = utcnow()
    with db_cursor() as conn:
        if group_id:
            row = conn.execute("SELECT id FROM order_mail_rule_groups WHERE id=? AND employee_id=?", (group_id, employee_id)).fetchone()
            if not row: raise ValueError("规则不存在")
            conn.execute("UPDATE order_mail_rule_groups SET name=?,action_type=?,enabled=?,updated_at=? WHERE id=?", (name, action_type, 1 if payload.get("enabled", True) else 0, now, group_id))
            conn.execute("DELETE FROM order_mail_rule_keywords WHERE group_id=?", (group_id,))
            saved_id = group_id
        else:
            cursor = conn.execute("INSERT INTO order_mail_rule_groups (employee_id,name,action_type,enabled,created_at,updated_at) VALUES (?,?,?,?,?,?)", (employee_id, name, action_type, 1 if payload.get("enabled", True) else 0, now, now))
            saved_id = int(cursor.lastrowid)
        for scope, keyword in keywords:
            conn.execute("INSERT INTO order_mail_rule_keywords (group_id,scope,keyword,created_at) VALUES (?,?,?,?)", (saved_id, scope, keyword, now))
    reclassify_cases(employee_id)
    return next(item for item in list_universal_rules(employee_id) if item["id"] == saved_id)


def save_universal_rule_scope(employee_id: str, group_id: int, scope: str, keywords: list[str]) -> dict[str, Any]:
    """Replace just one search area of a fixed top-level business category."""
    if scope not in SCOPE_LABELS:
        raise ValueError("无效的检索位置")
    cleaned: list[str] = []
    for value in keywords:
        word = str(value or "").strip()
        if word and word not in cleaned:
            cleaned.append(word)
    now = utcnow()
    with db_cursor() as conn:
        group = conn.execute(
            "SELECT id FROM order_mail_rule_groups WHERE id=? AND employee_id=? AND action_type IN ('new_order','quotation','order_change')",
            (group_id, employee_id),
        ).fetchone()
        if not group:
            raise ValueError("规则分类不存在")
        conn.execute("DELETE FROM order_mail_rule_keywords WHERE group_id=? AND scope=?", (group_id, scope))
        for word in cleaned:
            conn.execute("INSERT INTO order_mail_rule_keywords (group_id,scope,keyword,created_at) VALUES (?,?,?,?)", (group_id, scope, word, now))
        conn.execute("UPDATE order_mail_rule_groups SET updated_at=? WHERE id=?", (now, group_id))
    reclassify_cases(employee_id)
    return next(item for item in list_universal_rules(employee_id) if item["id"] == group_id)


def save_change_tag(employee_id: str, name: str, keywords: list[dict[str, Any]], tag_id: int | None = None) -> int:
    """Save one business-facing order change type and its recognition phrases.

    ``all`` is deliberately a UI convenience: a business phrase should work
    wherever it appears, while the database still stores the four explicit
    source scopes required by the routing engine.
    """
    pairs: list[tuple[str, str]] = []
    for item in keywords:
        scope = str(item.get("scope") or "").strip()
        word = str(item.get("keyword") or "").strip()
        if not word:
            continue
        if scope == "all":
            pairs.extend((source_scope, word) for source_scope in SCOPE_LABELS)
        else:
            pairs.append((scope, word))
    pairs = list(dict.fromkeys(pairs))
    if not name.strip() or not pairs or any(scope not in SCOPE_LABELS for scope, _ in pairs):
        raise ValueError("请填写变更类型名称和至少一个业务说法")
    now = utcnow()
    with db_cursor() as conn:
        if tag_id:
            existing = conn.execute("SELECT id FROM order_change_tags WHERE id=? AND employee_id=?", (tag_id, employee_id)).fetchone()
            if not existing:
                raise ValueError("变更类型不存在或无权修改")
            conn.execute("UPDATE order_change_tags SET name=?,updated_at=? WHERE id=? AND employee_id=?", (name.strip(), now, tag_id, employee_id))
            conn.execute("DELETE FROM order_change_tag_keywords WHERE tag_id=?", (tag_id,)); saved_id = tag_id
        else:
            existing = conn.execute("SELECT id FROM order_change_tags WHERE employee_id=? AND name=?", (employee_id, name.strip())).fetchone()
            if existing:
                raise ValueError("该变更类型已存在")
            cursor=conn.execute("INSERT INTO order_change_tags (employee_id,name,enabled,created_at,updated_at) VALUES (?,?,?,?,?)", (employee_id,name.strip(),1,now,now)); saved_id=int(cursor.lastrowid)
        for scope, keyword in pairs: conn.execute("INSERT INTO order_change_tag_keywords (tag_id,scope,keyword,created_at) VALUES (?,?,?,?)", (saved_id,scope,keyword,now))
    reclassify_cases(employee_id)
    return saved_id


def _mail_values(row: dict[str, Any]) -> dict[str, str]:
    return {"subject": str(row.get("subject") or "").lower(), "body": str(row.get("body_text") or "").lower(), "attachment_name": str(row.get("attachment_names") or "").lower(), "attachment_content": str(row.get("attachment_content") or "").lower()}


def _extract_attachment_text(path: str) -> tuple[str, str]:
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if not file_path.is_file(): return "", "missing"
    try:
        if suffix in {".txt", ".csv"}: return file_path.read_text(encoding="utf-8", errors="replace")[:120000], "parsed"
        if suffix in {".xlsx", ".xlsm"}:
            from openpyxl import load_workbook
            book=load_workbook(file_path, read_only=True, data_only=True); values=[]
            for sheet in book.worksheets:
                for row in sheet.iter_rows(values_only=True): values.extend(str(v) for v in row if v is not None)
            return " ".join(values)[:120000], "parsed"
        if suffix == ".pdf":
            import pdfplumber
            with pdfplumber.open(file_path) as pdf: return " ".join((page.extract_text() or "") for page in pdf.pages[:20])[:120000], "parsed"
        return "", "unsupported"
    except Exception: return "", "failed"


def _attachment_content(conn, mail_id: int) -> str:
    rows = conn.execute("SELECT id,stored_path FROM mail_attachments WHERE mail_id=? AND is_inline=0", (mail_id,)).fetchall(); parts=[]
    for row in rows:
        cached=conn.execute("SELECT text_content,parse_status FROM mail_attachment_texts WHERE attachment_id=?", (row["id"],)).fetchone()
        if cached is None:
            text,status=_extract_attachment_text(row["stored_path"])
            conn.execute("INSERT INTO mail_attachment_texts (attachment_id,text_content,parse_status,updated_at) VALUES (?,?,?,?)", (row["id"],text,status,utcnow()))
            parts.append(text)
        else: parts.append(cached["text_content"])
    return " ".join(parts)


def reclassify_cases(employee_id: str, account_id: int | None = None) -> None:
    ensure_universal_rules(employee_id)
    groups, tags = list_universal_rules(employee_id), list_change_tags(employee_id)
    clauses=["c.employee_id=?", "c.routing_source != 'manual'"]; params: list[Any]=[employee_id]
    if account_id: clauses.append("m.account_id=?"); params.append(account_id)
    with db_cursor() as conn:
        needs_attachment_content=any(k["scope"] == "attachment_content" for group in groups for k in group["keywords"]) or any(k["scope"] == "attachment_content" for tag in tags for k in tag["keywords"])
        rows=conn.execute(f"""SELECT c.id,c.mail_id,m.subject,m.body_text,COALESCE((SELECT GROUP_CONCAT(a.filename,' ') FROM mail_attachments a WHERE a.mail_id=m.id AND a.is_inline=0),'') attachment_names FROM order_intake_cases c JOIN mail_messages m ON m.id=c.mail_id WHERE {' AND '.join(clauses)}""",params).fetchall()
        now=utcnow()
        for raw in rows:
            row=dict(raw); row["attachment_content"]=_attachment_content(conn,row["mail_id"]) if needs_attachment_content else ""; values=_mail_values(row); matches=[]; clear_action_types=set(); assist_action_types=set(); tag_names=[]
            for group in groups:
                if not group["enabled"]: continue
                for keyword in group["keywords"]:
                    if keyword["keyword"].lower() in values[keyword["scope"]]:
                        strength = "clear" if keyword["scope"] in CLEAR_ROUTING_SCOPES else "assist"
                        matches.append({"scope":keyword["scope"],"keyword":keyword["keyword"],"group":group["name"],"action_type":group["action_type"],"strength":strength})
                        (clear_action_types if strength == "clear" else assist_action_types).add(group["action_type"])
            for tag in tags:
                tag_hits = [k for k in tag["keywords"] if tag["enabled"] and k["keyword"].lower() in values[k["scope"]]]
                if tag_hits:
                    tag_names.append(tag["name"])
                    for hit in tag_hits:
                        strength = "clear" if hit["scope"] in CLEAR_ROUTING_SCOPES else "assist"
                        matches.append({
                            "scope": hit["scope"], "keyword": hit["keyword"], "group": f"订单变更事项：{tag['name']}",
                            "action_type": "order_change", "source": "change_item", "strength": strength,
                        })
                        (clear_action_types if strength == "clear" else assist_action_types).add("order_change")

            # 先用主题、附件名称进行明确分流。正文和附件内容仅在没有明确
            # 结果时辅助判断；因此“采购订单 + PO + 附件内报价”仍明确属于录单。
            if len(clear_action_types) == 1:
                action_type = next(iter(clear_action_types)); state = "routed"; reason = "明确分流依据匹配"
            elif len(clear_action_types) > 1:
                action_type = "unclassified"; state = "needs_business_routing"; reason = "多个明确分流依据同时命中"
            elif len(assist_action_types) == 1:
                action_type = next(iter(assist_action_types)); state = "routed"; reason = "辅助识别线索匹配"
            elif len(assist_action_types) > 1:
                action_type = "unclassified"; state = "needs_business_routing"; reason = "多个辅助识别线索同时命中"
            else:
                action_type = "unclassified"; state = "unrouted"; reason = "未命中通用规则"
            conn.execute("UPDATE order_intake_cases SET action_type=?,routing_state=?,routing_source='keyword_rule',routing_reason=?,routing_matches_json=?,change_tags_json=?,updated_at=? WHERE id=?", (action_type,state,reason,json.dumps(matches,ensure_ascii=False),json.dumps(tag_names,ensure_ascii=False),now,row["id"]))


def bootstrap_cases(employee_id: str, account_id: int | None = None) -> None:
    now=utcnow()
    created_count = 0
    with db_cursor() as conn:
        rows=conn.execute("""SELECT m.id mail_id,COALESCE(t.customer_code,'') customer_code,COALESCE(t.customer_name,'') customer_name,COALESCE(t.order_number,'') order_number FROM mail_messages m JOIN mail_accounts a ON a.id=m.account_id LEFT JOIN mail_order_tasks t ON t.id=(SELECT id FROM mail_order_tasks WHERE mail_id=m.id ORDER BY id LIMIT 1) WHERE a.owner_employee_id=? AND (? IS NULL OR m.account_id=?)""",(employee_id,account_id,account_id)).fetchall()
        for row in rows:
            cursor = conn.execute("INSERT OR IGNORE INTO order_intake_cases (employee_id,mail_id,action_type,customer_code,customer_name,order_number,routing_state,routing_source,routing_reason,created_at,updated_at) VALUES (?,?,?,?,?,?, 'unrouted','keyword_rule','未命中通用规则',?,?)",(employee_id,row["mail_id"],"unclassified",row["customer_code"],row["customer_name"],row["order_number"],now,now))
            created_count += int(cursor.rowcount > 0)
        revision_key = f"order_intake_rule_engine_v6:{employee_id}"
        revision_needed = conn.execute("SELECT 1 FROM settings WHERE key=?", (revision_key,)).fetchone() is None
    if created_count or revision_needed:
        reclassify_cases(employee_id, account_id)
        with db_cursor() as conn:
            conn.execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (revision_key, now))


def list_cases(
    employee_id: str,
    target_date: str | None = None,
    action_type: str = "all",
    account_id: int | None = None,
    fetch_task_id: int | None = None,
    prepare: bool = True,
) -> list[dict[str, Any]]:
    if prepare:
        bootstrap_cases(employee_id, account_id)
    clauses = ["c.employee_id = ?"]
    values: list[Any] = [employee_id]
    if account_id:
        clauses.append("m.account_id = ?")
        values.append(account_id)
    if fetch_task_id:
        clauses.append(
            "EXISTS (SELECT 1 FROM mail_fetch_task_messages x "
            "WHERE x.fetch_task_id = ? AND x.mail_id = m.id AND x.is_new = 1)"
        )
        values.append(fetch_task_id)
    if action_type == "needs_business_routing":
        clauses.append("c.routing_state = 'needs_business_routing'")
    elif action_type == "unclassified":
        clauses.append("c.routing_state = 'unrouted'")
    elif action_type in ACTION_LABELS:
        clauses.append("c.action_type = ?")
        values.append(action_type)
    if target_date and not fetch_task_id:
        clauses.append("substr(COALESCE(NULLIF(m.sent_at, ''), NULLIF(m.received_at, ''), m.created_at), 1, 10) = ?")
        values.append(target_date)
    with db_cursor() as conn:
        rows = conn.execute(
            f"""
            SELECT c.*, m.subject, m.sender, m.sent_at, m.received_at, a.email AS mailbox_email,
                   rr.name AS routing_rule_name,
                   (SELECT COUNT(*) FROM mail_attachments a WHERE a.mail_id = m.id AND a.is_inline = 0) AS attachment_count,
                   (SELECT COUNT(*) FROM mail_order_tasks t WHERE t.mail_id = m.id) AS line_count
            FROM order_intake_cases c
            JOIN mail_messages m ON m.id = c.mail_id
            JOIN mail_accounts a ON a.id = m.account_id
            LEFT JOIN order_mail_routing_rules rr ON rr.id = c.routing_rule_id
            WHERE {' AND '.join(clauses)}
            ORDER BY CASE c.routing_state WHEN 'needs_business_routing' THEN 0 WHEN 'unrouted' THEN 4 ELSE 1 END,
                     COALESCE(m.sent_at, m.created_at) DESC, c.id DESC
            """,
            values,
        ).fetchall()
    today = business_today().isoformat()
    yesterday = (business_today() - timedelta(days=1)).isoformat()
    result = [dict(row) for row in rows]
    for item in result:
        item["received_day"] = str(item.get("sent_at") or item.get("received_at") or "")[:10]
        item["date_label"] = "今天" if item["received_day"] == today else "昨天" if item["received_day"] == yesterday else item["received_day"] or "日期待确认"
        item["routing_matches"] = json.loads(item.get("routing_matches_json") or "[]")
        item["change_tags"] = json.loads(item.get("change_tags_json") or "[]")
    return result


def list_date_counts(employee_id: str, account_id: int, days: int = 30, *, prepare: bool = True) -> list[dict[str, Any]]:
    if prepare:
        bootstrap_cases(employee_id, account_id)
    start_date = (business_today() - timedelta(days=max(1, days) - 1)).isoformat()
    with db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT substr(COALESCE(NULLIF(m.sent_at, ''), NULLIF(m.received_at, ''), m.created_at), 1, 10) AS mail_date,
                   COUNT(*) AS total,
                   SUM(CASE WHEN c.action_type = 'new_order' THEN 1 ELSE 0 END) AS new_order_count,
                   SUM(CASE WHEN c.action_type = 'order_change' THEN 1 ELSE 0 END) AS change_count,
                   SUM(CASE WHEN c.action_type = 'quotation' THEN 1 ELSE 0 END) AS quotation_count
            FROM order_intake_cases c
            JOIN mail_messages m ON m.id = c.mail_id
            WHERE c.employee_id = ? AND m.account_id = ?
              AND substr(COALESCE(NULLIF(m.sent_at, ''), NULLIF(m.received_at, ''), m.created_at), 1, 10) >= ?
            GROUP BY mail_date ORDER BY mail_date DESC
            """,
            (employee_id, account_id, start_date),
        ).fetchall()
    return [dict(row) for row in rows]


def get_case(case_id: int, employee_id: str) -> dict[str, Any] | None:
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT c.*, m.subject, m.sender, m.sent_at, m.received_at, m.body_html, m.body_text, m.eml_path,
                   rr.name AS routing_rule_name
            FROM order_intake_cases c
            JOIN mail_messages m ON m.id = c.mail_id
            LEFT JOIN order_mail_routing_rules rr ON rr.id = c.routing_rule_id
            WHERE c.id = ? AND c.employee_id = ?
            """,
            (case_id, employee_id),
        ).fetchone()
        if not row:
            return None
        data = dict(row)
        data["attachments"] = [dict(item) for item in conn.execute(
            "SELECT * FROM mail_attachments WHERE mail_id = ? ORDER BY is_inline, id", (data["mail_id"],)
        ).fetchall()]
        data["lines"] = [dict(item) for item in conn.execute(
            "SELECT * FROM mail_order_tasks WHERE mail_id = ? ORDER BY id", (data["mail_id"],)
        ).fetchall()]
        data["events"] = [dict(item) for item in conn.execute(
            "SELECT * FROM order_intake_case_events WHERE case_id = ? ORDER BY id DESC", (case_id,)
        ).fetchall()]
        data["routing_matches"] = json.loads(data.get("routing_matches_json") or "[]")
        data["change_tags"] = json.loads(data.get("change_tags_json") or "[]")
        from .mail_transcode_agent.mail_html_parser import extract_order_fields, safe_display_html
        data["display_html"] = safe_display_html(str(data.get("body_html") or ""), str(data.get("body_text") or ""))
        data["detected_fields"] = extract_order_fields(str(data.get("body_text") or ""), str(data.get("sender") or ""))
    return data


def get_attachment(attachment_id: int, employee_id: str) -> dict[str, Any] | None:
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT a.* FROM mail_attachments a
            JOIN order_intake_cases c ON c.mail_id = a.mail_id
            WHERE a.id = ? AND c.employee_id = ?
            """,
            (attachment_id, employee_id),
        ).fetchone()
    return dict(row) if row else None


def update_case(case_id: int, employee_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    current = get_case(case_id, employee_id)
    if not current:
        raise ValueError("订单接入任务不存在或无权操作")
    action_type = str(payload.get("action_type") or current["action_type"]).strip()
    status = str(payload.get("status") or current["status"]).strip()
    if action_type not in ACTION_LABELS or status not in STATUS_LABELS:
        raise ValueError("订单动作或处理状态无效")
    workflow_stage = str(payload.get("workflow_stage") or current["workflow_stage"]).strip()
    customer_match_status = str(payload.get("customer_match_status") or current["customer_match_status"]).strip()
    source_document_status = str(payload.get("source_document_status") or current["source_document_status"]).strip()
    mapping_status = str(payload.get("mapping_status") or current["mapping_status"]).strip()
    erp_prepare_status = str(payload.get("erp_prepare_status") or current["erp_prepare_status"]).strip()
    if workflow_stage not in WORKFLOW_STAGE_LABELS:
        raise ValueError("录单流程阶段无效")
    if customer_match_status not in MATCH_STATUS_LABELS or source_document_status not in DOCUMENT_STATUS_LABELS:
        raise ValueError("客户归属或订单资料状态无效")
    if mapping_status not in MAPPING_STATUS_LABELS or erp_prepare_status not in ERP_PREPARE_STATUS_LABELS:
        raise ValueError("映射或录单准备状态无效")
    after = {
        "action_type": action_type,
        "status": status,
        "customer_code": str(payload.get("customer_code") or "").strip(),
        "customer_name": str(payload.get("customer_name") or "").strip(),
        "order_number": str(payload.get("order_number") or "").strip(),
        "order_version": str(payload.get("order_version") or "").strip(),
        "parent_order_number": str(payload.get("parent_order_number") or "").strip(),
        "workflow_stage": workflow_stage,
        "customer_match_status": customer_match_status,
        "source_document_status": source_document_status,
        "mapping_status": mapping_status,
        "erp_prepare_status": erp_prepare_status,
        "handling_note": str(payload.get("handling_note") or "").strip(),
    }
    if action_type == "new_order" and status == "ready_for_erp" and not (after["customer_name"] and after["order_number"]):
        raise ValueError("确认进入 ERP 前，请补齐客户和客户 PO 号")
    now = utcnow()
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE order_intake_cases SET action_type=?, status=?, customer_code=?, customer_name=?,
                order_number=?, order_version=?, parent_order_number=?, workflow_stage=?, customer_match_status=?,
                source_document_status=?, mapping_status=?, erp_prepare_status=?, handling_note=?,
                routing_source='manual', routing_state='business_routed', routing_reason='业务人员手工调整', routing_rule_id=NULL, routed_by=?, routed_at=?,
                confirmed_by=?, confirmed_at=?, completed_at=?, updated_at=?
            WHERE id=? AND employee_id=?
            """,
            (*after.values(), employee_id, now, employee_id, now,
             now if status == "archived" and current["status"] != "archived" else (current.get("completed_at") if status == "archived" else None),
             now, case_id, employee_id),
        )
        conn.execute(
            """
            INSERT INTO order_intake_case_events (case_id, employee_id, action, before_json, after_json, created_at)
            VALUES (?, ?, 'manual_confirmation', ?, ?, ?)
            """,
            (case_id, employee_id, json.dumps({key: current[key] for key in after}, ensure_ascii=False), json.dumps(after, ensure_ascii=False), now),
        )
    return get_case(case_id, employee_id) or {}


def update_routing(case_id: int, employee_id: str, action_type: str) -> dict[str, Any]:
    current = get_case(case_id, employee_id)
    if not current:
        raise ValueError("邮件不存在或无权操作")
    if action_type not in ACTION_LABELS:
        raise ValueError("分流类型无效")
    now = utcnow()
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE order_intake_cases
            SET action_type = ?, routing_source = 'manual', routing_state = 'business_routed', routing_reason = '业务人员手工调整', routing_rule_id = NULL,
                routed_by = ?, routed_at = ?, updated_at = ?
            WHERE id = ? AND employee_id = ?
            """,
            (action_type, employee_id, now, now, case_id, employee_id),
        )
        conn.execute(
            """
            INSERT INTO order_intake_case_events
                (case_id, employee_id, action, before_json, after_json, created_at)
            VALUES (?, ?, 'routing_adjusted', ?, ?, ?)
            """,
            (
                case_id,
                employee_id,
                json.dumps({"action_type": current["action_type"]}, ensure_ascii=False),
                json.dumps({"action_type": action_type}, ensure_ascii=False),
                now,
            ),
        )
    return get_case(case_id, employee_id) or {}


def case_summary(cases: list[dict[str, Any]]) -> dict[str, int]:
    return {
        "total": len(cases),
        **{key: sum(1 for item in cases if item["action_type"] == key) for key in ACTION_LABELS},
        "needs_business_routing": sum(1 for item in cases if item.get("routing_state") == "needs_business_routing"),
        "unrouted": sum(1 for item in cases if item.get("routing_state") == "unrouted"),
    }


def work_summary(
    employee_id: str,
    account_id: int | None = None,
    *,
    target_date: str | None = None,
) -> dict[str, Any]:
    """Return workload for all cases whose source mail belongs to ``target_date``.

    When no date is supplied this remains the cross-date summary used by older
    callers.  A supplied date deliberately scopes both open and completed work
    to the source-mail date, so a user can close the loop for one mail day.
    """
    clauses = ["c.employee_id=?"]
    params: list[Any] = [employee_id]
    if account_id:
        clauses.append("m.account_id=?")
        params.append(account_id)
    if target_date:
        clauses.append(
            "substr(COALESCE(NULLIF(m.sent_at, ''), NULLIF(m.received_at, ''), m.created_at), 1, 10)=?"
        )
        params.append(target_date)
    with db_cursor() as conn:
        rows = conn.execute(
            f"""SELECT c.action_type,c.status,c.routing_state,c.completed_at
                FROM order_intake_cases c JOIN mail_messages m ON m.id=c.mail_id
                WHERE {' AND '.join(clauses)}""",
            params,
        ).fetchall()
    result: dict[str, Any] = {
        "needs_routing": 0, "pending": 0, "in_progress": 0,
        "awaiting_confirmation": 0, "on_hold": 0, "active_total": 0,
        "completed_today": 0, "by_type": {key: 0 for key in ROUTABLE_ACTION_TYPES},
    }
    today = business_today().isoformat()
    for raw in rows:
        item = dict(raw)
        if item["status"] == "archived":
            if target_date or _business_date(str(item.get("completed_at") or "")) == today:
                result["completed_today"] += 1
            continue
        if item["routing_state"] == "needs_business_routing":
            result["needs_routing"] += 1
            continue
        if item["action_type"] not in ROUTABLE_ACTION_TYPES:
            continue
        result["active_total"] += 1
        result["by_type"][item["action_type"]] += 1
        if item["status"] == "pending_triage":
            result["pending"] += 1
        elif item["status"] == "pending_review":
            result["in_progress"] += 1
        elif item["status"] == "ready_for_erp":
            result["awaiting_confirmation"] += 1
        elif item["status"] == "on_hold":
            result["on_hold"] += 1
    return result
