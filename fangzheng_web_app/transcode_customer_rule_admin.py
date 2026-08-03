from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Iterable, Mapping

from .db import db_cursor
from .transcode_agent_standard import OFFICIAL_GRADE_CODES
from .transcode_semantic_rule_finalizer import validate_atomic_conditions
from .transcode_semantic_service import TARGET_FIELDS


BUSINESS_FIELD_TARGETS = {
    "胶系": ("glue",),
    "基板厚度": ("thickness",),
    "铜箔规格": ("copper",),
    "基板尺寸": ("size",),
    "胶水类别": ("glue_category",),
    "铜箔类型+印字/非印字": ("copper_type", "print_mark"),
    "基板级别": ("grade_intent",),
    "总/芯厚": ("total_core",),
}
BUSINESS_FIELDS = tuple(BUSINESS_FIELD_TARGETS)
TARGET_FIELD_LABELS = {
    "glue": "胶系代码",
    "thickness": "基板厚度",
    "copper": "铜箔规格",
    "size": "基板尺寸",
    "glue_category": "胶水类别",
    "copper_type": "铜箔类型",
    "print_mark": "印字/非印字",
    "grade_intent": "基板级别",
    "total_core": "总/芯厚",
}
CONDITION_FIELDS = (
    "订单备注",
    "订单规格",
    "订单规格/订单备注",
    "胶系",
    "基板厚度",
    "铜箔规格",
    "铜箔上面oz",
    "铜箔下面oz",
    "客户规格",
    "客户物料编码",
    "客户料品名称",
    "品号/物料编号",
)
CONDITION_OPERATORS = (
    "contains_any",
    "contains_all",
    "not_contains",
    "equals",
    "not_equals",
    "in",
    "not_in",
    "lt",
    "lte",
    "gt",
    "gte",
    "missing",
    "present",
)
CONDITION_OPERATOR_LABELS = {
    "contains_any": "包含任一",
    "contains_all": "同时包含",
    "not_contains": "不包含",
    "equals": "等于",
    "not_equals": "不等于",
    "in": "属于",
    "not_in": "不属于",
    "lt": "小于",
    "lte": "小于等于",
    "gt": "大于",
    "gte": "大于等于",
    "missing": "为空",
    "present": "有值",
}
LIST_OPERATORS = {"contains_any", "contains_all", "not_contains", "in", "not_in"}
NUMBER_OPERATORS = {"lt", "lte", "gt", "gte"}
AGENT_ASSET_TYPE = "agent_deterministic"
SEMANTIC_ASSET_TYPE = "semantic"
AGENT_OVERRIDE_TO_BUSINESS_FIELD = {
    "glue_code": "胶系",
    "thickness_code": "基板厚度",
    "copper_code": "铜箔规格",
    "size_code": "基板尺寸",
    "glue_category_code": "胶水类别",
    "copper_type_code": "铜箔类型+印字/非印字",
    "grade_code": "基板级别",
    "tc_code": "总/芯厚",
}
AGENT_OVERRIDE_TO_TARGET = {
    "glue_code": "glue",
    "thickness_code": "thickness",
    "copper_code": "copper",
    "size_code": "size",
    "glue_category_code": "glue_category",
    "copper_type_code": "copper_type",
    "grade_code": "grade_intent",
    "tc_code": "total_core",
}


class CustomerRuleMaintenanceError(ValueError):
    pass


def ensure_customer_rule_maintenance_tables() -> None:
    with db_cursor() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS transcode_customer_rule_overrides (
                rule_id TEXT PRIMARY KEY,
                rule_json TEXT,
                deleted INTEGER NOT NULL DEFAULT 0,
                updated_by TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS transcode_customer_rule_changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rule_id TEXT NOT NULL,
                action TEXT NOT NULL,
                employee_id TEXT NOT NULL,
                before_json TEXT,
                after_json TEXT,
                created_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_transcode_customer_rule_changes_rule
            ON transcode_customer_rule_changes(rule_id, id DESC);

            CREATE TABLE IF NOT EXISTS transcode_agent_rule_overrides (
                rule_id TEXT PRIMARY KEY,
                rule_json TEXT,
                deleted INTEGER NOT NULL DEFAULT 0,
                updated_by TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            """
        )


def merge_customer_rule_overrides(base_rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ensure_customer_rule_maintenance_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT rule_id, rule_json, deleted FROM transcode_customer_rule_overrides"
        ).fetchall()
    overrides = {str(row["rule_id"]): row for row in rows}
    base_ids = {str(rule.get("rule_id") or "") for rule in base_rules}
    merged: list[dict[str, Any]] = []
    for rule in base_rules:
        rule_id = str(rule.get("rule_id") or "")
        override = overrides.get(rule_id)
        if override is None:
            merged.append(rule)
        elif not int(override["deleted"] or 0):
            merged.append(_parse_rule_json(override["rule_json"], rule_id))
    for rule_id, row in overrides.items():
        if rule_id in base_ids or int(row["deleted"] or 0):
            continue
        merged.append(_parse_rule_json(row["rule_json"], rule_id))
    return merged


def merge_agent_rule_overrides(base_rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Apply page-maintained overrides without modifying versioned Agent Excel assets."""
    ensure_customer_rule_maintenance_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT rule_id, rule_json, deleted FROM transcode_agent_rule_overrides"
        ).fetchall()
    overrides = {str(row["rule_id"]): row for row in rows}
    base_ids = {str(rule.get("规则ID") or "") for rule in base_rules}
    merged: list[dict[str, Any]] = []
    for rule in base_rules:
        rule_id = str(rule.get("规则ID") or "")
        override = overrides.get(rule_id)
        if override is None:
            merged.append(rule)
        elif not int(override["deleted"] or 0):
            merged.append(_parse_rule_json(override["rule_json"], rule_id))
    for rule_id, row in overrides.items():
        if rule_id in base_ids or int(row["deleted"] or 0):
            continue
        merged.append(_parse_rule_json(row["rule_json"], rule_id))
    return merged


def agent_rules_for_customer_workspace(rules: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return [_agent_rule_as_workspace_rule(rule) for rule in rules if _clean(rule.get("规则ID"))]


def customer_rule_workspace(
    rules: list[dict[str, Any]],
    *,
    search: str = "",
    customer_key: str = "",
    business_field: str = "",
    rule_id: str = "",
    rule_kind: str = "all",
) -> dict[str, Any]:
    selected_kind = rule_kind if rule_kind in {"all", "deterministic", "semantic"} else "all"
    rules = [rule for rule in rules if _rule_matches_kind(rule, selected_kind)]
    override_ids = _override_ids()
    customers_by_key: dict[str, dict[str, Any]] = {}
    for rule in rules:
        key = make_customer_key(rule.get("customer_code"), rule.get("customer_name"))
        customer = customers_by_key.setdefault(
            key,
            {
                "key": key,
                "code": _clean(rule.get("customer_code")),
                "name": _clean(rule.get("customer_name")) or "未命名客户",
                "rule_count": 0,
                "field_counts": {field: 0 for field in BUSINESS_FIELDS},
            },
        )
        customer["rule_count"] += 1
        field = _clean(rule.get("business_field"))
        if field in customer["field_counts"]:
            customer["field_counts"][field] += 1

    query = _normalize(search)
    customers = sorted(
        (
            customer
            for customer in customers_by_key.values()
            if not query
            or query in _normalize(customer["code"])
            or query in _normalize(customer["name"])
        ),
        key=lambda item: (-int(item["rule_count"]), item["name"]),
    )
    selected_key = customer_key if customer_key in customers_by_key else ""
    if not selected_key and customers:
        selected_key = customers[0]["key"]
    selected_customer = customers_by_key.get(selected_key)
    selected_field = business_field if business_field in BUSINESS_FIELDS else "基板级别"
    if selected_customer and not selected_customer["field_counts"].get(selected_field):
        selected_field = next(
            (
                field
                for field, count in selected_customer["field_counts"].items()
                if count
            ),
            selected_field,
        )

    selected_rules: list[dict[str, Any]] = []
    if selected_customer:
        selected_rules = [
            _rule_view(rule, overridden=_clean(rule.get("rule_id")) in override_ids)
            for rule in rules
            if make_customer_key(rule.get("customer_code"), rule.get("customer_name")) == selected_key
            and _clean(rule.get("business_field")) == selected_field
        ]
        selected_rules.sort(key=lambda item: (-int(item["priority"]), item["rule_id"]))
    selected_rule = next(
        (item for item in selected_rules if item["rule_id"] == rule_id),
        selected_rules[0] if selected_rules else None,
    )
    return {
        "customers": customers,
        "selected_customer": selected_customer,
        "selected_customer_key": selected_key,
        "business_fields": BUSINESS_FIELDS,
        "selected_field": selected_field,
        "rules": selected_rules,
        "selected_rule": selected_rule,
        "rule_count": len(rules),
        "override_count": sum(
            1 for rule in rules if _clean(rule.get("rule_id")) in override_ids
        ),
        "rule_kind": selected_kind,
    }


def _rule_matches_kind(rule: Mapping[str, Any], rule_kind: str) -> bool:
    if rule_kind == "all":
        return True
    source_fields = {
        _clean(item.get("field"))
        for item in rule.get("conditions") or []
        if isinstance(item, Mapping)
    }
    is_order_semantic = any("订单备注" in field for field in source_fields)
    return is_order_semantic if rule_kind == "semantic" else not is_order_semantic


def build_rule_from_form(
    form: Mapping[str, Any],
    *,
    existing_rule: dict[str, Any] | None = None,
) -> dict[str, Any]:
    rule_id = _clean(form.get("rule_id")) or _new_rule_id()
    customer_code = _clean(form.get("customer_code"))
    customer_name = _clean(form.get("customer_name"))
    business_field = _clean(form.get("business_field"))
    source_text = _clean(form.get("source_text"))
    target_field = _clean(form.get("target_field"))
    target_value = _clean(form.get("target_value"))
    if not customer_code and not customer_name:
        raise CustomerRuleMaintenanceError("客户代码和客户简称至少填写一个。")
    if business_field not in BUSINESS_FIELD_TARGETS:
        raise CustomerRuleMaintenanceError(f"不支持的维护参数：{business_field}")
    if target_field not in BUSINESS_FIELD_TARGETS[business_field]:
        raise CustomerRuleMaintenanceError("目标字段与维护参数不匹配。")
    if target_field not in TARGET_FIELDS or not target_value:
        raise CustomerRuleMaintenanceError("目标结果不能为空。")
    if not source_text:
        raise CustomerRuleMaintenanceError("业务触发条件不能为空。")

    condition_fields = _form_list(form, "condition_field")
    condition_operators = _form_list(form, "condition_operator")
    condition_values = _form_list(form, "condition_value")
    conditions = _build_conditions(condition_fields, condition_operators, condition_values)
    validate_atomic_conditions(conditions, context=f"规则{rule_id}")
    _validate_target_value(target_field, target_value)
    priority = _to_int(form.get("priority"), 100)
    enabled = _form_bool(form, "enabled", default=True)
    semantic_enabled = _form_bool(form, "semantic_enabled", default=False)
    if semantic_enabled and not any(item["field"] == "订单备注" for item in conditions):
        raise CustomerRuleMaintenanceError("启用模型语义标准化时，至少需要一条订单备注条件。")

    prior = existing_rule or {}
    semantic_type = _semantic_type(conditions)
    return {
        "rule_id": rule_id,
        "source_candidate_id": _clean(prior.get("source_candidate_id")) or f"PAGE-{rule_id}",
        "enabled": enabled,
        "customer_code": customer_code,
        "customer_name": customer_name,
        "source_row": int(prior.get("source_row") or 0),
        "source_column": "页面维护",
        "business_field": business_field,
        "source_text": source_text,
        "semantic_types": [semantic_type],
        "target_fields": [target_field],
        "normalized_values": [target_value],
        "stated_target_values": [target_value],
        "conditions": conditions,
        "required_input_fields": list(dict.fromkeys(item["field"] for item in conditions)),
        "execution_mode": "结构化后可确定性执行",
        "priority": priority,
        "model": _clean(prior.get("model")) or ("页面语义维护" if semantic_enabled else "页面确定性维护"),
        "prompt_sha256": _clean(prior.get("prompt_sha256")),
        "evidence_texts": [source_text],
        "approval": {
            "status": "confirmed",
            "basis": _clean(form.get("approval_basis")) or "客户特殊规则维护页面确认",
        },
        "note": "页面维护；保存后立即生效"
        + ("；允许订单备注模型语义标准化" if semantic_enabled else ""),
    }


def build_agent_rule_from_form(
    form: Mapping[str, Any],
    *,
    existing_rule: dict[str, Any] | None,
) -> dict[str, Any]:
    if not existing_rule:
        raise CustomerRuleMaintenanceError("Agent确定性规则请从已有基础规则修改。")
    rule = dict(existing_rule)
    rule_id = _clean(rule.get("规则ID"))
    customer_code = _clean(form.get("customer_code"))
    customer_name = _clean(form.get("customer_name"))
    business_field = _clean(form.get("business_field"))
    override_field = _clean(form.get("agent_override_field"))
    override_value = _clean(form.get("target_value")).upper()
    if not customer_code and not customer_name:
        raise CustomerRuleMaintenanceError("客户代码和客户简称至少填写一个。")
    if business_field not in BUSINESS_FIELDS:
        raise CustomerRuleMaintenanceError(f"不支持的维护参数：{business_field}")
    if override_field not in AGENT_OVERRIDE_TO_BUSINESS_FIELD:
        raise CustomerRuleMaintenanceError(f"Agent覆盖字段无效：{override_field}")
    if AGENT_OVERRIDE_TO_BUSINESS_FIELD[override_field] != business_field:
        raise CustomerRuleMaintenanceError("Agent覆盖字段与维护参数不匹配。")
    if not override_value:
        raise CustomerRuleMaintenanceError("目标结果不能为空。")
    _validate_agent_override_value(override_field, override_value)
    rule.update(
        {
            "_asset_type": AGENT_ASSET_TYPE,
            "规则ID": rule_id,
            "启用": "是" if _form_bool(form, "enabled", default=True) else "否",
            "客户代码": customer_code,
            "客户简称": customer_name,
            "原始字段": business_field,
            "规则文本": _clean(form.get("source_text")) or _clean(rule.get("规则文本")),
            "条件胶系": _clean(form.get("agent_condition_glue")),
            "条件关键词": _clean(form.get("agent_condition_keyword")),
            "条件铜厚": _clean(form.get("agent_condition_copper")),
            "条件厚度": _clean(form.get("agent_condition_thickness")),
            "条件尺寸": _clean(form.get("agent_condition_size")),
            "覆盖字段": override_field,
            "覆盖值": override_value,
            "优先级": str(_to_int(form.get("priority"), _to_int(rule.get("优先级"), 100))),
            "强制执行": "是",
            "待确认": "否",
            "物料类别": "CCL",
            "命中来源": "页面维护Agent确定性长期规则",
            "规则解释": _clean(form.get("approval_basis")) or "客户特殊规则维护页面确认",
        }
    )
    return rule


def save_rule_override(
    rule: dict[str, Any],
    *,
    updated_by: str,
    previous_rule: dict[str, Any] | None,
) -> None:
    ensure_customer_rule_maintenance_tables()
    rule_id = _clean(rule.get("rule_id"))
    if not rule_id:
        raise CustomerRuleMaintenanceError("规则ID不能为空。")
    now = _now()
    rule_json = _dump(rule)
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_customer_rule_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 0, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json = excluded.rule_json,
                deleted = 0,
                updated_by = excluded.updated_by,
                updated_at = excluded.updated_at
            """,
            (rule_id, rule_json, updated_by, now),
        )
        _insert_rule_change(
            conn,
            rule_id=rule_id,
            action="新增" if previous_rule is None else "修改",
            employee_id=updated_by,
            before=previous_rule,
            after=rule,
            created_at=now,
        )


def save_agent_rule_override(
    rule: dict[str, Any],
    *,
    updated_by: str,
    previous_rule: dict[str, Any] | None,
) -> None:
    ensure_customer_rule_maintenance_tables()
    rule_id = _clean(rule.get("规则ID"))
    if not rule_id:
        raise CustomerRuleMaintenanceError("Agent规则ID不能为空。")
    now = _now()
    rule_json = _dump(rule)
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_agent_rule_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 0, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json = excluded.rule_json,
                deleted = 0,
                updated_by = excluded.updated_by,
                updated_at = excluded.updated_at
            """,
            (rule_id, rule_json, updated_by, now),
        )
        _insert_rule_change(
            conn,
            rule_id=rule_id,
            action="新增" if previous_rule is None else "修改",
            employee_id=updated_by,
            before=previous_rule,
            after=rule,
            created_at=now,
        )


def validate_customer_maintained_rule(rule: dict[str, Any]) -> None:
    from .transcode_semantic_rules import _validate_machine_rule

    _validate_machine_rule(rule, 0)


def delete_rule_override(
    rule: dict[str, Any],
    *,
    updated_by: str,
) -> None:
    ensure_customer_rule_maintenance_tables()
    rule_id = _clean(rule.get("rule_id"))
    if not rule_id:
        raise CustomerRuleMaintenanceError("规则ID不能为空。")
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_customer_rule_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 1, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json = excluded.rule_json,
                deleted = 1,
                updated_by = excluded.updated_by,
                updated_at = excluded.updated_at
            """,
            (rule_id, _dump(rule), updated_by, now),
        )
        _insert_rule_change(
            conn,
            rule_id=rule_id,
            action="删除",
            employee_id=updated_by,
            before=rule,
            after=None,
            created_at=now,
        )


def delete_agent_rule_override(rule: dict[str, Any], *, updated_by: str) -> None:
    ensure_customer_rule_maintenance_tables()
    rule_id = _clean(rule.get("规则ID"))
    if not rule_id:
        raise CustomerRuleMaintenanceError("Agent规则ID不能为空。")
    now = _now()
    payload = dict(rule)
    payload["_asset_type"] = AGENT_ASSET_TYPE
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_agent_rule_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 1, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json = excluded.rule_json,
                deleted = 1,
                updated_by = excluded.updated_by,
                updated_at = excluded.updated_at
            """,
            (rule_id, _dump(payload), updated_by, now),
        )
        _insert_rule_change(
            conn,
            rule_id=rule_id,
            action="删除",
            employee_id=updated_by,
            before=payload,
            after=None,
            created_at=now,
        )


def list_customer_rule_changes(limit: int = 30) -> list[dict[str, Any]]:
    ensure_customer_rule_maintenance_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT id, rule_id, action, employee_id, before_json, after_json, created_at
            FROM transcode_customer_rule_changes
            ORDER BY id DESC
            LIMIT ?
            """,
            (max(1, min(int(limit), 100)),),
        ).fetchall()
    result = []
    for row in rows:
        payload = _optional_rule_json(row["after_json"]) or _optional_rule_json(row["before_json"]) or {}
        result.append(
            {
                "id": int(row["id"]),
                "rule_id": row["rule_id"],
                "action": row["action"],
                "employee_id": row["employee_id"],
                "created_at": row["created_at"],
                "customer": payload.get("customer_name") or payload.get("customer_code")
                or payload.get("客户简称") or payload.get("客户代码") or "",
                "business_field": payload.get("business_field") or payload.get("原始字段")
                or AGENT_OVERRIDE_TO_BUSINESS_FIELD.get(_clean(payload.get("覆盖字段")), ""),
                "source_text": payload.get("source_text") or payload.get("规则文本") or "",
            }
        )
    return result


def restore_customer_rule_change(change_id: int, *, updated_by: str) -> str:
    ensure_customer_rule_maintenance_tables()
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT id, rule_id, before_json, after_json
            FROM transcode_customer_rule_changes
            WHERE id = ?
            """,
            (int(change_id),),
        ).fetchone()
        if row is None:
            raise CustomerRuleMaintenanceError("修改记录不存在。")
        rule_id = _clean(row["rule_id"])
        before = _optional_rule_json(row["before_json"])
        after = _optional_rule_json(row["after_json"])
        asset_payload = before or after or {}
        asset_type = _clean(asset_payload.get("_asset_type") or asset_payload.get("asset_type"))
        override_table = (
            "transcode_agent_rule_overrides"
            if asset_type == AGENT_ASSET_TYPE
            else "transcode_customer_rule_overrides"
        )
        current_row = conn.execute(
            f"SELECT rule_json, deleted FROM {override_table} WHERE rule_id = ?",
            (rule_id,),
        ).fetchone()
        current = None
        if current_row is not None and not int(current_row["deleted"] or 0):
            current = _optional_rule_json(current_row["rule_json"])
        now = _now()
        if before is None:
            conn.execute(
                f"""
                INSERT INTO {override_table}
                    (rule_id, rule_json, deleted, updated_by, updated_at)
                VALUES (?, NULL, 1, ?, ?)
                ON CONFLICT(rule_id) DO UPDATE SET
                    rule_json = NULL,
                    deleted = 1,
                    updated_by = excluded.updated_by,
                    updated_at = excluded.updated_at
                """,
                (rule_id, updated_by, now),
            )
        else:
            conn.execute(
                f"""
                INSERT INTO {override_table}
                    (rule_id, rule_json, deleted, updated_by, updated_at)
                VALUES (?, ?, 0, ?, ?)
                ON CONFLICT(rule_id) DO UPDATE SET
                    rule_json = excluded.rule_json,
                    deleted = 0,
                    updated_by = excluded.updated_by,
                    updated_at = excluded.updated_at
                """,
                (rule_id, _dump(before), updated_by, now),
            )
        _insert_rule_change(
            conn,
            rule_id=rule_id,
            action="恢复",
            employee_id=updated_by,
            before=current,
            after=before,
            created_at=now,
        )
    return rule_id


def find_rule(rules: Iterable[dict[str, Any]], rule_id: str) -> dict[str, Any] | None:
    target = _clean(rule_id)
    return next((rule for rule in rules if _clean(rule.get("rule_id")) == target), None)


def make_customer_key(code: Any, name: Any) -> str:
    return f"{_clean(code)}\x1f{_clean(name)}"


def _agent_rule_as_workspace_rule(rule: dict[str, Any]) -> dict[str, Any]:
    override_field = _clean(rule.get("覆盖字段"))
    business_field = _clean(rule.get("原始字段"))
    if business_field not in BUSINESS_FIELDS:
        business_field = AGENT_OVERRIDE_TO_BUSINESS_FIELD.get(override_field, "")
    conditions: list[dict[str, Any]] = []
    condition_specs = (
        ("胶系", "equals", rule.get("条件胶系")),
        ("客户规格", "contains_any", rule.get("条件关键词")),
        ("铜箔规格", "equals", rule.get("条件铜厚")),
        ("基板厚度", "equals", rule.get("条件厚度")),
        ("客户规格", "contains_any", rule.get("条件尺寸")),
    )
    for field, operator, value in condition_specs:
        cleaned = _clean(value)
        if cleaned:
            conditions.append({"field": field, "operator": operator, "value": cleaned})
    if not conditions:
        conditions.append({"field": "客户规格", "operator": "present", "value": ""})
    source_text = _clean(rule.get("规则文本")) or _clean(rule.get("条件文本"))
    return {
        "asset_type": AGENT_ASSET_TYPE,
        "rule_id": _clean(rule.get("规则ID")),
        "customer_code": _clean(rule.get("客户代码")),
        "customer_name": _clean(rule.get("客户简称")),
        "business_field": business_field,
        "source_text": source_text,
        "source_column": _clean(rule.get("来源字段")),
        "conditions": conditions,
        "target_fields": [AGENT_OVERRIDE_TO_TARGET.get(override_field, override_field)],
        "normalized_values": [_clean(rule.get("覆盖值"))],
        "priority": _to_int(rule.get("优先级"), 0),
        "enabled": _clean(rule.get("启用")) == "是",
        "model": "Agent确定性规则",
        "approval": {"basis": _clean(rule.get("规则解释")) or _clean(rule.get("命中来源"))},
        "agent_rule": rule,
    }


def _rule_view(rule: dict[str, Any], *, overridden: bool) -> dict[str, Any]:
    conditions = list(rule.get("conditions") or [])
    source_fields = list(dict.fromkeys(_clean(item.get("field")) for item in conditions if _clean(item.get("field"))))
    editable_conditions = [
        {
            "field": _clean(item.get("field")),
            "operator": _clean(item.get("operator")),
            "value": _condition_value_text(item.get("value")),
        }
        for item in conditions
    ]
    target_field = _first(rule.get("target_fields"))
    source_column = _clean(rule.get("source_column"))
    model = _clean(rule.get("model"))
    asset_type = _clean(rule.get("asset_type")) or SEMANTIC_ASSET_TYPE
    if asset_type == AGENT_ASSET_TYPE:
        origin = "页面维护Agent长期规则" if overridden else "Agent确定性长期规则"
    elif source_column == "确认中心" or model == "确认中心人工规则":
        origin = "确认中心长期规则"
    elif overridden:
        origin = "页面维护"
    else:
        origin = "正式规则表"
    return {
        "rule_id": _clean(rule.get("rule_id")),
        "asset_type": asset_type,
        "customer_code": _clean(rule.get("customer_code")),
        "customer_name": _clean(rule.get("customer_name")),
        "business_field": _clean(rule.get("business_field")),
        "source_text": _clean(rule.get("source_text")),
        "input_source": " + ".join(source_fields),
        "conditions": editable_conditions,
        "condition_summary": _condition_summary(editable_conditions),
        "target_field": target_field,
        "target_value": _first(rule.get("normalized_values")),
        "priority": int(rule.get("priority") or 0),
        "enabled": bool(rule.get("enabled")),
        "semantic_enabled": "订单备注" in source_fields,
        "approval_basis": _clean((rule.get("approval") or {}).get("basis")),
        "origin": origin,
        "machine_rule": rule,
        "agent_rule": dict(rule.get("agent_rule") or {}),
    }


def _build_conditions(fields: list[str], operators: list[str], values: list[str]) -> list[dict[str, Any]]:
    conditions: list[dict[str, Any]] = []
    row_count = max(len(fields), len(operators), len(values))
    for index in range(row_count):
        field = _clean(fields[index] if index < len(fields) else "")
        operator = _clean(operators[index] if index < len(operators) else "")
        raw_value = _clean(values[index] if index < len(values) else "")
        if not field and not operator and not raw_value:
            continue
        if field not in CONDITION_FIELDS:
            raise CustomerRuleMaintenanceError(f"第{index + 1}条条件字段无效：{field}")
        if operator not in CONDITION_OPERATORS:
            raise CustomerRuleMaintenanceError(f"第{index + 1}条条件操作符无效：{operator}")
        if operator in {"missing", "present"}:
            value: Any = ""
        elif operator in LIST_OPERATORS:
            value = [item.strip() for item in raw_value.replace(",", "；").split("；") if item.strip()]
            if not value:
                raise CustomerRuleMaintenanceError(f"第{index + 1}条条件值不能为空。")
        elif operator in NUMBER_OPERATORS:
            try:
                value = float(raw_value)
            except ValueError as exc:
                raise CustomerRuleMaintenanceError(f"第{index + 1}条条件值必须是数字。") from exc
        else:
            if not raw_value:
                raise CustomerRuleMaintenanceError(f"第{index + 1}条条件值不能为空。")
            value = raw_value
        conditions.append(
            {
                "field": field,
                "operator": operator,
                "value": value,
                "source_scope": field,
            }
        )
    if not conditions:
        raise CustomerRuleMaintenanceError("至少维护一条触发条件。")
    return conditions


def _validate_target_value(target_field: str, value: str) -> None:
    raw = _clean(value)
    upper = raw.upper()
    if target_field == "grade_intent" and upper not in {str(item).upper() for item in OFFICIAL_GRADE_CODES}:
        raise CustomerRuleMaintenanceError(f"基板级别代码无效：{raw}")
    if target_field == "glue_category" and upper not in {"Y", "R"}:
        raise CustomerRuleMaintenanceError("胶水类别只允许Y或R。")
    if target_field == "total_core" and raw.lower() not in {
        "core",
        "total",
        "c",
        "t",
        "芯厚",
        "总厚",
        "core_after_total_to_core_conversion",
    }:
        raise CustomerRuleMaintenanceError("总/芯厚结果无效。")
    if target_field == "thickness" and (not raw.isdigit() or len(raw) != 5):
        raise CustomerRuleMaintenanceError("厚度结果必须是5位厚度代码。")
    if target_field == "size" and not (
        (raw.isdigit() and len(raw) == 8) or raw == "height_plus_0.3"
    ):
        raise CustomerRuleMaintenanceError("尺寸结果必须是8位尺寸代码或height_plus_0.3。")


def _validate_agent_override_value(override_field: str, value: str) -> None:
    upper = _clean(value).upper()
    if override_field == "grade_code" and upper not in {str(item).upper() for item in OFFICIAL_GRADE_CODES}:
        raise CustomerRuleMaintenanceError(f"基板级别代码无效：{value}")
    if override_field == "glue_category_code" and upper not in {"Y", "R"}:
        raise CustomerRuleMaintenanceError("胶水类别只允许Y或R。")
    if override_field == "tc_code" and upper not in {"C", "T"}:
        raise CustomerRuleMaintenanceError("总/芯厚只允许C或T。")
    if override_field == "thickness_code" and (not upper.isdigit() or len(upper) != 5):
        raise CustomerRuleMaintenanceError("厚度结果必须是5位厚度代码。")
    if override_field == "size_code" and (not upper.isdigit() or len(upper) != 8):
        raise CustomerRuleMaintenanceError("尺寸结果必须是8位尺寸代码。")


def _semantic_type(conditions: list[dict[str, Any]]) -> str:
    if len(conditions) > 1:
        return "multi_condition"
    operator = conditions[0]["operator"]
    if operator in {"contains_any", "contains_all"}:
        return "keyword_present"
    if operator == "not_contains":
        return "keyword_absent"
    if operator == "missing":
        return "default_when_missing"
    if operator in NUMBER_OPERATORS:
        return "comparison"
    return "explicit_fact"


def _condition_summary(conditions: list[dict[str, str]]) -> str:
    return "；".join(
        f"{item['field']} {CONDITION_OPERATOR_LABELS.get(item['operator'], item['operator'])}"
        + (f" {item['value']}" if item["value"] else "")
        for item in conditions
    )


def _override_ids() -> set[str]:
    ensure_customer_rule_maintenance_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT rule_id FROM transcode_customer_rule_overrides WHERE deleted = 0
            UNION
            SELECT rule_id FROM transcode_agent_rule_overrides WHERE deleted = 0
            """
        ).fetchall()
    return {str(row["rule_id"]) for row in rows}


def _insert_rule_change(
    conn,
    *,
    rule_id: str,
    action: str,
    employee_id: str,
    before: dict[str, Any] | None,
    after: dict[str, Any] | None,
    created_at: str,
) -> None:
    conn.execute(
        """
        INSERT INTO transcode_customer_rule_changes
            (rule_id, action, employee_id, before_json, after_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            rule_id,
            action,
            employee_id,
            _dump(before) if before else None,
            _dump(after) if after else None,
            created_at,
        ),
    )


def _parse_rule_json(value: Any, rule_id: str) -> dict[str, Any]:
    parsed = _optional_rule_json(value)
    if not isinstance(parsed, dict):
        raise CustomerRuleMaintenanceError(f"页面维护规则JSON无效：{rule_id}")
    return parsed


def _optional_rule_json(value: Any) -> dict[str, Any] | None:
    if not value:
        return None
    try:
        parsed = json.loads(str(value))
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def _condition_value_text(value: Any) -> str:
    if isinstance(value, list):
        return "；".join(str(item) for item in value)
    return _clean(value)


def _form_list(form: Mapping[str, Any], key: str) -> list[str]:
    getter = getattr(form, "getlist", None)
    if callable(getter):
        return [str(item or "") for item in getter(key)]
    value = form.get(key)
    return list(value) if isinstance(value, (list, tuple)) else [str(value or "")]


def _form_bool(form: Mapping[str, Any], key: str, *, default: bool) -> bool:
    if key not in form:
        return False if hasattr(form, "getlist") else default
    return _clean(form.get(key)).lower() in {"1", "true", "yes", "on", "是"}


def _new_rule_id() -> str:
    return datetime.now().strftime("TCR-%Y%m%d-%H%M%S-%f")


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _dump(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _normalize(value: Any) -> str:
    return "".join(_clean(value).upper().split())


def _first(values: Any) -> str:
    if isinstance(values, (list, tuple)) and values:
        return _clean(values[0])
    return ""


def _to_int(value: Any, default: int) -> int:
    try:
        return int(float(_clean(value)))
    except ValueError:
        return default
