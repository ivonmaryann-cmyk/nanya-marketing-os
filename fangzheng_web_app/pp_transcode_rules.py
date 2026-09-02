from __future__ import annotations

import json
import re
from datetime import datetime, timedelta
from typing import Any

from .database import transcode_cursor as db_cursor
from .db import utcnow
from .paths import STORAGE_DIR
from .transcode_agent_glue_resolver import is_retired_agent_glue_mapping, resolve_agent_glue
from .transcode_agent_rules import (
    ensure_default_transcode_agent_rule_version,
    get_active_transcode_agent_rule_version,
    load_transcode_agent_mapping_tables,
)


FEATURE_KEY = "pp_transcode_agent"
BACKUP_DIR = STORAGE_DIR / "pp_transcode_rule_backups"

# 规则页面使用的业务字段。字段顺序即 PP 编码位置顺序，避免业务维护时混淆。
PP_FIELDS = (
    ("glue_code", "胶系", 2),
    ("glass_style", "玻布规格", 4),
    ("pp_length", "PP长度", 4),
    ("formula_category", "配方类别", 1),
    ("resin_content", "树脂含量", 3),
    ("glass_vendor", "玻布厂家", 2),
    ("pp_grade", "PP级别", 1),
    ("pp_narrow_width", "PP窄幅", 1),
    ("customer_code_segment", "客户码", 3),
    ("gt_code", "GT长短秒", 3),
    ("customer_product_code", "客户产品码", 2),
    ("formula_code", "配方代码", 1),
)
FIELD_META = {key: {"label": label, "width": width} for key, label, width in PP_FIELDS}
# 胶系与营销转码 Agent 共用同一份活动映射，不在 PP 规则库复制或维护。
SHARED_FIELDS = {"glue_code"}
BASE_FIELDS = {"glass_style", "pp_length", "formula_category", "resin_content"}
CUSTOMER_FIELDS = set(FIELD_META) - BASE_FIELDS - SHARED_FIELDS


def _normalize(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def _clean_output(field_key: str, value: object) -> str:
    width = FIELD_META[field_key]["width"]
    text = _normalize(value)
    if not text:
        return "*" * width
    return text[:width].ljust(width, "*")


def ensure_pp_transcode_tables() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    with db_cursor() as conn:
        if getattr(conn, "dialect", "sqlite") == "postgresql":
            return
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS pp_transcode_base_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                field_key TEXT NOT NULL,
                input_value TEXT NOT NULL,
                output_value TEXT NOT NULL,
                business_note TEXT NOT NULL DEFAULT '',
                enabled INTEGER NOT NULL DEFAULT 1,
                source_kind TEXT NOT NULL DEFAULT 'page',
                created_by TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_by TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL,
                UNIQUE(field_key, input_value)
            );

            CREATE TABLE IF NOT EXISTS pp_transcode_customer_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                target_field TEXT NOT NULL,
                conditions_json TEXT NOT NULL DEFAULT '[]',
                output_value TEXT NOT NULL,
                business_note TEXT NOT NULL DEFAULT '',
                enabled INTEGER NOT NULL DEFAULT 1,
                created_by TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_by TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS pp_transcode_rule_changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rule_scope TEXT NOT NULL,
                rule_id INTEGER,
                action TEXT NOT NULL,
                before_json TEXT NOT NULL DEFAULT '{}',
                after_json TEXT NOT NULL DEFAULT '{}',
                changed_by TEXT NOT NULL DEFAULT '',
                changed_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS pp_transcode_confirmation_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER NOT NULL,
                employee_id TEXT NOT NULL,
                excel_row INTEGER NOT NULL,
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                spec TEXT NOT NULL,
                order_remark TEXT NOT NULL DEFAULT '',
                pending_code TEXT NOT NULL DEFAULT '',
                confirmed_pending_code TEXT NOT NULL DEFAULT '',
                confidence INTEGER NOT NULL DEFAULT 0,
                summary TEXT NOT NULL DEFAULT '',
                field_evidence_json TEXT NOT NULL DEFAULT '[]',
                confirmation_status TEXT NOT NULL DEFAULT 'pending',
                confirmation_basis TEXT NOT NULL DEFAULT '',
                confirmed_by TEXT NOT NULL DEFAULT '',
                confirmed_at TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(job_id, excel_row)
            );
            """
        )
        confirmation_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(pp_transcode_confirmation_items)").fetchall()
        }
        if "confirmed_pending_code" not in confirmation_columns:
            conn.execute(
                "ALTER TABLE pp_transcode_confirmation_items "
                "ADD COLUMN confirmed_pending_code TEXT NOT NULL DEFAULT ''"
            )


def _record_change(conn, scope: str, rule_id: int | None, action: str, before: dict, after: dict, employee_id: str) -> None:
    conn.execute(
        """
        INSERT INTO pp_transcode_rule_changes
        (rule_scope, rule_id, action, before_json, after_json, changed_by, changed_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (scope, rule_id, action, json.dumps(before, ensure_ascii=False), json.dumps(after, ensure_ascii=False), employee_id, utcnow()),
    )


def _shared_glue_mapping_tables() -> tuple[dict[str, list[dict]], str]:
    """Return the active Marketing Agent glue master and aliases for PP read-only use.

    PP deliberately excludes the CCL-only ``Agent胶系选择规则`` table. Those rows
    can include CCL customer and keyword conditions; PP customer conditions are
    maintained in PP's own customer-special-rule page.
    """
    version = get_active_transcode_agent_rule_version()
    if not version:
        version = ensure_default_transcode_agent_rule_version()
    tables = load_transcode_agent_mapping_tables(version)
    shared_tables = dict(tables)
    shared_tables["Agent胶系选择规则"] = []
    return shared_tables, version


def _shared_glue_rows() -> list[dict[str, Any]]:
    tables, version = _shared_glue_mapping_tables()
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for table_name, input_column, label in (
        ("Agent胶系主表", "胶系名称", "营销转码Agent胶系主表"),
        ("Agent胶系兼容别名", "兼容名称", "营销转码Agent胶系别名"),
    ):
        for row in tables.get(table_name, []):
            input_value = str(row.get(input_column) or "").strip()
            output_value = _normalize(row.get("输出胶系代码"))
            if not input_value or not re.fullmatch(r"[A-Z0-9]{2}", output_value):
                continue
            if is_retired_agent_glue_mapping(row):
                continue
            key = (table_name, _normalize(input_value), output_value)
            if key in seen:
                continue
            seen.add(key)
            rule_id = str(row.get("映射ID") or f"{table_name}:{input_value}")
            result.append(
                {
                    "id": f"shared:{rule_id}",
                    "field_key": "glue_code",
                    "field_label": FIELD_META["glue_code"]["label"],
                    "input_value": input_value,
                    "output_value": output_value,
                    "business_note": f"{label}；活动版本：{version}",
                    "enabled": 1,
                    "source_kind": "marketing_agent_shared",
                    "readonly": True,
                }
            )
    return result


def resolve_shared_pp_glue(spec: str) -> dict[str, Any] | None:
    """Resolve PP glue with the active Marketing Agent master/alias mapping only."""
    tables, version = _shared_glue_mapping_tables()
    resolved = resolve_agent_glue(tables, spec=str(spec or ""))
    if not resolved:
        return None
    code = _normalize(resolved.get("code"))
    if not re.fullmatch(r"[A-Z0-9]{2}", code):
        return None
    name = str(resolved.get("name") or "").strip()
    rule_id = str(resolved.get("rule_id") or "")
    return {
        "id": f"shared:{rule_id or name}",
        "field_key": "glue_code",
        "field_label": FIELD_META["glue_code"]["label"],
        "input_value": name,
        "output_value": code,
        "business_note": f"{resolved.get('source') or '营销转码Agent胶系映射'}；活动版本：{version}",
        "enabled": 1,
        "source_kind": "marketing_agent_shared",
        "readonly": True,
        "uncertain": bool(resolved.get("uncertain")),
        "conflict": str(resolved.get("conflict") or ""),
    }


def seed_pp_transcode_rules() -> None:
    """Seed PP-owned page rules once; glue is shared from Marketing Agent at runtime."""
    ensure_pp_transcode_tables()
    seed_rules = [
        ("glass_style", "106", "0106", "旧编码规范示例：三位玻布规格补足四位"),
        ("glass_style", "1067", "1067", "旧编码规范示例"),
        ("glass_style", "1078", "1078", "旧编码规范示例"),
        ("glass_style", "1080", "1080", "旧编码规范示例"),
        ("glass_style", "1086", "1086", "旧编码规范示例"),
        ("glass_style", "2113", "2113", "旧编码规范示例"),
        ("glass_style", "2313", "2313", "旧编码规范示例"),
        ("pp_length", "50M", "050M", "旧编码规范示例：长度补足三位"),
        ("pp_length", "100M", "100M", "旧编码规范示例"),
        ("pp_length", "150M", "150M", "旧编码规范示例"),
        ("pp_length", "200M", "200M", "旧编码规范示例"),
        ("pp_length", "250M", "250M", "旧编码规范示例"),
        ("pp_length", "300M", "300M", "旧编码规范示例"),
        ("pp_length", "350M", "350M", "旧编码规范示例"),
    ]
    now = utcnow()
    with db_cursor() as conn:
        for field_key, input_value, output_value, note in seed_rules:
            conn.execute(
                """
                INSERT OR IGNORE INTO pp_transcode_base_rules
                (field_key, input_value, output_value, business_note, enabled, source_kind, created_at, updated_at)
                VALUES (?, ?, ?, ?, 1, 'legacy_seed', ?, ?)
                """,
                (field_key, _normalize(input_value), _clean_output(field_key, output_value), note, now, now),
            )

def _row_dict(row) -> dict[str, Any]:
    return {key: row[key] for key in row.keys()}


def _get_base_rule_from_conn(conn, rule_id: int) -> dict[str, Any] | None:
    row = conn.execute("SELECT * FROM pp_transcode_base_rules WHERE id = ?", (rule_id,)).fetchone()
    return _row_dict(row) if row else None


def _get_customer_rule_from_conn(conn, rule_id: int) -> dict[str, Any] | None:
    row = conn.execute("SELECT * FROM pp_transcode_customer_rules WHERE id = ?", (rule_id,)).fetchone()
    if not row:
        return None
    item = _row_dict(row)
    try:
        item["conditions"] = json.loads(item["conditions_json"] or "[]")
    except json.JSONDecodeError:
        item["conditions"] = []
    return item


def _pagination(page: int | None, page_size: int | None) -> tuple[int, int] | None:
    if page_size is None:
        return None
    safe_size = max(10, min(int(page_size), 100))
    safe_page = max(1, int(page or 1))
    return safe_size, (safe_page - 1) * safe_size


def list_base_rules(
    field_key: str | None = None,
    keyword: str = "",
    enabled: str = "all",
    *,
    page: int | None = None,
    page_size: int | None = None,
) -> list[dict[str, Any]]:
    ensure_pp_transcode_tables()
    if field_key == "glue_code":
        rows = _shared_glue_rows()
        token = _normalize(keyword)
        if token:
            rows = [
                item for item in rows
                if token in _normalize(item["input_value"])
                or token in _normalize(item["output_value"])
                or token in _normalize(item["business_note"])
            ]
        if enabled == "disabled":
            rows = []
        rows.sort(key=lambda item: (_normalize(item["input_value"]), item["output_value"]))
        pagination = _pagination(page, page_size)
        if pagination:
            size, offset = pagination
            rows = rows[offset: offset + size]
        return rows
    clauses = ["1=1"]
    values: list[Any] = []
    if field_key in FIELD_META:
        clauses.append("field_key = ?")
        values.append(field_key)
    if keyword.strip():
        clauses.append("(input_value LIKE ? OR output_value LIKE ? OR business_note LIKE ?)")
        token = f"%{keyword.strip()}%"
        values.extend([token, token, token])
    if enabled in {"enabled", "disabled"}:
        clauses.append("enabled = ?")
        values.append(1 if enabled == "enabled" else 0)
    pagination = _pagination(page, page_size)
    query = f"SELECT * FROM pp_transcode_base_rules WHERE {' AND '.join(clauses)} ORDER BY field_key, id"
    if pagination:
        query += " LIMIT ? OFFSET ?"
        values.extend(pagination)
    with db_cursor() as conn:
        rows = conn.execute(
            query,
            values,
        ).fetchall()
    result = []
    for row in rows:
        item = _row_dict(row)
        item["field_label"] = FIELD_META.get(item["field_key"], {}).get("label", item["field_key"])
        result.append(item)
    return result


def count_base_rules(field_key: str | None = None, keyword: str = "", enabled: str = "all") -> int:
    if field_key == "glue_code":
        if enabled == "disabled":
            return 0
        rows = _shared_glue_rows()
        token = _normalize(keyword)
        if token:
            rows = [
                item for item in rows
                if token in _normalize(item["input_value"])
                or token in _normalize(item["output_value"])
                or token in _normalize(item["business_note"])
            ]
        return len(rows)
    clauses = ["1=1"]
    values: list[Any] = []
    if field_key in FIELD_META:
        clauses.append("field_key = ?")
        values.append(field_key)
    if keyword.strip():
        clauses.append("(input_value LIKE ? OR output_value LIKE ? OR business_note LIKE ?)")
        token = f"%{keyword.strip()}%"
        values.extend([token, token, token])
    if enabled in {"enabled", "disabled"}:
        clauses.append("enabled = ?")
        values.append(1 if enabled == "enabled" else 0)
    with db_cursor() as conn:
        row = conn.execute(
            f"SELECT COUNT(*) AS total FROM pp_transcode_base_rules WHERE {' AND '.join(clauses)}", values
        ).fetchone()
    return int(row["total"] if row else 0)


def get_base_rule(rule_id: int) -> dict[str, Any] | None:
    with db_cursor() as conn:
        return _get_base_rule_from_conn(conn, rule_id)


def save_base_rule(data: dict[str, Any], employee_id: str) -> int:
    field_key = str(data.get("field_key") or "").strip()
    if field_key in SHARED_FIELDS:
        raise ValueError("PP 胶系由营销转码Agent统一维护，请前往营销转码Agent规则配置修改。")
    if field_key not in BASE_FIELDS:
        raise ValueError("请选择可维护的 PP 基础字段")
    input_value = _normalize(data.get("input_value"))
    if not input_value:
        raise ValueError("请填写客户规格中的识别写法")
    output_value = _clean_output(field_key, data.get("output_value"))
    note = str(data.get("business_note") or "").strip()
    enabled = 1 if data.get("enabled", True) else 0
    rule_id = int(data.get("id") or 0)
    now = utcnow()
    with db_cursor() as conn:
        if rule_id:
            before = _get_base_rule_from_conn(conn, rule_id) or {}
            conn.execute(
                """
                UPDATE pp_transcode_base_rules
                SET field_key=?, input_value=?, output_value=?, business_note=?, enabled=?, updated_by=?, updated_at=?
                WHERE id=?
                """,
                (field_key, input_value, output_value, note, enabled, employee_id, now, rule_id),
            )
            after = _get_base_rule_from_conn(conn, rule_id) or {}
            _record_change(conn, "base", rule_id, "update", before, after, employee_id)
            return rule_id
        cursor = conn.execute(
            """
            INSERT INTO pp_transcode_base_rules
            (field_key, input_value, output_value, business_note, enabled, source_kind, created_by, created_at, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, 'page', ?, ?, ?, ?)
            """,
            (field_key, input_value, output_value, note, enabled, employee_id, now, employee_id, now),
        )
        new_id = int(cursor.lastrowid)
        _record_change(conn, "base", new_id, "create", {}, _get_base_rule_from_conn(conn, new_id) or {}, employee_id)
        return new_id


def list_customer_rules(
    customer_code: str = "",
    customer_name: str = "",
    field_key: str | None = None,
    enabled: str = "all",
    *,
    page: int | None = None,
    page_size: int | None = None,
) -> list[dict[str, Any]]:
    ensure_pp_transcode_tables()
    clauses = ["1=1"]
    values: list[Any] = []
    if customer_code.strip() or customer_name.strip():
        clauses.append("(customer_code = ? OR customer_name = ? OR (customer_code = '' AND customer_name = ''))")
        values.extend([customer_code.strip(), customer_name.strip()])
    if field_key in CUSTOMER_FIELDS:
        clauses.append("target_field = ?")
        values.append(field_key)
    if enabled in {"enabled", "disabled"}:
        clauses.append("enabled = ?")
        values.append(1 if enabled == "enabled" else 0)
    pagination = _pagination(page, page_size)
    query = f"SELECT * FROM pp_transcode_customer_rules WHERE {' AND '.join(clauses)} ORDER BY id DESC"
    if pagination:
        query += " LIMIT ? OFFSET ?"
        values.extend(pagination)
    with db_cursor() as conn:
        rows = conn.execute(
            query,
            values,
        ).fetchall()
    result = []
    for row in rows:
        item = _row_dict(row)
        try:
            item["conditions"] = json.loads(item["conditions_json"] or "[]")
        except json.JSONDecodeError:
            item["conditions"] = []
        item["target_label"] = FIELD_META.get(item["target_field"], {}).get("label", item["target_field"])
        result.append(item)
    return result


def count_customer_rules(
    customer_code: str = "", customer_name: str = "", field_key: str | None = None, enabled: str = "all"
) -> int:
    clauses = ["1=1"]
    values: list[Any] = []
    if customer_code.strip() or customer_name.strip():
        clauses.append("(customer_code = ? OR customer_name = ? OR (customer_code = '' AND customer_name = ''))")
        values.extend([customer_code.strip(), customer_name.strip()])
    if field_key in CUSTOMER_FIELDS:
        clauses.append("target_field = ?")
        values.append(field_key)
    if enabled in {"enabled", "disabled"}:
        clauses.append("enabled = ?")
        values.append(1 if enabled == "enabled" else 0)
    with db_cursor() as conn:
        row = conn.execute(
            f"SELECT COUNT(*) AS total FROM pp_transcode_customer_rules WHERE {' AND '.join(clauses)}", values
        ).fetchone()
    return int(row["total"] if row else 0)


def get_customer_rule(rule_id: int) -> dict[str, Any] | None:
    with db_cursor() as conn:
        return _get_customer_rule_from_conn(conn, rule_id)


def _validate_conditions(conditions: list[dict[str, Any]]) -> list[dict[str, str]]:
    allowed_fields = {"胶系", "玻布规格", "PP长度", "树脂含量", "客户规格", "订单备注"}
    result: list[dict[str, str]] = []
    for condition in conditions:
        field = str(condition.get("field") or "").strip()
        operator = str(condition.get("operator") or "contains").strip()
        value = str(condition.get("value") or "").strip()
        if not field or not value:
            continue
        if field not in allowed_fields or operator not in {"equals", "contains", "not_contains"}:
            raise ValueError("客户特殊规则的触发条件不正确")
        result.append({"field": field, "operator": operator, "value": value})
    if not result:
        raise ValueError("请至少填写一个客户特殊规则触发条件")
    return result


def save_customer_rule(data: dict[str, Any], employee_id: str) -> int:
    target_field = str(data.get("target_field") or "").strip()
    if target_field not in CUSTOMER_FIELDS:
        raise ValueError("请选择 PP 客户特殊规则可维护字段")
    conditions = _validate_conditions(data.get("conditions") or [])
    output_value = _clean_output(target_field, data.get("output_value"))
    customer_code = str(data.get("customer_code") or "").strip()
    customer_name = str(data.get("customer_name") or "").strip()
    note = str(data.get("business_note") or "").strip()
    enabled = 1 if data.get("enabled", True) else 0
    rule_id = int(data.get("id") or 0)
    now = utcnow()
    payload = (customer_code, customer_name, target_field, json.dumps(conditions, ensure_ascii=False), output_value, note, enabled)
    with db_cursor() as conn:
        if rule_id:
            before = _get_customer_rule_from_conn(conn, rule_id) or {}
            conn.execute(
                """
                UPDATE pp_transcode_customer_rules
                SET customer_code=?, customer_name=?, target_field=?, conditions_json=?, output_value=?, business_note=?, enabled=?, updated_by=?, updated_at=?
                WHERE id=?
                """,
                (*payload, employee_id, now, rule_id),
            )
            _record_change(conn, "customer", rule_id, "update", before, _get_customer_rule_from_conn(conn, rule_id) or {}, employee_id)
            return rule_id
        cursor = conn.execute(
            """
            INSERT INTO pp_transcode_customer_rules
            (customer_code, customer_name, target_field, conditions_json, output_value, business_note, enabled, created_by, created_at, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (*payload, employee_id, now, employee_id, now),
        )
        new_id = int(cursor.lastrowid)
        _record_change(conn, "customer", new_id, "create", {}, _get_customer_rule_from_conn(conn, new_id) or {}, employee_id)
        return new_id


def set_customer_rule_enabled(rule_id: int, enabled: bool, employee_id: str) -> None:
    with db_cursor() as conn:
        before = _get_customer_rule_from_conn(conn, rule_id) or {}
        conn.execute(
            "UPDATE pp_transcode_customer_rules SET enabled=?, updated_by=?, updated_at=? WHERE id=?",
            (1 if enabled else 0, employee_id, utcnow(), rule_id),
        )
        _record_change(
            conn,
            "customer",
            rule_id,
            "enable" if enabled else "disable",
            before,
            _get_customer_rule_from_conn(conn, rule_id) or {},
            employee_id,
        )


def list_rule_changes(limit: int = 100) -> list[dict[str, Any]]:
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT * FROM pp_transcode_rule_changes ORDER BY id DESC LIMIT ?", (max(1, min(limit, 500)),)
        ).fetchall()
    return [_row_dict(row) for row in rows]


def _backup_payload() -> dict[str, Any]:
    return {
        "created_at": utcnow(),
        "base_rules": list_base_rules(),
        "customer_rules": list_customer_rules(),
    }


def ensure_pp_transcode_daily_backup() -> Path:
    ensure_pp_transcode_tables()
    today = datetime.now().strftime("%Y%m%d")
    target = BACKUP_DIR / f"pp_transcode_rules_{today}.json"
    if not target.exists():
        target.write_text(json.dumps(_backup_payload(), ensure_ascii=False, indent=2), encoding="utf-8")
    cutoff = datetime.now() - timedelta(days=30)
    for path in BACKUP_DIR.glob("pp_transcode_rules_*.json"):
        try:
            if datetime.fromtimestamp(path.stat().st_mtime) < cutoff:
                path.unlink()
        except OSError:
            continue
    return target


def replace_pp_confirmation_items(job_id: int, employee_id: str, items: list[dict[str, Any]]) -> None:
    """Replace one PP job's first-pass results. PP never promotes a formal code here."""
    ensure_pp_transcode_tables()
    now = utcnow()
    with db_cursor() as conn:
        conn.execute("DELETE FROM pp_transcode_confirmation_items WHERE job_id = ?", (job_id,))
        for item in items:
            conn.execute(
                """
                INSERT INTO pp_transcode_confirmation_items
                (job_id, employee_id, excel_row, customer_code, customer_name, spec, order_remark,
                 pending_code, confidence, summary, field_evidence_json, confirmation_status,
                 created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
                """,
                (
                    job_id,
                    employee_id,
                    int(item["excel_row"]),
                    str(item.get("customer_code") or ""),
                    str(item.get("customer_name") or ""),
                    str(item.get("spec") or ""),
                    str(item.get("order_remark") or ""),
                    str(item.get("pending_code") or ""),
                    int(item.get("confidence") or 0),
                    str(item.get("summary") or ""),
                    json.dumps(item.get("field_evidence") or [], ensure_ascii=False),
                    now,
                    now,
                ),
            )


def list_pp_confirmation_items(job_id: int, employee_id: str) -> list[dict[str, Any]]:
    ensure_pp_transcode_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT * FROM pp_transcode_confirmation_items
            WHERE job_id = ? AND employee_id = ?
            ORDER BY excel_row
            """,
            (job_id, employee_id),
        ).fetchall()
    result: list[dict[str, Any]] = []
    for row in rows:
        item = _row_dict(row)
        try:
            item["field_evidence"] = json.loads(item["field_evidence_json"] or "[]")
        except json.JSONDecodeError:
            item["field_evidence"] = []
        result.append(item)
    return result


def update_pp_confirmation_item(
    item_id: int,
    employee_id: str,
    *,
    status: str,
    basis: str = "",
    confirmed_pending_code: str = "",
) -> dict[str, Any] | None:
    if status not in {"confirmed", "skipped"}:
        raise ValueError("PP 确认状态不正确")
    now = utcnow()
    with db_cursor() as conn:
        row = conn.execute(
            "SELECT * FROM pp_transcode_confirmation_items WHERE id = ? AND employee_id = ?",
            (item_id, employee_id),
        ).fetchone()
        if not row:
            return None
        conn.execute(
            """
            UPDATE pp_transcode_confirmation_items
            SET confirmation_status=?, confirmation_basis=?, confirmed_pending_code=?,
                confirmed_by=?, confirmed_at=?, updated_at=?
            WHERE id=?
            """,
            (
                status,
                str(basis or "").strip(),
                str(confirmed_pending_code or "").strip(),
                employee_id,
                now,
                now,
                item_id,
            ),
        )
        updated = conn.execute("SELECT * FROM pp_transcode_confirmation_items WHERE id = ?", (item_id,)).fetchone()
    return _row_dict(updated) if updated else None


def pp_confirmation_counts(job_id: int, employee_id: str) -> dict[str, int]:
    items = list_pp_confirmation_items(job_id, employee_id)
    return {
        "total": len(items),
        "pending": sum(item["confirmation_status"] == "pending" for item in items),
        "confirmed": sum(item["confirmation_status"] == "confirmed" for item in items),
        "skipped": sum(item["confirmation_status"] == "skipped" for item in items),
    }
