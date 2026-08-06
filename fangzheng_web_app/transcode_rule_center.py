from __future__ import annotations

import json
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Mapping

from .db import db_cursor, get_setting, set_setting
from .paths import STORAGE_DIR
from .transcode_agent_glue_resolver import is_retired_agent_glue_mapping


BACKUP_DIR = STORAGE_DIR / "transcode_rule_center_backups"
SCORE_SETTING_KEY = "transcode_rule_center_score_config"
BUSINESS_FIELDS = (
    "胶系",
    "基板厚度",
    "铜箔规格",
    "基板尺寸",
    "胶水类别",
    "铜箔类型+印字/非印字",
    "基板级别",
    "总/芯厚",
)
FIELD_TO_OVERRIDE = {
    "胶系": "glue_code",
    "基板厚度": "thickness_code",
    "铜箔规格": "copper_code",
    "基板尺寸": "size_code",
    "胶水类别": "glue_category_code",
    "铜箔类型+印字/非印字": "copper_type_code",
    "基板级别": "grade_code",
    "总/芯厚": "tc_code",
}
DEFAULT_SCORE_CONFIG = {
    "gate_threshold": 100,
    "semantic_supported_score": 98,
    "model_supported_score": 95,
    "ambiguous_score": 80,
    "missing_evidence_score": 60,
    "contradicted_score": 0,
}
FIELD_VALUE_PATTERNS = {
    "胶系": (re.compile(r"^[A-Z0-9]{2}$"), "胶系标准结果必须是2位字母或数字，例如2B、AH。"),
    "基板厚度": (re.compile(r"^[0-9]{5}$"), "基板厚度标准结果必须是5位数字，例如00800。"),
    "铜箔规格": (re.compile(r"^[A-Z0-9]{2}$"), "铜箔规格标准结果必须是2位字母或数字，例如HH、22。"),
    "基板尺寸": (re.compile(r"^[0-9]{8}$"), "基板尺寸标准结果必须是8位数字，例如37004900。"),
    "胶水类别": (re.compile(r"^[A-Z0-9]$"), "胶水类别标准结果必须是1位字母或数字，例如R、Y。"),
    "铜箔类型+印字/非印字": (
        re.compile(r"^[A-Z0-9]$"),
        "铜箔类型标准结果必须是1位字母或数字，例如W、R。",
    ),
    "基板级别": (re.compile(r"^[A-Z0-9]{2}$"), "基板级别标准结果必须是2位字母或数字，例如A1、AC。"),
    "总/芯厚": (re.compile(r"^[TC]$"), "总/芯厚标准结果只能填写T或C。"),
}
BACKUP_TABLES = (
    "transcode_rule_center_base_overrides",
    "transcode_rule_center_lookup_overrides",
    "transcode_rule_center_confirmation_overrides",
    "transcode_rule_center_changes",
    "transcode_customer_rule_overrides",
    "transcode_agent_rule_overrides",
    "transcode_customer_rule_changes",
    "transcode_rule_center_asset_overrides",
)
LOOKUP_GROUPS = {
    "glue_code": {"label": "胶系代码", "input": "胶系名称", "output": "场内胶系代码"},
    "glue_category": {"label": "胶水类别", "input": "胶系名称", "output": "胶水类别"},
    "grade_code": {"label": "合法基板级别", "input": "基板级别代码", "output": "有效代码"},
    "grade_trigger": {"label": "基板级别写法", "input": "业务等级说明", "output": "基板级别代码"},
    "standard_size": {"label": "毫米尺寸换算", "input": "毫米值", "output": "英寸值"},
    "high_speed_mil": {"label": "高频高速厚度", "input": "mil值", "output": "厚度(mm)"},
    "copper_micron": {"label": "微米铜厚换算", "input": "微米值", "output": "铜厚代码"},
    "copper_type": {"label": "铜箔类型", "input": "铜箔类型写法", "output": "铜箔类型代码"},
    "copper_valid": {"label": "合法铜箔组合", "input": "铜箔组合代码", "output": "有效代码"},
    "size_range": {"label": "标准尺寸区间", "input": "客户尺寸范围", "output": "厂内标准尺寸"},
    "total_to_core": {"label": "总厚转芯厚", "input": "铜箔组合", "output": "减少厚度(mm)"},
    "core_to_total": {"label": "芯厚转总厚", "input": "铜箔组合", "output": "增加厚度(mm)"},
}

# 这些写法虽然保留在历史 transcode_rules 中供旧链路兼容，但条件本身已经
# 限定到具体客户，不属于基础规则页面。客户规则维护入口负责承接这类规则。
CUSTOMER_LIMITED_BASE_LOOKUPS = {
    ("grade_trigger", "深南90022-1"),
    ("grade_trigger", "深南90022-2"),
    ("grade_trigger", "深南测试"),
    ("grade_trigger", "江西测试"),
}
CUSTOMER_LIMITED_BASE_LOOKUP_KEYS = frozenset(
    (group, unicodedata.normalize("NFKC", value).upper())
    for group, value in CUSTOMER_LIMITED_BASE_LOOKUPS
)

BASE_RULE_SCOPE_ORDER = (
    "正式映射表",
    "编码规范",
    "确定性算法",
    "总芯厚转换表",
    "业务补充",
)

BUSINESS_RULE_CATEGORIES = {
    "胶系": {
        "description": "按老版胶系表、最新版胶系表及业务正式补充维护场内胶系代码；同名多码会明确标记待核实。",
        "lookup_groups": ("glue_code",),
        "asset_groups": ("Agent胶系主表", "Agent胶系兼容别名"),
        "scope_order": ("老表", "新表", "额外正式补充"),
    },
    "基板厚度": {
        "description": "维护标准厚度、高频高速厚度写法及厚度换算关系。",
        "lookup_groups": ("high_speed_mil",),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
    "铜箔规格": {
        "description": "维护铜厚写法、铜箔组合及对应的标准结果。",
        "lookup_groups": ("copper_micron", "copper_valid"),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
    "基板尺寸": {
        "description": "维护毫米、英寸和厂内标准尺寸之间的换算关系。客户专属尺寸在对应客户下维护。",
        "lookup_groups": ("standard_size", "size_range"),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
    "胶水类别": {
        "description": "维护胶系与胶水类别之间的对应关系。",
        "lookup_groups": ("glue_category",),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
    "铜箔类型+印字/非印字": {
        "description": "维护铜箔类型及印字、非印字写法对应的标准结果。",
        "lookup_groups": ("copper_type",),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
    "基板级别": {
        "description": "维护业务正式来源中的基板级别编码和确定性判断规则。",
        "lookup_groups": ("grade_code", "grade_trigger"),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
    "总/芯厚": {
        "description": "维护总厚、芯厚判断及铜箔组合对应的厚度换算。",
        "lookup_groups": ("total_to_core", "core_to_total"),
        "asset_groups": (),
        "scope_order": BASE_RULE_SCOPE_ORDER,
    },
}

ASSET_RESULT_FIELDS_BY_CATEGORY = {
    "胶系": ("输出胶系代码", "覆盖胶系代码"),
    "胶水类别": ("覆盖胶水类别",),
    "基板级别": ("覆盖基板级别",),
}
ASSET_GROUPS = {
    "Agent胶系主表": {
        "label": "最新版胶系主表",
        "description": "按最新业务胶系编号和名称维护场内胶系代码。",
        "fields": ("启用", "胶系编号", "胶系名称", "胶系分类", "输出胶系代码", "备注"),
        "field_labels": {"输出胶系代码": "标准胶系代码"},
    },
    "Agent胶系兼容别名": {
        "label": "额外正式补充",
        "description": "维护老表和新表未收录、但已经由业务确认的正式胶系写法。",
        "fields": ("启用", "兼容名称", "标准胶系编号", "标准胶系名称", "输出胶系代码", "备注"),
        "field_labels": {"兼容名称": "业务名称", "输出胶系代码": "标准胶系代码"},
    },
    "Agent胶系选择规则": {
        "label": "胶系多码选择条件",
        "description": "同名胶系存在多个场内代码时，按客户或关键词选择。",
        "fields": ("启用", "胶系名称", "条件客户代码", "条件客户简称", "条件关键词", "输出胶系代码", "优先级", "备注"),
    },
    "Agent基础条件规则": {
        "label": "全客户特殊规则",
        "description": "适用于全部客户、但需要满足胶系或关键词条件的特殊覆盖。",
        "fields": ("启用", "物料类别", "条件胶系", "条件关键词", "关键词模式", "覆盖胶系代码", "覆盖胶水类别", "覆盖基板级别", "规则文本", "备注"),
    },
    "客户尺寸映射": {
        "label": "客户完整尺寸",
        "description": "按客户把一组客户尺寸转换为厂内标准尺寸。",
        "fields": ("客户代码", "客户简称", "客户尺寸W", "客户尺寸H", "厂内尺寸W", "厂内尺寸H", "规则文本", "备注"),
    },
    "客户单边尺寸映射": {
        "label": "客户单边尺寸",
        "description": "按客户替换单边尺寸，再组合生成厂内尺寸。",
        "fields": ("客户代码", "客户简称", "客户单边尺寸", "厂内单边尺寸", "适用字段", "规则文本", "备注"),
    },
    "客户尺寸算法": {
        "label": "客户尺寸算法",
        "description": "维护客户尺寸加大等已确认算法。",
        "fields": ("客户代码", "客户简称", "算法类型", "加大W", "加大H", "适用条件", "规则文本", "备注"),
    },
    "客户厚度映射": {
        "label": "客户厚度写法",
        "description": "把客户专用厚度写法转换为标准厚度和总芯厚口径。",
        "fields": ("客户代码", "客户简称", "客户厚度写法", "厚度mm", "厚度mil", "总芯厚口径", "规则文本", "备注"),
    },
    "客户物料编码口径": {
        "label": "客户物料编码口径",
        "description": "通过客户物料编码判断总厚或芯厚。",
        "fields": ("客户代码", "客户简称", "物料编码模式", "命中值", "总芯厚口径", "规则文本", "备注"),
    },
    "外部尺寸表引用": {
        "label": "外部尺寸表引用",
        "description": "记录仍依赖外部客户尺寸表的规则来源。",
        "fields": ("客户代码", "客户简称", "引用文件", "引用Sheet", "规则文本", "备注"),
    },
    "待接入规则": {
        "label": "待技术接入",
        "description": "业务规则已记录但当前尚未具备执行能力的项目。",
        "fields": ("启用", "客户代码", "客户简称", "技术类型", "原始规则", "规则来源说明", "建议处理", "备注"),
    },
}
CONFIRMATION_FIELDS = {
    "glue": "胶系",
    "thickness": "基板厚度",
    "copper": "铜箔规格",
    "size": "基板尺寸",
    "glue_category": "胶水类别",
    "copper_type": "铜箔类型+印字/非印字",
    "grade": "基板级别",
    "total_core": "总/芯厚",
}


class RuleCenterError(ValueError):
    pass


def ensure_rule_center_tables() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    with db_cursor() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS transcode_rule_center_base_overrides (
                rule_id TEXT PRIMARY KEY,
                rule_json TEXT,
                deleted INTEGER NOT NULL DEFAULT 0,
                updated_by TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS transcode_rule_center_changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT NOT NULL,
                object_id TEXT NOT NULL,
                action TEXT NOT NULL,
                employee_id TEXT NOT NULL,
                before_json TEXT,
                after_json TEXT,
                created_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_transcode_rule_center_changes
            ON transcode_rule_center_changes(category, object_id, id DESC);

            CREATE TABLE IF NOT EXISTS transcode_rule_center_lookup_overrides (
                group_key TEXT NOT NULL,
                input_value TEXT NOT NULL,
                output_value TEXT,
                deleted INTEGER NOT NULL DEFAULT 0,
                updated_by TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(group_key, input_value)
            );

            CREATE TABLE IF NOT EXISTS transcode_rule_center_confirmation_overrides (
                rule_id TEXT PRIMARY KEY,
                rule_json TEXT,
                deleted INTEGER NOT NULL DEFAULT 0,
                updated_by TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS transcode_rule_center_asset_overrides (
                asset_group TEXT NOT NULL,
                row_id TEXT NOT NULL,
                row_json TEXT,
                deleted INTEGER NOT NULL DEFAULT 0,
                updated_by TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(asset_group, row_id)
            );
            """
        )


def load_score_config() -> dict[str, int]:
    config = dict(DEFAULT_SCORE_CONFIG)
    raw = get_setting(SCORE_SETTING_KEY, "") or ""
    if raw:
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            payload = {}
        for key in config:
            if key in payload:
                try:
                    value = int(payload[key])
                except (TypeError, ValueError):
                    continue
                if key == "gate_threshold":
                    continue
                # 防止历史数据库中的非确定性100分绕过页面校验。
                config[key] = min(99, max(0, value))
    # 正式出码门禁是当前已确认安全口径，不允许页面误改为低于100。
    config["gate_threshold"] = 100
    return config


def save_score_config(form: Mapping[str, Any], *, updated_by: str) -> dict[str, int]:
    before = load_score_config()
    config = {"gate_threshold": 100}
    for key in (
        "semantic_supported_score",
        "model_supported_score",
        "ambiguous_score",
        "missing_evidence_score",
        "contradicted_score",
    ):
        try:
            value = int(str(form.get(key, "")).strip())
        except ValueError as exc:
            raise RuleCenterError("评分必须填写0到100之间的整数。") from exc
        if not 0 <= value <= 99:
            raise RuleCenterError("非确定性评分必须填写0到99之间的整数，不能达到正式出码标准。")
        config[key] = value
    set_setting(SCORE_SETTING_KEY, json.dumps(config, ensure_ascii=False, sort_keys=True))
    _record_change("评分标准", "score_config", "修改", updated_by, before, config)
    ensure_daily_backup()
    return config


def merge_score_matrix(matrix: dict[str, Any]) -> dict[str, Any]:
    config = load_score_config()
    merged = dict(matrix)
    merged["gate_threshold"] = config["gate_threshold"]
    merged["semantic_supported_score"] = config["semantic_supported_score"]
    merged["model_supported_score"] = config["model_supported_score"]
    verdict_scores = dict(merged.get("verdict_scores") or {})
    verdict_scores["contradicted"] = {"mode": "fixed", "value": config["contradicted_score"]}
    verdict_scores["ambiguous"] = {"mode": "cap", "value": config["ambiguous_score"]}
    verdict_scores["missing_evidence"] = {
        "mode": "cap",
        "value": config["missing_evidence_score"],
    }
    merged["verdict_scores"] = verdict_scores
    return merged


def list_base_overrides(*, include_deleted: bool = False) -> list[dict[str, Any]]:
    ensure_rule_center_tables()
    sql = "SELECT rule_id, rule_json, deleted, updated_by, updated_at FROM transcode_rule_center_base_overrides"
    if not include_deleted:
        sql += " WHERE deleted = 0"
    sql += " ORDER BY updated_at DESC, rule_id"
    with db_cursor() as conn:
        rows = conn.execute(sql).fetchall()
    result: list[dict[str, Any]] = []
    for row in rows:
        if not row["rule_json"]:
            continue
        item = json.loads(row["rule_json"])
        item["updated_by"] = row["updated_by"]
        item["updated_at"] = row["updated_at"]
        item["deleted"] = bool(row["deleted"])
        result.append(item)
    return result


def merge_base_rule_overrides(base_rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Append global page-maintained rules after loading bundled rule assets."""
    page_rules: list[dict[str, Any]] = []
    for item in list_base_overrides():
        business_field = _clean(item.get("business_field"))
        target_value = _clean(item.get("target_value")).upper()
        native_rule = dict(item.get("native_rule") or {})
        validation = FIELD_VALUE_PATTERNS.get(business_field)
        if (
            not native_rule.get("规则ID")
            or validation is None
            or not validation[0].fullmatch(target_value)
        ):
            continue
        native_rule["覆盖值"] = target_value
        page_rules.append(native_rule)
    return list(base_rules) + page_rules


def merge_lookup_overrides(tables: dict[str, Any]) -> dict[str, Any]:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT group_key, input_value, output_value, deleted FROM transcode_rule_center_lookup_overrides"
        ).fetchall()
    for row in rows:
        group = str(row["group_key"] or "")
        key, value = _normalize_lookup_values(
            group,
            row["input_value"],
            row["output_value"],
        )
        deleted = bool(row["deleted"])
        if group == "glue_code":
            for map_name in ("glue_model_map", "glue_exact_map"):
                target = tables.setdefault(map_name, {})
                if deleted:
                    target.pop(key, None)
                else:
                    target[key] = value.upper()
        elif group == "glue_category":
            target = tables.setdefault("glue_cat_map", {})
            if deleted:
                target.pop(key, None)
            else:
                target[key] = value
        elif group in {"grade_code", "grade_trigger"}:
            target = tables.setdefault("grade_desc_to_code", {})
            if deleted:
                target.pop(key, None)
            else:
                target[key] = value.upper()
                tables.setdefault("grade_code_map", {}).setdefault(value.upper(), key)
        elif group in {"total_to_core", "core_to_total"}:
            target = tables.setdefault(
                "thick_total_to_core" if group == "total_to_core" else "thick_core_to_total", {}
            )
            if deleted:
                target.pop(key, None)
            else:
                target[key] = float(value)
    return tables


def lookup_map_with_overrides(group: str, base_map: Mapping[Any, Any]) -> dict[Any, Any]:
    """Return a copy of a code/standard mapping with page maintenance applied."""
    ensure_rule_center_tables()
    result = dict(base_map)
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT input_value, output_value, deleted FROM transcode_rule_center_lookup_overrides WHERE group_key = ?",
            (group,),
        ).fetchall()
    for row in rows:
        key, value = _normalize_lookup_values(group, row["input_value"], row["output_value"])
        typed_key: Any = key
        typed_value: Any = value
        if group in {"standard_size", "high_speed_mil"}:
            typed_key, typed_value = float(key), float(value)
        if bool(row["deleted"]):
            result.pop(typed_key, None)
        else:
            result[typed_key] = typed_value
    return result


def build_rule_center_lookup_tables(
    engine_tables: dict[str, Any],
    mapping_tables: dict[str, list[dict]],
    *,
    official_grade_codes: set[str],
    standard_sizes: Mapping[float, float],
    high_speed_mil: Mapping[float, float],
    copper_micron: Mapping[str, str],
    copper_types: list[tuple[str, str]],
    copper_valid: set[str],
    size_ranges: list[tuple[float, float, float, float, float, float]],
) -> dict[str, Any]:
    """Build the business-facing mapping view from the same assets used at runtime."""
    tables = dict(engine_tables)
    old_glue: dict[str, str] = {}
    for map_name in ("glue_exact_map", "glue_model_map"):
        for name, code in (engine_tables.get(map_name) or {}).items():
            clean_name = _clean(name).upper()
            clean_code = _clean(code).upper()
            if clean_name and clean_code and not is_retired_agent_glue_mapping(
                {"胶系名称": clean_name, "输出胶系代码": clean_code}
            ):
                old_glue.setdefault(clean_name, clean_code)

    latest_glue: dict[str, str] = {}
    sources: dict[str, dict[str, str]] = {}
    for row in mapping_tables.get("Agent胶系主表", []):
        if _clean(row.get("启用")) == "否":
            continue
        name = _clean(row.get("胶系名称")).upper()
        code = _clean(row.get("输出胶系代码")).upper()
        if name and code:
            latest_glue.setdefault(name, code)
    # 老表与最新版主表在页面上分别展示。最新版主表由 Agent 映射资产承接，
    # 这里仅保留基础 transcode_rules 胶系表，避免同一批新表数据重复出现。
    tables["rule_center_glue_code"] = old_glue
    tables["rule_center_latest_glue_code"] = latest_glue
    sources["glue_code"] = {key: "老表" for key in old_glue}

    official = {str(code).upper(): str(code).upper() for code in official_grade_codes}
    tables["rule_center_grade_code"] = official
    sources["grade_code"] = {key: "官方编码规范" for key in official}
    tables["rule_center_grade_trigger"] = dict(engine_tables.get("grade_desc_to_code") or {})
    sources["grade_trigger"] = {
        str(key): "基础规则表" for key in tables["rule_center_grade_trigger"]
    }
    tables["rule_center_standard_size"] = dict(standard_sizes)
    tables["rule_center_high_speed_mil"] = dict(high_speed_mil)
    tables["rule_center_copper_micron"] = dict(copper_micron)
    tables["rule_center_copper_type"] = dict(copper_types)
    tables["rule_center_copper_valid"] = {
        value: str(value).replace("/", "") for value in copper_valid
    }
    tables["rule_center_size_range"] = {
        f"{w_min:g}-{w_max:g} × {h_min:g}-{h_max:g}": f"{std_w:g} × {std_h:g}"
        for w_min, w_max, h_min, h_max, std_w, std_h in size_ranges
    }
    for group, mapping in (
        ("standard_size", tables["rule_center_standard_size"]),
        ("high_speed_mil", tables["rule_center_high_speed_mil"]),
        ("copper_micron", tables["rule_center_copper_micron"]),
        ("copper_type", tables["rule_center_copper_type"]),
        ("copper_valid", tables["rule_center_copper_valid"]),
        ("size_range", tables["rule_center_size_range"]),
    ):
        sources[group] = {str(key): "官方编码规范" if group in {"standard_size", "high_speed_mil"} else "系统基础规则" for key in mapping}
    tables["__lookup_sources"] = sources
    return tables


def list_lookup_rows(
    tables: dict[str, Any],
    *,
    group_key: str = "glue_code",
    preserve_order: bool = False,
) -> list[dict[str, Any]]:
    group = group_key if group_key in LOOKUP_GROUPS else "glue_code"
    source_name = {
        "glue_code": "rule_center_glue_code",
        "glue_category": "glue_cat_map",
        "grade_code": "rule_center_grade_code",
        "grade_trigger": "rule_center_grade_trigger",
        "standard_size": "rule_center_standard_size",
        "high_speed_mil": "rule_center_high_speed_mil",
        "copper_micron": "rule_center_copper_micron",
        "copper_type": "rule_center_copper_type",
        "copper_valid": "rule_center_copper_valid",
        "size_range": "rule_center_size_range",
        "total_to_core": "thick_total_to_core",
        "core_to_total": "thick_core_to_total",
    }[group]
    base_map = dict(tables.get(source_name) or {})
    if group == "glue_code" and not base_map:
        base_map = dict(tables.get("glue_model_map") or {})
    elif group == "grade_code" and not base_map:
        base_map = {
            str(code): str(code) for code in (tables.get("grade_code_map") or {})
        }
    ensure_rule_center_tables()
    with db_cursor() as conn:
        overrides = conn.execute(
            """
            SELECT input_value, output_value, deleted, updated_by, updated_at
            FROM transcode_rule_center_lookup_overrides WHERE group_key = ?
            """,
            (group,),
        ).fetchall()
    override_map = {str(row["input_value"]): row for row in overrides}
    if preserve_order:
        keys = list(base_map)
        keys.extend(key for key in override_map if key not in base_map)
    else:
        keys = sorted(set(base_map) | set(override_map), key=lambda item: str(item).upper())
    result = []
    for key in keys:
        override = override_map.get(str(key))
        deleted = bool(override and override["deleted"])
        output = override["output_value"] if override and not deleted else base_map.get(key, "")
        baseline_source = str((tables.get("__lookup_sources") or {}).get(group, {}).get(str(key), "正式业务规则"))
        result.append(
            {
                "group_key": group,
                "group_label": LOOKUP_GROUPS[group]["label"],
                "input_value": str(key),
                "output_value": str(output),
                "source": (
                    "页面停用"
                    if deleted
                    else ("页面调整" if override else baseline_source)
                ),
                "has_baseline": key in base_map,
                "deleted": deleted,
                "updated_by": str(override["updated_by"] or "") if override else "",
                "updated_at": str(override["updated_at"] or "") if override else "",
            }
        )
    return result


def summarize_lookup_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Group equivalent lookup outputs for the business-facing read view."""
    grouped: dict[tuple[str, bool], dict[str, Any]] = {}
    for row in rows:
        output = _clean(row.get("output_value")) or "未配置"
        deleted = bool(row.get("deleted"))
        key = (output, deleted)
        item = grouped.setdefault(
            key,
            {
                "output_value": output,
                "deleted": deleted,
                "rows": [],
                "inputs": [],
                "sources": [],
            },
        )
        item["rows"].append(row)
        input_value = _clean(row.get("input_value"))
        source = _clean(row.get("source"))
        if input_value and input_value not in item["inputs"]:
            item["inputs"].append(input_value)
        if source and source not in item["sources"]:
            item["sources"].append(source)

    result = []
    for item in grouped.values():
        inputs = item["inputs"]
        item["count"] = len(item["rows"])
        item["input_preview"] = "、".join(inputs[:4]) + (f" 等{len(inputs)}种写法" if len(inputs) > 4 else "")
        item["source_summary"] = "、".join(item["sources"])
        result.append(item)
    return sorted(result, key=lambda item: (item["deleted"], item["output_value"].upper()))


def save_lookup_override(form: Mapping[str, Any], *, updated_by: str) -> None:
    ensure_rule_center_tables()
    group = _clean(form.get("lookup_group"))
    key, value = _normalize_lookup_values(
        group,
        form.get("lookup_input"),
        form.get("lookup_output"),
    )
    if group not in LOOKUP_GROUPS or not key or not value:
        raise RuleCenterError("映射类型、业务写法和标准结果都不能为空。")
    if group in {"total_to_core", "core_to_total"}:
        try:
            number = float(value)
        except ValueError as exc:
            raise RuleCenterError("厚度换算值必须是毫米数字。") from exc
        if not 0 <= number <= 1:
            raise RuleCenterError("厚度换算值应在0到1毫米之间。")
        value = str(number)
    elif group == "glue_code" and not re.fullmatch(r"[A-Z0-9]{2}", value):
        raise RuleCenterError("场内胶系代码必须是2位字母或数字。")
    elif group == "glue_category" and not re.fullmatch(r"[A-Z0-9]", value):
        raise RuleCenterError("胶水类别必须是1位字母或数字。")
    elif group in {"grade_code", "grade_trigger"} and not re.fullmatch(r"[A-Z0-9]{2}", value):
        raise RuleCenterError("基板级别代码必须是2位字母或数字。")
    elif group in {"standard_size", "high_speed_mil"}:
        try:
            source_number, target_number = float(key), float(value)
        except ValueError as exc:
            raise RuleCenterError("尺寸和厚度换算必须填写数字。") from exc
        if source_number <= 0 or target_number <= 0:
            raise RuleCenterError("尺寸和厚度换算必须大于0。")
        key, value = str(source_number), str(target_number)
    elif group == "copper_micron" and (not re.fullmatch(r"\d+(?:\.\d+)?", key) or not re.fullmatch(r"[A-Z0-9]", value)):
        raise RuleCenterError("微米铜厚必须填写数字，结果必须是1位铜厚代码。")
    elif group == "copper_type" and not re.fullmatch(r"[A-Z0-9]", value):
        raise RuleCenterError("铜箔类型结果必须是1位代码。")
    elif group == "copper_valid" and not re.fullmatch(r"[A-Z0-9]{2,4}", value):
        raise RuleCenterError("合法铜箔组合应填写去掉斜杠后的2到4位代码。")
    elif group == "size_range":
        if not re.fullmatch(r"\s*\d+(?:\.\d+)?-\d+(?:\.\d+)?\s*[×X*]\s*\d+(?:\.\d+)?-\d+(?:\.\d+)?\s*", key):
            raise RuleCenterError("客户尺寸范围格式应为37-37.3 × 49-49.3。")
        if not re.fullmatch(r"\s*\d+(?:\.\d+)?\s*[×X*]\s*\d+(?:\.\d+)?\s*", value):
            raise RuleCenterError("厂内标准尺寸格式应为37.3 × 49.3。")
    before = _lookup_override(group, key)
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_lookup_overrides
                (group_key, input_value, output_value, deleted, updated_by, updated_at)
            VALUES (?, ?, ?, 0, ?, ?)
            ON CONFLICT(group_key, input_value) DO UPDATE SET
                output_value=excluded.output_value, deleted=0,
                updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """,
            (group, key, value, updated_by, now),
        )
    after = {"group_key": group, "input_value": key, "output_value": value}
    _record_change("基础映射", f"{group}:{key}", "新增" if before is None else "修改", updated_by, before, after)
    ensure_daily_backup()


def delete_lookup_override(group: str, key: str, *, updated_by: str) -> None:
    ensure_rule_center_tables()
    group = _clean(group)
    key, _ = _normalize_lookup_values(group, key, "")
    if group not in LOOKUP_GROUPS or not key:
        raise RuleCenterError("要停用的映射不存在。")
    before = _lookup_override(group, key) or {"group_key": group, "input_value": key}
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_lookup_overrides
                (group_key, input_value, output_value, deleted, updated_by, updated_at)
            VALUES (?, ?, NULL, 1, ?, ?)
            ON CONFLICT(group_key, input_value) DO UPDATE SET
                deleted=1, updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """,
            (group, key, updated_by, now),
        )
    _record_change("基础映射", f"{group}:{key}", "停用", updated_by, before, None)
    ensure_daily_backup()


def lookup_group_meta() -> dict[str, dict[str, str]]:
    return {key: dict(value) for key, value in LOOKUP_GROUPS.items()}


def asset_group_meta() -> dict[str, dict[str, Any]]:
    return {key: {**value, "fields": tuple(value["fields"])} for key, value in ASSET_GROUPS.items()}


def business_rule_category_meta() -> dict[str, dict[str, Any]]:
    return {key: dict(value) for key, value in BUSINESS_RULE_CATEGORIES.items()}


def list_business_rule_rows(
    lookup_tables: Mapping[str, Any],
    mapping_tables: Mapping[str, list[dict[str, Any]]],
    *,
    category: str,
) -> list[dict[str, Any]]:
    """Project the runtime rule stores into one business-facing parameter list."""
    selected = category if category in BUSINESS_RULE_CATEGORIES else BUSINESS_FIELDS[0]
    meta = BUSINESS_RULE_CATEGORIES[selected]
    rows: list[dict[str, Any]] = []
    glue_pairs_by_scope: dict[str, set[tuple[str, str]]] = {
        "老表": set(),
        "新表": set(),
        "额外正式补充": set(),
    }
    for group in meta["lookup_groups"]:
        group_meta = LOOKUP_GROUPS[group]
        for row in list_lookup_rows(
            lookup_tables,
            group_key=group,
            preserve_order=selected == "胶系",
        ):
            if _is_customer_limited_base_lookup(group, row["input_value"]):
                continue
            source_scope = ""
            type_label = _lookup_business_type(group)
            if selected == "胶系":
                source_scope = "老表" if row.get("has_baseline") else "额外正式补充"
                type_label = source_scope
                if _is_hidden_glue_mapping(row["input_value"], row["output_value"]):
                    continue
                pair = _glue_display_pair(row["input_value"], row["output_value"])
                if pair in glue_pairs_by_scope[source_scope]:
                    continue
                glue_pairs_by_scope[source_scope].add(pair)
            else:
                source_scope = _lookup_rule_scope(group, row)
                type_label = source_scope
            rows.append(
                {
                    "kind": "lookup",
                    "category": selected,
                    "type_label": type_label,
                    "source_scope": source_scope,
                    "title": row["input_value"],
                    "detail": group_meta["label"],
                    "result": row["output_value"],
                    "enabled": not row["deleted"],
                    "lookup_group": group,
                    "lookup_input": row["input_value"],
                    "lookup_output": row["output_value"],
                    "conflict": False,
                    "eligible_for_formal_score": not row["deleted"],
                    "status_label": "生效中" if not row["deleted"] else "已停用",
                }
            )
    for group in meta["asset_groups"]:
        result_fields = ASSET_RESULT_FIELDS_BY_CATEGORY.get(selected, ())
        for row in list_asset_rows(mapping_tables, asset_group=group):
            result_field = next((field for field in result_fields if _clean(row.get(field))), "")
            if result_fields and not result_field:
                continue
            source_scope = ""
            if selected == "胶系":
                if group == "Agent胶系兼容别名":
                    source_scope = "额外正式补充"
                elif group == "Agent胶系主表" and not str(row.get("_row_id") or "").startswith("GLUE-PAGE-"):
                    source_scope = "新表"
                else:
                    source_scope = "额外正式补充"
                candidate_result = _clean(row.get(result_field)) if result_field else "-"
                candidate_title = _asset_business_title(row)
                if _is_hidden_glue_mapping(candidate_title, candidate_result):
                    continue
                pair = _glue_display_pair(candidate_title, candidate_result)
                if group == "Agent胶系兼容别名" and (
                    pair in glue_pairs_by_scope["老表"]
                    or pair in glue_pairs_by_scope["新表"]
                ):
                    continue
                if pair in glue_pairs_by_scope[source_scope]:
                    continue
                glue_pairs_by_scope[source_scope].add(pair)
            rows.append(
                {
                    "kind": "asset",
                    "category": selected,
                    "type_label": _asset_business_type(group),
                    "source_scope": source_scope,
                    "title": candidate_title if selected == "胶系" else _asset_business_title(row),
                    "detail": _asset_business_detail(row),
                    "result": _clean(row.get(result_field)) if result_field else "-",
                    "enabled": _clean(row.get("启用")) != "否",
                    "asset_group": group,
                    "asset_row_id": row.get("_row_id") or row.get("映射ID"),
                    "conflict": False,
                    "eligible_for_formal_score": _clean(row.get("启用")) != "否",
                    "status_label": "生效中" if _clean(row.get("启用")) != "否" else "已停用",
                }
            )
    if selected == "胶系":
        _mark_glue_conflicts(rows)
        source_rank = {"老表": 0, "新表": 1, "额外正式补充": 2}
        rows.sort(key=lambda item: source_rank.get(item.get("source_scope", ""), 3))
        previous_scope = ""
        for row in rows:
            scope = str(row.get("source_scope") or "")
            row["group_start"] = bool(scope and scope != previous_scope)
            previous_scope = scope
    else:
        rows = _dedupe_business_rule_rows(rows)
        rows.sort(key=lambda item: (not item["enabled"], item["type_label"], item["title"], item["result"]))
    return rows


def _lookup_business_type(group: str) -> str:
    algorithm_labels = {
        "total_to_core": "总厚转芯厚",
        "core_to_total": "芯厚转总厚",
        "standard_size": "毫米尺寸换算",
        "high_speed_mil": "高频高速厚度换算",
        "copper_micron": "微米铜厚换算",
        "size_range": "标准尺寸区间",
    }
    if group in algorithm_labels:
        return algorithm_labels[group]
    return "正式业务映射"


def _lookup_rule_scope(group: str, row: Mapping[str, Any]) -> str:
    """Return a business-readable scope without exposing storage implementation names."""
    if not bool(row.get("has_baseline")):
        return "业务补充"
    if group in {"total_to_core", "core_to_total"}:
        return "总芯厚转换表"
    if group in {"standard_size", "high_speed_mil", "grade_code"}:
        return "编码规范"
    if group in {"size_range"}:
        return "确定性算法"
    return "正式映射表"


def _is_customer_limited_base_lookup(group: str, value: Any) -> bool:
    normalized = unicodedata.normalize("NFKC", _clean(value)).upper()
    return (group, normalized) in CUSTOMER_LIMITED_BASE_LOOKUP_KEYS


def _asset_business_type(group: str) -> str:
    return {
        "Agent胶系主表": "新表",
        "Agent胶系兼容别名": "额外正式补充",
        "Agent胶系选择规则": "客户特殊规则",
        "Agent基础条件规则": "全客户特殊规则",
    }.get(group, "正式业务规则")


def _normalize_glue_identity(value: Any) -> str:
    # 保留中文用途限定，避免把“NY1600”和“NY1600降本”等不同正式胶系误判为同名冲突。
    normalized = unicodedata.normalize("NFKC", _clean(value)).upper()
    return re.sub(r"[\s\-_/]+", "", normalized)


def _glue_display_pair(name: Any, code: Any) -> tuple[str, str]:
    return _normalize_glue_identity(name), _clean(code).upper()


def _dedupe_business_rule_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Collapse duplicate projections while preserving distinct conversion directions."""
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str, bool]] = set()
    for row in rows:
        direction = _clean(row.get("lookup_group"))
        if direction not in {"total_to_core", "core_to_total"}:
            direction = ""
        key = (
            _clean(row.get("category")),
            direction,
            re.sub(r"\s+", "", _clean(row.get("title")).upper()),
            re.sub(r"\s+", "", _clean(row.get("result")).upper()),
            bool(row.get("enabled")),
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def _is_hidden_glue_mapping(name: Any, code: Any) -> bool:
    """NY-A1 -> 2Z is retired and must not be exposed as a maintainable rule."""
    return _normalize_glue_identity(name) == "NYA1" and _clean(code).upper() == "2Z"


def _mark_glue_conflicts(rows: list[dict[str, Any]]) -> None:
    codes_by_name: dict[str, list[str]] = {}
    for row in rows:
        if not row.get("enabled"):
            continue
        name = _normalize_glue_identity(row.get("title"))
        code = _clean(row.get("result")).upper()
        if not name or not code or code == "-":
            continue
        codes = codes_by_name.setdefault(name, [])
        if code not in codes:
            codes.append(code)
    for row in rows:
        codes = codes_by_name.get(_normalize_glue_identity(row.get("title")), [])
        row["conflict"] = bool(row.get("enabled")) and len(codes) > 1
        row["conflict_codes"] = tuple(codes)
        row["eligible_for_formal_score"] = bool(row.get("enabled")) and not row["conflict"]
        row["status_label"] = "冲突待核实" if row["conflict"] else ("生效中" if row.get("enabled") else "已停用")
        if row["conflict"]:
            detail = _clean(row.get("detail"))
            warning = f"同名胶系存在多个正式代码：{' / '.join(codes)}"
            row["detail"] = f"{detail}；{warning}" if detail and detail != "-" else warning


def _asset_business_title(row: Mapping[str, Any]) -> str:
    return next(
        (
            _clean(row.get(field))
            for field in ("胶系名称", "兼容名称", "条件胶系", "规则文本", "映射ID")
            if _clean(row.get(field))
        ),
        "未命名规则",
    )


def _asset_business_detail(row: Mapping[str, Any]) -> str:
    return next(
        (
            _clean(row.get(field))
            for field in ("规则文本", "条件关键词", "备注", "胶系分类")
            if _clean(row.get(field))
        ),
        "-",
    )


def _base_condition_summary(conditions: Mapping[str, Any]) -> str:
    parts = [f"{key}={value}" for key, value in conditions.items() if _clean(value)]
    return "；".join(parts) or "适用于全部规格"


def merge_agent_mapping_overrides(
    tables: dict[str, list[dict]],
) -> dict[str, list[dict]]:
    """Overlay page-maintained auxiliary assets without changing released workbooks."""
    ensure_rule_center_tables()
    merged = {name: [dict(row) for row in rows] for name, rows in tables.items()}
    with db_cursor() as conn:
        overrides = conn.execute(
            "SELECT asset_group, row_id, row_json, deleted FROM transcode_rule_center_asset_overrides"
        ).fetchall()
        glue_overrides = conn.execute(
            "SELECT input_value, output_value, deleted FROM transcode_rule_center_lookup_overrides WHERE group_key='glue_code'"
        ).fetchall()
    for override in overrides:
        group = str(override["asset_group"] or "")
        if group not in ASSET_GROUPS:
            continue
        row_id = str(override["row_id"] or "")
        rows = merged.setdefault(group, [])
        index = next((idx for idx, row in enumerate(rows) if _clean(row.get("映射ID")) == row_id), None)
        if bool(override["deleted"]):
            if index is not None:
                rows[index]["启用"] = "否"
            continue
        payload = json.loads(override["row_json"] or "{}")
        payload["映射ID"] = row_id
        if index is None:
            rows.append(payload)
        else:
            rows[index] = payload

    # The simple glue mapping editor is an exact-name shortcut over the latest master.
    master = merged.setdefault("Agent胶系主表", [])
    for override in glue_overrides:
        name = _clean(override["input_value"]).upper()
        matched = [row for row in master if _clean(row.get("胶系名称")).upper() == name]
        if bool(override["deleted"]):
            for row in matched:
                row["启用"] = "否"
            continue
        code = _clean(override["output_value"]).upper()
        if matched:
            for row in matched:
                row["输出胶系代码"] = code
                row["启用"] = "是"
        elif name and code:
            master.append(
                {
                    "映射ID": f"GLUE-PAGE-{name}",
                    "启用": "是",
                    "胶系编号": "",
                    "胶系名称": name,
                    "胶系分类": "页面新增",
                    "输出胶系代码": code,
                    "来源文件": "统一规则配置页面",
                    "来源行号": "",
                    "备注": "页面维护",
                }
            )
    for group in ("Agent胶系主表", "Agent胶系兼容别名", "Agent胶系选择规则"):
        merged[group] = [
            row for row in merged.get(group, [])
            if not is_retired_agent_glue_mapping(row)
        ]
    return merged


def list_asset_rows(
    mapping_tables: dict[str, list[dict]], *, asset_group: str
) -> list[dict[str, Any]]:
    group = asset_group if asset_group in ASSET_GROUPS else next(iter(ASSET_GROUPS))
    ensure_rule_center_tables()
    with db_cursor() as conn:
        override_rows = conn.execute(
            "SELECT row_id, deleted, updated_by, updated_at FROM transcode_rule_center_asset_overrides WHERE asset_group=?",
            (group,),
        ).fetchall()
    override_map = {str(row["row_id"]): row for row in override_rows}
    result = []
    for row in mapping_tables.get(group, []):
        item = dict(row)
        row_id = _clean(item.get("映射ID"))
        override = override_map.get(row_id)
        item["_row_id"] = row_id
        item["_source"] = "页面调整" if override and not bool(override["deleted"]) else "活动Agent资产"
        item["_updated_by"] = str(override["updated_by"] or "") if override else ""
        item["_updated_at"] = str(override["updated_at"] or "") if override else ""
        result.append(item)
    return result


def summarize_asset_rows(
    rows: list[dict[str, Any]], *, asset_group: str
) -> list[dict[str, Any]]:
    """Build a customer/code archive without changing the executable asset rows."""

    def first(row: dict[str, Any], *keys: str) -> str:
        return next((_clean(row.get(key)) for key in keys if _clean(row.get(key))), "")

    customer_groups = {
        "客户尺寸映射",
        "客户单边尺寸映射",
        "客户尺寸算法",
        "客户厚度映射",
        "客户物料编码口径",
        "外部尺寸表引用",
        "待接入规则",
    }
    grouped: dict[str, dict[str, Any]] = {}
    for row in rows:
        if asset_group in customer_groups:
            label = first(row, "客户简称", "客户代码") or "未指定客户"
            sublabel = first(row, "客户代码")
        elif asset_group == "Agent胶系主表":
            label = first(row, "输出胶系代码") or "未配置代码"
            sublabel = "场内胶系代码"
        elif asset_group == "Agent胶系兼容别名":
            label = first(row, "标准胶系名称", "标准胶系编号", "输出胶系代码") or "未指定标准胶系"
            sublabel = first(row, "输出胶系代码")
        else:
            label = first(row, "胶系名称", "条件胶系", "技术类型", "映射ID") or "其他规则"
            sublabel = first(row, "条件客户简称", "条件客户代码")

        title = first(row, "客户简称", "胶系名称", "兼容名称", "技术类型", "映射ID") or "未命名规则"
        size_detail = " × ".join(
            value for value in (first(row, "客户尺寸W"), first(row, "客户尺寸H")) if value
        )
        detail = first(
            row,
            "规则文本",
            "原始规则",
            "客户厚度写法",
            "条件关键词",
            "适用条件",
            "胶系分类",
        ) or size_detail or first(row, "备注") or "-"
        result_value = first(
            row,
            "输出胶系代码",
            "覆盖基板级别",
            "目标size_code",
            "厂内单边尺寸",
            "总芯厚口径",
            "算法类型",
        ) or "-"
        if (
            asset_group == "客户尺寸映射"
            and result_value == "-"
            and first(row, "厂内尺寸W")
            and first(row, "厂内尺寸H")
        ):
            result_value = f"{first(row, '厂内尺寸W')}*{first(row, '厂内尺寸H')}"
        key = f"{label}\x1f{sublabel}"
        item = grouped.setdefault(
            key,
            {"label": label, "sublabel": sublabel, "rows": [], "results": [], "sources": [], "enabled_count": 0},
        )
        item["rows"].append({"native": row, "title": title, "detail": detail, "result": result_value})
        if result_value not in item["results"]:
            item["results"].append(result_value)
        source = first(row, "_source")
        if source and source not in item["sources"]:
            item["sources"].append(source)
        if first(row, "启用") != "否":
            item["enabled_count"] += 1

    result = []
    for item in grouped.values():
        item["count"] = len(item["rows"])
        item["result_summary"] = "、".join(item["results"][:4]) + (
            f" 等{len(item['results'])}种结果" if len(item["results"]) > 4 else ""
        )
        item["source_summary"] = "、".join(item["sources"])
        result.append(item)
    return sorted(result, key=lambda item: (item["label"].upper(), item["sublabel"].upper()))


def summarize_semantic_rules(rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Group customer semantic rules by their confirmed business outcome."""
    business_value_labels = {
        "core": "芯厚",
        "total": "总厚",
        "core_after_total_to_core_conversion": "按总芯厚转换表转为芯厚",
        "total_after_core_to_total_conversion": "按总芯厚转换表转为总厚",
        "printed": "有印字/水印",
        "unprinted": "无印字/水印",
        "present": "有",
        "absent": "无",
    }
    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    for rule in rules:
        field = _clean(rule.get("business_field")) or "未分类参数"
        normalized = [
            _clean(value) for value in (rule.get("normalized_values") or []) if _clean(value)
        ]
        target = "、".join(business_value_labels.get(value, value) for value in normalized) or "仅语义识别"
        key = (field, target)
        item = grouped.setdefault(
            key,
            {"business_field": field, "target": target, "rules": [], "customers": [], "phrases": [], "enabled_count": 0},
        )
        item["rules"].append(rule)
        customer = _clean(rule.get("customer_name")) or _clean(rule.get("customer_code")) or "通用"
        if customer not in item["customers"]:
            item["customers"].append(customer)
        phrase = _clean(rule.get("source_text"))
        if phrase and phrase not in item["phrases"]:
            item["phrases"].append(phrase)
        if bool(rule.get("enabled", True)):
            item["enabled_count"] += 1

    result = []
    for item in grouped.values():
        item["count"] = len(item["rules"])
        item["customer_count"] = len(item["customers"])
        item["customer_preview"] = "、".join(item["customers"][:4]) + (
            f" 等{len(item['customers'])}个客户" if len(item["customers"]) > 4 else ""
        )
        item["phrase_preview"] = "；".join(item["phrases"][:2]) + (
            f" 等{len(item['phrases'])}种说法" if len(item["phrases"]) > 2 else ""
        )
        result.append(item)
    return sorted(result, key=lambda item: (item["business_field"], item["target"]))


def find_asset_row(
    mapping_tables: dict[str, list[dict]], asset_group: str, row_id: str
) -> dict[str, Any] | None:
    target = _clean(row_id)
    return next(
        (row for row in list_asset_rows(mapping_tables, asset_group=asset_group) if row.get("_row_id") == target),
        None,
    )


def build_asset_row_from_form(
    form: Mapping[str, Any], *, existing: dict[str, Any] | None = None
) -> tuple[str, dict[str, Any]]:
    group = _clean(form.get("asset_group"))
    if group not in ASSET_GROUPS:
        raise RuleCenterError("请选择要维护的规则分类。")
    prior = {key: value for key, value in (existing or {}).items() if not str(key).startswith("_")}
    row_id = _clean(form.get("asset_row_id")) or datetime.now().strftime("ASSET-PAGE-%Y%m%d-%H%M%S-%f")
    payload = dict(prior)
    payload["映射ID"] = row_id
    if "启用" not in payload:
        payload["启用"] = "是"
    for field in ASSET_GROUPS[group]["fields"]:
        if field == "启用":
            payload[field] = "是" if _form_bool(form.get("asset__启用"), default=False) else "否"
        else:
            payload[field] = _clean(form.get(f"asset__{field}"))
    required_fields = {
        "Agent胶系主表": ("胶系名称", "输出胶系代码"),
        "Agent胶系兼容别名": ("兼容名称", "输出胶系代码"),
        "Agent胶系选择规则": ("胶系名称", "输出胶系代码"),
        "客户尺寸映射": ("客户简称", "客户尺寸W", "客户尺寸H", "厂内尺寸W", "厂内尺寸H"),
        "客户单边尺寸映射": ("客户简称", "客户单边尺寸", "厂内单边尺寸"),
        "客户尺寸算法": ("客户简称", "算法类型"),
        "客户厚度映射": ("客户简称", "客户厚度写法"),
    }.get(group, ())
    missing = [field for field in required_fields if not _clean(payload.get(field))]
    if missing:
        raise RuleCenterError(f"请填写：{'、'.join(missing)}。")
    for field in ("输出胶系代码", "覆盖胶系代码"):
        value = _clean(payload.get(field)).upper()
        if value and not re.fullmatch(r"[A-Z0-9]{2}", value):
            raise RuleCenterError(f"{field}必须是2位字母或数字。")
        if value:
            payload[field] = value
    return group, payload


def save_asset_override(
    group: str,
    row: dict[str, Any],
    *,
    updated_by: str,
    record_change: bool = True,
) -> None:
    ensure_rule_center_tables()
    row_id = _clean(row.get("映射ID"))
    before = _asset_override(group, row_id)
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_asset_overrides
                (asset_group, row_id, row_json, deleted, updated_by, updated_at)
            VALUES (?, ?, ?, 0, ?, ?)
            ON CONFLICT(asset_group, row_id) DO UPDATE SET
                row_json=excluded.row_json, deleted=0,
                updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """,
            (group, row_id, _dump(row), updated_by, now),
        )
    if record_change:
        _record_change(
            "规则资产",
            f"{group}:{row_id}",
            "新增" if before is None else "修改",
            updated_by,
            before,
            row,
        )
    ensure_daily_backup()


def delete_asset_override(group: str, row_id: str, *, updated_by: str) -> None:
    ensure_rule_center_tables()
    if group not in ASSET_GROUPS or not _clean(row_id):
        raise RuleCenterError("要停用的规则资产不存在。")
    before = _asset_override(group, row_id)
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_asset_overrides
                (asset_group, row_id, row_json, deleted, updated_by, updated_at)
            VALUES (?, ?, NULL, 1, ?, ?)
            ON CONFLICT(asset_group, row_id) DO UPDATE SET
                row_json=NULL, deleted=1, updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """,
            (group, row_id, updated_by, now),
        )
    _record_change("规则资产", f"{group}:{row_id}", "停用", updated_by, before, None)
    ensure_daily_backup()


def merge_confirmation_policy_overrides(base_rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            "SELECT rule_id, rule_json, deleted FROM transcode_rule_center_confirmation_overrides"
        ).fetchall()
    overrides = {str(row["rule_id"]): row for row in rows}
    base_ids = {str(rule.get("rule_id") or "") for rule in base_rules}
    merged = []
    for rule in base_rules:
        rule_id = str(rule.get("rule_id") or "")
        override = overrides.get(rule_id)
        if override is None:
            merged.append(rule)
        elif not int(override["deleted"] or 0):
            merged.append(json.loads(override["rule_json"]))
    for rule_id, row in overrides.items():
        if rule_id not in base_ids and not int(row["deleted"] or 0):
            merged.append(json.loads(row["rule_json"]))
    return merged


def list_confirmation_policy_views(
    base_rules: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    if base_rules is None:
        from .transcode_confirmation_policy import DEFAULT_CONFIRMATION_RULE_PATH

        payload = json.loads(DEFAULT_CONFIRMATION_RULE_PATH.read_text(encoding="utf-8"))
        base_rules = payload.get("rules") or []
    result = []
    for rule in merge_confirmation_policy_overrides(base_rules):
        groups = ["/".join(str(value) for value in group) for group in rule.get("contains_any_groups") or []]
        field_key = str((rule.get("field_keys") or [""])[0])
        result.append(
            {
                "rule_id": str(rule.get("rule_id") or ""),
                "customers": "，".join(str(value) for value in rule.get("customers") or []),
                "contains_all": "，".join(str(value) for value in rule.get("contains_all") or []),
                "contains_any_groups": "；".join(groups),
                "not_contains_any": "，".join(str(value) for value in rule.get("not_contains_any") or []),
                "field_key": field_key,
                "field_label": CONFIRMATION_FIELDS.get(field_key, str(rule.get("field") or "")),
                "reason": str(rule.get("reason") or ""),
                "business_basis": str(rule.get("business_basis") or ""),
                "native_rule": rule,
            }
        )
    return result


def build_confirmation_policy_from_form(
    form: Mapping[str, Any], *, existing: dict[str, Any] | None = None
) -> dict[str, Any]:
    field_key = _clean(form.get("confirmation_field"))
    if field_key not in CONFIRMATION_FIELDS:
        raise RuleCenterError("请选择需要人工确认的参数。")
    customers = _split_business_values(form.get("confirmation_customers"))
    if not customers:
        raise RuleCenterError("人工确认规则至少需要一个客户，避免影响全部客户。")
    contains_all = _split_business_values(form.get("confirmation_contains_all"))
    group_texts = [item.strip() for item in str(form.get("confirmation_any_groups") or "").split("；") if item.strip()]
    any_groups = [_split_business_values(text, separators="/|,，、") for text in group_texts]
    any_groups = [group for group in any_groups if group]
    excluded = _split_business_values(form.get("confirmation_excluded"))
    if not (contains_all or any_groups or excluded):
        raise RuleCenterError("至少填写一个触发关键词或排除关键词。")
    reason = _clean(form.get("confirmation_reason"))
    if not reason:
        raise RuleCenterError("请填写业务需要人工确认的原因。")
    prior = existing or {}
    rule_id = _clean(form.get("confirmation_rule_id")) or datetime.now().strftime("CPR-PAGE-%Y%m%d-%H%M%S-%f")
    return {
        **prior,
        "rule_id": rule_id,
        "status": "approved" if _form_bool(form.get("confirmation_enabled"), default=True) else "disabled",
        "basis_type": prior.get("basis_type") or "condition_missing",
        "customers": customers,
        "contains_all": contains_all,
        "contains_any_groups": any_groups,
        "not_contains_any": excluded,
        "field": CONFIRMATION_FIELDS[field_key],
        "field_keys": [field_key],
        "reason": reason,
        "business_basis": _clean(form.get("confirmation_basis")) or reason,
        "created_at": prior.get("created_at") or _now(),
    }


def save_confirmation_policy(rule: dict[str, Any], *, updated_by: str) -> None:
    ensure_rule_center_tables()
    rule_id = _clean(rule.get("rule_id"))
    before = _confirmation_override(rule_id)
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_confirmation_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 0, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json=excluded.rule_json, deleted=0,
                updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """,
            (rule_id, _dump(rule), updated_by, now),
        )
    _record_change("人工确认规则", rule_id, "新增" if before is None else "修改", updated_by, before, rule)
    _clear_confirmation_cache()
    ensure_daily_backup()


def delete_confirmation_policy(rule_id: str, *, updated_by: str) -> None:
    ensure_rule_center_tables()
    before = _confirmation_override(rule_id)
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_confirmation_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, NULL, 1, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json=NULL, deleted=1, updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """,
            (rule_id, updated_by, now),
        )
    _record_change("人工确认规则", rule_id, "停用", updated_by, before, None)
    _clear_confirmation_cache()
    ensure_daily_backup()


def confirmation_field_meta() -> dict[str, str]:
    return dict(CONFIRMATION_FIELDS)


def build_base_rule_from_form(
    form: Mapping[str, Any], *, existing: dict[str, Any] | None = None
) -> dict[str, Any]:
    business_field = _clean(form.get("business_field"))
    if business_field not in FIELD_TO_OVERRIDE:
        raise RuleCenterError("请选择需要维护的业务参数。")
    target_value = _clean(form.get("target_value")).upper()
    if not target_value:
        raise RuleCenterError("请填写规则生效后的标准结果。")
    if len(target_value) > 8:
        raise RuleCenterError("标准结果长度不符合编码字段要求。")
    pattern, message = FIELD_VALUE_PATTERNS[business_field]
    if not pattern.fullmatch(target_value):
        raise RuleCenterError(message)
    source_text = _clean(form.get("source_text"))
    if not source_text:
        raise RuleCenterError("请用业务语言填写规则说明。")
    conditions = {
        "条件胶系": _clean(form.get("condition_glue")),
        "条件关键词": _clean(form.get("condition_keyword")),
        "条件铜厚": _clean(form.get("condition_copper")),
        "条件厚度": _clean(form.get("condition_thickness")),
        "条件尺寸": _clean(form.get("condition_size")),
    }
    if not any(conditions.values()):
        raise RuleCenterError("至少填写一个适用条件，防止规则覆盖全部规格。")
    prior = existing or {}
    rule_id = _clean(form.get("rule_id")) or _new_rule_id()
    try:
        priority = int(str(form.get("priority") or "100"))
    except ValueError as exc:
        raise RuleCenterError("优先顺序必须是整数。") from exc
    native = {
        "_global_rule": True,
        "规则ID": rule_id,
        "启用": "是" if _form_bool(form.get("enabled"), default=True) else "否",
        "客户代码": "",
        "客户简称": "",
        "物料类别": "CCL",
        "来源字段": "统一规则配置页面",
        "原始字段": business_field,
        "规则文本": source_text,
        "条件文本": source_text,
        **conditions,
        "覆盖字段": FIELD_TO_OVERRIDE[business_field],
        "覆盖值": target_value,
        "命中来源": "页面维护基础规则",
        "优先级": str(priority),
        "强制执行": "是",
        "待确认": "否",
        "来源行号": "",
        "规则解释": _clean(form.get("business_basis")) or source_text,
        "跳过原因": "",
    }
    return {
        "rule_id": rule_id,
        "business_field": business_field,
        "source_text": source_text,
        "target_value": target_value,
        "priority": priority,
        "enabled": native["启用"] == "是",
        "conditions": {
            "胶系": conditions["条件胶系"],
            "规格关键词": conditions["条件关键词"],
            "铜箔规格": conditions["条件铜厚"],
            "基板厚度": conditions["条件厚度"],
            "基板尺寸": conditions["条件尺寸"],
        },
        "business_basis": native["规则解释"],
        "native_rule": native,
        "created_at": prior.get("created_at") or _now(),
    }


def save_base_override(rule: dict[str, Any], *, updated_by: str) -> None:
    ensure_rule_center_tables()
    rule_id = _clean(rule.get("rule_id"))
    before = find_base_override(rule_id)
    now = _now()
    payload = dict(rule)
    payload["updated_at"] = now
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_base_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 0, ?, ?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_json=excluded.rule_json,
                deleted=0,
                updated_by=excluded.updated_by,
                updated_at=excluded.updated_at
            """,
            (rule_id, _dump(payload), updated_by, now),
        )
    _record_change("基础规则", rule_id, "新增" if before is None else "修改", updated_by, before, payload)
    ensure_daily_backup()


def delete_base_override(rule_id: str, *, updated_by: str) -> None:
    ensure_rule_center_tables()
    before = find_base_override(rule_id)
    if before is None:
        raise RuleCenterError("要删除的基础规则不存在。")
    now = _now()
    with db_cursor() as conn:
        conn.execute(
            "UPDATE transcode_rule_center_base_overrides SET deleted=1, updated_by=?, updated_at=? WHERE rule_id=?",
            (updated_by, now, rule_id),
        )
    _record_change("基础规则", rule_id, "删除", updated_by, before, None)
    ensure_daily_backup()


def find_base_override(rule_id: str) -> dict[str, Any] | None:
    target = _clean(rule_id)
    return next((item for item in list_base_overrides() if item.get("rule_id") == target), None)


def list_rule_center_changes(limit: int = 50) -> list[dict[str, Any]]:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT id, category, object_id, action, employee_id, before_json, after_json, created_at
            FROM transcode_rule_center_changes
            WHERE employee_id IS NULL OR employee_id <> 'system-cleanup'
            ORDER BY id DESC LIMIT ?
            """,
            (max(1, min(int(limit), 200)),),
        ).fetchall()
    return [dict(row) for row in rows]


def rule_center_summary(
    *, mapping_tables: dict[str, list[dict]], agent_rules: list[dict], semantic_rules: list[dict]
) -> dict[str, Any]:
    order_semantic = [
        rule
        for rule in semantic_rules
        if any(str(item.get("field") or "") == "订单备注" for item in rule.get("conditions") or [])
    ]
    customer_keys = {
        (_clean(rule.get("客户代码")), _clean(rule.get("客户简称")))
        for rule in agent_rules
        if _clean(rule.get("客户代码")) or _clean(rule.get("客户简称"))
    }
    customer_keys.update(
        (_clean(rule.get("customer_code")), _clean(rule.get("customer_name")))
        for rule in semantic_rules
        if _clean(rule.get("customer_code")) or _clean(rule.get("customer_name"))
    )
    mapping_counts = {
        name: sum(1 for row in rows if _clean(row.get("启用")) != "否")
        for name, rows in mapping_tables.items()
    }
    return {
        "base_override_count": len(list_base_overrides()),
        # 总览中的基础规则数量由业务参数投影单独计算。这里不再把客户映射表
        # 和语义资产混入“基础规则”统计。
        "base_mapping_count": 0,
        "mapping_counts": mapping_counts,
        "customer_count": len(customer_keys),
        "agent_rule_count": len(agent_rules),
        "semantic_rule_count": len(semantic_rules),
        "order_semantic_count": len(order_semantic),
        "score_config": load_score_config(),
    }


def ensure_daily_backup() -> Path:
    ensure_rule_center_tables()
    today = date.today().isoformat()
    path = BACKUP_DIR / f"transcode-rules-{today}.json"
    if not path.exists():
        create_backup(path=path, reason="每日自动备份")
    _cleanup_backups()
    return path


def create_backup(*, path: Path | None = None, reason: str = "手动备份") -> Path:
    ensure_rule_center_tables()
    from .transcode_customer_rule_admin import ensure_customer_rule_maintenance_tables

    ensure_customer_rule_maintenance_tables()
    target = path or BACKUP_DIR / f"transcode-rules-{datetime.now().strftime('%Y-%m-%d-%H%M%S-%f')}.json"
    payload: dict[str, Any] = {
        "schema_version": "1.0",
        "created_at": _now(),
        "reason": reason,
        "settings": {SCORE_SETTING_KEY: get_setting(SCORE_SETTING_KEY, "") or ""},
        "tables": {},
    }
    with db_cursor() as conn:
        existing = {
            row["name"]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }
        for table in BACKUP_TABLES:
            if table not in existing:
                payload["tables"][table] = []
                continue
            payload["tables"][table] = [dict(row) for row in conn.execute(f"SELECT * FROM {table}").fetchall()]
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    _cleanup_backups()
    return target


def list_backups() -> list[dict[str, Any]]:
    ensure_rule_center_tables()
    result = []
    for path in sorted(BACKUP_DIR.glob("transcode-rules-*.json"), reverse=True):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        result.append(
            {
                "name": path.name,
                "created_at": payload.get("created_at", ""),
                "reason": payload.get("reason", ""),
                "size_kb": max(1, round(path.stat().st_size / 1024)),
            }
        )
    return result


def restore_backup(name: str, *, updated_by: str) -> None:
    ensure_rule_center_tables()
    from .transcode_customer_rule_admin import ensure_customer_rule_maintenance_tables

    ensure_customer_rule_maintenance_tables()
    safe_name = Path(name).name
    path = BACKUP_DIR / safe_name
    if not path.exists() or path.parent != BACKUP_DIR:
        raise RuleCenterError("备份文件不存在。")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0":
        raise RuleCenterError("备份格式不受支持。")
    create_backup(reason=f"恢复{name}前自动备份")
    tables = payload.get("tables") or {}
    with db_cursor() as conn:
        for table in BACKUP_TABLES:
            if table not in tables:
                continue
            columns = [row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]
            if not columns:
                continue
            conn.execute(f"DELETE FROM {table}")
            placeholders = ",".join("?" for _ in columns)
            for row in tables[table]:
                conn.execute(
                    f"INSERT INTO {table} ({','.join(columns)}) VALUES ({placeholders})",
                    tuple(row.get(column) for column in columns),
                )
    score_raw = str((payload.get("settings") or {}).get(SCORE_SETTING_KEY) or "")
    set_setting(SCORE_SETTING_KEY, score_raw)
    _record_change("备份恢复", safe_name, "恢复", updated_by, None, {"backup": safe_name})
    _clear_confirmation_cache()


def _record_change(
    category: str,
    object_id: str,
    action: str,
    employee_id: str,
    before: Any,
    after: Any,
) -> None:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_changes
                (category, object_id, action, employee_id, before_json, after_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (category, object_id, action, employee_id, _dump(before), _dump(after), _now()),
        )


def _lookup_override(group: str, key: str) -> dict[str, Any] | None:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT group_key, input_value, output_value, deleted, updated_by, updated_at
            FROM transcode_rule_center_lookup_overrides
            WHERE group_key = ? AND input_value = ? AND deleted = 0
            """,
            (group, key),
        ).fetchone()
    return dict(row) if row else None


def _asset_override(group: str, row_id: str) -> dict[str, Any] | None:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT row_json FROM transcode_rule_center_asset_overrides
            WHERE asset_group=? AND row_id=? AND deleted=0
            """,
            (group, row_id),
        ).fetchone()
    return json.loads(row["row_json"]) if row and row["row_json"] else None


def _confirmation_override(rule_id: str) -> dict[str, Any] | None:
    ensure_rule_center_tables()
    with db_cursor() as conn:
        row = conn.execute(
            """
            SELECT rule_json FROM transcode_rule_center_confirmation_overrides
            WHERE rule_id = ? AND deleted = 0
            """,
            (rule_id,),
        ).fetchone()
    return json.loads(row["rule_json"]) if row and row["rule_json"] else None


def _clear_confirmation_cache() -> None:
    from .transcode_confirmation_policy import load_confirmation_policy_rules

    load_confirmation_policy_rules.cache_clear()


def _split_business_values(value: Any, *, separators: str = r"[,，、;；]+") -> list[str]:
    import re

    return [item.strip() for item in re.split(separators, str(value or "")) if item.strip()]


def _cleanup_backups() -> None:
    cutoff = datetime.now() - timedelta(days=30)
    for path in BACKUP_DIR.glob("transcode-rules-*.json"):
        if datetime.fromtimestamp(path.stat().st_mtime) < cutoff:
            path.unlink(missing_ok=True)


def _new_rule_id() -> str:
    return datetime.now().strftime("BASE-PAGE-%Y%m%d-%H%M%S-%f")


def _form_bool(value: Any, *, default: bool) -> bool:
    if value is None:
        return default
    return _clean(value).lower() in {"1", "true", "yes", "on", "是"}


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _normalize_lookup_values(group: str, key_value: Any, output_value: Any) -> tuple[str, str]:
    key = _clean(key_value)
    value = _clean(output_value)
    if group in {"glue_code", "glue_category"}:
        key = key.upper()
    elif group in {"grade_code", "grade_trigger", "copper_type", "copper_valid"}:
        key = key.upper()
    elif group in {"total_to_core", "core_to_total"}:
        key = re.sub(r"\s+", "", key).upper()
    if group in {
        "glue_code", "glue_category", "grade_code", "grade_trigger",
        "copper_micron", "copper_type",
        "copper_valid",
    }:
        value = value.upper()
    return key, value


def _dump(value: Any) -> str | None:
    if value is None:
        return None
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def _now() -> str:
    return datetime.now().replace(microsecond=0).isoformat()
