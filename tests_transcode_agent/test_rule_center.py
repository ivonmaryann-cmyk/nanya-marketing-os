from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from pathlib import Path

import pytest
from werkzeug.datastructures import MultiDict

from fangzheng_web_app import db
from fangzheng_web_app.transcode_agent_service import _customer_matches
from fangzheng_web_app.transcode_confirmation_policy import (
    apply_confirmation_rules_to_evidence,
    load_confirmation_policy_rules,
    match_confirmation_policy_rules,
)
from fangzheng_web_app.transcode_rule_center import (
    BUSINESS_FIELDS,
    build_rule_center_lookup_tables,
    build_base_rule_from_form,
    build_confirmation_policy_from_form,
    create_backup,
    delete_lookup_override,
    list_backups,
    ensure_daily_backup,
    find_base_override,
    list_business_rule_rows,
    list_lookup_rows,
    load_score_config,
    merge_base_rule_overrides,
    merge_agent_mapping_overrides,
    merge_lookup_overrides,
    merge_score_matrix,
    save_confirmation_policy,
    save_asset_override,
    save_base_override,
    save_lookup_override,
    save_score_config,
    summarize_asset_rows,
    summarize_lookup_rows,
    summarize_semantic_rules,
    restore_backup,
    RuleCenterError,
)


def _use_temp_database(monkeypatch, tmp_path) -> None:
    monkeypatch.setattr(db, "DATABASE_PATH", tmp_path / "app.db")
    monkeypatch.setattr(
        "fangzheng_web_app.transcode_rule_center.BACKUP_DIR",
        tmp_path / "backups",
    )
    db.init_db()


def test_legacy_global_override_remains_runtime_matchable(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("business_field", "基板级别"),
            ("source_text", "规格包含汽车板时基板级别为AC"),
            ("condition_keyword", "汽车板"),
            ("target_value", "AC"),
            ("priority", "150"),
            ("enabled", "1"),
        ]
    )
    rule = build_base_rule_from_form(form)
    save_base_override(rule, updated_by="cyb")
    merged = merge_base_rule_overrides([])

    assert len(merged) == 1
    assert merged[0]["覆盖字段"] == "grade_code"
    assert merged[0]["覆盖值"] == "AC"
    assert _customer_matches(merged[0], "任意代码", "任意客户") is True


def test_lookup_override_takes_priority_without_changing_baseline(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    tables = {
        "glue_model_map": {"NY2150": "2B"},
        "glue_exact_map": {"NY2150": "2B"},
        "glue_cat_map": {},
        "grade_desc_to_code": {},
        "grade_code_map": {},
        "thick_total_to_core": {},
        "thick_core_to_total": {},
    }
    save_lookup_override(
        MultiDict(
            [
                ("lookup_group", "glue_code"),
                ("lookup_input", "NY2150"),
                ("lookup_output", "2T"),
            ]
        ),
        updated_by="cyb",
    )
    merge_lookup_overrides(tables)

    assert tables["glue_model_map"]["NY2150"] == "2T"
    assert tables["glue_exact_map"]["NY2150"] == "2T"
    rows = list_lookup_rows({"glue_model_map": {"NY2150": "2B"}}, group_key="glue_code")
    assert rows[0]["output_value"] == "2T"
    assert rows[0]["source"] == "页面调整"


def test_rule_center_lookup_view_separates_old_and_latest_glue_tables(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    tables = build_rule_center_lookup_tables(
        {
            "grade_desc_to_code": {"汽车板": "AC"},
            "glue_exact_map": {"NY1140": "1A", "NY2150": "2B"},
            "glue_model_map": {"NY2150": "2B", "NY3150HC": "3H"},
        },
        {
            "Agent胶系主表": [
                {"启用": "是", "胶系名称": "NY2150", "输出胶系代码": "2B"},
                {"启用": "是", "胶系名称": "NY6300S", "输出胶系代码": "6C"},
            ]
        },
        official_grade_codes={"A1", "AC", "AY"},
        standard_sizes={940.0: 37.0, 1245.0: 49.0},
        high_speed_mil={3.0: 0.079},
        copper_micron={"18": "H"},
        copper_types=[("HTE", "W")],
        copper_valid={"H/H"},
        size_ranges=[(36.9, 37.1, 48.9, 49.1, 37.0, 49.0)],
    )

    glue_rows = list_lookup_rows(tables, group_key="glue_code")
    grade_rows = list_lookup_rows(tables, group_key="grade_code")
    size_rows = list_lookup_rows(tables, group_key="standard_size")

    assert {row["input_value"] for row in glue_rows} == {"NY1140", "NY2150", "NY3150HC"}
    assert all(row["source"] == "老表" for row in glue_rows)
    assert tables["rule_center_latest_glue_code"] == {"NY2150": "2B", "NY6300S": "6C"}
    assert {row["input_value"] for row in grade_rows} == {"A1", "AC", "AY"}
    assert {row["input_value"] for row in size_rows} == {"940.0", "1245.0"}


def test_glue_business_rows_keep_table_order_and_business_source_scope(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    lookup_tables = {
        "rule_center_glue_code": {"OLD-B": "OB", "OLD-A": "OA"},
        "__lookup_sources": {"glue_code": {"OLD-B": "老表", "OLD-A": "老表"}},
    }
    mapping_tables = {
        "Agent胶系主表": [
            {"映射ID": "NEW-2", "启用": "是", "胶系名称": "NEW-B", "输出胶系代码": "NB"},
            {"映射ID": "NEW-1", "启用": "是", "胶系名称": "NEW-A", "输出胶系代码": "NA"},
        ],
        "Agent胶系兼容别名": [
            {"映射ID": "ALIAS-DUP-OLD", "启用": "是", "兼容名称": "OLD-B", "输出胶系代码": "OB"},
            {"映射ID": "ALIAS-DUP-NEW", "启用": "是", "兼容名称": "NEW-B", "输出胶系代码": "NB"},
            {"映射ID": "ALIAS-1", "启用": "是", "兼容名称": "OLD-ALIAS", "输出胶系代码": "OA"},
            {"映射ID": "ALIAS-2", "启用": "是", "兼容名称": "OLD-ALIAS", "输出胶系代码": "OA"},
        ],
        "Agent胶系选择规则": [
            {"映射ID": "SELECT-1", "启用": "是", "胶系名称": "MULTI", "输出胶系代码": "MX"},
        ],
        "Agent基础条件规则": [
            {"映射ID": "COND-1", "启用": "是", "条件胶系": "TFT-GLUE", "覆盖胶系代码": "TG"},
        ],
    }

    rows = list_business_rule_rows(lookup_tables, mapping_tables, category="胶系")

    assert [(row["source_scope"], row["title"]) for row in rows] == [
        ("老表", "OLD-B"),
        ("老表", "OLD-A"),
        ("新表", "NEW-B"),
        ("新表", "NEW-A"),
        ("额外正式补充", "OLD-ALIAS"),
    ]
    assert [row["type_label"] for row in rows] == [
        "老表",
        "老表",
        "新表",
        "新表",
        "额外正式补充",
    ]
    assert "MULTI" not in {row["title"] for row in rows}
    assert "TFT-GLUE" not in {row["title"] for row in rows}
    assert [row["group_start"] for row in rows] == [True, False, True, False, True]


def test_all_eight_base_categories_keep_formal_mappings_maintainable(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    lookup_tables = {
        "rule_center_glue_code": {"NY2150": "2B"},
        "__lookup_sources": {"glue_code": {"NY2150": "老表"}},
        "rule_center_high_speed_mil": {3.0: 0.079},
        "rule_center_copper_micron": {"18": "H"},
        "rule_center_copper_valid": {"H/H": "HH"},
        "rule_center_standard_size": {940.0: 37.0},
        "rule_center_size_range": {"36.9-37.1 × 48.9-49.1": "37 × 49"},
        "glue_cat_map": {"NY2150": "R"},
        "rule_center_copper_type": {"HTE": "W"},
        "rule_center_grade_code": {"AC": "AC"},
        "rule_center_grade_trigger": {"汽车板": "AC"},
        "thick_total_to_core": {"H/H": 0.07},
        "thick_core_to_total": {"H/H": 0.07},
    }

    projected = {
        category: list_business_rule_rows(lookup_tables, {}, category=category)
        for category in BUSINESS_FIELDS
    }

    assert all(projected[category] for category in BUSINESS_FIELDS)
    assert all(
        row["kind"] in {"lookup", "asset"}
        for rows in projected.values()
        for row in rows
    )
    assert all(row["source_scope"] for rows in projected.values() for row in rows)
    assert {row["source_scope"] for row in projected["基板尺寸"]} == {
        "编码规范",
        "确定性算法",
    }
    total_core = projected["总/芯厚"]
    assert [(row["source_scope"], row["lookup_group"], row["detail"]) for row in total_core] == [
        ("总芯厚转换表", "total_to_core", "总厚转芯厚"),
        ("总芯厚转换表", "core_to_total", "芯厚转总厚"),
    ]


def test_glue_business_rows_hide_retired_ny_a1_2z_and_mark_multi_code_conflicts(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rows = list_business_rule_rows(
        {
            "rule_center_glue_code": {"NY-A1": "2Z", "NY2150": "2B"},
            "__lookup_sources": {"glue_code": {"NY-A1": "老表", "NY2150": "老表"}},
        },
        {
            "Agent胶系主表": [
                {"映射ID": "NEW-1", "启用": "是", "胶系名称": "NY2150", "输出胶系代码": "2T"},
            ]
        },
        category="胶系",
    )

    assert "NY-A1" not in {row["title"] for row in rows}
    ny2150 = [row for row in rows if row["title"] == "NY2150"]
    assert {row["result"] for row in ny2150} == {"2B", "2T"}
    assert all(row["conflict"] is True for row in ny2150)
    assert all(row["status_label"] == "冲突待核实" for row in ny2150)
    assert all(row["eligible_for_formal_score"] is False for row in ny2150)


def test_latest_glue_table_keeps_same_name_multi_code_visible_and_pending(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rows = list_business_rule_rows(
        {},
        {
            "Agent胶系主表": [
                {"映射ID": "NEW-1", "启用": "是", "胶系名称": "NY6300S", "输出胶系代码": "6C"},
                {"映射ID": "NEW-2", "启用": "是", "胶系名称": "NY6300S", "输出胶系代码": "B1"},
                {"映射ID": "NEW-3", "启用": "否", "胶系名称": "NY6300S", "输出胶系代码": "OLD"},
            ]
        },
        category="胶系",
    )

    assert [(row["source_scope"], row["title"], row["result"]) for row in rows] == [
        ("新表", "NY6300S", "6C"),
        ("新表", "NY6300S", "B1"),
        ("新表", "NY6300S", "OLD"),
    ]
    assert [row["status_label"] for row in rows] == ["冲突待核实", "冲突待核实", "已停用"]
    assert [row["eligible_for_formal_score"] for row in rows] == [False, False, False]


def test_glue_page_keeps_rc_compatibility_but_hides_only_retired_2z(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rows = list_business_rule_rows(
        {
            "rule_center_glue_code": {
                "NY-A1": "2Z",
                "NY-A1白纹改善": "RC",
                "2ZZN": "RC",
            },
            "__lookup_sources": {
                "glue_code": {
                    "NY-A1": "老表",
                    "NY-A1白纹改善": "老表",
                    "2ZZN": "老表",
                }
            },
        },
        {
            "Agent胶系主表": [],
            "Agent胶系兼容别名": [
                {
                    "映射ID": "ALIAS-NY-A1-RC",
                    "启用": "是",
                    "兼容名称": "NY-A1",
                    "输出胶系代码": "RC",
                }
            ],
        },
        category="胶系",
    )

    pairs = {(row["title"], row["result"]) for row in rows}
    assert ("NY-A1", "2Z") not in pairs
    assert ("NY-A1", "RC") in pairs
    assert ("NY-A1白纹改善", "RC") in pairs
    assert ("2ZZN", "RC") in pairs


def test_customer_limited_grade_writes_are_not_projected_as_base_rules(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rows = list_business_rule_rows(
        {
            "rule_center_grade_code": {"AC": "AC"},
            "rule_center_grade_trigger": {
                "汽车板": "AC",
                "深南90022-1": "D1",
                "深南90022-2": "D2",
                "深南测试": "D5",
                "江西测试": "A8",
            },
        },
        {},
        category="基板级别",
    )

    titles = {row["title"] for row in rows}
    assert "汽车板" in titles
    assert "AC" in titles
    assert titles.isdisjoint({"深南90022-1", "深南90022-2", "深南测试", "江西测试"})
    assert {row["source_scope"] for row in rows} == {"正式映射表", "编码规范"}


def test_base_rule_template_exposes_business_search_filters_and_explicit_edit_action():
    template = (
        Path(__file__).resolve().parents[1]
        / "templates"
        / "transcode_rule_center.html"
    ).read_text(encoding="utf-8")

    assert 'data-rule-list' in template
    assert 'data-source-label="规则范围"' in template
    assert 'data-source-all-label="全部范围"' in template
    assert 'data-status=' in template
    assert 'rc-edit-action' in template
    assert '>编辑</a>' in template
    assert '新增直接对应' not in template
    assert '新增条件规则' not in template


def test_lookup_business_summary_groups_same_output_without_changing_rows():
    rows = [
        {"input_value": "NY2150", "output_value": "2B", "source": "最新版胶系主表", "deleted": False},
        {"input_value": "NY2150H", "output_value": "2B", "source": "最新版胶系主表", "deleted": False},
        {"input_value": "NY3150HF", "output_value": "AH", "source": "页面调整", "deleted": False},
    ]

    summary = summarize_lookup_rows(rows)

    grouped = next(item for item in summary if item["output_value"] == "2B")
    assert grouped["count"] == 2
    assert grouped["inputs"] == ["NY2150", "NY2150H"]
    assert len(rows) == 3


def test_asset_business_summary_groups_customer_rules_by_customer():
    rows = [
        {"_row_id": "SIZE-1", "_source": "活动Agent资产", "启用": "是", "客户代码": "100", "客户简称": "客户A", "客户尺寸W": "37", "客户尺寸H": "49", "目标size_code": "37004900"},
        {"_row_id": "SIZE-2", "_source": "页面调整", "启用": "是", "客户代码": "100", "客户简称": "客户A", "客户尺寸W": "37.3", "客户尺寸H": "49.3", "目标size_code": "37304930"},
        {"_row_id": "SIZE-3", "_source": "活动Agent资产", "启用": "否", "客户代码": "200", "客户简称": "客户B", "客户尺寸W": "41", "客户尺寸H": "49", "目标size_code": "41004900"},
    ]

    summary = summarize_asset_rows(rows, asset_group="客户尺寸映射")

    customer_a = next(item for item in summary if item["label"] == "客户A")
    assert customer_a["count"] == 2
    assert customer_a["enabled_count"] == 2
    assert customer_a["result_summary"] == "37004900、37304930"


def test_semantic_business_summary_groups_shared_intent_across_customers():
    rules = [
        {"rule_id": "S1", "enabled": True, "customer_name": "客户A", "business_field": "基板级别", "source_text": "下汽车板", "normalized_values": ["AC"]},
        {"rule_id": "S2", "enabled": True, "customer_name": "客户B", "business_field": "基板级别", "source_text": "要汽板", "normalized_values": ["AC"]},
        {"rule_id": "S3", "enabled": True, "customer_name": "客户A", "business_field": "基板级别", "source_text": "需要AP板", "normalized_values": ["AP"]},
    ]

    summary = summarize_semantic_rules(rules)

    ac = next(item for item in summary if item["target"] == "AC")
    assert ac["count"] == 2
    assert ac["customer_count"] == 2
    assert ac["customer_preview"] == "客户A、客户B"


def test_page_asset_override_merges_into_runtime_without_mutating_source(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    source = {
        "Agent胶系主表": [
            {
                "映射ID": "TGM-1",
                "启用": "是",
                "胶系名称": "NY2150",
                "输出胶系代码": "2B",
            }
        ]
    }
    save_asset_override(
        "Agent胶系主表",
        {
            "映射ID": "TGM-1",
            "启用": "是",
            "胶系名称": "NY2150",
            "输出胶系代码": "2T",
        },
        updated_by="cyb",
    )

    merged = merge_agent_mapping_overrides(source)

    assert merged["Agent胶系主表"][0]["输出胶系代码"] == "2T"
    assert source["Agent胶系主表"][0]["输出胶系代码"] == "2B"


def test_lookup_input_is_normalized_and_deleted_baseline_is_visible_as_disabled(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    save_lookup_override(
        MultiDict(
            [
                ("lookup_group", "glue_code"),
                ("lookup_input", "ny2150"),
                ("lookup_output", "2t"),
            ]
        ),
        updated_by="cyb",
    )
    tables = {"glue_model_map": {"NY2150": "2B"}, "glue_exact_map": {"NY2150": "2B"}}
    merge_lookup_overrides(tables)
    assert tables["glue_model_map"]["NY2150"] == "2T"

    delete_lookup_override("glue_code", "ny2150", updated_by="cyb")
    runtime_tables = {
        "glue_model_map": {"NY2150": "2B"},
        "glue_exact_map": {"NY2150": "2B"},
    }
    merge_lookup_overrides(runtime_tables)
    assert "NY2150" not in runtime_tables["glue_model_map"]
    rows = list_lookup_rows({"glue_model_map": {"NY2150": "2B"}}, group_key="glue_code")
    assert rows[0]["input_value"] == "NY2150"
    assert rows[0]["output_value"] == "2B"
    assert rows[0]["source"] == "页面停用"
    assert rows[0]["deleted"] is True


def test_score_config_keeps_formal_gate_at_100(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    saved = save_score_config(
        MultiDict(
            [
                ("semantic_supported_score", "97"),
                ("model_supported_score", "94"),
                ("ambiguous_score", "75"),
                ("missing_evidence_score", "55"),
                ("contradicted_score", "0"),
            ]
        ),
        updated_by="cyb",
    )
    assert saved["gate_threshold"] == 100
    assert load_score_config()["semantic_supported_score"] == 97
    matrix = merge_score_matrix(
        {
            "gate_threshold": 90,
            "verdict_scores": {
                "supported": {"mode": "preserve"},
                "contradicted": {"mode": "fixed", "value": 1},
                "ambiguous": {"mode": "cap", "value": 1},
                "missing_evidence": {"mode": "cap", "value": 1},
            },
        }
    )
    assert matrix["gate_threshold"] == 100
    assert matrix["verdict_scores"]["missing_evidence"]["value"] == 55


@pytest.mark.parametrize(
    "unsafe_key",
    [
        "semantic_supported_score",
        "model_supported_score",
        "ambiguous_score",
        "missing_evidence_score",
        "contradicted_score",
    ],
)
def test_uncertain_scores_cannot_be_configured_as_formal_100(monkeypatch, tmp_path, unsafe_key):
    _use_temp_database(monkeypatch, tmp_path)
    values = {
        "semantic_supported_score": "98",
        "model_supported_score": "95",
        "ambiguous_score": "80",
        "missing_evidence_score": "60",
        "contradicted_score": "0",
    }
    values[unsafe_key] = "100"
    with pytest.raises(RuleCenterError, match="0到99"):
        save_score_config(MultiDict(values.items()), updated_by="cyb")
    assert load_score_config()["gate_threshold"] == 100


def test_unsafe_historical_score_config_is_clamped_at_runtime(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    db.set_setting(
        "transcode_rule_center_score_config",
        json.dumps(
            {
                "gate_threshold": 80,
                "semantic_supported_score": 100,
                "model_supported_score": "invalid",
                "ambiguous_score": 120,
                "missing_evidence_score": -5,
                "contradicted_score": 100,
            }
        ),
    )

    config = load_score_config()

    assert config["gate_threshold"] == 100
    assert config["semantic_supported_score"] == 99
    assert config["model_supported_score"] == 95
    assert config["ambiguous_score"] == 99
    assert config["missing_evidence_score"] == 0
    assert config["contradicted_score"] == 99


def test_malformed_historical_base_override_is_not_loaded_into_runtime(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    ensure_daily_backup()
    malformed = {
        "rule_id": "BASE-UNSAFE-001",
        "business_field": "基板级别",
        "target_value": "A",
        "native_rule": {"规则ID": "BASE-UNSAFE-001", "覆盖值": "A"},
    }
    with db.db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_rule_center_base_overrides
                (rule_id, rule_json, deleted, updated_by, updated_at)
            VALUES (?, ?, 0, ?, ?)
            """,
            ("BASE-UNSAFE-001", json.dumps(malformed, ensure_ascii=False), "legacy", "2026-01-01"),
        )

    assert merge_base_rule_overrides([]) == []


@pytest.mark.parametrize(
    ("business_field", "valid_value", "invalid_value"),
    [
        ("胶系", "AH", "A"),
        ("基板厚度", "00800", "800"),
        ("铜箔规格", "HH", "H/H"),
        ("基板尺寸", "37004900", "37X49"),
        ("胶水类别", "R", "RR"),
        ("铜箔类型+印字/非印字", "W", "HTE"),
        ("基板级别", "AC", "A"),
        ("总/芯厚", "T", "X"),
    ],
)
def test_base_rule_target_code_is_validated_by_business_field(
    business_field, valid_value, invalid_value
):
    base = [
        ("business_field", business_field),
        ("source_text", "测试业务规则"),
        ("condition_keyword", "测试条件"),
        ("enabled", "1"),
    ]
    rule = build_base_rule_from_form(MultiDict(base + [("target_value", valid_value)]))
    assert rule["target_value"] == valid_value
    with pytest.raises(RuleCenterError):
        build_base_rule_from_form(MultiDict(base + [("target_value", invalid_value)]))


def test_rule_center_backup_is_created(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    path = create_backup(reason="test")
    assert path.exists()
    backups = list_backups()
    assert backups and backups[0]["reason"] == "test"


def test_confirmation_policy_is_page_maintainable_and_runtime_effective(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    load_confirmation_policy_rules.cache_clear()
    rule = build_confirmation_policy_from_form(
        MultiDict(
            [
                ("confirmation_customers", "测试客户"),
                ("confirmation_field", "grade"),
                ("confirmation_contains_all", "NY2150H"),
                ("confirmation_any_groups", "HW/华为；汽车板/汽板"),
                ("confirmation_reason", "两类等级条件同时命中，需人工选择"),
                ("confirmation_enabled", "1"),
            ]
        )
    )
    save_confirmation_policy(rule, updated_by="cyb")

    matches = match_confirmation_policy_rules(
        "测试客户",
        "NY2150H 华为板",
        "下汽板",
    )
    assert [item["rule_id"] for item in matches] == [rule["rule_id"]]
    evidence = apply_confirmation_rules_to_evidence(
        [{"field_key": "grade", "field": "基板级别", "score": 100, "gate": True}],
        matches,
    )
    assert evidence[0]["score"] == 99
    assert evidence[0]["decision_state"] == "条件不足"


def test_editing_existing_confirmation_policy_preserves_advanced_condition():
    existing = {
        "rule_id": "CPR-EXISTING",
        "status": "approved",
        "basis_type": "non_unique_mapping",
        "customers": ["依顿"],
        "contains_all": ["NY2150H"],
        "contains_any_groups": [["HW", "华为"]],
        "copper_pair_mixed_threshold": 1.5,
        "field": "基板级别",
        "field_keys": ["grade"],
        "reason": "原因",
    }
    edited = build_confirmation_policy_from_form(
        MultiDict(
            [
                ("confirmation_rule_id", "CPR-EXISTING"),
                ("confirmation_customers", "依顿，伊顿"),
                ("confirmation_field", "grade"),
                ("confirmation_contains_all", "NY2150H"),
                ("confirmation_any_groups", "HW/华为"),
                ("confirmation_reason", "更新后原因"),
                ("confirmation_enabled", "1"),
            ]
        ),
        existing=existing,
    )
    assert edited["basis_type"] == "non_unique_mapping"
    assert edited["copper_pair_mixed_threshold"] == 1.5


def test_daily_backup_retains_only_30_days(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    backup_dir = tmp_path / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    old_path = backup_dir / "transcode-rules-2000-01-01.json"
    old_path.write_text("{}", encoding="utf-8")
    old_time = (datetime.now() - timedelta(days=31)).timestamp()
    os.utime(old_path, (old_time, old_time))

    daily_path = ensure_daily_backup()

    assert daily_path.exists()
    assert not old_path.exists()


def test_backup_restore_recovers_page_rule_and_score(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    original = build_base_rule_from_form(
        MultiDict(
            [
                ("business_field", "基板级别"),
                ("source_text", "规格包含安全类时为AY"),
                ("condition_keyword", "安全类"),
                ("target_value", "AY"),
                ("enabled", "1"),
            ]
        )
    )
    save_base_override(original, updated_by="cyb")
    save_score_config(
        MultiDict(
            [
                ("semantic_supported_score", "97"),
                ("model_supported_score", "94"),
                ("ambiguous_score", "75"),
                ("missing_evidence_score", "55"),
                ("contradicted_score", "0"),
            ]
        ),
        updated_by="cyb",
    )
    backup = create_backup(reason="restore-test")

    changed = dict(original)
    changed["target_value"] = "AC"
    changed["native_rule"] = {**original["native_rule"], "覆盖值": "AC"}
    save_base_override(changed, updated_by="cyb")
    restore_backup(backup.name, updated_by="cyb")

    restored = find_base_override(original["rule_id"])
    assert restored["target_value"] == "AY"
    assert restored["native_rule"]["覆盖值"] == "AY"
    assert load_score_config()["semantic_supported_score"] == 97
