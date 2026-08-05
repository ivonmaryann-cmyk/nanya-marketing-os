from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from werkzeug.datastructures import MultiDict

from fangzheng_web_app import db
from fangzheng_web_app.transcode_customer_rule_admin import (
    CONDITION_FIELDS,
    CUSTOMER_METADATA_ASSET_TYPE,
    CONDITION_OPERATORS,
    CONDITION_OPERATOR_LABELS,
    INFORMATION_ASSET_TYPE,
    CODE_MIGRATION_ASSET_TYPE,
    CUSTOMER_ORDER_ASSET_TYPE,
    LEGACY_CCL_ASSET_TYPE,
    CustomerRuleMaintenanceError,
    build_agent_rule_from_form,
    build_rule_from_form,
    customer_rule_workspace,
    delete_rule_override,
    list_customer_rule_changes,
    legacy_rule_assets_for_customer_workspace,
    make_customer_key,
    mapping_assets_for_customer_workspace,
    merge_agent_rule_overrides,
    merge_customer_rule_overrides,
    project_customer_rule_assets_for_workspace,
    resolve_customer_code_by_name,
    restore_customer_rule_change,
    save_agent_rule_override,
    save_rule_override,
    validate_customer_maintained_rule,
)


def _base_rule() -> dict:
    return {
        "rule_id": "TSR-TEST-001",
        "source_candidate_id": "SRC-TEST-001",
        "enabled": True,
        "customer_code": "103901",
        "customer_name": "广东依顿",
        "source_row": 1,
        "source_column": "CCL特殊规则",
        "business_field": "基板级别",
        "source_text": "订单备注含汽车板时下AC",
        "semantic_types": ["keyword_present"],
        "target_fields": ["grade_intent"],
        "normalized_values": ["AC"],
        "stated_target_values": ["AC"],
        "conditions": [
            {
                "field": "订单备注",
                "operator": "contains_any",
                "value": ["汽车板"],
                "source_scope": "订单备注",
            }
        ],
        "required_input_fields": ["订单备注"],
        "execution_mode": "结构化后可确定性执行",
        "priority": 120,
        "model": "test",
        "prompt_sha256": "",
        "evidence_texts": ["订单备注含汽车板时下AC"],
        "approval": {"status": "confirmed", "basis": "测试"},
        "note": "",
    }


def test_global_conditional_assets_are_grouped_under_all_customers():
    rules = mapping_assets_for_customer_workspace(
        {
            "Agent胶系选择规则": [
                {
                    "映射ID": "SELECT-1",
                    "启用": "是",
                    "胶系名称": "NY-P4",
                    "输出胶系代码": "RA",
                    "规则文本": "NY-P4按已确认结果RA",
                }
            ],
            "Agent基础条件规则": [
                {
                    "映射ID": "COND-1",
                    "启用": "是",
                    "条件胶系": "NY3150HF",
                    "条件关键词": "TFT",
                    "覆盖胶系代码": "3B",
                    "覆盖胶水类别": "Y",
                    "覆盖基板级别": "AT",
                    "规则文本": "NY3150HF含TFT时覆盖",
                }
            ],
        },
        include_semantic_pending=False,
    )

    workspace = customer_rule_workspace(rules)

    assert len(rules) == 4
    assert workspace["selected_customer"]["name"] == "全部客户"
    assert workspace["selected_customer"]["global_scope"] is True
    assert workspace["selected_customer"]["rule_count"] == 4


def _use_temp_database(monkeypatch, tmp_path) -> None:
    monkeypatch.setattr(db, "DATABASE_PATH", tmp_path / "app.db")
    db.init_db()


def test_page_rule_override_is_merged_immediately(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    base = _base_rule()
    form = MultiDict(
        [
            ("rule_id", base["rule_id"]),
            ("customer_code", "103901"),
            ("customer_name", "广东依顿"),
            ("business_field", "基板级别"),
            ("source_text", "订单备注表达MINILED时下AM"),
            ("semantic_enabled", "1"),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "MINILED；mini led"),
            ("target_field", "grade_intent"),
            ("target_value", "AM"),
            ("priority", "130"),
            ("enabled", "1"),
            ("approval_basis", "测试确认"),
        ]
    )
    updated = build_rule_from_form(form, existing_rule=base)
    validate_customer_maintained_rule(updated)
    save_rule_override(updated, updated_by="cyb", previous_rule=base)

    merged = merge_customer_rule_overrides([base])

    assert len(merged) == 1
    assert merged[0]["normalized_values"] == ["AM"]
    assert merged[0]["conditions"][0]["value"] == ["MINILED", "mini led"]


def test_delete_and_restore_rule(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    base = _base_rule()
    delete_rule_override(base, updated_by="cyb")
    assert merge_customer_rule_overrides([base]) == []

    delete_change = next(item for item in list_customer_rule_changes() if item["action"] == "删除")
    restore_customer_rule_change(delete_change["id"], updated_by="cyb")

    restored = merge_customer_rule_overrides([base])
    assert len(restored) == 1
    assert restored[0]["rule_id"] == base["rule_id"]
    assert restored[0]["normalized_values"] == ["AC"]


def test_new_multi_condition_rule_keeps_deterministic_conditions(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_code", "103901"),
            ("customer_name", "广东依顿"),
            ("business_field", "基板级别"),
            ("source_text", "NY2150H且订单备注含汽车板时下AC"),
            ("semantic_enabled", "1"),
            ("condition_field", "胶系"),
            ("condition_operator", "equals"),
            ("condition_value", "NY2150H"),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "汽车板；下汽车板；要汽板"),
            ("target_field", "grade_intent"),
            ("target_value", "AC"),
            ("priority", "140"),
            ("enabled", "1"),
        ]
    )

    rule = build_rule_from_form(form)
    validate_customer_maintained_rule(rule)

    assert rule["semantic_types"] == ["multi_condition"]
    assert rule["required_input_fields"] == ["胶系", "订单备注"]
    assert rule["normalized_values"] == ["AC"]


def test_customer_name_without_code_can_be_maintained(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_code", ""),
            ("customer_name", "确认中心测试客户"),
            ("business_field", "基板级别"),
            ("source_text", "订单备注含汽车板时下AC"),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "汽车板"),
            ("target_field", "grade_intent"),
            ("target_value", "AC"),
            ("enabled", "1"),
        ]
    )

    rule = build_rule_from_form(form)
    validate_customer_maintained_rule(rule)

    assert rule["customer_code"] == ""
    assert rule["customer_name"] == "确认中心测试客户"


def test_agent_deterministic_rule_auto_generates_rule_text(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_name", "台湾敬鹏"),
            ("business_field", "铜箔类型+印字/非印字"),
            ("agent_override_field", "copper_type_code"),
            ("agent_condition_glue", "NYHP-7350"),
            ("target_value", "R"),
            ("enabled", "1"),
        ]
    )

    rule = build_agent_rule_from_form(form, existing_rule=None)
    save_agent_rule_override(rule, updated_by="cyb", previous_rule=None)
    merged = merge_agent_rule_overrides([])

    assert rule["规则文本"] == "客户台湾敬鹏；胶系=NYHP-7350；铜箔类型+印字/非印字=R"
    assert rule["条件文本"] == rule["规则文本"]
    assert len(merged) == 1
    assert merged[0]["规则文本"] == rule["规则文本"]
    assert merged[0]["覆盖字段"] == "copper_type_code"
    assert merged[0]["覆盖值"] == "R"


def test_page_maintained_agent_rule_is_editable_active(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_name", "台湾敬鹏"),
            ("business_field", "铜箔类型+印字/非印字"),
            ("agent_override_field", "copper_type_code"),
            ("agent_condition_glue", "NYHP-7350"),
            ("target_value", "R"),
            ("enabled", "1"),
        ]
    )
    rule = build_agent_rule_from_form(form, existing_rule=None)
    save_agent_rule_override(rule, updated_by="cyb", previous_rule=None)

    workspace = customer_rule_workspace(
        project_customer_rule_assets_for_workspace(
            [],
            merge_agent_rule_overrides([]),
            {},
            include_semantic_pending=False,
        ),
        customer_key=make_customer_key("", "台湾敬鹏"),
        business_field="铜箔类型+印字/非印字",
    )

    selected = workspace["selected_rule"]
    assert selected is not None
    assert selected["scope_key"] == "customer"
    assert selected["status_label"] == "启用"
    assert selected["editable"] is True
    assert selected["origin"] == "页面维护Agent长期规则"


def test_page_maintained_agent_rule_with_legacy_marker_is_editable(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rule = {
        "规则ID": "PAGE-CUST-20260804-LEGACY-MARKER",
        "启用": "是",
        "客户简称": "台湾敬鹏",
        "条件胶系": "NYHP-7350",
        "覆盖字段": "copper_type_code",
        "覆盖值": "R",
        "命中来源": "页面维护Agent确定性长期规则",
        "来源行号": "页面新增",
        "规则文本": "客户台湾敬鹏；胶系=NYHP-7350；铜箔类型+印字/非印字=R",
        "优先级": "100",
    }
    workspace = customer_rule_workspace(
        project_customer_rule_assets_for_workspace(
            [],
            [rule],
            {},
            include_semantic_pending=False,
        ),
        customer_key=make_customer_key("", "台湾敬鹏"),
        business_field="铜箔类型+印字/非印字",
    )

    selected = workspace["selected_rule"]
    assert selected is not None
    assert selected["editable"] is True
    assert selected["scope_key"] == "customer"


def test_semantic_rule_still_requires_source_text(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_name", "台湾敬鹏"),
            ("business_field", "铜箔类型+印字/非印字"),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "按R"),
            ("target_field", "copper_type"),
            ("target_value", "R"),
        ]
    )

    with pytest.raises(CustomerRuleMaintenanceError, match="业务触发条件不能为空"):
        build_rule_from_form(form)


def test_condition_field_list_is_simplified():
    assert "订单规格" not in CONDITION_FIELDS
    assert "品号/物料编号" not in CONDITION_FIELDS
    assert "订单规格/订单备注" in CONDITION_FIELDS
    assert "客户规格" in CONDITION_FIELDS
    assert "客户物料编码" in CONDITION_FIELDS


def test_legacy_condition_fields_are_normalized_when_saving(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    order_spec_form = MultiDict(
        [
            ("customer_name", "台湾敬鹏"),
            ("business_field", "基板级别"),
            ("source_text", "客户订单上有汽车板字样=AC"),
            ("condition_field", "订单规格"),
            ("condition_operator", "contains_any"),
            ("condition_value", "汽车板"),
            ("target_field", "grade_intent"),
            ("target_value", "AC"),
            ("enabled", "1"),
        ]
    )
    part_number_form = MultiDict(
        [
            ("customer_name", "台湾敬鹏"),
            ("business_field", "总/芯厚"),
            ("source_text", "物料编号含特定值=芯厚"),
            ("condition_field", "品号/物料编号"),
            ("condition_operator", "contains_any"),
            ("condition_value", "ABC"),
            ("target_field", "total_core"),
            ("target_value", "C"),
            ("enabled", "1"),
        ]
    )

    order_spec_rule = build_rule_from_form(order_spec_form)
    part_number_rule = build_rule_from_form(part_number_form)

    assert order_spec_rule["conditions"][0]["field"] == "客户规格"
    assert part_number_rule["conditions"][0]["field"] == "客户物料编码"


def test_rule_view_normalizes_legacy_condition_field():
    rule = _base_rule()
    rule["conditions"] = [
        {
            "field": "订单规格",
            "operator": "contains_any",
            "value": ["汽车板"],
            "source_scope": "订单规格",
        }
    ]

    workspace = customer_rule_workspace(
        [rule],
        customer_key=make_customer_key("103901", "广东依顿"),
        business_field="基板级别",
    )

    selected = workspace["selected_rule"]
    assert selected["conditions"][0]["field"] == "客户规格"


def test_confirmation_rule_origin_and_disabled_state_are_visible(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rule = _base_rule()
    rule.update(
        {
            "rule_id": "TCR-CONFIRMATION-001",
            "source_column": "确认中心",
            "model": "确认中心人工规则",
            "note": "确认中心长期规则",
        }
    )
    save_rule_override(rule, updated_by="cyb", previous_rule=None)
    merged = merge_customer_rule_overrides([])
    workspace = customer_rule_workspace(
        merged,
        customer_key=make_customer_key(rule["customer_code"], rule["customer_name"]),
        business_field="基板级别",
        rule_id=rule["rule_id"],
    )
    assert workspace["selected_rule"]["origin"] == "确认中心长期规则"
    assert workspace["selected_rule"]["enabled"] is True

    disabled_form = MultiDict(
        [
            ("rule_id", rule["rule_id"]),
            ("customer_code", rule["customer_code"]),
            ("customer_name", rule["customer_name"]),
            ("business_field", "基板级别"),
            ("source_text", rule["source_text"]),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "汽车板"),
            ("target_field", "grade_intent"),
            ("target_value", "AC"),
        ]
    )
    disabled = build_rule_from_form(disabled_form, existing_rule=rule)
    save_rule_override(disabled, updated_by="cyb", previous_rule=rule)
    assert merge_customer_rule_overrides([])[0]["enabled"] is False


def test_disabled_rules_do_not_appear_in_active_workspace(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    active = _base_rule()
    disabled = dict(_base_rule(), rule_id="TSR-TEST-DISABLED", enabled=False)

    workspace = customer_rule_workspace([active, disabled], rule_scope="active")

    assert workspace["rule_count"] == 1
    assert workspace["display_scope_counts"]["active"] == 1
    assert all(rule["enabled"] for rule in workspace["rules"])


def test_agent_rule_is_grouped_by_output_field_and_raw_rule_is_readonly():
    agent = {
        "规则ID": "TAR-TEST-TOTAL-CORE",
        "启用": "是",
        "客户代码": "100",
        "客户简称": "客户A",
        "原始字段": "基板厚度",
        "条件厚度": "<0.8",
        "覆盖字段": "tc_code",
        "覆盖值": "C",
        "规则文本": "字段=基板厚度|覆盖字段=tc_code",
    }
    workspace = customer_rule_workspace(
        project_customer_rule_assets_for_workspace(
            [], [agent], {}, include_semantic_pending=False
        ),
        customer_key=make_customer_key("100", "客户A"),
        business_field="总/芯厚",
    )

    selected = workspace["selected_rule"]
    assert selected["business_field"] == "总/芯厚"
    assert selected["editable"] is False
    assert "字段=" not in selected["source_text"]
    assert selected["scope_key"] == "migration"


def _f1_agent_rule(rule_id: str, priority: int, *, page_maintained: bool = False) -> dict:
    rule = {
        "规则ID": rule_id,
        "启用": "是",
        "客户代码": "123036",
        "客户简称": "江威尔高",
        "条件胶系": "NY2140",
        "条件厚度": ">=0.8",
        "覆盖字段": "grade_code",
        "覆盖值": "F1",
        "优先级": str(priority),
        "待确认": "否",
        "规则文本": f"字段=基板级别 | {rule_id}",
        "条件文本": "胶系:NY2140；厚度>=0.8",
    }
    if page_maintained:
        rule.update(
            {
                "命中来源": "页面维护Agent确定性长期规则",
                "来源行号": "页面新增",
            }
        )
    return rule


def test_duplicate_background_agent_rules_keep_higher_priority():
    rules = project_customer_rule_assets_for_workspace(
        [],
        [
            _f1_agent_rule("TAR-LOW", 950),
            _f1_agent_rule("TAR-HIGH", 1050),
        ],
        {},
        include_semantic_pending=False,
    )
    workspace = customer_rule_workspace(
        rules,
        customer_key=make_customer_key("123036", "江威尔高"),
        business_field="基板级别",
    )
    mine = [
        rule
        for rule in rules
        if make_customer_key(rule.get("customer_code"), rule.get("customer_name"))
        == make_customer_key("123036", "江威尔高")
        and rule.get("business_field") == "基板级别"
        and rule.get("asset_type") == "agent_deterministic"
    ]

    assert len(mine) == 1
    assert workspace["selected_rule"]["rule_id"] == "TAR-HIGH"


def test_duplicate_page_maintained_rule_wins_over_background_rule():
    rules = project_customer_rule_assets_for_workspace(
        [],
        [
            _f1_agent_rule("TAR-BACKGROUND", 1050),
            _f1_agent_rule("PAGE-RULE", 950, page_maintained=True),
        ],
        {},
        include_semantic_pending=False,
    )
    workspace = customer_rule_workspace(
        rules,
        customer_key=make_customer_key("123036", "江威尔高"),
        business_field="基板级别",
    )
    mine = [
        rule
        for rule in rules
        if make_customer_key(rule.get("customer_code"), rule.get("customer_name"))
        == make_customer_key("123036", "江威尔高")
        and rule.get("business_field") == "基板级别"
        and rule.get("asset_type") == "agent_deterministic"
    ]

    assert len(mine) == 1
    assert workspace["selected_rule"]["rule_id"] == "PAGE-RULE"
    assert workspace["selected_rule"]["editable"] is True


def test_resolve_customer_code_by_name_uses_existing_rule_sources(monkeypatch):
    import fangzheng_web_app.transcode_agent_rules as agent_rules_mod
    import fangzheng_web_app.transcode_semantic_rules as semantic_rules_mod

    monkeypatch.setattr(
        agent_rules_mod,
        "load_transcode_agent_rules",
        lambda: [{"客户代码": "123036", "客户简称": "江威尔高"}],
    )
    monkeypatch.setattr(
        agent_rules_mod,
        "load_transcode_agent_mapping_tables",
        lambda: {
            "客户字段映射": [
                {"客户代码": "123036", "客户简称": "江威尔高"},
            ]
        },
    )
    monkeypatch.setattr(
        semantic_rules_mod,
        "get_active_transcode_semantic_rule_version",
        lambda: "",
    )
    monkeypatch.setattr(
        semantic_rules_mod,
        "load_transcode_semantic_rules",
        lambda version: [],
    )

    assert resolve_customer_code_by_name("江威尔高") == "123036"
    assert resolve_customer_code_by_name("不存在的客户") == ""


def test_long_term_rule_auto_fills_customer_code(monkeypatch):
    from fangzheng_web_app import transcode_agent_service as service

    monkeypatch.setattr(
        service,
        "resolve_customer_code_by_name",
        lambda name: "123036" if name == "江威尔高" else "",
    )
    item = {
        "customer_code": "",
        "customer_name": "江威尔高",
        "field_key": "grade",
        "field_label": "基板级别",
        "job_id": 55,
        "excel_row": 35,
    }
    payload = {
        "condition_field": "胶系",
        "condition_operator": "contains_any",
        "condition_value": "NY3150HC",
        "second_confirmed": True,
    }

    rule = service._build_long_term_confirmation_rule(
        item,
        "AL",
        basis="业务确认：基板级别按AL；条件：胶系 包含任一 NY3150HC",
        payload=payload,
    )

    assert rule["customer_code"] == "123036"
    assert rule["customer_name"] == "江威尔高"


def test_existing_semantic_rule_hides_duplicate_agent_projection():
    semantic = _base_rule()
    agent = {
        "规则ID": "TAR-DUPLICATE-001",
        "启用": "是",
        "客户代码": "103901",
        "客户简称": "广东依顿",
        "条件关键词": "当备注中有汽车板字样时",
        "覆盖字段": "grade_code",
        "覆盖值": "AC",
        "规则文本": "字段=基板级别|覆盖字段=grade_code",
    }

    rules = project_customer_rule_assets_for_workspace(
        [semantic], [agent], {}, include_semantic_pending=False
    )

    rule_ids = {rule["rule_id"] for rule in rules}
    assert semantic["rule_id"] in rule_ids
    assert agent["规则ID"] not in rule_ids


def test_all_condition_operators_have_chinese_labels():
    assert set(CONDITION_OPERATOR_LABELS) == set(CONDITION_OPERATORS)
    assert all(CONDITION_OPERATOR_LABELS[operator] for operator in CONDITION_OPERATORS)


def test_customer_mapping_assets_are_projected_into_customer_workspace(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    mapping_rules = mapping_assets_for_customer_workspace(
        {
            "客户尺寸映射": [
                {
                    "映射ID": "TAM-SIZE-TEST",
                    "启用": "是",
                    "客户代码": "103901",
                    "客户简称": "广东依顿",
                    "客户尺寸W": "36",
                    "客户尺寸H": "48",
                    "目标size_code": "37004900",
                    "规则文本": "36*48转37*49",
                }
            ]
        },
        include_semantic_pending=False,
    )
    workspace = customer_rule_workspace(
        mapping_rules,
        customer_key=make_customer_key("103901", "广东依顿"),
        business_field="基板尺寸",
    )

    selected = workspace["selected_rule"]
    assert selected["asset_type"] == "customer_mapping"
    assert selected["target_value"] == "37004900"
    assert selected["mapping_group"] == "客户尺寸映射"


def test_customer_identity_is_required(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("business_field", "基板级别"),
            ("source_text", "订单备注含汽车板时下AC"),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "汽车板"),
            ("target_field", "grade_intent"),
            ("target_value", "AC"),
        ]
    )
    with pytest.raises(CustomerRuleMaintenanceError, match="至少填写一个"):
        build_rule_from_form(form)


def test_invalid_grade_code_is_rejected(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_name", "确认中心测试客户"),
            ("business_field", "基板级别"),
            ("source_text", "订单备注含汽车板时下ZZ"),
            ("condition_field", "订单备注"),
            ("condition_operator", "contains_any"),
            ("condition_value", "汽车板"),
            ("target_field", "grade_intent"),
            ("target_value", "ZZ"),
        ]
    )
    with pytest.raises(CustomerRuleMaintenanceError, match="基板级别代码无效"):
        build_rule_from_form(form)


def test_semantic_switch_requires_order_remark_condition(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    form = MultiDict(
        [
            ("customer_name", "确认中心测试客户"),
            ("business_field", "基板级别"),
            ("source_text", "胶系为NY-A2时下AC"),
            ("semantic_enabled", "1"),
            ("condition_field", "胶系"),
            ("condition_operator", "equals"),
            ("condition_value", "NY-A2"),
            ("target_field", "grade_intent"),
            ("target_value", "AC"),
        ]
    )
    with pytest.raises(CustomerRuleMaintenanceError, match="订单备注条件"):
        build_rule_from_form(form)


def test_customer_workspace_separates_deterministic_and_order_semantic_rules():
    semantic = _base_rule()
    deterministic = dict(_base_rule())
    deterministic.update(
        {
            "rule_id": "TSR-TEST-DET-001",
            "source_text": "客户规格含NY-A2时下AC",
            "conditions": [
                {"field": "客户规格", "operator": "contains_any", "value": ["NY-A2"]}
            ],
        }
    )

    semantic_workspace = customer_rule_workspace([semantic, deterministic], rule_kind="semantic")
    deterministic_workspace = customer_rule_workspace(
        [semantic, deterministic], rule_kind="deterministic"
    )

    assert semantic_workspace["rule_count"] == 1
    assert semantic_workspace["selected_rule"]["rule_id"] == semantic["rule_id"]
    assert semantic_workspace["rule_kind"] == "semantic"
    assert deterministic_workspace["rule_count"] == 1
    assert deterministic_workspace["selected_rule"]["rule_id"] == deterministic["rule_id"]
    assert deterministic_workspace["rule_kind"] == "deterministic"


def test_customer_workspace_supports_business_scope_views_and_rule_text_search():
    specified = _base_rule()
    global_rule = dict(_base_rule())
    global_rule.update(
        {
            "rule_id": "TSR-GLOBAL-001",
            "customer_code": "",
            "customer_name": "全部客户",
            "global_scope": True,
            "source_text": "全客户订单备注含汽车板时下AC",
        }
    )
    pending_rule = dict(_base_rule())
    pending_rule.update(
        {
            "rule_id": "TSR-PENDING-001",
            "customer_code": "200",
            "customer_name": "待确认客户",
            "review_state": "pending",
            "source_text": "需要业务确认的MINILED规则",
        }
    )
    historical = dict(_base_rule())
    historical.update(
        {
            "rule_id": "TSR-HISTORY-001",
            "customer_code": "300",
            "customer_name": "历史建议客户",
            "historical_suggestion": True,
        }
    )
    rules = [specified, global_rule, pending_rule, historical]

    assert customer_rule_workspace(rules, rule_scope="global")["rule_count"] == 1
    assert customer_rule_workspace(rules, rule_scope="customer")["rule_count"] == 1
    assert customer_rule_workspace(rules, rule_scope="pending")["rule_count"] == 1
    assert customer_rule_workspace(rules, rule_scope="history")["rule_count"] == 1
    # 业务页面默认分层：生效范围只包含全客户和客户确定规则，历史资料不会混入。
    assert customer_rule_workspace(rules, rule_scope="active")["rule_count"] == 2
    assert customer_rule_workspace(rules, rule_scope="reference")["rule_count"] == 1
    searched = customer_rule_workspace(rules, search="MINILED")
    assert searched["selected_customer"]["name"] == "待确认客户"
    assert searched["scope_counts"] == {
        "all": 4,
        "global": 1,
        "customer": 1,
        "pending": 1,
        "technical": 0,
        "migration": 0,
        "reference": 0,
        "history": 1,
    }


def test_customer_workspace_replaces_placeholder_name_using_same_customer_code():
    canonical = _base_rule()
    canonical.update({"customer_code": "103890", "customer_name": "珠海景旺"})
    legacy_placeholder = dict(canonical)
    legacy_placeholder.update(
        {
            "rule_id": "LEGACY::客户下单转换::204::基板级别",
            "customer_name": "????",
            "source_text": "客户胶系NY-A1对应基板级别AC",
            "review_state": "migration",
        }
    )

    workspace = customer_rule_workspace(
        [canonical, legacy_placeholder],
        rule_scope="active",
        customer_key=make_customer_key("103890", "????"),
    )

    assert len(workspace["customers"]) == 1
    assert workspace["selected_customer"]["name"] == "珠海景旺"
    assert workspace["selected_rule"]["customer_name"] == "珠海景旺"


def test_customer_workspace_override_count_only_counts_current_category(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    semantic = _base_rule()
    save_rule_override(semantic, updated_by="cyb", previous_rule=None)
    deterministic = dict(_base_rule())
    deterministic.update(
        {
            "rule_id": "TSR-TEST-DET-COUNT",
            "conditions": [
                {"field": "客户规格", "operator": "contains_any", "value": ["NY-A2"]}
            ],
        }
    )

    semantic_workspace = customer_rule_workspace([semantic, deterministic], rule_kind="semantic")
    deterministic_workspace = customer_rule_workspace(
        [semantic, deterministic], rule_kind="deterministic"
    )

    assert semantic_workspace["override_count"] == 1
    assert deterministic_workspace["override_count"] == 0


def test_restore_new_rule_removes_it_from_runtime(monkeypatch, tmp_path):
    _use_temp_database(monkeypatch, tmp_path)
    rule = _base_rule()
    rule["rule_id"] = "TCR-NEW-ROLLBACK-001"
    save_rule_override(rule, updated_by="cyb", previous_rule=None)
    assert [item["rule_id"] for item in merge_customer_rule_overrides([])] == [rule["rule_id"]]

    add_change = next(item for item in list_customer_rule_changes() if item["action"] == "新增")
    restore_customer_rule_change(add_change["id"], updated_by="cyb")

    assert merge_customer_rule_overrides([]) == []


def test_workspace_includes_history_technical_reference_and_customer_aliases():
    rules = mapping_assets_for_customer_workspace(
        {
            "客户字段映射": [
                {
                    "映射ID": "HISTORY-1",
                    "启用": "是",
                    "客户代码": "100",
                    "客户简称": "南通深南",
                    "条件胶系": "NY2150",
                    "覆盖字段": "grade_code",
                    "覆盖值": "AC",
                    "规则文本": "历史样本建议下AC",
                }
            ],
            "外部尺寸表引用": [
                {
                    "映射ID": "EXT-1",
                    "启用": "是",
                    "客户代码": "100",
                    "客户简称": "南通深南",
                    "引用文件": "客户尺寸表.xlsx",
                    "规则文本": "尺寸参考客户对照表",
                }
            ],
            "待接入规则": [
                {
                    "映射ID": "TECH-1",
                    "启用": "否",
                    "客户代码": "100",
                    "客户简称": "南通深南",
                    "技术类型": "尺寸映射",
                    "原始规则": "36*48转37*49",
                    "建议处理": "接入尺寸覆盖",
                }
            ],
            "客户规则组": [
                {
                    "映射ID": "GROUP-1",
                    "启用": "是",
                    "规则组ID": "SHENNAN",
                    "规则组名称": "深南集团",
                    "客户代码": "100",
                    "客户简称": "南通深南",
                    "主规则客户代码": "100",
                    "主规则客户简称": "南通深南",
                },
                {
                    "映射ID": "GROUP-2",
                    "启用": "是",
                    "规则组ID": "SHENNAN",
                    "规则组名称": "深南集团",
                    "客户代码": "101",
                    "客户简称": "深南电路",
                    "主规则客户代码": "100",
                    "主规则客户简称": "南通深南",
                },
            ],
        },
        include_semantic_pending=False,
    )

    assert sum(rule.get("asset_type") == INFORMATION_ASSET_TYPE for rule in rules) == 3
    assert sum(rule.get("asset_type") == CUSTOMER_METADATA_ASSET_TYPE for rule in rules) == 2
    workspace = customer_rule_workspace(
        rules,
        customer_key=make_customer_key("100", "南通深南"),
    )

    assert workspace["rule_count"] == 3
    assert workspace["scope_counts"]["history"] == 1
    assert workspace["scope_counts"]["technical"] == 1
    assert workspace["scope_counts"]["reference"] == 1
    assert workspace["selected_customer"]["rule_group"] == "深南集团"
    assert workspace["selected_customer"]["aliases"] == ["深南电路（101）"]
    assert all("客户规则组" not in rule["source_text"] for rule in workspace["rules"])


def test_glue_selection_without_selection_condition_is_pending_not_active():
    rules = mapping_assets_for_customer_workspace(
        {
            "Agent胶系选择规则": [
                {
                    "映射ID": "SELECT-NO-CONDITION",
                    "启用": "是",
                    "胶系名称": "NY6300S",
                    "输出胶系代码": "6C",
                    "规则文本": "NY6300S存在多个候选代码",
                }
            ]
        },
        include_semantic_pending=False,
    )

    workspace = customer_rule_workspace(rules, rule_scope="pending")
    selected = workspace["selected_rule"]
    assert workspace["rule_count"] == 1
    assert selected["scope_key"] == "pending"
    assert selected["status_label"] == "缺选择条件/待业务确认"
    assert "选择条件" in selected["status_detail"]


def test_active_assets_project_all_required_read_only_categories():
    from fangzheng_web_app import create_app
    from fangzheng_web_app.transcode_agent_rules import load_transcode_agent_mapping_tables

    app = create_app()
    with app.app_context():
        rules = mapping_assets_for_customer_workspace(load_transcode_agent_mapping_tables())

    assert sum(rule.get("review_state") == "history" for rule in rules) >= 244
    assert sum(rule.get("review_state") == "technical" for rule in rules) >= 15
    assert sum(rule.get("review_state") == "pending" and rule.get("model") == "订单备注语义待确认" for rule in rules) >= 9
    assert sum(rule.get("review_state") == "reference" for rule in rules) >= 1
    assert sum(bool(rule.get("customer_metadata")) for rule in rules) >= 4
    assert all(
        rule.get("editable") is False
        for rule in rules
        if rule.get("review_state") in {"history", "technical", "reference", "migration"}
    )
    assert sum(rule.get("review_state") == "migration" for rule in rules) > 0
    assert any(rule.get("asset_type") == LEGACY_CCL_ASSET_TYPE for rule in rules)
    assert any(rule.get("asset_type") == CUSTOMER_ORDER_ASSET_TYPE for rule in rules)
    assert any(rule.get("asset_type") == CODE_MIGRATION_ASSET_TYPE for rule in rules)


def _write_legacy_rule_workbook(path: Path) -> None:
    workbook = openpyxl.Workbook()
    special = workbook.active
    special.title = "特殊需求"
    special.append(["客户代码", "客户简称", "物料类别（1为PP 2为基板）", "特殊需求"])
    special.append(["100", "CCL客户", 2, "NY2150板厚0.8mm，不含铜，汽车板"])
    special.append(["200", "PP客户", 1, "PP尺寸和厚度特殊规则"])
    order = workbook.create_sheet("客户下单与胶系基板转换")
    order.append(["客户编号", "客户简称", "客户胶系", "厂内胶系代码", "基板等级代码", "结构代码"])
    order.append(["100", "CCL客户", "NY2150", "2B", "AC", "Z"])
    workbook.save(path)


def test_legacy_runtime_projection_includes_ccl_and_customer_order_but_ignores_pp(tmp_path):
    workbook_path = tmp_path / "transcode_rules.xlsx"
    _write_legacy_rule_workbook(workbook_path)

    rules = legacy_rule_assets_for_customer_workspace(workbook_path)

    ccl_rules = [rule for rule in rules if rule["asset_type"] == LEGACY_CCL_ASSET_TYPE]
    order_rules = [rule for rule in rules if rule["asset_type"] == CUSTOMER_ORDER_ASSET_TYPE]
    assert ccl_rules
    assert {rule["business_field"] for rule in ccl_rules} >= {
        "胶系",
        "基板厚度",
        "基板级别",
        "总/芯厚",
    }
    assert len(order_rules) == 2
    assert {rule["business_field"] for rule in order_rules} == {"胶系", "基板级别"}
    assert all(rule["customer_name"] != "PP客户" for rule in rules)
    assert all(rule["review_state"] == "migration" for rule in rules)
    assert all(rule["priority"] == 0 and rule["editable"] is False for rule in rules)
    assert all("结构" not in rule["business_field"] for rule in rules)


def test_complete_projection_function_combines_all_sources_and_isolates_migration(tmp_path):
    workbook_path = tmp_path / "transcode_rules.xlsx"
    _write_legacy_rule_workbook(workbook_path)
    semantic = _base_rule()
    agent = {
        "规则ID": "AGENT-TEST-001",
        "启用": "是",
        "客户代码": "100",
        "客户简称": "CCL客户",
        "原始字段": "基板尺寸",
        "规则文本": "客户尺寸37*49",
        "覆盖字段": "size_code",
        "覆盖值": "37004900",
        "优先级": "100",
        "待确认": "否",
    }

    rules = project_customer_rule_assets_for_workspace(
        [semantic],
        [agent],
        {},
        base_workbook_path=workbook_path,
        include_semantic_pending=False,
    )
    workspace = customer_rule_workspace(rules, rule_scope="migration")

    assert any(rule.get("rule_id") == semantic["rule_id"] for rule in rules)
    assert any(rule.get("rule_id") == "AGENT-TEST-001" for rule in rules)
    assert any(rule.get("asset_type") == CODE_MIGRATION_ASSET_TYPE for rule in rules)
    assert workspace["rule_count"] > 0
    assert workspace["scope_counts"]["migration"] == workspace["rule_count"]
    assert all(rule["scope_key"] == "migration" for rule in workspace["rules"])


def test_historical_samples_never_become_formal_rules():
    rules = mapping_assets_for_customer_workspace(
        {
            "客户字段映射": [
                {
                    "映射ID": "HISTORY-NO-FORMAL",
                    "启用": "是",
                    "客户代码": "100",
                    "客户简称": "历史客户",
                    "覆盖字段": "grade_code",
                    "覆盖值": "AC",
                    "规则文本": "历史正确码样本建议",
                }
            ]
        },
        include_semantic_pending=False,
        include_runtime_legacy=False,
    )
    rule = rules[0]
    workspace = customer_rule_workspace(rules, rule_scope="history")

    assert rule["historical_suggestion"] is True
    assert rule["priority"] == 0
    assert rule["editable"] is False
    assert workspace["selected_rule"]["status_label"] == "历史样本建议"
