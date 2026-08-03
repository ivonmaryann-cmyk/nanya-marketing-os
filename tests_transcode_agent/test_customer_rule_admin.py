from __future__ import annotations

import pytest
from werkzeug.datastructures import MultiDict

from fangzheng_web_app import db
from fangzheng_web_app.transcode_customer_rule_admin import (
    CONDITION_OPERATORS,
    CONDITION_OPERATOR_LABELS,
    CustomerRuleMaintenanceError,
    build_rule_from_form,
    customer_rule_workspace,
    delete_rule_override,
    list_customer_rule_changes,
    make_customer_key,
    merge_customer_rule_overrides,
    restore_customer_rule_change,
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


def test_all_condition_operators_have_chinese_labels():
    assert set(CONDITION_OPERATOR_LABELS) == set(CONDITION_OPERATORS)
    assert all(CONDITION_OPERATOR_LABELS[operator] for operator in CONDITION_OPERATORS)


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
