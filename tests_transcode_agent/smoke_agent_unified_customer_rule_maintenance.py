from __future__ import annotations

import sys
from pathlib import Path
from tempfile import TemporaryDirectory


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import db
from fangzheng_web_app.transcode_agent_rules import (
    get_active_transcode_agent_rule_version,
    load_transcode_agent_rules,
)
from fangzheng_web_app.transcode_customer_rule_admin import (
    AGENT_ASSET_TYPE,
    agent_rules_for_customer_workspace,
    build_agent_rule_from_form,
    customer_rule_workspace,
    delete_agent_rule_override,
    list_customer_rule_changes,
    restore_customer_rule_change,
    save_agent_rule_override,
)
from fangzheng_web_app.transcode_semantic_rules import (
    get_active_transcode_semantic_rule_version,
    load_transcode_semantic_rules,
)


def main() -> None:
    agent_version = get_active_transcode_agent_rule_version()
    semantic_version = get_active_transcode_semantic_rule_version()
    base_agent_rules = load_transcode_agent_rules(agent_version, include_maintenance=False)
    base_semantic_rules = load_transcode_semantic_rules(
        semantic_version,
        include_maintenance=False,
    )
    selected = next(rule for rule in base_agent_rules if rule.get("覆盖字段") == "grade_code")
    before_value = selected["覆盖值"]
    changed_value = "AC" if before_value != "AC" else "A1"

    with TemporaryDirectory() as temp_dir_name:
        original_database_path = db.DATABASE_PATH
        db.DATABASE_PATH = Path(temp_dir_name) / "app.db"
        try:
            db.init_db()
            form = {
                "asset_type": AGENT_ASSET_TYPE,
                "customer_code": selected.get("客户代码", ""),
                "customer_name": selected.get("客户简称", ""),
                "business_field": "基板级别",
                "source_text": selected.get("规则文本", ""),
                "agent_condition_glue": selected.get("条件胶系", ""),
                "agent_condition_keyword": selected.get("条件关键词", ""),
                "agent_condition_copper": selected.get("条件铜厚", ""),
                "agent_condition_thickness": selected.get("条件厚度", ""),
                "agent_condition_size": selected.get("条件尺寸", ""),
                "agent_override_field": "grade_code",
                "target_value": changed_value,
                "priority": selected.get("优先级", "100"),
                "enabled": "1",
                "approval_basis": "统一客户规则维护回归",
            }
            maintained = build_agent_rule_from_form(form, existing_rule=selected)
            save_agent_rule_override(
                maintained,
                updated_by="tester",
                previous_rule=selected,
            )
            merged_agent_rules = load_transcode_agent_rules(agent_version)
            merged = next(rule for rule in merged_agent_rules if rule.get("规则ID") == selected["规则ID"])
            assert merged["覆盖值"] == changed_value, merged
            assert merged["命中来源"] == "页面维护Agent确定性长期规则"

            semantic_rules = load_transcode_semantic_rules(semantic_version)
            assert len(semantic_rules) == len(base_semantic_rules)
            workspace_rules = semantic_rules + agent_rules_for_customer_workspace(merged_agent_rules)
            workspace = customer_rule_workspace(workspace_rules)
            assert workspace["rule_count"] == len(base_semantic_rules) + len(base_agent_rules)
            assert workspace["override_count"] == 1

            delete_agent_rule_override(maintained, updated_by="tester")
            after_delete = load_transcode_agent_rules(agent_version)
            assert all(rule.get("规则ID") != selected["规则ID"] for rule in after_delete)
            changes = list_customer_rule_changes()
            assert changes[0]["action"] == "删除"
            restore_customer_rule_change(changes[0]["id"], updated_by="tester")
            restored = next(
                rule
                for rule in load_transcode_agent_rules(agent_version)
                if rule.get("规则ID") == selected["规则ID"]
            )
            assert restored["覆盖值"] == changed_value
        finally:
            db.DATABASE_PATH = original_database_path

    print("unified customer rule maintenance smoke passed base+override+delete+restore")


if __name__ == "__main__":
    main()
