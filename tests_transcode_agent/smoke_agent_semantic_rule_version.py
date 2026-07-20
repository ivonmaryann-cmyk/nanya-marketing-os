from __future__ import annotations

import sys
from pathlib import Path
from tempfile import TemporaryDirectory


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import db
from fangzheng_web_app import transcode_semantic_rules as semantic_rules


RULE_WORKBOOK = ROOT / "docs/develop0707/营销转码Agent最终模型语义规则表_20260710.xlsx"


def main() -> None:
    assert RULE_WORKBOOK.exists(), RULE_WORKBOOK
    rules, summary = semantic_rules.parse_semantic_rule_workbook(RULE_WORKBOOK)
    assert len(rules) == 51
    assert summary["pending_count"] == 3
    assert summary["models"] == ["deepseek-v4-pro"]
    assert len({rule["rule_id"] for rule in rules}) == 51
    assert len({rule["source_candidate_id"] for rule in rules}) == 39
    assert all(rule["execution_mode"] == "结构化后可确定性执行" for rule in rules)

    with TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        original_db_path = db.DATABASE_PATH
        original_rules_dir = semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR
        original_versions_dir = semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR
        try:
            db.DATABASE_PATH = temp_dir / "storage/app.db"
            db.DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
            db.init_db()
            semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR = temp_dir / "rules/transcode_semantic"
            semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR = (
                semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR / "versions"
            )
            semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR.mkdir(parents=True, exist_ok=True)

            first = semantic_rules.publish_transcode_semantic_rule_version(
                RULE_WORKBOOK,
                updated_by="smoke",
                approval_basis="P2-1C smoke approval",
                remark="first semantic version",
            )
            assert semantic_rules.get_active_transcode_semantic_rule_version() == first
            assert len(semantic_rules.load_transcode_semantic_rules(first)) == 51
            first_manifest = semantic_rules.validate_transcode_semantic_rule_version(first)
            assert first_manifest["pending_count"] == 3

            second = semantic_rules.publish_transcode_semantic_rule_version(
                RULE_WORKBOOK,
                updated_by="smoke",
                approval_basis="P2-1C smoke approval",
                remark="second semantic version",
            )
            assert second != first
            assert semantic_rules.get_active_transcode_semantic_rule_version() == second
            assert len(semantic_rules.get_transcode_semantic_rule_history()) == 2

            semantic_rules.activate_transcode_semantic_rule_version(first)
            assert semantic_rules.get_active_transcode_semantic_rule_version() == first
            assert len(semantic_rules.load_transcode_semantic_rules()) == 51

            machine_path = (
                semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR
                / second
                / semantic_rules.MACHINE_RULE_FILENAME
            )
            machine_path.write_text(machine_path.read_text(encoding="utf-8") + "\n", encoding="utf-8")
            try:
                semantic_rules.validate_transcode_semantic_rule_version(second)
            except semantic_rules.SemanticRuleAssetError as exc:
                assert "哈希不匹配" in str(exc)
            else:
                raise AssertionError("Tampered semantic rule version must be rejected")

            print(
                "semantic rule version smoke passed "
                f"rules={len(rules)} pending={summary['pending_count']} "
                f"first={first} second={second} rollback={first}"
            )
        finally:
            db.DATABASE_PATH = original_db_path
            semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR = original_rules_dir
            semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR = original_versions_dir


if __name__ == "__main__":
    main()
