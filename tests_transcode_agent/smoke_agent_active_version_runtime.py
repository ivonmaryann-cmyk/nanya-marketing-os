from __future__ import annotations

from pathlib import Path
import sys
from tempfile import TemporaryDirectory

from werkzeug.datastructures import FileStorage


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import db
from fangzheng_web_app import transcode_agent_rules as agent_rules
from fangzheng_web_app import transcode_rules
from fangzheng_web_app.transcode_agent_service import calculate_transcode_agent_quote


DRAFT_PATH = ROOT / "docs/develop0707/客户特殊规则结构化草稿_按原表行_20260708.xlsx"
EXPECTED_RULE_COUNT = 200


def main() -> None:
    assert DRAFT_PATH.exists(), DRAFT_PATH

    with TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        original_db_path = db.DATABASE_PATH
        original_agent_rules_dir = agent_rules.TRANSCODE_AGENT_RULES_DIR
        original_agent_versions_dir = agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR
        original_base_versions_dir = transcode_rules.TRANSCODE_RULES_VERSIONS_DIR
        try:
            db.DATABASE_PATH = temp_dir / "storage/app.db"
            db.DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
            db.init_db()

            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR = temp_dir / "rules/transcode/versions"
            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR.mkdir(parents=True, exist_ok=True)
            base_version = transcode_rules.ensure_default_transcode_rule_version()
            assert transcode_rules.get_transcode_rule_file_path(base_version).exists()

            agent_rules.TRANSCODE_AGENT_RULES_DIR = temp_dir / "rules/transcode_agent"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = agent_rules.TRANSCODE_AGENT_RULES_DIR / "versions"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR.mkdir(parents=True, exist_ok=True)

            with DRAFT_PATH.open("rb") as source:
                upload = FileStorage(
                    stream=source,
                    filename=DRAFT_PATH.name,
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                agent_version = agent_rules.save_new_transcode_agent_rule_version(
                    upload,
                    updated_by="smoke",
                    remark="active runtime smoke",
                )

            assert agent_rules.get_active_transcode_agent_rule_version() == agent_version
            assert len(agent_rules.load_transcode_agent_rules(agent_version)) == EXPECTED_RULE_COUNT

            quote = calculate_transcode_agent_quote(
                "NY3150HC 0.8mm 1/1 37*49 HVLP",
                customer="珠海景旺",
                customer_code="103890",
            )
            assert quote["agent_rule_version"] == agent_version, quote
            assert quote["rule_version"] == base_version, quote
            assert quote["status"] in {"成功", "待确认"}, quote
            assert _has_agent_override(quote, "glue_code", "RV"), quote
            assert _has_agent_override(quote, "copper_type_code", "O"), quote

            f7_quote = calculate_transcode_agent_quote(
                "NY2150 0.6mm 1/1 37*49 HS2-M2-VSP",
                customer="方正F7",
                customer_code="103891",
            )
            assert f7_quote["agent_rule_version"] == agent_version, f7_quote
            assert _has_agent_override(f7_quote, "glue_code", "2B"), f7_quote
            assert _has_agent_override(f7_quote, "tc_code", "C"), f7_quote
            assert _has_agent_override(f7_quote, "copper_type_code", "P"), f7_quote
            assert not any(item.get("field") == "结构码" and item.get("hit_type") == "Agent规则覆盖" for item in f7_quote["field_evidence"])

            print(
                "active runtime smoke passed "
                f"base_version={base_version} agent_version={agent_version} "
                f"rules={EXPECTED_RULE_COUNT} quote_status={quote['status']} f7_status={f7_quote['status']}"
            )
        finally:
            db.DATABASE_PATH = original_db_path
            agent_rules.TRANSCODE_AGENT_RULES_DIR = original_agent_rules_dir
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = original_agent_versions_dir
            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR = original_base_versions_dir


def _has_agent_override(quote: dict, field: str, value: str) -> bool:
    field_labels = {
        "glue_code": "胶系",
        "thickness_code": "厚度",
        "copper_code": "铜厚",
        "size_code": "尺寸",
        "glue_category_code": "胶水类别",
        "copper_type_code": "铜箔类型",
        "grade_code": "基板级别",
        "tc_code": "总/芯厚",
    }
    label = field_labels[field]
    return any(
        item.get("field") == label
        and item.get("code") == value
        and item.get("hit_type") == "Agent规则覆盖"
        for item in quote.get("field_evidence", [])
    )


if __name__ == "__main__":
    main()
