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


def main() -> None:
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
            transcode_rules.ensure_default_transcode_rule_version()

            agent_rules.TRANSCODE_AGENT_RULES_DIR = temp_dir / "rules/transcode_agent"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = agent_rules.TRANSCODE_AGENT_RULES_DIR / "versions"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR.mkdir(parents=True, exist_ok=True)

            with DRAFT_PATH.open("rb") as source:
                upload = FileStorage(
                    stream=source,
                    filename=DRAFT_PATH.name,
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                version = agent_rules.save_new_transcode_agent_rule_version(
                    upload,
                    updated_by="smoke",
                    remark="size mapping application smoke",
                )

            mapping_tables = agent_rules.load_transcode_agent_mapping_tables(version)
            assert len(mapping_tables["客户尺寸映射"]) == 48
            assert len(mapping_tables["客户单边尺寸映射"]) == 24
            assert len(mapping_tables["客户尺寸算法"]) == 1
            assert len(mapping_tables["外部尺寸表引用"]) == 1
            assert mapping_tables["外部尺寸表引用"][0]["启用"] == "是"

            _assert_size_mapping(
                "南通深南",
                "104370",
                "NY2150 0.8mm 1/1 36*48 HTE",
                "37004900",
            )
            _assert_size_mapping(
                "惠州特创",
                "103613",
                "NY2150 0.8mm 1/1 37*49 HTE",
                "37304930",
            )
            _assert_size_mapping(
                "安徽万奔",
                "133038",
                "NY2150 0.8mm 1/1 49*41 HTE",
                "41304930",
            )
            _assert_size_mapping(
                "江福昌发",
                "103031",
                "NY2140 0.8mm 1/1 1888*1049 HTE",
                "74304130",
            )
            _assert_size_mapping(
                "无新美亚",
                "104443",
                "NY2150 0.8mm 1/1 14*24 HTE 631",
                "14132421",
            )
            _assert_size_mapping(
                "无新美亚",
                "104443",
                "NY2150 0.8mm 1/1 14*24 HTE 632",
                "14022402",
            )

            print("size mapping application smoke passed")
        finally:
            db.DATABASE_PATH = original_db_path
            agent_rules.TRANSCODE_AGENT_RULES_DIR = original_agent_rules_dir
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = original_agent_versions_dir
            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR = original_base_versions_dir


def _assert_size_mapping(customer: str, customer_code: str, spec: str, expected_code: str) -> None:
    quote = calculate_transcode_agent_quote(spec, customer=customer, customer_code=customer_code)
    size_evidence = next(item for item in quote["field_evidence"] if item["field"] == "尺寸")
    assert size_evidence["code"] == expected_code, (customer, spec, expected_code, size_evidence, quote)
    assert size_evidence["hit_type"] == "Agent尺寸映射", (customer, spec, size_evidence, quote)
    assert size_evidence["score"] == 100, (customer, spec, size_evidence, quote)


if __name__ == "__main__":
    main()
