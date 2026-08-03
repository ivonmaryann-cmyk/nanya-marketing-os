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
                    remark="thickness mapping application smoke",
                )

            mapping_tables = agent_rules.load_transcode_agent_mapping_tables(version)
            assert len(mapping_tables["客户厚度映射"]) == 2
            assert all(row["启用"] == "是" for row in mapping_tables["客户厚度映射"])
            assert all(row["启用"] == "是" for row in mapping_tables["客户物料编码口径"])

            alias_quote = calculate_transcode_agent_quote(
                "NY2150 003 1/1 37*49 HTE",
                customer="江苏瀚宇",
                customer_code="104158",
            )
            thickness = next(item for item in alias_quote["field_evidence"] if item["field"] == "厚度")
            assert thickness["code"] == "00079", (thickness, alias_quote)
            assert thickness["hit_type"] == "Agent厚度映射", (thickness, alias_quote)
            assert thickness["score"] == 100, (thickness, alias_quote)
            alias_total_core = next(item for item in alias_quote["field_evidence"] if item["field"] == "总/芯厚")
            assert alias_total_core["code"] == "C", (alias_total_core, alias_quote)
            assert alias_total_core["hit_type"] == "Agent总芯厚映射", (alias_total_core, alias_quote)

            total_quote = calculate_transcode_agent_quote(
                "NY2150 43mil 1/1 37*49 HTE",
                customer="江苏瀚宇",
                customer_code="104158",
            )
            total_core = next(item for item in total_quote["field_evidence"] if item["field"] == "总/芯厚")
            assert total_core["code"] == "T", (total_core, total_quote)
            assert total_core["hit_type"] == "Agent总芯厚映射", (total_core, total_quote)

            material_core_quote = calculate_transcode_agent_quote(
                "NY2150 0.8mm 1/1 37*49 HTE 631",
                customer="无新美亚",
                customer_code="104443",
            )
            material_core = next(item for item in material_core_quote["field_evidence"] if item["field"] == "总/芯厚")
            assert material_core["code"] == "C", (material_core, material_core_quote)
            assert material_core["hit_type"] == "Agent物料编码口径", (material_core, material_core_quote)

            material_total_quote = calculate_transcode_agent_quote(
                "NY2150 0.8mm 1/1 37*49 HTE 632",
                customer="无新美亚",
                customer_code="104443",
            )
            material_total = next(item for item in material_total_quote["field_evidence"] if item["field"] == "总/芯厚")
            assert material_total["code"] == "T", (material_total, material_total_quote)
            assert material_total["hit_type"] == "Agent物料编码口径", (material_total, material_total_quote)

            other_customer_quote = calculate_transcode_agent_quote(
                "NY2150 0.8mm 1/1 37*49 HTE 631",
                customer="珠海景旺",
                customer_code="103890",
            )
            other_customer_total_core = next(item for item in other_customer_quote["field_evidence"] if item["field"] == "总/芯厚")
            assert other_customer_total_core["hit_type"] != "Agent物料编码口径", (other_customer_total_core, other_customer_quote)

            print("thickness mapping application smoke passed")
        finally:
            db.DATABASE_PATH = original_db_path
            agent_rules.TRANSCODE_AGENT_RULES_DIR = original_agent_rules_dir
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = original_agent_versions_dir
            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR = original_base_versions_dir


if __name__ == "__main__":
    main()
