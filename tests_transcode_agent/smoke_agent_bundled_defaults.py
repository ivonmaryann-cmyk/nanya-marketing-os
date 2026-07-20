from __future__ import annotations

from pathlib import Path
import sys
from tempfile import TemporaryDirectory


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import db
from fangzheng_web_app import transcode_agent_rules as agent_rules
from fangzheng_web_app import transcode_semantic_rules as semantic_rules


def main() -> None:
    with TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        original_db = db.DATABASE_PATH
        original_agent_dir = agent_rules.TRANSCODE_AGENT_RULES_DIR
        original_agent_versions = agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR
        original_semantic_dir = semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR
        original_semantic_versions = semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR
        try:
            db.DATABASE_PATH = temp_dir / "storage/app.db"
            db.DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
            db.init_db()
            agent_rules.TRANSCODE_AGENT_RULES_DIR = temp_dir / "rules/transcode_agent"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = agent_rules.TRANSCODE_AGENT_RULES_DIR / "versions"
            semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR = temp_dir / "rules/transcode_semantic"
            semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR = semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR / "versions"

            agent_version = agent_rules.ensure_default_transcode_agent_rule_version()
            semantic_version = semantic_rules.ensure_default_transcode_semantic_rule_version()

            assert agent_version == agent_rules.DEFAULT_AGENT_RULE_VERSION
            assert len(agent_rules.load_transcode_agent_rules(agent_version)) == 200
            assert len(semantic_rules.load_transcode_semantic_rules(semantic_version)) == 51
            assert agent_rules.get_transcode_agent_mapping_table_file_path(agent_version).exists()
            assert semantic_rules.validate_transcode_semantic_rule_version(semantic_version)["rule_count"] == 51
            print(f"bundled defaults smoke passed agent={agent_version} semantic={semantic_version}")
        finally:
            db.DATABASE_PATH = original_db
            agent_rules.TRANSCODE_AGENT_RULES_DIR = original_agent_dir
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = original_agent_versions
            semantic_rules.TRANSCODE_SEMANTIC_RULES_DIR = original_semantic_dir
            semantic_rules.TRANSCODE_SEMANTIC_RULES_VERSIONS_DIR = original_semantic_versions


if __name__ == "__main__":
    main()
