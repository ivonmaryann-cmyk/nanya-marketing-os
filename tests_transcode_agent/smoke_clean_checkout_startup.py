from __future__ import annotations

from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import create_app
from fangzheng_web_app.transcode_agent_rules import get_active_transcode_agent_rule_version, load_transcode_agent_rules
from fangzheng_web_app.transcode_semantic_rules import get_active_transcode_semantic_rule_version, load_transcode_semantic_rules


def main() -> None:
    app = create_app()
    assert app is not None
    agent_version = get_active_transcode_agent_rule_version()
    semantic_version = get_active_transcode_semantic_rule_version()
    loaded_agent_rules = load_transcode_agent_rules(agent_version)
    assert len(loaded_agent_rules) == 207
    assert any(rule.get("规则ID") == "TAR-CYB-20260721-001" for rule in loaded_agent_rules)
    loaded_semantic_rules = load_transcode_semantic_rules(semantic_version)
    assert len(loaded_semantic_rules) == 76
    assert any(rule.get("rule_id") == "TSR-20260724-023" for rule in loaded_semantic_rules)
    print(f"clean checkout startup smoke passed agent={agent_version} semantic={semantic_version}")


if __name__ == "__main__":
    main()
