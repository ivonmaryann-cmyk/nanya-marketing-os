from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[3]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from fangzheng_web_app.transcode_semantic_rules import (
    get_active_transcode_semantic_rule_version,
    publish_transcode_semantic_rule_version,
    validate_transcode_semantic_rule_version,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="发布业务已确认的营销转码模型语义规则版本")
    parser.add_argument("input", type=Path, help="最终模型语义规则xlsx")
    parser.add_argument("--updated-by", required=True, help="发布人")
    parser.add_argument("--approval-basis", required=True, help="业务确认依据")
    parser.add_argument("--remark", default="", help="版本备注")
    parser.add_argument("--no-activate", action="store_true", help="只生成版本，不设为active")
    args = parser.parse_args()

    version = publish_transcode_semantic_rule_version(
        args.input,
        updated_by=args.updated_by,
        approval_basis=args.approval_basis,
        remark=args.remark,
        activate=not args.no_activate,
    )
    manifest = validate_transcode_semantic_rule_version(version)
    print(
        json.dumps(
            {
                "version": version,
                "active_version": get_active_transcode_semantic_rule_version(),
                "rule_count": manifest["rule_count"],
                "pending_count": manifest["pending_count"],
                "runtime_mode": "shadow",
                "runtime_effect": "只输出影子证据，不覆盖编码和评分",
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
