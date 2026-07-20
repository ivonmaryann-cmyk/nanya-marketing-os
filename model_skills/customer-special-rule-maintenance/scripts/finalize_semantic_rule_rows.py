from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[3]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from fangzheng_web_app.transcode_semantic_rule_finalizer import (
    build_atomic_semantic_rule_rows,
    load_atomic_overrides,
    validate_atomic_semantic_rule_rows,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="将模型审计JSON生成一条语义项一行的正式规则行JSON")
    parser.add_argument("input", type=Path, help="客户特殊规则模型JSON化结果")
    parser.add_argument("output", type=Path, help="原子规则行JSON输出路径")
    parser.add_argument("--approval-basis", required=True, help="业务确认依据")
    parser.add_argument("--note", default="P2-2B原子化清洗；仅影子运行", help="正式规则备注")
    parser.add_argument("--overrides", type=Path, help="原子化结构修正JSON；默认读取Skill references内配置")
    args = parser.parse_args()

    audit_payload = json.loads(args.input.read_text(encoding="utf-8"))
    rules, pending = build_atomic_semantic_rule_rows(
        audit_payload,
        approval_basis=args.approval_basis,
        note=args.note,
        overrides=load_atomic_overrides(args.overrides) if args.overrides else None,
    )
    validate_atomic_semantic_rule_rows(rules)
    payload = {
        "schema_version": "1.0",
        "source_file": str(args.input),
        "model": audit_payload.get("model", ""),
        "prompt_sha256": audit_payload.get("prompt_sha256", ""),
        "rule_count": len(rules),
        "source_candidate_count": len({row["来源候选ID"] for row in rules}),
        "pending_count": len(pending),
        "rules": rules,
        "pending_candidate_ids": [item.get("candidate_id", "") for item in pending],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "rule_count": payload["rule_count"],
                "source_candidate_count": payload["source_candidate_count"],
                "pending_count": payload["pending_count"],
                "output": str(args.output),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
