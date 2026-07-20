from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[3]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from fangzheng_web_app.transcode_semantic_rule_maintenance import classify_draft_workbook


def main() -> None:
    parser = argparse.ArgumentParser(description="生成客户特殊规则标准/模型分流JSON")
    parser.add_argument("input", type=Path, help="结构化草稿xlsx路径")
    parser.add_argument("--output", type=Path, help="JSON输出路径；不填写则输出到终端")
    args = parser.parse_args()

    result = classify_draft_workbook(args.input)
    content = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(content, encoding="utf-8")
        print(args.output)
    else:
        print(content)


if __name__ == "__main__":
    main()
