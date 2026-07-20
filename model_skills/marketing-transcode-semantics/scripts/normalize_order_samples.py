from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys


PROJECT_DIR = Path(__file__).resolve().parents[3]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from fangzheng_web_app.transcode_semantic_service import (  # noqa: E402
    DeepSeekSemanticClient,
    load_semantic_model_config,
)


DEFAULT_INPUT = (
    PROJECT_DIR
    / "model_skills/marketing-transcode-semantics/references/order-semantic-samples-20260714.json"
)


def main() -> None:
    parser = argparse.ArgumentParser(description="批量执行营销转码订单语义标准化")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    env = dict(os.environ)
    env["TRANSCODE_SEMANTIC_MODEL_MODE"] = "active"
    config = load_semantic_model_config(environ=env)
    client = DeepSeekSemanticClient(config)

    results = []
    for sample in payload.get("samples", []):
        result = client.normalize(
            task_type="order_normalization",
            source_fields={"客户规格": sample["spec"]},
            customer_code=sample.get("customer_code", ""),
            customer_name=sample.get("customer_name", ""),
            relevant_rules=[
                {
                    "scope": "已批准订单语义提示",
                    "text": text,
                }
                for text in sample.get("approved_semantic_hints", [])
            ],
            task_context={
                "purpose": "只标准化八个CCL字段，不输出制造码",
                "approved_semantic_hints": sample.get("approved_semantic_hints", []),
            },
        )
        results.append(
            {
                "row": sample.get("row"),
                "customer_code": sample.get("customer_code", ""),
                "customer_name": sample.get("customer_name", ""),
                "spec": sample.get("spec", ""),
                "model": config.model,
                "normalization": result,
            }
        )

    output = {
        "schema_version": "1.0",
        "model": config.model,
        "sample_count": len(results),
        "results": results,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"model": config.model, "sample_count": len(results)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
