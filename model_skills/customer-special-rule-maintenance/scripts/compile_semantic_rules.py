from __future__ import annotations

import argparse
import hashlib
import json
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any


PROJECT_DIR = Path(__file__).resolve().parents[3]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from fangzheng_web_app.transcode_semantic_rule_compiler import (
    compile_semantic_candidate,
    evaluate_semantic_compilation,
)
from fangzheng_web_app.transcode_semantic_rule_maintenance import classify_draft_workbook
from fangzheng_web_app.transcode_semantic_service import (
    SEMANTIC_PROMPT_PATH,
    SEMANTIC_SCHEMA_PATH,
    DeepSeekSemanticClient,
    SemanticModelError,
    load_semantic_model_config,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="将客户特殊规则模型候选批量转换为受控JSON")
    parser.add_argument("input", type=Path, help="结构化草稿xlsx路径")
    parser.add_argument("--output", type=Path, required=True, help="编译结果JSON路径")
    parser.add_argument("--workers", type=int, default=3, help="并发数，默认3")
    parser.add_argument("--retries", type=int, default=1, help="失败后重试次数，默认1")
    parser.add_argument("--candidate-id", action="append", help="只处理指定候选ID，可重复")
    parser.add_argument("--resume", action="store_true", help="复用输出文件中程序校验通过的结果")
    args = parser.parse_args()

    classification = classify_draft_workbook(args.input)
    candidates = classification["candidates"]
    requested = set(args.candidate_id or [])
    if requested:
        candidates = [item for item in candidates if item["candidate_id"] in requested]
    if not candidates:
        raise SystemExit("没有需要处理的模型语义候选")
    if args.workers < 1 or args.workers > 8:
        raise SystemExit("workers必须在1到8之间")

    sibling_groups: dict[tuple[str, str, str], list[str]] = {}
    for candidate in classification["candidates"]:
        key = (
            str(candidate.get("customer_code") or ""),
            str(candidate.get("customer_name") or ""),
            str(candidate.get("business_field") or ""),
        )
        sibling_groups.setdefault(key, []).append(str(candidate.get("source_text") or ""))
    for candidate in candidates:
        key = (
            str(candidate.get("customer_code") or ""),
            str(candidate.get("customer_name") or ""),
            str(candidate.get("business_field") or ""),
        )
        source_text = str(candidate.get("source_text") or "")
        candidate["related_candidate_texts"] = [
            value for value in sibling_groups.get(key, []) if value and value != source_text
        ]

    config = load_semantic_model_config()
    client = DeepSeekSemanticClient(config)
    output_items: dict[str, dict[str, Any]] = {}
    existing_items: dict[str, dict[str, Any]] = {}
    if args.resume and args.output.exists():
        existing = json.loads(args.output.read_text(encoding="utf-8"))
        for item in existing.get("items", []):
            existing_items[str(item.get("candidate_id") or "")] = item
            if item.get("status") == "程序校验通过":
                output_items[str(item.get("candidate_id") or "")] = item

    candidate_by_id = {item["candidate_id"]: item for item in candidates}
    for candidate_id, existing_item in existing_items.items():
        if candidate_id in output_items or candidate_id not in candidate_by_id:
            continue
        raw_result = existing_item.get("raw_model_result") or {}
        if raw_result:
            revalidated = evaluate_semantic_compilation(
                candidate_by_id[candidate_id], raw_result
            ).to_dict()
            revalidated["attempts"] = existing_item.get("attempts", 1)
            output_items[candidate_id] = revalidated

    pending = [item for item in candidates if item["candidate_id"] not in output_items]
    write_lock = threading.Lock()

    def compile_one(candidate: dict[str, Any]) -> dict[str, Any]:
        last_error = ""
        for attempt in range(1, args.retries + 2):
            try:
                result = compile_semantic_candidate(candidate, client).to_dict()
                result["attempts"] = attempt
                return result
            except (SemanticModelError, ValueError) as exc:
                last_error = str(exc)
        return {
            **{
                key: candidate.get(key, "")
                for key in [
                    "candidate_id",
                    "customer_code",
                    "customer_name",
                    "source_row",
                    "business_field",
                    "source_text",
                    "required_input_fields",
                ]
            },
            "status": "程序校验失败",
            "recommended_execution_mode": "无法解析",
            "validation_result": f"模型调用失败：{last_error}",
            "business_question": "",
            "model_confidence": "low",
            "semantic_type_summary": "",
            "target_field_summary": "",
            "model_result": {},
            "attempts": args.retries + 1,
        }

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(compile_one, candidate): candidate for candidate in pending}
        for future in as_completed(futures):
            result = future.result()
            with write_lock:
                output_items[result["candidate_id"]] = result
                _write_output(args.output, args.input, config.model, candidates, output_items)
            print(
                f"{result['candidate_id']} {result['status']} "
                f"{result['recommended_execution_mode']}"
            )

    _write_output(args.output, args.input, config.model, candidates, output_items)
    print(json.dumps(_build_summary(output_items.values()), ensure_ascii=False))


def _write_output(
    output_path: Path,
    source_path: Path,
    model: str,
    candidates: list[dict[str, Any]],
    output_items: dict[str, dict[str, Any]],
) -> None:
    ordered = [
        output_items[item["candidate_id"]]
        for item in candidates
        if item["candidate_id"] in output_items
    ]
    payload = {
        "schema_version": "1.0",
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source_file": str(source_path.resolve()),
        "model_provider": "DeepSeek",
        "model": model,
        "prompt_sha256": _sha256(SEMANTIC_PROMPT_PATH),
        "schema_sha256": _sha256(SEMANTIC_SCHEMA_PATH),
        "candidate_count": len(candidates),
        "completed_count": len(ordered),
        "summary": _build_summary(ordered),
        "items": ordered,
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_path.replace(output_path)


def _build_summary(items) -> dict[str, dict[str, int]]:
    status_counts: dict[str, int] = {}
    mode_counts: dict[str, int] = {}
    for item in items:
        status = str(item.get("status") or "").strip()
        mode = str(item.get("recommended_execution_mode") or "").strip()
        if status:
            status_counts[status] = status_counts.get(status, 0) + 1
        if mode:
            mode_counts[mode] = mode_counts.get(mode, 0) + 1
    return {"status_counts": status_counts, "execution_mode_counts": mode_counts}


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


if __name__ == "__main__":
    main()
