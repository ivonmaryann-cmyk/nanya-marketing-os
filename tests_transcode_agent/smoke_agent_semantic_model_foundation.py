from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_semantic_service import (
    EVIDENCE_PROMPT_PATH,
    EVIDENCE_SCHEMA_PATH,
    MODEL_SKILL_DIR,
    PROJECT_MODEL_CONFIG_FILE,
    SEMANTIC_PROMPT_PATH,
    SEMANTIC_SCHEMA_PATH,
    DeepSeekSemanticClient,
    SemanticModelConfig,
    SemanticModelResponseError,
    load_semantic_model_config,
    validate_semantic_result,
)
from fangzheng_web_app.transcode_agent_service import _detect_semantic_input_columns
from fangzheng_web_app.transcode_engine import (
    build_context_text_from_row,
    detect_transcode_context_columns,
)


def main() -> None:
    test_skill_resources_exist()
    test_secret_file_config_loader()
    test_rule_semantic_normalization()
    test_order_normalization_clears_unstated_rule_target()
    test_direct_code_output_is_rejected()
    test_optional_order_remark_column_is_reserved()
    print("semantic model foundation smoke passed")


def test_skill_resources_exist() -> None:
    assert (MODEL_SKILL_DIR / "SKILL.md").exists()
    for path in [SEMANTIC_PROMPT_PATH, SEMANTIC_SCHEMA_PATH, EVIDENCE_PROMPT_PATH, EVIDENCE_SCHEMA_PATH]:
        assert path.exists(), path
    json.loads(SEMANTIC_SCHEMA_PATH.read_text(encoding="utf-8"))
    json.loads(EVIDENCE_SCHEMA_PATH.read_text(encoding="utf-8"))
    assert (ROOT / "config/model.example.yaml").exists()


def test_secret_file_config_loader() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        config_path = Path(tmp_dir) / "model.txt"
        config_path.write_text(
            "llm:\n"
            "  api_key: \"secret-test-value\"\n"
            "  base_url: \"https://api.deepseek.example\"\n"
            "  model: \"deepseek-test\"\n"
            "  timeout_seconds: 45\n",
            encoding="utf-8",
        )
        config = load_semantic_model_config(
            config_path,
            environ={"TRANSCODE_SEMANTIC_MODEL_MODE": "shadow"},
        )
        assert config.api_key == "secret-test-value"
        assert config.base_url == "https://api.deepseek.example"
        assert config.model == "deepseek-test"
        assert config.mode == "shadow"
        assert config.timeout_seconds == 45
        assert config.enabled
        assert "secret-test-value" not in repr(config)


def test_rule_semantic_normalization() -> None:
    source_text = (
        "基板级别：当备注中有X0A0/X0B0/car/汽车板/458/439字样时=AC，"
        "当备注中有MINILED字样时=AM；"
        "NY1600及NY2140外，其余均下汽车板料号"
    )
    fake_result = {
        "schema_version": "1.0",
        "task_type": "rule_structure",
        "material_scope": "CCL",
        "semantic_items": [
            {
                "semantic_type": "keyword_present",
                "target_field": "grade_intent",
                "normalized_value": "automotive",
                "stated_target_value": "AC",
                "conditions": [
                    {
                        "field": "订单备注",
                        "operator": "contains_any",
                        "value": ["X0A0", "X0B0", "car", "汽车板", "458", "439"],
                        "source_scope": "订单备注",
                    }
                ],
                "source_field": "CCL特殊规则",
                "evidence_text": "当备注中有X0A0/X0B0/car/汽车板/458/439字样时=AC",
                "confidence": "high",
                "deterministic_preferred": False,
            },
            {
                "semantic_type": "exclusion_set",
                "target_field": "grade_intent",
                "normalized_value": "automotive",
                "stated_target_value": "",
                "conditions": [
                    {
                        "field": "glue",
                        "operator": "not_in",
                        "value": ["NY1600", "NY2140"],
                        "source_scope": "客户规格",
                    }
                ],
                "source_field": "CCL特殊规则",
                "evidence_text": "NY1600及NY2140外，其余均下汽车板料号",
                "confidence": "high",
                "deterministic_preferred": True,
            },
        ],
        "ambiguities": [],
        "missing_inputs": [
            {"field": "订单备注", "reason": "执行关键词规则时需要订单备注字段"}
        ],
        "model_confidence": "high",
    }
    captured: dict = {}

    def fake_transport(url, headers, payload, timeout_seconds):
        captured.update(
            {
                "url": url,
                "authorization": headers.get("Authorization"),
                "payload": payload,
                "timeout": timeout_seconds,
            }
        )
        return {"choices": [{"message": {"content": json.dumps(fake_result, ensure_ascii=False)}}]}

    config = SemanticModelConfig(
        api_key="secret-test-value",
        base_url="https://api.deepseek.example",
        model="deepseek-test",
        mode="shadow",
    )
    client = DeepSeekSemanticClient(config, transport=fake_transport)
    result = client.normalize(
        task_type="rule_structure",
        source_fields={"CCL特殊规则": source_text},
        customer_code="105007",
        customer_name="湖奥士康",
    )

    assert result == fake_result
    assert captured["url"] == "https://api.deepseek.example/chat/completions"
    assert captured["authorization"] == "Bearer secret-test-value"
    assert captured["payload"]["model"] == "deepseek-test"
    assert captured["payload"]["temperature"] == 0
    assert result["semantic_items"][1]["conditions"][0]["operator"] == "not_in"


def test_direct_code_output_is_rejected() -> None:
    invalid = {
        "schema_version": "1.0",
        "task_type": "order_normalization",
        "material_scope": "CCL",
        "semantic_items": [],
        "ambiguities": [],
        "missing_inputs": [],
        "model_confidence": "high",
        "final_code": "2B015001137004900YWA1T",
    }
    try:
        validate_semantic_result(
            invalid,
            {"客户规格": "NY2150 1.5mm 1/1 37*49"},
            expected_task_type="order_normalization",
        )
    except SemanticModelResponseError:
        return
    raise AssertionError("Direct final_code output must be rejected")


def test_order_normalization_clears_unstated_rule_target() -> None:
    fake_result = {
        "schema_version": "1.0",
        "task_type": "order_normalization",
        "material_scope": "CCL",
        "semantic_items": [
            {
                "semantic_type": "keyword_present",
                "target_field": "grade_intent",
                "normalized_value": "automotive",
                "stated_target_value": "AC",
                "conditions": [],
                "source_field": "订单备注",
                "evidence_text": "要汽板",
                "confidence": "high",
                "deterministic_preferred": False,
            }
        ],
        "ambiguities": [],
        "missing_inputs": [],
        "model_confidence": "high",
    }

    def fake_transport(url, headers, payload, timeout_seconds):
        return {"choices": [{"message": {"content": json.dumps(fake_result, ensure_ascii=False)}}]}

    client = DeepSeekSemanticClient(
        SemanticModelConfig(
            api_key="secret-test-value",
            base_url="https://api.deepseek.example",
            model="deepseek-test",
            mode="shadow",
        ),
        transport=fake_transport,
    )
    result = client.normalize(
        task_type="order_normalization",
        source_fields={"订单备注": "要汽板"},
        relevant_rules=[{"stated_target_values": ["AC"]}],
    )
    assert result["semantic_items"][0]["normalized_value"] == "automotive"
    assert result["semantic_items"][0]["stated_target_value"] == ""


def test_optional_order_remark_column_is_reserved() -> None:
    frame = pd.DataFrame(
        [
            ["客户简称", "客户规格", "订单备注"],
            ["测试客户", "NY2150 0.8mm 1/1 37*49", "car / MINILED"],
        ]
    )
    semantic_columns = _detect_semantic_input_columns(frame, 1)
    assert semantic_columns["订单备注"]["indices"] == [2], semantic_columns
    context_columns = detect_transcode_context_columns(frame, 1, 0, None)
    assert context_columns == [2], context_columns
    assert build_context_text_from_row(frame.iloc[1], context_columns) == "car / MINILED"


if __name__ == "__main__":
    main()
