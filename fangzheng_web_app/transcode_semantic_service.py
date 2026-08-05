from __future__ import annotations

import json
import os
import re
import ssl
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from .paths import PROJECT_DIR


MODEL_SKILL_DIR = PROJECT_DIR / "model_skills" / "marketing-transcode-semantics"
SEMANTIC_PROMPT_PATH = MODEL_SKILL_DIR / "references" / "semantic-normalizer-prompt.txt"
SEMANTIC_SCHEMA_PATH = MODEL_SKILL_DIR / "references" / "semantic-normalizer.schema.json"
EVIDENCE_PROMPT_PATH = MODEL_SKILL_DIR / "references" / "evidence-reviewer-prompt.txt"
EVIDENCE_SCHEMA_PATH = MODEL_SKILL_DIR / "references" / "evidence-review.schema.json"
PROJECT_MODEL_CONFIG_FILE = PROJECT_DIR / "config" / "model.local.yaml"
LEGACY_DESKTOP_KEY_FILE = Path.home() / "Desktop" / "模型.txt"
DEFAULT_KEY_FILE = PROJECT_MODEL_CONFIG_FILE

SEMANTIC_TASK_TYPES = {"rule_structure", "order_normalization"}
MODEL_MODES = {"off", "shadow", "active"}
MODEL_CONFIDENCE_VALUES = {"high", "medium", "low"}
SEMANTIC_TYPES = {
    "keyword_present",
    "keyword_absent",
    "default_when_missing",
    "exclusion_set",
    "position_code",
    "alias_mapping",
    "explicit_fact",
    "comparison",
    "multi_condition",
    "external_lookup",
    "out_of_scope",
}
TARGET_FIELDS = {
    "glue",
    "thickness",
    "copper",
    "size",
    "glue_category",
    "copper_type",
    "print_mark",
    "grade_intent",
    "total_core",
    "combination_structure",
    "copper_vendor",
    "glass_vendor",
    "formula_code",
    "unknown",
}
CONDITION_OPERATORS = {
    "contains_any",
    "contains_all",
    "not_contains",
    "equals",
    "not_equals",
    "in",
    "not_in",
    "lt",
    "lte",
    "gt",
    "gte",
    "char_at",
    "missing",
    "present",
}
REVIEW_FIELDS = {
    "glue",
    "thickness",
    "copper",
    "size",
    "glue_category",
    "copper_type",
    "grade",
    "total_core",
}
REVIEW_VERDICTS = {"supported", "contradicted", "ambiguous", "missing_evidence"}


class SemanticModelError(RuntimeError):
    pass


class SemanticModelConfigError(SemanticModelError):
    pass


class SemanticModelResponseError(SemanticModelError):
    pass


@dataclass(frozen=True)
class SemanticModelConfig:
    api_key: str = field(repr=False)
    base_url: str = "https://api.deepseek.com"
    model: str = "deepseek-chat"
    mode: str = "off"
    timeout_seconds: float = 120.0
    max_order_calls: int = 50
    key_file: Path = DEFAULT_KEY_FILE

    @property
    def enabled(self) -> bool:
        return self.mode in {"shadow", "active"}


Transport = Callable[[str, dict[str, str], dict[str, Any], float], dict[str, Any]]


def load_semantic_model_config(
    key_file: str | Path | None = None,
    *,
    environ: dict[str, str] | None = None,
) -> SemanticModelConfig:
    env = os.environ if environ is None else environ
    configured_path = env.get("DEEPSEEK_API_KEY_FILE") or key_file or DEFAULT_KEY_FILE
    secret_path = Path(configured_path).expanduser()
    if not secret_path.exists() and key_file is None and not env.get("DEEPSEEK_API_KEY_FILE"):
        secret_path = LEGACY_DESKTOP_KEY_FILE
    file_values = _read_simple_llm_config(secret_path) if secret_path.exists() else {}

    api_key = str(env.get("DEEPSEEK_API_KEY") or file_values.get("api_key") or "").strip()
    base_url = str(env.get("DEEPSEEK_BASE_URL") or file_values.get("base_url") or "https://api.deepseek.com").strip()
    model = str(env.get("DEEPSEEK_MODEL") or file_values.get("model") or "deepseek-chat").strip()
    mode = str(
        env.get("TRANSCODE_SEMANTIC_MODEL_MODE")
        or file_values.get("mode")
        or "off"
    ).strip().lower()
    timeout_raw = str(
        env.get("TRANSCODE_SEMANTIC_MODEL_TIMEOUT")
        or file_values.get("timeout_seconds")
        or "120"
    ).strip()
    max_order_calls_raw = str(
        env.get("TRANSCODE_ORDER_SEMANTIC_MODEL_MAX_CALLS")
        or file_values.get("max_order_calls")
        or "50"
    ).strip()

    if not api_key:
        raise SemanticModelConfigError(
            "DeepSeek API key is missing. Set DEEPSEEK_API_KEY or DEEPSEEK_API_KEY_FILE."
        )
    if not base_url.startswith("https://"):
        raise SemanticModelConfigError("DeepSeek base_url must use HTTPS.")
    if not model:
        raise SemanticModelConfigError("DeepSeek model is missing.")
    if mode not in MODEL_MODES:
        raise SemanticModelConfigError(
            f"TRANSCODE_SEMANTIC_MODEL_MODE must be one of {sorted(MODEL_MODES)}."
        )
    try:
        timeout_seconds = float(timeout_raw)
    except ValueError as exc:
        raise SemanticModelConfigError("TRANSCODE_SEMANTIC_MODEL_TIMEOUT must be numeric.") from exc
    if timeout_seconds <= 0:
        raise SemanticModelConfigError("TRANSCODE_SEMANTIC_MODEL_TIMEOUT must be positive.")
    try:
        max_order_calls = int(max_order_calls_raw)
    except ValueError as exc:
        raise SemanticModelConfigError(
            "TRANSCODE_ORDER_SEMANTIC_MODEL_MAX_CALLS must be an integer."
        ) from exc
    if max_order_calls < 0:
        raise SemanticModelConfigError(
            "TRANSCODE_ORDER_SEMANTIC_MODEL_MAX_CALLS must not be negative."
        )

    return SemanticModelConfig(
        api_key=api_key,
        base_url=base_url.rstrip("/"),
        model=model,
        mode=mode,
        timeout_seconds=timeout_seconds,
        max_order_calls=min(max_order_calls, 500),
        key_file=secret_path,
    )


class DeepSeekSemanticClient:
    def __init__(
        self,
        config: SemanticModelConfig,
        *,
        transport: Transport | None = None,
    ) -> None:
        self.config = config
        self._transport = transport or _post_json

    def normalize(
        self,
        *,
        task_type: str,
        source_fields: dict[str, Any],
        customer_code: str = "",
        customer_name: str = "",
        relevant_rules: list[dict[str, Any]] | None = None,
        task_context: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        if task_type not in SEMANTIC_TASK_TYPES:
            raise ValueError(f"Unsupported semantic task_type: {task_type}")
        clean_sources = _clean_source_fields(source_fields)
        if not clean_sources:
            raise ValueError("At least one non-empty source field is required.")
        result = self._complete_json(
            system_prompt=_read_text(SEMANTIC_PROMPT_PATH),
            schema=_read_json(SEMANTIC_SCHEMA_PATH),
            task_payload={
                "task_type": task_type,
                "material_scope": "CCL",
                "customer": {
                    "customer_code": str(customer_code or "").strip(),
                    "customer_name": str(customer_name or "").strip(),
                },
                "source_fields": clean_sources,
                "task_context": task_context or {},
                "relevant_approved_rules": relevant_rules or [],
            },
        )
        if task_type == "order_normalization":
            _clear_unstated_order_targets(result, clean_sources)
        validate_semantic_result(result, clean_sources, expected_task_type=task_type)
        return result

    def review_evidence(
        self,
        *,
        source_fields: dict[str, Any],
        normalized_semantics: dict[str, Any],
        candidate_fields: dict[str, Any],
        field_evidence: list[dict[str, Any]],
        relevant_rules: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        clean_sources = _clean_source_fields(source_fields)
        if not clean_sources:
            raise ValueError("At least one non-empty source field is required.")
        result = self._complete_json(
            system_prompt=_read_text(EVIDENCE_PROMPT_PATH),
            schema=_read_json(EVIDENCE_SCHEMA_PATH),
            task_payload={
                "material_scope": "CCL",
                "source_fields": clean_sources,
                "normalized_semantics": normalized_semantics,
                "candidate_fields": candidate_fields,
                "field_evidence": field_evidence,
                "relevant_approved_rules": relevant_rules or [],
            },
        )
        validate_evidence_review(result, clean_sources)
        return result

    def _complete_json(
        self,
        *,
        system_prompt: str,
        schema: dict[str, Any],
        task_payload: dict[str, Any],
    ) -> dict[str, Any]:
        request_body = {
            "model": self.config.model,
            "temperature": 0,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": json.dumps(
                        {"input": task_payload, "output_schema": schema},
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                },
            ],
        }
        headers = {
            "Authorization": f"Bearer {self.config.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        endpoint = f"{self.config.base_url}/chat/completions"
        response = self._transport(endpoint, headers, request_body, self.config.timeout_seconds)
        try:
            content = response["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError) as exc:
            raise SemanticModelResponseError("DeepSeek response has no message content.") from exc
        result = _parse_json_content(content)
        if not isinstance(result, dict):
            raise SemanticModelResponseError("Model response must be a JSON object.")
        return result


def _clear_unstated_order_targets(
    result: dict[str, Any],
    source_fields: dict[str, str],
) -> None:
    """Remove target codes copied from approved rules but absent from order text."""
    for item in result.get("semantic_items") or []:
        if not isinstance(item, dict):
            continue
        stated_target = str(item.get("stated_target_value") or "").strip()
        if stated_target and not _text_present_in_sources(stated_target, source_fields):
            item["stated_target_value"] = ""


def validate_semantic_result(
    result: dict[str, Any],
    source_fields: dict[str, str],
    *,
    expected_task_type: str | None = None,
) -> None:
    expected_keys = {
        "schema_version",
        "task_type",
        "material_scope",
        "semantic_items",
        "ambiguities",
        "missing_inputs",
        "model_confidence",
    }
    _require_exact_keys(result, expected_keys, "semantic result")
    if result["schema_version"] != "1.0":
        raise SemanticModelResponseError("Unsupported semantic schema_version.")
    if result["task_type"] not in SEMANTIC_TASK_TYPES:
        raise SemanticModelResponseError("Invalid semantic task_type.")
    if expected_task_type and result["task_type"] != expected_task_type:
        raise SemanticModelResponseError("Semantic task_type does not match the request.")
    if result["material_scope"] not in {"CCL", "OUT_OF_SCOPE", "UNKNOWN"}:
        raise SemanticModelResponseError("Invalid material_scope.")
    if result["model_confidence"] not in MODEL_CONFIDENCE_VALUES:
        raise SemanticModelResponseError("Invalid model_confidence.")
    if not isinstance(result["semantic_items"], list):
        raise SemanticModelResponseError("semantic_items must be a list.")
    if not isinstance(result["ambiguities"], list) or not isinstance(result["missing_inputs"], list):
        raise SemanticModelResponseError("ambiguities and missing_inputs must be lists.")

    item_keys = {
        "semantic_type",
        "target_field",
        "normalized_value",
        "stated_target_value",
        "conditions",
        "source_field",
        "evidence_text",
        "confidence",
        "deterministic_preferred",
    }
    condition_keys = {"field", "operator", "value", "source_scope"}
    for index, item in enumerate(result["semantic_items"]):
        if not isinstance(item, dict):
            raise SemanticModelResponseError(f"semantic_items[{index}] must be an object.")
        _require_exact_keys(item, item_keys, f"semantic_items[{index}]")
        if item["semantic_type"] not in SEMANTIC_TYPES:
            raise SemanticModelResponseError(f"Invalid semantic_type at item {index}.")
        if item["target_field"] not in TARGET_FIELDS:
            raise SemanticModelResponseError(f"Invalid target_field at item {index}.")
        if item["confidence"] not in MODEL_CONFIDENCE_VALUES:
            raise SemanticModelResponseError(f"Invalid confidence at item {index}.")
        if not isinstance(item["deterministic_preferred"], bool):
            raise SemanticModelResponseError(f"deterministic_preferred must be boolean at item {index}.")
        _validate_evidence_span(item["source_field"], item["evidence_text"], source_fields, f"item {index}")
        stated_target = str(item["stated_target_value"] or "").strip()
        if stated_target and not _text_present_in_sources(stated_target, source_fields):
            raise SemanticModelResponseError(
                f"stated_target_value at item {index} is not present in source text."
            )
        if not isinstance(item["conditions"], list):
            raise SemanticModelResponseError(f"conditions must be a list at item {index}.")
        for condition_index, condition in enumerate(item["conditions"]):
            if not isinstance(condition, dict):
                raise SemanticModelResponseError(
                    f"condition {condition_index} at item {index} must be an object."
                )
            _require_exact_keys(
                condition,
                condition_keys,
                f"semantic_items[{index}].conditions[{condition_index}]",
            )
            if condition["operator"] not in CONDITION_OPERATORS:
                raise SemanticModelResponseError(
                    f"Invalid condition operator at item {index}, condition {condition_index}."
                )

    ambiguity_keys = {"field", "reason", "evidence_text"}
    for index, item in enumerate(result["ambiguities"]):
        if not isinstance(item, dict):
            raise SemanticModelResponseError(f"ambiguities[{index}] must be an object.")
        _require_exact_keys(item, ambiguity_keys, f"ambiguities[{index}]")
        evidence_text = str(item["evidence_text"] or "").strip()
        if evidence_text and not _text_present_in_sources(evidence_text, source_fields):
            raise SemanticModelResponseError(f"Ambiguity evidence {index} is not present in source text.")

    missing_keys = {"field", "reason"}
    for index, item in enumerate(result["missing_inputs"]):
        if not isinstance(item, dict):
            raise SemanticModelResponseError(f"missing_inputs[{index}] must be an object.")
        _require_exact_keys(item, missing_keys, f"missing_inputs[{index}]")


def validate_evidence_review(result: dict[str, Any], source_fields: dict[str, str]) -> None:
    _require_exact_keys(
        result,
        {"schema_version", "field_reviews", "hard_blockers", "model_confidence"},
        "evidence review",
    )
    if result["schema_version"] != "1.0":
        raise SemanticModelResponseError("Unsupported evidence schema_version.")
    if result["model_confidence"] not in MODEL_CONFIDENCE_VALUES:
        raise SemanticModelResponseError("Invalid evidence model_confidence.")
    if not isinstance(result["field_reviews"], list) or not isinstance(result["hard_blockers"], list):
        raise SemanticModelResponseError("field_reviews and hard_blockers must be lists.")
    review_keys = {"field", "verdict", "source_field", "evidence_text", "reason"}
    for index, item in enumerate(result["field_reviews"]):
        if not isinstance(item, dict):
            raise SemanticModelResponseError(f"field_reviews[{index}] must be an object.")
        _require_exact_keys(item, review_keys, f"field_reviews[{index}]")
        if item["field"] not in REVIEW_FIELDS:
            raise SemanticModelResponseError(f"Invalid review field at item {index}.")
        if item["verdict"] not in REVIEW_VERDICTS:
            raise SemanticModelResponseError(f"Invalid review verdict at item {index}.")
        evidence_text = str(item["evidence_text"] or "").strip()
        if item["verdict"] in {"supported", "contradicted"}:
            _validate_evidence_span(item["source_field"], evidence_text, source_fields, f"review {index}")
        elif evidence_text and not _text_present_in_sources(evidence_text, source_fields):
            raise SemanticModelResponseError(f"Review evidence {index} is not present in source text.")


def _read_simple_llm_config(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("#") or stripped == "llm:":
            continue
        match = re.match(
            r"^(api_key|key|base_url|model|mode|timeout_seconds|max_order_calls)\s*:\s*(.+?)\s*$",
            stripped,
        )
        if not match:
            continue
        key, value = match.groups()
        values["api_key" if key == "key" else key] = value.strip().strip("'\"")
    return values


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except OSError as exc:
        raise SemanticModelConfigError(f"Required model resource is missing: {path}") from exc


def _read_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(_read_text(path))
    except json.JSONDecodeError as exc:
        raise SemanticModelConfigError(f"Invalid JSON schema: {path}") from exc
    if not isinstance(value, dict):
        raise SemanticModelConfigError(f"JSON schema must be an object: {path}")
    return value


def _clean_source_fields(source_fields: dict[str, Any]) -> dict[str, str]:
    cleaned: dict[str, str] = {}
    for key, value in source_fields.items():
        name = str(key or "").strip()
        text = str(value or "").strip()
        if name and text and text.lower() not in {"nan", "none"}:
            cleaned[name] = text
    return cleaned


def _parse_json_content(content: Any) -> dict[str, Any]:
    text = str(content or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    try:
        result = json.loads(text)
    except json.JSONDecodeError as exc:
        raise SemanticModelResponseError("Model response is not valid JSON.") from exc
    if not isinstance(result, dict):
        raise SemanticModelResponseError("Model response must be a JSON object.")
    return result


def _require_exact_keys(value: dict[str, Any], expected: set[str], label: str) -> None:
    actual = set(value)
    if actual != expected:
        missing = sorted(expected - actual)
        unexpected = sorted(actual - expected)
        raise SemanticModelResponseError(
            f"Invalid {label} keys; missing={missing}, unexpected={unexpected}."
        )


def _validate_evidence_span(
    source_field: Any,
    evidence_text: Any,
    source_fields: dict[str, str],
    label: str,
) -> None:
    field_name = str(source_field or "").strip()
    evidence = str(evidence_text or "").strip()
    if field_name not in source_fields:
        matching_fields = [name for name, source in source_fields.items() if evidence and evidence in source]
        if len(matching_fields) == 1:
            return
        raise SemanticModelResponseError(f"Unknown source_field for {label}: {field_name}")
    if not evidence or evidence not in source_fields[field_name]:
        raise SemanticModelResponseError(f"Evidence for {label} is not a continuous source span.")


def _text_present_in_sources(text: str, source_fields: dict[str, str]) -> bool:
    needle = str(text or "").strip()
    return bool(needle) and any(needle in source for source in source_fields.values())


def _post_json(
    url: str,
    headers: dict[str, str],
    payload: dict[str, Any],
    timeout_seconds: float,
) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    ca_file = os.environ.get("DEEPSEEK_CA_FILE") or os.environ.get("SSL_CERT_FILE")
    if not ca_file:
        default_paths = ssl.get_default_verify_paths()
        system_ca = Path("/etc/ssl/cert.pem")
        if not default_paths.cafile and system_ca.exists():
            ca_file = str(system_ca)
    context = ssl.create_default_context(cafile=ca_file) if ca_file else ssl.create_default_context()
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds, context=context) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise SemanticModelError(f"DeepSeek HTTP {exc.code}: {detail}") from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise SemanticModelError(f"DeepSeek request failed: {exc}") from exc
    try:
        result = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise SemanticModelResponseError("DeepSeek API returned invalid JSON.") from exc
    if not isinstance(result, dict):
        raise SemanticModelResponseError("DeepSeek API response must be a JSON object.")
    return result
