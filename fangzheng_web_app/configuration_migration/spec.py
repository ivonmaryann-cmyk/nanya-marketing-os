from __future__ import annotations


TABLE_COLUMNS = {
    "settings": ("key", "value"),
    "pdf_excel_ai_config_versions": (
        "id", "enabled", "base_url", "model", "timeout_seconds", "max_rows",
        "api_key_ciphertext", "repair_instruction", "rebuild_instruction",
        "header_mapping_instruction", "config_fingerprint", "prompt_digest",
        "test_status", "test_message", "tested_at", "created_by", "created_at",
        "source_version_id", "activated_by", "activated_at",
    ),
}
TABLES = tuple(TABLE_COLUMNS)
PRIMARY_KEYS = {"settings": ("key",), "pdf_excel_ai_config_versions": ("id",)}
