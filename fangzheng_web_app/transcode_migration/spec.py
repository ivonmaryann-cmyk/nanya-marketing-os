from __future__ import annotations


TABLE_COLUMNS = {
    "jobs": ("id", "feature", "employee_id", "source_filename", "stored_input_path", "stored_result_path", "status", "success_count", "fail_count", "skip_count", "current_row", "total_rows", "worker_pid", "worker_started_at", "rule_version", "log_text", "error_message", "created_at", "completed_at", "confirm_count", "verify_count"),
    "pp_transcode_base_rules": ("id", "field_key", "input_value", "output_value", "business_note", "enabled", "source_kind", "created_by", "created_at", "updated_by", "updated_at"),
    "pp_transcode_customer_rules": ("id", "customer_code", "customer_name", "target_field", "conditions_json", "output_value", "business_note", "enabled", "created_by", "created_at", "updated_by", "updated_at"),
    "pp_transcode_rule_changes": ("id", "rule_scope", "rule_id", "action", "before_json", "after_json", "changed_by", "changed_at"),
    "pp_transcode_confirmation_items": ("id", "job_id", "employee_id", "excel_row", "customer_code", "customer_name", "spec", "order_remark", "pending_code", "confirmed_pending_code", "confidence", "summary", "field_evidence_json", "confirmation_status", "confirmation_basis", "confirmed_by", "confirmed_at", "created_at", "updated_at"),
    "transcode_model_configs": ("employee_id", "enabled", "base_url", "api_key", "model", "timeout_seconds", "max_order_calls", "updated_at"),
    "transcode_agent_confirmation_items": ("id", "job_id", "employee_id", "excel_row", "customer_code", "customer_name", "spec", "context_text", "field_key", "field_label", "current_code", "options_json", "pending_code", "score", "reason", "evidence_json", "analysis_json", "status", "confirmed_code", "confirmation_basis", "confirmed_by", "confirmed_at", "long_term_rule_id", "created_at", "updated_at", "pending_rule_id"),
    "transcode_agent_confirmation_events": ("id", "item_id", "job_id", "employee_id", "action", "before_json", "after_json", "created_at"),
    "transcode_agent_pending_rules": ("id", "rule_id", "rule_json", "employee_id", "customer_code", "customer_name", "business_field", "target_value", "condition_summary", "source_task_id", "source_excel_row", "status", "created_at", "updated_at", "updated_by", "processed_by", "processed_at"),
    "transcode_agent_row_verifications": ("id", "job_id", "excel_row", "employee_id", "action", "before_code", "after_code", "basis", "created_at"),
    "transcode_agent_rule_overrides": ("rule_id", "rule_json", "deleted", "updated_by", "updated_at"),
    "transcode_customer_rule_changes": ("id", "rule_id", "action", "employee_id", "before_json", "after_json", "created_at"),
    "transcode_customer_rule_overrides": ("rule_id", "rule_json", "deleted", "updated_by", "updated_at"),
    "transcode_rule_center_asset_overrides": ("asset_group", "row_id", "row_json", "deleted", "updated_by", "updated_at"),
    "transcode_rule_center_base_overrides": ("rule_id", "rule_json", "deleted", "updated_by", "updated_at"),
    "transcode_rule_center_changes": ("id", "category", "object_id", "action", "employee_id", "before_json", "after_json", "created_at"),
    "transcode_rule_center_confirmation_overrides": ("rule_id", "rule_json", "deleted", "updated_by", "updated_at"),
    "transcode_rule_center_lookup_overrides": ("group_key", "input_value", "output_value", "deleted", "updated_by", "updated_at"),
}

TABLES = tuple(TABLE_COLUMNS)
PRIMARY_KEYS = {table: ("id",) for table in TABLES}
PRIMARY_KEYS.update({
    "transcode_model_configs": ("employee_id",),
    "transcode_agent_rule_overrides": ("rule_id",),
    "transcode_customer_rule_overrides": ("rule_id",),
    "transcode_rule_center_asset_overrides": ("asset_group", "row_id"),
    "transcode_rule_center_base_overrides": ("rule_id",),
    "transcode_rule_center_confirmation_overrides": ("rule_id",),
    "transcode_rule_center_lookup_overrides": ("group_key", "input_value"),
})
IDENTITY_TABLES = tuple(table for table, keys in PRIMARY_KEYS.items() if keys == ("id",))
