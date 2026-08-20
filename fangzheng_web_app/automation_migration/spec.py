from __future__ import annotations


TABLES = (
    "mail_accounts", "mail_messages", "mail_attachments", "mail_attachment_texts",
    "mail_fetch_logs", "mail_fetch_tasks", "mail_fetch_task_messages", "mail_order_tasks",
    "mail_transcode_jobs", "order_intake_cases", "order_intake_case_events",
    "order_mail_routing_rules", "order_mail_routing_rule_events", "order_mail_rule_groups",
    "order_mail_rule_keywords", "order_change_tags", "order_change_tag_keywords",
    "order_entry_templates", "order_entry_template_lines", "order_entry_template_versions",
)

PRIMARY_KEYS = {table: ("id",) for table in TABLES}
PRIMARY_KEYS.update({
    "mail_attachment_texts": ("attachment_id",),
    "mail_fetch_task_messages": ("fetch_task_id", "mail_id"),
})

SENSITIVE_COLUMNS = {"mail_accounts": {"auth_code_ciphertext"}}
ATTACHMENT_COLUMNS = {"mail_messages": ("eml_path",), "mail_attachments": ("stored_path", "size_bytes", "sha256")}
