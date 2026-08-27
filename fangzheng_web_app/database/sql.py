from __future__ import annotations

import re


_IDENTITY_TABLES = {
    "mail_accounts", "mail_messages", "mail_attachments", "mail_fetch_logs", "mail_fetch_tasks",
    "mail_order_tasks", "mail_transcode_jobs", "order_intake_cases", "order_intake_case_events",
    "order_mail_routing_rules", "order_mail_routing_rule_events", "order_mail_rule_groups",
    "order_mail_rule_keywords", "order_change_tags", "order_change_tag_keywords",
    "order_entry_templates", "order_entry_template_lines", "order_entry_template_versions",
    "order_entry_template_tasks", "order_interface_configs", "order_interface_config_versions",
    "order_entry_detail_events", "order_interface_call_logs", "order_material_query_suggestions",
    "order_material_resolution_tasks",
    "automation_customers", "automation_customer_contacts", "automation_customer_routing_rules",
    "automation_customer_routing_conditions", "automation_customer_extraction_maps",
    "automation_customer_events", "automation_customer_spec_mappings",
    "jobs", "pp_transcode_base_rules", "pp_transcode_customer_rules",
    "pp_transcode_rule_changes", "pp_transcode_confirmation_items",
    "transcode_agent_confirmation_items", "transcode_agent_confirmation_events",
    "transcode_agent_pending_rules", "transcode_agent_row_verifications",
    "transcode_customer_rule_changes", "transcode_rule_center_changes",
}


def qmark_to_pyformat(sql: str) -> str:
    """Convert qmark parameters outside SQL strings/comments to psycopg placeholders."""
    output: list[str] = []
    index = 0
    state = "code"
    while index < len(sql):
        char = sql[index]
        following = sql[index + 1] if index + 1 < len(sql) else ""
        if state == "code":
            if char == "'":
                state = "single"
            elif char == '"':
                state = "double"
            elif char == "-" and following == "-":
                state = "line_comment"
            elif char == "/" and following == "*":
                state = "block_comment"
            elif char == "?":
                output.append("%s")
                index += 1
                continue
        elif state == "single" and char == "'":
            if following == "'":
                output.extend((char, following))
                index += 2
                continue
            state = "code"
        elif state == "double" and char == '"':
            if following == '"':
                output.extend((char, following))
                index += 2
                continue
            state = "code"
        elif state == "line_comment" and char in "\r\n":
            state = "code"
        elif state == "block_comment" and char == "*" and following == "/":
            output.extend((char, following))
            index += 2
            state = "code"
            continue
        output.append(char)
        index += 1
    return "".join(output)


def sqlite_to_postgresql(sql: str, *, map_automation_metadata: bool = True) -> tuple[str, bool]:
    """Translate the small, audited SQLite dialect surface used by automation."""
    statement = qmark_to_pyformat(sql)
    if re.fullmatch(r"\s*BEGIN\s+IMMEDIATE\s*;?\s*", statement, re.IGNORECASE):
        return "SELECT 1", False
    if map_automation_metadata:
        statement = re.sub(r"\bsettings\b", "automation_metadata", statement, flags=re.IGNORECASE)
    statement = re.sub(
        r"GROUP_CONCAT\(([^,()]+),\s*('(?:[^']|'')*')\)",
        r"STRING_AGG(\1, \2)", statement, flags=re.IGNORECASE,
    )

    replace_match = re.match(r"\s*INSERT\s+OR\s+REPLACE\s+INTO\s+mail_fetch_task_messages\b", statement, re.IGNORECASE)
    if replace_match:
        statement = re.sub(r"INSERT\s+OR\s+REPLACE", "INSERT", statement, count=1, flags=re.IGNORECASE)
        statement += " ON CONFLICT (fetch_task_id, mail_id) DO UPDATE SET is_new=EXCLUDED.is_new, created_at=EXCLUDED.created_at"

    ignore_match = re.match(r"\s*INSERT\s+OR\s+IGNORE\s+INTO\s+order_intake_cases\b", statement, re.IGNORECASE)
    if ignore_match:
        statement = re.sub(r"INSERT\s+OR\s+IGNORE", "INSERT", statement, count=1, flags=re.IGNORECASE)
        statement += " ON CONFLICT (employee_id, mail_id) DO NOTHING"

    pp_ignore_match = re.match(
        r"\s*INSERT\s+OR\s+IGNORE\s+INTO\s+pp_transcode_base_rules\b",
        statement,
        re.IGNORECASE,
    )
    if pp_ignore_match:
        statement = re.sub(r"INSERT\s+OR\s+IGNORE", "INSERT", statement, count=1, flags=re.IGNORECASE)
        statement += " ON CONFLICT (field_key, input_value) DO NOTHING"

    statement = re.sub(
        r"datetime\(COALESCE\(([^)]+)\)\)\s*<\s*datetime\((%s)\)",
        r"(COALESCE(\1))::timestamp < (\2)::timestamp",
        statement,
        flags=re.IGNORECASE,
    )

    insert_match = re.match(r"\s*INSERT\s+INTO\s+([a-zA-Z_][a-zA-Z0-9_]*)\b", statement, re.IGNORECASE)
    returns_identity = bool(
        insert_match
        and insert_match.group(1).lower() in _IDENTITY_TABLES
        and not re.search(r"\bRETURNING\b", statement, re.IGNORECASE)
    )
    if returns_identity:
        statement += " RETURNING id"
    return statement, returns_identity
