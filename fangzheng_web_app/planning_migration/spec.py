from __future__ import annotations


TABLE_COLUMNS = {
    "task_categories": (
        "id", "employee_id", "name", "short_label", "sort_order", "created_at", "updated_at",
    ),
    "personal_tasks": (
        "id", "employee_id", "category_id", "title", "description", "task_tag", "priority",
        "progress", "due_date", "archived_at", "sort_order", "priority_sort_order", "created_at",
        "updated_at",
    ),
    "feedback": (
        "id", "employee_id", "feedback_type", "feature", "material_desc", "system_result",
        "expected_result", "content", "daily_workload", "error_probability", "urgency", "status",
        "admin_note", "created_at", "updated_at",
    ),
}
TABLES = tuple(TABLE_COLUMNS)
PRIMARY_KEYS = {table: ("id",) for table in TABLES}
