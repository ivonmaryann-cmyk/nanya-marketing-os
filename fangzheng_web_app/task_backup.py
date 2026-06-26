from __future__ import annotations

import json
import os
from pathlib import Path

from .db import db_cursor, utcnow


BACKUP_ROOT = Path(os.environ.get("WORK_PLANNING_BACKUP_ROOT", r"D:\Carson\tmp\work_planning_backups"))
BACKUP_SCHEMA_VERSION = 1


def _backup_dir(employee_id: str) -> Path:
    return BACKUP_ROOT / employee_id


def latest_backup_path(employee_id: str) -> Path:
    return _backup_dir(employee_id) / "latest.json"


def get_task_backup_status(employee_id: str) -> dict:
    path = latest_backup_path(employee_id)
    if not path.exists():
        return {"exists": False, "path": str(path), "saved_at": "", "task_count": 0, "category_count": 0}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"exists": False, "path": str(path), "saved_at": "", "task_count": 0, "category_count": 0}
    return {
        "exists": True,
        "path": str(path),
        "saved_at": str(data.get("saved_at") or ""),
        "task_count": len(data.get("tasks") or []),
        "category_count": len(data.get("categories") or []),
    }


def save_task_backup(employee_id: str) -> dict:
    with db_cursor() as conn:
        categories = conn.execute(
            """
            SELECT id, name, short_label, sort_order, created_at, updated_at
            FROM task_categories
            WHERE employee_id = ?
            ORDER BY sort_order, id
            """,
            (employee_id,),
        ).fetchall()
        tasks = conn.execute(
            """
            SELECT
                t.title,
                t.description,
                t.task_tag,
                t.priority,
                t.progress,
                t.due_date,
                t.archived_at,
                t.sort_order,
                t.priority_sort_order,
                t.created_at,
                t.updated_at,
                c.name AS category_name
            FROM personal_tasks t
            LEFT JOIN task_categories c ON c.id = t.category_id AND c.employee_id = t.employee_id
            WHERE t.employee_id = ?
            ORDER BY COALESCE(c.sort_order, 999999), c.id, t.sort_order, t.id
            """,
            (employee_id,),
        ).fetchall()

    payload = {
        "schema_version": BACKUP_SCHEMA_VERSION,
        "employee_id": employee_id,
        "saved_at": utcnow(),
        "categories": [
            {
                "name": row["name"],
                "short_label": row["short_label"],
                "sort_order": row["sort_order"],
                "created_at": row["created_at"],
                "updated_at": row["updated_at"],
            }
            for row in categories
        ],
        "tasks": [
            {
                "title": row["title"],
                "description": row["description"],
                "task_tag": row["task_tag"],
                "priority": row["priority"],
                "progress": row["progress"],
                "due_date": row["due_date"],
                "archived_at": row["archived_at"],
                "sort_order": row["sort_order"],
                "priority_sort_order": row["priority_sort_order"],
                "created_at": row["created_at"],
                "updated_at": row["updated_at"],
                "category_name": row["category_name"],
            }
            for row in tasks
        ],
    }
    backup_dir = _backup_dir(employee_id)
    backup_dir.mkdir(parents=True, exist_ok=True)
    path = latest_backup_path(employee_id)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return get_task_backup_status(employee_id)


def restore_task_backup(employee_id: str) -> dict:
    path = latest_backup_path(employee_id)
    if not path.exists():
        raise FileNotFoundError(str(path))
    data = json.loads(path.read_text(encoding="utf-8"))
    categories = data.get("categories") or []
    tasks = data.get("tasks") or []
    now = utcnow()
    category_id_by_name: dict[str, int] = {}

    with db_cursor() as conn:
        conn.execute("DELETE FROM personal_tasks WHERE employee_id = ?", (employee_id,))
        conn.execute("DELETE FROM task_categories WHERE employee_id = ?", (employee_id,))

        for index, category in enumerate(categories, start=1):
            name = str(category.get("name") or "").strip()
            if not name or name in category_id_by_name:
                continue
            cursor = conn.execute(
                """
                INSERT INTO task_categories (
                    employee_id, name, short_label, sort_order, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    employee_id,
                    name,
                    str(category.get("short_label") or "").strip()[:12],
                    int(category.get("sort_order") or index),
                    category.get("created_at") or now,
                    category.get("updated_at") or now,
                ),
            )
            category_id_by_name[name] = int(cursor.lastrowid)

        for task in tasks:
            title = str(task.get("title") or "").strip()
            if not title:
                continue
            progress = str(task.get("progress") or "not_started")
            if progress == "waiting":
                progress = "not_started"
            category_name = str(task.get("category_name") or "").strip()
            conn.execute(
                """
                INSERT INTO personal_tasks (
                    employee_id, category_id, title, description, task_tag, priority, progress,
                    due_date, archived_at, sort_order, priority_sort_order, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    employee_id,
                    category_id_by_name.get(category_name),
                    title,
                    str(task.get("description") or ""),
                    str(task.get("task_tag") or ""),
                    str(task.get("priority") or "normal"),
                    progress,
                    task.get("due_date") or None,
                    task.get("archived_at") or None,
                    int(task.get("sort_order") or 0),
                    int(task.get("priority_sort_order") or 0),
                    task.get("created_at") or now,
                    task.get("updated_at") or now,
                ),
            )

    return get_task_backup_status(employee_id)
