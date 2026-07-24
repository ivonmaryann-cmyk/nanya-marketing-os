from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime

from werkzeug.security import check_password_hash, generate_password_hash

from .paths import DATABASE_PATH


def utcnow() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat()


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@contextmanager
def db_cursor():
    conn = get_connection()
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with db_cursor() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                feature TEXT NOT NULL DEFAULT 'fangzheng',
                employee_id TEXT NOT NULL,
                source_filename TEXT NOT NULL,
                stored_input_path TEXT NOT NULL,
                stored_result_path TEXT,
                status TEXT NOT NULL,
                success_count INTEGER NOT NULL DEFAULT 0,
                fail_count INTEGER NOT NULL DEFAULT 0,
                skip_count INTEGER NOT NULL DEFAULT 0,
                current_row INTEGER NOT NULL DEFAULT 0,
                total_rows INTEGER NOT NULL DEFAULT 0,
                worker_pid INTEGER,
                worker_started_at TEXT,
                rule_version TEXT NOT NULL,
                log_text TEXT NOT NULL DEFAULT '',
                error_message TEXT,
                created_at TEXT NOT NULL,
                completed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS users (
                employee_id TEXT PRIMARY KEY,
                display_name TEXT,
                department TEXT,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                enabled INTEGER NOT NULL DEFAULT 1,
                must_change_password INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS pdf_excel_ai_config_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                enabled INTEGER NOT NULL DEFAULT 0,
                base_url TEXT NOT NULL,
                model TEXT NOT NULL,
                timeout_seconds INTEGER NOT NULL,
                max_rows INTEGER NOT NULL,
                api_key_ciphertext TEXT NOT NULL DEFAULT '',
                repair_instruction TEXT NOT NULL DEFAULT '',
                rebuild_instruction TEXT NOT NULL DEFAULT '',
                header_mapping_instruction TEXT NOT NULL DEFAULT '',
                config_fingerprint TEXT NOT NULL,
                prompt_digest TEXT NOT NULL,
                test_status TEXT NOT NULL DEFAULT '',
                test_message TEXT NOT NULL DEFAULT '',
                tested_at TEXT,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL,
                activated_by TEXT NOT NULL,
                activated_at TEXT NOT NULL,
                source_version_id INTEGER,
                FOREIGN KEY(source_version_id) REFERENCES pdf_excel_ai_config_versions(id)
            );

            CREATE TABLE IF NOT EXISTS feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                feedback_type TEXT NOT NULL,
                feature TEXT,
                material_desc TEXT,
                system_result TEXT,
                expected_result TEXT,
                content TEXT,
                daily_workload TEXT,
                error_probability TEXT,
                urgency TEXT,
                status TEXT NOT NULL DEFAULT '待处理',
                admin_note TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS task_categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                name TEXT NOT NULL,
                short_label TEXT NOT NULL DEFAULT '',
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(employee_id, name)
            );

            CREATE TABLE IF NOT EXISTS personal_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                category_id INTEGER,
                title TEXT NOT NULL,
                description TEXT NOT NULL DEFAULT '',
                task_tag TEXT NOT NULL DEFAULT '',
                priority TEXT NOT NULL DEFAULT 'normal',
                progress TEXT NOT NULL DEFAULT 'not_started',
                due_date TEXT,
                archived_at TEXT,
                sort_order INTEGER NOT NULL DEFAULT 0,
                priority_sort_order INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(category_id) REFERENCES task_categories(id)
            );
            """
        )

        existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(jobs)").fetchall()}
        migrations = {
            "feature": "ALTER TABLE jobs ADD COLUMN feature TEXT NOT NULL DEFAULT 'fangzheng'",
            "current_row": "ALTER TABLE jobs ADD COLUMN current_row INTEGER NOT NULL DEFAULT 0",
            "total_rows": "ALTER TABLE jobs ADD COLUMN total_rows INTEGER NOT NULL DEFAULT 0",
            "log_text": "ALTER TABLE jobs ADD COLUMN log_text TEXT NOT NULL DEFAULT ''",
            "worker_pid": "ALTER TABLE jobs ADD COLUMN worker_pid INTEGER",
            "worker_started_at": "ALTER TABLE jobs ADD COLUMN worker_started_at TEXT",
        }
        for column, sql in migrations.items():
            if column not in existing_cols:
                conn.execute(sql)

        user_cols = {row["name"] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
        user_migrations = {
            "display_name": "ALTER TABLE users ADD COLUMN display_name TEXT",
            "department": "ALTER TABLE users ADD COLUMN department TEXT",
            "role": "ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'user'",
            "enabled": "ALTER TABLE users ADD COLUMN enabled INTEGER NOT NULL DEFAULT 1",
            "must_change_password": "ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 1",
            "created_at": "ALTER TABLE users ADD COLUMN created_at TEXT NOT NULL DEFAULT ''",
            "updated_at": "ALTER TABLE users ADD COLUMN updated_at TEXT NOT NULL DEFAULT ''",
        }
        for column, sql in user_migrations.items():
            if column not in user_cols:
                conn.execute(sql)

        ai_config_cols = {
            row["name"] for row in conn.execute("PRAGMA table_info(pdf_excel_ai_config_versions)").fetchall()
        }
        ai_config_migrations = {
            "activated_by": "ALTER TABLE pdf_excel_ai_config_versions ADD COLUMN activated_by TEXT NOT NULL DEFAULT ''",
            "activated_at": "ALTER TABLE pdf_excel_ai_config_versions ADD COLUMN activated_at TEXT NOT NULL DEFAULT ''",
        }
        for column, sql in ai_config_migrations.items():
            if column not in ai_config_cols:
                conn.execute(sql)

        task_cols = {row["name"] for row in conn.execute("PRAGMA table_info(personal_tasks)").fetchall()}
        task_migrations = {
            "sort_order": "ALTER TABLE personal_tasks ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0",
            "priority_sort_order": "ALTER TABLE personal_tasks ADD COLUMN priority_sort_order INTEGER NOT NULL DEFAULT 0",
            "task_tag": "ALTER TABLE personal_tasks ADD COLUMN task_tag TEXT NOT NULL DEFAULT ''",
        }
        for column, sql in task_migrations.items():
            if column not in task_cols:
                conn.execute(sql)
        conn.execute("UPDATE personal_tasks SET progress = 'not_started' WHERE progress = 'waiting'")

        category_cols = {row["name"] for row in conn.execute("PRAGMA table_info(task_categories)").fetchall()}
        category_migrations = {
            "short_label": "ALTER TABLE task_categories ADD COLUMN short_label TEXT NOT NULL DEFAULT ''",
        }
        for column, sql in category_migrations.items():
            if column not in category_cols:
                conn.execute(sql)
        conn.execute(
            """
            UPDATE task_categories
            SET short_label = CASE name
                WHEN '人力资源任务' THEN 'HR'
                WHEN 'AI开发任务' THEN 'AI'
                WHEN '其它业务任务' THEN '业务'
                ELSE short_label
            END
            WHERE short_label = ''
            """
        )

        existing_users = conn.execute("SELECT COUNT(*) AS total FROM users").fetchone()["total"]
        if existing_users == 0:
            employee_rows = conn.execute(
                "SELECT DISTINCT employee_id FROM jobs WHERE employee_id IS NOT NULL AND employee_id != ''"
            ).fetchall()
            now = utcnow()
            for row in employee_rows:
                employee_id = str(row["employee_id"]).strip()
                if employee_id:
                    conn.execute(
                        """
                        INSERT OR IGNORE INTO users (
                            employee_id, password_hash, role, enabled, must_change_password, created_at, updated_at
                        ) VALUES (?, ?, 'user', 1, 1, ?, ?)
                        """,
                        (employee_id, generate_password_hash(employee_id), now, now),
                    )

    if get_setting("admin_password_hash") is None:
        set_setting("admin_password_hash", generate_password_hash("admin123"))
    if get_setting("active_rule_version") is None:
        set_setting("active_rule_version", "")
    if get_setting("active_pdf_excel_ai_config_version") is None:
        set_setting("active_pdf_excel_ai_config_version", "")


def get_setting(key: str, default: str | None = None) -> str | None:
    with db_cursor() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else default


def set_setting(key: str, value: str) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO settings(key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def verify_admin_password(password: str) -> bool:
    hash_value = get_setting("admin_password_hash", "")
    return bool(hash_value and check_password_hash(hash_value, password))


def update_admin_password(new_password: str) -> None:
    set_setting("admin_password_hash", generate_password_hash(new_password))


def get_user(employee_id: str):
    with db_cursor() as conn:
        return conn.execute("SELECT * FROM users WHERE employee_id = ?", (employee_id,)).fetchone()


def list_users():
    with db_cursor() as conn:
        return conn.execute("SELECT * FROM users ORDER BY employee_id").fetchall()


def create_user(
    employee_id: str,
    *,
    display_name: str = "",
    department: str = "",
    role: str = "user",
    enabled: bool = True,
) -> None:
    now = utcnow()
    with db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO users (
                employee_id, display_name, department, password_hash, role, enabled,
                must_change_password, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?)
            ON CONFLICT(employee_id) DO UPDATE SET
                display_name = excluded.display_name,
                department = excluded.department,
                role = excluded.role,
                enabled = excluded.enabled,
                updated_at = excluded.updated_at
            """,
            (
                employee_id,
                display_name,
                department,
                generate_password_hash(employee_id),
                role,
                1 if enabled else 0,
                now,
                now,
            ),
        )


def verify_user_password(employee_id: str, password: str) -> bool:
    user = get_user(employee_id)
    if not user or not user["enabled"]:
        return False
    return check_password_hash(user["password_hash"], password)


def change_user_password(employee_id: str, new_password: str, *, must_change_password: bool = False) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE users
            SET password_hash = ?, must_change_password = ?, updated_at = ?
            WHERE employee_id = ?
            """,
            (generate_password_hash(new_password), 1 if must_change_password else 0, utcnow(), employee_id),
        )


def reset_user_password(employee_id: str) -> None:
    change_user_password(employee_id, employee_id, must_change_password=True)


def is_admin_user(employee_id: str | None) -> bool:
    if not employee_id:
        return False
    user = get_user(employee_id)
    return bool(user and user["enabled"] and user["role"] == "admin")


def ensure_bootstrap_user(employee_id: str, password: str) -> bool:
    """Allow first legacy login only when no users exist yet."""
    with db_cursor() as conn:
        total = conn.execute("SELECT COUNT(*) AS total FROM users").fetchone()["total"]
    if total != 0 or not employee_id or password != employee_id:
        return False
    create_user(employee_id, role="admin")
    return True


def create_job(employee_id: str, source_filename: str, stored_input_path: str, rule_version: str, feature: str = "fangzheng") -> int:
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            INSERT INTO jobs (
                feature, employee_id, source_filename, stored_input_path, status, rule_version, created_at
            ) VALUES (?, ?, ?, ?, 'queued', ?, ?)
            """,
            (feature, employee_id, source_filename, stored_input_path, rule_version, utcnow()),
        )
        return int(cursor.lastrowid)


def update_job_status(
    job_id: int,
    *,
    status: str,
    stored_result_path: str | None = None,
    success_count: int | None = None,
    fail_count: int | None = None,
    skip_count: int | None = None,
    current_row: int | None = None,
    total_rows: int | None = None,
    log_text: str | None = None,
    error_message: str | None = None,
    completed: bool = False,
) -> None:
    fields = ["status = ?"]
    params: list[object] = [status]

    if stored_result_path is not None:
        fields.append("stored_result_path = ?")
        params.append(stored_result_path)
    if success_count is not None:
        fields.append("success_count = ?")
        params.append(success_count)
    if fail_count is not None:
        fields.append("fail_count = ?")
        params.append(fail_count)
    if skip_count is not None:
        fields.append("skip_count = ?")
        params.append(skip_count)
    if current_row is not None:
        fields.append("current_row = ?")
        params.append(current_row)
    if total_rows is not None:
        fields.append("total_rows = ?")
        params.append(total_rows)
    if log_text is not None:
        fields.append("log_text = ?")
        params.append(log_text)
    if error_message is not None:
        fields.append("error_message = ?")
        params.append(error_message)
    if completed:
        fields.append("completed_at = ?")
        params.append(utcnow())

    params.append(job_id)
    with db_cursor() as conn:
        conn.execute(f"UPDATE jobs SET {', '.join(fields)} WHERE id = ?", params)


def set_job_worker(job_id: int, worker_pid: int | None) -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE jobs
            SET worker_pid = ?, worker_started_at = ?
            WHERE id = ?
            """,
            (worker_pid, utcnow() if worker_pid is not None else None, job_id),
        )


def append_job_log(
    job_id: int,
    message: str,
    *,
    success_count: int | None = None,
    fail_count: int | None = None,
    skip_count: int | None = None,
    current_row: int | None = None,
    total_rows: int | None = None,
) -> None:
    with db_cursor() as conn:
        row = conn.execute("SELECT log_text FROM jobs WHERE id = ?", (job_id,)).fetchone()
        existing = row["log_text"] if row and row["log_text"] else ""
        timestamp = datetime.now().strftime("%H:%M:%S")
        new_text = f"{existing}[{timestamp}] {message}\n"
        fields = ["log_text = ?"]
        params: list[object] = [new_text]
        if success_count is not None:
            fields.append("success_count = ?")
            params.append(success_count)
        if fail_count is not None:
            fields.append("fail_count = ?")
            params.append(fail_count)
        if skip_count is not None:
            fields.append("skip_count = ?")
            params.append(skip_count)
        if current_row is not None:
            fields.append("current_row = ?")
            params.append(current_row)
        if total_rows is not None:
            fields.append("total_rows = ?")
            params.append(total_rows)
        params.append(job_id)
        conn.execute(f"UPDATE jobs SET {', '.join(fields)} WHERE id = ?", params)


def get_job(job_id: int):
    with db_cursor() as conn:
        return conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()


def list_jobs(
    employee_id: str | None,
    start_date: str | None = None,
    end_date: str | None = None,
    limit: int = 100,
    feature: str | None = None,
    status: str | None = None,
    keyword: str | None = None,
    offset: int = 0,
):
    query = "SELECT * FROM jobs WHERE 1 = 1"
    params: list[object] = []

    if employee_id is not None:
        query += " AND employee_id = ?"
        params.append(employee_id)

    if feature:
        query += " AND feature = ?"
        params.append(feature)
    if status:
        query += " AND status = ?"
        params.append(status)
    if keyword:
        query += " AND (CAST(id AS TEXT) LIKE ? OR source_filename LIKE ?)"
        pattern = f"%{keyword}%"
        params.extend([pattern, pattern])
    if start_date:
        query += " AND date(created_at) >= date(?)"
        params.append(start_date)
    if end_date:
        query += " AND date(created_at) <= date(?)"
        params.append(end_date)

    query += " ORDER BY id DESC LIMIT ? OFFSET ?"
    params.extend([limit, offset])

    with db_cursor() as conn:
        return conn.execute(query, params).fetchall()


def count_jobs(
    employee_id: str | None,
    start_date: str | None = None,
    end_date: str | None = None,
    feature: str | None = None,
    status: str | None = None,
    keyword: str | None = None,
) -> int:
    query = "SELECT COUNT(*) AS total FROM jobs WHERE 1 = 1"
    params: list[object] = []

    if employee_id is not None:
        query += " AND employee_id = ?"
        params.append(employee_id)
    if feature:
        query += " AND feature = ?"
        params.append(feature)
    if status:
        query += " AND status = ?"
        params.append(status)
    if keyword:
        query += " AND (CAST(id AS TEXT) LIKE ? OR source_filename LIKE ?)"
        pattern = f"%{keyword}%"
        params.extend([pattern, pattern])
    if start_date:
        query += " AND date(created_at) >= date(?)"
        params.append(start_date)
    if end_date:
        query += " AND date(created_at) <= date(?)"
        params.append(end_date)

    with db_cursor() as conn:
        row = conn.execute(query, params).fetchone()
    return int(row["total"] if row else 0)


def get_active_job(employee_id: str, feature: str):
    with db_cursor() as conn:
        return conn.execute(
            """
            SELECT * FROM jobs
            WHERE employee_id = ? AND feature = ? AND status IN ('queued', 'running')
            ORDER BY id DESC
            LIMIT 1
            """,
            (employee_id, feature),
        ).fetchone()


def prune_jobs_for_employee(employee_id: str, keep_limit: int = 500) -> list[sqlite3.Row]:
    with db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT * FROM jobs
            WHERE employee_id = ?
            ORDER BY id DESC
            LIMIT -1 OFFSET ?
            """,
            (employee_id, keep_limit),
        ).fetchall()
        if rows:
            conn.executemany("DELETE FROM jobs WHERE id = ?", [(row["id"],) for row in rows])
    return rows


def delete_job(job_id: int) -> sqlite3.Row | None:
    with db_cursor() as conn:
        row = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
        if row:
            conn.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
    return row


def list_expired_terminal_jobs(cutoff: str) -> list[sqlite3.Row]:
    with db_cursor() as conn:
        return conn.execute(
            """
            SELECT * FROM jobs
            WHERE status IN ('completed', 'failed', 'canceled')
              AND datetime(COALESCE(completed_at, created_at)) < datetime(?)
            ORDER BY id
            """,
            (cutoff,),
        ).fetchall()


def delete_terminal_jobs(job_ids: list[int]) -> int:
    if not job_ids:
        return 0
    placeholders = ", ".join("?" for _ in job_ids)
    with db_cursor() as conn:
        cursor = conn.execute(
            f"""
            DELETE FROM jobs
            WHERE id IN ({placeholders})
              AND status IN ('completed', 'failed', 'canceled')
            """,
            job_ids,
        )
        return cursor.rowcount


def list_job_ids() -> set[int]:
    with db_cursor() as conn:
        return {int(row["id"]) for row in conn.execute("SELECT id FROM jobs").fetchall()}


def create_feedback(
    employee_id: str,
    feedback_type: str,
    *,
    feature: str = "",
    material_desc: str = "",
    system_result: str = "",
    expected_result: str = "",
    content: str = "",
    daily_workload: str = "",
    error_probability: str = "",
    urgency: str = "",
) -> int:
    now = utcnow()
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            INSERT INTO feedback (
                employee_id, feedback_type, feature, material_desc, system_result, expected_result,
                content, daily_workload, error_probability, urgency, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                employee_id,
                feedback_type,
                feature,
                material_desc,
                system_result,
                expected_result,
                content,
                daily_workload,
                error_probability,
                urgency,
                now,
                now,
            ),
        )
        return int(cursor.lastrowid)


def list_feedback(employee_id: str | None = None, *, feedback_type: str | None = None, status: str | None = None):
    query = "SELECT * FROM feedback WHERE 1 = 1"
    params: list[object] = []
    if employee_id:
        query += " AND employee_id = ?"
        params.append(employee_id)
    if feedback_type:
        query += " AND feedback_type = ?"
        params.append(feedback_type)
    if status:
        query += " AND status = ?"
        params.append(status)
    query += " ORDER BY id DESC"
    with db_cursor() as conn:
        return conn.execute(query, params).fetchall()


def get_feedback(feedback_id: int):
    with db_cursor() as conn:
        return conn.execute("SELECT * FROM feedback WHERE id = ?", (feedback_id,)).fetchone()


def update_feedback_status(feedback_id: int, status: str, admin_note: str = "") -> None:
    with db_cursor() as conn:
        conn.execute(
            """
            UPDATE feedback
            SET status = ?, admin_note = ?, updated_at = ?
            WHERE id = ?
            """,
            (status, admin_note, utcnow(), feedback_id),
        )


def get_rule_history() -> list[dict]:
    raw = get_setting("rule_history", "[]") or "[]"
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def append_rule_history(entry: dict) -> None:
    history = get_rule_history()
    history.insert(0, entry)
    set_setting("rule_history", json.dumps(history[:50], ensure_ascii=False))


DEFAULT_TASK_CATEGORIES = ["人力资源任务", "AI开发任务", "其它业务任务"]


def ensure_default_task_categories(employee_id: str) -> None:
    now = utcnow()
    with db_cursor() as conn:
        existing_count = conn.execute(
            "SELECT COUNT(*) AS total FROM task_categories WHERE employee_id = ?",
            (employee_id,),
        ).fetchone()["total"]
        if existing_count:
            return
        conn.executemany(
            """
            INSERT OR IGNORE INTO task_categories (
                employee_id, name, short_label, sort_order, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    employee_id,
                    name,
                    {"人力资源任务": "HR", "AI开发任务": "AI", "其它业务任务": "业务"}.get(name, ""),
                    index,
                    now,
                    now,
                )
                for index, name in enumerate(DEFAULT_TASK_CATEGORIES, start=1)
            ],
        )


def list_task_categories(employee_id: str):
    ensure_default_task_categories(employee_id)
    with db_cursor() as conn:
        return conn.execute(
            """
            SELECT
                c.*,
                COUNT(t.id) AS task_count,
                SUM(CASE WHEN t.archived_at IS NULL THEN 1 ELSE 0 END) AS active_count
            FROM task_categories c
            LEFT JOIN personal_tasks t ON t.category_id = c.id AND t.employee_id = c.employee_id
            WHERE c.employee_id = ?
            GROUP BY c.id
            ORDER BY c.sort_order, c.id
            """,
            (employee_id,),
        ).fetchall()


def get_task_category(category_id: int, employee_id: str):
    with db_cursor() as conn:
        return conn.execute(
            "SELECT * FROM task_categories WHERE id = ? AND employee_id = ?",
            (category_id, employee_id),
        ).fetchone()


def create_task_category(employee_id: str, name: str, short_label: str = "") -> int:
    now = utcnow()
    name = name.strip()
    with db_cursor() as conn:
        max_sort = conn.execute(
            "SELECT COALESCE(MAX(sort_order), 0) AS sort_order FROM task_categories WHERE employee_id = ?",
            (employee_id,),
        ).fetchone()["sort_order"]
        cursor = conn.execute(
            """
            INSERT INTO task_categories (employee_id, name, short_label, sort_order, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(employee_id, name) DO UPDATE SET updated_at = excluded.updated_at
            """,
            (employee_id, name, short_label.strip()[:12], int(max_sort) + 1, now, now),
        )
        if cursor.lastrowid:
            return int(cursor.lastrowid)
        row = conn.execute(
            "SELECT id FROM task_categories WHERE employee_id = ? AND name = ?",
            (employee_id, name),
        ).fetchone()
        return int(row["id"])


def update_task_category(category_id: int, employee_id: str, name: str, short_label: str = "") -> bool:
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            UPDATE task_categories
            SET name = ?, short_label = ?, updated_at = ?
            WHERE id = ? AND employee_id = ?
            """,
            (name.strip(), short_label.strip()[:12], utcnow(), category_id, employee_id),
        )
        return cursor.rowcount > 0


def delete_task_category(category_id: int, employee_id: str) -> bool:
    with db_cursor() as conn:
        category = conn.execute(
            "SELECT * FROM task_categories WHERE id = ? AND employee_id = ?",
            (category_id, employee_id),
        ).fetchone()
        if not category:
            return False
        conn.execute(
            "UPDATE personal_tasks SET category_id = NULL, updated_at = ? WHERE category_id = ? AND employee_id = ?",
            (utcnow(), category_id, employee_id),
        )
        conn.execute("DELETE FROM task_categories WHERE id = ? AND employee_id = ?", (category_id, employee_id))
        return True


def create_personal_task(
    employee_id: str,
    *,
    title: str,
    category_id: int | None,
    description: str = "",
    task_tag: str = "",
    priority: str = "normal",
    progress: str = "not_started",
    due_date: str | None = None,
) -> int:
    now = utcnow()
    with db_cursor() as conn:
        max_sort = conn.execute(
            """
            SELECT COALESCE(MAX(sort_order), 0) AS sort_order
            FROM personal_tasks
            WHERE employee_id = ? AND COALESCE(category_id, 0) = COALESCE(?, 0)
            """,
            (employee_id, category_id),
        ).fetchone()["sort_order"]
        max_priority_sort = conn.execute(
            """
            SELECT COALESCE(MAX(priority_sort_order), 0) AS sort_order
            FROM personal_tasks
            WHERE employee_id = ? AND priority = ?
            """,
            (employee_id, priority),
        ).fetchone()["sort_order"]
        cursor = conn.execute(
            """
            INSERT INTO personal_tasks (
                employee_id, category_id, title, description, task_tag, priority, progress,
                due_date, sort_order, priority_sort_order, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                employee_id,
                category_id,
                title.strip(),
                description.strip(),
                task_tag.strip(),
                priority,
                progress,
                due_date,
                int(max_sort) + 1,
                int(max_priority_sort) + 1,
                now,
                now,
            ),
        )
        return int(cursor.lastrowid)


def list_personal_tasks(
    employee_id: str,
    *,
    archived: bool | None = False,
    category_id: int | None = None,
    progress: str | None = None,
    due_on: str | None = None,
    due_before: str | None = None,
):
    query = """
        SELECT t.*, c.name AS category_name, c.short_label AS category_label
        FROM personal_tasks t
        LEFT JOIN task_categories c ON c.id = t.category_id AND c.employee_id = t.employee_id
        WHERE t.employee_id = ?
    """
    params: list[object] = [employee_id]
    if archived is True:
        query += " AND t.archived_at IS NOT NULL"
    elif archived is False:
        query += " AND t.archived_at IS NULL"
    if category_id is not None:
        query += " AND t.category_id = ?"
        params.append(category_id)
    if progress:
        query += " AND t.progress = ?"
        params.append(progress)
    if due_on:
        query += " AND t.due_date = ?"
        params.append(due_on)
    if due_before:
        query += " AND t.due_date IS NOT NULL AND date(t.due_date) < date(?)"
        params.append(due_before)
    query += """
        ORDER BY
            COALESCE(c.sort_order, 999999),
            c.id,
            t.sort_order,
            t.id DESC
    """
    with db_cursor() as conn:
        return conn.execute(query, params).fetchall()


def get_personal_task(task_id: int, employee_id: str):
    with db_cursor() as conn:
        return conn.execute(
            """
            SELECT t.*, c.name AS category_name, c.short_label AS category_label
            FROM personal_tasks t
            LEFT JOIN task_categories c ON c.id = t.category_id AND c.employee_id = t.employee_id
            WHERE t.id = ? AND t.employee_id = ?
            """,
            (task_id, employee_id),
        ).fetchone()


def update_personal_task(
    task_id: int,
    employee_id: str,
    *,
    title: str,
    category_id: int | None,
    description: str,
    task_tag: str,
    priority: str,
    progress: str,
    due_date: str | None,
) -> bool:
    archived_at = utcnow() if progress == "completed" else None
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            UPDATE personal_tasks
            SET title = ?,
                category_id = ?,
                description = ?,
                task_tag = ?,
                priority = ?,
                progress = ?,
                due_date = ?,
                archived_at = ?,
                updated_at = ?
            WHERE id = ? AND employee_id = ?
            """,
            (
                title.strip(),
                category_id,
                description.strip(),
                task_tag.strip(),
                priority,
                progress,
                due_date,
                archived_at,
                utcnow(),
                task_id,
                employee_id,
            ),
        )
        return cursor.rowcount > 0


def archive_personal_task(task_id: int, employee_id: str) -> bool:
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            UPDATE personal_tasks
            SET progress = 'completed', archived_at = ?, updated_at = ?
            WHERE id = ? AND employee_id = ?
            """,
            (utcnow(), utcnow(), task_id, employee_id),
        )
        return cursor.rowcount > 0


def restore_personal_task(task_id: int, employee_id: str) -> bool:
    with db_cursor() as conn:
        cursor = conn.execute(
            """
            UPDATE personal_tasks
            SET progress = 'in_progress', archived_at = NULL, updated_at = ?
            WHERE id = ? AND employee_id = ?
            """,
            (utcnow(), task_id, employee_id),
        )
        return cursor.rowcount > 0


def delete_personal_task(task_id: int, employee_id: str) -> bool:
    with db_cursor() as conn:
        cursor = conn.execute(
            "DELETE FROM personal_tasks WHERE id = ? AND employee_id = ?",
            (task_id, employee_id),
        )
        return cursor.rowcount > 0


def reorder_personal_tasks(employee_id: str, ordered_ids: list[int], category_id: int | None) -> int:
    if not ordered_ids:
        return 0
    with db_cursor() as conn:
        valid_rows = conn.execute(
            f"""
            SELECT id
            FROM personal_tasks
            WHERE employee_id = ?
              AND archived_at IS NULL
              AND COALESCE(category_id, 0) = COALESCE(?, 0)
              AND id IN ({",".join("?" for _ in ordered_ids)})
            """,
            [employee_id, category_id, *ordered_ids],
        ).fetchall()
        valid_ids = {int(row["id"]) for row in valid_rows}
        updates = [
            (index, task_id)
            for index, task_id in enumerate(ordered_ids, start=1)
            if task_id in valid_ids
        ]
        conn.executemany(
            "UPDATE personal_tasks SET sort_order = ?, updated_at = ? WHERE id = ? AND employee_id = ?",
            [(sort_order, utcnow(), task_id, employee_id) for sort_order, task_id in updates],
        )
        return len(updates)


def reorder_personal_tasks_by_priority(employee_id: str, ordered_ids: list[int], priority: str) -> int:
    if not ordered_ids:
        return 0
    with db_cursor() as conn:
        valid_rows = conn.execute(
            f"""
            SELECT id
            FROM personal_tasks
            WHERE employee_id = ?
              AND archived_at IS NULL
              AND priority = ?
              AND id IN ({",".join("?" for _ in ordered_ids)})
            """,
            [employee_id, priority, *ordered_ids],
        ).fetchall()
        valid_ids = {int(row["id"]) for row in valid_rows}
        updates = [
            (index, task_id)
            for index, task_id in enumerate(ordered_ids, start=1)
            if task_id in valid_ids
        ]
        conn.executemany(
            "UPDATE personal_tasks SET priority_sort_order = ?, updated_at = ? WHERE id = ? AND employee_id = ?",
            [(sort_order, utcnow(), task_id, employee_id) for sort_order, task_id in updates],
        )
        return len(updates)
