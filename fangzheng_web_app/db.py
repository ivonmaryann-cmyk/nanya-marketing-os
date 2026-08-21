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
                confirm_count INTEGER NOT NULL DEFAULT 0,
                verify_count INTEGER NOT NULL DEFAULT 0,
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

            CREATE TABLE IF NOT EXISTS transcode_model_configs (
                employee_id TEXT PRIMARY KEY,
                enabled INTEGER NOT NULL DEFAULT 0,
                base_url TEXT NOT NULL DEFAULT 'https://api.deepseek.com',
                api_key TEXT NOT NULL DEFAULT '',
                model TEXT NOT NULL DEFAULT 'deepseek-v4-pro',
                timeout_seconds REAL NOT NULL DEFAULT 60,
                max_order_calls INTEGER NOT NULL DEFAULT 50,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(employee_id) REFERENCES users(employee_id)
            );

            CREATE TABLE IF NOT EXISTS transcode_agent_confirmation_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER NOT NULL,
                employee_id TEXT NOT NULL,
                excel_row INTEGER NOT NULL,
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                spec TEXT NOT NULL DEFAULT '',
                context_text TEXT NOT NULL DEFAULT '',
                field_key TEXT NOT NULL,
                field_label TEXT NOT NULL,
                current_code TEXT NOT NULL DEFAULT '',
                options_json TEXT NOT NULL DEFAULT '[]',
                pending_code TEXT NOT NULL DEFAULT '',
                score INTEGER NOT NULL DEFAULT 0,
                reason TEXT NOT NULL DEFAULT '',
                evidence_json TEXT NOT NULL DEFAULT '{}',
                analysis_json TEXT NOT NULL DEFAULT '{}',
                status TEXT NOT NULL DEFAULT 'pending',
                confirmed_code TEXT,
                confirmation_basis TEXT,
                confirmed_by TEXT,
                confirmed_at TEXT,
                long_term_rule_id TEXT,
                pending_rule_id TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(job_id, excel_row, field_key),
                FOREIGN KEY(job_id) REFERENCES jobs(id)
            );

            CREATE TABLE IF NOT EXISTS transcode_agent_confirmation_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER,
                job_id INTEGER NOT NULL,
                employee_id TEXT NOT NULL,
                action TEXT NOT NULL,
                before_json TEXT NOT NULL DEFAULT '{}',
                after_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                FOREIGN KEY(item_id) REFERENCES transcode_agent_confirmation_items(id),
                FOREIGN KEY(job_id) REFERENCES jobs(id)
            );

            CREATE TABLE IF NOT EXISTS transcode_agent_pending_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rule_id TEXT NOT NULL,
                rule_json TEXT NOT NULL,
                employee_id TEXT NOT NULL,
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                business_field TEXT NOT NULL DEFAULT '',
                target_value TEXT NOT NULL DEFAULT '',
                condition_summary TEXT NOT NULL DEFAULT '',
                source_task_id INTEGER,
                source_excel_row INTEGER,
                status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                updated_by TEXT NOT NULL DEFAULT '',
                processed_by TEXT NOT NULL DEFAULT '',
                processed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS transcode_agent_row_verifications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER NOT NULL,
                excel_row INTEGER NOT NULL,
                employee_id TEXT NOT NULL,
                action TEXT NOT NULL,
                before_code TEXT NOT NULL DEFAULT '',
                after_code TEXT NOT NULL DEFAULT '',
                basis TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                UNIQUE(job_id, excel_row),
                FOREIGN KEY(job_id) REFERENCES jobs(id)
            );

            CREATE INDEX IF NOT EXISTS idx_transcode_agent_confirm_job_status
            ON transcode_agent_confirmation_items(job_id, status, excel_row);

            CREATE INDEX IF NOT EXISTS idx_transcode_agent_confirm_owner
            ON transcode_agent_confirmation_items(employee_id, job_id);

            CREATE INDEX IF NOT EXISTS idx_transcode_agent_pending_rules_status
            ON transcode_agent_pending_rules(status, employee_id);

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

        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS mail_accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL UNIQUE,
                owner_employee_id TEXT NOT NULL DEFAULT '',
                imap_host TEXT NOT NULL DEFAULT 'imap.163.com',
                imap_port INTEGER NOT NULL DEFAULT 993,
                auth_code_ciphertext TEXT NOT NULL DEFAULT '',
                enabled INTEGER NOT NULL DEFAULT 1,
                last_fetch_at TEXT,
                last_fetch_status TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS mail_messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_id INTEGER NOT NULL,
                folder TEXT NOT NULL DEFAULT 'INBOX',
                uid TEXT NOT NULL DEFAULT '',
                message_id TEXT NOT NULL DEFAULT '',
                subject TEXT NOT NULL DEFAULT '',
                sender TEXT NOT NULL DEFAULT '',
                sent_at TEXT NOT NULL DEFAULT '',
                received_at TEXT NOT NULL DEFAULT '',
                body_html TEXT,
                body_text TEXT,
                eml_path TEXT NOT NULL DEFAULT '',
                is_order INTEGER NOT NULL DEFAULT 0,
                fetch_task_id INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                UNIQUE(account_id, folder, uid)
            );

            CREATE TABLE IF NOT EXISTS mail_attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                mail_id INTEGER NOT NULL,
                filename TEXT NOT NULL DEFAULT '',
                content_type TEXT NOT NULL DEFAULT '',
                size_bytes INTEGER NOT NULL DEFAULT 0,
                sha256 TEXT NOT NULL DEFAULT '',
                stored_path TEXT NOT NULL DEFAULT '',
                is_inline INTEGER NOT NULL DEFAULT 0,
                parse_status TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS mail_order_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                mail_id INTEGER NOT NULL,
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                spec TEXT NOT NULL DEFAULT '',
                remark TEXT NOT NULL DEFAULT '',
                order_number TEXT NOT NULL DEFAULT '',
                source_type TEXT NOT NULL DEFAULT '',
                field_status TEXT NOT NULL DEFAULT 'missing',
                review_status TEXT NOT NULL DEFAULT 'pending_review',
                attachment_parse_status TEXT NOT NULL DEFAULT '',
                transcode_status TEXT NOT NULL DEFAULT 'not_started',
                transcode_code TEXT NOT NULL DEFAULT '',
                transcode_note TEXT NOT NULL DEFAULT '',
                transcode_confidence INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                reviewed_by TEXT NOT NULL DEFAULT '',
                reviewed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS mail_transcode_jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_ids TEXT NOT NULL DEFAULT '',
                input_path TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'running',
                success_count INTEGER NOT NULL DEFAULT 0,
                fail_count INTEGER NOT NULL DEFAULT 0,
                created_by TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                completed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS mail_fetch_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_id INTEGER NOT NULL,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                status TEXT NOT NULL,
                message TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS mail_fetch_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_id INTEGER NOT NULL,
                created_by TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'completed',
                email_count INTEGER NOT NULL DEFAULT 0,
                new_count INTEGER NOT NULL DEFAULT 0,
                duplicate_count INTEGER NOT NULL DEFAULT 0,
                order_count INTEGER NOT NULL DEFAULT 0,
                message TEXT NOT NULL DEFAULT '',
                started_at TEXT NOT NULL,
                completed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS mail_fetch_task_messages (
                fetch_task_id INTEGER NOT NULL,
                mail_id INTEGER NOT NULL,
                is_new INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                PRIMARY KEY (fetch_task_id, mail_id),
                FOREIGN KEY(fetch_task_id) REFERENCES mail_fetch_tasks(id),
                FOREIGN KEY(mail_id) REFERENCES mail_messages(id)
            );

            CREATE INDEX IF NOT EXISTS idx_mail_messages_account_uid
            ON mail_messages(account_id, folder, uid);

            CREATE INDEX IF NOT EXISTS idx_mail_attachments_mail
            ON mail_attachments(mail_id);

            CREATE INDEX IF NOT EXISTS idx_mail_order_tasks_review
            ON mail_order_tasks(review_status, transcode_status);
            """
        )

        mail_message_cols = {
            row["name"] for row in conn.execute("PRAGMA table_info(mail_messages)").fetchall()
        }
        if "fetch_task_id" not in mail_message_cols:
            conn.execute("ALTER TABLE mail_messages ADD COLUMN fetch_task_id INTEGER NOT NULL DEFAULT 0")

        mail_fetch_task_cols = {
            row["name"] for row in conn.execute("PRAGMA table_info(mail_fetch_tasks)").fetchall()
        }
        if "new_count" not in mail_fetch_task_cols:
            conn.execute("ALTER TABLE mail_fetch_tasks ADD COLUMN new_count INTEGER NOT NULL DEFAULT 0")
        if "duplicate_count" not in mail_fetch_task_cols:
            conn.execute("ALTER TABLE mail_fetch_tasks ADD COLUMN duplicate_count INTEGER NOT NULL DEFAULT 0")

        mail_account_cols = {
            row["name"] for row in conn.execute("PRAGMA table_info(mail_accounts)").fetchall()
        }
        if "owner_employee_id" not in mail_account_cols:
            conn.execute(
                "ALTER TABLE mail_accounts ADD COLUMN owner_employee_id TEXT NOT NULL DEFAULT ''"
            )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_mail_accounts_owner "
            "ON mail_accounts(owner_employee_id, enabled, id DESC)"
        )
        legacy_owner = conn.execute(
            "SELECT employee_id FROM users WHERE role = 'admin' ORDER BY employee_id LIMIT 1"
        ).fetchone()
        if not legacy_owner:
            legacy_owner = conn.execute(
                "SELECT employee_id FROM users ORDER BY employee_id LIMIT 1"
            ).fetchone()
        if legacy_owner:
            conn.execute(
                "UPDATE mail_accounts SET owner_employee_id = ? "
                "WHERE TRIM(owner_employee_id) = ''",
                (legacy_owner["employee_id"],),
            )

        # Order intake keeps the business decision separate from the source email.
        # A case is one mail-side business event; later ERP order/version tables can
        # extend it without rewriting the immutable source-mail evidence.
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS order_intake_cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                mail_id INTEGER NOT NULL,
                action_type TEXT NOT NULL DEFAULT 'unclassified',
                status TEXT NOT NULL DEFAULT 'pending_triage',
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                order_number TEXT NOT NULL DEFAULT '',
                order_version TEXT NOT NULL DEFAULT '',
                parent_order_number TEXT NOT NULL DEFAULT '',
                workflow_stage TEXT NOT NULL DEFAULT 'mail_triage',
                customer_match_status TEXT NOT NULL DEFAULT 'unmatched',
                source_document_status TEXT NOT NULL DEFAULT 'pending',
                mapping_status TEXT NOT NULL DEFAULT 'not_started',
                erp_prepare_status TEXT NOT NULL DEFAULT 'not_started',
                routing_source TEXT NOT NULL DEFAULT 'system',
                routing_reason TEXT NOT NULL DEFAULT '',
                routed_by TEXT NOT NULL DEFAULT '',
                routed_at TEXT,
                handling_note TEXT NOT NULL DEFAULT '',
                confirmed_by TEXT NOT NULL DEFAULT '',
                confirmed_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(employee_id, mail_id),
                FOREIGN KEY(mail_id) REFERENCES mail_messages(id)
            );

            CREATE TABLE IF NOT EXISTS order_intake_case_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                case_id INTEGER NOT NULL,
                employee_id TEXT NOT NULL,
                action TEXT NOT NULL,
                before_json TEXT NOT NULL DEFAULT '{}',
                after_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                FOREIGN KEY(case_id) REFERENCES order_intake_cases(id)
            );

            CREATE INDEX IF NOT EXISTS idx_order_intake_cases_employee_status
                ON order_intake_cases(employee_id, status, updated_at DESC);

            CREATE TABLE IF NOT EXISTS order_mail_routing_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                name TEXT NOT NULL,
                enabled INTEGER NOT NULL DEFAULT 1,
                priority INTEGER NOT NULL DEFAULT 100,
                sender_contains TEXT NOT NULL DEFAULT '',
                subject_contains TEXT NOT NULL DEFAULT '',
                attachment_contains TEXT NOT NULL DEFAULT '',
                action_type TEXT NOT NULL,
                customer_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS order_mail_routing_rule_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rule_id INTEGER NOT NULL,
                employee_id TEXT NOT NULL,
                action TEXT NOT NULL,
                before_json TEXT NOT NULL DEFAULT '{}',
                after_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                FOREIGN KEY(rule_id) REFERENCES order_mail_routing_rules(id)
            );

            CREATE INDEX IF NOT EXISTS idx_order_mail_routing_rules_owner
                ON order_mail_routing_rules(employee_id, enabled, priority, id);

            CREATE TABLE IF NOT EXISTS order_mail_rule_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                name TEXT NOT NULL,
                action_type TEXT NOT NULL,
                enabled INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS order_mail_rule_keywords (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                scope TEXT NOT NULL,
                keyword TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(group_id) REFERENCES order_mail_rule_groups(id)
            );

            CREATE TABLE IF NOT EXISTS order_change_tags (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                name TEXT NOT NULL,
                enabled INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(employee_id, name)
            );

            CREATE TABLE IF NOT EXISTS order_change_tag_keywords (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tag_id INTEGER NOT NULL,
                scope TEXT NOT NULL,
                keyword TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(tag_id) REFERENCES order_change_tags(id)
            );

            CREATE TABLE IF NOT EXISTS mail_attachment_texts (
                attachment_id INTEGER PRIMARY KEY,
                text_content TEXT NOT NULL DEFAULT '',
                parse_status TEXT NOT NULL DEFAULT 'pending',
                updated_at TEXT NOT NULL,
                FOREIGN KEY(attachment_id) REFERENCES mail_attachments(id)
            );

            CREATE TABLE IF NOT EXISTS automation_customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_code TEXT NOT NULL UNIQUE,
                customer_short_name TEXT NOT NULL DEFAULT '', quick_code TEXT NOT NULL DEFAULT '',
                customer_name TEXT NOT NULL DEFAULT '', group_name TEXT NOT NULL DEFAULT '', customer_type TEXT NOT NULL DEFAULT '',
                sales_name TEXT NOT NULL DEFAULT '', service_name TEXT NOT NULL DEFAULT '', internal_clerk_name TEXT NOT NULL DEFAULT '',
                internal_clerk_employee_id TEXT NOT NULL DEFAULT '', technical_support_name TEXT NOT NULL DEFAULT '',
                insurer_days TEXT NOT NULL DEFAULT '', credit_amount TEXT NOT NULL DEFAULT '', payment_terms TEXT NOT NULL DEFAULT '',
                grace_days TEXT NOT NULL DEFAULT '', invoice_address TEXT NOT NULL DEFAULT '', delivery_address TEXT NOT NULL DEFAULT '',
                contact_name TEXT NOT NULL DEFAULT '', contact_phone TEXT NOT NULL DEFAULT '', settlement_day TEXT NOT NULL DEFAULT '',
                transit_days TEXT NOT NULL DEFAULT '', first_trade_at TEXT NOT NULL DEFAULT '', last_order_at TEXT NOT NULL DEFAULT '',
                last_delivery_at TEXT NOT NULL DEFAULT '', last_payment_at TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'active',
                source_json TEXT NOT NULL DEFAULT '{}', note TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS automation_customer_contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT, customer_id INTEGER NOT NULL, contact_type TEXT NOT NULL,
                contact_value TEXT NOT NULL, enabled INTEGER NOT NULL DEFAULT 1, note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL, updated_at TEXT NOT NULL, UNIQUE(customer_id, contact_type, contact_value),
                FOREIGN KEY(customer_id) REFERENCES automation_customers(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS automation_customer_routing_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT, customer_id INTEGER NOT NULL, name TEXT NOT NULL, action_type TEXT NOT NULL,
                enabled INTEGER NOT NULL DEFAULT 1, priority INTEGER NOT NULL DEFAULT 100, note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL, updated_at TEXT NOT NULL, FOREIGN KEY(customer_id) REFERENCES automation_customers(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS automation_customer_routing_conditions (
                id INTEGER PRIMARY KEY AUTOINCREMENT, rule_id INTEGER NOT NULL, scope TEXT NOT NULL, keyword TEXT NOT NULL,
                created_at TEXT NOT NULL, FOREIGN KEY(rule_id) REFERENCES automation_customer_routing_rules(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS automation_customer_extraction_maps (
                id INTEGER PRIMARY KEY AUTOINCREMENT, customer_id INTEGER NOT NULL, target_field TEXT NOT NULL,
                source_kind TEXT NOT NULL DEFAULT 'attachment_table', source_label TEXT NOT NULL DEFAULT '',
                transform_type TEXT NOT NULL DEFAULT 'direct', transform_config_json TEXT NOT NULL DEFAULT '{}',
                enabled INTEGER NOT NULL DEFAULT 1, sort_order INTEGER NOT NULL DEFAULT 100, note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL, updated_at TEXT NOT NULL, FOREIGN KEY(customer_id) REFERENCES automation_customers(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS automation_customer_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT, customer_id INTEGER NOT NULL, action TEXT NOT NULL,
                before_json TEXT NOT NULL DEFAULT '{}', after_json TEXT NOT NULL DEFAULT '{}', operated_by TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL, FOREIGN KEY(customer_id) REFERENCES automation_customers(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_automation_customers_status ON automation_customers(status, customer_code);
            CREATE INDEX IF NOT EXISTS idx_automation_customer_contacts_lookup ON automation_customer_contacts(contact_type, contact_value, enabled);
            CREATE INDEX IF NOT EXISTS idx_automation_customer_rules_customer ON automation_customer_routing_rules(customer_id, enabled, priority, id);
            CREATE INDEX IF NOT EXISTS idx_automation_customer_conditions_rule ON automation_customer_routing_conditions(rule_id, scope);
            CREATE INDEX IF NOT EXISTS idx_automation_customer_maps_customer ON automation_customer_extraction_maps(customer_id, enabled, sort_order);

            -- One domestic order-entry workbook belongs to exactly one routed
            -- mail case. The current editable values live in the header/line
            -- tables; every save also records an immutable snapshot version.
            CREATE TABLE IF NOT EXISTS order_entry_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                case_id INTEGER NOT NULL UNIQUE,
                employee_id TEXT NOT NULL,
                template_key TEXT NOT NULL DEFAULT '151_domestic_v1',
                header_json TEXT NOT NULL DEFAULT '{}',
                current_version INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(case_id) REFERENCES order_intake_cases(id)
            );

            CREATE TABLE IF NOT EXISTS order_entry_template_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER NOT NULL,
                line_no INTEGER NOT NULL,
                values_json TEXT NOT NULL DEFAULT '{}',
                sources_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(template_id, line_no),
                FOREIGN KEY(template_id) REFERENCES order_entry_templates(id)
            );

            CREATE TABLE IF NOT EXISTS order_entry_template_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER NOT NULL,
                version_number INTEGER NOT NULL,
                header_json TEXT NOT NULL DEFAULT '{}',
                lines_json TEXT NOT NULL DEFAULT '[]',
                saved_by TEXT NOT NULL,
                saved_at TEXT NOT NULL,
                UNIQUE(template_id, version_number),
                FOREIGN KEY(template_id) REFERENCES order_entry_templates(id)
            );

            CREATE INDEX IF NOT EXISTS idx_order_entry_templates_case
                ON order_entry_templates(case_id, employee_id);
            CREATE INDEX IF NOT EXISTS idx_order_entry_lines_template
                ON order_entry_template_lines(template_id, line_no);
            """
        )

        order_intake_cols = {
            row["name"] for row in conn.execute("PRAGMA table_info(order_intake_cases)").fetchall()
        }
        order_intake_migrations = {
            "order_version": "TEXT NOT NULL DEFAULT ''",
            "parent_order_number": "TEXT NOT NULL DEFAULT ''",
            "workflow_stage": "TEXT NOT NULL DEFAULT 'mail_triage'",
            "customer_match_status": "TEXT NOT NULL DEFAULT 'unmatched'",
            "source_document_status": "TEXT NOT NULL DEFAULT 'pending'",
            "mapping_status": "TEXT NOT NULL DEFAULT 'not_started'",
            "erp_prepare_status": "TEXT NOT NULL DEFAULT 'not_started'",
            "routing_source": "TEXT NOT NULL DEFAULT 'system'",
            "routing_reason": "TEXT NOT NULL DEFAULT ''",
            "routed_by": "TEXT NOT NULL DEFAULT ''",
            "routed_at": "TEXT",
            "routing_rule_id": "INTEGER",
            "routing_state": "TEXT NOT NULL DEFAULT 'unrouted'",
            "routing_matches_json": "TEXT NOT NULL DEFAULT '[]'",
            "change_tags_json": "TEXT NOT NULL DEFAULT '[]'",
            "completed_at": "TEXT",
            "customer_id": "INTEGER",
            "customer_match_source": "TEXT NOT NULL DEFAULT ''",
            "customer_match_detail": "TEXT NOT NULL DEFAULT ''",
        }
        for column, definition in order_intake_migrations.items():
            if column not in order_intake_cols:
                conn.execute(f"ALTER TABLE order_intake_cases ADD COLUMN {column} {definition}")
        conn.execute("UPDATE order_intake_cases SET action_type = 'new_order' WHERE action_type = 'new'")
        conn.execute("UPDATE order_intake_cases SET action_type = 'order_change' WHERE action_type IN ('modify', 'cancel')")
        conn.execute("UPDATE order_intake_cases SET action_type = 'other' WHERE action_type = 'not_order'")
        conn.execute("UPDATE order_intake_cases SET action_type = 'order_change' WHERE action_type = 'delivery'")
        conn.execute("UPDATE order_intake_cases SET action_type = 'unclassified' WHERE action_type = 'other'")

        existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(jobs)").fetchall()}
        migrations = {
            "feature": "ALTER TABLE jobs ADD COLUMN feature TEXT NOT NULL DEFAULT 'fangzheng'",
            "current_row": "ALTER TABLE jobs ADD COLUMN current_row INTEGER NOT NULL DEFAULT 0",
            "total_rows": "ALTER TABLE jobs ADD COLUMN total_rows INTEGER NOT NULL DEFAULT 0",
            "log_text": "ALTER TABLE jobs ADD COLUMN log_text TEXT NOT NULL DEFAULT ''",
            "worker_pid": "ALTER TABLE jobs ADD COLUMN worker_pid INTEGER",
            "worker_started_at": "ALTER TABLE jobs ADD COLUMN worker_started_at TEXT",
            "confirm_count": "ALTER TABLE jobs ADD COLUMN confirm_count INTEGER NOT NULL DEFAULT 0",
            "verify_count": "ALTER TABLE jobs ADD COLUMN verify_count INTEGER NOT NULL DEFAULT 0",
        }
        for column, sql in migrations.items():
            if column not in existing_cols:
                conn.execute(sql)

        confirmation_cols = {
            row["name"]
            for row in conn.execute(
                "PRAGMA table_info(transcode_agent_confirmation_items)"
            ).fetchall()
        }
        if "long_term_rule_id" not in confirmation_cols:
            conn.execute(
                "ALTER TABLE transcode_agent_confirmation_items ADD COLUMN long_term_rule_id TEXT"
            )
        if "pending_rule_id" not in confirmation_cols:
            conn.execute(
                "ALTER TABLE transcode_agent_confirmation_items ADD COLUMN pending_rule_id TEXT"
            )

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

        from .automation_migration.outbox import ensure_sqlite_outbox

        ensure_sqlite_outbox(conn)

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


@contextmanager
def identity_db_cursor():
    from .database.identity import identity_cursor

    with identity_cursor() as conn:
        yield conn


@contextmanager
def transcode_db_cursor():
    from .database.transcode import transcode_cursor

    with transcode_cursor() as conn:
        yield conn


def get_user(employee_id: str):
    with identity_db_cursor() as conn:
        return conn.execute("SELECT * FROM users WHERE employee_id = ?", (employee_id,)).fetchone()


def list_users():
    with identity_db_cursor() as conn:
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
    with identity_db_cursor() as conn:
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
    with identity_db_cursor() as conn:
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


def get_transcode_model_config(employee_id: str):
    with transcode_db_cursor() as conn:
        return conn.execute(
            "SELECT * FROM transcode_model_configs WHERE employee_id = ?",
            (employee_id,),
        ).fetchone()


def save_transcode_model_config(
    employee_id: str,
    *,
    enabled: bool,
    base_url: str,
    api_key: str | None,
    model: str,
    timeout_seconds: float = 60,
    max_order_calls: int = 50,
) -> None:
    existing = get_transcode_model_config(employee_id)
    stored_key = str(existing["api_key"] or "") if existing else ""
    if api_key is not None:
        stored_key = str(api_key).strip()
    with transcode_db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_model_configs (
                employee_id, enabled, base_url, api_key, model,
                timeout_seconds, max_order_calls, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(employee_id) DO UPDATE SET
                enabled = excluded.enabled,
                base_url = excluded.base_url,
                api_key = excluded.api_key,
                model = excluded.model,
                timeout_seconds = excluded.timeout_seconds,
                max_order_calls = excluded.max_order_calls,
                updated_at = excluded.updated_at
            """,
            (
                employee_id,
                1 if enabled else 0,
                base_url,
                stored_key,
                model,
                timeout_seconds,
                max_order_calls,
                utcnow(),
            ),
        )


def ensure_bootstrap_user(employee_id: str, password: str) -> bool:
    """Allow first legacy login only when no users exist yet."""
    with identity_db_cursor() as conn:
        total = conn.execute("SELECT COUNT(*) AS total FROM users").fetchone()["total"]
    if total != 0 or not employee_id or password != employee_id:
        return False
    create_user(employee_id, role="admin")
    return True


def create_job(employee_id: str, source_filename: str, stored_input_path: str, rule_version: str, feature: str = "fangzheng") -> int:
    with transcode_db_cursor() as conn:
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
    confirm_count: int | None = None,
    verify_count: int | None = None,
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
    if confirm_count is not None:
        fields.append("confirm_count = ?")
        params.append(confirm_count)
    if verify_count is not None:
        fields.append("verify_count = ?")
        params.append(verify_count)
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
    with transcode_db_cursor() as conn:
        conn.execute(f"UPDATE jobs SET {', '.join(fields)} WHERE id = ?", params)


def set_job_worker(job_id: int, worker_pid: int | None) -> None:
    with transcode_db_cursor() as conn:
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
    confirm_count: int | None = None,
    verify_count: int | None = None,
    current_row: int | None = None,
    total_rows: int | None = None,
) -> None:
    with transcode_db_cursor() as conn:
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
        if confirm_count is not None:
            fields.append("confirm_count = ?")
            params.append(confirm_count)
        if verify_count is not None:
            fields.append("verify_count = ?")
            params.append(verify_count)
        if current_row is not None:
            fields.append("current_row = ?")
            params.append(current_row)
        if total_rows is not None:
            fields.append("total_rows = ?")
            params.append(total_rows)
        params.append(job_id)
        conn.execute(f"UPDATE jobs SET {', '.join(fields)} WHERE id = ?", params)


def get_job(job_id: int):
    with transcode_db_cursor() as conn:
        return conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()


def list_jobs(
    employee_id: str,
    start_date: str | None = None,
    end_date: str | None = None,
    limit: int = 100,
    feature: str | None = None,
):
    query = "SELECT * FROM jobs WHERE employee_id = ?"
    params: list[object] = [employee_id]

    if feature:
        query += " AND feature = ?"
        params.append(feature)
    if start_date:
        query += " AND date(created_at) >= date(?)"
        params.append(start_date)
    if end_date:
        query += " AND date(created_at) <= date(?)"
        params.append(end_date)

    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)

    with transcode_db_cursor() as conn:
        return conn.execute(query, params).fetchall()


def get_active_job(employee_id: str, feature: str):
    with transcode_db_cursor() as conn:
        return conn.execute(
            """
            SELECT * FROM jobs
            WHERE employee_id = ? AND feature = ? AND status IN ('queued', 'running')
            ORDER BY id DESC
            LIMIT 1
            """,
            (employee_id, feature),
        ).fetchone()


def replace_transcode_agent_confirmation_items(
    job_id: int,
    employee_id: str,
    items: list[dict],
) -> list[int]:
    now = utcnow()
    inserted_ids: list[int] = []
    with transcode_db_cursor() as conn:
        conn.execute(
            "DELETE FROM transcode_agent_confirmation_events WHERE job_id = ?",
            (job_id,),
        )
        conn.execute(
            "DELETE FROM transcode_agent_confirmation_items WHERE job_id = ?",
            (job_id,),
        )
        for item in items:
            cursor = conn.execute(
                """
                INSERT INTO transcode_agent_confirmation_items (
                    job_id, employee_id, excel_row, customer_code, customer_name,
                    spec, context_text, field_key, field_label, current_code,
                    options_json, pending_code, score, reason, evidence_json,
                    analysis_json, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
                """,
                (
                    job_id,
                    employee_id,
                    int(item.get("excel_row") or 0),
                    str(item.get("customer_code") or ""),
                    str(item.get("customer_name") or ""),
                    str(item.get("spec") or ""),
                    str(item.get("context_text") or ""),
                    str(item.get("field_key") or ""),
                    str(item.get("field_label") or ""),
                    str(item.get("current_code") or ""),
                    json.dumps(item.get("options") or [], ensure_ascii=False, default=str),
                    str(item.get("pending_code") or ""),
                    int(item.get("score") or 0),
                    str(item.get("reason") or ""),
                    json.dumps(item.get("evidence") or {}, ensure_ascii=False, default=str),
                    json.dumps(item.get("analysis") or {}, ensure_ascii=False, default=str),
                    now,
                    now,
                ),
            )
            inserted_ids.append(int(cursor.lastrowid))
    return inserted_ids


def list_transcode_agent_confirmation_items(
    job_id: int,
    employee_id: str,
    *,
    status: str | None = None,
) -> list[sqlite3.Row]:
    query = """
        SELECT * FROM transcode_agent_confirmation_items
        WHERE job_id = ? AND employee_id = ?
    """
    params: list[object] = [job_id, employee_id]
    if status:
        query += " AND status = ?"
        params.append(status)
    query += " ORDER BY excel_row, id"
    with transcode_db_cursor() as conn:
        return conn.execute(query, params).fetchall()


def get_transcode_agent_confirmation_item(item_id: int) -> sqlite3.Row | None:
    with transcode_db_cursor() as conn:
        return conn.execute(
            "SELECT * FROM transcode_agent_confirmation_items WHERE id = ?",
            (item_id,),
        ).fetchone()


def update_transcode_agent_confirmation_item(
    item_id: int,
    *,
    status: str,
    confirmed_code: str | None,
    confirmation_basis: str,
    confirmed_by: str,
    analysis: dict | None = None,
    long_term_rule_id: str | None = None,
    pending_rule_id: str | None = None,
) -> sqlite3.Row | None:
    now = utcnow()
    with transcode_db_cursor() as conn:
        before = conn.execute(
            "SELECT * FROM transcode_agent_confirmation_items WHERE id = ?",
            (item_id,),
        ).fetchone()
        if not before:
            return None
        fields = [
            "status = ?",
            "confirmed_code = ?",
            "confirmation_basis = ?",
            "confirmed_by = ?",
            "confirmed_at = ?",
            "updated_at = ?",
        ]
        params: list[object] = [
            status,
            confirmed_code,
            confirmation_basis,
            confirmed_by,
            now,
            now,
        ]
        if analysis is not None:
            fields.append("analysis_json = ?")
            params.append(json.dumps(analysis, ensure_ascii=False, default=str))
        if status in {"confirmed", "auto_resolved"} and confirmed_code:
            fields.extend(["current_code = ?", "score = 100"])
            params.append(confirmed_code)
        if long_term_rule_id is not None:
            fields.append("long_term_rule_id = ?")
            params.append(long_term_rule_id)
        if pending_rule_id is not None:
            fields.append("pending_rule_id = ?")
            params.append(pending_rule_id)
        params.append(item_id)
        conn.execute(
            f"UPDATE transcode_agent_confirmation_items SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        after = conn.execute(
            "SELECT * FROM transcode_agent_confirmation_items WHERE id = ?",
            (item_id,),
        ).fetchone()
        conn.execute(
            """
            INSERT INTO transcode_agent_confirmation_events (
                item_id, job_id, employee_id, action, before_json, after_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                item_id,
                before["job_id"],
                confirmed_by,
                status,
                json.dumps(dict(before), ensure_ascii=False, default=str),
                json.dumps(dict(after), ensure_ascii=False, default=str),
                now,
            ),
        )
        return after


def refresh_transcode_agent_confirmation_item(
    item_id: int,
    *,
    current_code: str,
    pending_code: str,
    score: int,
    reason: str,
    evidence: dict,
    analysis: dict,
) -> None:
    with transcode_db_cursor() as conn:
        conn.execute(
            """
            UPDATE transcode_agent_confirmation_items
            SET current_code = ?, pending_code = ?, score = ?, reason = ?,
                evidence_json = ?, analysis_json = ?, updated_at = ?
            WHERE id = ? AND status = 'pending'
            """,
            (
                current_code,
                pending_code,
                int(score),
                reason,
                json.dumps(evidence or {}, ensure_ascii=False, default=str),
                json.dumps(analysis or {}, ensure_ascii=False, default=str),
                utcnow(),
                item_id,
            ),
        )


def update_transcode_agent_row_analysis(
    job_id: int,
    excel_row: int,
    analysis: dict,
) -> None:
    with transcode_db_cursor() as conn:
        conn.execute(
            """
            UPDATE transcode_agent_confirmation_items
            SET analysis_json = ?, pending_code = ?, updated_at = ?
            WHERE job_id = ? AND excel_row = ?
            """,
            (
                json.dumps(analysis, ensure_ascii=False, default=str),
                str(analysis.get("candidate_code") or ""),
                utcnow(),
                job_id,
                excel_row,
            ),
        )


def transcode_agent_confirmation_counts(job_id: int) -> dict[str, int]:
    with transcode_db_cursor() as conn:
        rows = conn.execute(
            """
            SELECT status, COUNT(*) AS total
            FROM transcode_agent_confirmation_items
            WHERE job_id = ?
            GROUP BY status
            """,
            (job_id,),
        ).fetchall()
    counts = {"pending": 0, "confirmed": 0, "skipped": 0, "total": 0}
    for row in rows:
        counts[str(row["status"])] = int(row["total"])
        counts["total"] += int(row["total"])
    return counts


def list_transcode_agent_confirmation_events(job_id: int) -> list[sqlite3.Row]:
    with transcode_db_cursor() as conn:
        return conn.execute(
            """
            SELECT id, item_id, job_id, employee_id, action,
                   before_json, after_json, created_at
            FROM transcode_agent_confirmation_events
            WHERE job_id = ?
            ORDER BY id
            """,
            (job_id,),
        ).fetchall()


def create_transcode_agent_pending_rule(
    *,
    rule_id: str,
    rule_json: str,
    employee_id: str,
    customer_code: str,
    customer_name: str,
    business_field: str,
    target_value: str,
    condition_summary: str,
    source_task_id: int | None,
    source_excel_row: int | None,
) -> int:
    now = utcnow()
    with transcode_db_cursor() as conn:
        cursor = conn.execute(
            """
            INSERT INTO transcode_agent_pending_rules (
                rule_id, rule_json, employee_id, customer_code, customer_name,
                business_field, target_value, condition_summary,
                source_task_id, source_excel_row, status,
                created_at, updated_at, updated_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?)
            """,
            (
                rule_id,
                rule_json,
                employee_id,
                customer_code,
                customer_name,
                business_field,
                target_value,
                condition_summary,
                source_task_id,
                source_excel_row,
                now,
                now,
                employee_id,
            ),
        )
        return int(cursor.lastrowid)


def list_transcode_agent_pending_rules(
    status: str = "pending",
    *,
    employee_id: str = "",
) -> list[sqlite3.Row]:
    sql = "SELECT * FROM transcode_agent_pending_rules"
    conditions: list[str] = []
    params: list[object] = []
    if status:
        conditions.append("status = ?")
        params.append(status)
    if employee_id:
        conditions.append("employee_id = ?")
        params.append(employee_id)
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY created_at DESC, id DESC"
    with transcode_db_cursor() as conn:
        return conn.execute(sql, params).fetchall()


def get_transcode_agent_pending_rule(rule_id: int) -> sqlite3.Row | None:
    with transcode_db_cursor() as conn:
        return conn.execute(
            "SELECT * FROM transcode_agent_pending_rules WHERE id = ?",
            (rule_id,),
        ).fetchone()


def update_transcode_agent_pending_rule(
    rule_id: int,
    *,
    rule_json: str,
    customer_code: str,
    customer_name: str,
    business_field: str,
    target_value: str,
    condition_summary: str,
    updated_by: str,
) -> None:
    with transcode_db_cursor() as conn:
        conn.execute(
            """
            UPDATE transcode_agent_pending_rules
            SET rule_json = ?, customer_code = ?, customer_name = ?,
                business_field = ?, target_value = ?, condition_summary = ?,
                updated_at = ?, updated_by = ?
            WHERE id = ?
            """,
            (
                rule_json,
                customer_code,
                customer_name,
                business_field,
                target_value,
                condition_summary,
                utcnow(),
                updated_by,
                rule_id,
            ),
        )


def set_transcode_agent_pending_rule_status(
    rule_id: int,
    status: str,
    *,
    processed_by: str = "",
) -> None:
    with transcode_db_cursor() as conn:
        conn.execute(
            """
            UPDATE transcode_agent_pending_rules
            SET status = ?, processed_by = ?, processed_at = ?,
                updated_at = ?, updated_by = ?
            WHERE id = ?
            """,
            (status, processed_by, utcnow(), utcnow(), processed_by, rule_id),
        )


def upsert_transcode_agent_row_verification(
    *,
    job_id: int,
    excel_row: int,
    employee_id: str,
    action: str,
    before_code: str,
    after_code: str,
    basis: str = "",
) -> None:
    with transcode_db_cursor() as conn:
        conn.execute(
            """
            INSERT INTO transcode_agent_row_verifications (
                job_id, excel_row, employee_id, action,
                before_code, after_code, basis, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(job_id, excel_row) DO UPDATE SET
                employee_id = excluded.employee_id,
                action = excluded.action,
                before_code = excluded.before_code,
                after_code = excluded.after_code,
                basis = excluded.basis,
                created_at = excluded.created_at
            """,
            (
                job_id,
                excel_row,
                employee_id,
                action,
                before_code,
                after_code,
                basis,
                utcnow(),
            ),
        )


def get_transcode_agent_row_verification(
    job_id: int,
    excel_row: int,
) -> sqlite3.Row | None:
    with transcode_db_cursor() as conn:
        return conn.execute(
            """
            SELECT * FROM transcode_agent_row_verifications
            WHERE job_id = ? AND excel_row = ?
            """,
            (job_id, excel_row),
        ).fetchone()


def list_transcode_agent_row_verifications(
    job_id: int,
) -> list[sqlite3.Row]:
    with transcode_db_cursor() as conn:
        return conn.execute(
            """
            SELECT * FROM transcode_agent_row_verifications
            WHERE job_id = ?
            ORDER BY excel_row, id
            """,
            (job_id,),
        ).fetchall()


def prune_jobs_for_employee(employee_id: str, keep_limit: int = 500) -> list[sqlite3.Row]:
    with transcode_db_cursor() as conn:
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
            stale_ids = [row["id"] for row in rows]
            placeholders = ",".join("?" for _ in stale_ids)
            conn.execute(
                f"DELETE FROM transcode_agent_row_verifications WHERE job_id IN ({placeholders})",
                stale_ids,
            )
            conn.execute(
                f"DELETE FROM transcode_agent_confirmation_events WHERE job_id IN ({placeholders})",
                stale_ids,
            )
            conn.execute(
                f"DELETE FROM transcode_agent_confirmation_items WHERE job_id IN ({placeholders})",
                stale_ids,
            )
            conn.executemany("DELETE FROM jobs WHERE id = ?", [(row["id"],) for row in rows])
    return rows


def delete_job(job_id: int) -> sqlite3.Row | None:
    with transcode_db_cursor() as conn:
        row = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
        if row:
            conn.execute(
                "DELETE FROM transcode_agent_row_verifications WHERE job_id = ?",
                (job_id,),
            )
            conn.execute(
                "DELETE FROM transcode_agent_confirmation_events WHERE job_id = ?",
                (job_id,),
            )
            conn.execute(
                "DELETE FROM transcode_agent_confirmation_items WHERE job_id = ?",
                (job_id,),
            )
            conn.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
    return row


def list_expired_terminal_jobs(cutoff: str) -> list[sqlite3.Row]:
    with transcode_db_cursor() as conn:
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
    with transcode_db_cursor() as conn:
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
    with transcode_db_cursor() as conn:
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
