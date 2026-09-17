CREATE TABLE IF NOT EXISTS order_task_work_samples (
    id TEXT PRIMARY KEY,
    case_id BIGINT NOT NULL REFERENCES order_intake_cases(id),
    employee_id TEXT NOT NULL,
    work_date TEXT NOT NULL,
    minutes INTEGER NOT NULL CHECK(minutes BETWEEN 1 AND 480),
    kind TEXT NOT NULL CHECK(kind IN ('operation','rework')),
    note TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_task_work_owner_date
    ON order_task_work_samples(employee_id,work_date,case_id);
