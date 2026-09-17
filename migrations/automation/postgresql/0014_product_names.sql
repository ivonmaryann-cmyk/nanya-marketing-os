CREATE TABLE IF NOT EXISTS product_name_mappings (
    category TEXT NOT NULL, code TEXT NOT NULL, name TEXT NOT NULL,
    details TEXT NOT NULL DEFAULT '{}', enabled INTEGER NOT NULL DEFAULT 1 CHECK(enabled IN (0,1)),
    revision INTEGER NOT NULL DEFAULT 1, updated_by TEXT NOT NULL, updated_at TEXT NOT NULL,
    PRIMARY KEY(category,code)
);
CREATE TABLE IF NOT EXISTS product_name_audit (
    id TEXT PRIMARY KEY, category TEXT NOT NULL, code TEXT NOT NULL,
    before_json TEXT NOT NULL, after_json TEXT NOT NULL, actor TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS product_name_results (
    id TEXT PRIMARY KEY, product TEXT NOT NULL, employee_id TEXT NOT NULL,
    source_spec TEXT NOT NULL, input_json TEXT NOT NULL, result TEXT NOT NULL,
    snapshot_json TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS product_name_results_owner ON product_name_results(employee_id,created_at);
