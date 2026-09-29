CREATE TABLE IF NOT EXISTS shennan_settlement_customer_warehouses (
    customer_code TEXT NOT NULL, customer_name TEXT NOT NULL, warehouse_code TEXT NOT NULL,
    enabled INTEGER NOT NULL DEFAULT 1, revision INTEGER NOT NULL DEFAULT 1,
    updated_by TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (customer_code, warehouse_code)
);
CREATE TABLE IF NOT EXISTS shennan_settlement_batches (
    id TEXT PRIMARY KEY, employee_id TEXT NOT NULL, job_id INTEGER, status TEXT NOT NULL,
    manifest_path TEXT NOT NULL DEFAULT '', output_path TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL, updated_at TEXT NOT NULL, confirmed_at TEXT, cancelled_at TEXT
);
CREATE TABLE IF NOT EXISTS shennan_settlement_inputs (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, input_type TEXT NOT NULL, warehouse_code TEXT NOT NULL DEFAULT '',
    filename TEXT NOT NULL, path TEXT NOT NULL, row_count INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS shennan_settlement_consumptions (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, block_no INTEGER NOT NULL, source_row INTEGER NOT NULL,
    customer_code TEXT NOT NULL, customer_material_code TEXT NOT NULL, quantity REAL NOT NULL, price REAL NOT NULL,
    matched_quantity REAL NOT NULL DEFAULT 0, remaining_quantity REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS shennan_settlement_orders (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, order_no TEXT NOT NULL, line_no TEXT NOT NULL,
    order_date TEXT NOT NULL, customer_code TEXT NOT NULL, customer_material_code TEXT NOT NULL,
    factory_part_no TEXT NOT NULL, unfulfilled_quantity REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS shennan_settlement_inventory (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, warehouse_code TEXT NOT NULL, factory_part_no TEXT NOT NULL,
    lot_no TEXT NOT NULL, available_quantity REAL NOT NULL, stagnation_date TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS shennan_settlement_details (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, consumption_id TEXT NOT NULL, block_no INTEGER NOT NULL,
    source_row INTEGER NOT NULL, customer_code TEXT NOT NULL, customer_name TEXT NOT NULL,
    category TEXT NOT NULL, order_no TEXT NOT NULL, line_no TEXT NOT NULL, order_date TEXT NOT NULL,
    customer_material_code TEXT NOT NULL, factory_part_no TEXT NOT NULL, lot_no TEXT NOT NULL,
    warehouse_code TEXT NOT NULL, price REAL NOT NULL, quantity REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS shennan_settlement_unmatched (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, block_no INTEGER NOT NULL, source_row INTEGER NOT NULL,
    customer_name TEXT NOT NULL, customer_code TEXT NOT NULL, customer_material_code TEXT NOT NULL,
    original_quantity REAL NOT NULL, matched_quantity REAL NOT NULL, unmatched_quantity REAL NOT NULL,
    price REAL NOT NULL, reason TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS shennan_settlement_events (
    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, event_type TEXT NOT NULL, actor TEXT NOT NULL,
    message TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS shennan_settlement_batch_owner ON shennan_settlement_batches(employee_id, created_at);
CREATE INDEX IF NOT EXISTS shennan_settlement_detail_order ON shennan_settlement_details(order_no, line_no);
CREATE INDEX IF NOT EXISTS shennan_settlement_detail_inventory ON shennan_settlement_details(warehouse_code, factory_part_no, lot_no);
