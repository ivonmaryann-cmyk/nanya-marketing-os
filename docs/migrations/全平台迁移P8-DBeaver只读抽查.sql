-- 全平台 PostgreSQL 迁移 P8：DBeaver 只读抽查
-- 仅查询；请连接开发库后整段执行。
-- 请在 P8 人工业务验收记录中填写执行时间、截图路径或结果摘要。

SELECT
  current_database() AS database_name,
  current_user AS database_user,
  version() AS postgresql_version;

-- 页面列表、邮件、订单和接口交互的数量应与系统页面统计一致。
SELECT '客户档案' AS scope, COUNT(*) AS row_count FROM automation_customers
UNION ALL
SELECT '邮件', COUNT(*) FROM mail_messages
UNION ALL
SELECT '订单案例', COUNT(*) FROM order_intake_cases
UNION ALL
SELECT '内销录单模板', COUNT(*) FROM order_entry_templates
UNION ALL
SELECT '接口调用日志', COUNT(*) FROM order_interface_call_logs
ORDER BY scope;

-- 本次 P8 邮件同步与订单/接口流程的留痕摘要。
SELECT
  id AS fetch_task_id,
  status,
  email_count,
  new_count,
  duplicate_count,
  order_count,
  completed_at
FROM mail_fetch_tasks
WHERE id = 10;

SELECT
  id AS case_id,
  action_type,
  status,
  workflow_stage,
  erp_prepare_status,
  completed_at
FROM order_intake_cases
WHERE id = 296748;

SELECT
  interface_key,
  status,
  is_mock,
  COUNT(*) AS call_count
FROM order_interface_call_logs
WHERE template_id = 14
GROUP BY interface_key, status, is_mock
ORDER BY interface_key, status;

-- 迁移后必须为 0；非 0 表示仍有未验证外键。
SELECT COUNT(*) AS unvalidated_foreign_keys
FROM pg_constraint con
JOIN pg_namespace ns ON ns.oid = con.connamespace
WHERE ns.nspname = 'public'
  AND con.contype = 'f'
  AND NOT con.convalidated;
