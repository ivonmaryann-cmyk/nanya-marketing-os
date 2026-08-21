# PostgreSQL第三阶段PP与通用转码迁移开发报告

## 结论

第三阶段计划中的18张表已在本地整体迁移并切换到PostgreSQL，生产服务器未切换。共享`jobs`没有按功能拆分，因此价格计算、PDF转Excel、库存等所有后台任务的任务索引也随本阶段在本地改读PostgreSQL；任务文件仍保留在`storage/jobs`并按原路径访问。

## 完成范围

- 已完成：M7本地开发、全量复制、逐行核对、运行时切换、变更捕获、回滚回放和页面回归。
- 尚未完成：M7生产切换及稳定观察。阻塞项为生产凭据、备份恢复、监控告警、维护窗口、真实任务上传人工验收和多实例文件存储方案。
- 未进入本阶段：`settings`、`pdf_excel_ai_config_versions`、工作规划、反馈、附件和对象存储。
- 代码提交：`18ac4ca`
- 数据库迁移版本：`migrations/transcode/postgresql/0001_transcode_schema.sql`

## 环境变量

- `TRANSCODE_DATABASE_BACKEND`
- `TRANSCODE_POSTGRESQL_READ_WRITE_ENABLED`
- `TRANSCODE_DATABASE_URL`
- `TRANSCODE_TEST_DATABASE_URL`
- `TRANSCODE_DB_POOL_MIN_SIZE`
- `TRANSCODE_DB_POOL_MAX_SIZE`
- `TRANSCODE_DB_CONNECT_TIMEOUT_SECONDS`
- `TRANSCODE_DB_STATEMENT_TIMEOUT_MS`
- `TRANSCODE_MIGRATION_LOG_PATH`

连接账号、密码和真实连接串只能保存在环境变量或不提交的`config/local.env`中。

## 本地启动与迁移

本地PostgreSQL沿用第一、第二阶段实例。启动实例后执行：

```powershell
python -m fangzheng_web_app.transcode_migration migrate --sqlite storage/app.db
python -m fangzheng_web_app.transcode_migration verify --sqlite storage/app.db
```

迁移使用显式主键upsert并校准PostgreSQL序列，可重复执行；任一表核对失败会回滚整批PostgreSQL事务。默认迁移日志写入被Git忽略的`storage/migration_logs/transcode_migration.log`。

## 数据证据

- 18张表主键集合和逐行摘要全部一致。
- `jobs`156行、PP基础规则14行、Agent确认项11行、确认事件6行，其余表当前为0行。
- 切换前没有`queued`或`running`任务。
- 156条历史任务的输入和结果路径缺失均为0。
- 18张表的变更捕获触发器均已启用，本地主库待回放事件为0。
- SQLite完整性检查为`ok`，未删除或覆盖SQLite历史数据。

## 回滚

切换前备份：`D:\Carson\.postgres\backups\transcode_stage3_20260821_143651.sqlite3`。

```powershell
python -m fangzheng_web_app.transcode_migration rollback-replay `
  --sqlite storage/app.db `
  --backup D:\Carson\.postgres\backups\transcode_stage3_20260821_143651.sqlite3 `
  --backup-sha256 1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac `
  --confirm REPLAY-TRANSCODE-CHANGES-TO-SQLITE
```

重复执行回放直到`remaining=0`，再把本地`TRANSCODE_DATABASE_BACKEND`改回`sqlite`并重启应用。独立测试库已完成任务、PP规则、确认项、模型配置和规则中心变更的幂等回放，并验证迁移对象可完整清理。

## 测试和影响

- 第三阶段专项测试3项通过，覆盖DDL范围、配置锁、重复复制、任务读写、日期查询、PP规则、PP/Agent确认项、模型配置、规则中心备份恢复和回滚。
- 完整测试75项通过，7项因对应外部测试环境未配置而跳过。
- PP转码、通用转码、历史任务、PP规则和客户规则页面均返回HTTP 200。
- 仓库默认仍为SQLite，只有本地未提交配置启用PostgreSQL；生产不会因合并代码自动切换。

下一步由业务人员在本地分别上传一个PP转码任务和一个通用转码任务，确认进度、结果下载、确认项和历史任务页面，再决定是否推送合并。生产切换必须另行准备数据库凭据、备份、监控、维护窗口和文件共享方案。
