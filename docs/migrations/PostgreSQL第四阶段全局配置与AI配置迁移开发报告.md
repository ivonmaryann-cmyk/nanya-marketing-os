# PostgreSQL 第四阶段全局配置与 AI 配置迁移开发报告

## 1. 完成范围

- 本地完成`settings`和`pdf_excel_ai_config_versions`两张表的版本化建表、全量复制、逐行核对、独立数据访问和读写切换。
- 管理员密码哈希、活动规则/AI版本值、规则历史和AI配置密文均原值复制，日志不输出真实配置值。
- SQLite继续保留，PostgreSQL变更捕获已开启，支持将切换后的变更幂等回放到SQLite。
- 未执行生产切换、第五阶段迁移、附件移动或SQLite删除。

## 2. 提交与迁移版本

- 代码提交：`606763c`
- PostgreSQL DDL：`migrations/configuration/postgresql/0001_configuration_schema.sql`
- 本地备份：`D:\Carson\.postgres\backups\configuration_stage4_20260821_145512.sqlite3`
- 备份SHA-256：`1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac`

## 3. PostgreSQL 本地启动方式

沿用项目前三阶段已启动的本地PostgreSQL 16.15实例。应用连接信息只放在忽略提交的`config/local.env`，不得写入Git。

## 4. 环境变量清单

- `CONFIG_DATABASE_BACKEND`
- `CONFIG_POSTGRESQL_READ_WRITE_ENABLED`
- `CONFIG_DATABASE_URL`
- `CONFIG_TEST_DATABASE_URL`
- `CONFIG_DB_POOL_MIN_SIZE`
- `CONFIG_DB_POOL_MAX_SIZE`
- `CONFIG_DB_CONNECT_TIMEOUT_SECONDS`
- `CONFIG_DB_STATEMENT_TIMEOUT_MS`
- `CONFIG_MIGRATION_LOG_PATH`
- `PDF_EXCEL_AI_CONFIG_MASTER_KEY`

## 5. 迁移与校验命令

```powershell
python -m fangzheng_web_app.configuration_migration migrate --sqlite storage/app.db
python -m fangzheng_web_app.configuration_migration verify --sqlite storage/app.db
```

本次结果：`settings` 63/63，AI配置版本0/0；主键、逐行摘要、管理员密码哈希、活动AI版本值和密文集合一致，SQLite完整性为`ok`。

## 6. 回滚命令

先停止应用并把`CONFIG_DATABASE_BACKEND`改回`sqlite`，再执行：

```powershell
python -m fangzheng_web_app.configuration_migration rollback-replay `
  --sqlite storage/app.db `
  --backup D:\Carson\.postgres\backups\configuration_stage4_20260821_145512.sqlite3 `
  --backup-sha256 1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac `
  --confirm REPLAY-CONFIGURATION-CHANGES-TO-SQLITE
```

变更超过100条时重复执行，直到`remaining`为0。不要删除SQLite或直接覆盖数据库。

## 7. 验证结果

- 第四阶段单元测试：4项通过。
- 真实PostgreSQL集成测试：1项通过，覆盖重复复制、配置读写、AI密文保存与回滚回放。
- 全套测试：80项通过，8项因未设置各阶段专用测试连接而按设计跳过。
- 页面冒烟：仪表盘、PDF/Excel AI配置、规则管理、管理员密码页均为HTTP 200。

## 8. 影响与风险

本地全局设置与PDF/Excel AI配置已读写PostgreSQL，其余未迁移模块继续使用SQLite。生产环境没有切换。生产实施前仍需配置连接与AI主密钥、建立生产备份、验证恢复、安排维护窗口并完成人工配置保存验收。

## 9. 下一步

先在本地观察第四阶段读写和回滚日志，再经代码合并与生产准入后单独申请服务器切换。第五阶段和文件对象存储迁移不得随本次提交执行。
