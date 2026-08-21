# PostgreSQL 第五阶段工作规划与反馈迁移开发报告

## 1. 完成范围

- 本地完成`task_categories`、`personal_tasks`和`feedback`三张表的版本化建表、全量复制、逐行核对、独立数据访问和读写切换。
- 工作规划CRUD、筛选排序、反馈状态和任务JSON备份恢复均改走第五阶段独立连接。
- SQLite继续保留，PostgreSQL变更捕获已开启，支持将切换后的变更幂等回放到SQLite。
- 未执行生产切换、文件迁移、SQLite停止写入或删除。

## 2. 提交与迁移版本

- 代码提交：`d460d95`
- PostgreSQL DDL：`migrations/planning/postgresql/0001_planning_schema.sql`
- 本地备份：`D:\Carson\.postgres\backups\planning_stage5_20260821_151456.sqlite3`
- 备份SHA-256：`1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac`

## 3. PostgreSQL 本地启动方式

沿用前四阶段已启动的本地PostgreSQL 16.15实例。应用连接信息只放在忽略提交的`config/local.env`，不得写入Git。

## 4. 环境变量清单

- `PLANNING_DATABASE_BACKEND`
- `PLANNING_POSTGRESQL_READ_WRITE_ENABLED`
- `PLANNING_DATABASE_URL`
- `PLANNING_TEST_DATABASE_URL`
- `PLANNING_DB_POOL_MIN_SIZE`
- `PLANNING_DB_POOL_MAX_SIZE`
- `PLANNING_DB_CONNECT_TIMEOUT_SECONDS`
- `PLANNING_DB_STATEMENT_TIMEOUT_MS`
- `PLANNING_MIGRATION_LOG_PATH`
- `WORK_PLANNING_BACKUP_ROOT`

## 5. 迁移与校验命令

```powershell
python -m fangzheng_web_app.planning_migration migrate --sqlite storage/app.db
python -m fangzheng_web_app.planning_migration verify --sqlite storage/app.db
```

本次结果：分类3/3、任务0/0、反馈0/0；主键和逐行摘要一致，任务分类孤儿为0，SQLite完整性为`ok`。

## 6. 回滚命令

先停止应用并把`PLANNING_DATABASE_BACKEND`改回`sqlite`，确认当前终端已设置`PLANNING_DATABASE_URL`，再执行：

```powershell
python -m fangzheng_web_app.planning_migration rollback-replay `
  --sqlite storage/app.db `
  --backup D:\Carson\.postgres\backups\planning_stage5_20260821_151456.sqlite3 `
  --backup-sha256 1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac `
  --confirm REPLAY-PLANNING-CHANGES-TO-SQLITE
```

变更超过100条时重复执行，直到`remaining`为0。不要删除SQLite或直接覆盖数据库。

## 7. 验证结果

- 第五阶段单元测试：4项通过。
- 真实PostgreSQL集成测试：1项通过，覆盖重复复制、任务CRUD、筛选排序、反馈、JSON备份恢复和回滚回放。
- 本地主库回滚演练：7条真实探针变更全部回放，剩余0，并重新启用捕获。
- 全套测试：85项通过，9项因未设置各阶段专用测试连接而按设计跳过。
- 页面冒烟：工作规划、反馈中心、反馈管理均为HTTP 200。

## 8. 影响与风险

本地计划列出的44张业务表均已读写PostgreSQL，但生产环境没有切换。任务备份JSON、任务文件和附件仍保存在本地文件系统；稳定观察、生产备份恢复和人工业务验收尚未完成，因此SQLite仍必须保留。

## 9. 下一步

先观察本地五阶段读写与回滚日志，再执行文件存储FS-02至FS-06。生产数据库切换和SQLite最终下线必须分别申请确认，不能随本次提交执行。
