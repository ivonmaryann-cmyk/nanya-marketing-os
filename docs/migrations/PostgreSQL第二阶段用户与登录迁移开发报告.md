# PostgreSQL第二阶段用户与登录迁移开发报告

## 结论

第二阶段仅迁移`users`表。本地SQLite中的11个账号已全量复制并切换到本地PostgreSQL，密码哈希保持原值；生产服务器未切换。PP转码、通用转码、全局配置、附件和其他业务表均未进入本阶段。

## 完成范围

- 完成任务：M6本地部分；复用DB-01至DB-04的连接与版本化迁移原则，并完成`users`专用DDL、复制、核对、运行时隔离、变更捕获和回滚验证。
- 未完成任务：M6生产部分、DB-18生产观察。阻塞原因为尚未完成生产连接配置、备份恢复演练、监控告警、维护窗口和业务人工验收。
- 代码提交：`4949699`
- 数据库迁移版本：`migrations/identity/postgresql/0001_users.sql`

## 本地PostgreSQL

本地PostgreSQL沿用第一阶段实例。应用只从环境变量读取连接信息，不在仓库或文档中保存账号、密码和真实连接串。

环境变量名称：

- `IDENTITY_DATABASE_BACKEND`
- `IDENTITY_POSTGRESQL_READ_WRITE_ENABLED`
- `IDENTITY_DATABASE_URL`
- `IDENTITY_TEST_DATABASE_URL`
- `IDENTITY_DB_POOL_MIN_SIZE`
- `IDENTITY_DB_POOL_MAX_SIZE`
- `IDENTITY_DB_CONNECT_TIMEOUT_SECONDS`
- `IDENTITY_DB_STATEMENT_TIMEOUT_MS`

## 迁移与校验

```powershell
python -m fangzheng_web_app.identity_migration migrate --sqlite storage/app.db
python -m fangzheng_web_app.identity_migration verify --sqlite storage/app.db
```

本地核对结果：SQLite 11行、PostgreSQL 11行；主键、逐行摘要、密码哈希及关键状态计数全部一致，SQLite完整性检查为`ok`。迁移使用主键upsert，可重复执行且不生成重复账号；核对失败会回滚PostgreSQL事务。

## 回滚

切换前备份：`D:\Carson\.postgres\backups\identity_stage2_20260821_141257.sqlite3`。

```powershell
python -m fangzheng_web_app.identity_migration rollback-replay `
  --sqlite storage/app.db `
  --backup D:\Carson\.postgres\backups\identity_stage2_20260821_141257.sqlite3 `
  --backup-sha256 be8a868245104dad158326bc338e5f4fea9964aeb3e881c07261d23e3fad8294 `
  --confirm REPLAY-IDENTITY-CHANGES-TO-SQLITE
```

回放完成且`remaining=0`后，将本地`IDENTITY_DATABASE_BACKEND`改回`sqlite`并重启应用。测试库已完成回放及迁移对象清理演练，身份相关表剩余0张。不得直接覆盖或删除原SQLite。

## 测试

- 身份迁移单元测试：4项通过。
- 真实PostgreSQL集成测试：1项通过，覆盖重复复制、密码哈希、登录、改密、启停、角色和回滚回放。
- 全套测试：72项通过，6项因对应外部测试环境变量未配置而跳过。
- 本地运行时检查：身份后端为PostgreSQL，运行时账号数与SQLite均为11。

## 影响与下一步

仓库默认仍使用SQLite，只有本地未提交的`config/local.env`启用了PostgreSQL身份后端。现有页面、接口、密码规则和业务判断未改变，自动化库仍只通过`employee_id`逻辑关联用户，没有跨库外键。

下一步先由业务人员在本地完成真实账号的登录、改密、启停和角色页面验收，再决定是否推送并合并。生产切换必须另行配置凭据、备份、监控和维护窗口，不随本次代码提交自动发生。
