# SQLite 到 PostgreSQL 及文件存储全量迁移计划

> 文档用途：数据库与文件存储迁移的统一方案、执行清单和持续更新台账  
> 项目：南亚营销自动化平台  
> 首次编制日期：2026-08-20  
> 当前版本：v1.9
> 当前阶段：第一至第五阶段44张业务表已在本地切换到PostgreSQL；生产切换尚未开始
> 第一迁移模块：订单邮件自动化  

---

## 0. 当前实际进度快照

> 本节区分“开发工具完成”和“业务数据实际迁移完成”。两者不得混用。

| 数据库迁移阶段 | 开发状态 | 实际迁移状态 | 当前结论 |
|---|---|---|---|
| 第一阶段：订单邮件自动化20张表 | 迁移、核对、影子同步、回滚和观察工具已开发并完成真实后端业务测试 | 本地已切换；生产未切换 | 本地自动化模块读写PostgreSQL 16.15；仓库默认及生产仍保持SQLite |
| 第二阶段：用户与登录 `users` | 已完成 | 本地已迁移并切换 | 11个账号全量核对通过；生产登录读写未切换 |
| 第三阶段：PP与通用转码18张表 | 已完成 | 本地已迁移并切换 | 18表逐行核对通过；共享`jobs`整体切换；生产未切换 |
| 第四阶段：`settings`与AI配置 | 已完成 | 本地已迁移并切换 | 2表逐行核对通过；生产配置读写未切换 |
| 第五阶段：工作规划与反馈 | 已完成 | 本地已迁移并切换 | 3表逐行核对、业务回归和回滚重放通过；生产未切换 |
| 文件存储迁移 | FS-01本地抽象已开发，其余未开始 | 未开始 | 未上传、移动或删除任何真实附件 |

当前生产/本地业务事实：

- 本地未跟踪的`config/local.env`已分别启用自动化、身份、转码、全局配置和工作规划PostgreSQL后端；
- 本地PostgreSQL五个阶段的44张业务表均具备变更捕获，切换后新增、修改和删除可进入各阶段回滚日志；
- Git中的安全默认值仍为SQLite且PostgreSQL读写关闭；
- 第一至第五阶段共44张业务表仅在本地正式读写PostgreSQL，生产没有切换；
- 没有执行生产数据回放、正式切换或7至14天稳定观察；
- SQLite历史数据继续保留为回滚源，尚未停止写入或归档；本地任务文件、备份JSON和附件仍按原路径读取，禁止删除。

---

## 1. 文档使用要求

本文件不是一次性的方案说明，而是迁移工作的唯一执行台账。参与迁移的开发、测试和部署人员必须持续维护本文件，不能只修改代码而不更新迁移记录。

### 1.1 每次修改必须同步记录

每完成或变更一项迁移任务，执行人必须在对应任务下补充：

- 当前状态；
- 负责人；
- 开始和完成时间；
- 实际涉及的文件；
- 数据库迁移版本；
- Git 分支和提交号；
- 实际修改内容；
- 测试命令与测试结果；
- 数据核对结果；
- 回滚方式是否验证；
- 遗留问题和后续任务。

禁止仅将复选框改成“已完成”而不填写验证证据。

### 1.2 统一状态

| 状态 | 说明 |
|---|---|
| 未开始 | 尚未进行任何实现 |
| 进行中 | 正在开发或迁移 |
| 待验证 | 开发完成，但尚未通过完整验证 |
| 已完成 | 开发、测试、数据核对和回滚验证全部完成 |
| 已阻塞 | 因外部条件或问题无法继续，必须写明原因 |
| 已回滚 | 已撤销本项变更，必须记录回滚原因和结果 |

### 1.3 状态更新规则

1. 开始修改代码前，将对应任务更新为“进行中”。
2. 提交代码但未完成测试时，只能更新为“待验证”。
3. 只有满足该任务的全部验收条件后，才能更新为“已完成”。
4. 发生范围、架构或迁移顺序变化时，必须先更新本文档，再修改代码。
5. 所有正式切换和回滚操作必须有两人复核，并在迁移日志中记录。

---

## 2. 迁移目标与不可违反的原则

### 2.1 总体目标

逐步将现有 SQLite 数据迁移到 PostgreSQL，迁移期间允许两个数据库并存。第一阶段优先迁移“订单邮件自动化”模块，其余模块按照本文后续顺序逐步完成，最终停止 SQLite 业务写入并归档。

### 2.2 本次迁移不允许改变的内容

- 不改变现有页面业务含义；
- 不改变邮件只读抓取方式；
- 不改变邮件去重逻辑；
- 不改变邮件分流逻辑和规则优先级；
- 不改变录单、修改订单、报价、暂不分流等业务分类；
- 不改变订单任务状态的系统判断逻辑；
- 不改变内销录单模板的业务字段、必填规则和下载格式；
- 不改变当前登录、工号和数据归属逻辑；
- 不在首轮迁移中同步进行 ORM 重构、接口重构或页面重做；
- 不在未完成验证前删除 SQLite 中的历史表和数据；
- 不在数据库迁移过程中同时大规模迁移附件文件。

### 2.3 模块迁移边界原则

允许 SQLite 和 PostgreSQL 并存，但同一个业务操作不能依赖跨数据库事务，也不能在两个数据库之间直接联表。

因此迁移单位必须是完整业务边界，而不是单张表。订单邮件自动化中的邮件、附件、分流、业务任务和录单模板存在大量联表关系，必须作为一个整体迁移和切换。

### 2.4 兼容优先原则

首轮迁移优先保持数据语义一致，而不是立即使用 PostgreSQL 的所有高级类型：

- 当前以 ISO 文本保存的日期时间，第一阶段继续按文本迁移；
- 当前以 TEXT 保存的 JSON，第一阶段继续按 TEXT 迁移；
- 当前以 0/1 保存的布尔字段，第一阶段保持兼容；
- 现有整数主键原值迁移，不重新编号；
- 先完成数据库切换，再单独评估 `TIMESTAMPTZ`、`JSONB`、`BOOLEAN` 等类型优化。

这样可以避免日期归属、页面排序、JSON序列化和状态判断在迁移过程中发生隐性变化。

---

## 3. 当前系统基线

> 本节记录的是2026-08-20评估时的基线。正式迁移前必须重新采集一次，并将新结果补充到迁移日志。

### 3.1 SQLite现状

- 数据库文件：`storage/app.db`；
- 当前大小：约32MB；
- 业务表数量：44张；
- 当前数据库访问：Python标准库 `sqlite3`；
- 当前没有 PostgreSQL 驱动；
- 当前没有正式的数据库版本迁移框架；
- 当前 `journal_mode=DELETE`；
- 当前连接没有启用 SQLite 外键约束；
- 当前完整性检查结果：`ok`；
- 第一迁移模块逻辑外键扫描：未发现孤儿数据；
- 邮件唯一键扫描：未发现 `(account_id, folder, uid)` 重复记录。

### 3.2 第一模块数据规模

| 数据 | 当前数量 |
|---|---:|
| 邮箱配置 `mail_accounts` | 2 |
| 邮件 `mail_messages` | 825 |
| 附件记录 `mail_attachments` | 953 |
| 附件解析文本 `mail_attachment_texts` | 387 |
| 邮件业务任务 `mail_order_tasks` | 817 |
| 邮件抓取任务 `mail_fetch_tasks` | 7 |
| 抓取任务邮件关系 `mail_fetch_task_messages` | 1213 |
| 订单业务案例 `order_intake_cases` | 825 |
| 内销模板 `order_entry_templates` | 4 |
| 内销模板明细 `order_entry_template_lines` | 14 |
| 内销模板版本 `order_entry_template_versions` | 17 |

### 3.3 文件存储现状

- 项目 `storage` 目录总量约3GB；
- 邮件存储目录 `storage/mail_transcode` 约458MB；
- 当前邮件目录约1909个文件；
- 附件记录对应内容约180MB；
- 数据库不保存附件二进制，只保存文件路径和元数据；
- 当前数据库中的 `stored_path` 和 `eml_path` 是本机绝对路径；
- 当前邮件目录结构：

```text
storage/mail_transcode/
└── {mail_account_id}/
    └── {imap_uid}/
        ├── original.eml
        ├── attachments/
        │   └── 原始附件
        └── inline_images/
            └── 正文内嵌图片
```

### 3.4 当前数据库实现中的兼容风险

现有代码直接使用了以下 SQLite 特性，不能只替换连接字符串：

- `?` 参数占位符；
- `INTEGER PRIMARY KEY AUTOINCREMENT`；
- `INSERT OR IGNORE`；
- `INSERT OR REPLACE`；
- `GROUP_CONCAT`；
- `cursor.lastrowid`；
- `PRAGMA table_info`；
- `executescript`；
- `BEGIN IMMEDIATE`；
- `sqlite3.IntegrityError`；
- `sqlite3.OperationalError`；
- 启动时直接执行 `CREATE TABLE` 和 `ALTER TABLE`；
- 多处服务直接导入全局 `db_cursor()`。

### 3.5 当前高风险点

| 风险 | 等级 | 说明 |
|---|---|---|
| SQLite并发写锁 | 高 | 30名业务人员同时操作时可能产生锁等待和写入失败 |
| 原生SQL方言差异 | 高 | 直接切换连接会导致大量SQL运行失败 |
| 模块内跨表事务 | 高 | 按单表迁移会产生跨库联表和事务不一致 |
| 绝对文件路径 | 高 | 换服务器、目录或多实例部署后路径失效 |
| 双写不一致 | 高 | 简单依次写两个数据库可能只成功一个 |
| 日期时间语义变化 | 高 | 直接改为时区类型可能改变按天统计结果 |
| 邮箱授权密文 | 高 | 密文或密钥变化会导致企业邮箱无法登录 |
| SQLite未启用外键 | 中 | 只能依靠迁移前逻辑外键扫描发现异常 |
| PostgreSQL连接数 | 中 | 没有连接池会快速耗尽连接或降低性能 |
| 附件公开访问 | 高 | 邮件和订单附件含敏感业务信息，不能公开暴露 |

---

## 4. 目标生产架构

### 4.1 数据库迁移期间

```text
应用
├── 核心数据库访问层
│   └── SQLite：尚未迁移的原有模块
└── 自动化数据库访问层
    └── PostgreSQL：订单邮件自动化模块
```

建议环境变量：

```text
CORE_DATABASE_BACKEND=sqlite
CORE_SQLITE_PATH=storage/app.db

AUTOMATION_DATABASE_BACKEND=sqlite|shadow|postgres
AUTOMATION_DATABASE_URL=postgresql://...
```

运行模式：

| 模式 | 读取 | 主写入 | 用途 |
|---|---|---|---|
| `sqlite` | SQLite | SQLite | 当前生产模式、迁移前 |
| `shadow` | SQLite | SQLite，并通过可靠事件同步PostgreSQL | 影子迁移和结果对比 |
| `postgres` | PostgreSQL | PostgreSQL | 正式切换后 |

### 4.2 最终架构

```text
应用实例
├── PostgreSQL
│   ├── 全部业务数据
│   ├── 规则、任务和版本
│   └── 文件元数据及对象键
└── 对象存储或共享持久化存储
    ├── 原始邮件 original.eml
    ├── 普通附件
    ├── 正文内嵌图片
    └── 生成的Excel文件
```

不建议把大附件直接保存为 PostgreSQL `BYTEA`。数据库负责元数据、关系和事务，文件内容由对象存储或共享持久化存储负责。

---

## 5. 全量迁移阶段和顺序

### 5.1 总体阶段

| 阶段 | 内容 | 当前状态 | 允许进入下一阶段的条件 |
|---|---|---|---|
| M0 | 方案确认和基线冻结 | 已完成 | 本文档经负责人确认 |
| M1 | PostgreSQL基础设施与数据库访问层 | 待环境验证 | 双库连接、迁移版本、测试环境可用 |
| M2 | 第一模块SQL兼容与仓储隔离 | 待环境验证 | 自动化模块可分别运行在SQLite和PostgreSQL |
| M3 | 第一模块全量基线复制 | 实际执行未开始，工具已就绪 | 全量数据、主键、约束和文件引用核对通过 |
| M4 | 影子增量同步和影子读验证 | 实际执行未开始，工具已就绪 | 连续观察无数据差异、无业务影响 |
| M5 | 第一模块正式切换PostgreSQL | 实际执行未开始，工具已就绪 | 上线准入和回滚演练通过 |
| M6 | 用户与登录迁移 | 本地已完成 | 本地开发、迁移、切换和回滚验证通过；生产仍受准入与稳定观察约束 |
| M7 | PP转码及通用转码任务迁移 | 本地已完成 | 18表整体迁移、共享jobs切换、回滚回放和页面回归通过；生产未切换 |
| M8 | 规则中心、全局配置迁移 | 本地已完成 | 2表迁移、独立访问、密文核对、回滚回放和页面回归通过；生产未切换 |
| M9 | 工作规划、反馈和剩余模块迁移 | 本地已完成 | 3表迁移、任务备份恢复、反馈回归、回滚重放和页面验证通过；生产未切换 |
| M10 | 附件对象存储正式切换 | 未开始 | 文件双读、校验和回滚验证通过 |
| M11 | SQLite停止写入、只读归档和下线 | 未开始 | 全部业务数据迁移完成并稳定观察 |

### 5.2 建议实施节奏

以下是相对工作量估算，不代表固定上线日期：

| 工作 | 建议工作量 |
|---|---:|
| PostgreSQL基础层与迁移工具 | 3～5个工作日 |
| 第一模块SQL兼容和仓储隔离 | 5～8个工作日 |
| 自动化模块全量迁移工具与校验 | 3～5个工作日 |
| 影子同步和影子读 | 3～5个工作日 |
| 影子运行观察 | 至少7个自然日 |
| 正式切换及稳定观察 | 至少7～14个自然日 |
| 文件存储抽象和对象存储迁移 | 5～10个工作日，独立排期 |

不得为了赶时间跳过影子运行、全量核对和回滚演练。

---

## 6. 第一迁移模块：订单邮件自动化

### 6.1 业务范围

第一模块包括：

- 企业邮箱配置；
- 只读IMAP邮件抓取；
- 邮件原文、正文和附件记录；
- 附件文字识别结果；
- 手动同步与七天补充同步；
- 邮件去重；
- 通用邮件分流规则；
- 录单、修改订单、报价、暂不分流；
- 修改订单变更事项；
- 工作闭环与任务进度；
- 内销录单模板；
- 模板行、保存版本、刷新提取和下载Excel。

### 6.2 必须整体迁移的20张业务表

#### 邮件与抓取

1. `mail_accounts`
2. `mail_messages`
3. `mail_attachments`
4. `mail_attachment_texts`
5. `mail_fetch_logs`
6. `mail_fetch_tasks`
7. `mail_fetch_task_messages`
8. `mail_order_tasks`
9. `mail_transcode_jobs`

#### 业务分流与闭环

10. `order_intake_cases`
11. `order_intake_case_events`
12. `order_mail_routing_rules`
13. `order_mail_routing_rule_events`
14. `order_mail_rule_groups`
15. `order_mail_rule_keywords`
16. `order_change_tags`
17. `order_change_tag_keywords`

#### 内销录单模板

18. `order_entry_templates`
19. `order_entry_template_lines`
20. `order_entry_template_versions`

### 6.3 技术辅助表

技术辅助表不计入上述20张业务表，但迁移实现可能需要：

- `automation_schema_migrations`：记录PostgreSQL自动化库迁移版本；
- `automation_migration_outbox`：SQLite影子阶段的可靠增量事件；
- `automation_migration_inbox`：PostgreSQL幂等消费记录；
- `automation_change_log`：PostgreSQL切换后的增量回滚日志；
- `automation_metadata`：保存自动化规则初始化和引擎版本标记。

辅助表不得存放用户密码、邮箱授权码明文或附件原文。

### 6.4 `settings` 的特殊处理

全局 `settings` 表同时被报价、转码、管理员密码、AI配置和订单自动化使用，第一阶段不能整体迁移。

订单自动化当前使用的内部版本键包括但不限于：

- `order_mail_rule_seed_*`
- `order_mail_rule_cleanup_*`
- `order_change_item_refinement_*`
- `order_change_delivery_acceleration_*`
- `order_intake_rule_engine_*`

处理方案：

1. 首次迁移时，把订单自动化相关键复制到PostgreSQL的 `automation_metadata`；
2. 代码通过自动化仓储读取这些键；
3. 其他配置仍由SQLite的 `settings` 管理；
4. 不允许自动化事务同时写PostgreSQL业务表和SQLite `settings`；
5. 后续全局配置迁移时，再统一整合。

---

## 7. 详细修改计划与任务台账

## M1：PostgreSQL基础设施与数据库访问层

### DB-01 PostgreSQL环境与连接信息

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 15:36:55 +08:00
- 分支：`feature/postgresql-automation-migration`（基于 `dev@e09e2ae`）
- 修改文件：`compose.postgres.yaml`、`.env.example`
- 测试命令：配置脱敏单元测试；本机环境探测
- 测试结果：PostgreSQL 16.15隔离实例连接成功；`pg_dump`恢复到独立数据库并完成20表复核；SSL、最小权限账号和生产隔离仍待验证
- 数据核对结果：不涉及业务数据
- 回滚验证结果：默认后端保持SQLite；测试库可完整清理，备份恢复演练成功
- 遗留风险：当前实例仅监听`127.0.0.1:55432`并使用本地信任认证，只可用于隔离开发验证；仍需验证SSL和最小权限应用账号
- 目标：准备开发、测试和生产隔离的PostgreSQL实例。
- 修改内容：
  - 创建独立数据库和最小权限应用账号；
  - 开发、测试、生产使用不同数据库和账号；
  - 连接信息只能从环境变量或密钥管理读取；
  - 禁止将密码提交到Git；
  - 配置SSL连接；
  - 配置备份、保留周期和恢复演练。
- 验收：
  - 开发环境连接成功；
  - 无密码进入日志和Git；
  - 备份和恢复演练成功；
  - 应用账号不能创建无关数据库或访问系统库。
- Git提交：`a842e12`
- 验证证据：环境变量仅提交变量名和安全默认值；无密码进入Git；备份SHA256为`F62A326678EBDF230D14D3B8398B3681B653AD1EB9F3913817708DE5DD816B08`，恢复后20表核对通过；权限项待验证
- 回滚方式：移除新数据库环境配置，不影响SQLite。

### DB-02 引入PostgreSQL驱动和连接池

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 16:05:00 +08:00
- 修改文件：`requirements.txt`、`fangzheng_web_app/database/config.py`、`fangzheng_web_app/database/automation.py`
- 测试结果：驱动3.3.4、连接池3.3.1导入成功；配置、超时、脱敏测试通过；真实并发待PostgreSQL实例验证
- 建议依赖：`psycopg` 3 与 `psycopg_pool`。
- 预计涉及：`requirements.txt`、应用初始化、配置加载。
- 修改内容：
  - 增加PostgreSQL驱动；
  - 增加有上限的连接池；
  - 配置连接超时、事务超时和空闲连接回收；
  - 应用启动时做连接健康检查；
  - 日志中禁止输出完整连接串和密码。
- 验收：
  - 连接池并发测试通过；
  - 数据库断开时错误清晰；
  - 不发生无限重试或请求长时间卡死。
- Git提交：`a842e12`
- 验证证据：`psycopg 3.3.4`、`psycopg_pool 3.3.1`导入成功；连接配置不输出URL
- 回滚方式：功能开关保持 `sqlite`，卸载或停用PostgreSQL连接。

### DB-03 拆分核心库与自动化库访问入口

- [x] 状态：已完成
- 负责人：Codex
- 开始时间：2026-08-20 16:05:00 +08:00
- 完成时间：2026-08-20 19:31:00 +08:00
- 修改文件：`fangzheng_web_app/database/`及3个自动化服务入口
- 测试结果：SQLite自动化回归及真实PostgreSQL邮箱账号、邮件、分流案例、模板保存业务链路均通过；全套67项通过
- 目标：避免所有模块继续依赖同一个全局 `db_cursor()`。
- 建议结构：

```text
fangzheng_web_app/database/
├── core.py              # 尚未迁移模块，连接SQLite
├── automation.py        # 自动化模块，可路由SQLite/PostgreSQL
├── rows.py              # 统一字典行结果
├── errors.py            # 统一数据库异常
└── transactions.py      # 事务上下文
```

- 修改要求：
  - 首阶段不改业务服务的输入输出；
  - 自动化数据库异常统一转换，不在业务层捕获 `sqlite3.*`；
  - 事务发生异常必须显式回滚；
  - 读事务不能无条件提交；
  - SQLite和PostgreSQL都返回可转换为字典的行对象；
  - 禁止通过字符串替换自动转换全部SQL。
- 验收：
  - 非自动化模块仍只访问SQLite；
  - 自动化模块可通过配置选择数据库；
  - 自动化关闭或PostgreSQL不可用时，不误写其他数据库。
- Git提交：`a842e12`、`20d17d0`
- 验证证据：本地开关启用后临时邮箱仅写入PostgreSQL，SQLite账号行数保持`0→0`；Flask登录页返回200
- 数据核对结果：本地切换前20表行数、主键和逐行哈希全部一致
- 回滚验证结果：将本地后端配置恢复为`sqlite`即可回到保留的SQLite；切换前备份已生成
- 遗留风险：本地样本无真实邮件、附件、案例和模板历史数据，生产仍需代表性快照验证
- 回滚方式：自动化后端切回 `sqlite`。

### DB-04 建立版本化迁移脚本

- [x] 状态：已完成
- 负责人：Codex
- 开始时间：2026-08-20 16:20:00 +08:00
- 完成时间：2026-08-20 17:45:00 +08:00
- 修改文件：`migrations/automation/postgresql/0001_automation_schema.sql`至`0003_cutover_rollback.sql`、`automation_migration/schema.py`、`cli.py`、`tests/test_automation_postgresql_integration.py`
- 测试命令：`python -m unittest tests.test_automation_postgresql_integration -v`
- 测试结果：空PostgreSQL从0执行`0001`至`0003`成功，第二次执行无重复版本；完整测试回滚后迁移专属表和函数均为0
- 目标：替代启动时零散执行 `CREATE TABLE`、`PRAGMA` 和 `ALTER TABLE`。
- 要求：
  - 每个迁移版本不可修改历史内容，只能新增版本；
  - 支持升级状态查询；
  - 每个脚本有明确前置条件和回滚说明；
  - 自动化模块迁移版本与核心SQLite迁移版本分开；
  - 生产启动不自动执行破坏性DDL。
- 验收：
  - 空数据库可从0构建完整自动化结构；
  - 已存在数据库重复执行不会破坏数据；
  - 升级失败能明确定位版本。
- Git提交：`a842e12`、`ac266be`
- 验证证据：3条版本及校验和记录完整；20张业务表存在；`capture_changes=false`；真实PostgreSQL集成4项通过
- 数据核对结果：本任务只验证结构版本；业务数据核对由DB-11记录
- 回滚验证结果：固定确认词清理后迁移专属表和函数数量均为0
- 遗留风险：正式环境执行前仍需确认最小权限账号具备迁移所需DDL权限
- 回滚方式：按对应迁移版本说明处理。

---

## M2：第一模块SQL兼容与仓储隔离

### DB-05 自动化表PostgreSQL DDL

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 16:20:00 +08:00
- 修改文件：`migrations/automation/postgresql/0001_automation_schema.sql`
- 测试结果：静态范围测试及真实PostgreSQL空库建表通过，20张业务表全部存在；破坏性约束矩阵仍待补齐
- 修改内容：
  - 为20张业务表建立PostgreSQL DDL；
  - 保留字段含义、默认值和唯一约束；
  - 补齐明确的模块内外键；
  - 外键删除策略必须与现有业务一致，不得默认级联删除；
  - 邮箱配置删除后保留历史邮件的现有逻辑不得被外键破坏；
  - 为高频查询增加必要索引，但先不改变业务查询排序。
- 关键唯一约束：
  - `mail_messages(account_id, folder, uid)`；
  - `order_intake_cases(employee_id, mail_id)`；
  - `order_entry_templates(case_id)`；
  - `order_entry_template_lines(template_id, line_no)`；
  - `order_entry_template_versions(template_id, version_number)`。
- 验收：空库建表和约束测试全部通过。
- Git提交：`a842e12`
- 验证证据：静态范围测试确认20张业务表，无其他业务模块表
- 回滚方式：仅在空库或已备份环境执行结构回滚。

### DB-06 SQLite方言改造清单

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 17:05:00 +08:00
- 修改文件：`fangzheng_web_app/database/sql.py`、`automation.py`
- 测试结果：参数、Identity/RETURNING、冲突处理、聚合和元数据映射单元测试通过；双数据库业务矩阵待验证
- 必须逐处改造：

| SQLite写法 | PostgreSQL处理 |
|---|---|
| `?` | 使用驱动参数 `%s` 或仓储统一绑定 |
| `AUTOINCREMENT` | Identity/Sequence，并保留迁移ID |
| `INSERT OR IGNORE` | `ON CONFLICT DO NOTHING` |
| `INSERT OR REPLACE` | 明确冲突键和更新字段，禁止无条件替换整行 |
| `lastrowid` | `INSERT ... RETURNING id` |
| `GROUP_CONCAT` | `STRING_AGG`，明确排序和分隔符 |
| `PRAGMA table_info` | 版本迁移或 `information_schema` |
| `executescript` | 版本化SQL迁移 |
| `BEGIN IMMEDIATE` | PostgreSQL事务及明确行锁策略 |
| `sqlite3.*Error` | 统一数据库异常类型 |

- 重点文件：
  - `fangzheng_web_app/mail_transcode_agent/mail_store.py`
  - `fangzheng_web_app/mail_transcode_agent/mail_fetch_service.py`
  - `fangzheng_web_app/order_intake_service.py`
  - `fangzheng_web_app/order_entry_service.py`
  - `fangzheng_web_app/db.py`
- 验收：同一业务测试在SQLite和PostgreSQL结果一致。
- Git提交：`a842e12`
- 验证证据：方言单元测试通过；PostgreSQL端结果一致性待集成环境验证
- 回滚方式：保留SQLite仓储实现和后端开关。

### DB-07 自动化事务边界梳理

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 17:05:00 +08:00
- 修改文件：`fangzheng_web_app/database/automation.py`、`automation_migration/cli.py`
- 测试结果：迁移全量复制使用单事务且异常回滚；业务故障注入需PostgreSQL实例继续验证
- 必须覆盖：
  - 保存邮件和附件元数据；
  - 邮件去重；
  - 抓取任务和邮件关系；
  - 创建业务案例和分流记录；
  - 手工修改分流；
  - 规则和关键词维护；
  - 创建、保存、刷新内销模板；
  - 模板版本保存；
  - 邮箱配置编辑和删除。
- 要求：每个业务动作只能在一个数据库事务中完成，不能半成功。
- 验收：注入数据库异常后，不出现一半数据已写入的情况。
- Git提交：`a842e12`
- 验证证据：连接上下文和迁移事务均在异常时显式回滚；业务故障注入待验证
- 回滚方式：切回SQLite仓储。

### DB-08 自动化测试补齐

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 17:20:00 +08:00
- 修改文件：`tests/test_automation_migration_infrastructure.py`、`tests/test_automation_postgresql_integration.py`
- 测试结果：迁移与相关SQLite回归38项通过、PostgreSQL集成1项跳过；全套45项保留1个既有业务失败
- 必须覆盖：
  - 重复同步不新增重复邮件；
  - `BODY.PEEK`行为不改变原邮箱已读状态；
  - 手动同步、七天补充同步；
  - 日期归属、分页和分类数量；
  - 分流规则多命中和未命中；
  - 手工分流；
  - 工作闭环统计；
  - 邮件详情、正文和附件下载；
  - 内销模板创建、重新提取、编辑、保存、版本和下载；
  - 邮箱配置新增、编辑、停用和删除；
  - 权限和员工数据隔离。
- 验收：SQLite和PostgreSQL测试矩阵全部通过。
- Git提交：`a842e12`
- 验证证据：定向40项通过、PostgreSQL集成1项因无实例跳过；全套45项有1个迁移前既有失败
- 验证证据：待填写

---

## M3：全量基线复制

### DB-09 迁移前数据审计

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 16:45:00 +08:00
- 修改文件：`automation_migration/audit.py`、`snapshot.py`、`cli.py`
- 测试结果：本地`storage/app.db`完整审计通过；生产快照仍需执行
- 必查项目：
  - SQLite `integrity_check`；
  - 20张业务表行数；
  - 主键重复；
  - 唯一键重复；
  - 逻辑外键孤儿；
  - 邮箱账号与邮件归属；
  - 邮件、附件、业务案例和模板关系；
  - 所有 `eml_path`、`stored_path` 文件存在性；
  - 文件大小和SHA256；
  - 邮箱授权码密文非空且可正常解密；
  - 不输出授权码明文。
- 验收：审计报告无未解释异常。
- 审计报告路径：`outputs/migrations/sqlite-audit.json`和同名Markdown（本地忽略，不提交）
- Git提交：`a842e12`
- 数据核对结果：完整性`ok`；20表存在；6组唯一键和13组逻辑外键异常均为0；当前非零表为规则组3、规则词63、改单标签6、标签词50
- 回滚验证结果：审计使用SQLite在线备份临时快照，源库未修改；Windows文件句柄释放测试通过
- 遗留风险：本地样本无邮箱、邮件、附件、案例和模板业务记录，生产快照仍需审计

### DB-10 一致性快照与全量复制工具

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 16:35:00 +08:00
- 修改文件：`automation_migration/snapshot.py`、`copy.py`、`cli.py`
- 测试结果：真实PostgreSQL连续导入同一本地快照两次均通过；修复Psycopg批量写入必须使用Cursor的问题并增加集成测试
- 要求：
  - 使用SQLite在线备份API或一致性读事务；
  - 禁止在应用运行时直接复制正在写入的 `app.db` 文件作为迁移源；
  - 按依赖顺序导入；
  - 保留原始主键；
  - 邮箱授权密文逐字节复制；
  - 不复制密码、密钥到日志；
  - 支持断点、重复执行和幂等；
  - 导入后重置PostgreSQL序列到 `MAX(id)+1`；
  - 生成机器可读和人工可读报告。
- 验收：在测试快照上可以重复迁移且结果一致。
- Git提交：`a842e12`、`ac266be`
- 验证证据：两次复制结果一致；非零数据为规则组3、规则词63、改单标签6、标签词50、自动化元数据5
- 回滚验证结果：固定确认词清理后迁移专属表和函数数量均为0；SQLite 20表业务行数与迁移前审计一致，未执行业务数据写入
- 回滚方式：清空测试目标库并重新导入；生产必须依赖快照恢复。

### DB-11 全量核对

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 16:35:00 +08:00
- 修改文件：`automation_migration/verify.py`、`audit.py`
- 测试结果：本地快照与真实PostgreSQL的20表行数、主键和逐行哈希全部一致；备份恢复后的独立数据库再次核对通过
- 核对层级：
  1. 表行数；
  2. 主键集合；
  3. 每行规范化校验和；
  4. 外键与唯一键；
  5. 每日邮件数量；
  6. 每日各分类数量；
  7. 各任务状态数量；
  8. 规则、关键词和变更事项数量；
  9. 模板、模板行和模板版本内容；
  10. 文件路径、大小和SHA256。
- 当前数据量允许全量逐条比较，禁止只抽样。
- 验收：业务数据差异为0；所有差异均有书面解释并经负责人确认。
- Git提交：`a842e12`
- 验证报告：`outputs/migrations/postgresql-local-migration-first.json`和`postgresql-local-migration-second.json`均为`ok=true`
- 回滚验证结果：核对失败会回滚整批事务；隔离测试库完整清理及备份恢复均已演练
- 遗留风险：本地快照的邮件、附件、案例和模板数据为0，不能替代正式迁移前的代表性数据核对

---

## M4：影子同步与影子读

### DB-12 可靠增量同步

- [ ] 工具开发状态：待PostgreSQL环境验证；实际增量同步未开始
- 负责人：Codex
- 开始时间：2026-08-20 16:35:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 修改文件：`automation_migration/outbox.py`、`sync.py`、`db.py`、`0002_shadow_sync.sql`
- 测试结果：真实PostgreSQL短时断连恢复通过：首次投递失败1、待重试1，恢复后应用1，重复投递识别1
- 数据核对结果：目标业务记录1、Inbox记录1，重复投递未产生第二条业务数据
- 回滚验证结果：断连期间SQLite业务记录保留；停止同步器后仍由SQLite提供业务
- 遗留风险：长时间断线、较大积压、告警和批量追平性能仍待验证
- 推荐方案：SQLite事务内写业务数据和 `automation_migration_outbox`，后台同步器幂等写入PostgreSQL。
- 禁止方案：业务代码简单执行“先写SQLite、再写PostgreSQL”，因为第二次写入失败会造成静默分叉。
- 要求：
  - 每个事件有唯一ID；
  - PostgreSQL消费幂等；
  - 失败可重试；
  - 失败次数和最后错误可查询；
  - 有积压监控；
  - 不把附件二进制写入outbox；
  - 不记录敏感密文原文到普通日志。
- 验收：断开PostgreSQL后SQLite业务正常，恢复后积压可自动补齐。
- Git提交：`2467e15`、`ac266be`
- 验证证据：真实连接故障注入、恢复和同事件重复投递完成；全套66项通过
- 回滚方式：停止同步器，业务继续使用SQLite。

### DB-13 影子读比较

- [ ] 工具开发状态：待PostgreSQL环境验证；实际7天影子运行未开始
- 负责人：Codex
- 开始时间：2026-08-20 16:35:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 修改文件：`automation_migration/shadow.py`、`cli.py`
- 测试结果：真实PostgreSQL单次影子比较10项、差异0；修复PostgreSQL不接受未加引号`day`别名的问题；连续7天尚未完成
- 数据核对结果：影子日志仅保存计数、哈希、耗时和差异标记，无业务原文
- 回滚验证结果：影子命令不接管页面结果，停止后台命令即可取消
- 遗留风险：下载Excel关键单元格和人工页面结果仍需真实环境验收
- 规则：页面仍使用SQLite结果，同时读取PostgreSQL并记录差异，不把PostgreSQL差异结果展示给业务。
- 比较内容：
  - 日期栏数量；
  - 分类Tab数量；
  - 邮件列表排序和分页；
  - 工作闭环；
  - 邮件详情；
  - 规则命中；
  - 模板内容和版本；
  - 下载文件关键单元格。
- 验收：连续至少7天无未解释差异，且影子读取不显著增加页面响应时间。
- Git提交：`2467e15`、`ac266be`
- 验证证据：本地运行`0545253e64c346d39d5a2f9372cdd663`的10项指标全部匹配；连续7天真实观察尚未开始

### DB-14 性能与并发测试

- [ ] 工具开发状态：工具已就绪；实际性能与并发测试未开始
- 负责人：Codex
- 开始时间：2026-08-20 16:35:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 修改文件：`automation_migration/performance.py`、`cli.py`
- 测试结果：真实PostgreSQL只读基准完成，30并发、600请求、600成功、错误率0
- 数据核对结果：不涉及业务迁移
- 回滚验证结果：基准为只读，停止命令即可；未写入压测数据
- 遗留风险：阈值待负责人批准；写入并发、邮件同步并发和附件解析并发尚未执行
- 场景：
  - 30名业务同时浏览邮件列表；
  - 多人同时修改分流和模板；
  - 后台邮件同步与页面访问并发；
  - 七天补充同步；
  - 批量重新提取模板；
  - 附件解析任务与数据库写入并发。
- 监控：连接池、慢SQL、锁等待、错误率、P95响应时间、同步积压。
- 验收标准：由项目负责人和开发共同填写具体阈值后执行。
- 阈值：待填写
- Git提交：`2467e15`
- 结果：P50 8.27ms、P95 14.2ms、最大33.74ms；因阈值未批准，DB-14仍不得标记完成

---

## M5：正式切换与回滚

### DB-15 切换前准入检查

- [ ] 工具开发状态：检查器已就绪；实际准入条件未满足
- 负责人：Codex
- 开始时间：2026-08-20 16:35:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 修改文件：`automation_migration/preflight.py`、`config/automation_preflight.example.json`、`cli.py`
- 测试结果：空证据默认拒绝，完整合格证据单元测试通过；示例执行返回16个阻塞项和退出码2
- 数据核对结果：本地全量核对、单次影子和备份恢复已通过；仍未满足代表性生产快照、连续7天影子、人工验收、监控告警和三名负责人确认
- 回滚验证结果：检查器只读，不修改数据库或开关
- 遗留风险：全部准入条件满足前DB-16必须保持未开始
- 必须全部满足：
  - 代码评审完成；
  - 自动化测试通过；
  - 数据全量核对通过；
  - 影子运行至少7天；
  - PostgreSQL备份恢复演练通过；
  - SQLite回滚演练通过；
  - 文件路径和下载验证通过；
  - 邮箱抓取、分流、模板和下载人工验收通过；
  - 监控和告警已配置；
  - 切换负责人、复核人和回滚负责人明确。
- Git提交：`2467e15`
- 验证证据：示例证据执行结果`passed=false`，16个阻塞项，退出码2；未授权DB-16

### DB-16 正式切换步骤

- [~] 实际执行状态：本地第一阶段已切换；生产未开始
- 负责人：Codex
- 开始时间：2026-08-20 17:20:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 执行边界：用户明确授权跳过DB-15，仅把本地第一阶段切到本地PostgreSQL；未执行生产切换
- 修改文件：`automation_migration/cutover.py`、`cli.py`、`database/automation.py`、`tests/test_automation_postgresql_integration.py`、本地未跟踪`config/local.env`
- 测试结果：本地PostgreSQL业务链路及应用启动通过；全套67项通过
- 数据核对结果：本地切换前20表全量核对`ok=true`，附件缺失和哈希异常均为0
- 回滚验证结果：SQLite一致性备份和PostgreSQL切换前备份已生成；SQLite SHA256为`AF9E3972F2EF5CA18B81A4768F9603FF6601231D916B986A144E6FA716FDBFD9`，PostgreSQL备份SHA256为`87E2095A1482552E1FB3D91DC9251309CA9B3BF8CF5EFF8A0EBEF1C5AF7E646A`；SQLite保留且未删除
- 遗留风险：本地样本业务数据稀疏；生产维护窗口、人工验收、监控、权限、SSL和7天影子仍未完成
- 推荐步骤：
  1. 通知进入自动化短暂维护窗口；
  2. 暂停自动化模块写入；
  3. 确认所有outbox事件已同步；
  4. 执行最终增量复制；
  5. 执行全量关键校验；
  6. 创建SQLite一致性备份；
  7. 创建PostgreSQL切换前快照；
  8. 将 `AUTOMATION_DATABASE_BACKEND` 切换为 `postgres`；
  9. 运行健康检查；
  10. 运行邮件同步、分流、模板保存和下载冒烟测试；
  11. 恢复业务访问；
  12. 持续监控错误、性能和数据差异。
- 实际切换时间：2026-08-20 19:29:00 +08:00（仅本地）
- 切换结果：本地第一阶段成功；生产未切换
- 自检完成时间：2026-08-20 21:10:00 +08:00
- 自检修复：修复未配置邮箱时页面查询中的PostgreSQL空参数类型推断错误；`/order-automation`、`/order-automation/rules`和`/mail-transcode/accounts`均在本地PostgreSQL配置下返回正常页面
- 自检证据：20表数据量、主键及行哈希全部一致；附件缺失和哈希异常均为0；20张表均存在变更捕获触发器；真实新增和删除分别产生`insert`、`delete`日志，SQLite对应业务表未发生写入；全套67项测试通过且无跳过
- Git提交：`c8f9d97`、`20d17d0`、`3e798e3`
- 部署版本：本地`3e798e3`；未推送、未部署生产

### DB-17 回滚能力

- [ ] 实际演练状态：未开始（回滚工具已开发）
- 负责人：Codex
- 开始时间：2026-08-20 17:20:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 执行边界：仅开发变更日志和隔离回放工具；未回放生产数据
- 修改文件：`0003_cutover_rollback.sql`、`automation_migration/rollback.py`、`outbox.py`
- 测试结果：幂等插入/更新、重复删除、Outbox抑制和异常整批回滚测试通过；真实PostgreSQL回放待验证
- 数据核对结果：测试回放后SQLite记录一致且未产生循环Outbox；未使用生产数据
- 回滚验证结果：隔离SQLite故障注入通过；正式上线前完整演练仍未执行
- 本地自检补充：`capture_changes=true`；20张业务表触发器均存在；真实应用新增和删除均成功写入变更日志。验证产生的临时业务数据和日志已清理，当前未回放日志为0。该证据只验证捕获链路，不等同于完整回滚演练
- 遗留风险：完整回放到隔离SQLite、权限、敏感日志表保护和恢复时长仍需真实实例验证
- 回滚前提：正式切换后PostgreSQL新增的写入必须有可回放变更日志，否则不能安全切回旧SQLite。
- 回滚步骤：
  1. 暂停自动化写入；
  2. 导出PostgreSQL切换后的新增和修改；
  3. 按幂等方式回放到SQLite；
  4. 执行关键数据核对；
  5. 后端开关切回 `sqlite`；
  6. 冒烟测试；
  7. 恢复访问；
  8. 保存故障现场和日志。
- 要求：正式上线前至少完整演练一次。
- 演练时间：待填写
- 演练结果：待填写（仅完成隔离SQLite单元测试）
- Git提交：`c8f9d97`

### DB-18 稳定观察

- [ ] 正式观察状态：未开始（观察工具已开发）
- 负责人：Codex
- 开始时间：2026-08-20 17:20:00 +08:00
- 分支：`feature/postgresql-automation-migration`
- 执行边界：仅开发观察报告和门禁；正式观察期尚未开始
- 修改文件：`automation_migration/observation.py`、`cli.py`
- 测试结果：7个不同日期、全部人工通过、Outbox为0和影子差异为0的门禁测试通过
- 数据核对结果：无正式观察报告
- 回滚验证结果：观察工具只读，不修改业务数据或后端开关
- 遗留风险：正式切换后7至14天观察尚未发生，SQLite只读30天保留期未开始
- 观察期：正式切换后至少7～14天。
- 观察内容：
  - 数据库错误率；
  - 连接池使用；
  - 慢SQL；
  - 邮件去重；
  - 邮件日期归属；
  - 分流数量；
  - 模板保存、版本和下载；
  - 文件下载和路径；
  - 用户数据隔离。
- SQLite原表至少保留只读30天，不立即删除。
- 观察结论：待填写（正式观察尚未开始）
- Git提交：`c8f9d97`

---

## 8. 文件存储迁移与上线前处理

## 8.1 当前处理结论

数据库迁移和附件文件迁移必须拆成两个可独立回滚的阶段：

1. 第一模块切换PostgreSQL时，附件文件暂时留在当前目录；
2. PostgreSQL继续保存附件元数据和现有路径；
3. 数据库稳定后再切换文件存储；
4. 业务正式上线前，文件必须迁出开发者电脑和项目源码目录。

### 8.2 上线最低要求

如果对象存储暂时不能及时上线，至少必须满足：

- 文件存放在服务器持久化卷或共享NAS；
- 应用重新部署不会删除文件；
- 多应用实例可以访问同一份文件；
- 有每日备份和恢复验证；
- 数据库不再依赖 `/Users/...` 绝对路径；
- 文件目录不暴露为无鉴权的静态目录。

### 8.3 推荐目标：对象存储

优先支持S3兼容接口，可选MinIO、OSS、COS或其他对象存储。具体厂商由部署环境决定，不写死在业务代码中。

推荐对象键：

```text
mail/{account_id}/{uid}/original.eml
mail/{account_id}/{uid}/attachments/{attachment_id}/{safe_filename}
mail/{account_id}/{uid}/inline/{attachment_id}/{safe_filename}
order-entry/{case_id}/versions/{version}/template.xlsx
order-entry/{case_id}/exports/{export_id}.xlsx
```

### 8.4 文件元数据设计

建议在现有附件记录基础上增加或统一以下字段：

- `storage_backend`：`local`、`shared_fs`、`s3`；
- `object_key`：稳定的相对对象键；
- `original_filename`：原文件名；
- `content_type`；
- `size_bytes`；
- `sha256`；
- `storage_status`：待迁移、已复制、已校验、已切换、失败；
- `migrated_at`；
- `migration_error`。

旧的 `stored_path` 在过渡期保留，不能一开始就删除。

### 8.5 文件详细任务

### FS-01 文件存储抽象

- [ ] 状态：待验证
- 负责人：Codex
- 开始时间：2026-08-20 16:50:00 +08:00
- 修改文件：`fangzheng_web_app/file_storage/`及3个附件读取点
- 测试结果：原子保存、打开、存在性、校验和、删除接口、对象键越界防护和旧路径回退测试通过
- 数据核对结果：本地样本无附件，真实附件大小/SHA256与下载验证待执行
- 回滚验证结果：移除对象键配置后仍从旧路径读取；未搬迁或删除任何文件
- 遗留风险：对象存储适配器、受控临时链接和连续7天双读观察属于后续任务
- 目标：业务代码不再直接 `Path(stored_path).read_bytes()`。
- 统一接口至少包括：
  - `save(stream, object_key, metadata)`；
  - `open(object_key)`；
  - `exists(object_key)`；
  - `delete(object_key)`；
  - `checksum(object_key)`；
  - `temporary_download_url(object_key)` 或受控下载响应。
- 要求：同时支持旧本地路径和新对象存储。
- 验收：不修改业务页面即可读取旧附件和新附件。
- Git提交：`a842e12`
- 验证证据：存储接口和双读单元测试通过；现有3处附件读取改用兼容入口；未写入对象存储

### FS-02 新文件双写或直接写新存储

- [ ] 状态：未开始
- 负责人：待填写
- 推荐：新抓取文件写入对象存储，同时保留必要的本地回退期。
- 要求：
  - 先写临时对象，成功后原子确认；
  - 数据库记录只在文件写入成功后提交；
  - 失败时不能生成指向不存在对象的记录；
  - 文件名必须安全处理；
  - 相同邮件重复同步不能重复生成业务附件记录。
- 验收：断网和对象存储异常注入测试通过。
- Git提交：待填写
- 验证证据：待填写

### FS-03 历史文件清单与复制

- [ ] 状态：未开始
- 负责人：待填写
- 范围：
  - `original.eml`；
  - `attachments`；
  - `inline_images`；
  - 已保存的内销模板版本；
  - 需要长期保留的下载Excel。
- 要求：
  - 从数据库生成迁移清单；
  - 对每个文件记录源路径、目标键、大小和SHA256；
  - 支持断点续传；
  - 重复运行不重复上传；
  - 不因单个文件失败终止整个批次；
  - 生成失败清单并可重试。
- 验收：历史文件100%有结果，失败项必须为0或经书面确认。
- 迁移报告：待填写

### FS-04 文件双读验证

- [ ] 状态：未开始
- 负责人：待填写
- 读取顺序：优先对象存储，找不到时回退旧路径，并记录回退日志。
- 核对：文件存在、大小、SHA256、下载文件名、MIME类型。
- 验收：连续至少7天无未解释回退和校验差异。
- 验证证据：待填写

### FS-05 文件正式切换

- [ ] 状态：未开始
- 负责人：待填写
- 要求：
  - 新文件只写新存储；
  - 正式读取只使用对象键；
  - 旧路径仍保留只读回退开关；
  - 切换前创建本地文件备份；
  - 旧文件至少保留30天。
- 切换时间：待填写
- 结果：待填写

### FS-06 安全、备份和保留策略

- [ ] 状态：未开始
- 负责人：待填写
- 必须确认：
  - 存储桶禁止公开访问；
  - 下载必须经过登录和业务权限校验；
  - 临时下载链接有短有效期；
  - 服务端加密开启；
  - 访问密钥由密钥管理保存；
  - 日志不输出邮件正文、附件内容和签名链接；
  - 有版本控制或防误删除能力；
  - 有备份、恢复和灾难恢复演练；
  - 明确邮件、附件、模板的保留周期；
  - 明确删除邮箱配置是否保留历史邮件和附件。
- 验收：安全评审和恢复演练完成。
- 结果：待填写

---

## 9. 后续数据库迁移顺序

## 9.1 第二阶段：用户与登录

优先级较高，因为下个月约有30名业务人员使用。建议第一模块稳定后尽快迁移：

- `users`

注意事项：

- 密码哈希必须原样迁移；
- 不允许重置用户密码作为迁移手段；
- 登录、修改密码、启停用户和角色必须回归；
- 自动化库只通过 `employee_id` 关联用户，不在双库阶段跨库外键。

状态：本地已完成（2026-08-21）；生产未开始。

- 负责人：Codex
- 开始时间：2026-08-21 13:45:00 +08:00
- 完成时间：2026-08-21 14:17:17 +08:00
- 当前范围：独立身份库配置、`users`版本化迁移、密码哈希原样复制、数据核对、回滚验证和认证回归
- 当前边界：仅本地登录读写切换；仓库默认仍为SQLite；生产未切换；未迁移其他表，也未建立跨库外键
- 修改文件：`.env.example`、`database/config.py`、`database/identity.py`、`database/__init__.py`、`db.py`、`identity_migration/`、`migrations/identity/postgresql/0001_users.sql`、`tests/test_identity_migration.py`、`tests/test_identity_postgresql_integration.py`
- Git提交：`4949699`
- 数据库迁移版本：`0001_users.sql`
- 测试命令与结果：`python -m unittest tests.test_identity_migration -v`为4项通过；配置`IDENTITY_TEST_DATABASE_URL`后运行`tests.test_identity_postgresql_integration`为1项通过；完整`python -m unittest discover -s tests -p 'test_*.py'`为72项通过、6项按环境跳过
- 数据核对证据：本地SQLite与PostgreSQL均为11行；主键、逐行摘要、密码哈希、启用账号数、管理员数、强制改密数全部一致；SQLite完整性为`ok`
- 回滚验证结果：独立测试库完成PostgreSQL变更日志向临时SQLite幂等回放；随后删除测试迁移对象，剩余身份表为0；本地SQLite备份为`D:\Carson\.postgres\backups\identity_stage2_20260821_141257.sqlite3`，SHA256为`be8a868245104dad158326bc338e5f4fea9964aeb3e881c07261d23e3fad8294`
- 遗留风险：生产尚未配置连接、备份、监控和切换窗口；本地仅完成自动化测试及运行时计数核对，仍需业务人员以真实账号人工验收登录、改密、启停和角色页面；本地变更捕获已启用，正式回退时必须先回放未处理日志

## 9.2 第三阶段：PP转码与通用转码任务

建议整体评估以下18张表：

- `jobs`
- `pp_transcode_base_rules`
- `pp_transcode_customer_rules`
- `pp_transcode_rule_changes`
- `pp_transcode_confirmation_items`
- `transcode_model_configs`
- `transcode_agent_confirmation_items`
- `transcode_agent_confirmation_events`
- `transcode_agent_pending_rules`
- `transcode_agent_row_verifications`
- `transcode_agent_rule_overrides`
- `transcode_customer_rule_changes`
- `transcode_customer_rule_overrides`
- `transcode_rule_center_asset_overrides`
- `transcode_rule_center_base_overrides`
- `transcode_rule_center_changes`
- `transcode_rule_center_confirmation_overrides`
- `transcode_rule_center_lookup_overrides`

关键风险：`jobs` 是共享任务表，不能只迁某些行而让相同ID空间同时在两个数据库增长。必须先确认所有使用 `jobs` 的功能边界，再整体切换。

状态：本地已完成（2026-08-21）；生产未开始。

- 负责人：Codex
- 开始时间：2026-08-21 14:20:00 +08:00
- 完成时间：2026-08-21 14:42:16 +08:00
- 当前范围：计划列出的18张表整体迁移；`jobs`不得按功能拆分；任务文件继续保留在现有`storage/jobs`
- 当前边界：第四阶段`settings`和`pdf_excel_ai_config_versions`仍留在SQLite；不迁移附件、不删除SQLite历史任务、不执行生产切换
- 修改文件：`.env.example`、`database/config.py`、`database/transcode.py`、`database/sql.py`、`database/__init__.py`、`db.py`、`job_control.py`、`pp_transcode_rules.py`、`transcode_customer_rule_admin.py`、`transcode_rule_center.py`、`transcode_migration/`、`migrations/transcode/postgresql/0001_transcode_schema.sql`、`tests/test_transcode_migration.py`、`tests/test_transcode_postgresql_integration.py`
- Git提交：`18ac4ca`
- 数据库迁移版本：`0001_transcode_schema.sql`
- 测试命令与结果：第三阶段单元及真实PostgreSQL集成3项通过；完整`python -m unittest discover -s tests -p 'test_*.py'`为75项通过、7项按外部环境跳过
- 数据核对结果：18表主键与逐行摘要全部一致；`jobs`156、PP基础规则14、Agent确认项11、确认事件6，其余当前为0；切换前无`queued/running`任务；156条任务输入和结果文件路径缺失均为0；SQLite完整性为`ok`
- 页面核对结果：PP转码、通用转码、历史任务、PP规则、客户规则五个页面均返回HTTP 200
- 回滚验证结果：18表触发器全部存在，独立测试库完成多表变更向临时SQLite幂等回放并完整清理；本地主库变更捕获已启用且待回放事件为0；备份`D:\Carson\.postgres\backups\transcode_stage3_20260821_143651.sqlite3`的SHA256为`1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac`
- 遗留风险：共享`jobs`切换影响所有后台任务页面而不只PP；生产连接、备份恢复、监控、维护窗口和人工任务上传验收尚未完成；任务文件仍是本地文件，尚不支持多实例共享；本地变更捕获启用期间回退前必须先回放日志

## 9.3 第四阶段：全局配置与AI配置

- 开发状态：已完成
- 实际迁移状态：本地已迁移并切换；生产未开始

- `settings`
- `pdf_excel_ai_config_versions`

关键风险：`settings` 混合管理员密码、报价规则、转码规则和自动化版本标记。迁移前必须先按业务域拆分读取入口，不能直接把部分键迁移后仍让代码通过同一事务访问两库。

前置条件：第9.1用户与登录、第9.2 PP与通用转码已完成本地迁移；订单自动化种子标记已由第一阶段的`automation_metadata`隔离，不再与全局`settings`共用访问入口。

- 负责人：Codex
- 开始时间：2026-08-21 14:30:00 +08:00
- 完成时间：2026-08-21 14:58:00 +08:00
- 当前范围：整表迁移`settings`与`pdf_excel_ai_config_versions`；管理员密码哈希、规则活动版本、规则历史和AI配置密文保持原值
- 访问边界：订单自动化内部种子标记继续使用第一阶段`automation_metadata`；全局`get_setting/set_setting`和PDF AI配置改走第四阶段独立入口
- 当前边界：仅本地切换；不迁移第五阶段工作规划与反馈，不移动文件，不删除SQLite，不执行生产切换
- 修改文件：`.env.example`、`fangzheng_web_app/database/config.py`、`database/configuration.py`、`database/__init__.py`、`db.py`、`ai_repair_config.py`、`configuration_migration/`、`migrations/configuration/postgresql/0001_configuration_schema.sql`、`tests/test_configuration_migration.py`、`tests/test_configuration_postgresql_integration.py`
- Git提交：`606763c`
- 数据库迁移版本：`0001_configuration_schema.sql`
- 测试命令与结果：第四阶段单元4项及真实PostgreSQL集成1项通过；完整`python -m unittest discover -s tests -p 'test_*.py'`为80项通过、8项按外部测试连接配置跳过
- 数据核对证据：`settings`为63/63、`pdf_excel_ai_config_versions`为0/0；两表主键和逐行SHA-256摘要一致；管理员密码哈希、活动AI版本设置和AI密文集合一致；SQLite完整性为`ok`
- 页面核对结果：仪表盘、PDF/Excel AI配置、规则管理和管理员密码四个页面均返回HTTP 200
- 回滚验证结果：独立测试库完成设置新增/更新及AI配置密文写入向临时SQLite幂等回放，剩余事件为0；本地主库2表变更捕获已启用且待回放事件为0；备份`D:\Carson\.postgres\backups\configuration_stage4_20260821_145512.sqlite3`的SHA256为`1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac`
- 遗留风险：生产连接、备份恢复、维护窗口、监控和人工配置保存验收尚未完成；AI配置主密钥仍必须由部署环境单独提供；本地变更捕获启用期间回退前必须先回放日志；当前SQLite不得删除

## 9.4 第五阶段：工作规划与反馈

- `task_categories`
- `personal_tasks`
- `feedback`

该阶段数据量小、风险较低，适合最后迁移。

- 开发状态：已完成
- 实际迁移状态：本地已迁移并切换；生产未开始
- 负责人：Codex
- 开始时间：2026-08-21 15:05:00 +08:00
- 完成时间：2026-08-21 15:19:00 +08:00
- 当前范围：仅整表迁移`task_categories`、`personal_tasks`、`feedback`
- 访问边界：工作规划、任务备份恢复和反馈数据改走第五阶段独立入口；备份JSON文件仍保留现有路径
- 当前边界：仅本地切换；不移动或删除文件，不删除SQLite，不执行生产切换或SQLite最终下线
- 修改文件：`.env.example`、`fangzheng_web_app/database/config.py`、`database/planning.py`、`database/__init__.py`、`db.py`、`task_backup.py`、`planning_migration/`、`migrations/planning/postgresql/0001_planning_schema.sql`、`tests/test_planning_migration.py`、`tests/test_planning_postgresql_integration.py`
- Git提交：`d460d95`
- 数据库迁移版本：`0001_planning_schema.sql`
- 测试命令与结果：第五阶段单元4项及真实PostgreSQL集成1项通过；完整`python -m unittest discover -s tests -p 'test_*.py'`为85项通过、9项按外部测试连接配置跳过
- 数据核对证据：`task_categories`为3/3、`personal_tasks`为0/0、`feedback`为0/0；三表主键和逐行SHA-256摘要一致；任务分类孤儿为0；SQLite完整性为`ok`
- 业务与页面核对结果：真实PostgreSQL完成分类、任务、逾期筛选、排序、反馈状态、任务JSON备份恢复和清理；工作规划、反馈中心、反馈管理三个页面均返回HTTP 200
- 回滚验证结果：独立测试库完成完整CRUD和备份恢复向临时SQLite回放；本地主库实际产生7条探针变更并全部回放到SQLite，剩余0，随后重新开启捕获；备份`D:\Carson\.postgres\backups\planning_stage5_20260821_151456.sqlite3`的SHA256为`1c430d4e291b5be107ecaa5f4d3b6aea799a2610632cd1f77c37e803cf94fbac`
- 遗留风险：生产连接、备份恢复、维护窗口、监控和人工业务验收尚未完成；备份JSON和附件仍是本地文件；虽已完成44张业务表本地迁移，但文件存储、稳定观察和负责人确认未完成，SQLite不得下线或删除

## 9.5 SQLite最终下线

只有以下条件全部满足后才能停止SQLite：

- 44张业务表均已迁移；
- 所有模块已停止SQLite写入；
- PostgreSQL运行稳定；
- 文件存储已迁出开发机和项目目录；
- 历史数据全量核对通过；
- 备份和恢复演练通过；
- SQLite只读观察期结束；
- 项目负责人书面确认。

最终操作：

1. 生成最终SQLite一致性备份；
2. 记录文件SHA256；
3. 将SQLite标记为只读归档；
4. 从运行代码中移除SQLite写入；
5. 保留必要审计周期后再决定是否彻底下线；
6. 不得直接删除历史数据库文件。

---

## 10. 数据验证清单

### 10.1 结构验证

- [ ] 20张自动化业务表全部存在；
- [ ] 字段、默认值和非空约束符合方案；
- [ ] 主键和序列正确；
- [ ] 唯一约束正确；
- [ ] 模块内外键无孤儿；
- [ ] 索引已创建；
- [ ] 迁移版本记录完整。

### 10.2 数据验证

- [ ] 每张表行数一致；
- [ ] 主键集合一致；
- [ ] 规范化行哈希一致；
- [ ] 邮箱授权密文一致；
- [ ] 邮件去重键一致；
- [ ] 邮件正文和HTML一致；
- [ ] 附件元数据一致；
- [ ] 规则和关键词一致；
- [ ] 手工分流结果一致；
- [ ] 任务状态一致；
- [ ] 模板、行和版本一致。

### 10.3 业务验证

- [ ] 新增邮箱配置；
- [ ] 编辑邮箱配置；
- [ ] 删除邮箱配置但保留历史业务数据；
- [ ] IMAP连接测试；
- [ ] 手动同步；
- [ ] 七天补充同步；
- [ ] 重复同步不新增重复邮件；
- [ ] 原邮箱已读状态不改变；
- [ ] 日期筛选和日历；
- [ ] 10/20/50分页；
- [ ] 分类Tab数量；
- [ ] 待业务分流；
- [ ] 录单、修改订单、报价和暂不分流；
- [ ] 修改订单细分标签；
- [ ] 邮件详情返回来源页面；
- [ ] 工作闭环；
- [ ] 附件下载；
- [ ] 提取到内销模板；
- [ ] 刷新重新提取；
- [ ] 在线编辑、保存和版本；
- [ ] 下载Excel；
- [ ] 非100%明确字段继续留空。

### 10.4 文件验证

- [ ] 所有 `original.eml` 存在；
- [ ] 所有数据库附件记录对应文件存在；
- [ ] 大小一致；
- [ ] SHA256一致；
- [ ] 中文文件名下载正确；
- [ ] PDF、Excel、图片可打开；
- [ ] 正文内嵌图片可显示；
- [ ] 对象存储不可匿名访问；
- [ ] 旧路径回退受监控；
- [ ] 备份可恢复。

---

## 11. 监控和告警要求

### 11.1 PostgreSQL

- 连接池已用和等待连接数；
- 活跃连接数；
- 慢SQL；
- 锁等待和死锁；
- 事务回滚率；
- 数据库大小和增长；
- 备份成功状态；
- 复制或同步积压；
- 错误率和P95响应时间。

### 11.2 邮件与文件

- 邮件同步成功、失败和新增数量；
- 去重数量；
- 附件保存失败；
- 附件解析失败；
- 文件不存在；
- SHA256不一致；
- 对象存储请求失败；
- 本地旧路径回退次数；
- 文件存储容量和增长。

### 11.3 告警安全要求

告警中不得包含：

- 邮箱授权码；
- 数据库密码；
- 对象存储密钥；
- 完整邮件正文；
- 敏感附件内容；
- 可长期访问的下载链接。

---

## 12. 禁止事项

- 禁止未经备份直接修改或删除生产SQLite数据；
- 禁止一次性把44张表全部迁移；
- 禁止在第一阶段把20张自动化表拆到两个数据库；
- 禁止简单替换数据库连接字符串后直接上线；
- 禁止依赖自动字符串替换转换全部SQL；
- 禁止双写失败后静默忽略；
- 禁止以抽样代替当前小数据量下的全量核对；
- 禁止迁移时重新生成邮箱授权码密文；
- 禁止迁移时修改业务日期和时区语义；
- 禁止把附件放入Git或公开静态目录；
- 禁止把开发机绝对路径带到生产环境；
- 禁止没有回滚演练就正式切换；
- 禁止切换后立即删除SQLite表和本地附件；
- 禁止完成代码但不更新本文档。

---

## 13. 迁移决策记录

| 编号 | 日期 | 决策 | 原因 | 确认人 |
|---|---|---|---|---|
| ADR-001 | 2026-08-20 | 采用SQLite与PostgreSQL分模块并存迁移 | 降低一次性切换风险 | 项目负责人确认 |
| ADR-002 | 2026-08-20 | 第一迁移模块为订单邮件自动化20张表 | 模块内联表和事务紧密，必须整体迁移 | 项目负责人确认 |
| ADR-003 | 2026-08-20 | PP转码和PDF/图片转Excel不进入第一批 | 共享jobs、规则和配置边界需单独梳理 | 项目负责人确认 |
| ADR-004 | 2026-08-20 | 数据库迁移与附件文件迁移分开 | 避免同时改变两套高风险基础设施 | 项目负责人确认 |
| ADR-005 | 2026-08-20 | 文件最终优先使用对象存储，厂商待定 | 支持多实例、持久化、备份和扩展 | 项目负责人确认 |
| ADR-006 | 2026-08-20 | 第一阶段不改变日期、JSON和业务字段语义 | 确保只切换数据库层，不改变现有逻辑 | 项目负责人确认 |

新增或修改关键决策时，必须在此追加记录，不能覆盖旧决策。

---

## 14. 迁移执行日志

> 日志只能按时间追加，不得删除历史记录。代码变更、数据迁移、切换、回滚和重要验证均需记录。

| 时间 | 执行人 | 任务编号 | 操作 | Git提交/迁移版本 | 结果 | 证据或备注 |
|---|---|---|---|---|---|---|
| 2026-08-20 | Codex | M0 | 完成全量迁移评估并创建执行计划 | 文档v1.0 | 已完成 | 尚未修改代码和数据库 |
| 2026-08-20 | Codex | DB-01至DB-11、FS-01 | 完成第一阶段开发准备、SQLite审计和迁移验证工具 | `a842e12` / `0001_automation_schema.sql` | 待验证 | 定向40项通过、PG集成1项跳过；未切换读写、未迁移附件、未修改SQLite |
| 2026-08-20 | Codex | DB-12至DB-15 | 完成Outbox/Inbox、影子比较、只读基准和准入检查开发 | `2467e15` / `0002_shadow_sync.sql` | 待验证 | 全套56项通过、PG集成1项跳过；准入检查有16个阻塞项，DB-16保持未开始 |
| 2026-08-20 | Codex | DB-16至DB-18 | 完成切换准备、变更捕获、幂等回放和观察门禁开发 | `c8f9d97` / `0003_cutover_rollback.sql` | 待验证 | 全套63项通过、PG集成1项跳过；未切换、未回放生产数据、未开始正式观察 |
| 2026-08-20 | Codex | 全局状态校正 | 区分工具开发与实际数据迁移，纠正第二至第四阶段状态 | 文档v1.1 | 已完成 | 第一阶段实际迁移未开始；第二、第三、第四阶段均未迁移 |
| 2026-08-20 | Codex | DB-01、DB-04至DB-15 | 在本地PostgreSQL 16.15执行第一阶段隔离迁移、恢复、核对和测试库回滚演练 | `ac266be` / 文档v1.2 | 部分完成 | 66项测试通过；两次复制一致；影子10项零差异；正式切换、7天观察和人工验收未执行 |
| 2026-08-20 | Codex | DB-03、DB-16 | 经用户授权跳过DB-15，仅执行本地第一阶段切换 | `20d17d0` / 文档v1.3 | 本地成功 | 20表核对通过；SQLite与PostgreSQL均已备份；真实后端业务链路及全套67项通过；生产未切换 |
| 2026-08-20 | Codex | DB-16、DB-17自检 | 补启本地变更捕获并修复无邮箱用户打开订单自动化页面时的PostgreSQL参数类型错误 | `3e798e3` / 文档v1.4 | 本地通过 | 20表零差异；20个触发器；真实捕获insert/delete；三个第一阶段页面正常；67项无跳过通过；完整回滚演练和观察期仍未完成 |
| 2026-08-21 | Codex | M6第二阶段本地迁移 | 仅迁移并切换`users`，完成身份库隔离、全量核对、认证回归和回滚演练 | `4949699` / `0001_users.sql` / 文档v1.6 | 本地完成 | 11行零差异且密码哈希一致；真实PG集成1项、全套72项通过；生产未切换 |
| 2026-08-21 | Codex | M7第三阶段本地迁移 | 整体迁移并切换共享`jobs`及PP/通用转码18表 | `18ac4ca` / `0001_transcode_schema.sql` / 文档v1.7 | 本地完成 | 18表零差异；156条任务文件路径完整；18触发器；五页面200；全套75项通过；生产未切换 |
| 2026-08-21 | Codex | M8第四阶段本地迁移 | 整表迁移并切换`settings`和`pdf_excel_ai_config_versions` | `606763c` / `0001_configuration_schema.sql` / 文档v1.8 | 本地完成 | 63与0行零差异；敏感哈希和密文保持一致；回滚回放通过；四页面200；全套80项通过；生产未切换 |
| 2026-08-21 | Codex | M9第五阶段本地迁移 | 整表迁移并切换工作规划与反馈3表 | `d460d95` / `0001_planning_schema.sql` / 文档v1.9 | 本地完成 | 3/0/0行零差异；真实CRUD、备份恢复及7条变更回放通过；三页面200；全套85项通过；生产未切换 |

---

## 15. 每次开发提交模板

协作开发人员完成一批修改后，在对应任务和迁移日志中填写以下内容：

```text
任务编号：
状态：进行中 / 待验证 / 已完成 / 已阻塞 / 已回滚
负责人：
开始时间：
完成时间：
分支：
Git提交：
数据库迁移版本：
涉及文件：
修改内容：
业务逻辑是否变化：否（如为是，必须停止并重新确认范围）
测试命令：
测试结果：
数据核对结果：
性能结果：
回滚方式：
回滚是否演练：
遗留问题：
下一步：
```

---

## 16. 当前下一步

当前不应进入第二、第三或第四阶段。下一步严格按以下顺序完成第一阶段真实验证：

1. [x] 准备隔离的本地PostgreSQL测试实例，不使用生产连接；生产级SSL和最小权限账号仍待部署环境验证；
2. [x] 执行`0001_automation_schema.sql`至`0003_cutover_rollback.sql`并验证结构和默认关闭开关；权限验证仍待执行；
3. [~] 使用本地SQLite快照完成20张表全量复制、重复执行、全量核对和测试回滚；仍需含邮件、附件、案例和模板的代表性快照复验；
4. [x] 验证短时PostgreSQL断线期间SQLite业务正常、恢复后Outbox可追平且重复投递幂等；长时积压和告警仍待执行；
5. [~] 完成只读基准和备份恢复演练；批准阈值、写入并发和业务人工验收仍待执行；
6. 连续至少7天影子比较无未解释差异；
7. DB-15全部准入条件满足并由三名不同负责人确认后，才可另行申请DB-16正式切换；
8. 第一阶段正式切换并稳定观察完成后，才启动第9.1第二阶段用户与登录迁移。

本地第一阶段已按用户授权提前切换，用于开发验证和后续Git协作；该授权不适用于生产。生产仍必须完成第5至第7项后另行切换。

在以上验证完成前：禁止切换生产读写，禁止开始第二至第四阶段正式迁移，禁止删除SQLite数据或旧附件。
