# 南亚营销自动化平台 v1.9.0

这是内部网页版营销自动化工具。当前 `v1.9.0` 集成了 V1.6 视觉界面，并保留当前阶段已完成的业务功能：方正价格计算、营销自动化转码、深南价格计算、沪士价格计算、博敏价格计算、在途核对、历史记录、规则管理和反馈意见。

- 工号登录
- Excel 批量计算 / 转码
- 五个已接入功能支持单条即时计算 / 转码
- 在途核对支持上传含 `厂内明细` 和 `客户明细` 的 Excel，输出核对汇总、系统数据核对、待出货明细和客户数据核对
- 下载结果 Excel
- 按工号查看历史记录
- 管理员上传新的规则文件并立即生效
- 修改管理员密码
- 独立计算子进程与运行中任务停止

## 目录说明

- `app.py`: 启动入口
- `fangzheng_web_app/`: Web 应用代码、计算/转码引擎、默认 pkl 规则数据和默认规则种子
- `templates/`: 页面模板
- `static/`: 样式、品牌资源和前端脚本
- `storage/`: SQLite、规则版本、上传文件和结果文件。服务器升级时不要覆盖该目录

## 集成说明

方正、转码、深南、沪士、博敏相关逻辑已集成到网页项目内，部署服务器时只需要部署本项目目录，不再依赖外部桌面工具路径。

方正核心文件包括：

- `fangzheng_web_app/price_calculator_v3.py`
- `fangzheng_web_app/data_price.pkl`
- `fangzheng_web_app/data_account.pkl`

其他功能核心文件包括：

- `fangzheng_web_app/transcode_engine.py`
- `fangzheng_web_app/transcode_service.py`
- `fangzheng_web_app/shennan_service.py`
- `fangzheng_web_app/hushi_service.py`
- `fangzheng_web_app/bomin_service.py`
- `fangzheng_web_app/bomin_rules.py`
- `fangzheng_web_app/default_rules/hushi_rules.zip`
- `fangzheng_web_app/default_rules/bomin_price_rules.xlsx`

## 环境版本

请按 `requirements.txt` 中固定版本安装依赖，避免服务器和开发机版本不一致：

```powershell
python -m pip install -r requirements.txt
```

当前锁定版本：

- Flask 3.1.3
- Werkzeug 3.1.8
- pandas 3.0.2
- numpy 2.4.4
- openpyxl 3.1.5
- xlrd 2.0.1

## 首次启动

```powershell
python app.py
```

浏览器打开：`http://127.0.0.1:5000`

### NYEOS 内网 HTTPS 证书

当接口维护中的 NYEOS 地址使用 `https://nyeos.nouyatec.com` 或
`https://10.30.12.117`，且服务端使用企业自签名证书时，在不提交版本库的
`config/local.env` 配置证书路径：

```bash
NYEOS_CA_CERT_FILE=/absolute/path/to/api.crt
```

平台会保留默认 HTTPS 校验，并仅为上述 NYEOS 地址额外信任该证书；不会全局关闭
证书或主机名校验。修改后需重启服务。

## 南亚智能办公连接器（开发版）

本项目提供独立的 JSON REST + MCP Streamable HTTP 门面，直接调用邮件抓取与订单案件服务层；它不模拟网页登录、不返回 HTML，也不使用 Cookie。

首次启动前，为连接器设置**独立** Token（不是邮箱授权码、SMTP 密码或 NYEOS 凭据）：

```bash
export CONNECTOR_API_TOKEN='请生成一段随机长字符串'
export CONNECTOR_ALLOWED_EMPLOYEE_IDS=23582
export CONNECTOR_DEFAULT_EMPLOYEE_ID=23582
```

调用需带 `Authorization: Bearer $CONNECTOR_API_TOKEN`，可带 `X-Nanya-Employee-Id: 23582` 与 `X-Correlation-Id`。默认范围是工号 `23582`。

| 端点 | 用途 |
| --- | --- |
| `GET /api/connector/v1/health` | 如实返回邮箱、料号查询、录单、SMTP 的 enabled、Mock/Real mode 与 readiness；`real_configured_unverified` 只代表配置存在，绝不代表真实接口已验通。 |
| `GET /api/connector/v1/jobs/{id}` | 查询订单邮件同步任务。 |
| `GET /api/connector/v1/orders/cases` | 列出订单案件。 |
| `GET /api/connector/v1/orders/cases/{id}` | 读取案件与已提取的业务信息。 |
| `POST /api/connector/v1/mcp` | MCP Streamable HTTP JSON-RPC 入口。 |

MCP 工具及其业务边界：

- `marketing.mail.sync_orders`、`marketing.job.get`：真实调用已有 IMAP 后台抓取队列并轮询任务；
- `marketing.order.list_cases`、`marketing.order.get_case`：读取案件及已经提取的业务数据；
- `marketing.material.query`：调用既有 `build_material_query`，严格沿用接口维护中的 Mock/Real 状态、模板校验和接口审计；
- `marketing.order.prepare_entry`：调用既有录单模板校验/载荷映射，仅生成载荷预览，不请求 NYEOS、不写接口日志；
- `marketing.order.reply_draft`：生成可编辑邮件草稿，绝不发送；
- `marketing.order.submit_entry`：只有 `confirm=true` 才调用既有 `build_domestic_order_entry`，沿用其 Mock/Real、成功提交防重复和审计记录。

现有 SMTP 发信服务尚未提供幂等键，因此连接器**不暴露** `marketing.mail.send_reply`，避免 Agent 重试造成重复邮件。接口不会返回授权码、密码、Cookie、HTML 或服务器文件路径。

开发验收可进行一次真实 MCP Streamable HTTP 握手（只执行 `initialize` 与 `tools/list`，不会同步邮件；未设置 `CONNECTOR_BASE_URL` 时脚本会拉起一个不连接数据库的临时 Flask 测试宿主）：

```bash
CONNECTOR_API_TOKEN='同上' node tests/connector_api_mcp_handshake.mjs
```

## 历史任务清理

先预演超过 30 天的终态任务和任务日志：

```bash
python -m fangzheng_web_app.cleanup --days 30 --dry-run
```

确认预演结果后执行清理：

```bash
python -m fangzheng_web_app.cleanup --days 30
```

清理命令只处理 `completed`、`failed` 和 `canceled` 任务。排队中和运行中的任务、规则文件、PDF/Excel 模板、用户与系统配置不会被删除。

宝塔计划任务建议每天凌晨执行一次：

```bash
cd /www/wwwroot/nanya-marketing-os || exit 1
/apps/env/mkt_cal/bin/python3.11 -m fangzheng_web_app.cleanup --days 30
```

首次部署时先把命令末尾改成 `--days 30 --dry-run`，确认计划任务日志中的统计结果后再移除 `--dry-run`。

## 默认规则

首次启动时，系统会从项目内置文件初始化首个规则版本：

- `fangzheng_web_app/data_price.pkl`
- `fangzheng_web_app/data_account.pkl`
- `fangzheng_web_app/default_rules/hushi_rules.zip`
- `fangzheng_web_app/default_rules/bomin_price_rules.xlsx`

方正、转码、深南、沪士、博敏都会初始化默认规则版本。后续管理员在规则管理页面上传的规则文件会覆盖当前生效版本。

## 登录规则

- 首次无账号时，输入任意工号并使用同样工号作为密码，可创建首个管理员账号
- 后续用户需由管理员在账号与密码页面维护；初始密码为工号
- 管理员权限：在规则管理或管理员密码页输入管理员密码

默认管理员密码：

```text
admin123
```

## 规则文件要求

方正价格计算支持两份可版本化管理的规则文件：

1. `价格对账表`
2. `基板对照表`

转码、深南、沪士、博敏规则在各自功能页的规则管理入口中维护；沪士规则采用 ZIP 包整包上传方式维护多份报价 Excel。
