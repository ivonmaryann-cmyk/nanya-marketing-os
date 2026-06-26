# 南亚营销自动化平台 v1.9.0 部署说明

## 部署目标

本版本用于把服务器站点升级到 v1.9.0，新增在途核对功能，并保留 v1.8.0 已有的方正、转码、深南、沪士、博敏功能。

本次部署必须保留服务器现有 `storage/`，不要覆盖用户、密码、管理员密码、历史任务和现有规则版本。

## 发布包内容

发布包包含：

- `app.py`
- `fangzheng_web_app/`
- `templates/`
- `static/`
- `docs/`
- `requirements.txt`
- `README.md`
- `DEPLOY_v1.9.0.md`

发布包不包含：

- `storage/`
- `__pycache__/`
- 本地日志文件
- `.bak` 备份文件
- 临时文件

## 宝塔部署步骤

1. 在宝塔中停止 Python 项目 `fangzheng_web_app`。
2. 备份当前服务器目录，例如：`/www/wwwroot/fangzheng_web_app_bak_YYYYMMDD_HHMMSS`。
3. 上传 `nanya_platform_v190_release_YYYYMMDD_HHMMSS.zip` 到服务器临时目录。
4. 解压到临时目录，不要直接覆盖生产目录。
5. 确认并保留：`/www/wwwroot/fangzheng_web_app/storage`。
6. 用新包中的代码文件覆盖生产目录中的同名文件和目录：`app.py`、`fangzheng_web_app/`、`templates/`、`static/`、`docs/`、`requirements.txt`、`README.md`、`DEPLOY_v1.9.0.md`。
7. 安装或确认依赖：

```bash
/apps/env/mkt_cal/bin/python3.11 -m pip install -r /www/wwwroot/fangzheng_web_app/requirements.txt
```

8. 编译检查：

```bash
cd /www/wwwroot/fangzheng_web_app
/apps/env/mkt_cal/bin/python3.11 -m py_compile app.py fangzheng_web_app/*.py
```

9. 重启 Python/gunicorn 服务。

10. 服务器本机检查：

```bash
curl -I http://127.0.0.1:5000/login
```

11. 浏览器访问：

```text
http://ny-mkt-cal.nouyatec.com/
```

## 验证清单

- 原账号和密码可以登录。
- 功能中心显示在途核对入口。
- `/features/in-transit` 可以打开。
- 上传在途核对样例后任务能完成，并生成 `核对汇总`、`系统数据核对`、`待出货明细`、`客户数据核对` 四张 Sheet。
- 方正、转码、深南、沪士、博敏页面可以打开。
- 历史记录和管理员规则页面可以打开。

## 回滚

如果升级后出现核心功能异常：

1. 停止 Python 项目。
2. 将当前异常目录改名为 `fangzheng_web_app_failed_YYYYMMDD_HHMMSS`。
3. 将部署前备份目录恢复为 `fangzheng_web_app`。
4. 确认恢复后的目录仍包含原生产 `storage/`。
5. 重启 Python 项目并验证登录页。
