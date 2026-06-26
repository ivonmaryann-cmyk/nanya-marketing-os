# 协作开发说明

## 本地启动

```powershell
cd D:\Carson\tmp\server_release\nanya_platform_v190_release_20260615_112805
python app.py
```

浏览器打开：

```text
http://127.0.0.1:5000/login
```

## Git 协作规则

- `main`：稳定基线，可用于打包和部署。
- `dev`：日常集成分支。
- `feature/<name>`：新增功能分支。
- `fix/<name>`：问题修复分支。

每次开发前先从最新分支拉取，单个功能或修复完成后提交，再合并回 `dev` 验证。

## 不提交的内容

`storage/` 不进入 Git。里面包含本地/生产数据库、用户、管理员密码、规则版本、历史任务、上传文件和结果文件。

不要提交日志、缓存、临时包、备份文件和本地运行产物。

## 提交前检查

```powershell
python -m compileall -q .
```

关键功能变更后，至少打开本地网页确认：

- 登录页
- 功能中心
- 被修改的功能页
- 历史记录页

## 服务器部署提醒

服务器部署时只替换代码文件，继续保留服务器现有：

```text
/www/wwwroot/fangzheng_web_app/storage
```

除非明确要重置测试环境，否则不要覆盖生产 `storage/`。
