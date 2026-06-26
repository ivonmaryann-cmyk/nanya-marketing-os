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
