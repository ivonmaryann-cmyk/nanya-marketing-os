# 宝塔 Docker 部署说明

本文适用于当前生产服务器：

- 项目目录：`/www/wwwroot/nanya-marketing-os`
- 宿主机运行用户：`www`，UID/GID 均为 `1000`
- 容器监听：`127.0.0.1:5000`
- 数据目录：`/www/wwwroot/nanya-marketing-os/storage`
- 环境变量：`/www/wwwroot/nanya-marketing-os/.env`

## 部署原则

1. 第一次切换成功前，不删除旧宝塔 Python 项目和旧项目目录。
2. `storage/` 和 `.env` 只保存在服务器，不进入 Git，也不打进 Docker 镜像。
3. SQLite、历史任务、上传文件、结果文件和规则版本都通过 `storage/` 持久化。
4. 当前应用只运行一个 Gunicorn 工作进程，不要增加容器副本数量。

## 一、更新代码

```bash
cd /www/wwwroot/nanya-marketing-os
git checkout dev
git pull
```

确认运行数据存在：

```bash
ls -lah .env storage storage/app.db
```

如果 `.env` 或 `storage/app.db` 不存在，先上传完整运行数据，不要启动容器。

## 二、备份运行数据

备份必须放在项目目录外：

```bash
cd /www/wwwroot
backup_dir="nanya_backup_$(date +%Y%m%d_%H%M%S)"
mkdir -p "$backup_dir"
cp -a nanya-marketing-os/.env "$backup_dir/"
cp -a nanya-marketing-os/storage "$backup_dir/"
echo "备份目录：/www/wwwroot/$backup_dir"
```

## 三、准备权限

```bash
cd /www/wwwroot/nanya-marketing-os
chown -R 1000:1000 storage
chmod -R u+rwX,g+rwX storage
chown 1000:1000 .env
chmod 600 .env
```

## 四、构建镜像

旧 Python 项目此时可以继续运行：

```bash
cd /www/wwwroot/nanya-marketing-os
docker compose build --build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple/
```

首次构建需要下载 Python、Docling 和 OCR 依赖，时间较长属于正常情况。构建完成后检查：

```bash
docker compose images
```

## 五、切换到 Docker

先在宝塔的 Python 项目页面停止旧项目，避免占用端口，然后执行：

```bash
cd /www/wwwroot/nanya-marketing-os
docker compose up -d
docker compose ps
docker compose logs --tail=100 web
```

本机检查：

```bash
curl -I http://127.0.0.1:5000/login
```

确认登录、历史任务、规则和 PDF 转 Excel 功能正常后，再继续配置域名。

## 六、宝塔反向代理

在宝塔网站中添加或使用现有域名，将反向代理目标设置为：

```text
http://127.0.0.1:5000
```

浏览器通过域名访问 `/login`。端口 `5000` 只绑定到服务器本机，不需要在防火墙中开放。

## 七、更新版本

```bash
cd /www/wwwroot/nanya-marketing-os
git checkout dev
git pull
docker compose build --build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple/
docker compose up -d
docker compose ps
docker compose logs --tail=100 web
```

`docker compose up -d` 不会删除挂载在宿主机的 `storage/`。

## 八、自动清理历史任务

宝塔计划任务每天执行：

```bash
cd /www/wwwroot/nanya-marketing-os || exit 1
docker compose exec -T web python -m fangzheng_web_app.cleanup --days 30
```

第一次先预演：

```bash
cd /www/wwwroot/nanya-marketing-os || exit 1
docker compose exec -T web python -m fangzheng_web_app.cleanup --days 30 --dry-run
```

## 九、回滚

如果 Docker 启动后功能异常：

```bash
cd /www/wwwroot/nanya-marketing-os
docker compose down
```

然后在宝塔 Python 项目页面重新启动旧项目。不要执行 `docker compose down -v`，不要删除 `storage/`、`.env` 或备份目录。

## 常用排查命令

```bash
cd /www/wwwroot/nanya-marketing-os
docker compose ps
docker compose logs --tail=200 web
docker compose restart web
docker stats --no-stream nanya-marketing-os
df -h /
```
