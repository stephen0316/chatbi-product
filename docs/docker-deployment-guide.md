# ChatBI Docker 部署交付手册

本文用于交付给公司内网研发部署 Web 服务版 ChatBI。Docker 版不包含 Electron 桌面壳，适合部署在公司内网服务器，通过浏览器访问。

## 1. 部署前提

服务器只需要满足：

- Linux x64
- 已安装 Docker
- 如使用 `docker compose`，需安装 Docker Compose v2

服务器本机不需要安装 Node、npm、Python 或 openpyxl。Node、Python 和 Python 依赖都会封装在镜像里。

## 2. 研发需要接收的文件

把整个项目目录交付给研发，至少要包含：

- `Dockerfile`
- `.dockerignore`
- `docker-compose.yml`
- `.env.docker.example`
- `package.json`
- `package-lock.json`
- `server.js`
- `public/`
- `scripts/analyze_delisting.py`

不需要交付：

- `node_modules/`
- `dist/`
- `storage/`
- `outputs/`
- `vendor/`
- `electron/`

## 3. 在线构建部署

适用于服务器可以访问 npm 和 PyPI，或公司内网已配置代理/镜像源。

### 3.1 准备环境变量

```bash
cp .env.docker.example .env
vi .env
```

至少确认：

```bash
GEMINI_API_KEY=你的服务端Key
GEMINI_MODEL=gemini-2.5-flash
```

如果内网不能访问 Gemini API，可以先保留空值。本地分析和导出仍可用，智能问答会受影响。

### 3.2 构建镜像

```bash
docker build -t chatbi:latest .
```

如果需要使用公司 npm / pip 镜像源：

```bash
docker build \
  --build-arg NPM_REGISTRY=https://registry.npmmirror.com \
  --build-arg PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple \
  -t chatbi:latest .
```

### 3.3 启动服务

```bash
docker compose up -d
```

访问地址：

```text
http://服务器IP:3001
```

### 3.4 查看状态

```bash
docker compose ps
docker compose logs -f chatbi
curl http://127.0.0.1:3001/api/status
```

正常情况下 `/api/status` 会返回 JSON。

## 4. 离线交付部署

适用于内网服务器不能访问外网。

### 4.1 在可联网机器上构建镜像

```bash
docker build -t chatbi:latest .
docker save chatbi:latest -o chatbi-latest.tar
```

把以下文件交付给内网研发：

- `chatbi-latest.tar`
- `docker-compose.yml`
- `.env.docker.example`

### 4.2 在内网服务器加载镜像

```bash
docker load -i chatbi-latest.tar
cp .env.docker.example .env
vi .env
docker compose up -d
```

访问：

```text
http://服务器IP:3001
```

## 5. 数据目录和清理

Docker Compose 默认使用命名卷：

```text
chatbi-storage
```

容器内路径：

```text
/app/storage
```

上传文件、预检结果、分析结果、导出文件都会写入该目录。系统默认保留 3 天：

```bash
SESSION_RETENTION_DAYS=3
TMP_RETENTION_MINUTES=30
```

如需改成服务器固定目录，例如 `/data/chatbi/storage`，把 `docker-compose.yml` 中的 volume 改为：

```yaml
volumes:
  - /data/chatbi/storage:/app/storage
```

并在服务器上执行：

```bash
mkdir -p /data/chatbi/storage
chown -R 1000:1000 /data/chatbi/storage
```

## 6. 常用运维命令

启动：

```bash
docker compose up -d
```

停止：

```bash
docker compose down
```

查看日志：

```bash
docker compose logs -f chatbi
```

重启：

```bash
docker compose restart chatbi
```

更新镜像后重启：

```bash
docker compose up -d --force-recreate
```

## 7. 端口调整

当前配置为：

```yaml
ports:
  - "3001:3000"
```

含义：

- 服务器访问端口：`3001`
- 容器内部端口：`3000`

如果公司要求使用 8080，可改为：

```yaml
ports:
  - "8080:3000"
```

然后访问：

```text
http://服务器IP:8080
```

## 8. 注意事项

- Docker 版是 Web 服务部署，不是桌面应用。
- 不要在服务器上执行普通 `npm install` 安装 Electron 打包依赖。
- 如果需要重新构建镜像，只需要执行 `docker build`，不需要服务器本机安装 Node/Python。
- 如果 Gemini API 在内网不可访问，本地分析和 Excel 导出仍可用，但智能问答会报模型接口错误。
