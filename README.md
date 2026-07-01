# 产品清退官

产品清退官用于上传「产品全量列表」和「收入及直接成本明细表」，自动识别清退候选、生成退市建议表，并支持基于本次上传数据进行问答。

## 文档

- [桌面部署方案](docs/desktop-deployment-plan.md)
- [Windows 桌面版打包操作手册](docs/windows-electron-build-manual.md)
- [macOS 桌面版打包操作手册](docs/macos-electron-build-manual.md)
- [Docker 内网部署交付手册](docs/docker-deployment-guide.md)

## 启动

```bash
npm install
cp .env.example .env
npm run dev
```

浏览器打开：

```text
http://localhost:3000
```

## Docker 内网部署

服务器 Web 版推荐使用 Docker，避免内网服务器 Node/glibc/npm 版本不兼容。容器内会固定 Node、Python 和 openpyxl，运行时不会安装 Electron 打包依赖。

```bash
cp .env.docker.example .env
docker build -t chatbi:latest .
docker compose up -d
```

浏览器访问：

```text
http://服务器IP:3001
```

离线内网部署可先在可联网机器构建并导出镜像：

```bash
docker save chatbi:latest -o chatbi-latest.tar
```

详细步骤见 [Docker 内网部署交付手册](docs/docker-deployment-guide.md)。

## 模型问答配置

模型 key 只放在服务端 `.env`，不要写入前端代码：

```bash
GEMINI_API_KEY=your_key_here
```

未配置 `GEMINI_API_KEY` 时，系统仍支持常见统计类本地问答，例如强制退市数量、规则命中数量等。

## Windows 桌面版打包

桌面版目标是交付 Windows x64 绿色包，客户解压后双击 `产品清退官.exe` 使用，不需要安装 Node、Python 或额外依赖。

推荐在 Windows x64 或 Windows CI 中执行：

```bash
npm install
cp .env.example .env
npm run package:python-win
npm run build:win
```

构建前需要在 `.env` 中写入 `GEMINI_API_KEY`。构建脚本会生成 `build/electron/embedded-config.json`，并将 key 作为 Electron 资源打入安装包，不会暴露到前端代码。注意：内置 key 无法做到绝对保密，安装包被逆向后仍存在泄露风险。

Windows 产物：

```text
dist/产品清退官-win-x64.zip
```

桌面版运行数据写入 Electron `userData` 目录下的 `storage`，即 Windows 上通常位于：

```text
%APPDATA%/产品清退官/storage
```

上传文件、预检结果、分析结果和导出文件按 session 隔离保存，并沿用 3 天自动清理机制。

## macOS 桌面版打包

macOS 桌面版会把 Python 分析脚本封装成原生可执行文件，客户不需要安装 Python。

推荐在目标架构一致的 macOS 机器上执行。例如 Apple Silicon 客户用 arm64 Mac 打包，Intel 客户用 x64 Mac 打包。

```bash
npm install
cp .env.example .env
npm run package:analyzer-mac
npm run build:mac
```

构建前同样需要在 `.env` 中写入 `GEMINI_API_KEY`。

macOS 产物：

```text
dist/产品清退官-mac-arm64.zip
dist/产品清退官-mac-x64.zip
```

实际文件名取决于打包机器架构。当前配置生成 zip 包，不做签名和 notarization；首次打开时 macOS 可能提示来自未知开发者，需要在系统设置中允许打开。若要正式对外分发，建议后续补充 Apple Developer ID 签名和 notarization。

## 本地缓存

上传文件和分析结果会缓存在服务器本地：

```text
storage/uploads
storage/cache
```

重新上传并分析会覆盖上一批缓存。

## 当前分析口径

- 规则 1/2/3 的状态范围：`已上市`、`已入库`
- 规则 4：全部 `退市中` 产品纳入，暂不判断超过 1 年
- 2 年经营数据窗口：`2024-06` 至 `2026-05`
- 超过 2 年基准：早于 `2024-06-23`，或创建时间为空
- 顶部统计展示：符合退市条件的产品、强制退市、建议退市
- 结果表“理由”列仅展示具体规则内容，不显示 `规则 X：` 前缀
