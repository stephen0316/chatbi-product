import { app, BrowserWindow, dialog, session, shell } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

let mainWindow = null;
let serverRuntime = null;
let quitting = false;

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function readJsonIfExists(targetPath) {
  if (!(await pathExists(targetPath))) return {};
  const raw = await fs.readFile(targetPath, "utf8");
  return JSON.parse(raw);
}

async function readEmbeddedConfig() {
  const appRoot = app.getAppPath();
  const candidates = app.isPackaged
    ? [path.join(process.resourcesPath, "embedded-config.json")]
    : [
        path.join(appRoot, "build", "electron", "embedded-config.json"),
        path.join(appRoot, "electron", "embedded-config.json"),
      ];

  for (const candidate of candidates) {
    const config = await readJsonIfExists(candidate);
    if (Object.keys(config).length) return config;
  }
  return {};
}

async function configureRuntimeEnvironment() {
  const appRoot = app.getAppPath();
  const storageDir = path.join(app.getPath("userData"), "storage");
  await fs.mkdir(storageDir, { recursive: true });

  process.env.CHATBI_STORAGE_DIR = storageDir;
  process.env.SESSION_RETENTION_DAYS = process.env.SESSION_RETENTION_DAYS || "3";

  const config = await readEmbeddedConfig();
  if (!process.env.GEMINI_API_KEY && config.geminiApiKey) {
    process.env.GEMINI_API_KEY = config.geminiApiKey;
  }
  if (!process.env.GEMINI_MODEL && config.geminiModel) {
    process.env.GEMINI_MODEL = config.geminiModel;
  }

  const analyzerScript = app.isPackaged
    ? path.join(process.resourcesPath, "scripts", "analyze_delisting.py")
    : path.join(appRoot, "scripts", "analyze_delisting.py");
  process.env.CHATBI_ANALYZER_SCRIPT = analyzerScript;

  const packagedMacAnalyzer = path.join(process.resourcesPath, "analyzer-mac", "analyze_delisting");
  if (app.isPackaged && process.platform === "darwin") {
    if (await pathExists(packagedMacAnalyzer)) {
      process.env.CHATBI_ANALYZER_BIN = packagedMacAnalyzer;
      return;
    }
    throw new Error("缺少内置 macOS 分析程序：resources/analyzer-mac/analyze_delisting");
  }

  const packagedPython = path.join(process.resourcesPath, "python-win", "python.exe");
  if (await pathExists(packagedPython)) {
    process.env.PYTHON_BIN = packagedPython;
  } else if (app.isPackaged && process.platform === "win32") {
    throw new Error("缺少内置 Python 运行时：resources/python-win/python.exe");
  }
}

function isLocalAppUrl(url, port) {
  try {
    const parsed = new URL(url);
    return (
      parsed.protocol === "http:" &&
      ["127.0.0.1", "localhost"].includes(parsed.hostname) &&
      parsed.port === String(port)
    );
  } catch {
    return false;
  }
}

function restrictRendererNetwork(port) {
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    try {
      const parsed = new URL(details.url);
      const allowed =
        parsed.protocol === "data:" ||
        parsed.protocol === "blob:" ||
        isLocalAppUrl(details.url, port) ||
        (parsed.protocol === "https:" && parsed.hostname === "generativelanguage.googleapis.com");
      callback({ cancel: !allowed });
    } catch {
      callback({ cancel: true });
    }
  });
}

function createWindow(port) {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1100,
    minHeight: 720,
    title: "产品慧诊",
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  const localUrl = `http://127.0.0.1:${port}/`;
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (isLocalAppUrl(url, port)) return { action: "allow" };
    shell.openExternal(url).catch(() => {});
    return { action: "deny" };
  });
  mainWindow.webContents.on("will-navigate", (event, url) => {
    if (!isLocalAppUrl(url, port)) {
      event.preventDefault();
      shell.openExternal(url).catch(() => {});
    }
  });

  mainWindow.loadURL(localUrl);
  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

async function startLocalServer() {
  const appRoot = app.getAppPath();
  const serverEntry = path.join(appRoot, "server.js");
  const { startServer } = await import(pathToFileURL(serverEntry).href);
  serverRuntime = await startServer({ port: 0, host: "127.0.0.1" });
  return serverRuntime.port;
}

async function stopLocalServer() {
  if (!serverRuntime?.server?.listening) return;
  await new Promise((resolve) => {
    serverRuntime.server.close(() => resolve());
  });
}

async function bootstrap() {
  await configureRuntimeEnvironment();
  const port = await startLocalServer();
  restrictRendererNetwork(port);
  createWindow(port);
}

app.whenReady().then(bootstrap).catch((error) => {
  dialog.showErrorBox("产品慧诊启动失败", error.message || String(error));
  app.quit();
});

app.on("window-all-closed", () => {
  app.quit();
});

app.on("before-quit", async (event) => {
  if (quitting) return;
  event.preventDefault();
  quitting = true;
  await stopLocalServer().catch(() => {});
  app.quit();
});
