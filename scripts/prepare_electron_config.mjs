import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "build", "electron");
const outputPath = path.join(outputDir, "embedded-config.json");
const pythonExe = path.join(rootDir, "vendor", "python-win", "python.exe");
const macAnalyzer = path.join(rootDir, "vendor", "analyzer-mac", "analyze_delisting");
const platformArg = process.argv
  .slice(2)
  .find((arg) => arg.startsWith("--platform="))
  ?.split("=")
  .at(1);
const targetPlatform = platformArg || process.env.CHATBI_BUILD_PLATFORM || "win";

async function exists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

const geminiApiKey = process.env.GEMINI_API_KEY || "";
const geminiModel = process.env.GEMINI_MODEL || "gemini-2.5-flash";
const allowEmptyKey = process.env.ALLOW_EMPTY_GEMINI_KEY === "1";

if (!geminiApiKey && !allowEmptyKey) {
  throw new Error(
    "缺少 GEMINI_API_KEY。请在 .env 或环境变量中设置后再运行打包命令；如需测试空 key，可设置 ALLOW_EMPTY_GEMINI_KEY=1。",
  );
}

if (targetPlatform === "win" && !(await exists(pythonExe))) {
  throw new Error(
    "缺少 Windows 内置 Python：vendor/python-win/python.exe。请先在 Windows 环境运行 npm run package:python-win。",
  );
}

if (targetPlatform === "mac" && !(await exists(macAnalyzer))) {
  throw new Error(
    "缺少 macOS 分析程序：vendor/analyzer-mac/analyze_delisting。请先在 macOS 环境运行 npm run package:analyzer-mac。",
  );
}

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(
  outputPath,
  JSON.stringify({ geminiApiKey, geminiModel }, null, 2),
  "utf8",
);

console.log(`Electron embedded config written to ${path.relative(rootDir, outputPath)}`);
