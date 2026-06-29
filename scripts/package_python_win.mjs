import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const PYTHON_VERSION = "3.12.8";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const vendorDir = path.join(rootDir, "vendor", "python-win");
const tempDir = path.join(rootDir, "build", "python-win");
const zipUrl = `https://www.python.org/ftp/python/${PYTHON_VERSION}/python-${PYTHON_VERSION}-embed-amd64.zip`;
const getPipUrl = "https://bootstrap.pypa.io/get-pip.py";

function run(command, args, options = {}) {
  execFileSync(command, args, {
    cwd: rootDir,
    stdio: "inherit",
    ...options,
  });
}

function psQuote(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

if (process.platform !== "win32") {
  console.error("package:python-win 需要在 Windows x64 或 Windows CI 中运行。");
  console.error("原因：产物需要 Windows embeddable Python 及其 site-packages。");
  process.exit(1);
}

fs.rmSync(vendorDir, { recursive: true, force: true });
fs.rmSync(tempDir, { recursive: true, force: true });
fs.mkdirSync(vendorDir, { recursive: true });
fs.mkdirSync(tempDir, { recursive: true });

const zipPath = path.join(tempDir, "python-embed.zip");
const getPipPath = path.join(tempDir, "get-pip.py");

run("powershell.exe", [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-Command",
  [
    "$ErrorActionPreference = 'Stop'",
    `Invoke-WebRequest -Uri ${psQuote(zipUrl)} -OutFile ${psQuote(zipPath)}`,
    `Expand-Archive -Path ${psQuote(zipPath)} -DestinationPath ${psQuote(vendorDir)} -Force`,
    `Invoke-WebRequest -Uri ${psQuote(getPipUrl)} -OutFile ${psQuote(getPipPath)}`,
  ].join("; "),
]);

const pthFile = fs.readdirSync(vendorDir).find((file) => /^python\d+\._pth$/.test(file));
if (!pthFile) {
  throw new Error("未找到 Python embeddable ._pth 文件，无法启用 site-packages。");
}

const pthPath = path.join(vendorDir, pthFile);
const pthContent = fs.readFileSync(pthPath, "utf8").replace(/^#import site$/m, "import site");
fs.writeFileSync(pthPath, pthContent, "utf8");

const pythonExe = path.join(vendorDir, "python.exe");
run(pythonExe, [getPipPath]);
run(pythonExe, ["-m", "pip", "install", "--upgrade", "openpyxl"]);

fs.rmSync(tempDir, { recursive: true, force: true });

console.log(`Windows Python runtime is ready at ${path.relative(rootDir, vendorDir)}`);
