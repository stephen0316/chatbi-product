import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const venvDir = path.join(rootDir, "build", "analyzer-mac-venv");
const workDir = path.join(rootDir, "build", "pyinstaller-mac");
const distDir = path.join(rootDir, "vendor", "analyzer-mac");
const scriptPath = path.join(rootDir, "scripts", "analyze_delisting.py");
const pythonBin = path.join(venvDir, "bin", "python");

function run(command, args, options = {}) {
  execFileSync(command, args, {
    cwd: rootDir,
    stdio: "inherit",
    ...options,
  });
}

if (process.platform !== "darwin") {
  console.error("package:analyzer-mac 需要在 macOS 环境中运行。");
  console.error("原因：PyInstaller 只能为当前操作系统生成本机可执行文件。");
  process.exit(1);
}

fs.rmSync(venvDir, { recursive: true, force: true });
fs.rmSync(workDir, { recursive: true, force: true });
fs.rmSync(path.join(distDir, "analyze_delisting"), { force: true });
fs.mkdirSync(distDir, { recursive: true });

run(process.env.PYTHON_BIN || "python3", ["-m", "venv", venvDir]);
run(pythonBin, ["-m", "pip", "install", "--upgrade", "pip"]);
run(pythonBin, ["-m", "pip", "install", "--upgrade", "pyinstaller", "openpyxl"]);
run(pythonBin, [
  "-m",
  "PyInstaller",
  "--onefile",
  "--name",
  "analyze_delisting",
  "--distpath",
  distDir,
  "--workpath",
  workDir,
  "--specpath",
  workDir,
  scriptPath,
]);

fs.chmodSync(path.join(distDir, "analyze_delisting"), 0o755);

console.log(`macOS analyzer is ready at ${path.relative(rootDir, path.join(distDir, "analyze_delisting"))}`);
