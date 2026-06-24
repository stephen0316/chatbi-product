import "dotenv/config";
import express from "express";
import multer from "multer";
import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3000);
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "";
const GEMINI_MODEL = process.env.GEMINI_MODEL || "gemini-2.5-flash";
const DEFAULT_BUNDLED_PYTHON =
  "/Users/apple/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";
const PYTHON_BIN = process.env.PYTHON_BIN || DEFAULT_BUNDLED_PYTHON;

const STORAGE_DIR = path.join(__dirname, "storage");
const UPLOAD_DIR = path.join(STORAGE_DIR, "uploads");
const CACHE_DIR = path.join(STORAGE_DIR, "cache");
const RESULT_JSON = path.join(CACHE_DIR, "analysis-result.json");
const RESULT_XLSX = path.join(CACHE_DIR, "全量产品退市筛选结果.xlsx");
const ANALYZER_SCRIPT = path.join(__dirname, "scripts", "analyze_delisting.py");

await fs.mkdir(UPLOAD_DIR, { recursive: true });
await fs.mkdir(CACHE_DIR, { recursive: true });

const app = express();
app.use(express.json({ limit: "1mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({
  dest: path.join(STORAGE_DIR, "tmp"),
  limits: {
    fileSize: 100 * 1024 * 1024,
    files: 12,
  },
});

function normalize(value) {
  if (value == null || typeof value === "boolean") return "";
  return String(value).trim();
}

function decodeUploadName(name) {
  const text = normalize(name);
  if (!/[ÃÂäåæçèé]/.test(text)) return text;
  try {
    const decoded = Buffer.from(text, "latin1").toString("utf8");
    return /[\u4e00-\u9fff]/.test(decoded) ? decoded : text;
  } catch {
    return text;
  }
}

function safeFileName(name) {
  return decodeUploadName(name).replace(/[\\/]/g, "_");
}

async function removeDirContents(dir) {
  await fs.rm(dir, { recursive: true, force: true });
  await fs.mkdir(dir, { recursive: true });
}

async function resetCache() {
  await removeDirContents(UPLOAD_DIR);
  await removeDirContents(CACHE_DIR);
}

async function readCachedPayload() {
  const text = await fs.readFile(RESULT_JSON, "utf8");
  return JSON.parse(text);
}

async function writeCachedPayload(payload) {
  await fs.writeFile(RESULT_JSON, JSON.stringify(payload, null, 2), "utf8");
}

async function runAnalyzer(productPath, revenuePaths) {
  const args = [
    ANALYZER_SCRIPT,
    "--product-list",
    productPath,
    "--revenue-files",
    ...revenuePaths,
    "--output-json",
    RESULT_JSON,
    "--output-xlsx",
    RESULT_XLSX,
  ];
  try {
    const { stdout } = await execFileAsync(PYTHON_BIN, args, {
      cwd: __dirname,
      maxBuffer: 1024 * 1024 * 20,
    });
    const lastLine = stdout.trim().split("\n").filter(Boolean).at(-1);
    return lastLine ? JSON.parse(lastLine) : { ok: true };
  } catch (error) {
    const detail = error.stderr || error.stdout || error.message;
    throw new Error(`分析脚本执行失败：${detail}`);
  }
}

async function callGemini(prompt) {
  const requestBody = JSON.stringify({
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2 },
  });
  const requestPath = path.join(CACHE_DIR, `gemini-request-${Date.now()}.json`);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`;
  await fs.writeFile(requestPath, requestBody, "utf8");
  try {
    const { stdout } = await execFileAsync(
      "curl",
      ["-s", "-X", "POST", url, "-H", "Content-Type: application/json", "--data-binary", `@${requestPath}`],
      { maxBuffer: 1024 * 1024 * 10 },
    );
    const data = JSON.parse(stdout);
    if (data?.error) throw new Error(data.error.message || "模型接口调用失败");
    return (
      data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("").trim() ||
      "模型没有返回有效回答。"
    );
  } catch (error) {
    throw new Error(error.message || "模型接口调用失败");
  } finally {
    await fs.rm(requestPath, { force: true }).catch(() => {});
  }
}

function createAiContext(payload) {
  return JSON.stringify(
    {
      口径: {
        基准日期: payload.metadata.as_of_date,
        两年收入窗口: payload.metadata.revenue_window,
        规则123状态范围: payload.metadata.active_statuses_for_rules_1_to_3,
        规则4状态范围: "退市中，暂不判断超过1年",
      },
      汇总: {
        候选总数: payload.metadata.candidate_count,
        退市类型分布: payload.metadata.delisting_type_counts,
        命中规则分布: payload.metadata.rule_counts,
        候选状态分布: payload.metadata.candidate_status_counts,
      },
      样例明细: payload.rows.slice(0, 80),
    },
    null,
    2,
  );
}

function answerLocally(question, payload) {
  const q = question.trim();
  if (/强制/.test(q)) {
    return `强制退市 ${payload.metadata.delisting_type_counts["强制退市"] || 0} 条。`;
  }
  if (/建议/.test(q)) {
    return `建议退市 ${payload.metadata.delisting_type_counts["建议退市"] || 0} 条。`;
  }
  if (/规则\s*1|规则一/.test(q)) return `规则1命中 ${payload.metadata.rule_counts["1"] || 0} 条。`;
  if (/规则\s*2|规则二/.test(q)) return `规则2命中 ${payload.metadata.rule_counts["2"] || 0} 条。`;
  if (/规则\s*3|规则三/.test(q)) return `规则3命中 ${payload.metadata.rule_counts["3"] || 0} 条。`;
  if (/规则\s*4|规则四|退市中/.test(q)) return `规则4命中 ${payload.metadata.rule_counts["4"] || 0} 条。`;
  if (/已入库/.test(q)) {
    return `候选结果中“已入库”状态 ${payload.metadata.candidate_status_counts["已入库"] || 0} 条。`;
  }
  return "";
}

app.get("/api/status", async (_req, res) => {
  try {
    const payload = await readCachedPayload();
    res.json({ hasCache: true, metadata: payload.metadata });
  } catch {
    res.json({ hasCache: false });
  }
});

app.post(
  "/api/upload",
  upload.fields([
    { name: "productList", maxCount: 1 },
    { name: "revenueFiles", maxCount: 10 },
  ]),
  async (req, res) => {
    const uploadedFiles = [
      ...(req.files?.productList || []),
      ...(req.files?.revenueFiles || []),
    ];
    try {
      const productFile = req.files?.productList?.[0];
      const revenueFiles = (req.files?.revenueFiles || []).filter(
        (file) => !file.originalname.startsWith(".~"),
      );
      if (!productFile) throw new Error("请上传产品全量列表");
      if (!revenueFiles.length) throw new Error("请至少上传一份收入及直接成本明细表");

      await resetCache();
      const productOriginalName = safeFileName(productFile.originalname);
      const savedProductPath = path.join(
        UPLOAD_DIR,
        `product-list-${Date.now()}-${productOriginalName}`,
      );
      await fs.rename(productFile.path, savedProductPath);
      const savedRevenuePaths = [];
      const revenueOriginalNames = [];
      for (const file of revenueFiles) {
        const revenueOriginalName = safeFileName(file.originalname);
        revenueOriginalNames.push(revenueOriginalName);
        const savedPath = path.join(UPLOAD_DIR, `revenue-${Date.now()}-${revenueOriginalName}`);
        await fs.rename(file.path, savedPath);
        savedRevenuePaths.push(savedPath);
      }

      await runAnalyzer(savedProductPath, savedRevenuePaths);
      const payload = await readCachedPayload();
      payload.metadata.uploaded_files = {
        productList: productOriginalName,
        revenueFiles: revenueOriginalNames,
      };
      await writeCachedPayload(payload);
      res.json({ ok: true, metadata: payload.metadata, rows: payload.rows.slice(0, 50) });
    } catch (error) {
      for (const file of uploadedFiles) {
        await fs.rm(file.path, { force: true }).catch(() => {});
      }
      res.status(400).json({ ok: false, error: error.message });
    }
  },
);

app.get("/api/result", async (_req, res) => {
  try {
    const payload = await readCachedPayload();
    res.json({ ok: true, metadata: payload.metadata, rows: payload.rows });
  } catch {
    res.status(404).json({ ok: false, error: "还没有可用的分析结果" });
  }
});

app.get("/api/download", async (_req, res) => {
  try {
    await fs.access(RESULT_XLSX);
    res.download(RESULT_XLSX, "全量产品退市筛选结果.xlsx");
  } catch {
    res.status(404).json({ ok: false, error: "还没有可下载的分析结果" });
  }
});

app.delete("/api/cache", async (_req, res) => {
  await resetCache();
  res.json({ ok: true });
});

app.post("/api/chat", async (req, res) => {
  try {
    const question = normalize(req.body?.question);
    if (!question) throw new Error("请输入问题");
    const payload = await readCachedPayload();
    const localAnswer = answerLocally(question, payload);
    if (localAnswer && !GEMINI_API_KEY) {
      return res.json({ ok: true, answer: localAnswer, provider: "local" });
    }
    if (!GEMINI_API_KEY) {
      throw new Error("服务端未配置 GEMINI_API_KEY，无法调用模型问答");
    }

    const prompt = [
      "你是产品退市筛选分析助手。只能基于给定 JSON 上下文回答，不要编造不存在的数据。",
      "若问题需要完整明细但上下文只含样例，请说明可在结果表中下载查看完整数据。",
      `上下文：\n${createAiContext(payload)}`,
      `用户问题：${question}`,
    ].join("\n\n");
    const answer = await callGemini(prompt);
    res.json({ ok: true, answer, provider: GEMINI_MODEL });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Product delisting app listening on http://localhost:${PORT}`);
  console.log(GEMINI_API_KEY ? "Gemini Q&A enabled." : "Gemini Q&A disabled: set GEMINI_API_KEY in .env.");
});
