import "dotenv/config";
import express from "express";
import multer from "multer";
import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { randomUUID } from "node:crypto";
import { promisify } from "node:util";
import { fileURLToPath, pathToFileURL } from "node:url";

const execFileAsync = promisify(execFile);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 4000);
const LLM_API_KEY = process.env.LLM_API_KEY || "";
const LLM_BASE_URL = (process.env.LLM_BASE_URL || "https://onerouter.cmaiot.cn/v1").replace(/\/+$/, "");
const LLM_CHAT_COMPLETIONS_URL = process.env.LLM_CHAT_COMPLETIONS_URL || `${LLM_BASE_URL}/chat/completions`;
const LLM_MODEL = process.env.LLM_MODEL || "qwen3.7-max";
const DEFAULT_BUNDLED_PYTHON =
  "/Users/apple/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";
const PYTHON_BIN = process.env.PYTHON_BIN || DEFAULT_BUNDLED_PYTHON;
const ANALYZER_BIN = process.env.CHATBI_ANALYZER_BIN || "";

const STORAGE_DIR = process.env.CHATBI_STORAGE_DIR || process.env.STORAGE_DIR || path.join(__dirname, "storage");
const SESSION_COOKIE = "chatbi_sid";
const SESSIONS_DIR = path.join(STORAGE_DIR, "sessions");
const TMP_DIR = path.join(STORAGE_DIR, "tmp");
const SESSION_RETENTION_DAYS = Number(process.env.SESSION_RETENTION_DAYS || 3);
const SESSION_RETENTION_MS = SESSION_RETENTION_DAYS * 24 * 60 * 60 * 1000;
const TMP_RETENTION_MINUTES = Number(process.env.TMP_RETENTION_MINUTES || 30);
const TMP_RETENTION_MS = TMP_RETENTION_MINUTES * 60 * 1000;
const CLEANUP_INTERVAL_MS = 6 * 60 * 60 * 1000;
const configuredLocalAnswerDelayMs = Number(process.env.LOCAL_ANSWER_DELAY_MS ?? 900);
const LOCAL_ANSWER_DELAY_MS = Number.isFinite(configuredLocalAnswerDelayMs)
  ? Math.max(0, configuredLocalAnswerDelayMs)
  : 900;
const ANALYZER_SCRIPT =
  process.env.CHATBI_ANALYZER_SCRIPT || path.join(__dirname, "scripts", "analyze_delisting.py");
const ANALYZER_CWD = path.dirname(ANALYZER_BIN || ANALYZER_SCRIPT);

await fs.mkdir(SESSIONS_DIR, { recursive: true });
await fs.mkdir(TMP_DIR, { recursive: true });

const app = express();
app.set("trust proxy", 1);
app.use(express.json({ limit: "1mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({
  dest: TMP_DIR,
  limits: {
    fileSize: 100 * 1024 * 1024,
    files: 12,
  },
});

function normalize(value) {
  if (value == null || typeof value === "boolean") return "";
  return String(value).trim();
}

function sleep(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
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

function displayUploadName(name) {
  return safeFileName(name).replace(/^pending-\d+-\d+-/, "");
}

function isIgnoredUploadFile(file) {
  return normalize(file?.originalname).startsWith(".~");
}

async function removeUploadTempFiles(files) {
  await Promise.all((files || []).map((file) => fs.rm(file.path, { force: true }).catch(() => {})));
}

function parseCookies(header) {
  return Object.fromEntries(
    normalize(header)
      .split(";")
      .map((item) => item.trim())
      .filter(Boolean)
      .map((item) => {
        const separator = item.indexOf("=");
        if (separator === -1) return [item, ""];
        return [item.slice(0, separator), decodeURIComponent(item.slice(separator + 1))];
      }),
  );
}

function isValidSessionId(value) {
  return /^[a-f0-9-]{36}$/i.test(normalize(value));
}

function createSessionCookie(sessionId, req) {
  const parts = [
    `${SESSION_COOKIE}=${encodeURIComponent(sessionId)}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    "Max-Age=2592000",
  ];
  if (req.secure || req.headers["x-forwarded-proto"] === "https") {
    parts.push("Secure");
  }
  return parts.join("; ");
}

function sessionPaths(sessionId) {
  const base = path.join(SESSIONS_DIR, sessionId);
  const uploads = path.join(base, "uploads");
  const cache = path.join(base, "cache");
  return {
    base,
    uploads,
    cache,
    resultJson: path.join(cache, "analysis-result.json"),
    resultXlsx: path.join(cache, "全量产品退市筛选结果.xlsx"),
    inspectionJson: path.join(cache, "upload-inspection.json"),
  };
}

async function ensureSession(req, res, next) {
  try {
    const cookies = parseCookies(req.headers.cookie);
    let sessionId = cookies[SESSION_COOKIE];
    if (!isValidSessionId(sessionId)) {
      sessionId = randomUUID();
      res.setHeader("Set-Cookie", createSessionCookie(sessionId, req));
    }
    req.sessionId = sessionId;
    req.sessionPaths = sessionPaths(sessionId);
    await fs.mkdir(req.sessionPaths.uploads, { recursive: true });
    await fs.mkdir(req.sessionPaths.cache, { recursive: true });
    next();
  } catch (error) {
    next(error);
  }
}

async function removeDirContents(dir) {
  await fs.rm(dir, { recursive: true, force: true });
  await fs.mkdir(dir, { recursive: true });
}

async function newestMtimeMs(dir) {
  let stat;
  try {
    stat = await fs.stat(dir);
  } catch {
    return 0;
  }

  let newest = stat.mtimeMs;
  const entries = await fs.readdir(dir, { withFileTypes: true }).catch(() => []);
  for (const entry of entries) {
    const entryPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      newest = Math.max(newest, await newestMtimeMs(entryPath));
      continue;
    }
    const entryStat = await fs.stat(entryPath).catch(() => null);
    if (entryStat) newest = Math.max(newest, entryStat.mtimeMs);
  }
  return newest;
}

async function removeExpiredChildren(dir, cutoffMs) {
  const entries = await fs.readdir(dir, { withFileTypes: true }).catch(() => []);
  let removed = 0;
  for (const entry of entries) {
    const entryPath = path.join(dir, entry.name);
    const newest = entry.isDirectory()
      ? await newestMtimeMs(entryPath)
      : (await fs.stat(entryPath).catch(() => null))?.mtimeMs || 0;
    if (newest > 0 && newest < cutoffMs) {
      await fs.rm(entryPath, { recursive: true, force: true });
      removed += 1;
    }
  }
  return removed;
}

let cleanupRunning = false;
let cleanupTimer = null;

async function cleanupExpiredStorage() {
  if (cleanupRunning || !Number.isFinite(SESSION_RETENTION_MS) || SESSION_RETENTION_MS <= 0) return;
  cleanupRunning = true;
  try {
    const now = Date.now();
    const sessionCutoffMs = now - SESSION_RETENTION_MS;
    const tmpCutoffMs = Number.isFinite(TMP_RETENTION_MS) && TMP_RETENTION_MS > 0 ? now - TMP_RETENTION_MS : sessionCutoffMs;
    const removedSessions = await removeExpiredChildren(SESSIONS_DIR, sessionCutoffMs);
    const removedTmpFiles = await removeExpiredChildren(TMP_DIR, tmpCutoffMs);
    if (removedSessions || removedTmpFiles) {
      console.log(
        `Storage cleanup removed ${removedSessions} expired session(s) and ${removedTmpFiles} tmp item(s).`,
      );
    }
  } catch (error) {
    console.warn(`Storage cleanup failed: ${error.message}`);
  } finally {
    cleanupRunning = false;
  }
}

async function resetSessionStorage(paths) {
  await removeDirContents(paths.uploads);
  await removeDirContents(paths.cache);
}

async function readCachedPayload(paths) {
  const text = await fs.readFile(paths.resultJson, "utf8");
  return JSON.parse(text);
}

async function writeCachedPayload(paths, payload) {
  await fs.writeFile(paths.resultJson, JSON.stringify(payload, null, 2), "utf8");
}

async function runAnalyzerCli(args, errorPrefix) {
  const command = ANALYZER_BIN || PYTHON_BIN;
  const finalArgs = ANALYZER_BIN ? args : [ANALYZER_SCRIPT, ...args];
  try {
    const { stdout } = await execFileAsync(command, finalArgs, {
      cwd: ANALYZER_CWD,
      maxBuffer: 1024 * 1024 * 20,
    });
    const lastLine = stdout.trim().split("\n").filter(Boolean).at(-1);
    return lastLine ? JSON.parse(lastLine) : { ok: true };
  } catch (error) {
    const detail = error.stderr || error.stdout || error.message;
    throw new Error(`${errorPrefix}：${detail}`);
  }
}

async function runAnalyzer(productPath, revenuePaths, paths) {
  return runAnalyzerCli(
    [
    "--product-list",
    productPath,
    "--revenue-files",
    ...revenuePaths,
    "--output-json",
    paths.resultJson,
    "--output-xlsx",
    paths.resultXlsx,
    ],
    "分析脚本执行失败",
  );
}

async function runInspector(filePaths, paths) {
  return runAnalyzerCli(
    [
    "--inspect-files",
    ...filePaths,
    "--output-json",
    paths.inspectionJson,
    ],
    "预检脚本执行失败",
  );
}

async function runAnalyzerFromInspection(paths) {
  return runAnalyzerCli(
    [
    "--mapping-json",
    paths.inspectionJson,
    "--output-json",
    paths.resultJson,
    "--output-xlsx",
    paths.resultXlsx,
    ],
    "分析脚本执行失败",
  );
}

function getModelErrorMessage(data, status) {
  if (typeof data?.error === "string") return data.error;
  if (data?.error?.message) return data.error.message;
  if (data?.message) return data.message;
  return `模型接口调用失败：HTTP ${status}`;
}

async function callChatModel(prompt) {
  const requestBody = JSON.stringify({
    model: LLM_MODEL,
    messages: [{ role: "user", content: prompt }],
    temperature: 0.2,
  });
  try {
    const response = await fetch(LLM_CHAT_COMPLETIONS_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${LLM_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: requestBody,
    });
    const text = await response.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error(`模型接口返回非 JSON 内容：${text.slice(0, 120)}`);
    }
    if (!response.ok || data?.error) throw new Error(getModelErrorMessage(data, response.status));
    return data?.choices?.[0]?.message?.content?.trim() || "模型没有返回有效回答。";
  } catch (error) {
    if (error.message === "fetch failed") {
      throw new Error(`模型服务连接失败：无法访问 ${LLM_BASE_URL}，请检查网络、代理或模型网关配置。`);
    }
    throw new Error(error.message || "模型接口调用失败");
  }
}

function stripInspectionPaths(inspection) {
  return {
    ...inspection,
    files: (inspection.files || []).map(({ path: _path, ...file }) => file),
  };
}

function extractJsonObject(text) {
  const cleaned = text.replace(/```json|```/g, "").trim();
  const start = cleaned.indexOf("{");
  const end = cleaned.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) throw new Error("模型未返回可解析 JSON");
  return JSON.parse(cleaned.slice(start, end + 1));
}

function findSheet(inspection, fileId, sheetName) {
  const file = inspection.files.find((item) => item.id === fileId);
  if (!file) return null;
  return file.sheets.find((sheet) => sheet.name === sheetName) || null;
}

function columnsFromHeaders(sheet, type, aiColumns) {
  const source = type === "product" ? sheet.product_columns : sheet.revenue_columns;
  const columns = { ...source };
  for (const [field, header] of Object.entries(aiColumns || {})) {
    const index = sheet.headers.findIndex((item) => item === header);
    if (index >= 0) {
      columns[field] = { index, header, label: field };
    }
  }
  return columns;
}

async function refineInspectionWithModel(inspection) {
  if (!inspection.needs_ai || !LLM_API_KEY) return inspection;
  const prompt = [
    "你是 Excel 上传文件识别助手。只根据文件名、sheet 名、表头名、行数和月份范围判断，不要臆造不存在字段。",
    "请返回严格 JSON，不要解释。结构：",
    `{"selected":{"product_file_id":"","product_sheets":[{"file_id":"","sheet":"","columns":{"code":"","name":"","status":"","created":"","department":"","delisting_approval_completed":""}}],"revenue_sheets":[{"file_id":"","sheet":"","columns":{"code":"","revenue":"","gross_profit":"","month":""}}]},"warnings":[""]}`,
    "字段值必须是已给出的表头原文；无法确定就留空数组或空字符串。",
    `预检信息：\n${JSON.stringify(stripInspectionPaths(inspection), null, 2)}`,
  ].join("\n\n");

  try {
    const ai = extractJsonObject(await callChatModel(prompt));
    const selected = ai.selected || {};
    const productSheets = [];
    for (const item of selected.product_sheets || []) {
      const sheet = findSheet(inspection, item.file_id, item.sheet);
      if (!sheet) continue;
      productSheets.push({
        file_id: item.file_id,
        sheet: item.sheet,
        header_row: sheet.header_row,
        data_start_row: sheet.data_start_row,
        columns: columnsFromHeaders(sheet, "product", item.columns),
        row_count: sheet.product_row_count,
      });
    }
    const revenueSheets = [];
    for (const item of selected.revenue_sheets || []) {
      const sheet = findSheet(inspection, item.file_id, item.sheet);
      if (!sheet) continue;
      revenueSheets.push({
        file_id: item.file_id,
        sheet: item.sheet,
        header_row: sheet.header_row,
        data_start_row: sheet.data_start_row,
        columns: columnsFromHeaders(sheet, "revenue", item.columns),
        row_count: sheet.revenue_row_count,
        month_summary: sheet.month_summary,
      });
    }
    if (productSheets.length || revenueSheets.length) {
      inspection.selected = {
        product_file_id: selected.product_file_id || productSheets[0]?.file_id || inspection.selected.product_file_id,
        product_sheets: productSheets.length ? productSheets : inspection.selected.product_sheets,
        revenue_file_ids: revenueSheets.length
          ? [...new Set(revenueSheets.map((item) => item.file_id))]
          : inspection.selected.revenue_file_ids,
        revenue_sheets: revenueSheets.length ? revenueSheets : inspection.selected.revenue_sheets,
      };
      inspection.needs_ai = false;
    }
    inspection.warnings = [...(inspection.warnings || []), ...(ai.warnings || [])].filter(Boolean);
  } catch (error) {
    inspection.warnings = [...(inspection.warnings || []), `模型辅助识别失败：${error.message}`];
  }
  return inspection;
}

function createAiContext(payload) {
  return JSON.stringify(
    {
      口径: {
        基准日期: payload.metadata.as_of_date,
        两年收入窗口: payload.metadata.revenue_window,
        规则123状态范围: payload.metadata.active_statuses_for_rules_1_to_3,
        规则4状态范围: "退市中，且退市审批完成时间超过1年",
        规则4审批时间阈值: payload.metadata.rule4_approval_before,
      },
      汇总: {
        候选总数: payload.metadata.candidate_count,
        退市类型分布: payload.metadata.delisting_type_counts,
        命中规则分布: payload.metadata.rule_counts,
        候选状态分布: payload.metadata.candidate_status_counts,
        无法判断规则4数量: payload.metadata.missing_rule4_approval_count || 0,
      },
      样例明细: payload.rows.slice(0, 80),
      无法判断的产品样例: (payload.missing_rule4_approval_rows || []).slice(0, 40),
    },
    null,
    2,
  );
}

function answerLocally(question, payload, options = {}) {
  const includeNarrative = Boolean(options.includeNarrative);
  const q = question.trim();
  const metadata = payload.metadata || {};
  const rows = payload.rows || [];
  const missingRows = payload.missing_rule4_approval_rows || [];
  const delistingCounts = metadata.delisting_type_counts || {};
  const ruleCounts = metadata.rule_counts || {};
  const statusCounts = metadata.candidate_status_counts || {};

  const countBy = (items, key) => {
    const counts = new Map();
    for (const item of items) {
      const value = normalize(item[key]) || "未填写";
      counts.set(value, (counts.get(value) || 0) + 1);
    }
    return [...counts.entries()].sort((a, b) => b[1] - a[1]);
  };

  const formatTopCounts = (counts, limit = 5) => {
    if (!counts.length) return "暂无可统计数据";
    return counts.slice(0, limit).map(([name, count], index) => `${index + 1}. ${name}：${count} 个`).join("\n");
  };

  if (includeNarrative && /结论|总结|分析结果|得出什么/.test(q)) {
    const topRule = Object.entries(ruleCounts).sort((a, b) => Number(b[1] || 0) - Number(a[1] || 0))[0];
    const topDepartments = countBy(rows, "部门").slice(0, 3).map(([name, count]) => `${name} ${count} 个`).join("，") || "暂无";
    return [
      `本次分析共识别出 ${metadata.candidate_count || 0} 个退市候选产品。`,
      `其中强制退市 ${delistingCounts["强制退市"] || 0} 个，建议退市 ${delistingCounts["建议退市"] || 0} 个；规则命中最多的是规则${topRule?.[0] || "-"}（${topRule?.[1] || 0} 个）。`,
      `部门分布靠前的是：${topDepartments}。`,
      `另有 ${metadata.missing_rule4_approval_count || 0} 个退市中产品因缺少退市审批完成时间，暂时无法判断规则4。`,
    ].join("\n");
  }

  if (includeNarrative && /后续工作|团队|工作建议|怎么推进/.test(q)) {
    return [
      "建议按优先级推进：",
      `1. 先处理强制退市产品 ${delistingCounts["强制退市"] || 0} 个，优先核对规则1和规则2命中的产品口径。`,
      `2. 对建议退市产品 ${delistingCounts["建议退市"] || 0} 个做业务复核，重点确认近两年收入和毛利是否完整。`,
      `3. 补齐 ${metadata.missing_rule4_approval_count || 0} 个退市中产品的退市审批完成时间，避免规则4长期无法判断。`,
      "4. 将候选数量靠前的部门作为第一批沟通对象，逐项确认保留、退市或补数原因。",
    ].join("\n");
  }

  if (/部门/.test(q) && /最多|排行|排名|分布|应退未退/.test(q)) {
    const departmentCounts = countBy(rows, "部门");
    return `按当前退市候选结果统计，产品数量最多的部门如下：\n${formatTopCounts(departmentCounts)}`;
  }

  if (/强制/.test(q)) {
    return `强制退市 ${delistingCounts["强制退市"] || 0} 条。`;
  }
  if (/建议退市|建议类|建议产品/.test(q)) {
    return `建议退市 ${delistingCounts["建议退市"] || 0} 条。`;
  }
  if (/规则\s*1|规则一/.test(q)) return `规则1命中 ${ruleCounts["1"] || 0} 条。`;
  if (/规则\s*2|规则二/.test(q)) return `规则2命中 ${ruleCounts["2"] || 0} 条。`;
  if (/规则\s*3|规则三/.test(q)) return `规则3命中 ${ruleCounts["3"] || 0} 条。`;
  if (/无法判断|缺少|缺失|没有/.test(q) && /审批|退市审批|完成时间|规则\s*4|规则四/.test(q)) {
    const topDepartments = countBy(missingRows, "部门").slice(0, 3).map(([name, count]) => `${name} ${count} 个`).join("，");
    return `无法判断规则4的产品 ${metadata.missing_rule4_approval_count || 0} 条；这些产品状态为退市中，因缺少退市审批完成时间，无法判断是否退市中超过1年。${topDepartments ? `部门分布靠前的是：${topDepartments}。` : ""}`;
  }
  if (/规则\s*4|规则四|退市中/.test(q)) {
    return `规则4命中 ${ruleCounts["4"] || 0} 条；另有 ${metadata.missing_rule4_approval_count || 0} 条产品因缺少退市审批完成时间而无法判断规则4。`;
  }
  if (/已入库/.test(q)) {
    return `候选结果中“已入库”状态 ${statusCounts["已入库"] || 0} 条。`;
  }
  return "";
}

app.use("/api", ensureSession);

app.get("/api/status", async (req, res) => {
  try {
    const payload = await readCachedPayload(req.sessionPaths);
    res.json({ hasCache: true, metadata: payload.metadata });
  } catch {
    res.json({ hasCache: false });
  }
});

app.post("/api/inspect-upload", upload.array("analysisFiles", 12), async (req, res) => {
  const uploadedFiles = req.files || [];
  const savedPaths = [];
  const paths = req.sessionPaths;
  try {
    const ignoredFiles = uploadedFiles.filter(isIgnoredUploadFile);
    const files = uploadedFiles.filter((file) => !isIgnoredUploadFile(file));
    await removeUploadTempFiles(ignoredFiles);
    if (!files.length) throw new Error("请至少上传一份 Excel 文件");

    await resetSessionStorage(paths);
    for (const [index, file] of files.entries()) {
      const originalName = safeFileName(file.originalname);
      const savedPath = path.join(paths.uploads, `pending-${Date.now()}-${index + 1}-${originalName}`);
      await fs.rename(file.path, savedPath);
      savedPaths.push(savedPath);
    }

    await runInspector(savedPaths, paths);
    let inspection = JSON.parse(await fs.readFile(paths.inspectionJson, "utf8"));
    inspection.files = (inspection.files || []).map((file) => ({
      ...file,
      name: displayUploadName(file.name),
    }));
    inspection = await refineInspectionWithModel(inspection);
    await fs.writeFile(paths.inspectionJson, JSON.stringify(inspection, null, 2), "utf8");
    res.json({ ok: true, inspection: stripInspectionPaths(inspection) });
  } catch (error) {
    for (const file of uploadedFiles) {
      await fs.rm(file.path, { force: true }).catch(() => {});
    }
    for (const filePath of savedPaths) {
      await fs.rm(filePath, { force: true }).catch(() => {});
    }
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post("/api/analyze-inspection", async (req, res) => {
  const paths = req.sessionPaths;
  try {
    const inspection = JSON.parse(await fs.readFile(paths.inspectionJson, "utf8"));
    if (!inspection.selected?.product_file_id || !inspection.selected?.product_sheets?.length) {
      throw new Error("未识别到产品全量列表，无法分析");
    }
    if (!inspection.selected?.revenue_sheets?.length) {
      throw new Error("未识别到收入成本明细表，无法分析");
    }
    await runAnalyzerFromInspection(paths);
    const payload = await readCachedPayload(paths);
    const filesById = Object.fromEntries((inspection.files || []).map((file) => [file.id, file]));
    payload.metadata.uploaded_files = {
      productList: filesById[inspection.selected.product_file_id]?.name || "",
      revenueFiles: [...new Set(inspection.selected.revenue_sheets.map((item) => item.file_id))]
        .map((fileId) => filesById[fileId]?.name)
        .filter(Boolean),
    };
    payload.metadata.inspection_warnings = inspection.warnings || [];
    await writeCachedPayload(paths, payload);
    res.json({
      ok: true,
      metadata: payload.metadata,
      rows: payload.rows.slice(0, 50),
      missing_rule4_approval_rows: (payload.missing_rule4_approval_rows || []).slice(0, 50),
    });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post(
  "/api/upload",
  upload.fields([
    { name: "productList", maxCount: 1 },
    { name: "revenueFiles", maxCount: 10 },
  ]),
  async (req, res) => {
    const paths = req.sessionPaths;
    const uploadedFiles = [
      ...(req.files?.productList || []),
      ...(req.files?.revenueFiles || []),
    ];
    try {
      const ignoredFiles = uploadedFiles.filter(isIgnoredUploadFile);
      await removeUploadTempFiles(ignoredFiles);
      const productFile = (req.files?.productList || []).filter((file) => !isIgnoredUploadFile(file))[0];
      const revenueFiles = (req.files?.revenueFiles || []).filter((file) => !isIgnoredUploadFile(file));
      if (!productFile) throw new Error("请上传产品全量列表");
      if (!revenueFiles.length) throw new Error("请至少上传一份收入及直接成本明细表");

      await resetSessionStorage(paths);
      const productOriginalName = safeFileName(productFile.originalname);
      const savedProductPath = path.join(
        paths.uploads,
        `product-list-${Date.now()}-${productOriginalName}`,
      );
      await fs.rename(productFile.path, savedProductPath);
      const savedRevenuePaths = [];
      const revenueOriginalNames = [];
      for (const file of revenueFiles) {
        const revenueOriginalName = safeFileName(file.originalname);
        revenueOriginalNames.push(revenueOriginalName);
        const savedPath = path.join(paths.uploads, `revenue-${Date.now()}-${revenueOriginalName}`);
        await fs.rename(file.path, savedPath);
        savedRevenuePaths.push(savedPath);
      }

      await runAnalyzer(savedProductPath, savedRevenuePaths, paths);
      const payload = await readCachedPayload(paths);
      payload.metadata.uploaded_files = {
        productList: productOriginalName,
        revenueFiles: revenueOriginalNames,
      };
      await writeCachedPayload(paths, payload);
      res.json({
        ok: true,
        metadata: payload.metadata,
        rows: payload.rows.slice(0, 50),
        missing_rule4_approval_rows: (payload.missing_rule4_approval_rows || []).slice(0, 50),
      });
    } catch (error) {
      for (const file of uploadedFiles) {
        await fs.rm(file.path, { force: true }).catch(() => {});
      }
      res.status(400).json({ ok: false, error: error.message });
    }
  },
);

app.get("/api/result", async (req, res) => {
  try {
    const payload = await readCachedPayload(req.sessionPaths);
    res.json({
      ok: true,
      metadata: payload.metadata,
      rows: payload.rows,
      missing_rule4_approval_rows: payload.missing_rule4_approval_rows || [],
    });
  } catch {
    res.status(404).json({ ok: false, error: "还没有可用的分析结果" });
  }
});

app.get("/api/download", async (req, res) => {
  try {
    await fs.access(req.sessionPaths.resultXlsx);
    res.download(req.sessionPaths.resultXlsx, "全量产品退市筛选结果.xlsx");
  } catch {
    res.status(404).json({ ok: false, error: "还没有可下载的分析结果" });
  }
});

app.delete("/api/cache", async (req, res) => {
  await resetSessionStorage(req.sessionPaths);
  res.json({ ok: true });
});

app.post("/api/chat", async (req, res) => {
  try {
    const question = normalize(req.body?.question);
    if (!question) throw new Error("请输入问题");
    let payload;
    try {
      payload = await readCachedPayload(req.sessionPaths);
    } catch {
      throw new Error("当前会话还没有可用的分析结果，请先上传并分析。");
    }
    const localAnswer = answerLocally(question, payload, { includeNarrative: !LLM_API_KEY });
    if (localAnswer) {
      if (LOCAL_ANSWER_DELAY_MS > 0) {
        await sleep(LOCAL_ANSWER_DELAY_MS);
      }
      return res.json({ ok: true, answer: localAnswer, provider: "local" });
    }
    if (!LLM_API_KEY) {
      throw new Error("服务端未配置 LLM_API_KEY，无法调用模型问答");
    }

    const prompt = [
      "你是产品退市筛选分析助手。只能基于给定 JSON 上下文回答，不要编造不存在的数据。",
      "若问题需要完整明细但上下文只含样例，请说明可在结果表中下载查看完整数据。",
      `上下文：\n${createAiContext(payload)}`,
      `用户问题：${question}`,
    ].join("\n\n");
    const answer = await callChatModel(prompt);
    res.json({ ok: true, answer, provider: LLM_MODEL });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

export async function startServer(options = {}) {
  const port = Number(options.port ?? PORT);
  const host = options.host;
  await cleanupExpiredStorage();
  if (!cleanupTimer) {
    cleanupTimer = setInterval(cleanupExpiredStorage, CLEANUP_INTERVAL_MS);
    cleanupTimer.unref?.();
  }

  return new Promise((resolve, reject) => {
    const server = app.listen(port, host, () => {
      const address = server.address();
      const actualPort = typeof address === "object" && address ? address.port : port;
      const actualHost = host || "localhost";
      console.log(`Product delisting app listening on http://${actualHost}:${actualPort}`);
      console.log(LLM_API_KEY ? `Model Q&A enabled: ${LLM_MODEL}.` : "Model Q&A disabled: set LLM_API_KEY in .env.");
      resolve({ app, server, port: actualPort, host: actualHost });
    });
    server.once("error", reject);
  });
}

export { app };

function isMainModule() {
  return process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href;
}

if (isMainModule()) {
  await startServer();
}
