const uploadForm = document.querySelector("#uploadForm");
const productList = document.querySelector("#productList");
const revenueFiles = document.querySelector("#revenueFiles");
const productListFiles = document.querySelector("#productListFiles");
const revenueFileList = document.querySelector("#revenueFileList");
const statusText = document.querySelector("#statusText");
const downloadLink = document.querySelector("#downloadLink");
const resultBody = document.querySelector("#resultBody");
const summaryGrid = document.querySelector("#summaryGrid");
const clearCache = document.querySelector("#clearCache");
const searchInput = document.querySelector("#searchInput");
const typeFilter = document.querySelector("#typeFilter");
const chatForm = document.querySelector("#chatForm");
const questionInput = document.querySelector("#questionInput");
const chatLog = document.querySelector("#chatLog");

let allRows = [];
let metadata = null;

function formatNumber(value) {
  const number = Number(value || 0);
  return number.toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function setStatus(text) {
  statusText.textContent = text;
}

function setDownloadEnabled(enabled) {
  downloadLink.classList.toggle("disabled", !enabled);
}

function formatFileNames(files) {
  const names = Array.from(files || []).map((file) => file.name || file);
  if (!names.length) return "未选择文件";
  if (names.length === 1) return names[0];
  return `${names.length} 个文件：${names.join("、")}`;
}

function renderSelectedFiles() {
  productListFiles.textContent = formatFileNames(productList.files);
  revenueFileList.textContent = formatFileNames(revenueFiles.files);
}

function renderCachedFiles(uploadedFiles) {
  if (!uploadedFiles) return;
  if (uploadedFiles.productList) {
    productListFiles.textContent = uploadedFiles.productList;
  }
  if (uploadedFiles.revenueFiles?.length) {
    revenueFileList.textContent = formatFileNames(uploadedFiles.revenueFiles);
  }
}

function renderSummary() {
  const metrics = summaryGrid.querySelectorAll(".metric strong");
  if (!metadata) {
    metrics.forEach((metric) => {
      metric.textContent = "-";
    });
    return;
  }
  metrics[0].textContent = metadata.candidate_count ?? 0;
  metrics[1].textContent = metadata.delisting_type_counts?.["强制退市"] ?? 0;
  metrics[2].textContent = metadata.delisting_type_counts?.["建议退市"] ?? 0;
}

function rowMatches(row, query, type) {
  if (type && row["退市类型"] !== type) return false;
  if (!query) return true;
  const haystack = [
    row["产品编码"],
    row["产品名称"],
    row["产品状态"],
    row["部门"],
    row["退市类型"],
    row["理由"],
  ]
    .join(" ")
    .toLowerCase();
  return haystack.includes(query.toLowerCase());
}

function renderRows() {
  const query = searchInput.value.trim();
  const type = typeFilter.value;
  const rows = allRows.filter((row) => rowMatches(row, query, type)).slice(0, 500);

  if (!rows.length) {
    resultBody.innerHTML = `<tr><td colspan="9" class="empty">没有匹配的结果</td></tr>`;
    return;
  }

  resultBody.innerHTML = rows
    .map((row) => {
      const badgeClass = row["退市类型"] === "强制退市" ? "force" : "suggest";
      return `
        <tr>
          <td>${escapeHtml(row["产品编码"])}</td>
          <td>${escapeHtml(row["产品名称"])}</td>
          <td>${escapeHtml(row["产品状态"])}</td>
          <td>${escapeHtml(row["创建时间"])}</td>
          <td>${escapeHtml(row["部门"])}</td>
          <td class="num">${formatNumber(row["产品收入"])}</td>
          <td class="num">${formatNumber(row["产品毛利"])}</td>
          <td><span class="badge ${badgeClass}">${escapeHtml(row["退市类型"])}</span></td>
          <td>${escapeHtml(formatReason(row["理由"]))}</td>
        </tr>
      `;
    })
    .join("");
}

function formatReason(reason) {
  return String(reason ?? "")
    .split("；")
    .map((item) => item.replace(/^规则\s*\d+\s*[：:]\s*/, ""))
    .join("；");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function loadStatus() {
  const response = await fetch("/api/status");
  const data = await response.json();
  if (!data.hasCache) {
    metadata = null;
    allRows = [];
    setDownloadEnabled(false);
    setStatus("等待上传");
    renderSummary();
    renderRows();
    return;
  }
  metadata = data.metadata;
  setDownloadEnabled(true);
  setStatus(`已有缓存：${metadata.candidate_count} 条候选，上传新文件会覆盖当前缓存。`);
  renderCachedFiles(metadata.uploaded_files);
  renderSummary();
  await loadResult();
}

async function loadResult() {
  const response = await fetch("/api/result");
  if (!response.ok) return;
  const data = await response.json();
  metadata = data.metadata;
  allRows = data.rows || [];
  renderSummary();
  renderRows();
}

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData();
  if (!productList.files[0]) {
    setStatus("请先选择产品全量列表。");
    return;
  }
  if (!revenueFiles.files.length) {
    setStatus("请至少选择一份收入及直接成本明细表。");
    return;
  }
  formData.append("productList", productList.files[0]);
  for (const file of revenueFiles.files) {
    formData.append("revenueFiles", file);
  }

  setStatus("正在上传并分析，请稍候。");
  setDownloadEnabled(false);
  const response = await fetch("/api/upload", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) {
    setStatus(data.error || "上传分析失败。");
    return;
  }
  metadata = data.metadata;
  setStatus(`分析完成：${metadata.candidate_count} 条候选。`);
  renderCachedFiles(metadata.uploaded_files);
  setDownloadEnabled(true);
  await loadResult();
});

clearCache.addEventListener("click", async () => {
  const confirmed = window.confirm("确认删除服务器本地缓存和当前分析结果？");
  if (!confirmed) return;
  await fetch("/api/cache", { method: "DELETE" });
  metadata = null;
  allRows = [];
  productList.value = "";
  revenueFiles.value = "";
  renderSelectedFiles();
  setDownloadEnabled(false);
  setStatus("缓存已删除，可以重新上传。");
  renderSummary();
  renderRows();
});

searchInput.addEventListener("input", renderRows);
typeFilter.addEventListener("change", renderRows);
productList.addEventListener("change", renderSelectedFiles);
revenueFiles.addEventListener("change", renderSelectedFiles);

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
    document.querySelectorAll(".view").forEach((item) => item.classList.remove("active"));
    tab.classList.add("active");
    document.querySelector(`#${tab.dataset.view}View`).classList.add("active");
  });
});

function addMessage(className, text) {
  const node = document.createElement("div");
  node.className = className;
  node.textContent = text;
  chatLog.appendChild(node);
  chatLog.scrollTop = chatLog.scrollHeight;
  return node;
}

chatForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const question = questionInput.value.trim();
  if (!question) return;
  addMessage("user-msg", question);
  questionInput.value = "";
  const pending = addMessage("assistant-msg", "正在分析...");
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ question }),
  });
  const data = await response.json();
  pending.textContent = response.ok ? data.answer : data.error || "问答失败。";
});

loadStatus().catch((error) => {
  setStatus(error.message || "状态读取失败。");
});
