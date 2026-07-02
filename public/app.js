const uploadForm = document.querySelector("#uploadForm");
const analysisFiles = document.querySelector("#analysisFiles");
const analysisFileList = document.querySelector("#analysisFileList");
const analysisFilesIcon = document.querySelector("#analysisFilesIcon");
const selectedFileList = document.querySelector("#selectedFileList");
const inspectUpload = document.querySelector("#inspectUpload");
const confirmAnalysis = document.querySelector("#confirmAnalysis");
const inspectionPanel = document.querySelector("#inspectionPanel");
const inspectionLoading = document.querySelector("#inspectionLoading");
const uploadStepItems = document.querySelectorAll("#uploadStepper [data-step]");
const uploadStepFiles = document.querySelector("#uploadStepFiles");
const uploadStepInspection = document.querySelector("#uploadStepInspection");
const toast = document.querySelector("#toast");
const downloadLink = document.querySelector("#downloadLink");
const missingApprovalNotice = document.querySelector("#missingApprovalNotice");
const missingApprovalCount = document.querySelector("#missingApprovalCount");
const missingApprovalAction = document.querySelector("#missingApprovalAction");
const resultBody = document.querySelector("#resultBody");
const summaryGrid = document.querySelector("#summaryGrid");
const searchInput = document.querySelector("#searchInput");
const chatForm = document.querySelector("#chatForm");
const questionInput = document.querySelector("#questionInput");
const chatLog = document.querySelector("#chatLog");
const newChat = document.querySelector("#newChat");
const analysisArea = document.querySelector("#analysisArea");
const chatDrawer = document.querySelector("#chatDrawer");
const openChat = document.querySelector("#openChat");
const closeChat = document.querySelector("#closeChat");
const openUpload = document.querySelector("#openUpload");
const closeUpload = document.querySelector("#closeUpload");
const cancelUpload = document.querySelector("#cancelUpload");
const backToFiles = document.querySelector("#backToFiles");
const uploadModal = document.querySelector("#uploadModal");
const mainPanel = document.querySelector(".main-panel");
const listLoading = document.querySelector("#listLoading");
const loadingCodeStream = document.querySelector(".code-stream");
const loadingCodeLineNodes = loadingCodeStream ? Array.from(loadingCodeStream.querySelectorAll("span")) : [];
const loadingCodeLines = loadingCodeLineNodes.map((node) => node.textContent);

let allRows = [];
let missingApprovalRows = [];
let showingMissingApproval = false;
let metadata = null;
let toastTimer = null;
let pendingInspection = null;
let inspectionRequestId = 0;
let selectedAnalysisFiles = [];
let uploadStep = 1;
let loadingCodeTimer = null;
let loadingCodeRunId = 0;

const analysisFilesHint = "支持 .xlsx、.xls、.xlsm，最多 12 个文件，单个不超过 100MB；系统会自动识别产品列表和收入成本明细表";
const chatSuggestionQuestions = [
  "根据分析结果能得出什么结论？",
  "对团队后续工作有什么建议？",
  "哪些部门应退未退的产品最多？",
];
const PUBLIC_FIELD_LABELS = {
  code: "产品编码",
  name: "产品名称",
  status: "产品状态",
  created: "创建时间",
  department: "部门",
  delisting_approval_completed: "退市审批完成时间",
  revenue: "收入",
  gross_profit: "毛利",
  month: "月份",
};

function formatNumber(value) {
  const number = Number(value || 0);
  return number.toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function showToast(text) {
  window.clearTimeout(toastTimer);
  toast.textContent = text;
  toast.classList.add("visible");
  toastTimer = window.setTimeout(() => {
    toast.classList.remove("visible");
  }, 3200);
}

function appendInlineMarkdown(parent, text) {
  const pattern = /(\*\*([\s\S]+?)\*\*|`([^`]+)`|\[([^\]]+)\]\((https?:\/\/[^\s)]+)\))/g;
  let lastIndex = 0;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parent.appendChild(document.createTextNode(text.slice(lastIndex, match.index)));
    }

    if (match[2]) {
      const strong = document.createElement("strong");
      strong.textContent = match[2];
      parent.appendChild(strong);
    } else if (match[3]) {
      const code = document.createElement("code");
      code.textContent = match[3];
      parent.appendChild(code);
    } else if (match[4] && match[5]) {
      const link = document.createElement("a");
      link.textContent = match[4];
      link.href = match[5];
      link.target = "_blank";
      link.rel = "noreferrer";
      parent.appendChild(link);
    }

    lastIndex = pattern.lastIndex;
  }

  if (lastIndex < text.length) {
    parent.appendChild(document.createTextNode(text.slice(lastIndex)));
  }
}

function appendMarkdownParagraph(parent, lines) {
  if (!lines.length) return;
  const paragraph = document.createElement("p");
  lines.forEach((line, index) => {
    if (index > 0) paragraph.appendChild(document.createElement("br"));
    appendInlineMarkdown(paragraph, line);
  });
  parent.appendChild(paragraph);
}

function renderMarkdown(node, text) {
  node.replaceChildren();

  const lines = String(text || "").replace(/\r\n?/g, "\n").split("\n");
  let paragraphLines = [];
  let activeList = null;

  const flushParagraph = () => {
    appendMarkdownParagraph(node, paragraphLines);
    paragraphLines = [];
  };

  const flushList = () => {
    if (activeList) {
      node.appendChild(activeList);
      activeList = null;
    }
  };

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      flushParagraph();
      flushList();
      continue;
    }

    const heading = trimmed.match(/^(#{1,4})\s+(.+)$/);
    if (heading) {
      flushParagraph();
      flushList();
      const headingNode = document.createElement(heading[1].length <= 2 ? "h3" : "h4");
      appendInlineMarkdown(headingNode, heading[2]);
      node.appendChild(headingNode);
      continue;
    }

    const orderedItem = trimmed.match(/^(\d+)[.)]\s+(.+)$/);
    const bulletItem = trimmed.match(/^[-*+]\s+(.+)$/);
    const item = orderedItem || bulletItem;
    if (item) {
      flushParagraph();
      const tagName = orderedItem ? "ol" : "ul";
      if (!activeList || activeList.tagName.toLowerCase() !== tagName) {
        flushList();
        activeList = document.createElement(tagName);
        if (orderedItem && Number(orderedItem[1]) > 1) {
          activeList.start = Number(orderedItem[1]);
        }
      }
      const listItem = document.createElement("li");
      appendInlineMarkdown(listItem, orderedItem ? orderedItem[2] : item[1]);
      activeList.appendChild(listItem);
      continue;
    }

    flushList();
    paragraphLines.push(trimmed);
  }

  flushParagraph();
  flushList();
}

function setDownloadEnabled(enabled) {
  downloadLink.classList.toggle("disabled", !enabled);
}

function updateMissingApprovalToggle() {
  const count = missingApprovalRows.length || metadata?.missing_rule4_approval_count || 0;
  missingApprovalNotice.hidden = count === 0;
  missingApprovalCount.textContent = count;
  missingApprovalAction.textContent = showingMissingApproval ? "返回" : "查看";
  missingApprovalNotice.classList.toggle("disabled", count === 0);
  missingApprovalNotice.classList.toggle("active", showingMissingApproval);
  missingApprovalNotice.setAttribute("aria-disabled", String(count === 0));
  missingApprovalNotice.setAttribute("aria-pressed", String(showingMissingApproval));
  missingApprovalNotice.tabIndex = count === 0 ? -1 : 0;
}

function formatFileNames(files) {
  const names = Array.from(files || []).map((file) => file.name || file);
  if (!names.length) return "";
  if (names.length === 1) return names[0];
  return `${names.length} 个文件：${names.join("、")}`;
}

function fileKey(file) {
  return [file.name, file.size, file.lastModified].join("|");
}

function resetUploadCard() {
  analysisFileList.textContent = analysisFilesHint;
  analysisFilesIcon.classList.remove("has-file");
}

function renderSelectedFiles() {
  resetUploadCard();
  selectedFileList.innerHTML = selectedAnalysisFiles
    .map((file, index) => `
      <div class="selected-file-item">
        <span class="selected-file-name">${escapeHtml(file.name)}</span>
        <button class="file-remove" type="button" data-file-index="${index}" aria-label="删除 ${escapeHtml(file.name)}">删除</button>
      </div>
    `)
    .join("");
  inspectUpload.classList.toggle("disabled", !selectedAnalysisFiles.length);
}

function setUploadStep(step) {
  uploadStep = step;
  uploadStepFiles.hidden = step !== 1;
  uploadStepInspection.hidden = step !== 2;
  uploadStepItems.forEach((item) => {
    const itemStep = Number(item.dataset.step);
    item.classList.toggle("active", itemStep === step);
    item.classList.toggle("done", itemStep < step);
    const label = item.querySelector("span");
    if (label && itemStep === 1) {
      label.textContent = step > 1 ? "文件已上传" : "上传文件";
    }
  });
  cancelUpload.hidden = step !== 1;
  backToFiles.hidden = step !== 2;
  inspectUpload.hidden = step !== 1;
  confirmAnalysis.hidden = step !== 2;
}

function resetInspectionState() {
  pendingInspection = null;
  inspectionPanel.hidden = true;
  confirmAnalysis.classList.add("disabled");
}

function setUploadOpen(isOpen) {
  uploadModal.hidden = !isOpen;
  document.body.classList.toggle("modal-open", isOpen);
  if (isOpen) {
    if (!selectedAnalysisFiles.length) {
      clearUploadQueue();
    }
    const initialStep = canAnalyzeInspection(pendingInspection) ? 2 : 1;
    setUploadStep(initialStep);
    renderSelectedFiles();
    if (initialStep === 1) {
      window.setTimeout(() => analysisFiles.focus(), 80);
    } else {
      window.setTimeout(() => confirmAnalysis.focus(), 80);
    }
  } else {
    openUpload.focus();
  }
}

function canAnalyzeInspection(inspection) {
  return Boolean(inspection?.selected?.product_file_id && inspection?.selected?.revenue_sheets?.length);
}

function setInspectionLoading(isLoading) {
  inspectionLoading.hidden = !isLoading;
  analysisFiles.disabled = isLoading;
  selectedFileList.classList.toggle("is-disabled", isLoading);
  inspectUpload.classList.toggle("disabled", isLoading || !selectedAnalysisFiles.length);
  confirmAnalysis.classList.toggle("disabled", isLoading || !canAnalyzeInspection(pendingInspection));
}

function stopLoadingCode() {
  loadingCodeRunId += 1;
  window.clearTimeout(loadingCodeTimer);
  loadingCodeTimer = null;
  loadingCodeLineNodes.forEach((node, index) => {
    node.textContent = loadingCodeLines[index] || "";
    node.classList.remove("is-active");
    node.classList.add("is-typed");
  });
}

function renderLoadingCodeLine(node, text, showCaret) {
  node.textContent = text;
  if (!showCaret) return;
  const caret = document.createElement("span");
  caret.className = "typing-caret";
  caret.setAttribute("aria-hidden", "true");
  caret.textContent = "\u00a0";
  node.appendChild(caret);
}

function startLoadingCode() {
  if (!loadingCodeLineNodes.length) return;
  const runId = ++loadingCodeRunId;
  window.clearTimeout(loadingCodeTimer);
  loadingCodeLineNodes.forEach((node) => {
    node.textContent = "";
    node.classList.remove("is-active", "is-typed");
  });
  loadingCodeStream.scrollTop = 0;

  let lineIndex = 0;
  let charIndex = 0;
  const tick = () => {
    if (runId !== loadingCodeRunId) return;
    if (lineIndex >= loadingCodeLines.length) {
      loadingCodeTimer = window.setTimeout(() => {
        if (runId === loadingCodeRunId) startLoadingCode();
      }, 720);
      return;
    }

    const node = loadingCodeLineNodes[lineIndex];
    const code = loadingCodeLines[lineIndex] || "";
    node.classList.add("is-active");
    renderLoadingCodeLine(node, code.slice(0, charIndex), true);
    loadingCodeStream.scrollTop = loadingCodeStream.scrollHeight;

    if (charIndex < code.length) {
      charIndex += 1;
      loadingCodeTimer = window.setTimeout(tick, 18);
      return;
    }

    node.classList.remove("is-active");
    node.classList.add("is-typed");
    node.textContent = code;
    lineIndex += 1;
    charIndex = 0;
    loadingCodeTimer = window.setTimeout(tick, 130);
  };

  tick();
}

function setListLoading(isLoading) {
  mainPanel.classList.toggle("is-loading", isLoading);
  mainPanel.setAttribute("aria-busy", String(isLoading));
  listLoading.hidden = !isLoading;
  if (isLoading) {
    startLoadingCode();
  } else {
    stopLoadingCode();
  }
}

function clearUploadQueue() {
  selectedAnalysisFiles = [];
  analysisFiles.value = "";
  resetUploadCard();
  selectedFileList.innerHTML = "";
  inspectUpload.classList.add("disabled");
  resetInspectionState();
  setInspectionLoading(false);
}

function roleLabel(role) {
  if (role === "product_list") return "产品列表";
  if (role === "revenue_detail") return "收入明细";
  return "未确定";
}

function selectedMappingsForFile(inspection, fileId, type) {
  const key = type === "product_list" ? "product_sheets" : "revenue_sheets";
  return (inspection.selected?.[key] || [])
    .filter((item) => item.file_id === fileId);
}

function recognizedLabels(mappings, fields) {
  const labels = [];
  for (const field of fields) {
    const matched = mappings.some((mapping) => mapping.columns?.[field]?.header);
    if (matched) labels.push(PUBLIC_FIELD_LABELS[field] || field);
  }
  return labels;
}

function inspectionDescription(file, inspection) {
  const mappings = selectedMappingsForFile(inspection, file.id, file.role);
  if (file.role === "product_list") {
    if (!mappings.length) {
      return "该文件包含部分产品字段，但未通过产品全量列表预检，请查看下方提示。";
    }
    const labels = recognizedLabels(mappings, ["code", "name", "status", "created", "department", "delisting_approval_completed"]);
    return `已识别为产品全量列表，识别到${labels.length ? labels.join("、") : "关键"}等字段。`;
  }
  if (file.role === "revenue_detail") {
    if (!mappings.length) {
      return "该文件包含部分经营字段，但未通过收入成本明细表预检，请查看下方提示。";
    }
    const labels = recognizedLabels(mappings, ["code", "revenue", "gross_profit", "month"]);
    return `已识别为收入成本明细表，识别到${labels.length ? labels.join("、") : "关键"}等字段。`;
  }
  return "未识别为产品全量列表或收入成本明细表，请检查文件内容。";
}

function renderInspection(inspection) {
  pendingInspection = inspection;
  const productName = inspection.files.find((file) => file.id === inspection.selected?.product_file_id)?.name || "未识别";
  const revenueCount = inspection.selected?.revenue_file_ids?.length || 0;
  const items = inspection.files
    .map((file) => {
      const roleClass = file.role === "product_list" ? "product" : file.role === "revenue_detail" ? "revenue" : "";
      return `
        <div class="inspection-item">
          <div>
            <div class="inspection-name">${escapeHtml(file.name)}</div>
            <div class="inspection-meta">${escapeHtml(inspectionDescription(file, inspection))}</div>
          </div>
          <span class="role-pill ${roleClass}">${roleLabel(file.role)}</span>
        </div>
      `;
    })
    .join("");
  const warnings = (inspection.warnings || [])
    .map((warning) => `<div class="inspection-warning">${escapeHtml(warning)}</div>`)
    .join("");
  const revenueMonthSummary = inspection.selected?.revenue_month_summary;
  const revenueRange = revenueMonthSummary?.start_month
    ? `；月份覆盖：${escapeHtml(revenueMonthSummary.start_month)} 至 ${escapeHtml(revenueMonthSummary.end_month)}`
    : "";
  const resultNote = warnings || `<div class="inspection-ok">预检未发现阻断项，可以开始分析。</div>`;
  inspectionPanel.innerHTML = `
    <div class="inspection-title">
      <span>预检结果</span>
      <span>产品列表：${escapeHtml(productName)}；收入明细：${revenueCount} 个文件${revenueRange}</span>
    </div>
    <div class="inspection-list">${items}</div>
    ${resultNote}
  `;
  inspectionPanel.hidden = false;
  confirmAnalysis.classList.toggle("disabled", !canAnalyzeInspection(inspection));
}

async function inspectSelectedFiles() {
  renderSelectedFiles();
  resetInspectionState();
  setUploadStep(2);
  const requestId = ++inspectionRequestId;
  const formData = new FormData();
  if (!selectedAnalysisFiles.length) return;
  for (const file of selectedAnalysisFiles) {
    formData.append("analysisFiles", file);
  }

  setInspectionLoading(true);
  setDownloadEnabled(false);
  try {
    const response = await fetch("/api/inspect-upload", { method: "POST", body: formData });
    const data = await response.json();
    if (requestId !== inspectionRequestId) return;
    if (!response.ok) {
      showToast(data.error || "上传预检失败。");
      return;
    }
    renderInspection(data.inspection);
    showToast("预检完成，请确认识别结果。");
  } catch (error) {
    if (requestId === inspectionRequestId) {
      showToast(error.message || "上传预检失败。");
    }
  } finally {
    if (requestId === inspectionRequestId) {
      setInspectionLoading(false);
    }
  }
}

function addSelectedFiles(files) {
  const incoming = Array.from(files || []);
  if (!incoming.length) return false;
  const existingKeys = new Set(selectedAnalysisFiles.map(fileKey));
  const merged = [...selectedAnalysisFiles];
  let changed = false;
  for (const file of incoming) {
    if (!existingKeys.has(fileKey(file))) {
      merged.push(file);
      existingKeys.add(fileKey(file));
      changed = true;
    }
  }
  selectedAnalysisFiles = merged.slice(0, 12);
  if (merged.length > 12) {
    showToast("最多只能选择 12 个文件，已保留前 12 个。");
  }
  return changed;
}

function removeSelectedFile(index) {
  selectedAnalysisFiles = selectedAnalysisFiles.filter((_, itemIndex) => itemIndex !== index);
  analysisFiles.value = "";
  ++inspectionRequestId;
  setUploadStep(1);
  renderSelectedFiles();
  resetInspectionState();
  setInspectionLoading(false);
}

function renderSummary() {
  const metrics = summaryGrid.querySelectorAll(".metric-value");
  if (!metadata) {
    metrics.forEach((metric) => {
      metric.textContent = "-";
    });
    updateMissingApprovalToggle();
    return;
  }
  metrics[0].textContent = metadata.candidate_count ?? 0;
  metrics[1].textContent = metadata.delisting_type_counts?.["强制退市"] ?? 0;
  metrics[2].textContent = metadata.delisting_type_counts?.["建议退市"] ?? 0;
  updateMissingApprovalToggle();
}

function rowMatches(row, query) {
  if (!query) return true;
  const haystack = [
    row["产品编码"],
    row["产品名称"],
    row["部门"],
  ]
    .join(" ")
    .toLowerCase();
  return haystack.includes(query.toLowerCase());
}

function renderRows() {
  const query = searchInput.value.trim();
  const sourceRows = showingMissingApproval ? missingApprovalRows : allRows;
  const rows = sourceRows.filter((row) => rowMatches(row, query)).slice(0, 500);

  if (!rows.length) {
    const text = showingMissingApproval ? "没有无法判断的产品" : "没有匹配的结果";
    resultBody.innerHTML = `<tr><td colspan="9" class="empty">${text}</td></tr>`;
    return;
  }

  resultBody.innerHTML = rows
    .map((row) => {
      const isMissingApproval = showingMissingApproval || row["问题"];
      const badgeClass = isMissingApproval ? "pending" : row["退市类型"] === "强制退市" ? "force" : "suggest";
      return `
        <tr>
          <td>${escapeHtml(row["产品编码"])}</td>
          <td>${escapeHtml(row["产品名称"])}</td>
          <td>${escapeHtml(row["产品状态"])}</td>
          <td>${escapeHtml(row["创建时间"])}</td>
          <td>${escapeHtml(row["部门"])}</td>
          <td class="num">${formatNumber(row["产品收入"])}</td>
          <td class="num">${formatNumber(row["产品毛利"])}</td>
          <td><span class="badge ${badgeClass}">${escapeHtml(isMissingApproval ? "无法判断" : row["退市类型"])}</span></td>
          <td>${escapeHtml(isMissingApproval ? row["问题"] : formatReason(row["理由"]))}</td>
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
    missingApprovalRows = [];
    showingMissingApproval = false;
    setDownloadEnabled(false);
    renderSummary();
    renderRows();
    return;
  }
  metadata = data.metadata;
  setDownloadEnabled(true);
  clearUploadQueue();
  renderSummary();
  await loadResult();
}

async function loadResult() {
  const response = await fetch("/api/result");
  if (!response.ok) return;
  const data = await response.json();
  metadata = data.metadata;
  allRows = data.rows || [];
  missingApprovalRows = data.missing_rule4_approval_rows || [];
  if (showingMissingApproval && !missingApprovalRows.length) {
    showingMissingApproval = false;
  }
  renderSummary();
  renderRows();
}

uploadForm.addEventListener("submit", (event) => event.preventDefault());

confirmAnalysis.addEventListener("click", async () => {
  if (!pendingInspection || confirmAnalysis.classList.contains("disabled")) return;
  setUploadOpen(false);
  setListLoading(true);
  showToast("正在按预检结果分析，请稍候。");
  setDownloadEnabled(false);
  try {
    const response = await fetch("/api/analyze-inspection", { method: "POST" });
    const data = await response.json();
    if (!response.ok) {
      showToast(data.error || "分析失败。");
      return;
    }
    metadata = data.metadata;
    missingApprovalRows = data.missing_rule4_approval_rows || [];
    showingMissingApproval = false;
    showToast(`已找到${metadata.candidate_count} 条结果`);
    setDownloadEnabled(true);
    await loadResult();
  } catch (error) {
    showToast(error.message || "分析失败。");
  } finally {
    setListLoading(false);
  }
});

searchInput.addEventListener("input", renderRows);
function toggleMissingApprovalRows() {
  if (missingApprovalNotice.classList.contains("disabled")) return;
  showingMissingApproval = !showingMissingApproval;
  updateMissingApprovalToggle();
  renderRows();
}

missingApprovalNotice.addEventListener("click", toggleMissingApprovalRows);
missingApprovalNotice.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  toggleMissingApprovalRows();
});
analysisFiles.addEventListener("change", () => {
  const changed = addSelectedFiles(analysisFiles.files);
  analysisFiles.value = "";
  if (changed) {
    setUploadStep(1);
    renderSelectedFiles();
    resetInspectionState();
    setInspectionLoading(false);
  }
});
inspectUpload.addEventListener("click", () => {
  if (inspectUpload.classList.contains("disabled")) return;
  inspectSelectedFiles();
});
selectedFileList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-file-index]");
  if (!button) return;
  removeSelectedFile(Number(button.dataset.fileIndex));
});
openUpload.addEventListener("click", () => setUploadOpen(true));
closeUpload.addEventListener("click", () => setUploadOpen(false));
cancelUpload.addEventListener("click", () => setUploadOpen(false));
backToFiles.addEventListener("click", () => setUploadStep(1));
uploadModal.addEventListener("click", (event) => {
  if (event.target?.hasAttribute("data-close-upload")) {
    setUploadOpen(false);
  }
});

function setChatOpen(isOpen) {
  analysisArea.classList.toggle("chat-open", isOpen);
  chatDrawer.setAttribute("aria-hidden", String(!isOpen));
  openChat.setAttribute("aria-expanded", String(isOpen));
  if (isOpen) {
    window.setTimeout(() => questionInput.focus(), 180);
  } else {
    openChat.focus();
  }
}

openChat.addEventListener("click", () => setChatOpen(true));
closeChat.addEventListener("click", () => setChatOpen(false));
chatForm.addEventListener("click", (event) => {
  if (event.target.closest("button")) return;
  questionInput.focus();
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  if (!uploadModal.hidden) {
    setUploadOpen(false);
    return;
  }
  if (analysisArea.classList.contains("chat-open")) {
    setChatOpen(false);
  }
});

function addMessage(className, text, options = {}) {
  const node = document.createElement("div");
  node.className = className;
  if (options.markdown) {
    renderMarkdown(node, text);
  } else {
    node.textContent = text;
  }
  chatLog.appendChild(node);
  chatLog.scrollTop = chatLog.scrollHeight;
  return node;
}

function clearChatHint() {
  chatLog.querySelector(".chat-hint")?.remove();
}

function createChatEmptyState() {
  const state = document.createElement("div");
  state.className = "chat-empty chat-hint";

  const mascot = document.createElement("img");
  mascot.className = "chat-empty-mascot";
  mascot.src = "/assets/chat-assistant-mascot-v2.png";
  mascot.alt = "";

  const copy = document.createElement("div");
  copy.className = "chat-empty-copy";
  const title = document.createElement("h3");
  title.textContent = "产品慧诊助手";
  const description = document.createElement("p");
  description.textContent = "上传分析后，我可以帮你快速复盘结果。";
  copy.append(title, description);

  const suggestions = document.createElement("div");
  suggestions.className = "chat-suggestions";
  suggestions.setAttribute("aria-label", "推荐问题");
  chatSuggestionQuestions.forEach((question) => {
    const button = document.createElement("button");
    button.className = "chat-suggestion";
    button.type = "button";
    button.dataset.question = question;
    button.textContent = question;
    suggestions.appendChild(button);
  });

  state.append(mascot, copy, suggestions);
  return state;
}

function resetChatWindow() {
  chatLog.replaceChildren(createChatEmptyState());
  questionInput.value = "";
  chatLog.scrollTop = 0;
  questionInput.focus();
}

newChat.addEventListener("click", resetChatWindow);

async function submitQuestion(question) {
  if (!question) return;
  clearChatHint();
  addMessage("user-msg", question);
  questionInput.value = "";
  const pending = addMessage("assistant-msg", "正在分析...");
  try {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question }),
    });
    const data = await response.json();
    if (response.ok) {
      renderMarkdown(pending, data.answer);
    } else {
      pending.textContent = data.error || "问答失败。";
    }
  } catch {
    pending.textContent = "问答服务连接失败，请确认本地服务正在运行后重试。";
  }
}

chatLog.addEventListener("click", (event) => {
  const suggestion = event.target.closest(".chat-suggestion");
  if (!suggestion) return;
  submitQuestion(suggestion.dataset.question || suggestion.textContent.trim());
});

chatForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const question = questionInput.value.trim();
  await submitQuestion(question);
});

loadStatus().catch((error) => {
  showToast(error.message || "状态读取失败。");
});
