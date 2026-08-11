function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatValue(value) {
  if (Array.isArray(value)) return value.join("、");
  if (value && typeof value === "object") return JSON.stringify(value);
  return value;
}

function statusLabel(value) {
  return {
    queued: "排队中",
    pending: "等待中",
    running: "运行中",
    success: "成功",
    failed: "失败"
  }[value] || value || "等待中";
}

function downloadLabel(name) {
  return {
    "surface_water_section.docx": "下载地表水章节",
    "noise_section.docx": "下载声环境章节",
    "eia_outputs.zip": "下载完整成果包",
    "现状调查与评价.docx": "下载合并章节",
    "project_area_overview.docx": "下载区域环境概况"
  }[name] || name;
}

function isSurfaceWaterFactorColumn(header) {
  const excluded = new Set(["编号", "监测时间", "point_code", "needs_review", "warning"]);
  return !excluded.has(String(header));
}

function isExceedCell(header, value, tableType) {
  const normalizedHeader = String(header).toLowerCase();
  const number = Number(value);
  if (!Number.isFinite(number)) return false;
  if (tableType === "noise_compliance" && normalizedHeader.includes("exceed")) {
    return number > 0;
  }
  if (tableType === "surface_water_compliance" && isSurfaceWaterFactorColumn(header)) {
    return number > 1;
  }
  return false;
}

function renderTable(containerId, table, tableType, jobId) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = renderTableMarkup(table || {}, tableType, jobId);
}

function renderTableGroup(containerId, table, tableType, jobId) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const subtables = Array.isArray(table?.subtables)
    ? table.subtables.filter((item) => item && item.exists !== false && Array.isArray(item.rows) && item.rows.length)
    : [];
  if (!subtables.length) {
    container.innerHTML = renderTableMarkup(table || {}, tableType, jobId);
    return;
  }
  container.innerHTML = subtables.map((subtable) => renderTableMarkup(subtable, tableType, jobId)).join("");
}

function renderTableMarkup(table, tableType, jobId) {
  const headers = table.headers || [];
  const rows = table.rows || [];
  const previewRows = rows.slice(0, 50);
  const sourcePath = table.source_path || "";
  const title = table.title || "暂无数据";

  if (!table.exists || !rows.length) {
    return `<div class="empty-state" style="padding:32px;"><div class="empty-state-icon">—</div><p>暂无可展示数据</p></div>`;
  }

  const downloadHref = `/api/jobs/${jobId}/download/${encodeURI(sourcePath)}`;
  const note = rows.length > 50 ? `仅预览前 50 行，共 ${rows.length} 行；可下载完整 JSON。` : `共 ${rows.length} 行。`;
  const head = headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("");
  const body = previewRows.map((row) => {
    const review = row.needs_review === true || Boolean(row.warning);
    const cells = headers.map((header) => {
      const value = row[header] ?? "";
      const exceed = isExceedCell(header, value, tableType);
      return `<td class="${exceed ? "exceed-cell" : ""}">${escapeHtml(formatValue(value))}</td>`;
    }).join("");
    return `<tr class="${review ? "review-row" : ""}">${cells}</tr>`;
  }).join("");

  return `
    <div class="table-toolbar">
      <div>
        <div class="table-title">${escapeHtml(title)}</div>
        <div class="table-note">${escapeHtml(note)}</div>
      </div>
      <a class="download-button secondary" href="${downloadHref}" target="_blank">下载 JSON</a>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr>${head}</tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;
}

function renderNoiseSummary(summary) {
  const container = document.getElementById("noiseSummary");
  if (!container) return;
  const sensitive = summary.sensitive || {};
  const attenuation = summary.attenuation || {};
  const cards = [
    ["敏感点总数", sensitive.total_count ?? "-", false],
    ["敏感点超标", sensitive.exceed_count ?? "-", Number(sensitive.exceed_count) > 0],
    ["衰减断面总数", attenuation.total_count ?? "-", false],
    ["衰减断面超标", attenuation.exceed_count ?? "-", Number(attenuation.exceed_count) > 0]
  ];
  container.innerHTML = cards.map(([label, value, danger]) => `
    <div class="summary-card">
      <span>${escapeHtml(label)}</span>
      <strong class="${danger ? "danger" : ""}">${escapeHtml(String(value))}</strong>
    </div>
  `).join("");
}

function renderTextDownloads(groups, jobId) {
  const container = document.getElementById("textDownloads");
  if (!container) return;
  const orderedGroups = ["text_output", "monitoring_extraction", "compliance"];
  const sections = orderedGroups
    .map((key) => ({ key, title: groups[key]?.title || key, files: groups[key]?.files || [] }))
    .filter((group) => group.files.length);
  container.innerHTML = "";
  if (!sections.length) {
    container.innerHTML = `<div class="empty-state" style="padding:24px;width:100%;"><p>暂无文本输出文件</p></div>`;
    return;
  }
  sections.forEach((section) => {
    const heading = document.createElement("div");
    heading.className = "download-group-title";
    heading.textContent = section.title;
    container.appendChild(heading);
    section.files.forEach((file) => {
      const link = document.createElement("a");
      link.className = "download-button";
      link.href = `/api/jobs/${jobId}/download/${encodeURI(file.path)}`;
      link.textContent = file.label || downloadLabel(file.name);
      link.target = "_blank";
      container.appendChild(link);
    });
  });
}

function renderLlmTextPolishStatus(llmStatus, targetId) {
  const target = document.getElementById(targetId || "llmTextPolishText");
  if (!target) return;
  const state = llmStatus.state || "unknown";
  const warnings = Array.isArray(llmStatus.warnings) ? llmStatus.warnings.filter(Boolean) : [];
  const detail =
    state === "fallback" || state === "success" || !warnings.length ? "" : `：${warnings.join("；")}`;
  target.className = `llm-state ${state}`;
  target.textContent = `${llmStatus.label || "暂无结果"}${detail}`;
}

function renderSchemaFallbackStatus(schemaStatus, targetId) {
  const target = document.getElementById(targetId || "schemaFallbackText");
  if (!target) return;
  const state = schemaStatus.state || "pending";
  const reasons = schemaReviewReasons(schemaStatus);
  const preview = reasons.slice(0, 3).join("；");
  const more = reasons.length > 3 ? `；另有${reasons.length - 3}条` : "";
  target.className = `llm-state ${state}`;
  target.textContent = `${schemaStatus.label || "等待结构解析"}${preview ? `：${preview}${more}` : ""}`;
  target.title = reasons.join("\n");
  renderSchemaReviewBanner(reasons, state);
}

function schemaReviewReasons(schemaStatus) {
  if (Array.isArray(schemaStatus.review_reasons) && schemaStatus.review_reasons.length) {
    return [...new Set(schemaStatus.review_reasons.filter(Boolean).map(String))];
  }
  const details = Array.isArray(schemaStatus.details) ? schemaStatus.details : [];
  const reasons = [];
  details.forEach((item) => {
    if (!item || typeof item !== "object") return;
    if (Array.isArray(item.warnings)) {
      item.warnings.forEach((warning) => {
        if (warning) reasons.push(String(warning));
      });
    }
    if (Array.isArray(item.missing_after_llm) && item.missing_after_llm.length) {
      reasons.push(`${item.schema || "结构解析"} 缺少字段: ${item.missing_after_llm.join("、")}`);
    }
    if (item.llm_error) {
      reasons.push(`${item.schema || "结构解析"} LLM结构解析失败: ${item.llm_error}`);
    }
  });
  return [...new Set(reasons)];
}

function renderSchemaReviewBanner(reasons, state) {
  const banner = document.getElementById("schemaReviewBanner");
  const list = document.getElementById("schemaReviewReasons");
  if (!banner || !list) return;
  if (state !== "needs_review" || !reasons.length) {
    banner.style.display = "none";
    list.innerHTML = "";
    return;
  }
  list.innerHTML = reasons.map((reason) => `<li>${escapeHtml(reason)}</li>`).join("");
  banner.style.display = "";
}

function renderInputValidationStatus(validation, jobId) {
  const banner = document.getElementById("inputValidationBanner");
  const list = document.getElementById("inputValidationReasons");
  const downloads = document.getElementById("inputValidationDownloads");
  if (!banner || !list || !downloads) return;
  const errors = Array.isArray(validation.errors) ? validation.errors.filter(Boolean) : [];
  if (validation.state !== "failed" || !errors.length) {
    banner.style.display = "none";
    list.innerHTML = "";
    downloads.innerHTML = "";
    return;
  }
  banner.style.display = "";
  const visibleErrors = errors.slice(0, 20);
  if (errors.length > visibleErrors.length) {
    visibleErrors.push(`另有 ${errors.length - visibleErrors.length} 项，详见下载的校核结果`);
  }
  const files = [
    [validation.report_path, "下载输入校核结果"],
    [validation.correspondence_path, "下载点位对应清单"]
  ].filter(([path]) => path);
  downloads.innerHTML = files.map(([path, label]) => `
    <a class="download-button secondary"
      href="/api/jobs/${encodeURIComponent(jobId)}/download/${encodeURI(path)}"
      target="_blank">${escapeHtml(label)}</a>
  `).join("");
}

function bindTabButtons() {
  document.addEventListener("click", (event) => {
    const button = event.target.closest(".tab-button");
    if (!button) return;
    const group = button.parentElement;
    const section = group.parentElement;
    group.querySelectorAll(".tab-button").forEach((item) => item.classList.remove("active"));
    section.querySelectorAll(".table-card").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    document.getElementById(button.dataset.target).classList.add("active");
  });
}

function bindFileInput(inputId, labelId) {
  const input = document.getElementById(inputId);
  const label = document.getElementById(labelId);
  if (!input || !label) return;
  input.addEventListener("change", () => {
    label.textContent = input.files[0]?.name || "选择 .docx 文件";
  });
}
