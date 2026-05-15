/* ─────────────────────────────────────────────────────────
   Правовая экспертиза договора — frontend logic
   ───────────────────────────────────────────────────────── */

const state = {
  sessionId: null,
  analysis: null,
  file: null,
};

/* ── DOM refs ── */
const uploadSection   = document.getElementById("upload-section");
const loadingSection  = document.getElementById("loading-section");
const resultsSection  = document.getElementById("results-section");
const actionsSection  = document.getElementById("actions-section");
const dropZone        = document.getElementById("drop-zone");
const fileInput       = document.getElementById("file-input");
const filePill        = document.getElementById("file-pill");
const filePillName    = document.getElementById("file-pill-name");
const fileRemoveBtn   = document.getElementById("file-remove");
const checkBtn        = document.getElementById("check-btn");
const toast           = document.getElementById("toast");
const modalBackdrop   = document.getElementById("modal-backdrop");

/* ─── Upload wiring ─── */

dropZone.addEventListener("click", () => fileInput.click());

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const f = e.dataTransfer.files[0];
  if (f) setFile(f);
});

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

fileRemoveBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  clearFile();
});

function setFile(f) {
  const ext = f.name.split(".").pop().toLowerCase();
  if (!["docx", "txt"].includes(ext)) {
    showToast("Поддерживаются форматы: .docx, .txt", true);
    return;
  }
  state.file = f;
  filePillName.textContent = f.name;
  filePill.classList.add("visible");
  checkBtn.disabled = false;
}

function clearFile() {
  state.file = null;
  fileInput.value = "";
  filePill.classList.remove("visible");
  checkBtn.disabled = true;
}

/* ─── Check button ─── */

checkBtn.addEventListener("click", startAnalysis);

async function startAnalysis() {
  if (!state.file) return;

  show(loadingSection);
  hide(uploadSection);
  hide(resultsSection);
  hide(actionsSection);

  const fd = new FormData();
  fd.append("contract", state.file);

  try {
    const res = await fetch("/api/analyze", { method: "POST", body: fd });
    const data = await res.json();

    if (!res.ok || data.error) {
      showToast(data.error || "Ошибка сервера", true);
      show(uploadSection);
      hide(loadingSection);
      return;
    }

    state.sessionId = data.session_id;
    state.analysis  = data.analysis;
    renderResults(data.analysis);

    hide(loadingSection);
    show(resultsSection);
    show(actionsSection);
    resultsSection.scrollIntoView({ behavior: "smooth", block: "start" });

  } catch (err) {
    showToast("Ошибка соединения с сервером", true);
    show(uploadSection);
    hide(loadingSection);
  }
}

/* ─── Render results ─── */

function renderResults(a) {
  const missing  = a.missing_clauses  || [];
  const modified = a.modified_clauses || [];
  const compliant = a.compliant_clauses || [];
  const hasIssues = missing.length > 0 || modified.length > 0;

  // Status banner
  const banner = document.getElementById("status-banner");
  if (hasIssues) {
    banner.className = "status-banner has-issues";
    banner.innerHTML = `<span class="status-dot"></span>
      ВЫЯВЛЕНЫ ЗАМЕЧАНИЯ: необходима доработка договора`;
  } else {
    banner.className = "status-banner no-issues";
    banner.innerHTML = `<span class="status-dot"></span>
      ЗАМЕЧАНИЙ НЕТ: договор соответствует требованиям Заказчика`;
  }

  // Requisites
  const req = a.contract_requisites || {};
  document.getElementById("req-number").textContent  = req.number  || "не указан";
  document.getElementById("req-date").textContent    = req.date    || "не указана";
  document.getElementById("req-supplier").textContent = req.supplier || "не определён";
  document.getElementById("req-buyer").textContent   = req.buyer   || "не определён";
  document.getElementById("req-subject").textContent = req.subject || "не указан";

  // Risk badge
  const riskEl = document.getElementById("risk-badge");
  const risk = (a.risk_level || "").toLowerCase();
  riskEl.textContent = a.risk_level || "—";
  riskEl.className = "risk-badge " + (risk === "высокий" ? "high" : risk === "средний" ? "medium" : "low");

  // Assessment
  document.getElementById("assessment-text").textContent = a.overall_assessment || "";

  // Counters
  document.getElementById("cnt-missing").textContent  = missing.length;
  document.getElementById("cnt-modified").textContent = modified.length;
  document.getElementById("cnt-compliant").textContent = compliant.length;

  // Detail lists
  renderGroup("group-missing",  missing,  "missing");
  renderGroup("group-modified", modified, "modified");
  renderGroup("group-compliant", compliant, "compliant");

  // Show/hide groups
  toggle("group-missing-wrap",  missing.length > 0);
  toggle("group-modified-wrap", modified.length > 0);
  toggle("group-compliant-wrap", compliant.length > 0);

  // Buttons
  const sid = state.sessionId;
  const isDocx = state.file && state.file.name.toLowerCase().endsWith(".docx");

  const btnChecked = document.getElementById("btn-checked");
  btnChecked.disabled = !isDocx || !hasIssues;
  btnChecked.onclick = () => triggerDownload(`/api/download/checked/${sid}`);

  document.getElementById("btn-conclusion").onclick = () =>
    triggerDownload(`/api/download/conclusion/${sid}`);

  // Protocol vs agreement logic
  const btnLegal = document.getElementById("btn-legal");
  if (!hasIssues) {
    btnLegal.disabled = true;
  } else {
    btnLegal.disabled = false;
    btnLegal.onclick = () => openModal(a);
  }
}

function renderGroup(containerId, clauses, type) {
  const container = document.getElementById(containerId);
  container.innerHTML = "";
  clauses.forEach(cl => {
    const item = document.createElement("div");
    item.className = `detail-item ${type}`;

    let inner = `<div class="detail-item-head">
      ${cl.requirement_name || ""}
      <span class="detail-item-id">${cl.requirement_id || ""}</span>
    </div>`;

    if (type === "missing") {
      inner += `<div class="detail-item-body">${cl.description || ""}</div>`;
      if (cl.recommended_wording) {
        inner += `<div class="detail-item-rec">Рекомендуемая формулировка: ${cl.recommended_wording}</div>`;
      }
    } else if (type === "modified") {
      if (cl.contract_wording) {
        inner += `<div class="detail-item-cite">В договоре: «${cl.contract_wording}»</div>`;
      }
      inner += `<div class="detail-item-body">${cl.issue || ""}</div>`;
      if (cl.recommended_correction) {
        inner += `<div class="detail-item-rec">Рекомендация: ${cl.recommended_correction}</div>`;
      }
    } else {
      if (cl.contract_wording) {
        inner += `<div class="detail-item-cite">«${cl.contract_wording}»</div>`;
      }
    }

    item.innerHTML = inner;
    container.appendChild(item);
  });
}

/* ─── Group toggles ─── */

document.querySelectorAll(".group-toggle").forEach(btn => {
  btn.addEventListener("click", () => {
    const body = btn.nextElementSibling;
    const chevron = btn.querySelector(".chevron");
    body.classList.toggle("collapsed");
    chevron.style.transform = body.classList.contains("collapsed") ? "rotate(-90deg)" : "";
  });
});

/* ─── Modal ─── */

function openModal(a) {
  const missing  = (a.missing_clauses  || []).length;
  const modified = (a.modified_clauses || []).length;

  // Recommend protocol if there are modified clauses (need to change existing), else agreement
  const preferProtocol = modified > 0;

  const noteEl = document.getElementById("modal-note");
  noteEl.textContent = preferProtocol
    ? `В договоре выявлены ${modified} усл. с замечаниями и ${missing} отсутствующих условий. ` +
      "Рекомендуется Протокол разногласий."
    : `В договоре отсутствуют ${missing} обязательных условий (изменений редакции нет). ` +
      "Рекомендуется Дополнительное соглашение.";

  const cardProto = document.getElementById("card-protocol");
  const cardAgree = document.getElementById("card-agreement");
  const tagProto  = document.getElementById("rec-tag-protocol");
  const tagAgree  = document.getElementById("rec-tag-agreement");

  if (preferProtocol) {
    cardProto.classList.add("recommended");
    cardAgree.classList.remove("recommended");
    tagProto.classList.remove("hidden");
    tagAgree.classList.add("hidden");
  } else {
    cardAgree.classList.add("recommended");
    cardProto.classList.remove("recommended");
    tagAgree.classList.remove("hidden");
    tagProto.classList.add("hidden");
  }

  const sid = state.sessionId;
  document.getElementById("btn-download-protocol").onclick = () => {
    triggerDownload(`/api/download/protocol/${sid}`);
    closeModal();
  };
  document.getElementById("btn-download-agreement").onclick = () => {
    triggerDownload(`/api/download/agreement/${sid}`);
    closeModal();
  };

  modalBackdrop.classList.add("open");
}

function closeModal() {
  modalBackdrop.classList.remove("open");
}

document.getElementById("modal-close-btn").addEventListener("click", closeModal);
modalBackdrop.addEventListener("click", (e) => {
  if (e.target === modalBackdrop) closeModal();
});

/* ─── Download helper ─── */

async function triggerDownload(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      showToast(data.error || "Ошибка при генерации документа", true);
      return;
    }
    const blob = await res.blob();
    const disposition = res.headers.get("content-disposition") || "";
    let filename = "document.docx";
    const m = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/i);
    if (m) filename = decodeURIComponent(m[1].replace(/['"]/g, ""));

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
    showToast("Документ загружен");
  } catch {
    showToast("Ошибка при скачивании файла", true);
  }
}

/* ─── Toast ─── */

let toastTimer;
function showToast(msg, isError = false) {
  toast.textContent = msg;
  toast.className = "toast" + (isError ? " error" : "");
  toast.classList.add("visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("visible"), 3500);
}

/* ─── Utilities ─── */

function show(el) { el.classList.remove("hidden"); }
function hide(el) { el.classList.add("hidden"); }
function toggle(id, visible) {
  const el = document.getElementById(id);
  if (el) visible ? show(el) : hide(el);
}
