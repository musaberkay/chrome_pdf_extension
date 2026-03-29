import { PDFDocument } from "pdf-lib";
import { configurePdfWorker, convertPdfsToDocx } from "./pdf-to-word.js";

configurePdfWorker();

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileList = document.getElementById("fileList");
const mergeBtn = document.getElementById("mergeBtn");
const wordBtn = document.getElementById("wordBtn");
const clearBtn = document.getElementById("clearBtn");
const statusEl = document.getElementById("status");

const viewHome = document.getElementById("viewHome");
const viewWorkspace = document.getElementById("viewWorkspace");
const backBtn = document.getElementById("backBtn");
const workspaceTitle = document.getElementById("workspaceTitle");
const workspaceDesc = document.getElementById("workspaceDesc");

const WORKSPACE = {
  merge: {
    title: "Merge PDF",
    desc: "Combine PDFs in the order you want. Drag to reorder, then merge into one file.",
  },
  word: {
    title: "PDF to Word",
    desc: "Convert PDF text to an editable .docx file. Layout, images, and complex tables are not preserved.",
  },
};

/** @type {File[]} */
let files = [];

/** @type {number | null} */
let dragSourceIndex = null;

function formatKb(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  return `${(bytes / 1024).toFixed(1)} KB`;
}

function setStatus(text, ok = false) {
  statusEl.textContent = text;
  statusEl.classList.toggle("status--ok", ok);
}

function updateButtons() {
  const hasFiles = files.length > 0;
  mergeBtn.disabled = !hasFiles;
  wordBtn.disabled = !hasFiles;
  clearBtn.disabled = !hasFiles;
}

function clearDropHighlights() {
  fileList.querySelectorAll(".file-item--over").forEach((el) => {
    el.classList.remove("file-item--over");
  });
}

function hasFilePayload(dataTransfer) {
  return dataTransfer?.types?.includes?.("Files") === true;
}

/**
 * @param {File[]} picked
 */
function addPdfFiles(picked) {
  const pdfs = picked.filter(
    (f) => f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf")
  );
  if (pdfs.length < picked.length) {
    setStatus("Some files were skipped (not PDF).");
  } else {
    setStatus("");
  }
  const seen = new Set(files.map((f) => `${f.name}-${f.size}-${f.lastModified}`));
  for (const f of pdfs) {
    const key = `${f.name}-${f.size}-${f.lastModified}`;
    if (!seen.has(key)) {
      seen.add(key);
      files.push(f);
    }
  }
  renderList();
}

function renderList() {
  fileList.innerHTML = "";
  files.forEach((file, index) => {
    const li = document.createElement("li");
    li.className = "file-item";
    li.draggable = true;
    li.title = "Drag to reorder";

    const grip = document.createElement("span");
    grip.className = "file-item__grip";
    grip.setAttribute("aria-hidden", "true");

    const nameEl = document.createElement("span");
    nameEl.className = "file-item__name";
    nameEl.textContent = file.name;
    nameEl.title = file.name;

    const meta = document.createElement("span");
    meta.className = "file-item__meta";
    meta.textContent = formatKb(file.size);

    li.append(grip, nameEl, meta);

    li.addEventListener("dragstart", (e) => {
      dragSourceIndex = index;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", String(index));
      li.classList.add("file-item--dragging");
    });

    li.addEventListener("dragend", () => {
      dragSourceIndex = null;
      li.classList.remove("file-item--dragging");
      clearDropHighlights();
    });

    li.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      if (dragSourceIndex !== null && dragSourceIndex !== index) {
        li.classList.add("file-item--over");
      }
    });

    li.addEventListener("dragleave", (e) => {
      const rel = e.relatedTarget;
      if (rel instanceof Node && li.contains(rel)) return;
      li.classList.remove("file-item--over");
    });

    li.addEventListener("drop", (e) => {
      e.preventDefault();
      li.classList.remove("file-item--over");
      const fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
      if (!Number.isFinite(fromIdx) || fromIdx === index) return;
      const [moved] = files.splice(fromIdx, 1);
      files.splice(index, 0, moved);
      renderList();
    });

    fileList.appendChild(li);
  });
  updateButtons();
}

fileInput.addEventListener("change", () => {
  addPdfFiles(Array.from(fileInput.files || []));
  fileInput.value = "";
});

dropZone.addEventListener("dragenter", (e) => {
  if (!hasFilePayload(e.dataTransfer)) return;
  e.preventDefault();
  dropZone.classList.add("drop-zone--active");
});

dropZone.addEventListener("dragleave", (e) => {
  if (!hasFilePayload(e.dataTransfer)) return;
  const rel = e.relatedTarget;
  if (rel instanceof Node && dropZone.contains(rel)) return;
  dropZone.classList.remove("drop-zone--active");
});

dropZone.addEventListener("dragover", (e) => {
  if (!hasFilePayload(e.dataTransfer)) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = "copy";
});

dropZone.addEventListener("drop", (e) => {
  if (!hasFilePayload(e.dataTransfer)) return;
  e.preventDefault();
  dropZone.classList.remove("drop-zone--active");
  const list = e.dataTransfer?.files;
  if (!list?.length) return;
  addPdfFiles(Array.from(list));
});

clearBtn.addEventListener("click", () => {
  files = [];
  setStatus("");
  renderList();
});

mergeBtn.addEventListener("click", async () => {
  if (files.length === 0) return;
  setStatus("Merging…");
  mergeBtn.disabled = true;
  wordBtn.disabled = true;
  clearBtn.disabled = true;
  try {
    const merged = await PDFDocument.create();
    for (const file of files) {
      const bytes = await file.arrayBuffer();
      const doc = await PDFDocument.load(bytes, { ignoreEncryption: false });
      const indices = doc.getPageIndices();
      const pages = await merged.copyPages(doc, indices);
      pages.forEach((p) => merged.addPage(p));
    }
    const out = await merged.save();
    const blob = new Blob([out], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merged.pdf";
    a.click();
    URL.revokeObjectURL(url);
    setStatus("Done. Check your downloads.", true);
  } catch (e) {
    const msg =
      e instanceof Error ? e.message : "Merge failed.";
    setStatus(
      msg.includes("encrypt")
        ? "A PDF is password-protected or encrypted."
        : `Error: ${msg}`
    );
  } finally {
    updateButtons();
  }
});

wordBtn.addEventListener("click", async () => {
  if (files.length === 0) return;
  setStatus("Converting to Word…");
  mergeBtn.disabled = true;
  wordBtn.disabled = true;
  clearBtn.disabled = true;
  try {
    const blob = await convertPdfsToDocx(files);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "converted.docx";
    a.click();
    URL.revokeObjectURL(url);
    setStatus("Word document downloaded. Text-only; layout may differ.", true);
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Conversion failed.";
    setStatus(
      msg.includes("encrypt") || msg.includes("password")
        ? "Password-protected PDFs are not supported for conversion."
        : `Error: ${msg}`
    );
  } finally {
    updateButtons();
  }
});

function showHome() {
  if (!viewHome || !viewWorkspace) return;
  viewHome.hidden = false;
  viewWorkspace.hidden = true;
  setStatus("", false);
}

function showWorkspace(tool) {
  if (!viewWorkspace || !workspaceTitle || !workspaceDesc) return;
  const cfg = WORKSPACE[tool];
  if (!cfg) return;
  if (viewHome) viewHome.hidden = true;
  viewWorkspace.hidden = false;
  workspaceTitle.textContent = cfg.title;
  workspaceDesc.textContent = cfg.desc;
  mergeBtn.hidden = tool !== "merge";
  wordBtn.hidden = tool !== "word";
  updateButtons();
}

if (viewHome) {
  viewHome.querySelectorAll("[data-tool]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tool = btn.dataset.tool;
      if (tool === "merge" || tool === "word") showWorkspace(tool);
    });
  });
}

if (backBtn) backBtn.addEventListener("click", showHome);

updateButtons();
