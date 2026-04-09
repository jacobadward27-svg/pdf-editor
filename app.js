/* global pdfjsLib, PDFLib — loaded by UMD scripts in index.html */
const PDFDocument =
  typeof PDFLib !== "undefined"
    ? PDFLib.PDFDocument || (PDFLib.default && PDFLib.default.PDFDocument)
    : null;

const THUMB_SCALE = 0.22;
const EXPORT_SCALE = 2;
const MAX_UNDO = 45;
const DB_NAME = "docsfix-editor";
const IDB_STORE = "kv";

/** 24 distinct highlighter colors for selection fills */
const HIGHLIGHT_PALETTE = [
  "#FFEB3B",
  "#FFC107",
  "#FF9800",
  "#FF5722",
  "#F44336",
  "#E91E63",
  "#9C27B0",
  "#673AB7",
  "#3F51B5",
  "#2196F3",
  "#03A9F4",
  "#00BCD4",
  "#009688",
  "#4CAF50",
  "#8BC34A",
  "#CDDC39",
  "#FFEE58",
  "#FFD54F",
  "#FF8A65",
  "#A1887F",
  "#90A4AE",
  "#78909C",
  "#B39DDB",
  "#80DEEA",
];

const els = {
  filePicker: document.getElementById("filePicker"),
  pdfLoading: document.getElementById("pdfLoading"),
  thumbnails: document.getElementById("thumbnails"),
  pdfCanvas: document.getElementById("pdfCanvas"),
  overlayCanvas: document.getElementById("overlayCanvas"),
  selectionHud: document.getElementById("selectionHud"),
  pageCanvasInner: document.getElementById("pageCanvasInner"),
  selectionToolbar: document.getElementById("selectionToolbar"),
  stColors: document.getElementById("stColors"),
  emptyState: document.getElementById("emptyState"),
  pageStack: document.getElementById("pageStack"),
  docTitle: document.getElementById("docTitle"),
  statusText: document.getElementById("statusText"),
  zoomLabel: document.getElementById("zoomLabel"),
  zoomIn: document.getElementById("zoomIn"),
  zoomOut: document.getElementById("zoomOut"),
  btnExport: document.getElementById("btnExport"),
  btnSave: document.getElementById("btnSave"),
  btnUndo: document.getElementById("btnUndo"),
  btnRedo: document.getElementById("btnRedo"),
  btnAnotherPdf: document.getElementById("btnAnotherPdf"),
  btnDuplicate: document.getElementById("btnDuplicate"),
  btnInfo: document.getElementById("btnInfo"),
  infoDialog: document.getElementById("infoDialog"),
  infoBody: document.getElementById("infoBody"),
  searchInput: document.getElementById("searchInput"),
  toast: document.getElementById("toast"),
  stCut: document.getElementById("stCut"),
  stCopy: document.getElementById("stCopy"),
  stDelete: document.getElementById("stDelete"),
  stRewrite: document.getElementById("stRewrite"),
  stShapeRect: document.getElementById("stShapeRect"),
  stShapeCircle: document.getElementById("stShapeCircle"),
  stShapeArrow: document.getElementById("stShapeArrow"),
  searchPanel: document.getElementById("searchPanel"),
  searchResults: document.getElementById("searchResults"),
  searchPanelClose: document.getElementById("searchPanelClose"),
  searchPanelTitle: document.getElementById("searchPanelTitle"),
  mergePicker: document.getElementById("mergePicker"),
  imagePicker: document.getElementById("imagePicker"),
  btnAddPage: document.getElementById("btnAddPage"),
  btnDelPage: document.getElementById("btnDelPage"),
  btnMergePdf: document.getElementById("btnMergePdf"),
  btnSplitPdf: document.getElementById("btnSplitPdf"),
  btnReplaceText: document.getElementById("btnReplaceText"),
  btnSignature: document.getElementById("btnSignature"),
  btnInsertImage: document.getElementById("btnInsertImage"),
  signatureDialog: document.getElementById("signatureDialog"),
  signaturePad: document.getElementById("signaturePad"),
  sigClear: document.getElementById("sigClear"),
  signatureForm: document.getElementById("signatureForm"),
};

let pdfDoc = null;
let pdfBytes = null;
let fileName = "document.pdf";
let numPages = 0;
let currentPage = 1;
let viewScale = 1;
let activeTool = "draw";
/** @type {number[]} */
let exportOrder = [];
/** @type {Record<number, string>} */
let overlayStorage = {};
/** @type {Record<number, string[]>} */
const pageUndo = {};
/** @type {Record<number, string[]>} */
const pageRedo = {};

let isDrawing = false;
let lastPoint = null;
let drawDidMove = false;
let drawPrimedUndo = false;

let isSelecting = false;
let selStart = null;
let selEnd = null;
/** @type {{ x: number; y: number; w: number; h: number } | null} */
let activeSelection = null;

let dbPromise = null;
let persistTimer = null;

/** @type {Record<number, { w: number; h: number }>} PDF page size at scale 1 (canonical overlay space) */
let pageBaseViewport = {};

/** Display order: PDF page numbers (1-based) in sidebar order */
let pageListOrder = [];

/** @type {'sig' | 'img' | null} */
let placementMode = null;
let placementDataUrl = null;
let placementImgEl = null;

function showToast(msg) {
  els.toast.textContent = msg;
  els.toast.classList.remove("hidden");
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => els.toast.classList.add("hidden"), 2800);
}

function setPdfLoading(on) {
  if (els.pdfLoading) els.pdfLoading.classList.toggle("hidden", !on);
}

function openDb() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onerror = () => reject(req.error);
    req.onsuccess = () => resolve(req.result);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(IDB_STORE)) {
        req.result.createObjectStore(IDB_STORE);
      }
    };
  });
  return dbPromise;
}

async function idbSet(key, val) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(val, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function idbGet(key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const r = tx.objectStore(IDB_STORE).get(key);
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

function schedulePersist() {
  if (!pdfDoc || !pdfBytes) return;
  clearTimeout(persistTimer);
  persistTimer = setTimeout(() => saveSession(), 450);
}

async function saveSession() {
  if (!pdfBytes || !pdfDoc) return;
  stashOverlay();
  try {
    const copy = new Uint8Array(pdfBytes);
    await idbSet("pdfBytes", copy);
    await idbSet("fileName", fileName);
    await idbSet("exportOrder", [...exportOrder]);
    await idbSet("overlayStorage", { ...overlayStorage });
    await idbSet("currentPage", currentPage);
    await idbSet("viewScale", viewScale);
    await idbSet("savedAt", Date.now());
  } catch (e) {
    console.warn("Could not save session", e);
  }
}

function clearUndoRedoAll() {
  for (const k of Object.keys(pageUndo)) delete pageUndo[k];
  for (const k of Object.keys(pageRedo)) delete pageRedo[k];
  updateUndoButtons();
}

function getCanonicalOverlayDataUrl() {
  const base = pageBaseViewport[currentPage];
  if (!base || !els.overlayCanvas.width) {
    return els.overlayCanvas.toDataURL("image/png");
  }
  const c = document.createElement("canvas");
  c.width = base.w;
  c.height = base.h;
  c.getContext("2d").drawImage(
    els.overlayCanvas,
    0,
    0,
    els.overlayCanvas.width,
    els.overlayCanvas.height,
    0,
    0,
    base.w,
    base.h
  );
  return c.toDataURL("image/png");
}

function beforeOverlayMutation() {
  if (!pdfDoc) return;
  const p = currentPage;
  const u = pageUndo[p] || (pageUndo[p] = []);
  u.push(getCanonicalOverlayDataUrl());
  if (u.length > MAX_UNDO) u.shift();
  pageRedo[p] = [];
  updateUndoButtons();
}

function updateUndoButtons() {
  const p = currentPage;
  const u = pageUndo[p];
  const r = pageRedo[p];
  if (els.btnUndo) els.btnUndo.disabled = !pdfDoc || !u || u.length === 0;
  if (els.btnRedo) els.btnRedo.disabled = !pdfDoc || !r || r.length === 0;
}

async function applySnap(url) {
  const ctx = els.overlayCanvas.getContext("2d");
  const w = els.overlayCanvas.width;
  const h = els.overlayCanvas.height;
  ctx.clearRect(0, 0, w, h);
  if (!url) return;
  await new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      ctx.drawImage(img, 0, 0, img.naturalWidth, img.naturalHeight, 0, 0, w, h);
      resolve();
    };
    img.onerror = reject;
    img.src = url;
  });
}

async function undo() {
  const p = currentPage;
  const u = pageUndo[p];
  if (!u || !u.length) return;
  const r = pageRedo[p] || (pageRedo[p] = []);
  r.push(getCanonicalOverlayDataUrl());
  const prev = u.pop();
  await applySnap(prev);
  overlayStorage[p] = prev;
  updateUndoButtons();
  schedulePersist();
  hideSelectionToolbar();
}

async function redo() {
  const p = currentPage;
  const r = pageRedo[p];
  if (!r || !r.length) return;
  const u = pageUndo[p] || (pageUndo[p] = []);
  u.push(getCanonicalOverlayDataUrl());
  const next = r.pop();
  await applySnap(next);
  overlayStorage[p] = next;
  updateUndoButtons();
  schedulePersist();
  hideSelectionToolbar();
}

function setTool(tool) {
  hideSelectionToolbar();
  clearSelectionHud();
  activeTool = tool;
  document.querySelectorAll(".tool-btn[data-tool]").forEach((b) => {
    b.classList.toggle("active", b.dataset.tool === tool);
  });
  const overlay = els.overlayCanvas;
  const interactive = tool === "draw" || tool === "highlight" || tool === "text" || tool === "select";
  overlay.classList.toggle("interactive", interactive);
  overlay.classList.toggle("cursor-select", tool === "select");
  if (tool === "draw" || tool === "highlight") {
    overlay.style.cursor = "crosshair";
  } else if (tool === "text") {
    overlay.style.cursor = "text";
  } else if (tool === "select") {
    overlay.style.cursor = "crosshair";
  } else {
    overlay.style.cursor = "default";
  }
}

function stashOverlay() {
  if (!pdfDoc || !numPages) return;
  if (els.overlayCanvas.width > 0 && els.overlayCanvas.height > 0) {
    overlayStorage[currentPage] = getCanonicalOverlayDataUrl();
  }
  schedulePersist();
}

function restoreOverlay(pageNum) {
  const ctx = els.overlayCanvas.getContext("2d");
  const url = overlayStorage[pageNum];
  const W = els.overlayCanvas.width;
  const H = els.overlayCanvas.height;
  ctx.clearRect(0, 0, W, H);
  if (!url) return;
  const img = new Image();
  img.onload = () => {
    ctx.clearRect(0, 0, W, H);
    ctx.drawImage(img, 0, 0, img.naturalWidth, img.naturalHeight, 0, 0, W, H);
  };
  img.src = url;
}

function syncOverlaySize() {
  const pdfCanvas = els.pdfCanvas;
  const overlay = els.overlayCanvas;
  const hud = els.selectionHud;
  overlay.width = pdfCanvas.width;
  overlay.height = pdfCanvas.height;
  overlay.style.width = pdfCanvas.style.width || `${pdfCanvas.width}px`;
  overlay.style.height = pdfCanvas.style.height || `${pdfCanvas.height}px`;
  if (hud) {
    hud.width = pdfCanvas.width;
    hud.height = pdfCanvas.height;
    hud.style.width = overlay.style.width;
    hud.style.height = overlay.style.height;
  }
}

function clearSelectionHud() {
  const h = els.selectionHud;
  if (!h || !h.width) return;
  h.getContext("2d").clearRect(0, 0, h.width, h.height);
}

function normalizeRect(x1, y1, x2, y2) {
  const x = Math.min(x1, x2);
  const y = Math.min(y1, y2);
  const w = Math.abs(x2 - x1);
  const h = Math.abs(y2 - y1);
  return { x, y, w, h };
}

function drawSelectionHud() {
  const h = els.selectionHud;
  if (!h || !selStart || !selEnd) return;
  const ctx = h.getContext("2d");
  ctx.clearRect(0, 0, h.width, h.height);
  const { x, y, w, h: hh } = normalizeRect(selStart.x, selStart.y, selEnd.x, selEnd.y);
  if (w < 1 || hh < 1) return;
  ctx.strokeStyle = "#2b7fff";
  ctx.lineWidth = 2;
  ctx.setLineDash([6, 4]);
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, hh - 1);
  ctx.setLineDash([]);
  ctx.fillStyle = "rgba(43, 127, 255, 0.12)";
  ctx.fillRect(x, y, w, hh);
}

function hideSelectionToolbar() {
  activeSelection = null;
  if (els.selectionToolbar) els.selectionToolbar.classList.add("hidden");
}

function positionSelectionToolbar(rect) {
  const tb = els.selectionToolbar;
  if (!tb) return;
  const pad = 8;
  const ob = els.overlayCanvas.getBoundingClientRect();
  const scaleX = ob.width / els.overlayCanvas.width;
  const scaleY = ob.height / els.overlayCanvas.height;
  let left = ob.left + rect.x * scaleX;
  let top = ob.top + (rect.y + rect.h) * scaleY + pad;
  tb.classList.remove("hidden");
  const tw = tb.offsetWidth || 260;
  const th = tb.offsetHeight || 120;
  if (left + tw > window.innerWidth - 12) left = window.innerWidth - tw - 12;
  if (left < 8) left = 8;
  if (top + th > window.innerHeight - 12) top = ob.top + rect.y * scaleY - th - pad;
  if (top < 8) top = 8;
  tb.style.left = `${left}px`;
  tb.style.top = `${top}px`;
}

function finalizeSelection() {
  if (!selStart || !selEnd) {
    isSelecting = false;
    clearSelectionHud();
    return;
  }
  const r = normalizeRect(selStart.x, selStart.y, selEnd.x, selEnd.y);
  isSelecting = false;
  selStart = null;
  selEnd = null;
  clearSelectionHud();
  if (r.w < 6 || r.h < 6) {
    hideSelectionToolbar();
    return;
  }
  activeSelection = r;
  positionSelectionToolbar(r);
}

function getOverlayPoint(e) {
  const overlay = els.overlayCanvas;
  const rect = overlay.getBoundingClientRect();
  const scaleX = overlay.width / rect.width;
  const scaleY = overlay.height / rect.height;
  return {
    x: (e.clientX - rect.left) * scaleX,
    y: (e.clientY - rect.top) * scaleY,
  };
}

function primeDrawUndo() {
  if (drawPrimedUndo) return;
  beforeOverlayMutation();
  drawPrimedUndo = true;
}

function startDraw(e) {
  if (!pdfDoc) return;
  if (activeTool === "select") {
    hideSelectionToolbar();
    isSelecting = true;
    selStart = getOverlayPoint(e);
    selEnd = { ...selStart };
    return;
  }
  if (activeTool !== "draw" && activeTool !== "highlight") return;
  primeDrawUndo();
  isDrawing = true;
  drawDidMove = false;
  lastPoint = getOverlayPoint(e);
}

function moveDraw(e) {
  if (activeTool === "select" && isSelecting && selStart) {
    selEnd = getOverlayPoint(e);
    drawSelectionHud();
    return;
  }
  if (!isDrawing || !lastPoint) return;
  const p = getOverlayPoint(e);
  if (Math.hypot(p.x - lastPoint.x, p.y - lastPoint.y) > 0.5) drawDidMove = true;
  const ctx = els.overlayCanvas.getContext("2d");
  ctx.beginPath();
  ctx.moveTo(lastPoint.x, lastPoint.y);
  ctx.lineTo(p.x, p.y);
  if (activeTool === "highlight") {
    ctx.strokeStyle = "rgba(255, 220, 0, 0.45)";
    ctx.lineWidth = 16;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
  } else {
    ctx.strokeStyle = "#1a2b4a";
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
  }
  ctx.stroke();
  lastPoint = p;
}

function endDraw() {
  if (activeTool === "select" && isSelecting) {
    finalizeSelection();
    return;
  }
  if (isDrawing) {
    if (!drawDidMove && lastPoint) {
      const ctx = els.overlayCanvas.getContext("2d");
      ctx.beginPath();
      ctx.arc(lastPoint.x, lastPoint.y, activeTool === "highlight" ? 8 : 1.2, 0, Math.PI * 2);
      ctx.fillStyle = activeTool === "highlight" ? "rgba(255, 220, 0, 0.45)" : "#1a2b4a";
      ctx.fill();
    }
    stashOverlay();
    schedulePersist();
  }
  isDrawing = false;
  lastPoint = null;
  drawDidMove = false;
  drawPrimedUndo = false;
}

function onTextClick(e) {
  if (activeTool !== "text" || !pdfDoc) return;
  beforeOverlayMutation();
  const p = getOverlayPoint(e);
  const label = window.prompt("Type the text to place on the page:", "");
  if (!label) {
    const u = pageUndo[currentPage];
    if (u && u.length) u.pop();
    updateUndoButtons();
    return;
  }
  const ctx = els.overlayCanvas.getContext("2d");
  ctx.font = "600 14px DM Sans, system-ui, sans-serif";
  ctx.fillStyle = "#1a2b4a";
  ctx.strokeStyle = "rgba(255,255,255,0.9)";
  ctx.lineWidth = 3;
  ctx.strokeText(label, p.x, p.y);
  ctx.fillText(label, p.x, p.y);
  stashOverlay();
  schedulePersist();
}

function getSelectionImageData(rect) {
  const c = document.createElement("canvas");
  c.width = Math.max(1, Math.floor(rect.w));
  c.height = Math.max(1, Math.floor(rect.h));
  const x = Math.floor(rect.x);
  const y = Math.floor(rect.y);
  c.getContext("2d").drawImage(els.overlayCanvas, x, y, rect.w, rect.h, 0, 0, c.width, c.height);
  return c;
}

async function copySelectionToClipboard(cutAfter) {
  const rect = activeSelection;
  if (!rect) return;
  const c = getSelectionImageData(rect);
  let copied = false;
  try {
    const blob = await new Promise((res) => c.toBlob(res, "image/png"));
    if (blob && navigator.clipboard && window.ClipboardItem) {
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
      copied = true;
      showToast(cutAfter ? "Cut to clipboard." : "Copied to clipboard.");
    } else {
      showToast("Clipboard not supported in this browser.");
    }
  } catch (err) {
    console.warn(err);
    showToast("Could not use clipboard — try another browser.");
  }
  if (cutAfter && copied) {
    beforeOverlayMutation();
    els.overlayCanvas.getContext("2d").clearRect(rect.x, rect.y, rect.w, rect.h);
    stashOverlay();
    schedulePersist();
  }
  hideSelectionToolbar();
}

function deleteSelection() {
  const rect = activeSelection;
  if (!rect) return;
  beforeOverlayMutation();
  els.overlayCanvas.getContext("2d").clearRect(rect.x, rect.y, rect.w, rect.h);
  stashOverlay();
  schedulePersist();
  hideSelectionToolbar();
  showToast("Selection cleared.");
}

function highlightSelection(hex) {
  const rect = activeSelection;
  if (!rect) return;
  beforeOverlayMutation();
  const ctx = els.overlayCanvas.getContext("2d");
  ctx.save();
  ctx.fillStyle = hex + "59";
  ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
  ctx.restore();
  stashOverlay();
  schedulePersist();
  hideSelectionToolbar();
  showToast("Highlight applied.");
}

function buildColorSwatches() {
  if (!els.stColors) return;
  els.stColors.innerHTML = "";
  HIGHLIGHT_PALETTE.forEach((hex) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "st-swatch";
    b.style.background = hex;
    b.title = hex;
    b.addEventListener("click", () => highlightSelection(hex));
    els.stColors.appendChild(b);
  });
}

document.addEventListener("mousedown", (e) => {
  if (!els.selectionToolbar || els.selectionToolbar.classList.contains("hidden")) return;
  if (els.selectionToolbar.contains(e.target)) return;
  if (e.target.closest(".page-canvas-inner")) return;
  hideSelectionToolbar();
});

els.overlayCanvas.addEventListener("mousedown", (e) => {
  if (placementMode) return;
  if (activeTool === "text") onTextClick(e);
  else startDraw(e);
});
els.overlayCanvas.addEventListener("mousemove", moveDraw);
els.overlayCanvas.addEventListener("mouseup", endDraw);
els.overlayCanvas.addEventListener("mouseleave", (e) => {
  if (activeTool === "select" && isSelecting) finalizeSelection();
  else endDraw();
});

els.overlayCanvas.addEventListener(
  "touchstart",
  (e) => {
    if (placementMode) return;
    e.preventDefault();
    const t = e.touches[0];
    if (activeTool === "text") {
      onTextClick({ clientX: t.clientX, clientY: t.clientY });
    } else startDraw({ clientX: t.clientX, clientY: t.clientY });
  },
  { passive: false }
);
els.overlayCanvas.addEventListener(
  "touchmove",
  (e) => {
    e.preventDefault();
    const t = e.touches[0];
    moveDraw({ clientX: t.clientX, clientY: t.clientY });
  },
  { passive: false }
);
els.overlayCanvas.addEventListener("touchend", endDraw);

document.querySelectorAll(".tool-btn[data-tool]").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (btn.dataset.tool === "organize") {
      if (!pdfDoc) {
        showToast("Open a PDF first.");
        return;
      }
      showToast(`Export order: ${exportOrder.length} pages. Use Duplicate to add copies.`);
      return;
    }
    setTool(btn.dataset.tool);
  });
});

if (els.btnUndo) els.btnUndo.addEventListener("click", () => undo());
if (els.btnRedo) els.btnRedo.addEventListener("click", () => redo());

if (els.btnAnotherPdf) {
  els.btnAnotherPdf.addEventListener("click", () => {
    if (els.filePicker) els.filePicker.click();
  });
}

if (els.stCut) els.stCut.addEventListener("click", () => copySelectionToClipboard(true));
if (els.stCopy) els.stCopy.addEventListener("click", () => copySelectionToClipboard(false));
if (els.stDelete) els.stDelete.addEventListener("click", () => deleteSelection());

document.addEventListener("keydown", (e) => {
  if (!pdfDoc) return;
  const mod = e.ctrlKey || e.metaKey;
  if (!mod) return;
  if (e.key === "z" && !e.shiftKey) {
    e.preventDefault();
    undo();
  } else if (e.key === "y" || (e.key === "z" && e.shiftKey)) {
    e.preventDefault();
    redo();
  }
});

if (els.filePicker) {
  els.filePicker.addEventListener("change", (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) loadPdf(f);
    e.target.value = "";
  });
}

async function loadPdfFromBytes(bytes, name, restoredOverlays, restoredOrder, restoredPage, restoredScale) {
  pdfBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  fileName = name || "document.pdf";
  els.docTitle.textContent = fileName.replace(/\.pdf$/i, "") || "Document";

  const loadingTask = pdfjsLib.getDocument({
    data: pdfBytes.slice(0),
    verbosity: 0,
  });
  pdfDoc = await loadingTask.promise;
  numPages = pdfDoc.numPages;
  if (Array.isArray(restoredOrder) && restoredOrder.length > 0) {
    exportOrder = restoredOrder.filter((n) => Number.isFinite(n) && n >= 1 && n <= numPages);
  }
  if (!exportOrder.length) {
    exportOrder = Array.from({ length: numPages }, (_, i) => i + 1);
  }
  overlayStorage = restoredOverlays && typeof restoredOverlays === "object" ? { ...restoredOverlays } : {};
  currentPage = Math.min(Math.max(1, restoredPage || 1), numPages);
  viewScale = typeof restoredScale === "number" ? Math.min(2.5, Math.max(0.5, restoredScale)) : 1;
  pageBaseViewport = {};
  pageListOrder = Array.from({ length: numPages }, (_, i) => i + 1);

  els.emptyState.classList.add("hidden");
  els.pageStack.classList.remove("hidden");
  els.btnExport.disabled = false;
  els.btnSave.disabled = false;

  clearUndoRedoAll();
  await renderThumbnails();
  await renderCurrentPage();
  updateUndoButtons();
  els.statusText.textContent = `${numPages} page${numPages === 1 ? "" : "s"} · ${fileName}`;
}

async function loadPdf(file) {
  if (!file) return;
  if (typeof pdfjsLib === "undefined") {
    showToast("PDF engine not loaded. Refresh the page or run from the project folder.");
    return;
  }

  setPdfLoading(true);
  await new Promise((r) => requestAnimationFrame(r));

  try {
    stashOverlay();
    const raw = new Uint8Array(await file.arrayBuffer());
    if (raw.byteLength < 8) {
      throw new Error("File is empty or too small to be a PDF.");
    }
    clearUndoRedoAll();
    await loadPdfFromBytes(raw, file.name || "document.pdf", {}, null, 1, 1);
    await saveSession();
    showToast("PDF ready — edits are saved in this browser automatically.");
  } catch (err) {
    console.error(err);
    const msg =
      err && err.message
        ? err.message
        : "Could not open this file. Make sure it is a valid PDF.";
    showToast(msg);
    els.statusText.textContent = "Could not load PDF.";
    pdfDoc = null;
    numPages = 0;
    els.emptyState.classList.remove("hidden");
    els.pageStack.classList.add("hidden");
    els.btnExport.disabled = true;
    els.btnSave.disabled = true;
  } finally {
    setPdfLoading(false);
  }
}

async function tryRestoreSession() {
  if (typeof pdfjsLib === "undefined") return;
  try {
    const buf = await idbGet("pdfBytes");
    if (!buf) return;
    const fn = (await idbGet("fileName")) || "document.pdf";
    const ord = await idbGet("exportOrder");
    const ovs = await idbGet("overlayStorage");
    const cp = await idbGet("currentPage");
    const vs = await idbGet("viewScale");
    setPdfLoading(true);
    await new Promise((r) => requestAnimationFrame(r));
    let bytes;
    if (buf instanceof Uint8Array) bytes = buf;
    else if (buf instanceof ArrayBuffer) bytes = new Uint8Array(buf);
    else if (buf && buf.buffer) bytes = new Uint8Array(buf.buffer);
    else return;
    await loadPdfFromBytes(bytes, fn, ovs, ord, cp, vs);
    await saveSession();
    showToast("Restored your last document from this browser.");
  } catch (e) {
    console.warn("Session restore failed", e);
  } finally {
    setPdfLoading(false);
  }
}

let dragFromPage = null;

async function renderThumbnails() {
  els.thumbnails.innerHTML = "";
  const order = pageListOrder.length ? pageListOrder : Array.from({ length: numPages }, (_, i) => i + 1);
  pageListOrder = [...order];

  let slot = 0;
  for (const pdfPage of pageListOrder) {
    const page = await pdfDoc.getPage(pdfPage);
    const vp = page.getViewport({ scale: THUMB_SCALE });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = vp.width;
    canvas.height = vp.height;
    await page.render({ canvasContext: ctx, viewport: vp }).promise;

    const wrap = document.createElement("button");
    wrap.type = "button";
    wrap.className = "thumb" + (pdfPage === currentPage ? " active" : "");
    wrap.dataset.page = String(pdfPage);
    wrap.dataset.slot = String(slot);
    wrap.draggable = true;
    wrap.appendChild(canvas);
    const badge = document.createElement("span");
    badge.className = "thumb-num";
    badge.textContent = String(pdfPage);
    wrap.appendChild(badge);

    wrap.addEventListener("click", async () => {
      stashOverlay();
      currentPage = pdfPage;
      document.querySelectorAll(".thumb").forEach((t) =>
        t.classList.toggle("active", +t.dataset.page === pdfPage)
      );
      await renderCurrentPage();
      updateUndoButtons();
    });

    wrap.addEventListener("dragstart", (e) => {
      dragFromPage = pdfPage;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("application/x-docsfix-page", String(pdfPage));
      wrap.style.opacity = "0.5";
    });
    wrap.addEventListener("dragend", () => {
      wrap.style.opacity = "1";
      document.querySelectorAll(".thumb").forEach((t) => t.classList.remove("drag-over"));
    });
    wrap.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      wrap.classList.add("drag-over");
    });
    wrap.addEventListener("dragleave", () => wrap.classList.remove("drag-over"));
    wrap.addEventListener("drop", async (e) => {
      e.preventDefault();
      wrap.classList.remove("drag-over");
      const fromP = dragFromPage;
      const toP = pdfPage;
      if (fromP == null || fromP === toP) return;
      const arr = [...pageListOrder];
      const iFrom = arr.indexOf(fromP);
      const iTo = arr.indexOf(toP);
      if (iFrom < 0 || iTo < 0) return;
      arr.splice(iFrom, 1);
      arr.splice(iTo, 0, fromP);
      pageListOrder = arr;
      await reorderPdfPages(arr);
    });

    els.thumbnails.appendChild(wrap);
    slot++;
  }
}

async function reorderPdfPages(order1Based) {
  if (!PDFDocument || !pdfBytes || order1Based.length !== numPages) return;
  try {
    setPdfLoading(true);
    const src = await PDFDocument.load(pdfBytes.slice(0));
    const out = await PDFDocument.create();
    const idx = order1Based.map((p) => p - 1);
    const copied = await out.copyPages(src, idx);
    copied.forEach((p) => out.addPage(p));
    const newBytes = await out.save();
    const newOverlays = {};
    order1Based.forEach((oldPage, i) => {
      const k = String(oldPage);
      if (overlayStorage[oldPage] != null) newOverlays[i + 1] = overlayStorage[oldPage];
      else if (overlayStorage[k] != null) newOverlays[i + 1] = overlayStorage[k];
    });
    overlayStorage = newOverlays;
    pdfBytes = new Uint8Array(newBytes);
    const task = pdfjsLib.getDocument({ data: pdfBytes.slice(0), verbosity: 0 });
    if (pdfDoc && pdfDoc.destroy) {
      try {
        await pdfDoc.destroy();
      } catch (_) {}
    }
    pdfDoc = await task.promise;
    numPages = pdfDoc.numPages;
    pageListOrder = Array.from({ length: numPages }, (_, i) => i + 1);
    exportOrder = [...pageListOrder];
    currentPage = Math.min(Math.max(1, currentPage), numPages);
    pageBaseViewport = {};
    clearUndoRedoAll();
    await renderThumbnails();
    await renderCurrentPage();
    await saveSession();
    showToast("Page order updated.");
  } catch (err) {
    console.error(err);
    showToast("Could not reorder pages.");
  } finally {
    setPdfLoading(false);
  }
}

/**
 * @param {Uint8Array|ArrayBuffer} newBytes
 * @param {{ nextName?: string; overlayMap?: Record<number, string> | null }} [opts]
 */
async function reloadPdfFromLib(newBytes, opts) {
  if (!pdfBytes) return;
  const nextName = opts && opts.nextName;
  const overlayMap = opts && opts.overlayMap;
  setPdfLoading(true);
  try {
    stashOverlay();
    pdfBytes = newBytes instanceof Uint8Array ? newBytes : new Uint8Array(newBytes);
    if (nextName) fileName = nextName;
    const task = pdfjsLib.getDocument({ data: pdfBytes.slice(0), verbosity: 0 });
    if (pdfDoc && pdfDoc.destroy) {
      try {
        await pdfDoc.destroy();
      } catch (_) {}
    }
    pdfDoc = await task.promise;
    numPages = pdfDoc.numPages;
    pageListOrder = Array.from({ length: numPages }, (_, i) => i + 1);
    exportOrder = [...pageListOrder];
    overlayStorage = overlayMap && typeof overlayMap === "object" ? { ...overlayMap } : {};
    pageBaseViewport = {};
    currentPage = Math.min(Math.max(1, currentPage), numPages);
    clearUndoRedoAll();
    await renderThumbnails();
    await renderCurrentPage();
    await saveSession();
  } catch (e) {
    console.error(e);
    showToast(e.message || "PDF update failed.");
  } finally {
    setPdfLoading(false);
  }
}

async function renderCurrentPage() {
  if (!pdfDoc) return;
  hideSelectionToolbar();
  clearSelectionHud();
  const page = await pdfDoc.getPage(currentPage);
  const baseVp = page.getViewport({ scale: 1 });
  pageBaseViewport[currentPage] = { w: baseVp.width, h: baseVp.height };
  const viewport = page.getViewport({ scale: viewScale });

  const pdfCanvas = els.pdfCanvas;
  const ctx = pdfCanvas.getContext("2d");
  pdfCanvas.width = viewport.width;
  pdfCanvas.height = viewport.height;
  await page.render({ canvasContext: ctx, viewport }).promise;

  syncOverlaySize();
  restoreOverlay(currentPage);
  els.zoomLabel.textContent = `${Math.round(viewScale * 100)}%`;
  updateUndoButtons();
}

els.zoomIn.addEventListener("click", async () => {
  viewScale = Math.min(2.5, viewScale + 0.15);
  stashOverlay();
  await renderCurrentPage();
  schedulePersist();
});
els.zoomOut.addEventListener("click", async () => {
  viewScale = Math.max(0.5, viewScale - 0.15);
  stashOverlay();
  await renderCurrentPage();
  schedulePersist();
});

els.btnDuplicate.addEventListener("click", () => {
  if (!pdfDoc) return;
  let idx = exportOrder.lastIndexOf(currentPage);
  if (idx < 0) idx = exportOrder.length - 1;
  exportOrder.splice(idx + 1, 0, currentPage);
  showToast(`Copy of page ${currentPage} added to export order.`);
  schedulePersist();
});

els.btnInfo.addEventListener("click", () => {
  if (!pdfDoc) {
    showToast("Open a PDF first.");
    return;
  }
  els.infoBody.innerHTML = `
    <strong>File:</strong> ${escapeHtml(fileName)}<br/>
    <strong>Original pages:</strong> ${numPages}<br/>
    <strong>Export pages:</strong> ${exportOrder.length}<br/>
    <em>Edits are kept in this browser (IndexedDB) until you open another PDF.</em>
  `;
  els.infoDialog.showModal();
});

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}

els.searchInput.addEventListener("keydown", async (e) => {
  if (e.key !== "Enter" || !pdfDoc) return;
  const q = els.searchInput.value.trim();
  if (/^\d+$/.test(q)) {
    const n = parseInt(q, 10);
    if (n >= 1 && n <= numPages) {
      stashOverlay();
      currentPage = n;
      document.querySelectorAll(".thumb").forEach((t) => t.classList.toggle("active", +t.dataset.page === n));
      await renderCurrentPage();
      document.querySelector(`.thumb[data-page="${n}"]`)?.scrollIntoView({ block: "nearest" });
      showToast(`Page ${n}.`);
    } else {
      showToast("Invalid page number.");
    }
    return;
  }
  e.preventDefault();
  const hits = await runTextSearchAll(q);
  if (!hits.length) {
    showToast("No text matches.");
    if (els.searchPanel) els.searchPanel.classList.add("hidden");
  } else {
    showSearchPanel(hits, q);
  }
});

async function runTextSearchAll(query) {
  const q = query.trim();
  if (!q || !pdfDoc) return [];
  const hits = [];
  const lower = q.toLowerCase();
  for (let pi = 1; pi <= numPages; pi++) {
    const page = await pdfDoc.getPage(pi);
    const tc = await page.getTextContent({ normalizeWhitespace: true });
    for (const it of tc.items) {
      if (!it.str) continue;
      if (it.str.toLowerCase().includes(lower)) {
        hits.push({ page: pi, text: it.str.slice(0, 160) });
        break;
      }
    }
  }
  return hits;
}

function showSearchPanel(hits, query) {
  if (!els.searchResults || !els.searchPanel) return;
  els.searchResults.innerHTML = "";
  if (els.searchPanelTitle) els.searchPanelTitle.textContent = `${hits.length} match(es) for "${query}"`;
  hits.forEach((h) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "search-hit";
    b.innerHTML = `<strong>Page ${h.page}</strong><small>${escapeHtml(h.text)}</small>`;
    b.addEventListener("click", async () => {
      stashOverlay();
      currentPage = h.page;
      document.querySelectorAll(".thumb").forEach((t) => t.classList.toggle("active", +t.dataset.page === h.page));
      await renderThumbnails();
      await renderCurrentPage();
      els.searchPanel.classList.add("hidden");
    });
    els.searchResults.appendChild(b);
  });
  els.searchPanel.classList.remove("hidden");
}

if (els.searchPanelClose) {
  els.searchPanelClose.addEventListener("click", () => {
    if (els.searchPanel) els.searchPanel.classList.add("hidden");
  });
}

async function compositePageImage(pageNum) {
  const page = await pdfDoc.getPage(pageNum);
  const baseVp = page.getViewport({ scale: 1 });
  const viewport = page.getViewport({ scale: EXPORT_SCALE });
  const canvas = document.createElement("canvas");
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext("2d");
  await page.render({ canvasContext: ctx, viewport }).promise;

  const url = overlayStorage[pageNum];
  if (url) {
    await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        ctx.drawImage(img, 0, 0, baseVp.width, baseVp.height, 0, 0, viewport.width, viewport.height);
        resolve();
      };
      img.onerror = reject;
      img.src = url;
    });
  }

  const blob = await new Promise((res) => canvas.toBlob(res, "image/png"));
  return { blob, width: viewport.width, height: viewport.height };
}

function downloadBlob(data, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(data);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}

els.btnExport.addEventListener("click", async () => {
  if (!PDFDocument) {
    showToast("Export library not loaded — check your connection and refresh.");
    return;
  }
  if (!pdfDoc || !pdfBytes) return;
  stashOverlay();
  els.btnExport.disabled = true;
  els.statusText.textContent = "Building PDF…";
  try {
    const out = await PDFDocument.create();
    for (const pageNum of exportOrder) {
      const { blob, width, height } = await compositePageImage(pageNum);
      const bytes = new Uint8Array(await blob.arrayBuffer());
      const img = await out.embedPng(bytes);
      const page = out.addPage([width, height]);
      page.drawImage(img, { x: 0, y: 0, width, height });
    }
    const outBytes = await out.save();
    const base = fileName.replace(/\.pdf$/i, "") || "document";
    downloadBlob(new Blob([outBytes], { type: "application/pdf" }), `${base}-docsfix.pdf`);
    showToast("Exported — check your downloads.");
    els.statusText.textContent = `${exportOrder.length} pages exported · ${fileName}`;
  } catch (err) {
    console.error(err);
    showToast("Export failed — see console.");
    els.statusText.textContent = "Export failed";
  } finally {
    els.btnExport.disabled = false;
  }
});

els.btnSave.addEventListener("click", () => {
  stashOverlay();
  const draft = {
    version: 2,
    fileName,
    exportOrder: [...exportOrder],
    overlays: { ...overlayStorage },
  };
  const blob = new Blob([JSON.stringify(draft)], { type: "application/json" });
  downloadBlob(blob, `${(fileName || "doc").replace(/\.pdf$/i, "")}-docsfix-draft.json`);
  showToast("Draft JSON downloaded.");
});

async function addBlankPage() {
  if (!PDFDocument || !pdfBytes) return;
  try {
    const src = await PDFDocument.load(pdfBytes.slice(0));
    const last = src.getPage(src.getPageCount() - 1);
    const { width, height } = last.getSize();
    src.addPage([width, height]);
    const nb = await src.save();
    await reloadPdfFromLib(new Uint8Array(nb), {});
    showToast("Blank page added.");
  } catch (e) {
    console.error(e);
    showToast("Could not add page.");
  }
}

async function deleteCurrentPdfPage() {
  if (!PDFDocument || !pdfBytes || numPages <= 1) {
    showToast("Cannot delete the only page.");
    return;
  }
  try {
    const src = await PDFDocument.load(pdfBytes.slice(0));
    src.removePage(currentPage - 1);
    const nb = await src.save();
    const newO = {};
    for (let p = 1; p <= numPages; p++) {
      if (p < currentPage) newO[p] = overlayStorage[p];
      else if (p > currentPage) newO[p - 1] = overlayStorage[p];
    }
    currentPage = Math.max(1, Math.min(currentPage, numPages - 1));
    await reloadPdfFromLib(new Uint8Array(nb), { overlayMap: newO });
    showToast("Page deleted.");
  } catch (e) {
    console.error(e);
    showToast("Could not delete page.");
  }
}

async function mergePdfFiles(fileList) {
  if (!PDFDocument || !pdfBytes || !fileList || !fileList.length) return;
  try {
    const merged = await PDFDocument.create();
    const first = await PDFDocument.load(pdfBytes.slice(0));
    (await merged.copyPages(first, first.getPageIndices())).forEach((p) => merged.addPage(p));
    for (const f of fileList) {
      const doc = await PDFDocument.load(await f.arrayBuffer());
      (await merged.copyPages(doc, doc.getPageIndices())).forEach((p) => merged.addPage(p));
    }
    const out = await merged.save();
    const base = fileName.replace(/\.pdf$/i, "") || "document";
    await reloadPdfFromLib(new Uint8Array(out), { nextName: base + "-merged.pdf" });
    showToast("PDFs merged.");
  } catch (e) {
    console.error(e);
    showToast("Merge failed.");
  }
}

async function splitPdfHere() {
  if (!PDFDocument || !pdfBytes) return;
  try {
    const src = await PDFDocument.load(pdfBytes.slice(0));
    const n = src.getPageCount();
    if (n < 2) {
      showToast("Need at least 2 pages to split.");
      return;
    }
    const cp = Math.min(Math.max(1, currentPage), n - 1);
    const idxA = Array.from({ length: cp }, (_, i) => i);
    const idxB = Array.from({ length: n - cp }, (_, i) => i + cp);
    const a = await PDFDocument.create();
    const b = await PDFDocument.create();
    (await a.copyPages(src, idxA)).forEach((p) => a.addPage(p));
    (await b.copyPages(src, idxB)).forEach((p) => b.addPage(p));
    const base = fileName.replace(/\.pdf$/i, "") || "document";
    downloadBlob(new Blob([await a.save()], { type: "application/pdf" }), base + "-part1.pdf");
    downloadBlob(new Blob([await b.save()], { type: "application/pdf" }), base + "-part2.pdf");
    showToast(`Part 1: pages 1–${cp}. Part 2: pages ${cp + 1}–${n}. Downloaded.`);
  } catch (e) {
    console.error(e);
    showToast("Split failed.");
  }
}

async function applyReplaceOnCurrentPage(find, replace) {
  if (!pdfDoc || !find) return;
  const page = await pdfDoc.getPage(currentPage);
  const vp = page.getViewport({ scale: viewScale });
  const tc = await page.getTextContent();
  beforeOverlayMutation();
  const ctx = els.overlayCanvas.getContext("2d");
  let n = 0;
  const util = pdfjsLib.Util;
  for (const it of tc.items) {
    if (!it.str || !it.str.includes(find)) continue;
    const tr = util ? util.transform(vp.transform, it.transform) : it.transform;
    const x = tr[4];
    const y = tr[5];
    const fh = Math.sqrt(tr[2] * tr[2] + tr[3] * tr[3]) || 12;
    const tw = Math.max((it.width || fh * it.str.length * 0.45) * 1.1, fh * find.length * 0.35);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(x, y - fh * 0.85, tw, fh * 1.35);
    ctx.font = `600 ${fh}px DM Sans, sans-serif`;
    ctx.fillStyle = "#1a2b4a";
    ctx.fillText(it.str.split(find).join(replace), x, y);
    n++;
  }
  stashOverlay();
  schedulePersist();
  showToast(n ? `Updated ${n} text run(s) (overlay).` : "No matching text on this page.");
}

function rewriteSelectionText() {
  const r = activeSelection;
  if (!r) return;
  const t = window.prompt("New text for this area:", "");
  if (t === null) return;
  beforeOverlayMutation();
  const ctx = els.overlayCanvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(r.x, r.y, r.w, r.h);
  ctx.font = "600 14px DM Sans, sans-serif";
  ctx.fillStyle = "#1a2b4a";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(t, r.x + r.w / 2, r.y + r.h / 2);
  ctx.textAlign = "left";
  ctx.textBaseline = "alphabetic";
  stashOverlay();
  schedulePersist();
  hideSelectionToolbar();
  showToast("Text area updated.");
}

function drawShapeInSelection(kind) {
  const r = activeSelection;
  if (!r) return;
  beforeOverlayMutation();
  const ctx = els.overlayCanvas.getContext("2d");
  ctx.strokeStyle = "#1a2b4a";
  ctx.fillStyle = "rgba(43, 127, 255, 0.15)";
  ctx.lineWidth = 2.5;
  if (kind === "rect") {
    ctx.strokeRect(r.x + 1, r.y + 1, r.w - 2, r.h - 2);
  } else if (kind === "circle") {
    ctx.beginPath();
    ctx.ellipse(r.x + r.w / 2, r.y + r.h / 2, r.w / 2 - 1, r.h / 2 - 1, 0, 0, Math.PI * 2);
    ctx.stroke();
    ctx.fill();
  } else if (kind === "arrow") {
    const x1 = r.x + r.w * 0.15;
    const y1 = r.y + r.h * 0.5;
    const x2 = r.x + r.w * 0.85;
    const y2 = r.y + r.h * 0.5;
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    const ang = Math.atan2(y2 - y1, x2 - x1);
    const hs = Math.min(14, r.h * 0.35);
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - hs * Math.cos(ang - Math.PI / 6), y2 - hs * Math.sin(ang - Math.PI / 6));
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - hs * Math.cos(ang + Math.PI / 6), y2 - hs * Math.sin(ang + Math.PI / 6));
    ctx.stroke();
  }
  stashOverlay();
  schedulePersist();
  hideSelectionToolbar();
}

function initSignaturePad() {
  const pad = els.signaturePad;
  if (!pad) return;
  const pctx = pad.getContext("2d");
  let down = false;
  let lx = 0,
    ly = 0;
  function pos(ev) {
    const rect = pad.getBoundingClientRect();
    const x = ((ev.clientX - rect.left) / rect.width) * pad.width;
    const y = ((ev.clientY - rect.top) / rect.height) * pad.height;
    return { x, y };
  }
  function start(ev) {
    down = true;
    const p = pos(ev);
    lx = p.x;
    ly = p.y;
  }
  function move(ev) {
    if (!down) return;
    const p = pos(ev);
    pctx.strokeStyle = "#1a2b4a";
    pctx.lineWidth = 2;
    pctx.lineCap = "round";
    pctx.beginPath();
    pctx.moveTo(lx, ly);
    pctx.lineTo(p.x, p.y);
    pctx.stroke();
    lx = p.x;
    ly = p.y;
  }
  function end() {
    down = false;
  }
  pad.addEventListener("mousedown", start);
  pad.addEventListener("mousemove", move);
  pad.addEventListener("mouseup", end);
  pad.addEventListener("mouseleave", end);
  pad.addEventListener(
    "touchstart",
    (e) => {
      e.preventDefault();
      start(e.touches[0]);
    },
    { passive: false }
  );
  pad.addEventListener(
    "touchmove",
    (e) => {
      e.preventDefault();
      move(e.touches[0]);
    },
    { passive: false }
  );
  pad.addEventListener("touchend", end);
  if (els.sigClear) {
    els.sigClear.addEventListener("click", () => pctx.clearRect(0, 0, pad.width, pad.height));
  }
}

if (els.btnAddPage) els.btnAddPage.addEventListener("click", () => addBlankPage());
if (els.btnDelPage) els.btnDelPage.addEventListener("click", () => deleteCurrentPdfPage());
if (els.btnMergePdf) {
  els.btnMergePdf.addEventListener("click", () => els.mergePicker && els.mergePicker.click());
}
if (els.mergePicker) {
  els.mergePicker.addEventListener("change", (e) => {
    const fs = e.target.files;
    if (fs && fs.length) mergePdfFiles(fs);
    e.target.value = "";
  });
}
if (els.btnSplitPdf) els.btnSplitPdf.addEventListener("click", () => splitPdfHere());

if (els.btnReplaceText) {
  els.btnReplaceText.addEventListener("click", async () => {
    if (!pdfDoc) {
      showToast("Open a PDF first.");
      return;
    }
    const find = window.prompt("Find text on this page:");
    if (!find) return;
    const rep = window.prompt("Replace with:");
    if (rep === null) return;
    await applyReplaceOnCurrentPage(find, rep);
  });
}

if (els.btnSignature) {
  els.btnSignature.addEventListener("click", () => {
    if (!pdfDoc) {
      showToast("Open a PDF first.");
      return;
    }
    if (els.signatureDialog) {
      els.signatureDialog.showModal();
      const p = els.signaturePad;
      if (p) p.getContext("2d").clearRect(0, 0, p.width, p.height);
    }
  });
}

if (els.signatureForm) {
  els.signatureForm.addEventListener("submit", (e) => {
    e.preventDefault();
    if (!els.signaturePad) return;
    placementDataUrl = els.signaturePad.toDataURL("image/png");
    placementMode = "sig";
    if (els.signatureDialog) els.signatureDialog.close();
    showToast("Click on the page to place your signature.");
  });
}

if (els.btnInsertImage) {
  els.btnInsertImage.addEventListener("click", () => {
    if (!pdfDoc) {
      showToast("Open a PDF first.");
      return;
    }
    els.imagePicker && els.imagePicker.click();
  });
}
if (els.imagePicker) {
  els.imagePicker.addEventListener("change", (e) => {
    const f = e.target.files && e.target.files[0];
    if (!f || !f.type.startsWith("image/")) return;
    const img = new Image();
    img.onload = () => {
      placementImgEl = img;
      placementMode = "img";
      showToast("Click on the page to place the image.");
    };
    img.src = URL.createObjectURL(f);
    e.target.value = "";
  });
}

els.overlayCanvas.addEventListener(
  "click",
  (e) => {
    if (!placementMode || !pdfDoc) return;
    e.stopPropagation();
    e.preventDefault();
    const p = getOverlayPoint(e);
    beforeOverlayMutation();
    const ctx = els.overlayCanvas.getContext("2d");
    if (placementMode === "sig" && placementDataUrl) {
      const im = new Image();
      im.onload = () => {
        const w = Math.min(200, im.naturalWidth);
        const h = (im.naturalHeight / im.naturalWidth) * w;
        ctx.drawImage(im, p.x, p.y, w, h);
        stashOverlay();
        schedulePersist();
      };
      im.src = placementDataUrl;
    } else if (placementMode === "img" && placementImgEl) {
      const im = placementImgEl;
      const w = Math.min(240, im.naturalWidth);
      const h = (im.naturalHeight / im.naturalWidth) * w;
      ctx.drawImage(im, p.x, p.y, w, h);
      stashOverlay();
      schedulePersist();
    }
    placementMode = null;
    placementDataUrl = null;
    placementImgEl = null;
  },
  true
);

if (els.stRewrite) els.stRewrite.addEventListener("click", () => rewriteSelectionText());
if (els.stShapeRect) els.stShapeRect.addEventListener("click", () => drawShapeInSelection("rect"));
if (els.stShapeCircle) els.stShapeCircle.addEventListener("click", () => drawShapeInSelection("circle"));
if (els.stShapeArrow) els.stShapeArrow.addEventListener("click", () => drawShapeInSelection("arrow"));

initSignaturePad();
buildColorSwatches();
setTool("draw");

if (typeof pdfjsLib === "undefined") {
  els.statusText.textContent = "Missing vendor/pdf.min.js — open the folder that contains index.html.";
} else if (!PDFDocument) {
  els.statusText.textContent = "Missing pdf-lib — check vendor/pdf-lib.min.js next to index.html.";
} else {
  tryRestoreSession();
}
