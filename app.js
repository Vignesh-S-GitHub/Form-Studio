const fieldCatalog = [
  // Insert
  { type: "image", label: "Picture", icon: "IMG", group: "Insert" },
  { type: "shaperect", label: "Rectangle", icon: "RECT", group: "Insert" },
  { type: "shapecircle", label: "Circle", icon: "CIR", group: "Insert" },
  { type: "shapeline", label: "Line", icon: "LINE", group: "Insert" },
  { type: "elementstar", label: "Star Element", icon: "STAR", group: "Insert" },
  { type: "elementarrow", label: "Arrow Element", icon: "ARW", group: "Insert" },
  // Layout & Structure
  { type: "table", label: "Table Grid", icon: "TBL", group: "Layout" },
  { type: "sectionheader", label: "Section Header", icon: "SEC", group: "Layout" },
  { type: "divider", label: "Divider Line", icon: "DIV", group: "Layout" },
  { type: "constant", label: "Constant/Label", icon: "TXT", group: "Layout" },
  // Text Input
  { type: "text", label: "Short Text", icon: "T", group: "Input" },
  { type: "email", label: "Email", icon: "@", group: "Input" },
  { type: "phone", label: "Phone", icon: "TEL", group: "Input" },
  { type: "url", label: "URL", icon: "URL", group: "Input" },
  { type: "textarea", label: "Long Text", icon: "PAR", group: "Input" },
  // Numbers & Currency
  { type: "number", label: "Number", icon: "#", group: "Number" },
  { type: "currency", label: "Currency", icon: "$", group: "Number" },
  { type: "percentage", label: "Percentage", icon: "%", group: "Number" },
  // Selection
  { type: "select", label: "Dropdown", icon: "SEL", group: "Choice" },
  { type: "radio", label: "Radio Buttons", icon: "RAD", group: "Choice" },
  { type: "checkbox", label: "Checkboxes", icon: "CHK", group: "Choice" },
  { type: "multiselect", label: "Multi-Select", icon: "MUL", group: "Choice" },
  // Date & Time
  { type: "date", label: "Date", icon: "DT", group: "Date/Time" },
  { type: "time", label: "Time", icon: "TM", group: "Date/Time" },
  { type: "datetime", label: "Date & Time", icon: "DTT", group: "Date/Time" },
  // File & Media
  { type: "file", label: "File Upload", icon: "FIL", group: "Media" },
  { type: "photo", label: "Photo Upload", icon: "PHT", group: "Media" },
  // Special
  { type: "signature", label: "Signature", icon: "SIG", group: "Special" },
  { type: "rating", label: "Rating", icon: "RATE", group: "Special" }
];

const pageFormats = {
  A5: { w: 148, h: 210 },
  A4: { w: 210, h: 297 },
  A3: { w: 297, h: 420 },
  Letter: { w: 216, h: 279 },
  Legal: { w: 216, h: 356 },
  Tabloid: { w: 279, h: 432 },
  Executive: { w: 184, h: 267 }
};

const defaultByType = {
  constant: { label: "Header Text", value: "Your Text Here", required: false, w: 280, h: 60 },
  text: { label: "Text Field", placeholder: "Enter text", required: true, w: 240, h: 50 },
  email: { label: "Email Address", placeholder: "name@example.com", required: true, w: 240, h: 50 },
  phone: { label: "Phone Number", placeholder: "+1 (555) 000-0000", required: false, w: 240, h: 50 },
  url: { label: "Website URL", placeholder: "https://example.com", required: false, w: 240, h: 50 },
  number: { label: "Number", placeholder: "0", required: false, w: 180, h: 50 },
  currency: { label: "Amount", placeholder: "$0.00", currency: "USD", required: false, w: 200, h: 50 },
  percentage: { label: "Percentage", placeholder: "0%", required: false, w: 180, h: 50 },
  textarea: { label: "Description", placeholder: "Write details here", required: false, w: 280, h: 100 },
  select: { label: "Choose Option", options: "Option 1,Option 2,Option 3", required: true, w: 240, h: 50 },
  radio: { label: "Select One", options: "Yes,No,Maybe", required: false, w: 240, h: 80 },
  checkbox: { label: "Accept Terms", options: "I agree", required: true, w: 240, h: 40 },
  multiselect: { label: "Select Multiple", options: "Item 1,Item 2,Item 3,Item 4", required: false, w: 260, h: 100 },
  date: { label: "Date", required: false, w: 200, h: 50 },
  time: { label: "Time", required: false, w: 180, h: 50 },
  datetime: { label: "Date and Time", required: false, w: 240, h: 50 },
  file: { label: "Attach File", accept: ".pdf,.doc,.docx", maxFiles: 1, required: false, w: 240, h: 60 },
  photo: { label: "Photo", accept: "image/*", maxFiles: 1, required: false, w: 240, h: 80 },
  signature: { label: "Signature", required: true, w: 280, h: 100 },
  sectionheader: { label: "Section Title", value: "New Section", required: false, w: 280, h: 40 },
  table: {
    label: "Table",
    rows: 3,
    cols: 3,
    colWidths: "",
    rowHeights: "",
    cellMode: "text",
    cellLabels: "",
    borderColor: "#5e75b8",
    borderWidth: 1,
    cellBg: "#ffffff",
    textColor: "#24366b",
    required: false,
    w: 360,
    h: 220
  },
  divider: { label: "Divider", required: false, w: 280, h: 8 },
  rating: { label: "Rating", maxRating: 5, required: false, w: 240, h: 50 },
  image: { label: "Picture", imageSrc: "", fit: "cover", radius: 10, opacity: 100, rotate: 0, required: false, w: 240, h: 180 },
  shaperect: { label: "Rectangle", fillColor: "#dff1eb", strokeColor: "#3b7f71", strokeWidth: 1, radius: 10, opacity: 100, rotate: 0, required: false, w: 180, h: 100 },
  shapecircle: { label: "Circle", fillColor: "#f8ead9", strokeColor: "#ba7a36", strokeWidth: 1, opacity: 100, rotate: 0, required: false, w: 120, h: 120 },
  shapeline: { label: "Line", strokeColor: "#4f6f67", strokeWidth: 2, opacity: 100, rotate: 0, required: false, w: 220, h: 8 },
  elementstar: { label: "Star", fillColor: "#f3ca62", opacity: 100, rotate: 0, required: false, w: 80, h: 80 },
  elementarrow: { label: "Arrow", fillColor: "#93b9ad", opacity: 100, rotate: 0, required: false, w: 120, h: 50 }
};

const insertTypes = ["image", "shaperect", "shapecircle", "shapeline", "elementstar", "elementarrow"];

const state = {
  selectedId: null,
  selectedIds: [],
  fields: [],
  page: { size: "A4", orientation: "portrait" },
  dragging: null,
  splitSizes: {
    palette: 280,
    canvas: "1fr",
    config: 320
  }
};

const historyState = {
  past: [],
  future: [],
  limit: 120,
  suspend: false
};

const templateStorageKey = "formStudio.latestTemplate";
const draftStorageKey = "formStudio.builderDraft";
let saveTimer = null;
let historyTimer = null;
let suppressNextSelectionClick = false;

const appEl = document.getElementById("builderApp");
const setupDialogEl = document.getElementById("setupDialog");
const setupFormEl = document.getElementById("setupForm");
const setupNameEl = document.getElementById("setupName");
const setupSizeEl = document.getElementById("setupSize");
const setupOrientationEl = document.getElementById("setupOrientation");
const paletteEl = document.getElementById("palette");
const sheetEl = document.getElementById("sheet");
const sheetStageEl = document.getElementById("sheetStage");
const propertyFormEl = document.getElementById("propertyForm");
const fieldCountEl = document.getElementById("fieldCount");
const templateNameEl = document.getElementById("templateName");
const pageSizeSelectEl = document.getElementById("pageSizeSelect");
const orientationSelectEl = document.getElementById("orientationSelect");
const layerListEl = document.getElementById("layerList");
const vGuideEl = document.getElementById("vGuide");
const hGuideEl = document.getElementById("hGuide");
const xMeasureGuideEl = document.getElementById("xMeasureGuide");
const yMeasureGuideEl = document.getElementById("yMeasureGuide");
const xMeasureLabelEl = document.getElementById("xMeasureLabel");
const yMeasureLabelEl = document.getElementById("yMeasureLabel");
const jsonOutputEl = document.getElementById("jsonOutput");
const jsonDialogEl = document.getElementById("jsonDialog");
const btnUndoEl = document.getElementById("btnUndo");
const btnRedoEl = document.getElementById("btnRedo");
const btnBringFrontEl = document.getElementById("btnBringFront");
const btnSendBackEl = document.getElementById("btnSendBack");
const btnToggleLockEl = document.getElementById("btnToggleLock");
const divider1El = document.getElementById("divider1");
const divider2El = document.getElementById("divider2");
const splitPaneEl = document.querySelector(".split-pane");
const essentialToolsEl = document.getElementById("essentialTools");
const toolDecFontEl = document.getElementById("toolDecFont");
const toolIncFontEl = document.getElementById("toolIncFont");
const toolFontSizeEl = document.getElementById("toolFontSize");
const toolBoldEl = document.getElementById("toolBold");
const toolItalicEl = document.getElementById("toolItalic");
const toolUnderlineEl = document.getElementById("toolUnderline");
const toolAlignLeftEl = document.getElementById("toolAlignLeft");
const toolAlignCenterEl = document.getElementById("toolAlignCenter");
const toolAlignRightEl = document.getElementById("toolAlignRight");

let splitDragState = null;

function uid() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function snapshotsEqual(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

function captureSnapshot() {
  return {
    templateName: templateNameEl.value,
    selectedId: state.selectedId,
    selectedIds: deepClone(state.selectedIds),
    page: deepClone(state.page),
    fields: deepClone(state.fields)
  };
}

function applySnapshot(snapshot) {
  historyState.suspend = true;
  state.page = deepClone(snapshot.page || { size: "A4", orientation: "portrait" });
  state.fields = Array.isArray(snapshot.fields) ? deepClone(snapshot.fields) : [];
  state.selectedId = snapshot.selectedId || state.fields[0]?.id || null;
  state.selectedIds = Array.isArray(snapshot.selectedIds) ? [...snapshot.selectedIds] : (state.selectedId ? [state.selectedId] : []);
  normalizeSelection();

  const nextName = (snapshot.templateName || "Untitled Template").toString();
  templateNameEl.value = nextName;
  setupNameEl.value = nextName;
  pageSizeSelectEl.value = state.page.size;
  setupSizeEl.value = state.page.size;
  orientationSelectEl.value = state.page.orientation;
  setupOrientationEl.value = state.page.orientation;

  render();
  historyState.suspend = false;
}

function updateHistoryButtons() {
  btnUndoEl.disabled = historyState.past.length <= 1;
  btnRedoEl.disabled = historyState.future.length === 0;
}

function updateLayerButtons() {
  const hasSelection = state.selectedIds.length > 0;
  btnBringFrontEl.disabled = !hasSelection;
  btnSendBackEl.disabled = !hasSelection;
  btnToggleLockEl.disabled = !hasSelection;
}

function normalizeSelection() {
  const validIds = new Set(state.fields.map((f) => f.id));
  state.selectedIds = state.selectedIds.filter((id) => validIds.has(id));

  if (state.selectedId && !validIds.has(state.selectedId)) {
    state.selectedId = null;
  }
  if (!state.selectedId && state.selectedIds.length) {
    state.selectedId = state.selectedIds[0];
  }
  if (!state.selectedIds.length && state.selectedId) {
    state.selectedIds = [state.selectedId];
  }
}

function isSelected(id) {
  return state.selectedIds.includes(id);
}

function selectOnly(id) {
  state.selectedId = id;
  state.selectedIds = id ? [id] : [];
}

function toggleSelection(id) {
  const idx = state.selectedIds.indexOf(id);
  if (idx >= 0) {
    state.selectedIds.splice(idx, 1);
    if (state.selectedId === id) {
      state.selectedId = state.selectedIds[0] || null;
    }
  } else {
    state.selectedIds.push(id);
    state.selectedId = id;
  }
  normalizeSelection();
}

function getSelectedFields() {
  const set = new Set(state.selectedIds);
  return state.fields.filter((f) => set.has(f.id));
}

function recordHistory() {
  if (historyState.suspend) {
    return;
  }

  const next = captureSnapshot();
  const current = historyState.past[historyState.past.length - 1];
  if (current && snapshotsEqual(current, next)) {
    updateHistoryButtons();
    return;
  }

  historyState.past.push(next);
  if (historyState.past.length > historyState.limit) {
    historyState.past.shift();
  }
  historyState.future = [];
  updateHistoryButtons();
}

function scheduleHistoryRecord() {
  clearTimeout(historyTimer);
  historyTimer = window.setTimeout(recordHistory, 220);
}

function undo() {
  if (historyState.past.length <= 1) {
    return;
  }

  const current = historyState.past.pop();
  historyState.future.push(current);
  const previous = historyState.past[historyState.past.length - 1];
  applySnapshot(previous);
  scheduleDraftSave();
  updateHistoryButtons();
}

function redo() {
  if (!historyState.future.length) {
    return;
  }

  const next = historyState.future.pop();
  historyState.past.push(next);
  applySnapshot(next);
  scheduleDraftSave();
  updateHistoryButtons();
}

function getSheetRect() {
  return sheetEl.getBoundingClientRect();
}

function hideGuides() {
  vGuideEl.classList.add("hidden");
  hGuideEl.classList.add("hidden");
  xMeasureGuideEl.classList.add("hidden");
  yMeasureGuideEl.classList.add("hidden");
  xMeasureLabelEl.classList.add("hidden");
  yMeasureLabelEl.classList.add("hidden");
}

function showVGuide(x) {
  const left = sheetEl.offsetLeft + x;
  vGuideEl.style.left = `${left}px`;
  vGuideEl.style.top = `${sheetEl.offsetTop}px`;
  vGuideEl.style.height = `${sheetEl.clientHeight}px`;
  vGuideEl.classList.remove("hidden");
}

function showHGuide(y) {
  const top = sheetEl.offsetTop + y;
  hGuideEl.style.top = `${top}px`;
  hGuideEl.style.left = `${sheetEl.offsetLeft}px`;
  hGuideEl.style.width = `${sheetEl.clientWidth}px`;
  hGuideEl.classList.remove("hidden");
}

function formatGuideDistance(px) {
  const sheetPx = Math.max(1, sheetEl.clientWidth);
  const size = computeSheetSize();
  const mmPerPx = size.widthMm / sheetPx;
  const mm = px * mmPerPx;
  const inches = mm / 25.4;
  return inches.toFixed(1);
}

function showMeasureGuides(field) {
  const sheetW = sheetEl.clientWidth;
  const sheetH = sheetEl.clientHeight;
  const hasVSnapGuide = !vGuideEl.classList.contains("hidden");
  const hasHSnapGuide = !hGuideEl.classList.contains("hidden");
  const rightGap = Math.max(0, sheetW - (field.x + field.w));
  const bottomGap = Math.max(0, sheetH - (field.y + field.h));

  const horizontalCandidates = [
    { value: field.x, start: 0, end: field.x },
    { value: rightGap, start: field.x + field.w, end: sheetW }
  ];

  const verticalCandidates = [
    { value: field.y, start: 0, end: field.y },
    { value: bottomGap, start: field.y + field.h, end: sheetH }
  ];

  const xGap = horizontalCandidates.sort((a, b) => a.value - b.value)[0];
  const yGap = verticalCandidates.sort((a, b) => a.value - b.value)[0];

  const baseLeft = sheetEl.offsetLeft;
  const baseTop = sheetEl.offsetTop;

  const xLineTop = baseTop + Math.min(sheetH - 12, field.y + field.h + 10);
  if (hasHSnapGuide) {
    xMeasureGuideEl.classList.add("hidden");
    xMeasureLabelEl.classList.add("hidden");
  } else {
    xMeasureGuideEl.style.left = `${baseLeft + xGap.start}px`;
    xMeasureGuideEl.style.top = `${xLineTop}px`;
    xMeasureGuideEl.style.width = `${Math.max(0, xGap.end - xGap.start)}px`;
    xMeasureGuideEl.classList.remove("hidden");

    xMeasureLabelEl.textContent = formatGuideDistance(xGap.value);
    xMeasureLabelEl.style.left = `${baseLeft + (xGap.start + xGap.end) / 2}px`;
    xMeasureLabelEl.style.top = `${xLineTop - 14}px`;
    xMeasureLabelEl.classList.remove("hidden");
  }

  const yLineLeft = baseLeft + Math.min(sheetW - 12, field.x + field.w + 10);
  if (hasVSnapGuide) {
    yMeasureGuideEl.classList.add("hidden");
    yMeasureLabelEl.classList.add("hidden");
  } else {
    yMeasureGuideEl.style.left = `${yLineLeft}px`;
    yMeasureGuideEl.style.top = `${baseTop + yGap.start}px`;
    yMeasureGuideEl.style.height = `${Math.max(0, yGap.end - yGap.start)}px`;
    yMeasureGuideEl.classList.remove("hidden");

    yMeasureLabelEl.textContent = formatGuideDistance(yGap.value);
    yMeasureLabelEl.style.left = `${yLineLeft + 16}px`;
    yMeasureLabelEl.style.top = `${baseTop + (yGap.start + yGap.end) / 2}px`;
    yMeasureLabelEl.classList.remove("hidden");
  }
}

function computeSheetSize() {
  const format = pageFormats[state.page.size] || pageFormats.A4;
  if (state.page.orientation === "landscape") {
    return { widthMm: format.h, heightMm: format.w };
  }
  return { widthMm: format.w, heightMm: format.h };
}

function applySheetSize() {
  const size = computeSheetSize();
  const ratio = `${size.widthMm} / ${size.heightMm}`;
  sheetEl.style.setProperty("--sheet-ratio", ratio);
}

function supportsEssentialTextTools(field) {
  if (!field) {
    return false;
  }
  return !["divider", ...insertTypes].includes(field.type);
}

function isInsertObject(field) {
  return Boolean(field && insertTypes.includes(field.type));
}

function ensureTextStyleDefaults(field) {
  if (!field) {
    return;
  }
  if (typeof field.fontSize !== "number" || Number.isNaN(field.fontSize)) {
    field.fontSize = 14;
  }
  if (!field.fontWeight) {
    field.fontWeight = 500;
  }
  if (!field.fontStyle) {
    field.fontStyle = "normal";
  }
  if (!field.textDecoration) {
    field.textDecoration = "none";
  }
  if (!field.textAlign) {
    field.textAlign = "left";
  }
}

function renderEssentialTools() {
  const selected = getSelectedField();
  if (!selected || !supportsEssentialTextTools(selected)) {
    essentialToolsEl.classList.add("is-disabled");
    [toolDecFontEl, toolIncFontEl, toolFontSizeEl, toolBoldEl, toolItalicEl, toolUnderlineEl, toolAlignLeftEl, toolAlignCenterEl, toolAlignRightEl]
      .forEach((el) => {
        el.disabled = true;
      });
    toolFontSizeEl.value = "14";
    toolBoldEl.classList.remove("active");
    toolItalicEl.classList.remove("active");
    toolUnderlineEl.classList.remove("active");
    toolAlignLeftEl.classList.remove("active");
    toolAlignCenterEl.classList.remove("active");
    toolAlignRightEl.classList.remove("active");
    return;
  }

  ensureTextStyleDefaults(selected);
  essentialToolsEl.classList.remove("is-disabled");
  [toolDecFontEl, toolIncFontEl, toolFontSizeEl, toolBoldEl, toolItalicEl, toolUnderlineEl, toolAlignLeftEl, toolAlignCenterEl, toolAlignRightEl]
    .forEach((el) => {
      el.disabled = false;
    });

  toolFontSizeEl.value = selected.fontSize;
  toolBoldEl.classList.toggle("active", Number(selected.fontWeight) >= 700);
  toolItalicEl.classList.toggle("active", selected.fontStyle === "italic");
  toolUnderlineEl.classList.toggle("active", selected.textDecoration === "underline");
  toolAlignLeftEl.classList.toggle("active", selected.textAlign === "left");
  toolAlignCenterEl.classList.toggle("active", selected.textAlign === "center");
  toolAlignRightEl.classList.toggle("active", selected.textAlign === "right");
}

function updateSelectedTextStyle(updates) {
  const selected = getSelectedField();
  if (!selected || !supportsEssentialTextTools(selected)) {
    return;
  }

  ensureTextStyleDefaults(selected);
  Object.assign(selected, updates);

  selected.fontSize = clamp(Number(selected.fontSize) || 14, 8, 72);
  if (!["left", "center", "right"].includes(selected.textAlign)) {
    selected.textAlign = "left";
  }

  renderCanvas();
  renderLayerList();
  renderEssentialTools();
  scheduleDraftSave();
  scheduleHistoryRecord();
}

function nearestSnap(candidate, options, threshold = 7) {
  let snapped = candidate;
  let snappedGuide = null;

  options.forEach((point) => {
    const diff = Math.abs(point - candidate);
    if (diff <= threshold && (snappedGuide === null || diff < Math.abs(snappedGuide - candidate))) {
      snapped = point;
      snappedGuide = point;
    }
  });

  return { value: snapped, guide: snappedGuide };
}

function snapPosition(dragged, nextX, nextY, excludeIds = []) {
  const sheetWidth = sheetEl.clientWidth;
  const sheetHeight = sheetEl.clientHeight;

  let x = clamp(nextX, 0, sheetWidth - dragged.w);
  let y = clamp(nextY, 0, sheetHeight - dragged.h);

  const draggedPointsX = [x, x + dragged.w / 2, x + dragged.w];
  const draggedPointsY = [y, y + dragged.h / 2, y + dragged.h];

  const snapTargetsX = [0, sheetWidth / 2, sheetWidth];
  const snapTargetsY = [0, sheetHeight / 2, sheetHeight];

  state.fields
    .filter((f) => f.id !== dragged.id && !excludeIds.includes(f.id))
    .forEach((field) => {
      snapTargetsX.push(field.x, field.x + field.w / 2, field.x + field.w);
      snapTargetsY.push(field.y, field.y + field.h / 2, field.y + field.h);
    });

  const xChecks = [
    { edge: "left", point: draggedPointsX[0], offset: 0 },
    { edge: "center", point: draggedPointsX[1], offset: dragged.w / 2 },
    { edge: "right", point: draggedPointsX[2], offset: dragged.w }
  ];

  const yChecks = [
    { edge: "top", point: draggedPointsY[0], offset: 0 },
    { edge: "center", point: draggedPointsY[1], offset: dragged.h / 2 },
    { edge: "bottom", point: draggedPointsY[2], offset: dragged.h }
  ];

  let xGuide = null;
  let yGuide = null;

  xChecks.forEach((check) => {
    const snap = nearestSnap(check.point, snapTargetsX);
    if (snap.guide !== null) {
      x = snap.value - check.offset;
      xGuide = snap.guide;
    }
  });

  yChecks.forEach((check) => {
    const snap = nearestSnap(check.point, snapTargetsY);
    if (snap.guide !== null) {
      y = snap.value - check.offset;
      yGuide = snap.guide;
    }
  });

  x = clamp(x, 0, sheetWidth - dragged.w);
  y = clamp(y, 0, sheetHeight - dragged.h);

  if (xGuide !== null) {
    showVGuide(xGuide);
  } else {
    vGuideEl.classList.add("hidden");
  }

  if (yGuide !== null) {
    showHGuide(yGuide);
  } else {
    hGuideEl.classList.add("hidden");
  }

  return { x, y };
}

function addField(type, x = 22, y = 22) {
  const defaults = defaultByType[type];
  if (!defaults) {
    return;
  }

  const sheetWidth = sheetEl.clientWidth || 794;
  const sheetHeight = sheetEl.clientHeight || 1123;

  const newField = {
    id: uid(),
    type,
    ...defaults,
    locked: false,
    x: clamp(x, 0, sheetWidth - defaults.w),
    y: clamp(y, 0, sheetHeight - defaults.h)
  };

  ensureTextStyleDefaults(newField);

  state.fields.push(newField);
  selectOnly(newField.id);
  render();
  recordHistory();
}

function duplicateSelectedField() {
  const selectedFields = getSelectedFields();
  if (!selectedFields.length) {
    return;
  }

  const createdIds = [];
  selectedFields.forEach((selected) => {
    const next = {
      ...selected,
      id: uid(),
      locked: false,
      x: clamp(selected.x + 16, 0, Math.max(0, sheetEl.clientWidth - selected.w)),
      y: clamp(selected.y + 16, 0, Math.max(0, sheetEl.clientHeight - selected.h))
    };

    state.fields.push(next);
    createdIds.push(next.id);
  });

  state.selectedId = createdIds[0] || null;
  state.selectedIds = createdIds;
  render();
  recordHistory();
}

function clearSheetObjects() {
  if (!state.fields.length) {
    return;
  }

  const approved = window.confirm("Clear all objects from this sheet?");
  if (!approved) {
    return;
  }

  state.fields = [];
  state.selectedId = null;
  state.selectedIds = [];
  render();
  recordHistory();
}

function removeField(id) {
  state.fields = state.fields.filter((f) => f.id !== id);
  normalizeSelection();
  if (!state.selectedId) {
    state.selectedId = state.fields[0]?.id || null;
    state.selectedIds = state.selectedId ? [state.selectedId] : [];
  }
  render();
  recordHistory();
}

function removeSelectedFields() {
  const ids = new Set(state.selectedIds);
  if (!ids.size) {
    return;
  }

  state.fields = state.fields.filter((f) => !ids.has(f.id));
  state.selectedId = state.fields[0]?.id || null;
  state.selectedIds = state.selectedId ? [state.selectedId] : [];
  render();
  recordHistory();
}

function getSelectedField() {
  normalizeSelection();
  return state.fields.find((f) => f.id === state.selectedId) || null;
}

function bringSelectedToFront() {
  const selectedIds = new Set(state.selectedIds);
  if (!selectedIds.size) {
    return;
  }

  const keep = state.fields.filter((f) => !selectedIds.has(f.id));
  const moved = state.fields.filter((f) => selectedIds.has(f.id));
  state.fields = [...keep, ...moved];
  render();
  recordHistory();
}

function sendSelectedToBack() {
  const selectedIds = new Set(state.selectedIds);
  if (!selectedIds.size) {
    return;
  }

  const moved = state.fields.filter((f) => selectedIds.has(f.id));
  const keep = state.fields.filter((f) => !selectedIds.has(f.id));
  state.fields = [...moved, ...keep];
  render();
  recordHistory();
}

function toggleSelectedLock() {
  const selectedFields = getSelectedFields();
  if (!selectedFields.length) {
    return;
  }

  const anyUnlocked = selectedFields.some((field) => !field.locked);
  selectedFields.forEach((field) => {
    field.locked = anyUnlocked;
  });
  render();
  recordHistory();
}

function alignSelected(mode) {
  const selectedFields = getSelectedFields();
  if (!selectedFields.length) {
    return;
  }

  selectedFields.forEach((selected) => {
    const maxX = Math.max(0, sheetEl.clientWidth - selected.w);
    const maxY = Math.max(0, sheetEl.clientHeight - selected.h);

    if (mode === "left") {
      selected.x = 0;
    } else if (mode === "hcenter") {
      selected.x = Math.round(maxX / 2);
    } else if (mode === "right") {
      selected.x = maxX;
    } else if (mode === "top") {
      selected.y = 0;
    } else if (mode === "vcenter") {
      selected.y = Math.round(maxY / 2);
    } else if (mode === "bottom") {
      selected.y = maxY;
    }
  });

  render();
  recordHistory();
}

function applyInsertQuickAction(action) {
  const selected = getSelectedField();
  if (!selected || !isInsertObject(selected)) {
    return;
  }

  const sheetW = sheetEl.clientWidth;
  const sheetH = sheetEl.clientHeight;
  const ratio = selected.w > 0 && selected.h > 0 ? selected.w / selected.h : 1;

  if (action === "fill-sheet") {
    selected.x = 0;
    selected.y = 0;
    selected.w = sheetW;
    selected.h = sheetH;
  } else if (action === "fit-width") {
    selected.x = 0;
    selected.w = sheetW;
  } else if (action === "fit-height") {
    selected.y = 0;
    selected.h = sheetH;
  } else if (action === "center-sheet") {
    selected.x = Math.round((sheetW - selected.w) / 2);
    selected.y = Math.round((sheetH - selected.h) / 2);
  } else if (action === "cover-sheet" && selected.type === "image") {
    let nextW = sheetW;
    let nextH = Math.round(nextW / ratio);
    if (nextH < sheetH) {
      nextH = sheetH;
      nextW = Math.round(nextH * ratio);
    }
    selected.w = nextW;
    selected.h = nextH;
    selected.x = Math.round((sheetW - selected.w) / 2);
    selected.y = Math.round((sheetH - selected.h) / 2);
  } else if (action === "contain-sheet" && selected.type === "image") {
    let nextW = sheetW;
    let nextH = Math.round(nextW / ratio);
    if (nextH > sheetH) {
      nextH = sheetH;
      nextW = Math.round(nextH * ratio);
    }
    selected.w = nextW;
    selected.h = nextH;
    selected.x = Math.round((sheetW - selected.w) / 2);
    selected.y = Math.round((sheetH - selected.h) / 2);
  }

  selected.w = clamp(Math.round(selected.w), 40, sheetW);
  selected.h = clamp(Math.round(selected.h), 40, sheetH);
  selected.x = clamp(Math.round(selected.x), 0, Math.max(0, sheetW - selected.w));
  selected.y = clamp(Math.round(selected.y), 0, Math.max(0, sheetH - selected.h));

  render();
  recordHistory();
}

function previewFor(field) {
  switch (field.type) {
    case "table":
      return `${Math.max(1, Number(field.rows) || 3)} x ${Math.max(1, Number(field.cols) || 3)} table`;
    case "image":
      return field.imageSrc ? "Inserted picture" : "Upload picture from Properties";
    case "shaperect":
      return "Rectangle shape";
    case "shapecircle":
      return "Circle shape";
    case "shapeline":
      return "Line element";
    case "elementstar":
      return "Star element";
    case "elementarrow":
      return "Arrow element";
    case "constant":
    case "sectionheader":
      return field.value || "Text";
    case "text":
    case "email":
    case "phone":
    case "url":
    case "number":
    case "currency":
    case "percentage":
    case "date":
    case "time":
    case "datetime":
      return field.placeholder || "Input field";
    case "textarea":
      return field.placeholder || "Multi-line text";
    case "select":
      return `Dropdown: ${(field.options || "").split(",").slice(0, 2).join(", ")}...`;
    case "radio":
      return `Radio: ${(field.options || "").split(",").slice(0, 2).join(", ")}...`;
    case "checkbox":
      return `☑ ${(field.options || "Check")}`;
    case "multiselect":
      return `Multiple: ${(field.options || "").split(",").slice(0, 2).join(", ")}...`;
    case "file":
      return `📎 Files (${field.maxFiles || 1} max)`;
    case "photo":
      return `🖼 Photos (${field.maxFiles || 1} max)`;
    case "signature":
      return "✍ Signature field";
    case "rating":
      return `★ Rating (max ${field.maxRating || 5})`;
    case "divider":
      return "──────────────";
    default:
      return "Field";
  }
}

function getTableMatrix(field) {
  const rows = clamp(Math.round(Number(field.rows) || 3), 1, 12);
  const cols = clamp(Math.round(Number(field.cols) || 3), 1, 12);
  const raw = String(field.cellLabels || "").trim();

  if (!raw) {
    return Array.from({ length: rows }, () => Array.from({ length: cols }, () => ""));
  }

  const lines = raw.split(/\r?\n/).map((line) => line.trim());
  return Array.from({ length: rows }, (_, rowIndex) => {
    const rowLine = lines[rowIndex] || "";
    const parts = rowLine.split("|").map((part) => part.trim());
    return Array.from({ length: cols }, (_, colIndex) => parts[colIndex] || "");
  });
}

function parseTableTrackList(raw, count) {
  const tokens = String(raw || "")
    .split(",")
    .map((part) => Number(part.trim()))
    .filter((value) => Number.isFinite(value) && value > 0);

  return Array.from({ length: count }, (_, idx) => tokens[idx] || 1);
}

function getTableTrackTemplates(field) {
  const rows = clamp(Math.round(Number(field.rows) || 3), 1, 12);
  const cols = clamp(Math.round(Number(field.cols) || 3), 1, 12);
  const colTracks = parseTableTrackList(field.colWidths, cols);
  const rowTracks = parseTableTrackList(field.rowHeights, rows);

  return {
    colsTemplate: colTracks.map((value) => `${value}fr`).join(" "),
    rowsTemplate: rowTracks.map((value) => `${value}fr`).join(" ")
  };
}

function renderObjectBody(field, labelStyle, previewStyle) {
  if (field.type === "table") {
    const matrix = getTableMatrix(field);
    const trackTemplates = getTableTrackTemplates(field);
    const borderColor = escapeHtml(field.borderColor || "#5e75b8");
    const borderWidth = Math.max(1, Number(field.borderWidth) || 1);
    const cellBg = escapeHtml(field.cellBg || "#ffffff");
    const textColor = escapeHtml(field.textColor || "#24366b");

    const cells = matrix
      .flatMap((row) => row)
      .map((cell) => `<div class="object-table-cell" style="background:${cellBg};color:${textColor};border:${borderWidth}px solid ${borderColor};${labelStyle}">${escapeHtml(cell || " ")}</div>`)
      .join("");

    return `<div class="object-table-grid" style="grid-template-columns:${trackTemplates.colsTemplate};grid-template-rows:${trackTemplates.rowsTemplate};">${cells}</div>`;
  }

  if (field.type === "image") {
    if (field.imageSrc) {
      return `<div class="object-insert object-image" style="border-radius:${Number(field.radius) || 0}px;"><img src="${escapeHtml(field.imageSrc)}" alt="Inserted" style="object-fit:${escapeHtml(field.fit || "cover")};" /></div>`;
    }
    return `<div class="object-insert object-image object-image-empty" style="border-radius:${Number(field.radius) || 0}px;">Picture placeholder</div>`;
  }

  if (field.type === "shaperect") {
    return `<div class="object-insert shape-rect" style="background:${escapeHtml(field.fillColor || "#dff1eb")};border:${Number(field.strokeWidth) || 1}px solid ${escapeHtml(field.strokeColor || "#3b7f71")};border-radius:${Number(field.radius) || 0}px;"></div>`;
  }

  if (field.type === "shapecircle") {
    return `<div class="object-insert shape-circle" style="background:${escapeHtml(field.fillColor || "#f8ead9")};border:${Number(field.strokeWidth) || 1}px solid ${escapeHtml(field.strokeColor || "#ba7a36")};"></div>`;
  }

  if (field.type === "shapeline") {
    return `<div class="object-insert shape-line" style="background:${escapeHtml(field.strokeColor || "#4f6f67")};height:${Math.max(1, Number(field.strokeWidth) || 2)}px;"></div>`;
  }

  if (field.type === "elementstar") {
    return `<div class="object-insert element-star" style="color:${escapeHtml(field.fillColor || "#f3ca62")};">★</div>`;
  }

  if (field.type === "elementarrow") {
    return `<div class="object-insert element-arrow" style="color:${escapeHtml(field.fillColor || "#93b9ad")};">➜</div>`;
  }

  return `
      <div class="object-meta">
        <span>${escapeHtml(field.type.toUpperCase())}</span>
        <span>${field.locked ? "Locked" : (field.required ? "Required" : "Optional")}</span>
      </div>
      <div class="object-label" style="${labelStyle}">${escapeHtml(field.label || "Untitled")}</div>
      <div class="object-preview" style="${previewStyle}">${escapeHtml(previewFor(field))}</div>
  `;
}

function renderCanvas() {
  normalizeSelection();
  fieldCountEl.textContent = `${state.fields.length} objects`;
  sheetEl.innerHTML = "";

  if (!state.fields.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.style.position = "absolute";
    empty.style.inset = "12px";
    empty.style.display = "grid";
    empty.style.placeItems = "center";
    empty.textContent = "Drag blocks from left panel and drop on the sheet.";
    sheetEl.appendChild(empty);
    return;
  }

  state.fields.forEach((field) => {
    ensureTextStyleDefaults(field);
    const selected = isSelected(field.id);
    const item = document.createElement("article");
    item.className = `canvas-object ${selected ? "selected" : ""} ${field.locked ? "locked" : ""}`;
    item.style.left = `${field.x}px`;
    item.style.top = `${field.y}px`;
    item.style.width = `${field.w}px`;
    item.style.height = `${field.h}px`;
    item.style.opacity = `${clamp((Number(field.opacity) || 100) / 100, 0.05, 1)}`;
    item.style.transform = `rotate(${clamp(Number(field.rotate) || 0, -180, 180)}deg)`;
    item.style.transformOrigin = "center center";
    item.dataset.id = field.id;

    const labelStyle = `font-size:${clamp(Number(field.fontSize) || 14, 8, 72)}px;font-weight:${field.fontWeight};font-style:${field.fontStyle};text-decoration:${field.textDecoration};text-align:${field.textAlign};`;
    const previewStyle = `font-size:${Math.max(11, (Number(field.fontSize) || 14) - 2)}px;text-align:${field.textAlign};`;

    item.innerHTML = `
      ${renderObjectBody(field, labelStyle, previewStyle)}
      ${state.selectedIds.length === 1 && selected ? `
        <div class="resize-handle tl" data-handle="tl"></div>
        <div class="resize-handle tr" data-handle="tr"></div>
        <div class="resize-handle bl" data-handle="bl"></div>
        <div class="resize-handle br" data-handle="br"></div>
        <div class="resize-handle tc" data-handle="tc"></div>
        <div class="resize-handle bc" data-handle="bc"></div>
        <div class="resize-handle lc" data-handle="lc"></div>
        <div class="resize-handle rc" data-handle="rc"></div>
      ` : ""}
    `;

    item.addEventListener("pointerdown", (event) => {
      if (event.shiftKey || event.ctrlKey || event.metaKey) {
        return;
      }
      const handle = event.target.closest(".resize-handle");
      if (handle) {
        event.stopPropagation();
        startResizeField(event, field.id, handle.dataset.handle);
      } else {
        startDragField(event, field.id);
      }
    });
    item.addEventListener("click", (event) => {
      if (suppressNextSelectionClick) {
        suppressNextSelectionClick = false;
        return;
      }
      if (event.shiftKey || event.ctrlKey || event.metaKey) {
        toggleSelection(field.id);
      } else {
        selectOnly(field.id);
      }
      render();
    });

    sheetEl.appendChild(item);
  });
}

function renderLayerList() {
  if (!state.fields.length) {
    layerListEl.innerHTML = "<div class='empty'>No layers yet.</div>";
    return;
  }

  const rows = [...state.fields]
    .map((field, idx) => ({ field, z: idx + 1 }))
    .reverse()
    .map(({ field, z }) => {
      const selectedClass = isSelected(field.id) ? "selected" : "";
      const locked = field.locked ? "locked" : "";
      return `
        <button type="button" class="layer-item ${selectedClass} ${locked}" data-layer-id="${field.id}">
          <span class="layer-name">${escapeHtml(field.label || field.type)}</span>
          <span class="layer-meta">Z${z} ${field.locked ? "| Locked" : ""}</span>
        </button>
      `;
    })
    .join("");

  layerListEl.innerHTML = rows;
  layerListEl.querySelectorAll("[data-layer-id]").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      if (event.shiftKey || event.ctrlKey || event.metaKey) {
        toggleSelection(btn.dataset.layerId);
      } else {
        selectOnly(btn.dataset.layerId);
      }
      render();
    });
  });
}

function renderProperties() {
  if (state.selectedIds.length > 1) {
    propertyFormEl.innerHTML = `<div class='empty'>${state.selectedIds.length} objects selected. Drag any selected object to move all.</div>`;
    return;
  }

  const selected = state.fields.find((f) => f.id === state.selectedId);
  if (!selected) {
    propertyFormEl.innerHTML = "<div class='empty'>No object selected.</div>";
    return;
  }

  const isInsert = isInsertObject(selected);

  const commonRows = `
    <div class="form-row">
      <label for="fLabel">Label</label>
      <input id="fLabel" name="label" value="${escapeHtml(selected.label || "")}" />
    </div>
    ${isInsert ? "" : `
      <label class="check-row" for="fRequired">
        <input id="fRequired" name="required" type="checkbox" ${selected.required ? "checked" : ""} />
        Required field
      </label>
    `}
    <label class="check-row" for="fLocked">
      <input id="fLocked" name="locked" type="checkbox" ${selected.locked ? "checked" : ""} />
      Lock layer position
    </label>
    <div class="form-row">
      <label for="fX">X Position</label>
      <input id="fX" name="x" type="number" min="0" value="${selected.x}" />
    </div>
    <div class="form-row">
      <label for="fY">Y Position</label>
      <input id="fY" name="y" type="number" min="0" value="${selected.y}" />
    </div>
    <div class="form-row">
      <label for="fW">Width</label>
      <input id="fW" name="w" type="number" min="70" value="${selected.w}" />
    </div>
    <div class="form-row">
      <label for="fH">Height</label>
      <input id="fH" name="h" type="number" min="40" value="${selected.h}" />
    </div>
  `;

  let typeRows = "";

  // Placeholder fields
  if (["text", "email", "phone", "url", "number", "currency", "percentage", "textarea", "date", "time", "datetime"].includes(selected.type)) {
    typeRows += `
      <div class="form-row">
        <label for="fPlaceholder">Placeholder</label>
        <input id="fPlaceholder" name="placeholder" value="${escapeHtml(selected.placeholder || "")}" />
      </div>
    `;
  }

  // Constant/Section header value
  if (["constant", "sectionheader"].includes(selected.type)) {
    typeRows += `
      <div class="form-row">
        <label for="fValue">Text Value</label>
        <textarea id="fValue" name="value" rows="3">${escapeHtml(selected.value || "")}</textarea>
      </div>
    `;
  }

  if (selected.type === "table") {
    typeRows += `
      <div class="form-row">
        <label for="fRows">Rows</label>
        <input id="fRows" name="rows" type="number" min="1" max="12" value="${clamp(Math.round(Number(selected.rows) || 3), 1, 12)}" />
      </div>
      <div class="form-row">
        <label for="fCols">Columns</label>
        <input id="fCols" name="cols" type="number" min="1" max="12" value="${clamp(Math.round(Number(selected.cols) || 3), 1, 12)}" />
      </div>
      <div class="form-row">
        <label for="fColWidths">Column Widths (ratio)</label>
        <input id="fColWidths" name="colWidths" value="${escapeHtml(selected.colWidths || "")}" placeholder="1,1,2" />
      </div>
      <div class="form-row">
        <label for="fRowHeights">Row Heights (ratio)</label>
        <input id="fRowHeights" name="rowHeights" value="${escapeHtml(selected.rowHeights || "")}" placeholder="1,2,1" />
      </div>
      <div class="form-row">
        <label for="fCellMode">Cell Mode</label>
        <select id="fCellMode" name="cellMode">
          <option value="text" ${selected.cellMode === "text" ? "selected" : ""}>Static Text</option>
          <option value="input" ${selected.cellMode === "input" ? "selected" : ""}>Input Cells</option>
        </select>
      </div>
      <div class="form-row">
        <label for="fCellLabels">Cell Text (one row per line, use | between cells)</label>
        <textarea id="fCellLabels" name="cellLabels" rows="6" placeholder="Name | Age | City\nArun | 27 | Chennai">${escapeHtml(selected.cellLabels || "")}</textarea>
      </div>
      <div class="form-row">
        <label for="fBorderColor">Border Color</label>
        <input id="fBorderColor" name="borderColor" type="color" value="${escapeHtml(selected.borderColor || "#5e75b8")}" />
      </div>
      <div class="form-row">
        <label for="fBorderWidth">Border Width</label>
        <input id="fBorderWidth" name="borderWidth" type="number" min="1" max="8" value="${Math.max(1, Number(selected.borderWidth) || 1)}" />
      </div>
      <div class="form-row">
        <label for="fCellBg">Cell Fill</label>
        <input id="fCellBg" name="cellBg" type="color" value="${escapeHtml(selected.cellBg || "#ffffff")}" />
      </div>
      <div class="form-row">
        <label for="fTextColor">Text Color</label>
        <input id="fTextColor" name="textColor" type="color" value="${escapeHtml(selected.textColor || "#24366b")}" />
      </div>
    `;
  }

  // Options-based fields (select, radio, checkbox, multiselect)
  if (["select", "radio", "checkbox", "multiselect"].includes(selected.type)) {
    typeRows += `
      <div class="form-row">
        <label for="fOptions">Options (comma separated)</label>
        <textarea id="fOptions" name="options" rows="3">${escapeHtml(selected.options || "")}</textarea>
      </div>
    `;
  }

  // Currency specific
  if (selected.type === "currency") {
    typeRows += `
      <div class="form-row">
        <label for="fCurrency">Currency Code</label>
        <input id="fCurrency" name="currency" value="${escapeHtml(selected.currency || "USD")}" />
      </div>
    `;
  }

  // File upload options
  if (["file", "photo"].includes(selected.type)) {
    typeRows += `
      <div class="form-row">
        <label for="fAccept">Accept</label>
        <input id="fAccept" name="accept" value="${escapeHtml(selected.accept || (selected.type === "photo" ? "image/*" : ".pdf,.doc,.docx"))}" />
      </div>
      <div class="form-row">
        <label for="fMaxFiles">Max files</label>
        <input id="fMaxFiles" name="maxFiles" type="number" min="1" max="20" value="${selected.maxFiles || 1}" />
      </div>
    `;
  }

  // Picture insert options
  if (selected.type === "image") {
    typeRows += `
      <div class="form-row">
        <label for="fImageUpload">Insert Picture</label>
        <input id="fImageUpload" type="file" accept="image/*" />
      </div>
      <div class="form-row">
        <label for="fFit">Fit Mode</label>
        <select id="fFit" name="fit">
          <option value="cover" ${selected.fit === "cover" ? "selected" : ""}>Cover</option>
          <option value="contain" ${selected.fit === "contain" ? "selected" : ""}>Contain</option>
          <option value="fill" ${selected.fit === "fill" ? "selected" : ""}>Fill</option>
        </select>
      </div>
      <div class="form-row">
        <label for="fRadius">Corner Radius</label>
        <input id="fRadius" name="radius" type="number" min="0" max="120" value="${Number(selected.radius) || 0}" />
      </div>
    `;
  }

  // Shape/element style options
  if (["shaperect", "shapecircle", "shapeline", "elementstar", "elementarrow"].includes(selected.type)) {
    typeRows += `
      <div class="form-row">
        <label for="fFillColor">Fill Color</label>
        <input id="fFillColor" name="fillColor" type="color" value="${escapeHtml(selected.fillColor || "#dff1eb")}" />
      </div>
    `;
  }

  if (["shaperect", "shapecircle", "shapeline"].includes(selected.type)) {
    typeRows += `
      <div class="form-row">
        <label for="fStrokeColor">Stroke Color</label>
        <input id="fStrokeColor" name="strokeColor" type="color" value="${escapeHtml(selected.strokeColor || "#3b7f71")}" />
      </div>
      <div class="form-row">
        <label for="fStrokeWidth">Stroke Width</label>
        <input id="fStrokeWidth" name="strokeWidth" type="number" min="1" max="12" value="${Number(selected.strokeWidth) || 1}" />
      </div>
    `;
  }

  if (selected.type === "shaperect") {
    typeRows += `
      <div class="form-row">
        <label for="fRadius">Corner Radius</label>
        <input id="fRadius" name="radius" type="number" min="0" max="80" value="${Number(selected.radius) || 0}" />
      </div>
    `;
  }

  // Rating specific
  if (selected.type === "rating") {
    typeRows += `
      <div class="form-row">
        <label for="fMaxRating">Max Rating</label>
        <input id="fMaxRating" name="maxRating" type="number" min="1" max="10" value="${selected.maxRating || 5}" />
      </div>
    `;
  }

  if (isInsert) {
    typeRows += `
      <div class="form-row">
        <label for="fOpacity">Opacity (%)</label>
        <input id="fOpacity" name="opacity" type="range" min="5" max="100" value="${clamp(Number(selected.opacity) || 100, 5, 100)}" />
      </div>
      <div class="form-row">
        <label for="fRotate">Rotation (deg)</label>
        <input id="fRotate" name="rotate" type="number" min="-180" max="180" value="${clamp(Number(selected.rotate) || 0, -180, 180)}" />
      </div>
      <div class="form-row">
        <label>Insert Quick Actions</label>
        <div class="align-grid">
          <button class="btn btn-ghost" type="button" data-insert-action="fill-sheet">Fill Sheet</button>
          <button class="btn btn-ghost" type="button" data-insert-action="fit-width">Fit Width</button>
          <button class="btn btn-ghost" type="button" data-insert-action="fit-height">Fit Height</button>
          <button class="btn btn-ghost" type="button" data-insert-action="center-sheet">Center</button>
          ${selected.type === "image" ? `
            <button class="btn btn-ghost" type="button" data-insert-action="cover-sheet">Cover Sheet</button>
            <button class="btn btn-ghost" type="button" data-insert-action="contain-sheet">Contain Sheet</button>
          ` : ""}
        </div>
      </div>
    `;
  }

  propertyFormEl.innerHTML = `
    ${commonRows}
    ${typeRows}
    <div class="form-row">
      <label>Quick Align</label>
      <div class="align-grid">
        <button class="btn btn-ghost" type="button" data-align="left">Left</button>
        <button class="btn btn-ghost" type="button" data-align="hcenter">Center</button>
        <button class="btn btn-ghost" type="button" data-align="right">Right</button>
        <button class="btn btn-ghost" type="button" data-align="top">Top</button>
        <button class="btn btn-ghost" type="button" data-align="vcenter">Middle</button>
        <button class="btn btn-ghost" type="button" data-align="bottom">Bottom</button>
      </div>
    </div>
    <button class="btn btn-danger" id="btnDeleteField" type="button">Delete Object</button>
  `;

  propertyFormEl.querySelector("#btnDeleteField").addEventListener("click", () => removeField(selected.id));
  propertyFormEl.querySelectorAll("[data-align]").forEach((btn) => {
    btn.addEventListener("click", () => {
      alignSelected(btn.dataset.align);
    });
  });
  propertyFormEl.querySelectorAll("[data-insert-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      applyInsertQuickAction(btn.dataset.insertAction);
    });
  });

  propertyFormEl.querySelectorAll("input, textarea, select").forEach((input) => {
    input.addEventListener("input", (event) => {
      const { name, type, value, checked } = event.target;
      if (!name || type === "file") {
        return;
      }
      const target = state.fields.find((f) => f.id === selected.id);
      if (!target) {
        return;
      }

      if (type === "checkbox") {
        target[name] = checked;
      } else if (["x", "y", "w", "h", "maxFiles", "maxRating", "strokeWidth", "radius", "rotate", "opacity", "rows", "cols", "borderWidth"].includes(name)) {
        target[name] = Number(value) || 0;
      } else {
        target[name] = value;
      }

      if (typeof target.rows === "number") {
        target.rows = clamp(Math.round(target.rows), 1, 12);
      }
      if (typeof target.cols === "number") {
        target.cols = clamp(Math.round(target.cols), 1, 12);
      }
      if (typeof target.colWidths === "string") {
        target.colWidths = target.colWidths
          .split(",")
          .map((part) => part.trim())
          .filter(Boolean)
          .slice(0, clamp(Math.round(Number(target.cols) || 3), 1, 12))
          .join(",");
      }
      if (typeof target.rowHeights === "string") {
        target.rowHeights = target.rowHeights
          .split(",")
          .map((part) => part.trim())
          .filter(Boolean)
          .slice(0, clamp(Math.round(Number(target.rows) || 3), 1, 12))
          .join(",");
      }

      if (typeof target.opacity === "number") {
        target.opacity = clamp(target.opacity, 5, 100);
      }
      if (typeof target.rotate === "number") {
        target.rotate = clamp(target.rotate, -180, 180);
      }
      target.w = clamp(target.w, 70, sheetEl.clientWidth);
      target.h = clamp(target.h, 40, sheetEl.clientHeight);
      target.x = clamp(target.x, 0, Math.max(0, sheetEl.clientWidth - target.w));
      target.y = clamp(target.y, 0, Math.max(0, sheetEl.clientHeight - target.h));

      renderCanvas();
      renderLayerList();
      scheduleDraftSave();
      scheduleHistoryRecord();
    });
  });

  const imageUpload = propertyFormEl.querySelector("#fImageUpload");
  if (imageUpload) {
    imageUpload.addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) {
        return;
      }
      const target = state.fields.find((f) => f.id === selected.id);
      if (!target) {
        return;
      }

      const reader = new FileReader();
      reader.onload = () => {
        target.imageSrc = String(reader.result || "");
        renderCanvas();
        renderLayerList();
        scheduleDraftSave();
        recordHistory();
      };
      reader.readAsDataURL(file);
    });
  }
}

function render() {
  applySheetSize();
  renderCanvas();
  renderProperties();
  renderLayerList();
  renderEssentialTools();
  updateHistoryButtons();
  updateLayerButtons();
  scheduleDraftSave();
}

function buildTemplatePayload() {
  return {
    templateName: templateNameEl.value,
    version: "0.4.0",
    page: {
      size: state.page.size,
      orientation: state.page.orientation,
      standards: Object.keys(pageFormats)
    },
    fields: state.fields.map((f, order) => ({ ...f, zIndex: order + 1 }))
  };
}

function exportTemplate() {
  const payload = buildTemplatePayload();
  jsonOutputEl.textContent = JSON.stringify(payload, null, 2);
  localStorage.setItem(templateStorageKey, JSON.stringify(payload));
  jsonDialogEl.showModal();
}

function saveDraft() {
  const payload = buildTemplatePayload();
  const text = JSON.stringify(payload);
  localStorage.setItem(draftStorageKey, text);
  localStorage.setItem(templateStorageKey, text);
}

function scheduleDraftSave() {
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(saveDraft, 180);
}

function restoreFromPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return false;
  }

  const restoredName = (payload.templateName || "Untitled Template").toString();
  const restoredPage = payload.page || {};
  const restoredFields = Array.isArray(payload.fields) ? payload.fields : [];

  state.page.size = pageFormats[restoredPage.size] ? restoredPage.size : "A4";
  state.page.orientation = restoredPage.orientation === "landscape" ? "landscape" : "portrait";
  state.fields = restoredFields
    .filter((f) => defaultByType[f.type])
    .map((field) => ({
      ...defaultByType[field.type],
      ...field,
      locked: Boolean(field.locked),
      id: field.id || uid()
    }));
  state.fields.forEach((field) => ensureTextStyleDefaults(field));
  state.selectedId = state.fields[0]?.id || null;
  state.selectedIds = state.selectedId ? [state.selectedId] : [];

  templateNameEl.value = restoredName;
  setupNameEl.value = restoredName;
  pageSizeSelectEl.value = state.page.size;
  setupSizeEl.value = state.page.size;
  orientationSelectEl.value = state.page.orientation;
  setupOrientationEl.value = state.page.orientation;
  return true;
}

function tryRestoreDraft() {
  const raw = localStorage.getItem(draftStorageKey) || localStorage.getItem(templateStorageKey);
  if (!raw) {
    return false;
  }

  try {
    const parsed = JSON.parse(raw);
    return restoreFromPayload(parsed);
  } catch {
    return false;
  }
}

function downloadTemplate() {
  const payload = buildTemplatePayload();
  localStorage.setItem(templateStorageKey, JSON.stringify(payload));
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const safeName = (payload.templateName || "template").trim().replace(/[^a-z0-9]+/gi, "-").replace(/(^-|-$)/g, "").toLowerCase();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${safeName || "template"}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function openFillerPage() {
  const payload = buildTemplatePayload();
  localStorage.setItem(templateStorageKey, JSON.stringify(payload));
  window.location.href = "filler.html?source=builder";
}

function buildFormatSelectors() {
  const options = Object.entries(pageFormats)
    .map(([name, def]) => `<option value="${name}">${name} (${def.w}x${def.h} mm)</option>`)
    .join("");

  setupSizeEl.innerHTML = options;
  pageSizeSelectEl.innerHTML = options;
  setupSizeEl.value = state.page.size;
  pageSizeSelectEl.value = state.page.size;
  setupOrientationEl.value = state.page.orientation;
  orientationSelectEl.value = state.page.orientation;
}

function buildPalette() {
  paletteEl.innerHTML = "";
  const groups = new Map();
  fieldCatalog.forEach((item) => {
    const key = item.group || "Other";
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(item);
  });

  groups.forEach((items, groupName) => {
    const section = document.createElement("section");
    section.className = "palette-section";

    const heading = document.createElement("h4");
    heading.className = "palette-section-title";
    heading.textContent = groupName;
    section.appendChild(heading);

    const list = document.createElement("div");
    list.className = "palette-group";

    items.forEach((item) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn palette-item";
      btn.draggable = true;
      btn.innerHTML = `<span class="palette-icon">${item.icon}</span><span>${item.label}</span>`;

      btn.addEventListener("click", () => addField(item.type, 24, 24));
      btn.addEventListener("dragstart", (event) => {
        event.dataTransfer?.setData("text/plain", item.type);
        event.dataTransfer.effectAllowed = "copy";
      });

      list.appendChild(btn);
    });

    section.appendChild(list);
    paletteEl.appendChild(section);
  });
}

function startDragField(event, fieldId) {
  const field = state.fields.find((f) => f.id === fieldId);
  if (!field || field.locked || event.button !== 0) {
    return;
  }

  if (!isSelected(field.id)) {
    selectOnly(field.id);
  }

  const selectedFields = getSelectedFields().filter((f) => !f.locked);
  if (!selectedFields.length) {
    return;
  }

  if (selectedFields.length !== state.selectedIds.length) {
    state.selectedIds = selectedFields.map((f) => f.id);
    state.selectedId = selectedFields[0]?.id || null;
  }

  const rect = getSheetRect();
  const ids = selectedFields.map((f) => f.id);
  const groupStarts = {};
  selectedFields.forEach((f) => {
    groupStarts[f.id] = { x: f.x, y: f.y };
  });

  const bounds = {
    minX: Math.min(...selectedFields.map((f) => f.x)),
    minY: Math.min(...selectedFields.map((f) => f.y)),
    maxX: Math.max(...selectedFields.map((f) => f.x + f.w)),
    maxY: Math.max(...selectedFields.map((f) => f.y + f.h))
  };

  state.dragging = {
    id: field.id,
    ids,
    pointerId: event.pointerId,
    offsetX: event.clientX - rect.left - field.x,
    offsetY: event.clientY - rect.top - field.y,
    startX: field.x,
    startY: field.y,
    groupStarts,
    groupBounds: bounds,
    moved: false
  };

  event.currentTarget.setPointerCapture(event.pointerId);
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
  render();
}

function onPointerMove(event) {
  if (!state.dragging) {
    return;
  }

  const field = state.fields.find((f) => f.id === state.dragging.id);
  if (!field) {
    return;
  }

  const rect = getSheetRect();
  const proposedX = event.clientX - rect.left - state.dragging.offsetX;
  const proposedY = event.clientY - rect.top - state.dragging.offsetY;
  const snapped = snapPosition(field, proposedX, proposedY, state.dragging.ids || []);

  let dx = snapped.x - state.dragging.startX;
  let dy = snapped.y - state.dragging.startY;
  const bounds = state.dragging.groupBounds;
  const sheetW = sheetEl.clientWidth;
  const sheetH = sheetEl.clientHeight;
  dx = clamp(dx, -bounds.minX, sheetW - bounds.maxX);
  dy = clamp(dy, -bounds.minY, sheetH - bounds.maxY);

  (state.dragging.ids || [field.id]).forEach((id) => {
    const target = state.fields.find((f) => f.id === id);
    const start = state.dragging.groupStarts?.[id];
    if (!target || !start) {
      return;
    }
    target.x = Math.round(start.x + dx);
    target.y = Math.round(start.y + dy);
  });

  showMeasureGuides(field);
  state.dragging.moved = state.dragging.moved || dx !== 0 || dy !== 0;
  renderCanvas();
  renderLayerList();
  renderProperties();
}

function onPointerUp() {
  const shouldRecord = Boolean(state.dragging?.moved);
  if (shouldRecord) {
    suppressNextSelectionClick = true;
  }
  state.dragging = null;
  hideGuides();
  window.removeEventListener("pointermove", onPointerMove);
  window.removeEventListener("pointerup", onPointerUp);

  if (shouldRecord) {
    scheduleDraftSave();
    recordHistory();
  }
}

function startResizeField(event, fieldId, handleType) {
  const field = state.fields.find((f) => f.id === fieldId);
  if (!field || field.locked) {
    return;
  }

  selectOnly(fieldId);

  state.dragging = {
    id: fieldId,
    pointerId: event.pointerId,
    type: "resize",
    handle: handleType,
    startX: field.x,
    startY: field.y,
    startW: field.w,
    startH: field.h,
    startClientX: event.clientX,
    startClientY: event.clientY,
    moved: false
  };

  event.currentTarget.setPointerCapture(event.pointerId);
  window.addEventListener("pointermove", onResizePointerMove);
  window.addEventListener("pointerup", onResizePointerUp);
}

function onResizePointerMove(event) {
  if (!state.dragging || state.dragging.type !== "resize") {
    return;
  }

  const field = state.fields.find((f) => f.id === state.dragging.id);
  if (!field) {
    return;
  }

  const deltaX = event.clientX - state.dragging.startClientX;
  const deltaY = event.clientY - state.dragging.startClientY;
  const minSize = 40;
  const handle = state.dragging.handle;

  let newX = state.dragging.startX;
  let newY = state.dragging.startY;
  let newW = state.dragging.startW;
  let newH = state.dragging.startH;

  // Corner handles
  if (handle === "tl") {
    newX = state.dragging.startX + deltaX;
    newY = state.dragging.startY + deltaY;
    newW = state.dragging.startW - deltaX;
    newH = state.dragging.startH - deltaY;
  } else if (handle === "tr") {
    newY = state.dragging.startY + deltaY;
    newW = state.dragging.startW + deltaX;
    newH = state.dragging.startH - deltaY;
  } else if (handle === "bl") {
    newX = state.dragging.startX + deltaX;
    newW = state.dragging.startW - deltaX;
    newH = state.dragging.startH + deltaY;
  } else if (handle === "br") {
    newW = state.dragging.startW + deltaX;
    newH = state.dragging.startH + deltaY;
  }
  // Edge handles
  else if (handle === "tc") {
    newY = state.dragging.startY + deltaY;
    newH = state.dragging.startH - deltaY;
  } else if (handle === "bc") {
    newH = state.dragging.startH + deltaY;
  } else if (handle === "lc") {
    newX = state.dragging.startX + deltaX;
    newW = state.dragging.startW - deltaX;
  } else if (handle === "rc") {
    newW = state.dragging.startW + deltaX;
  }

  // Enforce minimum size
  if (newW < minSize) {
    if (handle.includes("l")) {
      newX = state.dragging.startX + (state.dragging.startW - minSize);
    }
    newW = minSize;
  }
  if (newH < minSize) {
    if (handle.includes("t")) {
      newY = state.dragging.startY + (state.dragging.startH - minSize);
    }
    newH = minSize;
  }

  const sheetW = sheetEl.clientWidth;
  const sheetH = sheetEl.clientHeight;

  field.x = clamp(Math.round(newX), 0, Math.max(0, sheetW - minSize));
  field.y = clamp(Math.round(newY), 0, Math.max(0, sheetH - minSize));
  field.w = clamp(Math.round(newW), minSize, Math.max(minSize, sheetW - field.x));
  field.h = clamp(Math.round(newH), minSize, Math.max(minSize, sheetH - field.y));
  showMeasureGuides(field);
  state.dragging.moved = true;

  renderCanvas();
  renderLayerList();
  renderProperties();
}

function onResizePointerUp() {
  const shouldRecord = Boolean(state.dragging?.moved);
  if (shouldRecord) {
    suppressNextSelectionClick = true;
  }
  state.dragging = null;
  hideGuides();
  window.removeEventListener("pointermove", onResizePointerMove);
  window.removeEventListener("pointerup", onResizePointerUp);

  if (shouldRecord) {
    scheduleDraftSave();
    recordHistory();
  }
}

function initDropZone() {
  sheetEl.addEventListener("dragover", (event) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  });

  sheetEl.addEventListener("drop", (event) => {
    event.preventDefault();
    const type = event.dataTransfer?.getData("text/plain");
    if (!type || !defaultByType[type]) {
      return;
    }

    const rect = getSheetRect();
    const defaults = defaultByType[type];
    const x = event.clientX - rect.left - defaults.w / 2;
    const y = event.clientY - rect.top - 18;
    addField(type, x, y);
  });
}

function initSplitPaneDraggers() {
  // Setup divider1 (between palette and canvas)
  divider1El.addEventListener("mousedown", (event) => startSplitDrag(event, 1));
  
  // Setup divider2 (between canvas and config)
  divider2El.addEventListener("mousedown", (event) => startSplitDrag(event, 2));
}

function startSplitDrag(event, dividerId) {
  event.preventDefault();
  
  const startX = event.clientX;
  const startLeftWidth = document.querySelector(".col-palette").getBoundingClientRect().width;
  const startRightWidth = document.querySelector(".col-config").getBoundingClientRect().width;
  const paneRect = splitPaneEl.getBoundingClientRect();
  
  const divider = dividerId === 1 ? divider1El : divider2El;
  divider.classList.add("dragging");
  
  const onDrag = (moveEvent) => {
    const deltaX = moveEvent.clientX - startX;
    
    if (dividerId === 1) {
      // Resize palette and canvas (divider1)
      const newPaletteWidth = Math.max(200, Math.min(startLeftWidth + deltaX, 500));
      document.querySelector(".col-palette").style.flex = `0 0 ${newPaletteWidth}px`;
    } else {
      // Resize config and canvas (divider2)
      const newConfigWidth = Math.max(220, Math.min(startRightWidth - deltaX, 500));
      document.querySelector(".col-config").style.flex = `0 0 ${newConfigWidth}px`;
    }
  };
  
  const onDragEnd = () => {
    divider.classList.remove("dragging");
    document.removeEventListener("mousemove", onDrag);
    document.removeEventListener("mouseup", onDragEnd);
    saveSplitState();
  };
  
  document.addEventListener("mousemove", onDrag);
  document.addEventListener("mouseup", onDragEnd);
}

function saveSplitState() {
  const splitState = {
    palette: document.querySelector(".col-palette").style.flex,
    config: document.querySelector(".col-config").style.flex
  };
  localStorage.setItem("formStudio.splitState", JSON.stringify(splitState));
}

function restoreSplitState() {
  const saved = localStorage.getItem("formStudio.splitState");
  if (!saved) {
    return;
  }
  try {
    const parsed = JSON.parse(saved);
    if (parsed.palette) {
      document.querySelector(".col-palette").style.flex = parsed.palette;
    }
    if (parsed.config) {
      document.querySelector(".col-config").style.flex = parsed.config;
    }
  } catch (e) {
    console.error("Failed to restore split state:", e);
  }
}

function bindEvents() {
  document.getElementById("btnExport").addEventListener("click", exportTemplate);
  document.getElementById("btnDownloadTemplate").addEventListener("click", downloadTemplate);
  document.getElementById("btnOpenFiller").addEventListener("click", openFillerPage);
  btnUndoEl.addEventListener("click", undo);
  btnRedoEl.addEventListener("click", redo);
  document.getElementById("btnDuplicate").addEventListener("click", duplicateSelectedField);
  document.getElementById("btnClearSheet").addEventListener("click", clearSheetObjects);
  document.getElementById("btnBringFront").addEventListener("click", bringSelectedToFront);
  document.getElementById("btnSendBack").addEventListener("click", sendSelectedToBack);
  document.getElementById("btnToggleLock").addEventListener("click", toggleSelectedLock);
  document.getElementById("btnCloseDialog").addEventListener("click", () => jsonDialogEl.close());

  toolDecFontEl.addEventListener("click", () => {
    const selected = getSelectedField();
    if (!selected) {
      return;
    }
    ensureTextStyleDefaults(selected);
    updateSelectedTextStyle({ fontSize: selected.fontSize - 1 });
  });

  toolIncFontEl.addEventListener("click", () => {
    const selected = getSelectedField();
    if (!selected) {
      return;
    }
    ensureTextStyleDefaults(selected);
    updateSelectedTextStyle({ fontSize: selected.fontSize + 1 });
  });

  toolFontSizeEl.addEventListener("input", () => {
    updateSelectedTextStyle({ fontSize: Number(toolFontSizeEl.value) || 14 });
  });

  toolBoldEl.addEventListener("click", () => {
    const selected = getSelectedField();
    if (!selected) {
      return;
    }
    ensureTextStyleDefaults(selected);
    updateSelectedTextStyle({ fontWeight: Number(selected.fontWeight) >= 700 ? 500 : 700 });
  });

  toolItalicEl.addEventListener("click", () => {
    const selected = getSelectedField();
    if (!selected) {
      return;
    }
    ensureTextStyleDefaults(selected);
    updateSelectedTextStyle({ fontStyle: selected.fontStyle === "italic" ? "normal" : "italic" });
  });

  toolUnderlineEl.addEventListener("click", () => {
    const selected = getSelectedField();
    if (!selected) {
      return;
    }
    ensureTextStyleDefaults(selected);
    updateSelectedTextStyle({ textDecoration: selected.textDecoration === "underline" ? "none" : "underline" });
  });

  toolAlignLeftEl.addEventListener("click", () => updateSelectedTextStyle({ textAlign: "left" }));
  toolAlignCenterEl.addEventListener("click", () => updateSelectedTextStyle({ textAlign: "center" }));
  toolAlignRightEl.addEventListener("click", () => updateSelectedTextStyle({ textAlign: "right" }));

  // Initialize split-pane draggers
  initSplitPaneDraggers();

  pageSizeSelectEl.addEventListener("change", () => {
    state.page.size = pageSizeSelectEl.value;
    setupSizeEl.value = state.page.size;
    render();
    recordHistory();
  });

  orientationSelectEl.addEventListener("change", () => {
    state.page.orientation = orientationSelectEl.value;
    setupOrientationEl.value = state.page.orientation;
    render();
    recordHistory();
  });

  templateNameEl.addEventListener("input", () => {
    setupNameEl.value = templateNameEl.value;
    scheduleDraftSave();
    scheduleHistoryRecord();
  });

  setupFormEl.addEventListener("submit", (event) => {
    event.preventDefault();
    state.page.size = setupSizeEl.value;
    state.page.orientation = setupOrientationEl.value;
    templateNameEl.value = setupNameEl.value.trim() || "Untitled Template";
    pageSizeSelectEl.value = state.page.size;
    orientationSelectEl.value = state.page.orientation;
    appEl.classList.remove("hidden");
    setupDialogEl.close();
    restoreSplitState();
    render();
    recordHistory();
  });

  window.addEventListener("keydown", (event) => {
    const activeTag = document.activeElement?.tagName;
    const isTyping = ["INPUT", "TEXTAREA", "SELECT"].includes(activeTag);
    if (isTyping) {
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
      event.preventDefault();
      duplicateSelectedField();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "z") {
      event.preventDefault();
      undo();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "y") {
      event.preventDefault();
      redo();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === "z") {
      event.preventDefault();
      redo();
      return;
    }

    const selectedFields = getSelectedFields().filter((f) => !f.locked);
    if (!selectedFields.length) {
      return;
    }

    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      removeSelectedFields();
      return;
    }

    const step = event.shiftKey ? 10 : 1;
    const minX = Math.min(...selectedFields.map((f) => f.x));
    const minY = Math.min(...selectedFields.map((f) => f.y));
    const maxRight = Math.max(...selectedFields.map((f) => f.x + f.w));
    const maxBottom = Math.max(...selectedFields.map((f) => f.y + f.h));
    const maxDx = sheetEl.clientWidth - maxRight;
    const minDx = -minX;
    const maxDy = sheetEl.clientHeight - maxBottom;
    const minDy = -minY;
    let dx = 0;
    let dy = 0;
    let moved = false;

    if (event.key === "ArrowLeft") {
      dx = -step;
      moved = true;
    } else if (event.key === "ArrowRight") {
      dx = step;
      moved = true;
    } else if (event.key === "ArrowUp") {
      dy = -step;
      moved = true;
    } else if (event.key === "ArrowDown") {
      dy = step;
      moved = true;
    }

    if (moved) {
      dx = clamp(dx, minDx, maxDx);
      dy = clamp(dy, minDy, maxDy);
      selectedFields.forEach((selected) => {
        selected.x = Math.round(selected.x + dx);
        selected.y = Math.round(selected.y + dy);
      });
      event.preventDefault();
      renderCanvas();
      renderLayerList();
      renderProperties();
      scheduleDraftSave();
      scheduleHistoryRecord();
    }
  });
}

function boot() {
  buildFormatSelectors();
  buildPalette();
  initDropZone();
  bindEvents();
  restoreSplitState();
  updateHistoryButtons();

  const restored = tryRestoreDraft();
  if (restored) {
    appEl.classList.remove("hidden");
    render();
    recordHistory();
    return;
  }

  setupDialogEl.showModal();
}

boot();
