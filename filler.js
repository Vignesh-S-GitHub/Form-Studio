const templateStorageKey = "formStudio.latestTemplate";
const supportedFieldTypes = new Set([
  "constant", "sectionheader", "divider",
  "table",
  "image", "shaperect", "shapecircle", "shapeline", "elementstar", "elementarrow",
  "text", "email", "phone", "url", "number", "currency", "percentage",
  "textarea",
  "select", "radio", "checkbox", "multiselect",
  "date", "time", "datetime",
  "file", "photo",
  "signature", "rating"
]);

const pageFormats = {
  A5: { w: 148, h: 210 },
  A4: { w: 210, h: 297 },
  A3: { w: 297, h: 420 },
  Letter: { w: 216, h: 279 },
  Legal: { w: 216, h: 356 },
  Tabloid: { w: 279, h: 432 },
  Executive: { w: 184, h: 267 }
};

const MM_PER_INCH = 25.4;

const exportQualityPresets = {
  standard: { dpi: 220, oversample: 1.4, jpegQuality: 0.93, pdfImageType: "PNG" },
  high: { dpi: 300, oversample: 1.8, jpegQuality: 0.97, pdfImageType: "PNG" },
  print: { dpi: 420, oversample: 2.2, jpegQuality: 1, pdfImageType: "PNG" }
};

let currentTemplate = null;

const sheetEl = document.getElementById("fillSheet");
const titleEl = document.getElementById("fillerTitle");
const statusEl = document.getElementById("fillerStatus");
const fileInputEl = document.getElementById("templateFile");
const exportFormatEl = document.getElementById("exportFormat");
const exportQualityEl = document.getElementById("exportQuality");
const btnBackBuilderEl = document.getElementById("btnBackBuilder");
const btnFocusCanvasEl = document.getElementById("btnFocusCanvas");
const btnExitCanvasEl = document.getElementById("btnExitCanvas");
const fillerEssentialToolsEl = document.getElementById("fillerEssentialTools");
const fToolDecFontEl = document.getElementById("fToolDecFont");
const fToolIncFontEl = document.getElementById("fToolIncFont");
const fToolFontSizeEl = document.getElementById("fToolFontSize");
const fToolBoldEl = document.getElementById("fToolBold");
const fToolItalicEl = document.getElementById("fToolItalic");
const fToolUnderlineEl = document.getElementById("fToolUnderline");
const fToolAlignLeftEl = document.getElementById("fToolAlignLeft");
const fToolAlignCenterEl = document.getElementById("fToolAlignCenter");
const fToolAlignRightEl = document.getElementById("fToolAlignRight");

let isCanvasFocusMode = false;
let activeStylableControl = null;

function isStylableControl(el) {
  if (!el || !(el instanceof HTMLElement)) {
    return false;
  }

  if (!(el.matches("input, textarea") && el.classList.contains("fill-control"))) {
    return false;
  }

  if (el.tagName === "TEXTAREA") {
    return true;
  }

  const inputType = (el.getAttribute("type") || "text").toLowerCase();
  return !["radio", "checkbox", "file", "hidden"].includes(inputType);
}

function isInsideEssentialTools(el) {
  return Boolean(el && fillerEssentialToolsEl && fillerEssentialToolsEl.contains(el));
}

function setFillerToolDisabled(disabled) {
  [
    fToolDecFontEl,
    fToolIncFontEl,
    fToolFontSizeEl,
    fToolBoldEl,
    fToolItalicEl,
    fToolUnderlineEl,
    fToolAlignLeftEl,
    fToolAlignCenterEl,
    fToolAlignRightEl
  ].forEach((el) => {
    if (el) {
      el.disabled = disabled;
    }
  });
  fillerEssentialToolsEl?.classList.toggle("is-disabled", disabled);
}

function setActiveToolButton(activeKey) {
  fToolAlignLeftEl?.classList.toggle("active", activeKey === "left");
  fToolAlignCenterEl?.classList.toggle("active", activeKey === "center");
  fToolAlignRightEl?.classList.toggle("active", activeKey === "right");
}

function syncFillerToolsState() {
  const control = activeStylableControl;
  if (!isStylableControl(control)) {
    activeStylableControl = null;
    setFillerToolDisabled(true);
    if (fToolFontSizeEl) {
      fToolFontSizeEl.value = "14";
    }
    fToolBoldEl?.classList.remove("active");
    fToolItalicEl?.classList.remove("active");
    fToolUnderlineEl?.classList.remove("active");
    setActiveToolButton("");
    return;
  }

  setFillerToolDisabled(false);
  const computed = window.getComputedStyle(control);
  const fontSize = Math.round(parseFloat(computed.fontSize || "14"));
  if (fToolFontSizeEl) {
    fToolFontSizeEl.value = String(Number.isFinite(fontSize) ? fontSize : 14);
  }

  fToolBoldEl?.classList.toggle("active", Number(computed.fontWeight) >= 700);
  fToolItalicEl?.classList.toggle("active", computed.fontStyle === "italic");
  fToolUnderlineEl?.classList.toggle("active", computed.textDecorationLine.includes("underline"));

  const align = (control.style.textAlign || computed.textAlign || "left").toLowerCase();
  if (["left", "center", "right"].includes(align)) {
    setActiveToolButton(align);
  } else {
    setActiveToolButton("left");
  }
}

function applyStyleToActiveControl(styles) {
  if (!isStylableControl(activeStylableControl)) {
    return;
  }

  Object.assign(activeStylableControl.style, styles);
  syncFillerToolsState();
}

function toggleStyleOnActiveControl(property, onValue, offValue, computedValue) {
  if (!isStylableControl(activeStylableControl)) {
    return;
  }

  const current = computedValue();
  activeStylableControl.style[property] = current ? offValue : onValue;
  syncFillerToolsState();
}

function setCanvasFocusMode(enabled) {
  const shell = document.querySelector(".filler-shell");
  if (!shell) {
    return;
  }

  isCanvasFocusMode = Boolean(enabled);
  shell.classList.toggle("canvas-focus", isCanvasFocusMode);

  if (btnFocusCanvasEl) {
    btnFocusCanvasEl.textContent = isCanvasFocusMode ? "Exit Canvas" : "Maximize Canvas";
  }

  if (btnExitCanvasEl) {
    btnExitCanvasEl.classList.toggle("hidden", !isCanvasFocusMode);
    btnExitCanvasEl.classList.toggle("visible", isCanvasFocusMode);
  }

  if (currentTemplate) {
    window.requestAnimationFrame(() => renderTemplate(currentTemplate));
  }
}

function toggleCanvasFocusMode() {
  setCanvasFocusMode(!isCanvasFocusMode);
}

function getSelectedQualityPreset() {
  const selected = exportQualityEl?.value || "high";
  return exportQualityPresets[selected] || exportQualityPresets.high;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function computeSheetSize(page) {
  const format = pageFormats[page?.size] || pageFormats.A4;
  if (page?.orientation === "landscape") {
    return { widthMm: format.h, heightMm: format.w };
  }
  return { widthMm: format.w, heightMm: format.h };
}

function applySheetSize(template) {
  const size = computeSheetSize(template.page);
  sheetEl.style.setProperty("--sheet-ratio", `${size.widthMm} / ${size.heightMm}`);
}

function clearSheet(message) {
  sheetEl.innerHTML = `<div class="filler-empty">${message}</div>`;
}

function normalizeType(type) {
  if (type === "paragraph") {
    return "textarea";
  }
  return type;
}

function getSafeFields(template) {
  return (template.fields || [])
    .map((field) => ({ ...field, type: normalizeType(field.type) }))
    .filter((field) => supportedFieldTypes.has(field.type));
}

function updatePhotoPreview(inputEl, previewEl) {
  const file = inputEl.files?.[0];
  if (!file) {
    previewEl.innerHTML = "";
    return;
  }

  const src = URL.createObjectURL(file);
  const img = document.createElement("img");
  img.src = src;
  img.alt = file.name;
  previewEl.innerHTML = "";
  previewEl.appendChild(img);
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

function getTableTrackTemplates(field, rows, cols) {
  const colTracks = parseTableTrackList(field.colWidths, cols);
  const rowTracks = parseTableTrackList(field.rowHeights, rows);

  return {
    colsTemplate: colTracks.map((value) => `${value}fr`).join(" "),
    rowsTemplate: rowTracks.map((value) => `${value}fr`).join(" ")
  };
}

function createInput(field) {
  if (field.type === "table") {
    const matrix = getTableMatrix(field);
    const rows = matrix.length;
    const cols = matrix[0]?.length || 1;
    const trackTemplates = getTableTrackTemplates(field, rows, cols);
    const wrap = document.createElement("div");
    wrap.className = "table-fill-wrap";
    wrap.style.gridTemplateColumns = trackTemplates.colsTemplate;
    wrap.style.gridTemplateRows = trackTemplates.rowsTemplate;
    wrap.style.setProperty("--tbl-border-color", field.borderColor || "#5e75b8");
    wrap.style.setProperty("--tbl-border-width", `${Math.max(1, Number(field.borderWidth) || 1)}px`);
    wrap.style.setProperty("--tbl-bg", field.cellBg || "#ffffff");
    wrap.style.setProperty("--tbl-text", field.textColor || "#24366b");

    const useInputCells = (field.cellMode || "text") === "input";
    matrix.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        const cellEl = document.createElement("div");
        cellEl.className = "table-cell";

        if (useInputCells) {
          const input = document.createElement("input");
          input.className = "table-cell-input";
          input.type = "text";
          input.placeholder = cell || "";
          input.name = `${field.id}__r${rowIndex}c${colIndex}`;
          input.dataset.tableOwner = field.id;
          input.dataset.row = String(rowIndex);
          input.dataset.col = String(colIndex);
          cellEl.appendChild(input);
        } else {
          const text = document.createElement("span");
          text.className = "table-cell-text";
          text.textContent = cell;
          cellEl.appendChild(text);
        }

        wrap.appendChild(cellEl);
      });
    });

    return wrap;
  }

  if (field.type === "image") {
    const wrap = document.createElement("div");
    wrap.className = "insert-image";
    if (field.imageSrc) {
      const img = document.createElement("img");
      img.src = field.imageSrc;
      img.alt = field.label || "Picture";
      img.style.objectFit = field.fit || "cover";
      wrap.appendChild(img);
    } else {
      wrap.textContent = "Picture";
      wrap.classList.add("insert-empty");
    }
    return wrap;
  }

  if (field.type === "shaperect") {
    const div = document.createElement("div");
    div.className = "insert-shape rect";
    div.style.background = field.fillColor || "#dff1eb";
    div.style.border = `${Number(field.strokeWidth) || 1}px solid ${field.strokeColor || "#3b7f71"}`;
    div.style.borderRadius = `${Number(field.radius) || 0}px`;
    return div;
  }

  if (field.type === "shapecircle") {
    const div = document.createElement("div");
    div.className = "insert-shape circle";
    div.style.background = field.fillColor || "#f8ead9";
    div.style.border = `${Number(field.strokeWidth) || 1}px solid ${field.strokeColor || "#ba7a36"}`;
    return div;
  }

  if (field.type === "shapeline") {
    const div = document.createElement("div");
    div.className = "insert-shape line";
    div.style.background = field.strokeColor || "#4f6f67";
    div.style.height = `${Math.max(1, Number(field.strokeWidth) || 2)}px`;
    return div;
  }

  if (field.type === "elementstar") {
    const div = document.createElement("div");
    div.className = "insert-element";
    div.style.color = field.fillColor || "#f3ca62";
    div.textContent = "★";
    return div;
  }

  if (field.type === "elementarrow") {
    const div = document.createElement("div");
    div.className = "insert-element";
    div.style.color = field.fillColor || "#93b9ad";
    div.textContent = "➜";
    return div;
  }

  // Static display fields
  if (field.type === "constant" || field.type === "sectionheader") {
    const div = document.createElement("div");
    div.className = field.type === "sectionheader" ? "section-header" : "constant-text";
    div.textContent = field.value || "-";
    return div;
  }

  if (field.type === "divider") {
    const div = document.createElement("div");
    div.className = "divider-line";
    return div;
  }

  // Text-based inputs
  if (["text", "email", "phone", "url", "number", "currency", "percentage", "date", "time", "datetime"].includes(field.type)) {
    const input = document.createElement("input");
    input.className = "fill-control";
    input.type = field.type === "currency" || field.type === "percentage" ? "number" : field.type;
    input.placeholder = field.placeholder || "";
    input.name = field.id;
    if (field.type === "number" || field.type === "currency" || field.type === "percentage") {
      input.step = "any";
    }
    return input;
  }

  // Textarea
  if (field.type === "textarea") {
    const el = document.createElement("textarea");
    el.className = "fill-control";
    el.placeholder = field.placeholder || "";
    el.name = field.id;
    return el;
  }

  // Dropdown select
  if (field.type === "select") {
    const select = document.createElement("select");
    select.className = "fill-control";
    select.name = field.id;
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = field.placeholder || "Choose...";
    select.appendChild(defaultOption);
    const options = (field.options || "").split(",").map((o) => o.trim());
    options.forEach((opt) => {
      if (opt) {
        const option = document.createElement("option");
        option.value = opt;
        option.textContent = opt;
        select.appendChild(option);
      }
    });
    return select;
  }

  // Radio buttons
  if (field.type === "radio") {
    const wrap = document.createElement("div");
    wrap.className = "radio-group";
    const options = (field.options || "").split(",").map((o) => o.trim());
    options.forEach((opt) => {
      if (opt) {
        const label = document.createElement("label");
        label.className = "radio-label";
        const input = document.createElement("input");
        input.type = "radio";
        input.name = field.id;
        input.value = opt;
        label.appendChild(input);
        label.appendChild(document.createTextNode(opt));
        wrap.appendChild(label);
      }
    });
    return wrap;
  }

  // Checkboxes
  if (field.type === "checkbox") {
    const wrap = document.createElement("div");
    wrap.className = "checkbox-group";
    const options = (field.options || "").split(",").map((o) => o.trim());
    options.forEach((opt) => {
      if (opt) {
        const label = document.createElement("label");
        label.className = "checkbox-label";
        const input = document.createElement("input");
        input.type = "checkbox";
        input.name = field.id;
        input.value = opt;
        label.appendChild(input);
        label.appendChild(document.createTextNode(opt));
        wrap.appendChild(label);
      }
    });
    return wrap;
  }

  // Multi-select
  if (field.type === "multiselect") {
    const select = document.createElement("select");
    select.className = "fill-control";
    select.name = field.id;
    select.multiple = true;
    const options = (field.options || "").split(",").map((o) => o.trim());
    options.forEach((opt) => {
      if (opt) {
        const option = document.createElement("option");
        option.value = opt;
        option.textContent = opt;
        select.appendChild(option);
      }
    });
    return select;
  }

  // File upload
  if (field.type === "file") {
    const wrap = document.createElement("div");
    wrap.className = "file-wrap";
    const input = document.createElement("input");
    input.className = "fill-control";
    input.type = "file";
    input.accept = field.accept || ".pdf,.doc,.docx";
    input.name = field.id;
    input.multiple = (field.maxFiles || 1) > 1;
    wrap.appendChild(input);
    return wrap;
  }

  // Photo upload
  if (field.type === "photo") {
    const wrap = document.createElement("div");
    wrap.className = "photo-wrap";
    const input = document.createElement("input");
    input.className = "fill-control";
    input.type = "file";
    input.accept = field.accept || "image/*";
    input.name = field.id;
    input.multiple = (field.maxFiles || 1) > 1;
    const preview = document.createElement("div");
    preview.className = "photo-preview";
    input.addEventListener("change", () => updatePhotoPreview(input, preview));
    wrap.appendChild(input);
    wrap.appendChild(preview);
    return wrap;
  }

  // Signature field
  if (field.type === "signature") {
    const wrap = document.createElement("div");
    wrap.className = "signature-wrap";
    const canvas = document.createElement("canvas");
    canvas.className = "signature-pad";
    canvas.width = 400;
    canvas.height = 150;
    const style = document.createElement("style");
    style.textContent = `.signature-pad { border: 2px solid #ccc; background: white; cursor: crosshair; }`;
    wrap.appendChild(style);
    wrap.appendChild(canvas);
    const hidden = document.createElement("input");
    hidden.type = "hidden";
    hidden.name = field.id;
    wrap.appendChild(hidden);
    return wrap;
  }

  // Rating field
  if (field.type === "rating") {
    const wrap = document.createElement("div");
    wrap.className = "rating-group";
    const maxRating = field.maxRating || 5;
    const hiddenInput = document.createElement("input");
    hiddenInput.type = "hidden";
    hiddenInput.name = field.id;
    hiddenInput.value = 0;
    wrap.appendChild(hiddenInput);
    for (let i = 1; i <= maxRating; i++) {
      const label = document.createElement("label");
      label.className = "rating-label";
      const input = document.createElement("input");
      input.type = "radio";
      input.name = `${field.id}-rating`;
      input.value = i;
      input.addEventListener("change", () => {
        hiddenInput.value = i;
      });
      const star = document.createElement("span");
      star.textContent = "★";
      label.appendChild(input);
      label.appendChild(star);
      wrap.appendChild(label);
    }
    return wrap;
  }

  // Default text input for any unknown type
  const input = document.createElement("input");
  input.className = "fill-control";
  input.type = "text";
  input.placeholder = field.placeholder || "";
  input.name = field.id;
  return input;
}

function captureInputState() {
  if (!currentTemplate) {
    return {};
  }

  const stateMap = {};

  currentTemplate.fields.forEach((field) => {
    if (field.type === "table" && (field.cellMode || "text") === "input") {
      const matrix = getTableMatrix(field).map((row) => [...row]);
      const inputs = Array.from(sheetEl.querySelectorAll(`[data-table-owner="${CSS.escape(field.id)}"]`));
      inputs.forEach((input) => {
        const rowIndex = Number(input.dataset.row);
        const colIndex = Number(input.dataset.col);
        if (Number.isInteger(rowIndex) && Number.isInteger(colIndex) && matrix[rowIndex] && typeof matrix[rowIndex][colIndex] !== "undefined") {
          matrix[rowIndex][colIndex] = input.value || "";
        }
      });
      stateMap[field.id] = { type: field.type, grid: matrix };
      return;
    }

    if (field.type === "radio") {
      const checked = sheetEl.querySelector(`input[type="radio"][name="${CSS.escape(field.id)}"]:checked`);
      stateMap[field.id] = { type: field.type, value: checked ? checked.value : "" };
      return;
    }

    if (field.type === "checkbox") {
      const checked = Array.from(sheetEl.querySelectorAll(`input[type="checkbox"][name="${CSS.escape(field.id)}"]:checked`)).map((el) => el.value);
      stateMap[field.id] = { type: field.type, value: checked };
      return;
    }

    if (field.type === "multiselect") {
      const selectEl = sheetEl.querySelector(`select[name="${CSS.escape(field.id)}"]`);
      const selectedValues = selectEl ? Array.from(selectEl.selectedOptions).map((opt) => opt.value) : [];
      stateMap[field.id] = { type: field.type, value: selectedValues };
      return;
    }

    const el = sheetEl.querySelector(`[name="${CSS.escape(field.id)}"]`);
    if (el) {
      stateMap[field.id] = { type: field.type, value: el.value || "" };
    }
  });

  return stateMap;
}

function applyInputState(stateMap) {
  if (!stateMap || typeof stateMap !== "object") {
    return;
  }

  currentTemplate.fields.forEach((field) => {
    const saved = stateMap[field.id];
    if (!saved) {
      return;
    }

    if (field.type === "table" && (field.cellMode || "text") === "input" && Array.isArray(saved.grid)) {
      const inputs = Array.from(sheetEl.querySelectorAll(`[data-table-owner="${CSS.escape(field.id)}"]`));
      inputs.forEach((input) => {
        const rowIndex = Number(input.dataset.row);
        const colIndex = Number(input.dataset.col);
        if (Number.isInteger(rowIndex) && Number.isInteger(colIndex) && saved.grid[rowIndex] && typeof saved.grid[rowIndex][colIndex] !== "undefined") {
          input.value = saved.grid[rowIndex][colIndex] || "";
        }
      });
      return;
    }

    if (field.type === "radio") {
      const radios = Array.from(sheetEl.querySelectorAll(`input[type="radio"][name="${CSS.escape(field.id)}"]`));
      radios.forEach((radio) => {
        radio.checked = radio.value === String(saved.value || "");
      });
      return;
    }

    if (field.type === "checkbox") {
      const values = Array.isArray(saved.value) ? saved.value : [];
      const checks = Array.from(sheetEl.querySelectorAll(`input[type="checkbox"][name="${CSS.escape(field.id)}"]`));
      checks.forEach((checkbox) => {
        checkbox.checked = values.includes(checkbox.value);
      });
      return;
    }

    if (field.type === "multiselect") {
      const selectEl = sheetEl.querySelector(`select[name="${CSS.escape(field.id)}"]`);
      if (selectEl) {
        const values = Array.isArray(saved.value) ? saved.value : [];
        Array.from(selectEl.options).forEach((opt) => {
          opt.selected = values.includes(opt.value);
        });
      }
      return;
    }

    const el = sheetEl.querySelector(`[name="${CSS.escape(field.id)}"]`);
    if (el && typeof saved.value !== "undefined") {
      el.value = saved.value;
    }
  });
}

function renderTemplate(template, preserveValues = false) {
  const previousState = preserveValues ? captureInputState() : null;
  const safeFields = getSafeFields(template);
  currentTemplate = { ...template, fields: safeFields };
  titleEl.textContent = currentTemplate.templateName || "Fill Form";
  statusEl.textContent = `${safeFields.length} objects`;
  applySheetSize(currentTemplate);

  sheetEl.innerHTML = "";
  const sheetWidth = sheetEl.clientWidth || 794;
  const sheetHeight = sheetEl.clientHeight || 1123;

  if (!safeFields.length) {
    clearSheet("Template has no objects. Go back to builder and add fields.");
    return;
  }

  const sorted = [...safeFields].sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0));
  sorted.forEach((field) => {
    const card = document.createElement("div");
    card.className = `fill-object ${field.type}`;

    const safeW = clamp(Number(field.w) || 160, 70, sheetWidth);
    const safeH = clamp(Number(field.h) || 60, 40, sheetHeight);
    const safeX = clamp(Number(field.x) || 0, 0, Math.max(0, sheetWidth - safeW));
    const safeY = clamp(Number(field.y) || 0, 0, Math.max(0, sheetHeight - safeH));

    card.style.left = `${safeX}px`;
    card.style.top = `${safeY}px`;
    card.style.width = `${safeW}px`;
    card.style.height = `${safeH}px`;

    card.appendChild(createInput(field));
    sheetEl.appendChild(card);
  });

  if (previousState) {
    applyInputState(previousState);
  }
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}

async function collectResponseData() {
  const response = {
    templateName: currentTemplate.templateName,
    createdAt: new Date().toISOString(),
    values: {}
  };

  for (const field of currentTemplate.fields) {
    if (field.type === "constant") {
      response.values[field.id] = {
        label: field.label,
        type: field.type,
        value: field.value || ""
      };
      continue;
    }

    if (field.type === "table") {
      const matrix = getTableMatrix(field);
      const useInputCells = (field.cellMode || "text") === "input";

      if (!useInputCells) {
        response.values[field.id] = {
          label: field.label,
          type: field.type,
          value: matrix.map((row) => row.join(" | ")).join("\n"),
          grid: matrix
        };
        continue;
      }

      const inputs = Array.from(sheetEl.querySelectorAll(`[data-table-owner="${CSS.escape(field.id)}"]`));
      const outGrid = matrix.map((row) => [...row]);

      inputs.forEach((input) => {
        const rowIndex = Number(input.dataset.row);
        const colIndex = Number(input.dataset.col);
        if (Number.isInteger(rowIndex) && Number.isInteger(colIndex) && outGrid[rowIndex] && typeof outGrid[rowIndex][colIndex] !== "undefined") {
          outGrid[rowIndex][colIndex] = (input.value || "").trim();
        }
      });

      response.values[field.id] = {
        label: field.label,
        type: field.type,
        value: outGrid.map((row) => row.join(" | ")).join("\n"),
        grid: outGrid
      };
      continue;
    }

    if (field.type === "radio") {
      const checked = sheetEl.querySelector(`input[type="radio"][name="${CSS.escape(field.id)}"]:checked`);
      response.values[field.id] = { label: field.label, type: field.type, value: checked ? checked.value : "" };
      continue;
    }

    if (field.type === "checkbox") {
      const checkedValues = Array.from(sheetEl.querySelectorAll(`input[type="checkbox"][name="${CSS.escape(field.id)}"]:checked`)).map((el) => el.value);
      response.values[field.id] = { label: field.label, type: field.type, value: checkedValues };
      continue;
    }

    if (field.type === "multiselect") {
      const selectEl = sheetEl.querySelector(`select[name="${CSS.escape(field.id)}"]`);
      const selectedValues = selectEl ? Array.from(selectEl.selectedOptions).map((opt) => opt.value) : [];
      response.values[field.id] = { label: field.label, type: field.type, value: selectedValues };
      continue;
    }

    const element = sheetEl.querySelector(`[name="${CSS.escape(field.id)}"]`);
    if (!element) {
      response.values[field.id] = { label: field.label, type: field.type, value: field.value || "" };
      continue;
    }

    if (field.type === "photo") {
      const files = Array.from(element.files || []);
      const images = [];
      for (const file of files) {
        try {
          const dataUrl = await readFileAsDataUrl(file);
          images.push({ name: file.name, dataUrl });
        } catch {
          images.push({ name: file.name, dataUrl: "" });
        }
      }

      response.values[field.id] = {
        label: field.label,
        type: field.type,
        value: files.map((f) => f.name),
        images
      };
      continue;
    }

    if (field.type === "file") {
      const files = Array.from(element.files || []);
      response.values[field.id] = {
        label: field.label,
        type: field.type,
        value: files.map((f) => f.name)
      };
      continue;
    }

    response.values[field.id] = { label: field.label, type: field.type, value: (element.value || "").trim() };
  }

  return response;
}

function buildSafeName() {
  return (currentTemplate?.templateName || "filled-form")
    .trim()
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/(^-|-$)/g, "")
    .toLowerCase() || "filled-form";
}

function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function goToBuilder() {
  window.location.href = "index.html";
}

function getRenderableEntries(response) {
  const entries = [];

  currentTemplate.fields.forEach((field) => {
    const row = response.values[field.id];

    const fallbackRow = {
      label: field.label,
      type: field.type,
      value: field.value || ""
    };

    entries.push({ field, row: row || fallbackRow });
  });

  return entries;
}

function createSnapshotSheet(response, targetWidth, targetHeight) {
  const baseWidth = Math.max(1, sheetEl.clientWidth || 794);
  const baseHeight = Math.max(1, sheetEl.clientHeight || 1123);
  const width = targetWidth || baseWidth;
  const height = targetHeight || baseHeight;
  const scaleX = width / baseWidth;
  const scaleY = height / baseHeight;

  const snapshot = document.createElement("div");
  snapshot.className = "sheet fill-sheet";
  snapshot.style.boxShadow = "none";
  snapshot.style.background = "#ffffff";
  snapshot.style.width = `${width}px`;
  snapshot.style.height = `${height}px`;
  snapshot.style.position = "relative";
  snapshot.style.overflow = "hidden";

  const entries = getRenderableEntries(response);

  if (!entries.length) {
    const empty = document.createElement("div");
    empty.className = "filler-empty";
    empty.textContent = "No filled values to export.";
    snapshot.appendChild(empty);
    return snapshot;
  }

  entries.forEach(({ field, row }) => {
    const block = document.createElement("div");
    block.className = `export-object ${field.type}`;
    block.style.left = `${(Number(field.x) || 0) * scaleX}px`;
    block.style.top = `${(Number(field.y) || 0) * scaleY}px`;
    block.style.width = `${(Number(field.w) || 160) * scaleX}px`;
    block.style.height = `${(Number(field.h) || 60) * scaleY}px`;
    block.style.fontSize = `${Math.max(10, 14 * Math.min(scaleX, scaleY))}px`;
    block.style.lineHeight = "1.35";

    if (field.type === "photo") {
      const firstImage = row.images?.[0]?.dataUrl;
      if (firstImage) {
        const img = document.createElement("img");
        img.src = firstImage;
        img.alt = row.images?.[0]?.name || "photo";
        block.appendChild(img);
      }
    } else if (field.type === "sectionheader") {
      block.style.fontWeight = "700";
      block.style.fontSize = `${Math.max(14, 16 * Math.min(scaleX, scaleY))}px`;
      block.textContent = String(field.value || field.label || "Section");
    } else if (field.type === "divider") {
      block.style.background = "#9fb7af";
      block.style.height = `${Math.max(1, 2 * scaleY)}px`;
      block.style.top = `${((Number(field.y) || 0) + ((Number(field.h) || 8) / 2)) * scaleY}px`;
      block.style.borderRadius = "999px";
    } else if (field.type === "table") {
      const rows = clamp(Math.round(Number(field.rows) || 3), 1, 12);
      const cols = clamp(Math.round(Number(field.cols) || 3), 1, 12);
      const trackTemplates = getTableTrackTemplates(field, rows, cols);
      const matrix = Array.isArray(row.grid) ? row.grid : getTableMatrix(field);
      const table = document.createElement("div");
      table.className = "export-table-grid";
      table.style.gridTemplateColumns = trackTemplates.colsTemplate;
      table.style.gridTemplateRows = trackTemplates.rowsTemplate;
      table.style.setProperty("--tbl-border-color", field.borderColor || "#5e75b8");
      table.style.setProperty("--tbl-border-width", `${Math.max(1, Number(field.borderWidth) || 1)}px`);
      table.style.setProperty("--tbl-bg", field.cellBg || "#ffffff");
      table.style.setProperty("--tbl-text", field.textColor || "#24366b");

      for (let r = 0; r < rows; r += 1) {
        for (let c = 0; c < cols; c += 1) {
          const cell = document.createElement("div");
          cell.className = "export-table-cell";
          cell.textContent = matrix[r]?.[c] || "";
          table.appendChild(cell);
        }
      }

      block.appendChild(table);
    } else if (field.type === "image") {
      if (field.imageSrc) {
        const img = document.createElement("img");
        img.src = field.imageSrc;
        img.alt = field.label || "Picture";
        img.style.objectFit = field.fit || "cover";
        block.appendChild(img);
      }
    } else if (field.type === "shaperect") {
      block.style.background = field.fillColor || "#dff1eb";
      block.style.border = `${Number(field.strokeWidth) || 1}px solid ${field.strokeColor || "#3b7f71"}`;
      block.style.borderRadius = `${Number(field.radius) || 0}px`;
    } else if (field.type === "shapecircle") {
      block.style.background = field.fillColor || "#f8ead9";
      block.style.border = `${Number(field.strokeWidth) || 1}px solid ${field.strokeColor || "#ba7a36"}`;
      block.style.borderRadius = "999px";
    } else if (field.type === "shapeline") {
      block.style.background = field.strokeColor || "#4f6f67";
      block.style.height = `${Math.max(1, Number(field.strokeWidth) || 2) * scaleY}px`;
      block.style.top = `${((Number(field.y) || 0) + ((Number(field.h) || 8) / 2)) * scaleY}px`;
      block.style.borderRadius = "999px";
    } else if (field.type === "elementstar") {
      block.style.display = "grid";
      block.style.placeItems = "center";
      block.style.fontSize = `${Math.max(20, Math.min((Number(field.w) || 80) * scaleX, (Number(field.h) || 80) * scaleY))}px`;
      block.style.color = field.fillColor || "#f3ca62";
      block.textContent = "★";
    } else if (field.type === "elementarrow") {
      block.style.display = "grid";
      block.style.placeItems = "center";
      block.style.fontSize = `${Math.max(18, Math.min((Number(field.w) || 120) * scaleX, (Number(field.h) || 50) * scaleY))}px`;
      block.style.color = field.fillColor || "#93b9ad";
      block.textContent = "➜";
    } else {
      const rawValue = row?.value;
      const textValue = Array.isArray(rawValue) ? rawValue.join(", ") : String(rawValue || "");

      // Export input-like fields as text only, without form box visuals.
      block.style.display = "block";
      block.style.padding = "0";
      block.style.border = "none";
      block.style.background = "transparent";
      block.style.borderRadius = "0";
      block.style.whiteSpace = "pre-wrap";
      block.style.wordBreak = "break-word";
      block.style.color = "#17322f";
      block.style.fontSize = `${Math.max(10, 13 * Math.min(scaleX, scaleY))}px`;
      block.style.lineHeight = "1.35";
      block.textContent = textValue;
    }

    snapshot.appendChild(block);
  });

  return snapshot;
}

function getExportPixelSize(quality) {
  const page = currentTemplate?.page || { size: "A4", orientation: "portrait" };
  const mm = computeSheetSize(page);
  const dpi = quality?.dpi || exportQualityPresets.high.dpi;

  const widthPx = Math.max(1200, Math.round((mm.widthMm / MM_PER_INCH) * dpi));
  const heightPx = Math.max(1600, Math.round((mm.heightMm / MM_PER_INCH) * dpi));

  return { widthPx, heightPx };
}

function setupOffscreenSnapshot(snapshot, widthPx, heightPx) {
  snapshot.style.position = "fixed";
  snapshot.style.left = "0";
  snapshot.style.top = "0";
  snapshot.style.opacity = "1";
  snapshot.style.pointerEvents = "none";
  snapshot.style.zIndex = "-1";
  snapshot.style.width = `${widthPx}px`;
  snapshot.style.height = `${heightPx}px`;
  document.body.appendChild(snapshot);
}

function getRenderDimensions(widthPx, heightPx, quality) {
  const oversample = Math.max(1, Number(quality?.oversample) || 1);
  const maxEdge = 12000;

  let renderWidth = Math.round(widthPx * oversample);
  let renderHeight = Math.round(heightPx * oversample);

  const largestEdge = Math.max(renderWidth, renderHeight);
  if (largestEdge > maxEdge) {
    const scaleDown = maxEdge / largestEdge;
    renderWidth = Math.max(1, Math.round(renderWidth * scaleDown));
    renderHeight = Math.max(1, Math.round(renderHeight * scaleDown));
  }

  return { renderWidth, renderHeight };
}

async function renderHighQualityCanvas(snapshot, widthPx, heightPx, quality) {
  const { renderWidth, renderHeight } = getRenderDimensions(widthPx, heightPx, quality);

  if (window.htmlToImage) {
    try {
      const options = {
        cacheBust: true,
        backgroundColor: "#ffffff",
        pixelRatio: 1,
        canvasWidth: renderWidth,
        canvasHeight: renderHeight,
        skipAutoScale: true,
        style: {
          transform: "none"
        }
      };

      if (quality?.pdfImageType === "JPEG") {
        return await window.htmlToImage.toJpeg(snapshot, {
          ...options,
          quality: quality.jpegQuality || 0.95
        });
      }

      return await window.htmlToImage.toPng(snapshot, options);
    } catch {
      // Fall through to html2canvas fallback
    }
  }

  if (!window.html2canvas) {
    throw new Error("No export renderer is available");
  }

  const canvas = await window.html2canvas(snapshot, {
    backgroundColor: "#ffffff",
    scale: Math.max(1, renderWidth / Math.max(1, widthPx)),
    useCORS: true,
    allowTaint: false,
    imageTimeout: 0,
    logging: false
  });

  if (quality?.pdfImageType === "JPEG") {
    return canvas.toDataURL("image/jpeg", quality.jpegQuality || 0.95);
  }
  return canvas.toDataURL("image/png");
}

async function exportAsImage(response, imageType) {
  const qualitySetting = getSelectedQualityPreset();
  const { widthPx, heightPx } = getExportPixelSize(qualitySetting);
  const snapshot = createSnapshotSheet(response, widthPx, heightPx);
  setupOffscreenSnapshot(snapshot, widthPx, heightPx);

  try {
    const engineQuality = imageType === "jpg"
      ? { ...qualitySetting, pdfImageType: "JPEG" }
      : { ...qualitySetting, pdfImageType: "PNG" };
    const dataUrl = await renderHighQualityCanvas(snapshot, widthPx, heightPx, engineQuality);
    const res = await fetch(dataUrl);
    const blob = await res.blob();
    triggerDownload(blob, `${buildSafeName()}-filled.${imageType}`);
  } finally {
    snapshot.remove();
  }
}

async function exportAsPdf(response) {
  const qualitySetting = getSelectedQualityPreset();
  const { widthPx, heightPx } = getExportPixelSize(qualitySetting);
  const snapshot = createSnapshotSheet(response, widthPx, heightPx);
  setupOffscreenSnapshot(snapshot, widthPx, heightPx);

  try {
    const pdfImageType = qualitySetting.pdfImageType || "PNG";
    const imgData = await renderHighQualityCanvas(snapshot, widthPx, heightPx, qualitySetting);
    const mm = computeSheetSize(currentTemplate?.page || { size: "A4", orientation: "portrait" });
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation: mm.widthMm >= mm.heightMm ? "landscape" : "portrait",
      unit: "mm",
      format: [mm.widthMm, mm.heightMm]
    });

    try {
      pdf.addImage(imgData, pdfImageType, 0, 0, mm.widthMm, mm.heightMm, undefined, "FAST");
    } catch {
      // Fallback path for environments where PNG embedding can fail.
      const fallbackData = await renderHighQualityCanvas(snapshot, widthPx, heightPx, {
        ...qualitySetting,
        pdfImageType: "JPEG",
        jpegQuality: Math.min(1, Math.max(0.85, qualitySetting.jpegQuality || 0.95))
      });
      pdf.addImage(fallbackData, "JPEG", 0, 0, mm.widthMm, mm.heightMm, undefined, "FAST");
    }

    pdf.save(`${buildSafeName()}-filled.pdf`);
  } finally {
    snapshot.remove();
  }
}

async function downloadFilledData() {
  if (!currentTemplate) {
    window.requestAnimationFrame(() => renderTemplate(currentTemplate, true));
  }

  const response = await collectResponseData();
  const format = exportFormatEl?.value || "pdf";

  if (format === "pdf") {
    await exportAsPdf(response);
    return;
  }

  if (format === "png" || format === "jpg") {
    await exportAsImage(response, format);
  }
}

function loadLatestTemplate() {
  const raw = localStorage.getItem(templateStorageKey);
  if (!raw) {
    clearSheet("No template found. Redirecting to builder...");
    statusEl.textContent = "No template loaded";
    const fromBuilder = new URLSearchParams(window.location.search).get("source") === "builder";
    if (fromBuilder) {
      window.setTimeout(goToBuilder, 1200);
    }
    return;
  }

  try {
    renderTemplate(JSON.parse(raw));
  } catch {
    clearSheet("Saved template is invalid JSON.");
    statusEl.textContent = "Template error";
  }
}

function initFillerSplitPaneDraggers() {
  const divider = document.getElementById("fillerDivider");
  if (divider) {
    divider.addEventListener("mousedown", startFillerSplitDrag);
  }
}

function startFillerSplitDrag(event) {
  event.preventDefault();
  
  const startX = event.clientX;
  const startInfoWidth = document.querySelector(".col-info").getBoundingClientRect().width;
  const divider = event.target.closest(".split-divider");
  
  divider.classList.add("dragging");
  
  const onDrag = (moveEvent) => {
    const deltaX = moveEvent.clientX - startX;
    const newInfoWidth = Math.max(180, Math.min(startInfoWidth + deltaX, 500));
    document.querySelector(".col-info").style.flex = `0 0 ${newInfoWidth}px`;
  };
  
  const onDragEnd = () => {
    divider.classList.remove("dragging");
    document.removeEventListener("mousemove", onDrag);
    document.removeEventListener("mouseup", onDragEnd);
    saveFillerSplitState();
  };
  
  document.addEventListener("mousemove", onDrag);
  document.addEventListener("mouseup", onDragEnd);
}

function saveFillerSplitState() {
  const splitState = {
    info: document.querySelector(".col-info")?.style.flex
  };
  localStorage.setItem("formStudio.fillerSplitState", JSON.stringify(splitState));
}

function restoreFillerSplitState() {
  const saved = localStorage.getItem("formStudio.fillerSplitState");
  if (!saved) {
    return;
  }
  try {
    const parsed = JSON.parse(saved);
    if (parsed.info) {
      document.querySelector(".col-info").style.flex = parsed.info;
    }
  } catch (e) {
    console.error("Failed to restore filler split state:", e);
  }
}

function bindEvents() {
  // Initialize split-pane draggers
  initFillerSplitPaneDraggers();
  
  if (btnBackBuilderEl) {
    btnBackBuilderEl.addEventListener("click", goToBuilder);
  }

  document.getElementById("btnLoadLatest").addEventListener("click", loadLatestTemplate);
  document.getElementById("btnDownloadFilled").addEventListener("click", downloadFilledData);

  if (btnFocusCanvasEl) {
    btnFocusCanvasEl.addEventListener("click", toggleCanvasFocusMode);
  }

  if (btnExitCanvasEl) {
    btnExitCanvasEl.addEventListener("click", () => setCanvasFocusMode(false));
  }

  if (fillerEssentialToolsEl) {
    setFillerToolDisabled(true);

    [
      fToolDecFontEl,
      fToolIncFontEl,
      fToolBoldEl,
      fToolItalicEl,
      fToolUnderlineEl,
      fToolAlignLeftEl,
      fToolAlignCenterEl,
      fToolAlignRightEl
    ].forEach((btn) => {
      btn?.addEventListener("mousedown", (event) => {
        event.preventDefault();
      });
    });

    fToolDecFontEl?.addEventListener("click", () => {
      const current = Number(fToolFontSizeEl?.value || 14);
      const next = clamp(current - 1, 8, 72);
      applyStyleToActiveControl({ fontSize: `${next}px` });
    });

    fToolIncFontEl?.addEventListener("click", () => {
      const current = Number(fToolFontSizeEl?.value || 14);
      const next = clamp(current + 1, 8, 72);
      applyStyleToActiveControl({ fontSize: `${next}px` });
    });

    fToolFontSizeEl?.addEventListener("input", () => {
      const next = clamp(Number(fToolFontSizeEl.value || 14), 8, 72);
      applyStyleToActiveControl({ fontSize: `${next}px` });
    });

    fToolBoldEl?.addEventListener("click", () => {
      toggleStyleOnActiveControl("fontWeight", "700", "500", () => {
        const value = window.getComputedStyle(activeStylableControl).fontWeight;
        return Number(value) >= 700;
      });
    });

    fToolItalicEl?.addEventListener("click", () => {
      toggleStyleOnActiveControl("fontStyle", "italic", "normal", () => {
        return window.getComputedStyle(activeStylableControl).fontStyle === "italic";
      });
    });

    fToolUnderlineEl?.addEventListener("click", () => {
      toggleStyleOnActiveControl("textDecoration", "underline", "none", () => {
        const value = window.getComputedStyle(activeStylableControl).textDecorationLine;
        return value.includes("underline");
      });
    });

    fToolAlignLeftEl?.addEventListener("click", () => applyStyleToActiveControl({ textAlign: "left" }));
    fToolAlignCenterEl?.addEventListener("click", () => applyStyleToActiveControl({ textAlign: "center" }));
    fToolAlignRightEl?.addEventListener("click", () => applyStyleToActiveControl({ textAlign: "right" }));

    sheetEl.addEventListener("focusin", (event) => {
      const target = event.target;
      if (isStylableControl(target)) {
        activeStylableControl = target;
      } else {
        activeStylableControl = null;
      }
      syncFillerToolsState();
    });

    sheetEl.addEventListener("focusout", () => {
      window.setTimeout(() => {
        const active = document.activeElement;
        if (isInsideEssentialTools(active)) {
          return;
        }
        if (!isStylableControl(active)) {
          activeStylableControl = null;
          syncFillerToolsState();
        }
      }, 0);
    });
  }

  fileInputEl.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    const text = await file.text();
    try {
      const parsed = JSON.parse(text);
      localStorage.setItem(templateStorageKey, JSON.stringify(parsed));
      renderTemplate(parsed);
    } catch {
      clearSheet("Uploaded file is not valid JSON.");
      statusEl.textContent = "Template error";
    }
  });

  window.addEventListener("resize", () => {
    if (currentTemplate) {
      renderTemplate(currentTemplate, true);
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && isCanvasFocusMode) {
      setCanvasFocusMode(false);
    }
  });
}

bindEvents();
restoreFillerSplitState();
loadLatestTemplate();
