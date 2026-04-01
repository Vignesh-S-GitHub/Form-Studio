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
  standard: { dpi: 150, jpegQuality: 0.88, pdfImageType: "JPEG" },
  high:     { dpi: 220, jpegQuality: 0.95, pdfImageType: "PNG" },
  print:    { dpi: 300, jpegQuality: 1.00, pdfImageType: "PNG" }
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
    div.textContent = "â˜…";
    return div;
  }

  if (field.type === "elementarrow") {
    const div = document.createElement("div");
    div.className = "insert-element";
    div.style.color = field.fillColor || "#93b9ad";
    div.textContent = "âžœ";
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
      star.textContent = "â˜…";
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
    if (!row) {
      return;
    }

    if (["constant", "sectionheader", "divider", "table", "image", "shaperect", "shapecircle", "shapeline", "elementstar", "elementarrow"].includes(field.type)) {
      entries.push({ field, row });
      return;
    }

    if (field.type === "photo") {
      if ((row.images || []).length) {
        entries.push({ field, row });
      }
      return;
    }

    if (String(row.value || "").trim()) {
      entries.push({ field, row });
    }
  });

  return entries;
}

// Capture inline styles the user applied via Essential Tools on the fill sheet
function captureInputStyles() {
  const styles = {};
  if (!currentTemplate) return styles;
  currentTemplate.fields.forEach((field) => {
    const el = sheetEl.querySelector(`[name="${CSS.escape(field.id)}"]`);
    if (!el) return;
    const s = el.style;
    const captured = {};
    if (s.fontSize)      captured.fontSize      = parseInt(s.fontSize, 10) || field.fontSize;
    if (s.fontWeight)    captured.fontWeight    = s.fontWeight;
    if (s.fontStyle)     captured.fontStyle     = s.fontStyle;
    if (s.textDecoration) captured.textDecoration = s.textDecoration;
    if (s.textAlign)     captured.textAlign     = s.textAlign;
    if (Object.keys(captured).length) styles[field.id] = captured;
  });
  return styles;
}

// Canvas 2D helper â€” rounded rectangle path
function ctxRoundRect(ctx, x, y, w, h, r) {
  const r2 = Math.min(r, w / 2, h / 2);
  ctx.beginPath();
  ctx.moveTo(x + r2, y);
  ctx.lineTo(x + w - r2, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r2);
  ctx.lineTo(x + w, y + h - r2);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r2, y + h);
  ctx.lineTo(x + r2, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r2);
  ctx.lineTo(x, y + r2);
  ctx.quadraticCurveTo(x, y, x + r2, y);
  ctx.closePath();
}

// Canvas 2D helper â€” load and draw an image with cover/contain fit
function drawImageToCanvas(ctx, src, x, y, w, h, fit, radius) {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      ctx.save();
      if (radius > 0) { ctxRoundRect(ctx, x, y, w, h, radius); ctx.clip(); }
      const nw = img.naturalWidth || 1;
      const nh = img.naturalHeight || 1;
      const s = fit === "cover" ? Math.max(w / nw, h / nh) : Math.min(w / nw, h / nh);
      ctx.drawImage(img, x + (w - nw * s) / 2, y + (h - nh * s) / 2, nw * s, nh * s);
      ctx.restore();
      resolve();
    };
    img.onerror = () => resolve();
    img.src = src;
  });
}

// Pure Canvas 2D renderer â€” replaces html2canvas entirely.
// Draws every field directly at the target pixel dimensions for full DPI quality.
async function renderFormToCanvas(response, widthPx, heightPx, inputStyles) {
  const canvas = document.createElement("canvas");
  canvas.width  = widthPx;
  canvas.height = heightPx;
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, widthPx, heightPx);

  const baseWidth  = Math.max(1, sheetEl.clientWidth  || 794);
  const baseHeight = Math.max(1, sheetEl.clientHeight || 1123);
  const scaleX = widthPx  / baseWidth;
  const scaleY = heightPx / baseHeight;
  const scale  = Math.min(scaleX, scaleY);

  for (const { field, row } of getRenderableEntries(response)) {
    // Merge live essential-tool overrides into the field definition
    const ef = { ...field, ...(inputStyles?.[field.id] || {}) };

    const fx = Math.round((Number(ef.x) || 0) * scaleX);
    const fy = Math.round((Number(ef.y) || 0) * scaleY);
    const fw = Math.round((Number(ef.w) || 160) * scaleX);
    const fh = Math.round((Number(ef.h) || 60)  * scaleY);

    ctx.save();
    if (typeof ef.rotate === "number" && ef.rotate !== 0) {
      ctx.translate(fx + fw / 2, fy + fh / 2);
      ctx.rotate((ef.rotate * Math.PI) / 180);
      ctx.translate(-(fx + fw / 2), -(fy + fh / 2));
    }
    ctx.globalAlpha = typeof ef.opacity === "number" ? ef.opacity / 100 : 1;
    ctx.beginPath(); ctx.rect(fx, fy, fw, fh); ctx.clip();

    if (ef.type === "shaperect") {
      const sw = Math.max(1, Math.round((Number(ef.strokeWidth) || 1) * scale));
      ctx.fillStyle = ef.fillColor || "#dff1eb";
      ctxRoundRect(ctx, fx, fy, fw, fh, Number(ef.radius) || 0); ctx.fill();
      ctx.strokeStyle = ef.strokeColor || "#3b7f71"; ctx.lineWidth = sw;
      ctxRoundRect(ctx, fx, fy, fw, fh, Number(ef.radius) || 0); ctx.stroke();

    } else if (ef.type === "shapecircle") {
      const sw = Math.max(1, Math.round((Number(ef.strokeWidth) || 1) * scale));
      ctx.fillStyle = ef.fillColor || "#f8ead9";
      ctx.beginPath(); ctx.ellipse(fx + fw/2, fy + fh/2, fw/2, fh/2, 0, 0, Math.PI*2); ctx.fill();
      ctx.strokeStyle = ef.strokeColor || "#ba7a36"; ctx.lineWidth = sw; ctx.stroke();

    } else if (ef.type === "shapeline") {
      const lh = Math.max(1, Math.round((Number(ef.strokeWidth) || 2) * scale));
      ctx.fillStyle = ef.strokeColor || "#4f6f67";
      ctx.fillRect(fx, fy + fh/2 - lh/2, fw, lh);

    } else if (ef.type === "divider") {
      const lh = Math.max(1, Math.round(2 * scale));
      ctx.fillStyle = "#9fb7af";
      ctx.fillRect(fx, fy + fh/2 - lh/2, fw, lh);

    } else if (ef.type === "elementstar") {
      const fs = Math.max(20, Math.round(Math.min(fw, fh) * 0.85));
      ctx.font = `${fs}px Arial`; ctx.fillStyle = ef.fillColor || "#f3ca62";
      ctx.textAlign = "center"; ctx.textBaseline = "middle";
      ctx.fillText("â˜…", fx + fw/2, fy + fh/2);

    } else if (ef.type === "elementarrow") {
      const fs = Math.max(18, Math.round(Math.min(fw, fh) * 0.75));
      ctx.font = `${fs}px Arial`; ctx.fillStyle = ef.fillColor || "#93b9ad";
      ctx.textAlign = "center"; ctx.textBaseline = "middle";
      ctx.fillText("âžœ", fx + fw/2, fy + fh/2);

    } else if (ef.type === "image" && ef.imageSrc) {
      await drawImageToCanvas(ctx, ef.imageSrc, fx, fy, fw, fh, ef.fit || "cover", Number(ef.radius) || 0);

    } else if (ef.type === "photo") {
      const src = row.images?.[0]?.dataUrl;
      if (src) await drawImageToCanvas(ctx, src, fx, fy, fw, fh, "cover", 0);

    } else if (ef.type === "sectionheader") {
      const baseFontSize = (typeof ef.fontSize === "number" && ef.fontSize > 0) ? ef.fontSize : 16;
      const fs = Math.max(8, Math.round(baseFontSize * scale));
      ctx.font = `${ef.fontStyle || "normal"} ${ef.fontWeight || "700"} ${fs}px 'Space Grotesk',Arial`;
      ctx.fillStyle = ef.textColor || "#1a1a2e";
      ctx.textAlign = ef.textAlign || "left"; ctx.textBaseline = "middle";
      const tx = ctx.textAlign === "center" ? fx+fw/2 : ctx.textAlign === "right" ? fx+fw-4 : fx+4;
      ctx.fillText(ef.value || ef.label || "Section", tx, fy + fh/2);
      const lh = Math.max(1, Math.round(2 * scale));
      ctx.fillRect(fx, fy + fh - lh, fw, lh);

    } else if (ef.type === "constant") {
      const baseFontSize = (typeof ef.fontSize === "number" && ef.fontSize > 0) ? ef.fontSize : 14;
      const fs = Math.max(8, Math.round(baseFontSize * scale));
      ctx.font = `${ef.fontStyle || "normal"} ${ef.fontWeight || "400"} ${fs}px 'Space Grotesk',Arial`;
      ctx.fillStyle = ef.textColor || "#1a1a2e";
      ctx.textAlign = ef.textAlign || "left"; ctx.textBaseline = "middle";
      const tx = ctx.textAlign === "center" ? fx+fw/2 : ctx.textAlign === "right" ? fx+fw-4 : fx+4;
      ctx.fillText(String(ef.value || ""), tx, fy + fh/2);

    } else if (ef.type === "table") {
      const rows = clamp(Math.round(Number(ef.rows) || 3), 1, 12);
      const cols = clamp(Math.round(Number(ef.cols) || 3), 1, 12);
      const matrix = Array.isArray(row.grid) ? row.grid : getTableMatrix(ef);
      const borderW = Math.max(1, Number(ef.borderWidth) || 1);
      const colTr = parseTableTrackList(ef.colWidths, cols);
      const rowTr = parseTableTrackList(ef.rowHeights, rows);
      const totalC = colTr.reduce((s, v) => s + v, 0) || 1;
      const totalR = rowTr.reduce((s, v) => s + v, 0) || 1;
      const cfs = Math.max(8, Math.round(12 * scale));
      let cy = fy;
      for (let r = 0; r < rows; r++) {
        const ch = Math.round((rowTr[r] / totalR) * fh);
        let cx = fx;
        for (let c = 0; c < cols; c++) {
          const cw = Math.round((colTr[c] / totalC) * fw);
          ctx.fillStyle = ef.cellBg || "#ffffff"; ctx.fillRect(cx, cy, cw, ch);
          ctx.strokeStyle = ef.borderColor || "#5e75b8"; ctx.lineWidth = borderW; ctx.strokeRect(cx, cy, cw, ch);
          ctx.font = `${cfs}px Arial`; ctx.fillStyle = ef.textColor || "#24366b";
          ctx.textAlign = "left"; ctx.textBaseline = "middle";
          ctx.fillText(matrix[r]?.[c] || "", cx + 4, cy + ch/2);
          cx += cw;
        }
        cy += ch;
      }

    } else {
      // All regular input fields (text, email, textarea, select, radio, checkbox, date, etc.)
      const baseFontSize = (typeof ef.fontSize === "number" && ef.fontSize > 0) ? ef.fontSize : 14;
      const fs     = Math.max(8, Math.round(baseFontSize * scale));
      const weight = String(ef.fontWeight || "400");
      const fStyle = ef.fontStyle      || "normal";
      const tDeco  = ef.textDecoration || "none";
      const tAlign = ef.textAlign      || "left";
      const pad    = Math.round(4 * scale);

      // Value only — no label, no underline border in the exported document
      let displayValue = Array.isArray(row.value) ? row.value.join(", ") : String(row.value ?? "");
      ctx.font = `${fStyle} ${weight} ${fs}px 'Space Grotesk',Arial`;
      ctx.textBaseline = "middle";

      if (!displayValue && ef.placeholder) {
        ctx.fillStyle = "#9ca3af"; ctx.textAlign = "left";
        ctx.fillText(`(${ef.placeholder})`, fx + pad, fy + fh / 2);
      } else if (displayValue) {
        ctx.fillStyle = ef.textColor || "#1a1a2e";
        ctx.textAlign = tAlign === "center" ? "center" : tAlign === "right" ? "right" : "left";
        const tx = tAlign === "center" ? fx + fw / 2 : tAlign === "right" ? fx + fw - pad : fx + pad;
        ctx.fillText(displayValue, tx, fy + fh / 2);
        // Manual underline (from Essential Tools U button only — not the field border)
        if (tDeco === "underline") {
          const tm = ctx.measureText(displayValue);
          const ulx = tAlign === "center" ? tx - tm.width / 2 : tAlign === "right" ? tx - tm.width : tx;
          ctx.fillStyle = ef.textColor || "#1a1a2e";
          ctx.fillRect(ulx, fy + fh / 2 + fs / 2 + 1, tm.width, Math.max(1, Math.round(scale)));
        }
      }
    }

    ctx.restore();
  }

  return canvas;
}

function getExportPixelSize(quality) {
  const page = currentTemplate?.page || { size: "A4", orientation: "portrait" };
  const mm   = computeSheetSize(page);
  const dpi  = quality?.dpi || exportQualityPresets.high.dpi;
  return {
    widthPx:  Math.round((mm.widthMm  / MM_PER_INCH) * dpi),
    heightPx: Math.round((mm.heightMm / MM_PER_INCH) * dpi),
    widthMm:  mm.widthMm,
    heightMm: mm.heightMm
  };
}

async function exportAsImage(response, imageType) {
  const qualitySetting = getSelectedQualityPreset();
  const { widthPx, heightPx } = getExportPixelSize(qualitySetting);
  const canvas = await renderFormToCanvas(response, widthPx, heightPx, captureInputStyles());
  const mime   = imageType === "jpg" ? "image/jpeg" : "image/png";
  const q      = imageType === "jpg" ? (qualitySetting.jpegQuality || 0.95) : 1;
  const dataUrl = canvas.toDataURL(mime, q);
  const blob    = await (await fetch(dataUrl)).blob();
  triggerDownload(blob, `${buildSafeName()}-filled.${imageType}`);
}

async function exportAsPdf(response) {
  const qualitySetting = getSelectedQualityPreset();
  const { widthPx, heightPx, widthMm, heightMm } = getExportPixelSize(qualitySetting);
  const canvas = await renderFormToCanvas(response, widthPx, heightPx, captureInputStyles());

  const pdfImageType = qualitySetting.pdfImageType || "PNG";
  const imgData = pdfImageType === "JPEG"
    ? canvas.toDataURL("image/jpeg", qualitySetting.jpegQuality || 0.95)
    : canvas.toDataURL("image/png");

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({
    orientation: widthMm >= heightMm ? "landscape" : "portrait",
    unit: "mm",
    format: [widthMm, heightMm],
    compress: true
  });
  pdf.addImage(imgData, pdfImageType, 0, 0, widthMm, heightMm, undefined, "FAST");
  pdf.save(`${buildSafeName()}-filled.pdf`);
}





async function exportAsDocx(response) {
  if (!window.docx) {
    throw new Error("DOCX library failed to load");
  }

  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType } = window.docx;
  const lines = getRenderableEntries(response);

  function makeTextRun(text, opts = {}) {
    return new TextRun({ text: String(text), ...opts });
  }

  const children = [
    new Paragraph({
      children: [makeTextRun(response.templateName || "Filled Form", { bold: true, size: 36, font: "Arial" })],
      heading: HeadingLevel.HEADING_1,
      spacing: { after: 120 }
    }),
    new Paragraph({
      children: [makeTextRun(`Generated: ${new Date(response.createdAt).toLocaleString()}`, { size: 18, color: "888888", font: "Arial" })],
      spacing: { after: 240 }
    }),
    new Paragraph({ text: "", spacing: { after: 120 } })
  ];

  lines.forEach(({ field, row }) => {
    // Skip pure layout/decoration elements in docx (they have no text value)
    if (["shaperect", "shapecircle", "shapeline", "elementstar", "elementarrow", "image"].includes(field.type)) {
      return;
    }

    if (field.type === "divider") {
      children.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "9fb7af" } },
        spacing: { before: 120, after: 120 }
      }));
      return;
    }

    if (field.type === "sectionheader") {
      children.push(new Paragraph({
        children: [makeTextRun(String(field.value || field.label || ""), { bold: true, size: 26, font: "Arial", color: "2d6a4f" })],
        spacing: { before: 240, after: 80 }
      }));
      children.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2d6a4f" } }, spacing: { after: 160 } }));
      return;
    }

    if (field.type === "constant") {
      children.push(new Paragraph({
        children: [makeTextRun(String(field.value || ""), { size: 22, font: "Arial" })],
        spacing: { after: 80 }
      }));
      return;
    }

    if (field.type === "table") {
      const matrix = Array.isArray(row.grid) ? row.grid : getTableMatrix(field);
      const numCols = matrix[0]?.length || 1;
      const tableRows = matrix.map((rowData) =>
        new TableRow({
          children: rowData.map((cellText) =>
            new TableCell({
              children: [new Paragraph({
                children: [makeTextRun(String(cellText || ""), { size: 18, font: "Arial" })]
              })],
              width: { size: Math.round(9000 / numCols), type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: (field.cellBg || "ffffff").replace("#", "") }
            })
          )
        })
      );
      children.push(new Table({ rows: tableRows, width: { size: 9000, type: WidthType.DXA } }));
      children.push(new Paragraph({ text: "", spacing: { after: 160 } }));
      return;
    }

    if (field.type === "photo") {
      const names = Array.isArray(row.value) ? row.value.join(", ") : (row.value || "");
      children.push(new Paragraph({
        children: [
          makeTextRun(`${field.label || "Photo"}: `, { bold: true, size: 20, font: "Arial" }),
          makeTextRun(names || "(image attached)", { size: 20, font: "Arial", color: "555555" })
        ],
        spacing: { after: 100 }
      }));
      return;
    }

    // Regular input fields
    const rawValue = Array.isArray(row.value) ? row.value.join(", ") : String(row.value ?? "");
    const displayValue = rawValue.trim() || "â€”";

    children.push(new Paragraph({
      children: [makeTextRun((field.label || field.type) + ":", { bold: true, size: 20, font: "Arial", color: "374151" })],
      spacing: { before: 120, after: 40 }
    }));
    children.push(new Paragraph({
      children: [makeTextRun(displayValue, { size: 22, font: "Arial" })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "d1d5db" } },
      spacing: { after: 160 }
    }));
  });

  const doc = new Document({
    creator: "Form Studio",
    description: response.templateName || "Filled Form",
    sections: [{
      properties: {},
      children
    }]
  });
  const blob = await Packer.toBlob(doc);
  triggerDownload(blob, `${buildSafeName()}-filled.docx`);
}

function exportAsTxt(response) {
  const lines = [
    `Template: ${response.templateName || "-"}`,
    `Generated: ${response.createdAt}`,
    "",
    "Values:"
  ];

  currentTemplate.fields.forEach((field) => {
    if (!supportedFieldTypes.has(field.type)) {
      return;
    }

    const row = response.values[field.id];
    const value = Array.isArray(row?.value) ? row.value.join(", ") : row?.value;
    lines.push(`- ${field.label || field.id} (${field.type}): ${value ?? ""}`);
  });

  const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
  triggerDownload(blob, `${buildSafeName()}-filled.txt`);
}

async function downloadFilledData() {
  if (!currentTemplate) {
    window.requestAnimationFrame(() => renderTemplate(currentTemplate, true));
  }

  const response = await collectResponseData();
  const format = exportFormatEl?.value || "json";

  if (format === "json") {
    const blob = new Blob([JSON.stringify(response, null, 2)], { type: "application/json" });
    triggerDownload(blob, `${buildSafeName()}-response.json`);
    return;
  }

  if (format === "txt") {
    exportAsTxt(response);
    return;
  }

  if (format === "docx") {
    await exportAsDocx(response);
    return;
  }

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
