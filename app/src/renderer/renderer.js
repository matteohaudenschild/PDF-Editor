const state = {
  serviceBaseUrl: null,
  document: null,
  currentPage: 1,
  zoom: 1,
  fitZoom: 1,
  selectedBlockId: null,
  editingBlockId: null,
  pendingFocusBlockId: null,
  drag: null,
  resize: null,
  rotate: null,
  suppressClickBlockId: null,
  lastManualTextBlockId: null,
  lastExportPath: null,
  autoSaveTimer: null,
  updateInfo: null,
  updateInstalling: false,
  backgroundLoadToken: 0,
  renderedBackgroundPage: null,
  renderedBackgroundUrl: "",
  activePdfTool: "select",
  manualTextStyle: {
    fontFamily: "Helvetica",
    fontKey: "Helvetica",
    cssFontFamily: "Arial, Helvetica, sans-serif",
    fontSize: 8,
    fontWeight: "400",
    fontStyle: "normal",
    textDecoration: "none",
  },
  documentHistory: {
    entries: [],
    index: -1,
    isRestoring: false,
  },
  templateLearning: {
    active: false,
  },
  whiteboard: {
    mode: "pen",
    isDrawing: false,
    lastPoint: null,
    color: "#000000",
    penSize: 3,
    pointerId: null,
    cursorVisible: false,
    history: [],
    historyIndex: -1,
    hasUncommittedChange: false,
    isRestoringHistory: false,
  },
};

const CUSTOM_BLOCK_DEFAULT_WIDTH = 110;
const CUSTOM_BLOCK_DEFAULT_HEIGHT = 18;
const CUSTOM_BLOCK_MIN_WIDTH = 36;
const CUSTOM_BLOCK_MIN_HEIGHT = 14;
const MANUAL_TEXT_DEFAULT_WIDTH = 85;
const MANUAL_TEXT_DEFAULT_HEIGHT = 10;
const MANUAL_TEXT_FONT_SIZE_PER_HEIGHT = 8 / MANUAL_TEXT_DEFAULT_HEIGHT;
const MANUAL_TEXT_DESCENDER_RATIO = 0.32;
const MANUAL_COVER_DEFAULT_WIDTH = MANUAL_TEXT_DEFAULT_WIDTH;
const MANUAL_COVER_DEFAULT_HEIGHT = 11;
const MANUAL_OVERLAY_MIN_SIZE = 1;
const MANUAL_CHECKBOX_SIZE = 9;
const MANUAL_CHECKBOX_MIN_SIZE = 6;
const DRAG_THRESHOLD_PX = 4;
const WHITEBOARD_HISTORY_LIMIT = 30;
const DOCUMENT_HISTORY_LIMIT = 80;
const MANUAL_TEXT_FONT_OPTIONS = [
  {
    value: "Helvetica",
    fontKey: "Helvetica",
    cssFontFamily: "Arial, Helvetica, sans-serif",
  },
  {
    value: "Times-Roman",
    fontKey: "Times-Roman",
    cssFontFamily: '"Times New Roman", Times, serif',
  },
  {
    value: "Courier",
    fontKey: "Courier",
    cssFontFamily: '"Courier New", Courier, monospace',
  },
];

const el = {
  openButton: document.getElementById("openButton"),
  webPdfInput: document.getElementById("webPdfInput"),
  templateLearningButton: document.getElementById("templateLearningButton"),
  templateNameInput: document.getElementById("templateNameInput"),
  saveTemplateButton: document.getElementById("saveTemplateButton"),
  whiteboardPenButton: document.getElementById("whiteboardPenButton"),
  whiteboardEraserButton: document.getElementById("whiteboardEraserButton"),
  whiteboardColorLabel: document.getElementById("whiteboardColorLabel"),
  whiteboardColorInput: document.getElementById("whiteboardColorInput"),
  whiteboardThicknessLabel: document.getElementById("whiteboardThicknessLabel"),
  whiteboardThicknessInput: document.getElementById("whiteboardThicknessInput"),
  whiteboardThicknessValue: document.getElementById("whiteboardThicknessValue"),
  whiteboardClearButton: document.getElementById("whiteboardClearButton"),
  whiteboardUndoButton: document.getElementById("whiteboardUndoButton"),
  whiteboardRedoButton: document.getElementById("whiteboardRedoButton"),
  selectToolButton: document.getElementById("selectToolButton"),
  addTextButton: document.getElementById("addTextButton"),
  fontFamilyLabel: document.getElementById("fontFamilyLabel"),
  fontFamilySelect: document.getElementById("fontFamilySelect"),
  fontSizeLabel: document.getElementById("fontSizeLabel"),
  fontSizeInput: document.getElementById("fontSizeInput"),
  boldButton: document.getElementById("boldButton"),
  italicButton: document.getElementById("italicButton"),
  underlineButton: document.getElementById("underlineButton"),
  eraseTextButton: document.getElementById("eraseTextButton"),
  checkButton: document.getElementById("checkButton"),
  uncheckButton: document.getElementById("uncheckButton"),
  undoButton: document.getElementById("undoButton"),
  redoButton: document.getElementById("redoButton"),
  prevButton: document.getElementById("prevButton"),
  nextButton: document.getElementById("nextButton"),
  zoomOutButton: document.getElementById("zoomOutButton"),
  zoomFitButton: document.getElementById("zoomFitButton"),
  zoomInButton: document.getElementById("zoomInButton"),
  saveDraftButton: document.getElementById("saveDraftButton"),
  exportButton: document.getElementById("exportButton"),
  revealButton: document.getElementById("revealButton"),
  updateBanner: document.getElementById("updateBanner"),
  updateMessage: document.getElementById("updateMessage"),
  updateInstallButton: document.getElementById("updateInstallButton"),
  updateDismissButton: document.getElementById("updateDismissButton"),
  status: document.getElementById("status"),
  meta: document.getElementById("meta"),
  errors: document.getElementById("errors"),
  pageArea: document.getElementById("pageArea"),
  whiteboardContainer: document.getElementById("whiteboardContainer"),
  whiteboardCanvas: document.getElementById("whiteboardCanvas"),
  whiteboardCursor: document.getElementById("whiteboardCursor"),
  pageContainer: document.getElementById("pageContainer"),
  backgroundImage: document.getElementById("backgroundImage"),
  textLayer: document.getElementById("textLayer"),
};

function setStatus(message) {
  el.status.textContent = message;
}

function ensureDocumentFieldCompatibility(documentModel) {
  if (!documentModel || typeof documentModel !== "object") {
    return documentModel;
  }

  const initialFields = Array.isArray(documentModel.fields)
    ? documentModel.fields
    : (Array.isArray(documentModel.blocks) ? documentModel.blocks : []);
  documentModel.fields = initialFields;

  const descriptor = Object.getOwnPropertyDescriptor(documentModel, "blocks");
  if (!descriptor || typeof descriptor.get !== "function" || typeof descriptor.set !== "function") {
    Object.defineProperty(documentModel, "blocks", {
      configurable: true,
      enumerable: false,
      get() {
        return this.fields;
      },
      set(value) {
        this.fields = Array.isArray(value) ? value : [];
      },
    });
  }

  return documentModel;
}

function isDesktopRuntime() {
  return Boolean(window.desktopAPI?.getServiceBaseUrl);
}

function formatBytes(bytes) {
  const numericBytes = Number(bytes) || 0;
  if (numericBytes <= 0) {
    return "";
  }
  const megabytes = numericBytes / (1024 * 1024);
  return `${megabytes.toFixed(megabytes >= 10 ? 0 : 1)} MB`;
}

function updateUpdateBanner() {
  if (!el.updateBanner || !el.updateMessage || !el.updateInstallButton) {
    return;
  }

  const info = state.updateInfo;
  const showBanner = Boolean(info?.available);
  el.updateBanner.hidden = !showBanner;
  if (!showBanner) {
    el.updateMessage.textContent = "";
    return;
  }

  const sizeText = info.assetSize ? ` (${formatBytes(info.assetSize)})` : "";
  el.updateMessage.textContent =
    `Update ${info.latestVersion} ist verfügbar${sizeText}. Deine installierte Version ist ${info.currentVersion}.`;
  el.updateInstallButton.disabled = state.updateInstalling;
  el.updateInstallButton.textContent = state.updateInstalling
    ? "Update wird vorbereitet..."
    : "Jetzt installieren";
}

async function checkForAppUpdates() {
  if (!isDesktopRuntime() || typeof window.desktopAPI?.checkForUpdates !== "function") {
    return;
  }

  const info = await window.desktopAPI.checkForUpdates();
  if (info?.error) {
    console.warn("Update check failed:", info.error);
  }
  state.updateInfo = info;
  updateUpdateBanner();
}

async function installAvailableUpdate() {
  if (!state.updateInfo?.available || state.updateInstalling) {
    return;
  }
  const confirmed = window.confirm(
    "Das Update wird heruntergeladen, der Installer startet und die App schliesst sich danach. Jetzt fortfahren?",
  );
  if (!confirmed) {
    return;
  }

  state.updateInstalling = true;
  updateUpdateBanner();
  setStatus("Update wird heruntergeladen...");
  try {
    if (hasEditableDocument()) {
      await saveDraft();
    }
    await window.desktopAPI.installUpdate();
    setStatus("Update-Installer wurde gestartet.");
  } catch (error) {
    console.error(error);
    state.updateInstalling = false;
    updateUpdateBanner();
    const fallbackText = state.updateInfo?.releaseUrl
      ? ` Bitte lade die neueste Version manuell: ${state.updateInfo.releaseUrl}`
      : "";
    setStatus(`${error.message || "Update konnte nicht installiert werden."}${fallbackText}`);
  }
}

function resolveServiceUrl(pathOrUrl) {
  if (!pathOrUrl) {
    return "";
  }
  if (/^https?:\/\//i.test(pathOrUrl)) {
    return pathOrUrl;
  }
  if (pathOrUrl.startsWith("/")) {
    return `${state.serviceBaseUrl}${pathOrUrl}`;
  }
  return `${state.serviceBaseUrl}/${pathOrUrl}`;
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return window.btoa(binary);
}

function getDownloadFilename(contentDisposition, fallbackName) {
  if (typeof contentDisposition !== "string" || !contentDisposition.trim()) {
    return fallbackName;
  }

  const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match?.[1]) {
    return decodeURIComponent(utf8Match[1]);
  }

  const simpleMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
  if (simpleMatch?.[1]) {
    return simpleMatch[1];
  }

  return fallbackName;
}

function triggerBlobDownload(blob, filename) {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => {
    URL.revokeObjectURL(objectUrl);
  }, 1000);
}

async function readErrorMessage(response, fallbackMessage) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const payload = await response.json().catch(() => null);
    return payload?.detail || fallbackMessage;
  }
  const text = await response.text().catch(() => "");
  return text || fallbackMessage;
}

function normalizeFontName(fontName) {
  let normalized = String(fontName || "").replaceAll("_", "-").trim().toLowerCase();
  if (normalized.includes("+")) {
    const [prefix, suffix] = normalized.split("+", 2);
    if (prefix.length === 6 && /^[a-z]+$/.test(prefix)) {
      normalized = suffix;
    }
  }
  const aliases = {
    "arialmt": "arial",
    "arial-boldmt": "arial-bold",
    "arial-italicmt": "arial-italic",
    "arial-bolditalicmt": "arial-bolditalic",
    "helv": "helvetica",
  };
  return aliases[normalized] || normalized;
}

function inferFontFaceStyle(fontName) {
  const normalized = normalizeFontName(fontName);
  return {
    fontWeight: normalized.includes("bold") ? "700" : "400",
    fontStyle: normalized.includes("italic") || normalized.includes("oblique") ? "italic" : "normal",
  };
}

function isBoldFontWeight(fontWeight) {
  const value = String(fontWeight || "").trim().toLowerCase();
  if (!value) {
    return false;
  }
  if (value.includes("bold")) {
    return true;
  }
  const numericValue = Number(value);
  return Number.isFinite(numericValue) && numericValue >= 600;
}

function getManualDisplayFontWeight(block) {
  return isManualTextBlock(block) && isBoldFontWeight(block.fontWeight)
    ? "700"
    : (block.fontWeight || "400");
}

function getTextDecorationThickness(block, scale) {
  if (String(block.textDecoration || "").toLowerCase() !== "underline") {
    return "";
  }
  const fontSize = (Number(block.fontSize) || 8) * scale;
  return isBoldFontWeight(block.fontWeight)
    ? `${Math.max(1.6, fontSize * 0.18)}px`
    : `${Math.max(1, fontSize * 0.08)}px`;
}

function getManualTextFontOption(fontFamily) {
  const normalized = normalizeFontName(fontFamily);
  if (normalized.includes("times")) {
    return MANUAL_TEXT_FONT_OPTIONS.find((option) => option.value === "Times-Roman") || MANUAL_TEXT_FONT_OPTIONS[0];
  }
  if (normalized.includes("courier") || normalized.startsWith("cour")) {
    return MANUAL_TEXT_FONT_OPTIONS.find((option) => option.value === "Courier") || MANUAL_TEXT_FONT_OPTIONS[0];
  }
  return MANUAL_TEXT_FONT_OPTIONS.find((option) => option.value === "Helvetica") || MANUAL_TEXT_FONT_OPTIONS[0];
}

function normalizeManualTextFontSize(value) {
  const numericValue = Number(value);
  const fallbackValue = Number(state.manualTextStyle.fontSize) || 8;
  const fontSize = Number.isFinite(numericValue) && numericValue > 0 ? numericValue : fallbackValue;
  return clamp(Math.round(fontSize * 10) / 10, 1, 72);
}

function getManualTextLineHeight(fontSize) {
  return Math.max(1, Math.round(fontSize * 1.05 * 10) / 10);
}

function getManualTextDescenderPadding(fontSize) {
  return Math.max(1, Math.round(fontSize * MANUAL_TEXT_DESCENDER_RATIO * 10) / 10);
}

function getManualTextBaselineOffset(fontSize) {
  return Math.max(1, Math.round(fontSize * 0.74 * 10) / 10);
}

function clampManualTextBaselineForBox(bbox, fontSize, preferredBaseline) {
  const y0 = Number(bbox?.y0) || 0;
  const y1 = Number(bbox?.y1) || y0;
  const boxHeight = Math.max(0, y1 - y0);
  const fallbackBaseline = y0 + Math.max(1, Math.min(boxHeight, getManualTextBaselineOffset(fontSize)));
  const rawBaseline = Number.isFinite(Number(preferredBaseline)) ? Number(preferredBaseline) : fallbackBaseline;
  if (boxHeight <= 0) {
    return rawBaseline;
  }

  const descenderPadding = getManualTextDescenderPadding(fontSize);
  const minBaseline = y0 + Math.min(boxHeight, Math.max(1, fontSize * 0.5));
  const maxBaseline = y0 + Math.max(1, boxHeight - descenderPadding);
  return clamp(rawBaseline, Math.min(minBaseline, maxBaseline), Math.max(minBaseline, maxBaseline));
}

function getManualTextBaselineForBox(bbox, fontSize, lineHeight) {
  const boxHeight = Math.max(0, (Number(bbox?.y1) || 0) - (Number(bbox?.y0) || 0));
  const textLineHeight = Math.max(1, Number(lineHeight) || getManualTextLineHeight(fontSize));
  const editorTop = Math.max(0, boxHeight - textLineHeight);
  return clampManualTextBaselineForBox(
    bbox,
    fontSize,
    (Number(bbox?.y0) || 0) + editorTop + getManualTextBaselineOffset(fontSize),
  );
}

function syncManualTextBaselineToBox(block) {
  if (!isManualTextBlock(block)) {
    return;
  }
  block.baseline = clampManualTextBaselineForBox(
    block.bbox,
    Number(block.fontSize) || 8,
    Number.isFinite(Number(block.baseline))
      ? Number(block.baseline)
      : getManualTextBaselineForBox(block.bbox, block.fontSize, block.lineHeight),
  );
}

function getManualTextFontSizeForHeight(height) {
  const safeHeight = Math.max(MANUAL_OVERLAY_MIN_SIZE, Number(height) || MANUAL_TEXT_DEFAULT_HEIGHT);
  return normalizeManualTextFontSize(safeHeight * MANUAL_TEXT_FONT_SIZE_PER_HEIGHT);
}

function hasManualTextContent(block) {
  return Boolean(String(block?.currentText || "").trim());
}

function syncManualTextSizeToBox(block) {
  if (!isManualTextBlock(block) || hasManualTextContent(block)) {
    return;
  }

  const height = block.bbox.y1 - block.bbox.y0;
  const fontSize = getManualTextFontSizeForHeight(height);
  block.fontSize = fontSize;
  block.lineHeight = getManualTextLineHeight(fontSize);
  block.baseline = getManualTextBaselineForBox(block.bbox, block.fontSize, block.lineHeight);
  block.minFontSize = 1;
}

function syncManualTextEditorBox(editor, block, scale) {
  if (!(editor instanceof HTMLElement) || !isManualTextBlock(block)) {
    return;
  }

  syncManualTextBaselineToBox(block);
  const baselineOffset = getManualTextBaselineOffset(Number(block.fontSize) || 8);
  const editorTop = Math.max(0, (Number(block.baseline) || block.bbox.y0) - block.bbox.y0 - baselineOffset) * scale;
  const visibleHeight = getBlockHeight(block, scale);
  editor.style.fontSize = `${block.fontSize * scale}px`;
  editor.style.lineHeight = `${block.lineHeight * scale}px`;
  editor.style.textDecorationThickness = getTextDecorationThickness(block, scale);
  editor.style.top = `${editorTop}px`;
  editor.style.bottom = "auto";
  editor.style.height = `${Math.max(block.lineHeight * scale, visibleHeight - editorTop)}px`;
}

function createBlockId(prefix) {
  if (window.crypto?.randomUUID) {
    return `${prefix}-${window.crypto.randomUUID()}`;
  }
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getBlockWidth(block, scale) {
  const minWidth = (isManualTextBlock(block) || isManualCoverBlock(block)) ? MANUAL_OVERLAY_MIN_SIZE : 4;
  return Math.max(minWidth, (block.bbox.x1 - block.bbox.x0) * scale);
}

function getBlockHeight(block, scale) {
  const minHeight = (isManualTextBlock(block) || isManualCoverBlock(block)) ? MANUAL_OVERLAY_MIN_SIZE : 4;
  return Math.max(minHeight, (block.bbox.y1 - block.bbox.y0) * scale);
}

function measureTextWidth(text, block, scale) {
  const canvas = measureTextWidth.canvas || (measureTextWidth.canvas = document.createElement("canvas"));
  const context = canvas.getContext("2d");
  if (!context) {
    return 0;
  }
  const fontSize = block.fontSize * scale;
  const fontStyle = block.fontStyle || "normal";
  const fontWeight = block.fontWeight || "400";
  context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${block.cssFontFamily}`;
  return context.measureText(text).width;
}

function getMaskedTemplateLayout(template, block, scale) {
  const characters = Array.from(template);
  const layout = [];
  let previousRight = 0;

  for (let index = 0; index < characters.length; index += 1) {
    const right = measureTextWidth(template.slice(0, index + 1), block, scale);
    layout.push({
      index,
      char: characters[index],
      left: previousRight,
      width: Math.max(1, right - previousRight),
    });
    previousRight = right;
  }

  return {
    characters: layout,
    totalWidth: previousRight,
  };
}

function getScaledMaskCharacterLayout(template, block, scale, totalWidth) {
  const layout = getMaskedTemplateLayout(template, block, scale);
  const measuredWidth = layout.totalWidth;
  const widthScale = measuredWidth > 0 ? (totalWidth / measuredWidth) : 1;
  return layout.characters.map((entry) => ({
    ...entry,
    left: entry.left * widthScale,
    width: Math.max(1, entry.width * widthScale),
  }));
}

function appendPositionedTemplateCharacters(target, characters) {
  let activeRun = null;
  const flushRun = () => {
    if (!activeRun) {
      return;
    }
    const node = document.createElement("span");
    node.className = "masked-template-char";
    node.textContent = activeRun.text;
    node.style.left = `${activeRun.left}px`;
    node.style.width = `${Math.max(1, activeRun.right - activeRun.left)}px`;
    target.appendChild(node);
    activeRun = null;
  };

  for (const entry of characters) {
    if (entry.char === " ") {
      flushRun();
      continue;
    }

    if (!activeRun) {
      activeRun = {
        text: entry.char,
        left: entry.left,
        right: entry.left + entry.width,
      };
      continue;
    }

    activeRun.text += entry.char;
    activeRun.right = entry.left + entry.width;
  }

  flushRun();
}

function getMaskBounds(template) {
  const first = template.indexOf("_");
  const last = template.lastIndexOf("_");
  if (first < 0 || last < first) {
    return null;
  }
  return { first, last };
}

function getMaskedSlotRegion(template, block, scale) {
  const bounds = getMaskBounds(template);
  const blockWidth = getBlockWidth(block, scale);
  const measuredTemplateWidth = measureTextWidth(template, block, scale);
  const totalWidth = Math.max(blockWidth, Math.ceil(measuredTemplateWidth));
  const widthScale = measuredTemplateWidth > 0 ? (totalWidth / measuredTemplateWidth) : 1;
  if (!bounds) {
    return {
      prefix: "",
      slotTemplate: template,
      suffix: "",
      startIndex: 0,
      endIndex: Math.max(0, template.length - 1),
      left: 0,
      width: Math.max(4, totalWidth),
      totalWidth,
    };
  }

  const prefix = template.slice(0, bounds.first);
  const suffix = template.slice(bounds.last + 1);
  const left = clamp(
    measureTextWidth(prefix, block, scale) * widthScale,
    0,
    Math.max(0, totalWidth - 4),
  );
  const suffixWidth = measureTextWidth(suffix, block, scale) * widthScale;
  const right = clamp(totalWidth - suffixWidth, left + 4, totalWidth);

  return {
    prefix,
    slotTemplate: template.slice(bounds.first, bounds.last + 1),
    suffix,
    startIndex: bounds.first,
    endIndex: bounds.last,
    left,
    width: Math.max(4, Math.ceil(right - left)),
    totalWidth,
  };
}

function getMaskedRuns(template, block, scale, totalWidth) {
  const characters = getScaledMaskCharacterLayout(template, block, scale, totalWidth);
  const runs = [];
  let activeRun = null;

  for (const entry of characters) {
    if (entry.char !== "_") {
      if (activeRun) {
        runs.push(activeRun);
        activeRun = null;
      }
      continue;
    }

    if (!activeRun) {
      activeRun = {
        startIndex: entry.index,
        endIndex: entry.index,
        left: entry.left,
        right: entry.left + entry.width,
      };
      continue;
    }

    activeRun.endIndex = entry.index;
    activeRun.right = entry.left + entry.width;
  }

  if (activeRun) {
    runs.push(activeRun);
  }

  return runs.map((run) => ({
    ...run,
    width: Math.max(4, run.right - run.left),
    template: template.slice(run.startIndex, run.endIndex + 1),
  }));
}

function getMaskedGroups(template, block, scale, totalWidth) {
  const characters = getScaledMaskCharacterLayout(template, block, scale, totalWidth);
  const groups = [];
  let activeGroup = null;

  for (const entry of characters) {
    if (entry.char === "_") {
      if (!activeGroup) {
        activeGroup = {
          startIndex: entry.index,
          endIndex: entry.index,
          left: entry.left,
          right: entry.left + entry.width,
        };
      } else {
        activeGroup.endIndex = entry.index;
        activeGroup.right = entry.left + entry.width;
      }
      continue;
    }

    if (activeGroup && entry.char !== " ") {
      groups.push(activeGroup);
      activeGroup = null;
    }
  }

  if (activeGroup) {
    groups.push(activeGroup);
  }

  return groups.map((group) => ({
    ...group,
    width: Math.max(4, group.right - group.left),
    template: template.slice(group.startIndex, group.endIndex + 1),
  }));
}

function getIbanTemplateSlotPairs(template) {
  const underscoreIndexes = [];
  for (let index = 0; index < template.length; index += 1) {
    if (template[index] === "_") {
      underscoreIndexes.push(index);
    }
  }

  const slots = [];
  for (let index = 0; index < underscoreIndexes.length; index += 2) {
    const startIndex = underscoreIndexes[index];
    const endIndex = underscoreIndexes[index + 1] ?? startIndex;
    slots.push({ startIndex, endIndex });
  }
  return slots;
}

function getIbanVisualSlots(template, block, scale, totalWidth) {
  const characters = getScaledMaskCharacterLayout(template, block, scale, totalWidth)
    .filter((entry) => entry.char === "_");
  const slots = [];
  for (let index = 0; index < characters.length; index += 2) {
    const first = characters[index];
    const second = characters[index + 1] ?? first;
    slots.push({
      startIndex: first.index,
      endIndex: second.index,
      left: first.left,
      width: Math.max(2, (second.left + second.width) - first.left),
      template: template.slice(first.index, second.index + 1),
    });
  }
  return slots;
}

function isUnderlineTemplateBlock(block) {
  return !block.isCustom
    && !block.isCheckbox
    && typeof block.originalText === "string"
    && block.originalText.includes("__")
    && !block.originalText.includes("\n");
}

function getMaskedTemplate(block) {
  return block.originalText || "";
}

function isGermanIbanTemplateBlock(block) {
  const template = getMaskedTemplate(block).toUpperCase();
  return template.includes("IBAN") && template.includes("DE");
}

function extractMaskedValue(template, currentText) {
  const safeCurrent = typeof currentText === "string" && currentText.length ? currentText : template;
  let value = "";
  for (let index = 0; index < template.length; index += 1) {
    if (template[index] !== "_") {
      continue;
    }
    const char = safeCurrent[index] ?? "_";
    if (char !== "_") {
      value += char;
    }
  }
  return value;
}

function applyMaskedValue(template, value) {
  const chars = Array.from(template);
  let valueIndex = 0;
  for (let index = 0; index < chars.length; index += 1) {
    if (chars[index] !== "_") {
      continue;
    }
    chars[index] = valueIndex < value.length ? value[valueIndex] : "_";
    valueIndex += 1;
  }
  return chars.join("");
}

function sanitizeMaskedValue(block, value) {
  const rawValue = String(value || "");
  if (isGermanIbanTemplateBlock(block)) {
    return rawValue.replace(/\D+/g, "");
  }
  return rawValue;
}

function sanitizeIbanSlotValue(value) {
  return String(value || "").replace(/\D+/g, "").slice(0, 1);
}

function normalizeIbanCurrentText(template, rawCurrent) {
  const slots = getIbanTemplateSlotPairs(template);
  const safeCurrent = typeof rawCurrent === "string" ? rawCurrent : "";
  const digits = [];

  if (safeCurrent.length === template.length) {
    const chars = Array.from(safeCurrent);
    for (const slot of slots) {
      const slotChars = chars.slice(slot.startIndex, slot.endIndex + 1);
      digits.push(slotChars.find((char) => /\d/.test(char)) ?? "");
    }
  } else {
    digits.push(...safeCurrent.replace(/\D+/g, "").slice(0, slots.length));
  }

  const chars = Array.from(template);
  for (let slotIndex = 0; slotIndex < slots.length; slotIndex += 1) {
    const slot = slots[slotIndex];
    const digit = /\d/.test(digits[slotIndex] || "") ? digits[slotIndex] : "_";
    chars[slot.startIndex] = digit;
    for (let index = slot.startIndex + 1; index <= slot.endIndex; index += 1) {
      chars[index] = "_";
    }
  }

  return chars.join("");
}

function normalizeMaskedBlockByPosition(block) {
  const template = getMaskedTemplate(block);
  const safeCurrent = typeof block.currentText === "string" && block.currentText.length === template.length
    ? block.currentText
    : template;
  const chars = Array.from(template);
  for (let index = 0; index < chars.length; index += 1) {
    if (chars[index] !== "_") {
      continue;
    }
    const currentChar = safeCurrent[index] ?? "_";
    chars[index] = currentChar === "_" ? "_" : currentChar;
  }
  block.currentText = chars.join("");
}

function normalizeMaskedBlockText(block) {
  const template = getMaskedTemplate(block);
  if (!isGermanIbanTemplateBlock(block)) {
    normalizeMaskedBlockByPosition(block);
    return;
  }
  block.currentText = normalizeIbanCurrentText(template, block.currentText);
}

function getMaskedOverlayText(template, currentText) {
  const safeCurrent = currentText && currentText.length === template.length
    ? currentText
    : applyMaskedValue(template, extractMaskedValue(template, currentText));
  const chars = [];
  for (let index = 0; index < template.length; index += 1) {
    if (template[index] === "_") {
      chars.push(safeCurrent[index] && safeCurrent[index] !== "_" ? safeCurrent[index] : " ");
    } else {
      chars.push(" ");
    }
  }
  return chars.join("");
}

function isFixedGeneratedField(block) {
  return typeof block.groupKind === "string"
    && (block.groupKind.startsWith("generated-") || block.groupKind.startsWith("widget-"));
}

function isContractIdNumberBlock(block) {
  const blockId = String(block?.id || "").toLowerCase();
  const groupKind = String(block?.groupKind || "").toLowerCase();
  if (blockId.includes("creditor-id")) {
    return false;
  }
  const text = String(block?.currentText || block?.originalText || "").trim();
  return blockId.includes("generated-id-number")
    || blockId.includes("generated-instruction-id")
    || blockId.endsWith("-id-number")
    || groupKind === "generated-id-number-field"
    || /^200\d{7,10}$/.test(text);
}

function syncContractIdUnderline(underline, block, scale) {
  const text = String(block?.currentText || block?.originalText || "").trim();
  if (!text) {
    underline.style.display = "none";
    return;
  }

  const boldBlock = { ...block, fontWeight: "700" };
  const underlineWidth = Math.max(1, Math.ceil(measureTextWidth(text, boldBlock, scale)));
  const baseline = typeof block.baseline === "number"
    ? (block.baseline - block.bbox.y0) * scale
    : block.fontSize * 0.88 * scale;
  const underlineTop = baseline + Math.max(1, block.fontSize * 0.11 * scale);
  const thickness = Math.max(1, Math.round(0.75 * scale));

  underline.style.display = "block";
  underline.style.left = "0px";
  underline.style.top = `${underlineTop}px`;
  underline.style.width = `${underlineWidth}px`;
  underline.style.height = `${thickness}px`;
  underline.style.background = block.color || "#000000";
}

function isNormalizedGeneratedDocument() {
  const templateId = String(state.document?.detectedTemplateId || "").trim().toLowerCase();
  const templateFamily = String(state.document?.detectedTemplateFamily || "").trim().toLowerCase();
  return templateId === "sicherheit_nord_vt_handlungsanweisung_scan_9696"
    || templateFamily === "sicherheit_nord_vt_handlungsanweisung";
}

function isScanTemplateDocument() {
  if (isNormalizedGeneratedDocument()) {
    return false;
  }

  const templateId = String(state.document?.detectedTemplateId || "").trim().toLowerCase();
  if (templateId.includes("scan")) {
    return true;
  }

  const warnings = state.document?.supportStatus?.warnings ?? [];
  return warnings.some((warning) => {
    const normalized = String(warning || "").trim().toLowerCase();
    return normalized.includes("scan") || normalized.includes("gescannt");
  });
}

function isScanHeaderField(block) {
  return block?.groupKind === "generated-contract-party-field"
    || block?.groupKind === "generated-contract-object-line-field";
}

function isScanFieldChanged(block) {
  return String(block?.currentText || "") !== String(block?.originalText || "");
}

function isUnchangedScanGeneratedTextField(block) {
  return isScanGeneratedTextField(block)
    && !isContractIdNumberBlock(block)
    && !isScanFieldChanged(block)
    && Boolean(String(block?.currentText || "").trim());
}

function isUnchangedScanGeneratedCheckbox(block) {
  return isScanTemplateDocument()
    && isFixedGeneratedField(block)
    && Boolean(block?.isCheckbox)
    && !isScanFieldChanged(block)
    && Boolean(String(block?.currentText || "").trim());
}

function shouldRenderScanOriginalCover(block) {
  return isScanTemplateDocument()
    && !block?.isCustom
    && !block?.isCheckbox
    && typeof block?.groupKind === "string"
    && block.groupKind.startsWith("generated-")
    && (isScanFieldChanged(block) || isContractIdNumberBlock(block))
    && Boolean(String(block.currentText || "").trim() || String(block.originalText || "").trim());
}

function shouldRenderScanCheckboxCover(block) {
  return isScanTemplateDocument()
    && Boolean(block?.isCheckbox)
    && typeof block?.groupKind === "string"
    && block.groupKind.startsWith("generated-")
    && isScanFieldChanged(block)
    && Boolean(String(block.currentText || "").trim() || String(block.originalText || "").trim());
}

function clampRectToPage(rect, page) {
  return {
    x0: Math.max(0, Math.min(rect.x0, page.width)),
    y0: Math.max(0, Math.min(rect.y0, page.height)),
    x1: Math.max(0, Math.min(rect.x1, page.width)),
    y1: Math.max(0, Math.min(rect.y1, page.height)),
  };
}

function getScanCheckboxInnerRect(block, page) {
  const rect = {
    x0: block.bbox.x0,
    y0: block.bbox.y0,
    x1: block.bbox.x1,
    y1: block.bbox.y1,
  };
  const inset = Math.max(0.25, Math.min(rect.x1 - rect.x0, rect.y1 - rect.y0) * 0.03);
  return clampRectToPage({
    x0: rect.x0 + inset,
    y0: rect.y0 + inset,
    x1: rect.x1 - inset,
    y1: rect.y1 - inset,
  }, page);
}

function getScanReplacementRect(block, page) {
  const baseline = typeof block.baseline === "number" ? block.baseline : block.bbox.y1;
  if (isContractIdNumberBlock(block)) {
    const text = String(block.currentText || block.originalText || "").trim();
    const textWidth = text ? measureTextWidth(text, { ...block, fontWeight: "700" }, 1) : 0;
    const top = Math.min(block.bbox.y0, baseline - (block.fontSize * 1.05)) - 0.8;
    const bottom = Math.max(block.bbox.y1, baseline + (block.fontSize * 0.42)) + 1.2;
    return clampRectToPage({
      x0: block.bbox.x0 - 1.0,
      y0: top,
      x1: Math.max(block.bbox.x1, block.bbox.x0 + textWidth + Math.max(4.0, block.fontSize * 0.45)) + 1.0,
      y1: bottom,
    }, page);
  }
  if (isScanHeaderField(block)) {
    const scaleX = page.width / 595.0;
    const x0 = block.groupKind === "generated-contract-party-field" ? (98.5 * scaleX) : (285.6 * scaleX);
    const x1 = block.groupKind === "generated-contract-party-field" ? (286.0 * scaleX) : (565.0 * scaleX);
    return clampRectToPage({
      x0: x0 - 0.8,
      y0: block.bbox.y0 - 1.2,
      x1: x1 + 0.8,
      y1: block.bbox.y1 + 1.8,
    }, page);
  }

  const top = Math.min(block.bbox.y0, baseline - (block.fontSize * 1.05)) - 0.5;
  const bottom = Math.max(block.bbox.y1, baseline + (block.fontSize * 0.42)) + 0.5;
  let x1 = block.bbox.x0;
  for (const text of [block.originalText, block.currentText]) {
    const trimmed = String(text || "").trim();
    if (!trimmed) {
      continue;
    }
    x1 = Math.min(block.bbox.x1, Math.max(x1, block.bbox.x0 + measureTextWidth(trimmed, block, 1) + 5.0));
  }
  return clampRectToPage({
    x0: block.bbox.x0 - 0.8,
    y0: top,
    x1: x1 + 0.8,
    y1: bottom,
  }, page);
}

function appendScanCoverNode(rect, scale) {
  const node = document.createElement("div");
  node.className = "scan-cover";
  node.style.position = "absolute";
  node.style.left = `${rect.x0 * scale}px`;
  node.style.top = `${rect.y0 * scale}px`;
  node.style.width = `${Math.max(1, (rect.x1 - rect.x0) * scale)}px`;
  node.style.height = `${Math.max(1, (rect.y1 - rect.y0) * scale)}px`;
  node.style.background = "#ffffff";
  node.style.pointerEvents = "none";
  node.style.zIndex = "1";
  el.textLayer.appendChild(node);
}

function appendScanGuideLine(x0, y0, x1, y1, scale) {
  const node = document.createElement("div");
  node.className = "scan-guide-line";
  node.style.position = "absolute";
  const isHorizontal = Math.abs(y0 - y1) <= Math.abs(x0 - x1);
  const thickness = 1;
  if (isHorizontal) {
    node.style.left = `${Math.min(x0, x1) * scale}px`;
    node.style.top = `${(Math.min(y0, y1) * scale) - (thickness / 2)}px`;
    node.style.width = `${Math.max(1, Math.abs(x1 - x0) * scale)}px`;
    node.style.height = `${thickness}px`;
  } else {
    node.style.left = `${(Math.min(x0, x1) * scale) - (thickness / 2)}px`;
    node.style.top = `${Math.min(y0, y1) * scale}px`;
    node.style.width = `${thickness}px`;
    node.style.height = `${Math.max(1, Math.abs(y1 - y0) * scale)}px`;
  }
  node.style.background = "#000000";
  node.style.pointerEvents = "none";
  node.style.zIndex = "1";
  el.textLayer.appendChild(node);
}

function renderScanTemplateCovers(page, blocks, scale) {
  if (!isScanTemplateDocument()) {
    return;
  }

  let hasHeaderField = false;
  for (const block of blocks) {
    if (shouldRenderScanCheckboxCover(block)) {
      appendScanCoverNode(getScanCheckboxInnerRect(block, page), scale);
      continue;
    }
    if (!shouldRenderScanOriginalCover(block)) {
      continue;
    }
    appendScanCoverNode(getScanReplacementRect(block, page), scale);
    hasHeaderField = hasHeaderField || isScanHeaderField(block);
  }

  if (!hasHeaderField) {
    return;
  }

  const scaleX = page.width / 595.0;
  const scaleY = page.height / 842.0;
  const sx = (value) => value * scaleX;
  const sy = (value) => value * scaleY;
  const rows = [55.05, 65.45, 75.75, 86.10, 96.45, 106.95];
  const leftX0 = 28.4;
  const leftSplit = 98.5;
  const leftX1 = 286.0;
  const rightX0 = 285.6;
  const rightX1 = 565.0;

  for (const y of rows) {
    appendScanGuideLine(sx(leftX0), sy(y), sx(leftX1), sy(y), scale);
    appendScanGuideLine(sx(rightX0), sy(y), sx(rightX1), sy(y), scale);
  }

  for (const [x0, y0, x1, y1] of [
    [leftX0, rows[0], leftX0, rows[rows.length - 1]],
    [leftSplit, rows[0], leftSplit, rows[rows.length - 1]],
    [leftX1, rows[0], leftX1, rows[rows.length - 1]],
    [rightX0, rows[0], rightX0, rows[rows.length - 1]],
    [rightX1, rows[0], rightX1, rows[rows.length - 1]],
  ]) {
    appendScanGuideLine(sx(x0), sy(y0), sx(x1), sy(y1), scale);
  }
}

function autoSizeTextBlock(node, editor, block, scale) {
  if (isFixedGeneratedField(block)) {
    node.style.width = `${getBlockWidth(block, scale)}px`;
    node.style.height = `${getBlockHeight(block, scale)}px`;
    return;
  }

  if (isManualOverlayBlock(block)) {
    node.style.width = `${getBlockWidth(block, scale)}px`;
    node.style.height = `${getBlockHeight(block, scale)}px`;
    return;
  }

  if (block.isCustom) {
    const page = state.document?.pages.find((entry) => entry.pageNumber === block.page) ?? null;
    const currentWidth = Math.max(CUSTOM_BLOCK_MIN_WIDTH, block.bbox.x1 - block.bbox.x0);
    const currentHeight = Math.max(CUSTOM_BLOCK_MIN_HEIGHT, block.bbox.y1 - block.bbox.y0);
    const requiredWidth = Math.max(currentWidth, Math.ceil((editor.scrollWidth + 2) / scale));
    const requiredHeight = Math.max(currentHeight, Math.ceil((editor.scrollHeight + 2) / scale));
    const maxWidth = page ? Math.max(CUSTOM_BLOCK_MIN_WIDTH, page.width - block.bbox.x0) : requiredWidth;
    const nextWidth = Math.min(maxWidth, requiredWidth);

    if (nextWidth > currentWidth) {
      block.bbox.x1 = block.bbox.x0 + nextWidth;
    }
    if (requiredHeight > currentHeight) {
      block.bbox.y1 = block.bbox.y0 + requiredHeight;
    }

    node.style.width = `${Math.max(currentWidth, nextWidth) * scale}px`;
    node.style.height = `${Math.max(currentHeight, requiredHeight) * scale}px`;
    return;
  }

  const baseWidth = getBlockWidth(block, scale);
  const baseHeight = getBlockHeight(block, scale);

  node.style.width = `${baseWidth}px`;
  node.style.height = `${baseHeight}px`;

  const neededWidth = Math.max(baseWidth, editor.scrollWidth + 2);
  const neededHeight = Math.max(baseHeight, editor.scrollHeight + 2);

  node.style.width = `${Math.ceil(neededWidth)}px`;
  node.style.height = `${Math.ceil(neededHeight)}px`;
}

function growCustomBlockByLine(node, block, scale, lineCount = 1) {
  const lineStep = Math.max(CUSTOM_BLOCK_MIN_HEIGHT, Math.ceil(block.lineHeight));
  const currentHeight = Math.max(CUSTOM_BLOCK_MIN_HEIGHT, block.bbox.y1 - block.bbox.y0);
  block.bbox.y1 = block.bbox.y0 + currentHeight + (lineStep * lineCount);
  node.style.height = `${(block.bbox.y1 - block.bbox.y0) * scale}px`;
}

function ensureEditablePlaceholder(node) {
  if (node.textContent === "") {
    node.innerHTML = "<br>";
  }
}

function readEditableText(node) {
  const text = node.innerText.replace(/\r/g, "");
  if (text === "\n") {
    return "";
  }
  return text;
}

function placeCaretAtEnd(node) {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  range.selectNodeContents(node);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

function selectEditableText(node) {
  if (node instanceof HTMLInputElement || node instanceof HTMLTextAreaElement) {
    node.setSelectionRange(0, node.value.length);
    return;
  }

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  range.selectNodeContents(node);
  selection.removeAllRanges();
  selection.addRange(range);
}

function focusEditableTarget(target, options = {}) {
  const { selectAll = false } = options;
  if (target.classList.contains("text-editor")) {
    ensureEditablePlaceholder(target);
  }

  target.focus();
  if (selectAll) {
    selectEditableText(target);
    return;
  }

  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
    target.setSelectionRange(target.value.length, target.value.length);
    return;
  }

  placeCaretAtEnd(target);
}

function isScanGeneratedTextField(block) {
  return isScanTemplateDocument()
    && isFixedGeneratedField(block)
    && !block?.isCheckbox
    && !isUnderlineTemplateBlock(block);
}

function syncSelectedBlock() {
  const nodes = el.textLayer.querySelectorAll(".text-block");
  for (const node of nodes) {
    node.classList.toggle("selected", node.dataset.blockId === state.selectedBlockId);
    node.classList.toggle("manual-editing", node.dataset.blockId === state.editingBlockId);
  }

  if (state.pendingFocusBlockId) {
    const target = el.textLayer.querySelector(
      `[data-block-id="${state.pendingFocusBlockId}"] .text-editor, `
      + `[data-block-id="${state.pendingFocusBlockId}"] .masked-template-capture`,
    );
    if (target) {
      state.pendingFocusBlockId = null;
      const blockId = target.closest(".text-block")?.dataset.blockId ?? null;
      const targetBlock = blockId ? getBlockById(blockId) : null;
      requestAnimationFrame(() => {
        focusEditableTarget(target, { selectAll: isScanGeneratedTextField(targetBlock) });
      });
    }
  }

  updateMeta();
  syncManualTextStyleControls();
}

function syncDocumentBlockValues() {
  for (const block of state.document?.blocks || []) {
    block.currentValue = block.currentText || "";
    if (typeof block.originalValue !== "string") {
      block.originalValue = block.originalText || "";
    }
  }
}

function getCurrentPageModel() {
  return state.document?.pages.find((page) => page.pageNumber === state.currentPage) ?? null;
}

function getCurrentPageBlocks() {
  return (state.document?.blocks ?? []).filter((block) => block.page === state.currentPage);
}

function getSelectedBlock() {
  return getCurrentPageBlocks().find((block) => block.id === state.selectedBlockId) ?? null;
}

function getSelectedManualTextBlock() {
  const selected = getSelectedBlock();
  return isManualTextBlock(selected) ? selected : null;
}

function getActiveManualTextStyleSource() {
  return getSelectedManualTextBlock() || state.manualTextStyle;
}

function setPressedState(button, active) {
  if (!button) {
    return;
  }
  button.classList.toggle("active-text-style", active);
  button.setAttribute("aria-pressed", active ? "true" : "false");
}

function updateManualTextStyleControlsVisibility(hasDoc) {
  const showTextStyleControls = hasDoc && Boolean(getSelectedManualTextBlock());
  document.body.classList.toggle("text-style-mode", showTextStyleControls);
  const elements = [
    el.fontFamilyLabel,
    el.fontFamilySelect,
    el.fontSizeLabel,
    el.fontSizeInput,
    el.boldButton,
    el.italicButton,
    el.underlineButton,
  ];
  for (const element of elements) {
    if (!element) {
      continue;
    }
    element.hidden = !showTextStyleControls;
    if ("disabled" in element) {
      element.disabled = !showTextStyleControls;
    }
  }
}

function syncManualTextStyleControls() {
  const hasDoc = hasEditableDocument();
  updateManualTextStyleControlsVisibility(hasDoc);
  if (!hasDoc) {
    return;
  }

  const style = getActiveManualTextStyleSource();
  const fontOption = getManualTextFontOption(style.fontFamily);
  if (el.fontFamilySelect) {
    el.fontFamilySelect.value = fontOption.value;
  }
  if (el.fontSizeInput) {
    el.fontSizeInput.value = String(normalizeManualTextFontSize(style.fontSize));
  }
  setPressedState(el.boldButton, isBoldFontWeight(style.fontWeight));
  setPressedState(el.italicButton, String(style.fontStyle || "").toLowerCase() === "italic");
  setPressedState(el.underlineButton, String(style.textDecoration || "").toLowerCase() === "underline");
}

function applyManualTextStylePatch(patch) {
  const nextPatch = {};

  if (Object.prototype.hasOwnProperty.call(patch, "fontFamily")) {
    const fontOption = getManualTextFontOption(patch.fontFamily);
    nextPatch.fontFamily = fontOption.value;
    nextPatch.fontKey = fontOption.fontKey;
    nextPatch.cssFontFamily = fontOption.cssFontFamily;
    nextPatch.fontAssetId = null;
  }

  if (Object.prototype.hasOwnProperty.call(patch, "fontSize")) {
    const fontSize = normalizeManualTextFontSize(patch.fontSize);
    nextPatch.fontSize = fontSize;
    nextPatch.lineHeight = getManualTextLineHeight(fontSize);
  }

  if (Object.prototype.hasOwnProperty.call(patch, "fontWeight")) {
    nextPatch.fontWeight = isBoldFontWeight(patch.fontWeight) ? "700" : "400";
  }

  if (Object.prototype.hasOwnProperty.call(patch, "fontStyle")) {
    nextPatch.fontStyle = String(patch.fontStyle || "").toLowerCase() === "italic" ? "italic" : "normal";
  }

  if (Object.prototype.hasOwnProperty.call(patch, "textDecoration")) {
    nextPatch.textDecoration = String(patch.textDecoration || "").toLowerCase() === "underline" ? "underline" : "none";
  }

  Object.assign(state.manualTextStyle, nextPatch);

  const selected = getSelectedManualTextBlock();
  if (selected) {
    Object.assign(selected, nextPatch);
    if (Object.prototype.hasOwnProperty.call(nextPatch, "fontSize")) {
      selected.baseline = getManualTextBaselineForBox(selected.bbox, selected.fontSize, selected.lineHeight);
      if (!isManualTextBlock(selected) && (selected.bbox.y1 - selected.bbox.y0) < nextPatch.lineHeight) {
        selected.bbox.y1 = selected.bbox.y0 + nextPatch.lineHeight;
      }
    }
    rememberManualTextBlock(selected);
    renderTextLayer();
    commitDocumentHistory();
    scheduleDraftSave();
  }

  syncManualTextStyleControls();
}

function getBlockById(blockId) {
  return state.document?.blocks.find((block) => block.id === blockId) ?? null;
}

function findTemplateBlock(x, y) {
  const blocks = getCurrentPageBlocks().filter((block) => !block.isCustom && !block.isCheckbox && block.editable !== false);
  if (!blocks.length) {
    return state.document?.blocks?.find((block) => !block.isCustom && !block.isCheckbox && block.editable !== false) ?? null;
  }

  let best = blocks[0];
  let bestDistance = Number.POSITIVE_INFINITY;

  for (const block of blocks) {
    const centerX = (block.bbox.x0 + block.bbox.x1) / 2;
    const centerY = (block.bbox.y0 + block.bbox.y1) / 2;
    const dx = centerX - x;
    const dy = centerY - y;
    const distance = (dx * dx) + (dy * dy);
    if (distance < bestDistance) {
      best = block;
      bestDistance = distance;
    }
  }

  return best;
}

function findTemplateBlockNearPoint(x, y, options = {}) {
  const {
    screenPaddingX = 8,
    screenPaddingY = 10,
  } = options;
  const blocks = getCurrentPageBlocks().filter((block) => !block.isCustom && !block.isCheckbox && block.editable !== false);
  if (!blocks.length) {
    return null;
  }

  const paddingX = screenPaddingX / Math.max(state.zoom, 0.01);
  const paddingY = screenPaddingY / Math.max(state.zoom, 0.01);
  const candidates = blocks.filter((block) => (
    x >= (block.bbox.x0 - paddingX)
    && x <= (block.bbox.x1 + paddingX)
    && y >= (block.bbox.y0 - paddingY)
    && y <= (block.bbox.y1 + paddingY)
  ));
  if (!candidates.length) {
    return null;
  }

  let best = candidates[0];
  let bestDistance = Number.POSITIVE_INFINITY;
  for (const block of candidates) {
    const centerX = (block.bbox.x0 + block.bbox.x1) / 2;
    const centerY = (block.bbox.y0 + block.bbox.y1) / 2;
    const dx = centerX - x;
    const dy = centerY - y;
    const distance = (dx * dx) + (dy * dy);
    if (distance < bestDistance) {
      best = block;
      bestDistance = distance;
    }
  }

  return best;
}

function createDefaultTemplateBlock(pageNumber) {
  return {
    id: `default-template-page-${pageNumber}`,
    page: pageNumber,
    bbox: {
      x0: 0,
      y0: 0,
      x1: CUSTOM_BLOCK_DEFAULT_WIDTH,
      y1: CUSTOM_BLOCK_DEFAULT_HEIGHT,
    },
    originalText: "",
    currentText: "",
    fontFamily: "Helvetica",
    fontKey: "Helvetica",
    fontSize: 10,
    color: "#000000",
    lineHeight: 12,
    align: "left",
    rotation: 0,
    groupKind: "fallback-template",
    minFontSize: 6,
    editable: false,
    cssFontFamily: "Arial, sans-serif",
    fontAssetId: null,
    fontWeight: "400",
    fontStyle: "normal",
    textDecoration: "none",
    baseline: 8.2,
    isCheckbox: false,
    isCustom: false,
  };
}

function toggleCheckbox(block) {
  if (!block?.isCheckbox) {
    return;
  }
  block.currentText = block.currentText.trim() ? "" : "x";
  block.currentValue = block.currentText;
  renderTextLayer();
  commitDocumentHistory();
  scheduleDraftSave();
}

function isManualOverlayBlock(block) {
  return Boolean(block?.isCustom) && String(block?.groupKind || "").startsWith("manual-");
}

function isManualCoverBlock(block) {
  return isManualOverlayBlock(block) && String(block?.groupKind || "") === "manual-cover";
}

function isManualTextBlock(block) {
  return isManualOverlayBlock(block) && String(block?.groupKind || "") === "manual-text";
}

function isManualCheckboxBlock(block) {
  return isManualOverlayBlock(block) && String(block?.groupKind || "") === "manual-checkbox";
}

function rememberManualTextBlock(block) {
  if (isManualTextBlock(block)) {
    state.lastManualTextBlockId = block.id;
  }
}

function getManualTextBlockFromElement(element) {
  const owner = element instanceof HTMLElement ? element.closest(".text-block") : null;
  const block = owner?.dataset.blockId ? getBlockById(owner.dataset.blockId) : null;
  return isManualTextBlock(block) ? block : null;
}

function getManualTextBlockForKeyboard(activeElement = document.activeElement, options = {}) {
  const { includeLast = true } = options;
  const candidates = [
    getManualTextBlockFromElement(activeElement),
    state.editingBlockId ? getBlockById(state.editingBlockId) : null,
    state.selectedBlockId ? getBlockById(state.selectedBlockId) : null,
    includeLast && state.lastManualTextBlockId ? getBlockById(state.lastManualTextBlockId) : null,
  ];
  const seen = new Set();
  for (const block of candidates) {
    if (!isManualTextBlock(block) || seen.has(block.id) || block.page !== state.currentPage) {
      continue;
    }
    seen.add(block.id);
    return block;
  }
  return null;
}

function setActivePdfTool(tool) {
  state.activePdfTool = ["select", "text", "erase", "check", "uncheck"].includes(tool) ? tool : "select";
  updateButtons();
}

function clampToPage(page, value, size, padding = 2) {
  return Math.min(Math.max(value, 0), Math.max(0, (page - size) - padding));
}

function createManualOverlayBlock(event, tool = state.activePdfTool) {
  if (!state.document?.supportStatus?.supported) {
    return;
  }
  if (typeof event.button === "number" && event.button !== 0) {
    return;
  }

  const page = getCurrentPageModel();
  const containerRect = el.pageContainer.getBoundingClientRect();
  const x = (event.clientX - containerRect.left) / state.zoom;
  const y = (event.clientY - containerRect.top) / state.zoom;
  if (!page) {
    return;
  }
  const template = findTemplateBlock(x, y) || createDefaultTemplateBlock(state.currentPage);

  const isTextTool = tool === "text";
  const isCheckboxTool = tool === "check" || tool === "uncheck";
  const isEraseTool = tool === "erase";
  const manualTextStyle = state.manualTextStyle;
  const manualFontOption = getManualTextFontOption(manualTextStyle.fontFamily);
  const manualFontSize = normalizeManualTextFontSize(manualTextStyle.fontSize);
  const manualLineHeight = getManualTextLineHeight(manualFontSize);
  const width = isCheckboxTool
    ? MANUAL_CHECKBOX_SIZE
    : (isEraseTool ? MANUAL_COVER_DEFAULT_WIDTH : MANUAL_TEXT_DEFAULT_WIDTH);
  const height = isCheckboxTool
    ? MANUAL_CHECKBOX_SIZE
    : (isEraseTool ? MANUAL_COVER_DEFAULT_HEIGHT : MANUAL_TEXT_DEFAULT_HEIGHT);
  const x0 = clampToPage(page.width, isCheckboxTool ? x - (width / 2) : x, width);
  const y0 = clampToPage(page.height, isCheckboxTool ? y - (height / 2) : y, height);
  const baselineOffset = typeof template.baseline === "number"
      ? template.baseline - template.bbox.y0
      : Math.max(template.fontSize * 0.82, 1);
  const text = tool === "check" ? "x" : "";
  const groupKind = {
    text: "manual-text",
    erase: "manual-cover",
    check: "manual-checkbox",
    uncheck: "manual-checkbox",
  }[tool] || "manual-text";

  const block = {
    id: createBlockId(`custom-page-${state.currentPage}`),
    page: state.currentPage,
    fieldType: isCheckboxTool ? "checkbox" : "text-line",
    bbox: {
      x0,
      y0,
      x1: x0 + width,
      y1: y0 + height,
    },
    originalText: "",
    currentText: text,
    originalValue: "",
    currentValue: text,
    fontFamily: isCheckboxTool ? "Helvetica" : (isTextTool ? manualFontOption.value : template.fontFamily),
    fontKey: isCheckboxTool ? "Helvetica" : (isTextTool ? manualFontOption.fontKey : template.fontKey),
    fontSize: isCheckboxTool ? MANUAL_CHECKBOX_SIZE : (isTextTool ? manualFontSize : template.fontSize),
    color: template.color,
    lineHeight: isCheckboxTool ? MANUAL_CHECKBOX_SIZE : (isTextTool ? manualLineHeight : template.lineHeight),
    align: template.align,
    rotation: 0,
    groupKind,
    minFontSize: isTextTool ? 1 : template.minFontSize,
    editable: true,
    cssFontFamily: isCheckboxTool ? "Arial, sans-serif" : (isTextTool ? manualFontOption.cssFontFamily : template.cssFontFamily),
    fontAssetId: isCheckboxTool || isTextTool ? null : template.fontAssetId,
    fontWeight: isCheckboxTool ? "700" : (isTextTool ? manualTextStyle.fontWeight : template.fontWeight),
    fontStyle: isTextTool ? manualTextStyle.fontStyle : template.fontStyle,
    textDecoration: isTextTool ? manualTextStyle.textDecoration : (template.textDecoration || "none"),
    baseline: isTextTool
      ? getManualTextBaselineForBox({ y0, y1: y0 + height }, manualFontSize, manualLineHeight)
      : y0 + baselineOffset,
    isCheckbox: isCheckboxTool,
    isCustom: true,
  };

  state.document.blocks.push(block);
  rememberManualTextBlock(block);
  state.selectedBlockId = tool === "text" ? null : block.id;
  state.editingBlockId = null;
  state.pendingFocusBlockId = null;
  renderTextLayer();
  commitDocumentHistory();
  scheduleDraftSave();
  if (isTextTool || isEraseTool || isCheckboxTool) {
    setActivePdfTool("select");
  }
  setStatus(
    tool === "text"
      ? "Weißes Textfeld gesetzt. Zum Schreiben anklicken."
      : (tool === "erase" ? "Weiße Löschfläche gesetzt." : "Ankreuzfeld gesetzt."),
  );
}

function createCustomTextBlock(event) {
  createManualOverlayBlock(event, "text");
}

function beginPotentialDrag(block, node, event) {
  if (!block.isCustom || event.button !== 0) {
    return;
  }

  state.drag = {
    blockId: block.id,
    node,
    pageNumber: block.page,
    startClientX: event.clientX,
    startClientY: event.clientY,
    originX0: block.bbox.x0,
    originY0: block.bbox.y0,
    originBaseline: typeof block.baseline === "number" ? block.baseline : null,
    width: block.bbox.x1 - block.bbox.x0,
    height: block.bbox.y1 - block.bbox.y0,
    dragging: false,
  };
}

function clearDragState() {
  if (state.drag?.node) {
    state.drag.node.classList.remove("manual-overlay-dragging");
  }
  state.drag = null;
  document.body.classList.remove("dragging-block");
}

function beginResize(block, node, event, handle = "se") {
  if (!block.isCustom || event.button !== 0) {
    return;
  }

  state.resize = {
    blockId: block.id,
    node,
    pageNumber: block.page,
    startClientX: event.clientX,
    startClientY: event.clientY,
    originWidth: block.bbox.x1 - block.bbox.x0,
    originHeight: block.bbox.y1 - block.bbox.y0,
    originX0: block.bbox.x0,
    originY0: block.bbox.y0,
    originX1: block.bbox.x1,
    originY1: block.bbox.y1,
    originBaseline: typeof block.baseline === "number" ? block.baseline : null,
    originBaselineOffset: typeof block.baseline === "number" ? block.baseline - block.bbox.y0 : null,
    handle,
    handleNode: event.currentTarget instanceof HTMLElement ? event.currentTarget : null,
    resizing: false,
  };
}

function clearResizeState() {
  if (state.resize?.handleNode) {
    state.resize.handleNode.classList.remove("manual-resize-active");
  }
  if (state.resize?.node) {
    state.resize.node.classList.remove("manual-overlay-resizing");
  }
  state.resize = null;
  document.body.classList.remove("resizing-block");
}

function getPagePointFromClient(event) {
  const containerRect = el.pageContainer.getBoundingClientRect();
  return {
    x: (event.clientX - containerRect.left) / state.zoom,
    y: (event.clientY - containerRect.top) / state.zoom,
  };
}

function normalizeRotation(degrees) {
  let value = degrees % 360;
  if (value > 180) {
    value -= 360;
  }
  if (value < -180) {
    value += 360;
  }
  return Math.round(value * 10) / 10;
}

function beginRotate(block, node, event) {
  if (!block.isCustom || event.button !== 0) {
    return;
  }

  const centerX = (block.bbox.x0 + block.bbox.x1) / 2;
  const centerY = (block.bbox.y0 + block.bbox.y1) / 2;
  const point = getPagePointFromClient(event);
  state.rotate = {
    blockId: block.id,
    node,
    pageNumber: block.page,
    centerX,
    centerY,
    startAngle: Math.atan2(point.y - centerY, point.x - centerX) * 180 / Math.PI,
    originRotation: Number(block.rotation) || 0,
    rotating: false,
  };
}

function clearRotateState() {
  state.rotate = null;
  document.body.classList.remove("rotating-block");
}

function handleDragMove(event) {
  if (state.resize || state.rotate || !state.drag || !state.document?.supportStatus?.supported) {
    return;
  }

  const drag = state.drag;
  const dx = event.clientX - drag.startClientX;
  const dy = event.clientY - drag.startClientY;

  if (!drag.dragging) {
    if (Math.abs(dx) < DRAG_THRESHOLD_PX && Math.abs(dy) < DRAG_THRESHOLD_PX) {
      return;
    }
    drag.dragging = true;
    document.body.classList.add("dragging-block");
    drag.node?.classList.add("manual-overlay-dragging");
    const selection = window.getSelection();
    selection?.removeAllRanges();
  }

  const page = state.document.pages.find((entry) => entry.pageNumber === drag.pageNumber);
  const block = getBlockById(drag.blockId);
  if (!page || !block) {
    clearDragState();
    return;
  }

  event.preventDefault();

  const nextX0 = clampToPage(page.width, drag.originX0 + (dx / state.zoom), drag.width);
  const nextY0 = clampToPage(page.height, drag.originY0 + (dy / state.zoom), drag.height);

  block.bbox.x0 = nextX0;
  block.bbox.x1 = nextX0 + drag.width;
  block.bbox.y0 = nextY0;
  block.bbox.y1 = nextY0 + drag.height;
  if (typeof drag.originBaseline === "number") {
    block.baseline = drag.originBaseline + (dy / state.zoom);
  }

  if (drag.node) {
    drag.node.style.left = `${nextX0 * state.zoom}px`;
    drag.node.style.top = `${nextY0 * state.zoom}px`;
  }
}

function handleResizeMove(event) {
  if (!state.resize || !state.document?.supportStatus?.supported) {
    return;
  }

  const resize = state.resize;
  const dx = event.clientX - resize.startClientX;
  const dy = event.clientY - resize.startClientY;

  if (!resize.resizing) {
    if (Math.abs(dx) < DRAG_THRESHOLD_PX && Math.abs(dy) < DRAG_THRESHOLD_PX) {
      return;
    }
    resize.resizing = true;
    document.body.classList.add("resizing-block");
    resize.node?.classList.add("manual-overlay-resizing");
    resize.handleNode?.classList.add("manual-resize-active");
    const selection = window.getSelection();
    selection?.removeAllRanges();
  }

  const page = state.document.pages.find((entry) => entry.pageNumber === resize.pageNumber);
  const block = getBlockById(resize.blockId);
  if (!page || !block) {
    clearResizeState();
    return;
  }

  event.preventDefault();

  const isFlexibleManualOverlay = isManualTextBlock(block) || isManualCoverBlock(block);
  const minWidth = isManualCheckboxBlock(block)
    ? MANUAL_CHECKBOX_MIN_SIZE
    : isFlexibleManualOverlay
    ? MANUAL_OVERLAY_MIN_SIZE
    : CUSTOM_BLOCK_MIN_WIDTH;
  const minHeight = isManualCheckboxBlock(block)
    ? MANUAL_CHECKBOX_MIN_SIZE
    : isFlexibleManualOverlay
    ? MANUAL_OVERLAY_MIN_SIZE
    : CUSTOM_BLOCK_MIN_HEIGHT;
  let nextX0 = resize.originX0;
  let nextY0 = resize.originY0;
  let nextX1 = resize.originX1;
  let nextY1 = resize.originY1;
  const docDx = dx / state.zoom;
  const docDy = dy / state.zoom;

  if (resize.handle.includes("w")) {
    nextX0 = Math.min(resize.originX1 - minWidth, resize.originX0 + docDx);
  }
  if (resize.handle.includes("e")) {
    nextX1 = Math.max(resize.originX0 + minWidth, resize.originX1 + docDx);
  }
  if (resize.handle.includes("n")) {
    nextY0 = Math.min(resize.originY1 - minHeight, resize.originY0 + docDy);
  }
  if (resize.handle.includes("s")) {
    nextY1 = Math.max(resize.originY0 + minHeight, resize.originY1 + docDy);
  }

  nextX0 = clamp(nextX0, 0, page.width - minWidth);
  nextY0 = clamp(nextY0, 0, page.height - minHeight);
  nextX1 = clamp(nextX1, nextX0 + minWidth, page.width);
  nextY1 = clamp(nextY1, nextY0 + minHeight, page.height);

  block.bbox.x0 = nextX0;
  block.bbox.y0 = nextY0;
  block.bbox.x1 = nextX1;
  block.bbox.y1 = nextY1;
  syncManualTextSizeToBox(block);
  if (isManualTextBlock(block)) {
    block.baseline = getManualTextBaselineForBox(block.bbox, block.fontSize, block.lineHeight);
  } else if (typeof resize.originBaselineOffset === "number") {
    block.baseline = nextY0 + resize.originBaselineOffset;
  }

  if (resize.node) {
    resize.node.style.left = `${nextX0 * state.zoom}px`;
    resize.node.style.top = `${nextY0 * state.zoom}px`;
    resize.node.style.width = `${(nextX1 - nextX0) * state.zoom}px`;
    resize.node.style.height = `${(nextY1 - nextY0) * state.zoom}px`;
    syncManualTextEditorBox(resize.node.querySelector(".text-editor"), block, state.zoom);
  }
}

function handleRotateMove(event) {
  if (!state.rotate || !state.document?.supportStatus?.supported) {
    return;
  }

  const rotate = state.rotate;
  const point = getPagePointFromClient(event);
  const currentAngle = Math.atan2(point.y - rotate.centerY, point.x - rotate.centerX) * 180 / Math.PI;
  const delta = currentAngle - rotate.startAngle;

  if (!rotate.rotating) {
    if (Math.abs(delta) < 1.5) {
      return;
    }
    rotate.rotating = true;
    document.body.classList.add("rotating-block");
    const selection = window.getSelection();
    selection?.removeAllRanges();
  }

  const block = getBlockById(rotate.blockId);
  if (!block) {
    clearRotateState();
    return;
  }

  event.preventDefault();
  block.rotation = normalizeRotation(rotate.originRotation + delta);
  if (rotate.node) {
    rotate.node.style.transform = `rotate(${block.rotation}deg)`;
  }
}

function handleDragEnd() {
  if (state.rotate) {
    const wasRotating = state.rotate.rotating;
    const blockId = state.rotate.blockId;
    clearRotateState();
    if (wasRotating) {
      state.suppressClickBlockId = blockId;
      commitDocumentHistory();
      scheduleDraftSave();
    }
  }

  if (state.resize) {
    const wasResizing = state.resize.resizing;
    const blockId = state.resize.blockId;
    clearResizeState();
    if (wasResizing) {
      state.suppressClickBlockId = blockId;
      syncManualTextStyleControls();
      commitDocumentHistory();
      scheduleDraftSave();
    }
  }

  if (state.drag) {
    const wasDragging = state.drag.dragging;
    const blockId = state.drag.blockId;
    clearDragState();
    if (wasDragging) {
      state.suppressClickBlockId = blockId;
      commitDocumentHistory();
      scheduleDraftSave();
    }
  }
}

function removeSelectedCustomBlock() {
  const selected = getSelectedBlock();
  if (!selected?.isCustom || !state.document) {
    return false;
  }

  state.document.blocks = state.document.blocks.filter((block) => block.id !== selected.id);
  if (state.lastManualTextBlockId === selected.id) {
    state.lastManualTextBlockId = null;
  }
  state.selectedBlockId = null;
  state.editingBlockId = null;
  state.pendingFocusBlockId = null;
  renderTextLayer();
  commitDocumentHistory();
  scheduleDraftSave();
  setStatus("Manuelles Feld entfernt.");
  return true;
}

function removeManualTextBlockIfEmpty(block, editor = null) {
  if (!isManualTextBlock(block) || !state.document) {
    return false;
  }

  const text = editor instanceof HTMLElement
    ? readEditableText(editor)
    : (block.currentText || "");
  if (text.trim()) {
    return false;
  }

  state.selectedBlockId = block.id;
  state.editingBlockId = null;
  block.currentText = "";
  block.currentValue = "";
  if (editor instanceof HTMLElement) {
    editor.blur();
  }
  return removeSelectedCustomBlock();
}

function handleDeleteForEmptyManualTextBlock(event) {
  if (event.key !== "Delete" && event.key !== "Backspace") {
    return false;
  }

  const active = document.activeElement;
  const block = getManualTextBlockForKeyboard(active, { includeLast: false });
  const activeEditor = active instanceof HTMLElement
    && active.classList.contains("text-editor")
    && getManualTextBlockFromElement(active)?.id === block?.id
    ? active
    : null;

  if (!removeManualTextBlockIfEmpty(block, activeEditor)) {
    return false;
  }

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  return true;
}

function isKeyboardDeletableManualOverlayBlock(block) {
  return isManualCoverBlock(block) || isManualCheckboxBlock(block);
}

function handleDeleteForSelectedManualOverlayBlock(event) {
  if (event.key !== "Delete" && event.key !== "Backspace") {
    return false;
  }

  const active = document.activeElement;
  if (
    active instanceof HTMLElement
    && (active.classList.contains("text-editor") || active.classList.contains("masked-template-capture"))
  ) {
    return false;
  }

  const activeOwner = active instanceof HTMLElement ? active.closest(".text-block") : null;
  const activeBlock = activeOwner?.dataset.blockId ? getBlockById(activeOwner.dataset.blockId) : null;
  const block = isKeyboardDeletableManualOverlayBlock(activeBlock) ? activeBlock : getSelectedBlock();
  if (!isKeyboardDeletableManualOverlayBlock(block)) {
    return false;
  }

  state.selectedBlockId = block.id;
  state.editingBlockId = null;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  return removeSelectedCustomBlock();
}

function clearSelection(options = {}) {
  const { blurEditor = true } = options;
  state.selectedBlockId = null;
  state.pendingFocusBlockId = null;

  if (blurEditor) {
    const active = document.activeElement;
    if (
      active instanceof HTMLElement
      && (active.classList.contains("text-editor") || active.classList.contains("masked-template-capture"))
    ) {
      active.blur();
    }
  }

  state.editingBlockId = null;
  syncSelectedBlock();
}

function hasEditableDocument() {
  return Boolean(state.document?.supportStatus?.supported);
}

function clonePlainData(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function createDocumentHistorySnapshot() {
  if (!hasEditableDocument()) {
    return null;
  }

  return {
    currentPage: state.currentPage,
    selectedBlockId: state.selectedBlockId,
    lastManualTextBlockId: state.lastManualTextBlockId,
    manualTextStyle: clonePlainData(state.manualTextStyle),
    blocks: clonePlainData(state.document.blocks || []),
  };
}

function getDocumentHistoryKey(snapshot) {
  if (!snapshot) {
    return "";
  }
  return JSON.stringify({
    manualTextStyle: snapshot.manualTextStyle,
    blocks: snapshot.blocks,
  });
}

function canUndoDocument() {
  return hasEditableDocument()
    && !state.documentHistory.isRestoring
    && state.documentHistory.index > 0;
}

function canRedoDocument() {
  return hasEditableDocument()
    && !state.documentHistory.isRestoring
    && state.documentHistory.index >= 0
    && state.documentHistory.index < state.documentHistory.entries.length - 1;
}

function resetDocumentHistory() {
  state.documentHistory.entries = [];
  state.documentHistory.index = -1;
  state.documentHistory.isRestoring = false;

  const snapshot = createDocumentHistorySnapshot();
  if (snapshot) {
    state.documentHistory.entries.push({
      snapshot,
      key: getDocumentHistoryKey(snapshot),
    });
    state.documentHistory.index = 0;
  }

  updateButtons();
}

function commitDocumentHistory() {
  if (!hasEditableDocument() || state.documentHistory.isRestoring) {
    return false;
  }

  const snapshot = createDocumentHistorySnapshot();
  if (!snapshot) {
    return false;
  }

  const key = getDocumentHistoryKey(snapshot);
  const currentEntry = state.documentHistory.entries[state.documentHistory.index] ?? null;
  if (currentEntry?.key === key) {
    updateButtons();
    return false;
  }

  if (state.documentHistory.index < state.documentHistory.entries.length - 1) {
    state.documentHistory.entries.splice(state.documentHistory.index + 1);
  }

  state.documentHistory.entries.push({ snapshot, key });
  if (state.documentHistory.entries.length > DOCUMENT_HISTORY_LIMIT) {
    const overflow = state.documentHistory.entries.length - DOCUMENT_HISTORY_LIMIT;
    state.documentHistory.entries.splice(0, overflow);
  }

  state.documentHistory.index = state.documentHistory.entries.length - 1;
  updateButtons();
  return true;
}

async function restoreDocumentHistorySnapshot(snapshot) {
  if (!snapshot || !state.document) {
    return;
  }

  const pageCount = Math.max(1, Number(state.document.pageCount) || 1);
  const nextPage = clamp(Math.round(Number(snapshot.currentPage) || 1), 1, pageCount);
  const pageChanged = state.currentPage !== nextPage || state.renderedBackgroundPage !== nextPage;
  state.currentPage = nextPage;
  state.document.blocks = clonePlainData(snapshot.blocks) || [];
  state.manualTextStyle = {
    ...state.manualTextStyle,
    ...(clonePlainData(snapshot.manualTextStyle) || {}),
  };

  const blockExists = (blockId) => Boolean(blockId && state.document.blocks.some((block) => block.id === blockId));
  state.selectedBlockId = blockExists(snapshot.selectedBlockId) ? snapshot.selectedBlockId : null;
  state.editingBlockId = null;
  state.pendingFocusBlockId = null;
  state.suppressClickBlockId = null;
  state.lastManualTextBlockId = blockExists(snapshot.lastManualTextBlockId)
    ? snapshot.lastManualTextBlockId
    : null;

  syncDocumentBlockValues();
  updateButtons();
  updateMeta();
  if (pageChanged) {
    await loadBackground({ preserveCurrent: true });
  } else {
    renderTextLayer();
  }
  scheduleDraftSave();
}

async function moveDocumentHistory(step) {
  if (state.documentHistory.isRestoring || !hasEditableDocument()) {
    return;
  }

  const nextIndex = state.documentHistory.index + step;
  if (nextIndex < 0 || nextIndex >= state.documentHistory.entries.length) {
    return;
  }

  state.documentHistory.isRestoring = true;
  clearDragState();
  clearResizeState();
  clearRotateState();
  state.documentHistory.index = nextIndex;
  updateButtons();

  try {
    await restoreDocumentHistorySnapshot(state.documentHistory.entries[nextIndex]?.snapshot);
    setStatus(step < 0 ? "Ein Schritt zurück." : "Ein Schritt vor.");
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  } finally {
    state.documentHistory.isRestoring = false;
    updateButtons();
  }
}

function hasManualTemplateFields() {
  return Boolean(state.document?.blocks?.some((block) => block.isCustom));
}

function getDefaultTemplateName() {
  const sourcePath = String(state.document?.sourcePath || "").trim();
  const sourceName = sourcePath.split(/[\\/]/).pop() || "neue-vorlage";
  const stem = sourceName.replace(/\.pdf$/i, "").trim();
  return stem || "neue-vorlage";
}

function getTemplateNameValue() {
  const rawValue = String(el.templateNameInput?.value || "").trim();
  return rawValue || getDefaultTemplateName();
}

function setTemplateLearningActive(active) {
  state.templateLearning.active = Boolean(active) && hasEditableDocument();
  updateButtons();
  updateMeta();
}

function isWhiteboardMode() {
  return !hasEditableDocument();
}

function getWhiteboardPenSize() {
  return clamp(Math.round(Number(state.whiteboard.penSize) || 3), 1, 24);
}

function getWhiteboardStrokeWidth() {
  return state.whiteboard.mode === "eraser" ? 18 : getWhiteboardPenSize();
}

function canUndoWhiteboard() {
  return state.whiteboard.historyIndex > 0;
}

function canRedoWhiteboard() {
  return state.whiteboard.historyIndex >= 0
    && state.whiteboard.historyIndex < state.whiteboard.history.length - 1;
}

function createWhiteboardSnapshot() {
  const canvas = el.whiteboardCanvas;
  if (!canvas.width || !canvas.height) {
    return null;
  }

  return {
    dataUrl: canvas.toDataURL("image/png"),
  };
}

function resetWhiteboardHistory() {
  const snapshot = createWhiteboardSnapshot();
  state.whiteboard.history = snapshot ? [snapshot] : [];
  state.whiteboard.historyIndex = snapshot ? 0 : -1;
  state.whiteboard.hasUncommittedChange = false;
  updateButtons();
}

function commitWhiteboardHistory() {
  const snapshot = createWhiteboardSnapshot();
  if (!snapshot) {
    return false;
  }

  const currentSnapshot = state.whiteboard.history[state.whiteboard.historyIndex] ?? null;
  if (currentSnapshot?.dataUrl === snapshot.dataUrl) {
    state.whiteboard.hasUncommittedChange = false;
    updateButtons();
    return false;
  }

  if (state.whiteboard.historyIndex < state.whiteboard.history.length - 1) {
    state.whiteboard.history.splice(state.whiteboard.historyIndex + 1);
  }

  state.whiteboard.history.push(snapshot);
  if (state.whiteboard.history.length > WHITEBOARD_HISTORY_LIMIT) {
    const overflow = state.whiteboard.history.length - WHITEBOARD_HISTORY_LIMIT;
    state.whiteboard.history.splice(0, overflow);
  }

  state.whiteboard.historyIndex = state.whiteboard.history.length - 1;
  state.whiteboard.hasUncommittedChange = false;
  updateButtons();
  return true;
}

async function restoreWhiteboardSnapshot(snapshot) {
  resizeWhiteboardCanvas(false);
  const canvas = el.whiteboardCanvas;
  const context = getWhiteboardContext();
  resetWhiteboardCanvas(context, canvas.width, canvas.height);

  if (!snapshot?.dataUrl) {
    return;
  }

  const image = new Image();
  await new Promise((resolve, reject) => {
    image.onload = resolve;
    image.onerror = () => reject(new Error("Whiteboard-Verlauf konnte nicht geladen werden."));
    image.src = snapshot.dataUrl;
  });
  context.drawImage(image, 0, 0, canvas.clientWidth, canvas.clientHeight);
}

async function moveWhiteboardHistory(step) {
  if (state.whiteboard.isRestoringHistory) {
    return;
  }

  const nextIndex = state.whiteboard.historyIndex + step;
  if (nextIndex < 0 || nextIndex >= state.whiteboard.history.length) {
    return;
  }

  state.whiteboard.isRestoringHistory = true;
  stopWhiteboardDrawing();
  state.whiteboard.hasUncommittedChange = false;
  updateButtons();

  try {
    await restoreWhiteboardSnapshot(state.whiteboard.history[nextIndex]);
    state.whiteboard.historyIndex = nextIndex;
    setStatus(step < 0 ? "Whiteboard: ein Schritt zurück." : "Whiteboard: ein Schritt vor.");
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  } finally {
    state.whiteboard.isRestoringHistory = false;
    updateButtons();
  }
}

function updateWhiteboardToolButtons() {
  const visible = isWhiteboardMode();
  el.whiteboardPenButton.hidden = !visible;
  el.whiteboardEraserButton.hidden = !visible;
  el.whiteboardColorLabel.hidden = !visible;
  el.whiteboardColorInput.hidden = !visible;
  el.whiteboardThicknessLabel.hidden = !visible;
  el.whiteboardThicknessInput.hidden = !visible;
  el.whiteboardThicknessValue.hidden = !visible;
  el.whiteboardClearButton.hidden = !visible;
  el.whiteboardUndoButton.hidden = !visible;
  el.whiteboardRedoButton.hidden = !visible;
  el.whiteboardPenButton.classList.toggle("active-tool", visible && state.whiteboard.mode === "pen");
  el.whiteboardEraserButton.classList.toggle("active-tool", visible && state.whiteboard.mode === "eraser");
  el.whiteboardColorInput.disabled = !visible || state.whiteboard.mode === "eraser";
  el.whiteboardThicknessInput.disabled = !visible || state.whiteboard.mode === "eraser";
  el.whiteboardUndoButton.disabled = !visible || state.whiteboard.isRestoringHistory || !canUndoWhiteboard();
  el.whiteboardRedoButton.disabled = !visible || state.whiteboard.isRestoringHistory || !canRedoWhiteboard();
  const penSize = getWhiteboardPenSize();
  el.whiteboardColorInput.value = state.whiteboard.color;
  el.whiteboardThicknessInput.value = String(penSize);
  el.whiteboardThicknessValue.textContent = `${penSize} px`;
  el.whiteboardContainer.hidden = !visible;
  if (visible) {
    updateWhiteboardCursorAppearance();
    el.pageContainer.hidden = true;
    return;
  }
  hideWhiteboardCursor();
}

function updateWhiteboardCursorAppearance() {
  const eraserMode = state.whiteboard.mode === "eraser";
  const strokeWidth = getWhiteboardStrokeWidth();
  el.whiteboardCanvas.classList.toggle("eraser-mode", eraserMode);
  el.whiteboardCursor.classList.toggle("eraser", eraserMode);
  el.whiteboardCursor.style.setProperty(
    "--whiteboard-cursor-size",
    `${eraserMode ? strokeWidth : Math.max(12, strokeWidth + 6)}px`,
  );
  el.whiteboardCursor.style.setProperty(
    "--whiteboard-cursor-fill",
    eraserMode ? "rgba(255, 255, 255, 0.45)" : state.whiteboard.color,
  );
}

function hideWhiteboardCursor() {
  state.whiteboard.cursorVisible = false;
  el.whiteboardCursor.hidden = true;
  el.whiteboardCanvas.classList.remove("cursor-hidden");
}

function showWhiteboardCursor(point, pointerType = "") {
  if (!isWhiteboardMode() || pointerType === "touch") {
    hideWhiteboardCursor();
    return;
  }

  updateWhiteboardCursorAppearance();
  el.whiteboardCursor.style.left = `${point.x}px`;
  el.whiteboardCursor.style.top = `${point.y}px`;
  el.whiteboardCursor.hidden = false;
  state.whiteboard.cursorVisible = true;
  el.whiteboardCanvas.classList.add("cursor-hidden");
}

function getWhiteboardContext() {
  const context = el.whiteboardCanvas.getContext("2d");
  if (!context) {
    throw new Error("Whiteboard-Kontext konnte nicht erstellt werden.");
  }
  return context;
}

function resetWhiteboardCanvas(context, width, height) {
  context.save();
  context.setTransform(1, 0, 0, 1, 0, 0);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.restore();
}

function resizeWhiteboardCanvas(preserveDrawing = true) {
  const canvas = el.whiteboardCanvas;
  const containerWidth = Math.max(320, el.pageArea.clientWidth - 24);
  const containerTop = el.pageArea.getBoundingClientRect().top;
  const displayHeight = Math.max(600, Math.floor(window.innerHeight - containerTop - 36));
  const displayWidth = Math.floor(containerWidth);
  const dpr = Math.max(1, Math.min(window.devicePixelRatio || 1, 2));
  const targetWidth = Math.max(1, Math.round(displayWidth * dpr));
  const targetHeight = Math.max(1, Math.round(displayHeight * dpr));

  if (canvas.width === targetWidth && canvas.height === targetHeight) {
    return;
  }

  const snapshot = preserveDrawing && canvas.width && canvas.height
    ? (() => {
      const image = document.createElement("canvas");
      image.width = canvas.width;
      image.height = canvas.height;
      const imageContext = image.getContext("2d");
      imageContext?.drawImage(canvas, 0, 0);
      return image;
    })()
    : null;

  canvas.width = targetWidth;
  canvas.height = targetHeight;
  canvas.style.width = `${displayWidth}px`;
  canvas.style.height = `${displayHeight}px`;

  const context = getWhiteboardContext();
  context.setTransform(dpr, 0, 0, dpr, 0, 0);
  context.lineCap = "round";
  context.lineJoin = "round";
  resetWhiteboardCanvas(context, targetWidth, targetHeight);

  if (snapshot) {
    context.drawImage(snapshot, 0, 0, displayWidth, displayHeight);
  }
}

function clearWhiteboard(setMessage = true) {
  resizeWhiteboardCanvas(true);
  const context = getWhiteboardContext();
  resetWhiteboardCanvas(context, el.whiteboardCanvas.width, el.whiteboardCanvas.height);
  stopWhiteboardDrawing();
  state.whiteboard.hasUncommittedChange = false;
  if (state.whiteboard.history.length === 0) {
    resetWhiteboardHistory();
  } else {
    commitWhiteboardHistory();
  }
  if (setMessage) {
    setStatus("Whiteboard geleert.");
  }
}

function setWhiteboardMode(mode) {
  state.whiteboard.mode = mode === "eraser" ? "eraser" : "pen";
  updateWhiteboardCursorAppearance();
  updateWhiteboardToolButtons();
}

function getWhiteboardPoint(event) {
  const rect = el.whiteboardCanvas.getBoundingClientRect();
  return {
    x: clamp(event.clientX - rect.left, 0, rect.width),
    y: clamp(event.clientY - rect.top, 0, rect.height),
  };
}

function syncWhiteboardCursorFromEvent(event) {
  const point = getWhiteboardPoint(event);
  showWhiteboardCursor(point, event.pointerType);
  return point;
}

function drawWhiteboardSegment(fromPoint, toPoint) {
  const context = getWhiteboardContext();
  const strokeWidth = getWhiteboardStrokeWidth();
  context.save();
  context.strokeStyle = state.whiteboard.mode === "eraser" ? "#ffffff" : state.whiteboard.color;
  context.fillStyle = state.whiteboard.mode === "eraser" ? "#ffffff" : state.whiteboard.color;
  context.lineWidth = strokeWidth;
  context.beginPath();
  context.moveTo(fromPoint.x, fromPoint.y);
  context.lineTo(toPoint.x, toPoint.y);
  context.stroke();
  context.beginPath();
  context.arc(toPoint.x, toPoint.y, context.lineWidth / 2, 0, Math.PI * 2);
  context.fill();
  context.restore();
}

function finishWhiteboardStroke(pointerId = null) {
  const shouldCommit = state.whiteboard.hasUncommittedChange;
  stopWhiteboardDrawing(pointerId);
  if (shouldCommit) {
    commitWhiteboardHistory();
  } else {
    updateButtons();
  }
}

function stopWhiteboardDrawing(pointerId = null) {
  state.whiteboard.isDrawing = false;
  state.whiteboard.lastPoint = null;
  const activePointerId = pointerId ?? state.whiteboard.pointerId;
  if (
    activePointerId !== null
    && typeof el.whiteboardCanvas.hasPointerCapture === "function"
    && el.whiteboardCanvas.hasPointerCapture(activePointerId)
  ) {
    try {
      el.whiteboardCanvas.releasePointerCapture(activePointerId);
    } catch (error) {
      // Ignore invalid release attempts.
    }
  }
  state.whiteboard.pointerId = null;
}

async function exportWhiteboardPdf() {
  resizeWhiteboardCanvas(true);
  setStatus("Whiteboard wird als PDF exportiert...");
  const payload = {
    imageDataUrl: el.whiteboardCanvas.toDataURL("image/png"),
    width: el.whiteboardCanvas.clientWidth,
    height: el.whiteboardCanvas.clientHeight,
  };

  if (isDesktopRuntime()) {
    const targetPath = await window.desktopAPI.savePdfDialog("whiteboard.pdf");
    if (!targetPath) {
      return;
    }

    const response = await fetch(`${state.serviceBaseUrl}/whiteboard/export`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ...payload,
        targetPath,
      }),
    });

    const responsePayload = await response.json().catch(() => null);
    if (!response.ok) {
      throw new Error(responsePayload?.detail || "Whiteboard-Export fehlgeschlagen.");
    }

    state.lastExportPath = responsePayload.outputPath;
    updateButtons();
    setStatus(`Whiteboard exportiert: ${responsePayload.outputPath}`);
    return;
  }

  const response = await fetch(`${state.serviceBaseUrl}/whiteboard/export-download`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(await readErrorMessage(response, "Whiteboard-Export fehlgeschlagen."));
  }

  const filename = getDownloadFilename(
    response.headers.get("content-disposition"),
    "whiteboard.pdf",
  );
  const blob = await response.blob();
  state.lastExportPath = null;
  updateButtons();
  triggerBlobDownload(blob, filename);
  setStatus(`Whiteboard exportiert: ${filename}`);
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function updateButtons() {
  const hasDoc = hasEditableDocument();
  const hasService = Boolean(state.serviceBaseUrl);
  const toolButtons = [
    [el.selectToolButton, "select"],
    [el.addTextButton, "text"],
    [el.eraseTextButton, "erase"],
    [el.checkButton, "check"],
    [el.uncheckButton, "uncheck"],
  ];
  for (const [button, tool] of toolButtons) {
    if (!button) {
      continue;
    }
    button.hidden = !hasDoc;
    button.disabled = !hasDoc;
    button.classList.toggle("active-tool", hasDoc && state.activePdfTool === tool);
  }
  el.prevButton.hidden = !hasDoc;
  el.nextButton.hidden = !hasDoc;
  el.undoButton.hidden = !hasDoc;
  el.redoButton.hidden = !hasDoc;
  el.zoomOutButton.hidden = !hasDoc;
  el.zoomFitButton.hidden = !hasDoc;
  el.zoomInButton.hidden = !hasDoc;
  el.templateLearningButton.hidden = true;
  el.templateNameInput.hidden = true;
  el.saveTemplateButton.hidden = true;
  el.prevButton.disabled = !hasDoc || state.currentPage <= 1;
  el.nextButton.disabled = !hasDoc || state.currentPage >= (state.document?.pageCount ?? 0);
  el.undoButton.disabled = !canUndoDocument();
  el.redoButton.disabled = !canRedoDocument();
  el.zoomOutButton.disabled = !hasDoc;
  el.zoomFitButton.disabled = !hasDoc;
  el.zoomInButton.disabled = !hasDoc;
  el.templateLearningButton.disabled = !hasDoc;
  el.templateLearningButton.classList.toggle("active-tool", false);
  el.saveTemplateButton.disabled = !hasDoc || !hasManualTemplateFields();
  el.saveDraftButton.disabled = !hasDoc;
  el.exportButton.disabled = !hasService;
  el.exportButton.textContent = hasDoc ? "PDF exportieren" : "Whiteboard exportieren";
  el.revealButton.disabled = !state.lastExportPath;
  el.revealButton.hidden = !isDesktopRuntime();
  if (el.updateInstallButton) {
    el.updateInstallButton.disabled = state.updateInstalling;
  }
  updateWhiteboardToolButtons();
  syncManualTextStyleControls();
}

function updateMeta() {
  if (!hasEditableDocument()) {
    el.meta.textContent = "Whiteboard aktiv.";
    return;
  }

  const sourceName = String(state.document.sourcePath || "Dokument.pdf").split(/[\\/]/).pop() || "Dokument.pdf";
  el.meta.textContent = `${sourceName} | Seite ${state.currentPage}/${state.document.pageCount}`;
}

function showSupportStatus() {
  if (!state.document) {
    el.errors.textContent = "";
    el.errors.hidden = true;
    return;
  }

  const { reasons = [] } = state.document.supportStatus || {};
  const lines = [];
  if (reasons.length) {
    lines.push("Ablehnung:");
    reasons.forEach((reason) => lines.push(`- ${reason}`));
  }
  el.errors.textContent = lines.join("\n");
  el.errors.hidden = lines.length === 0;
}

function injectFontFaces() {
  const existing = document.getElementById("dynamic-fonts");
  if (existing) {
    existing.remove();
  }

  if (!state.document?.fonts?.length) {
    return;
  }

  const style = document.createElement("style");
  style.id = "dynamic-fonts";
  style.textContent = state.document.fonts
    .filter((font) => font.loadUrl)
    .map((font) => {
      const face = inferFontFaceStyle(font.family);
      return `@font-face { font-family: "${font.cssFamily}"; src: url("${resolveServiceUrl(font.loadUrl)}"); font-weight: ${face.fontWeight}; font-style: ${face.fontStyle}; }`;
    })
    .join("\n");
  document.head.appendChild(style);
}

function applyPageDisplaySize(page, displayWidth, displayHeight) {
  el.backgroundImage.style.width = `${displayWidth}px`;
  el.backgroundImage.style.height = `${displayHeight}px`;
  el.pageContainer.style.width = `${displayWidth}px`;
  el.pageContainer.style.height = `${displayHeight}px`;
  el.textLayer.style.width = `${displayWidth}px`;
  el.textLayer.style.height = `${displayHeight}px`;
}

async function loadBackground(options = {}) {
  const { preserveCurrent = false } = options;
  const page = getCurrentPageModel();
  if (!page) {
    state.backgroundLoadToken += 1;
    state.renderedBackgroundPage = null;
    state.renderedBackgroundUrl = "";
    el.pageContainer.hidden = true;
    el.backgroundImage.hidden = true;
    el.backgroundImage.removeAttribute("src");
    el.backgroundImage.onload = null;
    el.backgroundImage.onerror = null;
    el.textLayer.innerHTML = "";
    return;
  }

  const displayWidth = page.width * state.zoom;
  const displayHeight = page.height * state.zoom;
  const deviceScale = Math.max(2, Math.min(window.devicePixelRatio || 1, 3));
  const targetWidth = Math.max(1, Math.round(displayWidth * deviceScale));
  const backgroundUrl = `${state.serviceBaseUrl}/documents/${state.document.id}/pages/${state.currentPage}/background?width=${targetWidth}`;
  const canPreserveCurrent = Boolean(
    preserveCurrent
    && !el.pageContainer.hidden
    && !el.backgroundImage.hidden
    && state.renderedBackgroundPage === state.currentPage
    && el.backgroundImage.getAttribute("src"),
  );

  if (canPreserveCurrent) {
    applyPageDisplaySize(page, displayWidth, displayHeight);
    renderTextLayer();
  } else {
    el.pageContainer.hidden = true;
    el.backgroundImage.hidden = true;
    el.textLayer.innerHTML = "";
  }

  if (state.renderedBackgroundPage === state.currentPage && state.renderedBackgroundUrl === backgroundUrl) {
    applyPageDisplaySize(page, displayWidth, displayHeight);
    el.backgroundImage.hidden = false;
    el.pageContainer.hidden = false;
    renderTextLayer();
    return;
  }

  const loadToken = state.backgroundLoadToken + 1;
  state.backgroundLoadToken = loadToken;

  await new Promise((resolve) => {
    const nextImage = new Image();
    nextImage.onload = () => {
      if (loadToken !== state.backgroundLoadToken) {
        resolve();
        return;
      }
      applyPageDisplaySize(page, displayWidth, displayHeight);
      el.backgroundImage.onload = null;
      el.backgroundImage.onerror = null;
      el.backgroundImage.src = backgroundUrl;
      el.backgroundImage.hidden = false;
      el.pageContainer.hidden = false;
      state.renderedBackgroundPage = state.currentPage;
      state.renderedBackgroundUrl = backgroundUrl;
      renderTextLayer();
      resolve();
    };
    nextImage.onerror = () => {
      if (loadToken !== state.backgroundLoadToken) {
        resolve();
        return;
      }
      if (!canPreserveCurrent) {
        el.backgroundImage.hidden = true;
        el.pageContainer.hidden = true;
        el.textLayer.innerHTML = "";
      }
      setStatus("PDF-Hintergrund konnte nicht geladen werden.");
      resolve();
    };
    nextImage.src = backgroundUrl;
  });
}

function appendCustomBlockControls(node, block) {
  if (!block.isCustom || isManualOverlayBlock(block)) {
    return;
  }

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "delete-handle";
  deleteButton.textContent = "X";
  deleteButton.title = "Feld löschen";
  deleteButton.addEventListener("mousedown", (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
  deleteButton.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    state.selectedBlockId = block.id;
    removeSelectedCustomBlock();
  });
  node.appendChild(deleteButton);

  for (const handle of ["nw", "ne", "sw", "se"]) {
    const resizeHandle = document.createElement("div");
    resizeHandle.className = `resize-handle resize-${handle}`;
    resizeHandle.contentEditable = "false";
    resizeHandle.tabIndex = -1;
    resizeHandle.title = "Größe ändern";
    resizeHandle.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      state.selectedBlockId = block.id;
      state.editingBlockId = null;
      rememberManualTextBlock(block);
      syncSelectedBlock();
      beginResize(block, node, event, handle);
    });
    node.appendChild(resizeHandle);
  }

  const rotateHandle = document.createElement("div");
  rotateHandle.className = "rotate-handle";
  rotateHandle.contentEditable = "false";
  rotateHandle.tabIndex = -1;
  rotateHandle.title = "Drehen";
  rotateHandle.addEventListener("mousedown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    state.selectedBlockId = block.id;
    state.editingBlockId = null;
    syncSelectedBlock();
    beginRotate(block, node, event);
  });
  node.appendChild(rotateHandle);
}

function appendManualTransformHandles(node, block) {
  for (const handle of ["nw", "n", "ne", "e", "se", "s", "sw", "w"]) {
    const resizeHandle = document.createElement("div");
    resizeHandle.className = `resize-handle resize-${handle}`;
    resizeHandle.contentEditable = "false";
    resizeHandle.tabIndex = -1;
    resizeHandle.title = "Größe ändern";
    resizeHandle.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      state.selectedBlockId = block.id;
      state.editingBlockId = null;
      rememberManualTextBlock(block);
      syncSelectedBlock();
      beginResize(block, node, event, handle);
    });
    node.appendChild(resizeHandle);
  }
}

function renderTextLayer() {
  el.textLayer.innerHTML = "";
  const page = getCurrentPageModel();
  if (!page) {
    return;
  }

  const scale = state.zoom;
  for (const line of page.lineOverlays || []) {
    const node = document.createElement("div");
    node.className = "line-overlay";
    const isHorizontal = Math.abs(line.y0 - line.y1) <= Math.abs(line.x0 - line.x1);
    const thickness = 1;
    if (isHorizontal) {
      node.style.left = `${Math.min(line.x0, line.x1) * scale}px`;
      node.style.top = `${(Math.min(line.y0, line.y1) * scale) - (thickness / 2)}px`;
      node.style.width = `${Math.abs(line.x1 - line.x0) * scale}px`;
      node.style.height = `${thickness}px`;
    } else {
      node.style.left = `${(Math.min(line.x0, line.x1) * scale) - (thickness / 2)}px`;
      node.style.top = `${Math.min(line.y0, line.y1) * scale}px`;
      node.style.width = `${thickness}px`;
      node.style.height = `${Math.abs(line.y1 - line.y0) * scale}px`;
    }
    node.style.background = line.color || "#000000";
    el.textLayer.appendChild(node);
  }

  const blocks = getCurrentPageBlocks();
  renderScanTemplateCovers(page, blocks, scale);

  for (const block of blocks) {
    if (block.editable === false && !block.isCheckbox) {
      continue;
    }
    if (isManualTextBlock(block)) {
      syncManualTextBaselineToBox(block);
    }

    const node = document.createElement("div");
    node.className = "text-block";
    if (block.isCheckbox) {
      node.classList.add("checkbox-block");
    }
    if (block.isCustom) {
      node.classList.add("custom-block");
    }
    if (isManualOverlayBlock(block)) {
      node.classList.add("manual-overlay-block");
    }
    if (isManualCoverBlock(block)) {
      node.classList.add("manual-cover-block");
    }
    if (isManualTextBlock(block) && state.editingBlockId === block.id) {
      node.classList.add("manual-editing");
    }
    if (isFixedGeneratedField(block)) {
      node.classList.add("generated-block");
    }
    if (isScanGeneratedTextField(block)) {
      node.classList.add("scan-generated-block");
    }
    if (isUnchangedScanGeneratedTextField(block)) {
      node.classList.add("scan-unchanged-block");
    }
    if (isContractIdNumberBlock(block)) {
      node.classList.add("contract-id-number-block");
    }
    if (block.id === state.selectedBlockId) {
      node.classList.add("selected");
    }
    node.dataset.blockId = block.id;
    if (block.groupKind) {
      node.dataset.groupKind = block.groupKind;
    }
    node.style.left = `${block.bbox.x0 * scale}px`;
    node.style.top = `${block.bbox.y0 * scale}px`;
    node.style.width = `${getBlockWidth(block, scale)}px`;
    node.style.height = `${getBlockHeight(block, scale)}px`;
    node.style.transform = `rotate(${Number(block.rotation) || 0}deg)`;
    node.style.setProperty("--masked-text-color", block.color || "#000");

    if (block.isCheckbox) {
      const renderMark = block.currentText.trim() && !isUnchangedScanGeneratedCheckbox(block);
      const manualCheckbox = isManualCheckboxBlock(block);
      if (manualCheckbox && renderMark) {
        node.classList.add("manual-checkbox-checked");
      }
      node.title = manualCheckbox ? "Kreuzfeld" : "Ankreuzen";
      node.tabIndex = manualCheckbox ? 0 : node.tabIndex;
      node.setAttribute("role", "checkbox");
      node.setAttribute("aria-checked", block.currentText.trim() ? "true" : "false");
      const mark = document.createElement("div");
      mark.className = "checkbox-mark";
      mark.textContent = renderMark ? "x" : "";
      mark.style.fontSize = `${Math.max(8, Math.min(getBlockWidth(block, scale), getBlockHeight(block, scale)) * 0.95)}px`;
      node.appendChild(mark);
      node.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (isManualOverlayBlock(block)) {
          state.selectedBlockId = block.id;
          state.editingBlockId = null;
          syncSelectedBlock();
          node.focus({ preventScroll: true });
          beginPotentialDrag(block, node, event);
        }
      });
      node.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (state.suppressClickBlockId === block.id) {
          state.suppressClickBlockId = null;
          return;
        }
        const eventTarget = event.target instanceof HTMLElement ? event.target : null;
        if (eventTarget?.closest(".resize-handle, .rotate-handle")) {
          return;
        }
        if (manualCheckbox) {
          state.selectedBlockId = block.id;
          state.editingBlockId = null;
          syncSelectedBlock();
          node.focus({ preventScroll: true });
          return;
        }
        toggleCheckbox(block);
      });
      el.textLayer.appendChild(node);
      continue;
    }

    if (isManualCoverBlock(block)) {
      node.title = "Weiße Fläche";
      node.tabIndex = 0;
      node.addEventListener("mousedown", (event) => {
        if (event.button !== 0) {
          return;
        }
        state.selectedBlockId = block.id;
        state.editingBlockId = null;
        syncSelectedBlock();
        node.focus({ preventScroll: true });
        beginPotentialDrag(block, node, event);
      });
      node.addEventListener("click", (event) => {
        if (state.suppressClickBlockId === block.id) {
          state.suppressClickBlockId = null;
          event.preventDefault();
          return;
        }
        const eventTarget = event.target instanceof HTMLElement ? event.target : null;
        if (eventTarget?.closest(".resize-handle, .rotate-handle")) {
          return;
        }
        state.selectedBlockId = block.id;
        state.editingBlockId = null;
        syncSelectedBlock();
        node.focus({ preventScroll: true });
      });
      appendManualTransformHandles(node, block);
      el.textLayer.appendChild(node);
      continue;
    }

    if (isUnderlineTemplateBlock(block)) {
      normalizeMaskedBlockText(block);
      const template = getMaskedTemplate(block);
      const slotRegion = getMaskedSlotRegion(template, block, scale);
      const isGermanIban = isGermanIbanTemplateBlock(block);
      const totalMaskWidth = slotRegion.totalWidth || measureTextWidth(template, block, scale);
      const requiredWidth = Math.max(
        getBlockWidth(block, scale),
        Math.ceil(totalMaskWidth) + 2,
      );
      const requiredHeight = Math.max(
        getBlockHeight(block, scale),
        Math.ceil(block.lineHeight * scale) + 2,
      );
      const maskRuns = !isGermanIban
        ? getMaskedRuns(template, block, scale, totalMaskWidth || requiredWidth)
        : [];
      const ibanSlots = isGermanIban
        ? getIbanVisualSlots(template, block, scale, totalMaskWidth || requiredWidth)
        : [];
      node.classList.add("masked-template-block");
      node.style.width = `${requiredWidth}px`;
      node.style.height = `${requiredHeight}px`;

      const maskedView = document.createElement("div");
      maskedView.className = "masked-template-view";
      maskedView.style.fontFamily = block.cssFontFamily;
      maskedView.style.fontWeight = block.fontWeight || "400";
      maskedView.style.fontStyle = block.fontStyle || "normal";
      maskedView.style.fontSize = `${block.fontSize * scale}px`;
      maskedView.style.lineHeight = `${block.lineHeight * scale}px`;
      maskedView.style.color = block.color;
      maskedView.textContent = isGermanIban ? "" : template;

      const runStates = [];
      if (isGermanIban) {
        for (const slot of ibanSlots) {
          const slotRegionNode = document.createElement("div");
          slotRegionNode.className = "masked-template-region";
          slotRegionNode.style.left = `${slot.left}px`;
          slotRegionNode.style.width = `${Math.max(2, slot.width)}px`;

          const slotFill = document.createElement("div");
          slotFill.className = "masked-template-fill";
          slotFill.style.fontFamily = block.cssFontFamily;
          slotFill.style.fontWeight = block.fontWeight || "400";
          slotFill.style.fontStyle = block.fontStyle || "normal";
          slotFill.style.fontSize = `${block.fontSize * scale}px`;
          slotFill.style.lineHeight = `${block.lineHeight * scale}px`;
          slotFill.style.color = block.color;
          slotFill.style.textAlign = "center";

          const slotNode = document.createElement("span");
          slotNode.className = "masked-template-slot-fill";
          slotNode.style.left = "0px";
          slotNode.style.width = `${Math.max(2, slot.width)}px`;
          slotFill.appendChild(slotNode);

          const slotCapture = document.createElement("input");
          slotCapture.type = "text";
          slotCapture.className = "masked-template-capture";
          slotCapture.autocomplete = "off";
          slotCapture.spellcheck = false;
          slotCapture.maxLength = 1;
          slotCapture.inputMode = "numeric";
          slotCapture.setAttribute("aria-label", block.originalText);
          slotCapture.style.fontFamily = block.cssFontFamily;
          slotCapture.style.fontWeight = block.fontWeight || "400";
          slotCapture.style.fontStyle = block.fontStyle || "normal";
          slotCapture.style.fontSize = `${block.fontSize * scale}px`;
          slotCapture.style.lineHeight = `${block.lineHeight * scale}px`;

          slotRegionNode.appendChild(slotFill);
          slotRegionNode.appendChild(slotCapture);
          node.appendChild(slotRegionNode);
          runStates.push({
            startIndex: slot.startIndex,
            endIndex: slot.endIndex,
            template: slot.template,
            region: slotRegionNode,
            fill: slotFill,
            capture: slotCapture,
            slotNodes: [slotNode],
          });
        }
      } else {
        for (const run of maskRuns) {
          const runRegion = document.createElement("div");
          runRegion.className = "masked-template-region";
          runRegion.style.left = `${run.left}px`;
          runRegion.style.width = `${run.width}px`;

          const runFill = document.createElement("div");
          runFill.className = "masked-template-fill";
          runFill.style.fontFamily = block.cssFontFamily;
          runFill.style.fontWeight = block.fontWeight || "400";
          runFill.style.fontStyle = block.fontStyle || "normal";
          runFill.style.fontSize = `${block.fontSize * scale}px`;
          runFill.style.lineHeight = `${block.lineHeight * scale}px`;
          runFill.style.color = block.color;

          const runCapture = document.createElement("input");
          runCapture.type = "text";
          runCapture.className = "masked-template-capture";
          runCapture.autocomplete = "off";
          runCapture.spellcheck = false;
          runCapture.maxLength = (run.template.match(/_/g) || []).length;
          runCapture.setAttribute("aria-label", block.originalText);
          runCapture.style.fontFamily = block.cssFontFamily;
          runCapture.style.fontWeight = block.fontWeight || "400";
          runCapture.style.fontStyle = block.fontStyle || "normal";
          runCapture.style.fontSize = `${block.fontSize * scale}px`;
          runCapture.style.lineHeight = `${block.lineHeight * scale}px`;

          runRegion.appendChild(runFill);
          runRegion.appendChild(runCapture);
          node.appendChild(runRegion);
          runStates.push({
            ...run,
            region: runRegion,
            fill: runFill,
            capture: runCapture,
          });
        }
      }

      const syncActiveMaskedRegion = () => {
        const activeElement = document.activeElement;
        let activeRun = null;
        for (const run of runStates) {
          const isActive = run.capture === activeElement;
          run.region.classList.toggle("active", isActive);
          if (isActive) {
            activeRun = run;
          }
        }
        node.classList.toggle("editing", !isGermanIban && Boolean(activeRun));
      };

      const focusIbanRun = (runIndex) => {
        if (!isGermanIban || runIndex < 0 || runIndex >= runStates.length) {
          return;
        }
        const targetRun = runStates[runIndex];
        if (!targetRun?.capture) {
          return;
        }
        targetRun.capture.focus();
        targetRun.capture.setSelectionRange(0, targetRun.capture.value.length);
      };

      const rebuildMaskedField = () => {
        const currentText = isGermanIban
          ? normalizeIbanCurrentText(template, block.currentText)
          : (block.currentText && block.currentText.length === template.length
            ? block.currentText
            : applyMaskedValue(template, extractMaskedValue(template, block.currentText)));
        block.currentText = currentText;
        if (isGermanIban) {
          for (const slot of runStates) {
            const slotText = currentText.slice(slot.startIndex, slot.endIndex + 1);
            const slotChar = Array.from(slotText).find((char) => /\d/.test(char)) ?? "";
            slot.slotNodes.forEach((node) => {
              node.textContent = slotChar;
            });
            slot.capture.value = slotChar;
          }
        } else {
          for (const run of runStates) {
            const runText = currentText.slice(run.startIndex, run.endIndex + 1);
            run.fill.textContent = getMaskedOverlayText(run.template, runText);
            run.capture.value = extractMaskedValue(run.template, runText);
          }
        }
      };

      const commitMaskedValue = (rawValue, run = null) => {
        if (isGermanIban) {
          const sanitizedValue = sanitizeIbanSlotValue(rawValue);
          const normalizedCurrentText = normalizeIbanCurrentText(template, block.currentText);
          const slotText = sanitizedValue || "_";
          block.currentText = `${normalizedCurrentText.slice(0, run.startIndex)}${slotText}${"_".repeat(Math.max(0, run.endIndex - run.startIndex))}${normalizedCurrentText.slice(run.endIndex + 1)}`;
          rebuildMaskedField();
          scheduleDraftSave();
          return;
        }

        const normalizedCurrentText = block.currentText && block.currentText.length === template.length
          ? block.currentText
          : applyMaskedValue(template, extractMaskedValue(template, block.currentText));
        const runText = applyMaskedValue(run.template, rawValue);
        block.currentText = `${normalizedCurrentText.slice(0, run.startIndex)}${runText}${normalizedCurrentText.slice(run.endIndex + 1)}`;
        rebuildMaskedField();
        scheduleDraftSave();
      };

      rebuildMaskedField();
      for (const run of runStates) {
        run.capture.addEventListener("focus", () => {
          state.selectedBlockId = block.id;
          syncSelectedBlock();
          syncActiveMaskedRegion();
          if (isGermanIban) {
            requestAnimationFrame(() => {
              run.capture.setSelectionRange(0, run.capture.value.length);
            });
          }
        });
        run.capture.addEventListener("mousedown", () => {
          state.selectedBlockId = block.id;
          syncSelectedBlock();
        });
        run.capture.addEventListener("keydown", (event) => {
          if (!isGermanIban) {
            return;
          }
          if (event.key.length === 1 && !/[0-9]/.test(event.key)) {
            event.preventDefault();
            return;
          }
          if (event.key !== "Backspace" && event.key !== "Delete") {
            return;
          }
          event.preventDefault();
          const currentIndex = runStates.indexOf(run);
          const hasCurrentValue = Boolean(run.capture.value);
          const targetIndex = hasCurrentValue
            ? currentIndex
            : currentIndex - 1;
          if (targetIndex < 0 || targetIndex >= runStates.length) {
            requestAnimationFrame(() => {
              run.capture.focus();
              run.capture.setSelectionRange(0, 0);
            });
            return;
          }

          commitMaskedValue("", runStates[targetIndex]);
          requestAnimationFrame(() => {
            focusIbanRun(targetIndex);
            runStates[targetIndex].capture.setSelectionRange(0, 0);
          });
        });
        run.capture.addEventListener("input", () => {
          if (isGermanIban) {
            run.capture.value = sanitizeIbanSlotValue(run.capture.value);
          }
          commitMaskedValue(run.capture.value, run);
          if (isGermanIban && run.capture.value) {
            const currentIndex = runStates.indexOf(run);
            if (currentIndex >= 0 && currentIndex < runStates.length - 1) {
              requestAnimationFrame(() => {
                focusIbanRun(currentIndex + 1);
              });
            }
          }
        });
        run.capture.addEventListener("blur", () => {
          requestAnimationFrame(() => {
            normalizeMaskedBlockText(block);
            rebuildMaskedField();
            syncActiveMaskedRegion();
          });
        });
      }

      node.addEventListener("mousedown", () => {
        state.selectedBlockId = block.id;
        syncSelectedBlock();
      });
      node.addEventListener("click", (event) => {
        const eventTarget = event.target instanceof HTMLElement ? event.target : null;
        if (eventTarget?.classList.contains("masked-template-capture")) {
          return;
        }

        let targetCapture = null;
        const clickedRegion = eventTarget?.closest(".masked-template-region");
        if (clickedRegion) {
          targetCapture = clickedRegion.querySelector(".masked-template-capture");
        }
        if (!targetCapture) {
          targetCapture = runStates[0]?.capture;
        }
        if (!targetCapture) {
          return;
        }
        targetCapture.focus();
        targetCapture.setSelectionRange(targetCapture.value.length, targetCapture.value.length);
      });

      node.appendChild(maskedView);
      if (isGermanIban) {
        // Group regions were already appended above.
      }
      el.textLayer.appendChild(node);
      continue;
    }

    const editor = document.createElement("div");
    editor.className = "text-editor";
    editor.spellcheck = false;
    editor.contentEditable = isManualTextBlock(block) && state.editingBlockId !== block.id ? "false" : "plaintext-only";
    if (block.currentText) {
      editor.textContent = block.currentText;
    } else {
      ensureEditablePlaceholder(editor);
    }
    editor.style.fontFamily = block.cssFontFamily;
    editor.style.fontWeight = isContractIdNumberBlock(block) ? "700" : getManualDisplayFontWeight(block);
    editor.style.fontStyle = block.fontStyle || "normal";
    editor.style.fontSize = `${block.fontSize * scale}px`;
    editor.style.lineHeight = `${block.lineHeight * scale}px`;
    editor.style.color = block.color;
    editor.style.textAlign = block.align;
    editor.style.textDecoration = String(block.textDecoration || "").toLowerCase() === "underline" ? "underline" : "none";
    editor.style.textDecorationThickness = getTextDecorationThickness(block, scale);
    editor.style.textUnderlineOffset = String(block.textDecoration || "").toLowerCase() === "underline" ? "0.08em" : "";
    let contractIdUnderline = null;
    if (isContractIdNumberBlock(block)) {
      editor.style.textDecoration = "none";
      contractIdUnderline = document.createElement("div");
      contractIdUnderline.className = "contract-id-underline";
      syncContractIdUnderline(contractIdUnderline, block, scale);
    }
    syncManualTextEditorBox(editor, block, scale);

    editor.addEventListener("focus", () => {
      state.selectedBlockId = block.id;
      rememberManualTextBlock(block);
      if (isManualTextBlock(block)) {
        state.editingBlockId = block.id;
        editor.contentEditable = "plaintext-only";
      }
      if (!block.currentText) {
        ensureEditablePlaceholder(editor);
        requestAnimationFrame(() => {
          placeCaretAtEnd(editor);
        });
      }
      syncSelectedBlock();
      if (isScanGeneratedTextField(block)) {
        requestAnimationFrame(() => {
          selectEditableText(editor);
        });
      }
    });

    editor.addEventListener("mousedown", (event) => {
      state.selectedBlockId = block.id;
      rememberManualTextBlock(block);
      syncSelectedBlock();
      if (isManualTextBlock(block) && state.editingBlockId !== block.id) {
        event.preventDefault();
        beginPotentialDrag(block, node, event);
      }
    });

    editor.addEventListener("keydown", (event) => {
      if (handleDeleteForEmptyManualTextBlock(event)) {
        return;
      }

      if (!block.isCustom || event.key !== "Enter" || event.shiftKey) {
        return;
      }

      requestAnimationFrame(() => {
        growCustomBlockByLine(node, block, scale);
        autoSizeTextBlock(node, editor, block, scale);
        scheduleDraftSave();
      });
    });

    editor.addEventListener("input", () => {
      block.currentText = readEditableText(editor);
      block.currentValue = block.currentText;
      autoSizeTextBlock(node, editor, block, scale);
      if (contractIdUnderline) {
        syncContractIdUnderline(contractIdUnderline, block, scale);
      }
      scheduleDraftSave();
    });

    editor.addEventListener("blur", () => {
      block.currentText = readEditableText(editor);
      block.currentValue = block.currentText;
      commitDocumentHistory();
      if (isManualTextBlock(block) && state.editingBlockId === block.id) {
        state.editingBlockId = null;
        editor.contentEditable = "false";
        syncSelectedBlock();
      }
      if (!block.currentText) {
        ensureEditablePlaceholder(editor);
      }
      if (isScanGeneratedTextField(block) && isScanFieldChanged(block)) {
        renderTextLayer();
      }
    });

    if (isManualTextBlock(block)) {
      node.addEventListener("click", (event) => {
        if (state.suppressClickBlockId === block.id) {
          state.suppressClickBlockId = null;
          event.preventDefault();
          return;
        }
        const eventTarget = event.target instanceof HTMLElement ? event.target : null;
        if (eventTarget?.closest(".delete-handle, .resize-handle, .rotate-handle")) {
          return;
        }
        const wasEditing = state.editingBlockId === block.id;
        state.selectedBlockId = block.id;
        state.editingBlockId = block.id;
        rememberManualTextBlock(block);
        editor.contentEditable = "plaintext-only";
        syncSelectedBlock();
        if (!wasEditing) {
          requestAnimationFrame(() => {
            focusEditableTarget(editor);
          });
        }
      });
    }

    node.appendChild(editor);
    if (contractIdUnderline) {
      node.appendChild(contractIdUnderline);
    }
    el.textLayer.appendChild(node);
    if (isManualTextBlock(block)) {
      appendManualTransformHandles(node, block);
    }
    if (block.isCustom && !isManualOverlayBlock(block)) {
      const deleteButton = document.createElement("button");
      deleteButton.type = "button";
      deleteButton.className = "delete-handle";
      deleteButton.textContent = "X";
      deleteButton.title = "Feld löschen";
      deleteButton.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
      });
      deleteButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        state.selectedBlockId = block.id;
        removeSelectedCustomBlock();
      });
      node.appendChild(deleteButton);

      const dragHandle = document.createElement("div");
      dragHandle.className = "drag-handle";
      dragHandle.contentEditable = "false";
      dragHandle.tabIndex = -1;
      dragHandle.title = "Verschieben";
      dragHandle.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
        state.selectedBlockId = block.id;
        syncSelectedBlock();
        beginPotentialDrag(block, node, event);
      });
      node.appendChild(dragHandle);

      const resizeHandle = document.createElement("div");
      resizeHandle.className = "resize-handle";
      resizeHandle.contentEditable = "false";
      resizeHandle.tabIndex = -1;
      resizeHandle.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
        state.selectedBlockId = block.id;
        syncSelectedBlock();
        beginResize(block, node, event);
      });
      node.appendChild(resizeHandle);
    }
    autoSizeTextBlock(node, editor, block, scale);
  }

  syncSelectedBlock();
}

async function saveDraft() {
  if (!state.document?.supportStatus?.supported) {
    return;
  }

  syncDocumentBlockValues();
  const payload = {
    fields: state.document.fields,
  };

  const response = await fetch(`${state.serviceBaseUrl}/documents/${state.document.id}/draft`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(error || "Entwurf konnte nicht gespeichert werden.");
  }
}

async function saveLearnedTemplate() {
  if (!state.document?.supportStatus?.supported) {
    return;
  }
  if (!hasManualTemplateFields()) {
    throw new Error("Bitte zuerst die benötigten Felder setzen.");
  }

  const templateName = getTemplateNameValue();
  setStatus("Vorlage wird gespeichert...");

  const response = await fetch(`${state.serviceBaseUrl}/documents/${state.document.id}/learn-template`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: templateName,
      fields: state.document.fields,
    }),
  });

  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(payload?.detail || "Vorlage konnte nicht gespeichert werden.");
  }

  state.document.detectedTemplateId = payload.templateId;
  state.document.detectedTemplateFamily = "user_learned";
  updateButtons();
  updateMeta();
  setStatus(
    payload.replacedExisting
      ? `Vorlage aktualisiert: ${payload.templateName}`
      : `Vorlage gespeichert: ${payload.templateName}`,
  );
}

function scheduleDraftSave() {
  clearTimeout(state.autoSaveTimer);
  state.autoSaveTimer = setTimeout(async () => {
    try {
      await saveDraft();
    } catch (error) {
      console.error(error);
      setStatus("Bearbeitungsstand konnte nicht aktualisiert werden.");
    }
  }, 300);
}

async function importPdf(sourcePath) {
  clearTimeout(state.autoSaveTimer);
  state.autoSaveTimer = null;
  setStatus("PDF wird importiert...");
  const response = await fetch(`${state.serviceBaseUrl}/documents/import`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sourcePath }),
  });

  if (!response.ok) {
    throw new Error(await readErrorMessage(response, "Import fehlgeschlagen."));
  }

  await applyImportedDocument(await response.json());
}

async function importPdfFile(file) {
  clearTimeout(state.autoSaveTimer);
  state.autoSaveTimer = null;
  setStatus("PDF wird importiert...");

  const response = await fetch(`${state.serviceBaseUrl}/documents/upload`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileName: file.name,
      fileDataBase64: arrayBufferToBase64(await file.arrayBuffer()),
    }),
  });

  if (!response.ok) {
    throw new Error(await readErrorMessage(response, "Import fehlgeschlagen."));
  }

  await applyImportedDocument(await response.json());
}

async function applyImportedDocument(documentModel) {
  state.document = ensureDocumentFieldCompatibility(documentModel);
  state.currentPage = 1;
  state.selectedBlockId = null;
  state.editingBlockId = null;
  state.activePdfTool = "select";
  state.lastExportPath = null;
  state.backgroundLoadToken += 1;
  state.renderedBackgroundPage = null;
  state.renderedBackgroundUrl = "";
  if (el.templateNameInput) {
    el.templateNameInput.value = getDefaultTemplateName();
  }
  injectFontFaces();
  showSupportStatus();
  resetDocumentHistory();

  if (!state.document.supportStatus.supported) {
    el.backgroundImage.removeAttribute("src");
    el.textLayer.innerHTML = "";
    el.pageContainer.hidden = true;
    updateButtons();
    updateMeta();
    setStatus("PDF wird nur eingeschränkt oder gar nicht unterstützt.");
    return;
  }

  const page = getCurrentPageModel();
  const availableWidth = Math.max(320, el.pageArea.clientWidth - 24);
  state.fitZoom = availableWidth / page.width;
  state.zoom = state.fitZoom;
  clearWhiteboard(false);
  updateButtons();
  updateMeta();
  await loadBackground();
  setStatus("PDF geladen. Werkzeug wählen und weiße Felder direkt auf die Seite setzen.");
}

async function exportPdf() {
  setStatus("PDF wird exportiert...");
  const response = await fetch(`${state.serviceBaseUrl}/documents/${state.document.id}/export`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });

  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(payload?.detail || "Export fehlgeschlagen.");
  }

  state.lastExportPath = payload.outputPath;
  updateButtons();
  setStatus(`Exportiert: ${payload.outputPath}`);
}

async function exportPdfDownload() {
  setStatus("PDF wird exportiert...");
  const response = await fetch(`${state.serviceBaseUrl}/documents/${state.document.id}/export-download`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });

  if (!response.ok) {
    throw new Error(await readErrorMessage(response, "Export fehlgeschlagen."));
  }

  const filename = getDownloadFilename(
    response.headers.get("content-disposition"),
    `${(state.document?.sourcePath || "dokument").replace(/\.pdf$/i, "")}-bearbeitet.pdf`,
  );
  const blob = await response.blob();
  state.lastExportPath = null;
  updateButtons();
  triggerBlobDownload(blob, filename);
  setStatus(`Exportiert: ${filename}`);
}

async function initialize() {
  state.serviceBaseUrl = isDesktopRuntime()
    ? await window.desktopAPI.getServiceBaseUrl()
    : window.location.origin;
  resizeWhiteboardCanvas(false);
  clearWhiteboard(false);
  setWhiteboardMode("pen");
  updateButtons();
  updateMeta();
  checkForAppUpdates().catch((error) => {
    console.warn("Update check failed:", error);
  });
}

el.openButton.addEventListener("click", async () => {
  try {
    if (isDesktopRuntime()) {
      const sourcePath = await window.desktopAPI.openPdfDialog();
      if (!sourcePath) {
        return;
      }
      await importPdf(sourcePath);
      return;
    }

    el.webPdfInput.value = "";
    el.webPdfInput.click();
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  }
});

el.templateLearningButton.addEventListener("click", () => {
  const nextState = !state.templateLearning.active;
  setTemplateLearningActive(nextState);
  setStatus(
    nextState
      ? "Lernmodus aktiv. Felder setzen und danach die Vorlage speichern."
      : "Lernmodus beendet.",
  );
});

el.saveTemplateButton.addEventListener("click", async () => {
  try {
    await saveLearnedTemplate();
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  }
});

el.webPdfInput.addEventListener("change", async () => {
  const [file] = el.webPdfInput.files || [];
  if (!file) {
    return;
  }

  try {
    await importPdfFile(file);
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  } finally {
    el.webPdfInput.value = "";
  }
});

el.undoButton.addEventListener("click", async () => {
  await moveDocumentHistory(-1);
});

el.redoButton.addEventListener("click", async () => {
  await moveDocumentHistory(1);
});

el.prevButton.addEventListener("click", async () => {
  state.currentPage -= 1;
  state.selectedBlockId = null;
  await loadBackground();
  updateButtons();
});

el.nextButton.addEventListener("click", async () => {
  state.currentPage += 1;
  state.selectedBlockId = null;
  await loadBackground();
  updateButtons();
});

el.zoomOutButton.addEventListener("click", async () => {
  state.zoom = clamp(state.zoom / 1.15, 0.25, 4);
  updateButtons();
  await loadBackground({ preserveCurrent: true });
  updateButtons();
});

el.zoomFitButton.addEventListener("click", async () => {
  state.zoom = state.fitZoom;
  updateButtons();
  await loadBackground({ preserveCurrent: true });
  updateButtons();
});

el.zoomInButton.addEventListener("click", async () => {
  state.zoom = clamp(state.zoom * 1.15, 0.25, 4);
  updateButtons();
  await loadBackground({ preserveCurrent: true });
  updateButtons();
});

el.saveDraftButton.addEventListener("click", async () => {
  try {
    await saveDraft();
    setStatus("Bearbeitungsstand nur für diese Sitzung aktualisiert.");
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  }
});

el.selectToolButton?.addEventListener("click", () => {
  setActivePdfTool("select");
  setStatus("Auswahl aktiv.");
});

el.addTextButton?.addEventListener("click", () => {
  setActivePdfTool("text");
  setStatus("Text hinzufügen aktiv. Auf die PDF klicken, um ein weißes Textfeld zu setzen.");
});

el.fontFamilySelect?.addEventListener("change", () => {
  applyManualTextStylePatch({ fontFamily: el.fontFamilySelect.value });
});

el.fontSizeInput?.addEventListener("change", () => {
  applyManualTextStylePatch({ fontSize: el.fontSizeInput.value });
});

el.boldButton?.addEventListener("click", () => {
  const style = getActiveManualTextStyleSource();
  applyManualTextStylePatch({ fontWeight: isBoldFontWeight(style.fontWeight) ? "400" : "700" });
});

el.italicButton?.addEventListener("click", () => {
  const style = getActiveManualTextStyleSource();
  applyManualTextStylePatch({
    fontStyle: String(style.fontStyle || "").toLowerCase() === "italic" ? "normal" : "italic",
  });
});

el.underlineButton?.addEventListener("click", () => {
  const style = getActiveManualTextStyleSource();
  applyManualTextStylePatch({
    textDecoration: String(style.textDecoration || "").toLowerCase() === "underline" ? "none" : "underline",
  });
});

el.eraseTextButton?.addEventListener("click", () => {
  setActivePdfTool("erase");
  setStatus("Text löschen aktiv. Auf die PDF klicken, um eine weiße Fläche zu setzen.");
});

el.checkButton?.addEventListener("click", () => {
  setActivePdfTool("check");
  setStatus("Ankreuzen aktiv. Auf die Box klicken, um ein weißes Feld mit Kreuz zu setzen.");
});

el.uncheckButton?.addEventListener("click", () => {
  setActivePdfTool("uncheck");
  setStatus("Kreuz löschen aktiv. Auf die Box klicken, um sie weiß zu überdecken.");
});

el.exportButton.addEventListener("click", async () => {
  try {
    if (hasEditableDocument()) {
      await saveDraft();
      if (isDesktopRuntime()) {
        await exportPdf();
      } else {
        await exportPdfDownload();
      }
    } else {
      await exportWhiteboardPdf();
    }
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  }
});

el.revealButton.addEventListener("click", async () => {
  if (!isDesktopRuntime() || !state.lastExportPath) {
    return;
  }
  await window.desktopAPI.revealFile(state.lastExportPath);
});

if (el.updateInstallButton) {
  el.updateInstallButton.addEventListener("click", async () => {
    await installAvailableUpdate();
  });
}

if (el.updateDismissButton) {
  el.updateDismissButton.addEventListener("click", () => {
    state.updateInfo = null;
    updateUpdateBanner();
  });
}

el.whiteboardPenButton.addEventListener("click", () => {
  setWhiteboardMode("pen");
});

el.whiteboardEraserButton.addEventListener("click", () => {
  setWhiteboardMode("eraser");
});

el.whiteboardUndoButton.addEventListener("click", async () => {
  await moveWhiteboardHistory(-1);
});

el.whiteboardRedoButton.addEventListener("click", async () => {
  await moveWhiteboardHistory(1);
});

el.whiteboardClearButton.addEventListener("click", () => {
  clearWhiteboard(true);
});

el.whiteboardColorInput.addEventListener("input", () => {
  state.whiteboard.color = el.whiteboardColorInput.value || "#000000";
  updateWhiteboardCursorAppearance();
  updateWhiteboardToolButtons();
});

el.whiteboardThicknessInput.addEventListener("input", () => {
  state.whiteboard.penSize = clamp(
    Math.round(Number(el.whiteboardThicknessInput.value) || 3),
    1,
    24,
  );
  el.whiteboardThicknessValue.textContent = `${getWhiteboardPenSize()} px`;
  updateWhiteboardCursorAppearance();
  updateWhiteboardToolButtons();
});

el.whiteboardCanvas.addEventListener("pointerenter", (event) => {
  if (!isWhiteboardMode()) {
    hideWhiteboardCursor();
    return;
  }
  syncWhiteboardCursorFromEvent(event);
});

el.whiteboardCanvas.addEventListener("pointerdown", (event) => {
  if (!isWhiteboardMode() || state.whiteboard.isRestoringHistory || event.button !== 0) {
    return;
  }
  resizeWhiteboardCanvas(true);
  event.preventDefault();
  const point = syncWhiteboardCursorFromEvent(event);
  state.whiteboard.isDrawing = true;
  state.whiteboard.lastPoint = point;
  state.whiteboard.pointerId = event.pointerId;
  state.whiteboard.hasUncommittedChange = true;
  el.whiteboardCanvas.setPointerCapture(event.pointerId);
  drawWhiteboardSegment(point, point);
});

el.whiteboardCanvas.addEventListener("pointermove", (event) => {
  if (!isWhiteboardMode() || state.whiteboard.isRestoringHistory) {
    hideWhiteboardCursor();
    return;
  }
  const point = syncWhiteboardCursorFromEvent(event);
  if (!state.whiteboard.isDrawing || !state.whiteboard.lastPoint) {
    return;
  }
  event.preventDefault();
  drawWhiteboardSegment(state.whiteboard.lastPoint, point);
  state.whiteboard.lastPoint = point;
  state.whiteboard.hasUncommittedChange = true;
});

el.whiteboardCanvas.addEventListener("pointerup", (event) => {
  syncWhiteboardCursorFromEvent(event);
  finishWhiteboardStroke(event.pointerId);
});

el.whiteboardCanvas.addEventListener("pointercancel", (event) => {
  finishWhiteboardStroke(event.pointerId);
  hideWhiteboardCursor();
});

el.whiteboardCanvas.addEventListener("pointerleave", () => {
  if (!state.whiteboard.isDrawing) {
    hideWhiteboardCursor();
  }
});

window.addEventListener("blur", () => {
  finishWhiteboardStroke();
  hideWhiteboardCursor();
});

el.pageContainer.addEventListener("contextmenu", (event) => {
  if (state.activePdfTool && state.activePdfTool !== "select") {
    event.preventDefault();
  }
});

el.pageContainer.addEventListener("click", (event) => {
  if (!state.document?.supportStatus?.supported || el.pageContainer.hidden) {
    return;
  }
  if (event.button !== 0) {
    return;
  }

  const target = event.target;
  if (!(target instanceof Element) || target.closest(".text-block")) {
    return;
  }

  if (state.activePdfTool && state.activePdfTool !== "select") {
    event.preventDefault();
    createManualOverlayBlock(event, state.activePdfTool);
    return;
  }

  if (!isScanTemplateDocument()) {
    return;
  }

  const containerRect = el.pageContainer.getBoundingClientRect();
  const x = (event.clientX - containerRect.left) / state.zoom;
  const y = (event.clientY - containerRect.top) / state.zoom;
  const block = findTemplateBlockNearPoint(x, y);
  if (!block) {
    return;
  }

  state.selectedBlockId = block.id;
  state.pendingFocusBlockId = block.id;
  syncSelectedBlock();
});

window.addEventListener("mousedown", (event) => {
  if (!state.document?.supportStatus?.supported) {
    return;
  }

  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }

  if (target.closest(".text-block")) {
    return;
  }

  if (target.closest("#toolbar")) {
    return;
  }

  if (state.drag || state.resize) {
    return;
  }

  clearSelection();
});

window.addEventListener("mousemove", handleResizeMove);
window.addEventListener("mousemove", handleDragMove);
window.addEventListener("mousemove", handleRotateMove);
window.addEventListener("mouseup", handleDragEnd);

document.addEventListener("keydown", handleDeleteForSelectedManualOverlayBlock, true);
document.addEventListener("keydown", handleDeleteForEmptyManualTextBlock, true);

window.addEventListener("keydown", (event) => {
  if (handleDeleteForSelectedManualOverlayBlock(event)) {
    return;
  }

  if (handleDeleteForEmptyManualTextBlock(event)) {
    return;
  }

  if (event.key === "Escape") {
    const active = document.activeElement;
    if (
      active instanceof HTMLElement
      && (active.classList.contains("text-editor") || active.classList.contains("masked-template-capture"))
    ) {
      active.blur();
      event.preventDefault();
    }
    return;
  }

  if (event.key !== "Delete") {
    return;
  }

  const active = document.activeElement;
  const selected = getSelectedBlock();
  if (!selected?.isCustom) {
    return;
  }

  if (
    active instanceof HTMLElement
    && (active.classList.contains("text-editor") || active.classList.contains("masked-template-capture"))
  ) {
    const owner = active.closest(".text-block");
    if (owner?.dataset.blockId === selected.id && selected.currentText.trim()) {
      return;
    }
  }

  event.preventDefault();
  removeSelectedCustomBlock();
});

window.addEventListener("resize", async () => {
  if (hasEditableDocument()) {
    const page = getCurrentPageModel();
    if (!page) {
      return;
    }
    const availableWidth = Math.max(320, el.pageArea.clientWidth - 24);
    state.fitZoom = availableWidth / page.width;
  } else {
    resizeWhiteboardCanvas(true);
  }
});

initialize().catch((error) => {
  console.error(error);
  setStatus("Initialisierung fehlgeschlagen.");
});
