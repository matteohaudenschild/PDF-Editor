const state = {
  serviceBaseUrl: null,
  document: null,
  currentPage: 1,
  zoom: 1,
  fitZoom: 1,
  selectedBlockId: null,
  pendingFocusBlockId: null,
  drag: null,
  resize: null,
  lastExportPath: null,
  autoSaveTimer: null,
  updateInfo: null,
  updateInstalling: false,
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
const DRAG_THRESHOLD_PX = 4;
const WHITEBOARD_HISTORY_LIMIT = 30;

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
    `Update ${info.latestVersion} ist verfuegbar${sizeText}. Deine installierte Version ist ${info.currentVersion}.`;
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
    setStatus(error.message || "Update konnte nicht installiert werden.");
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

function createBlockId(prefix) {
  if (window.crypto?.randomUUID) {
    return `${prefix}-${window.crypto.randomUUID()}`;
  }
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getBlockWidth(block, scale) {
  return Math.max(4, (block.bbox.x1 - block.bbox.x0) * scale);
}

function getBlockHeight(block, scale) {
  return Math.max(4, (block.bbox.y1 - block.bbox.y0) * scale);
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
    && isScanFieldChanged(block)
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
  renderTextLayer();
  scheduleDraftSave();
}

function clampToPage(page, value, size, padding = 2) {
  return Math.min(Math.max(value, 0), Math.max(0, (page - size) - padding));
}

function createCustomTextBlock(event) {
  if (!state.document?.supportStatus?.supported) {
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

  const width = CUSTOM_BLOCK_DEFAULT_WIDTH;
  const height = CUSTOM_BLOCK_DEFAULT_HEIGHT;
  const x0 = clampToPage(page.width, x, width);
  const y0 = clampToPage(page.height, y, height);
  const baselineOffset = typeof template.baseline === "number"
    ? template.baseline - template.bbox.y0
    : Math.max(template.fontSize * 0.82, 1);

  const block = {
    id: createBlockId(`custom-page-${state.currentPage}`),
    page: state.currentPage,
    bbox: {
      x0,
      y0,
      x1: x0 + width,
      y1: y0 + height,
    },
    originalText: "",
    currentText: "",
    fontFamily: template.fontFamily,
    fontKey: template.fontKey,
    fontSize: template.fontSize,
    color: template.color,
    lineHeight: template.lineHeight,
    align: template.align,
    rotation: 0,
    groupKind: "manual",
    minFontSize: template.minFontSize,
    editable: true,
    cssFontFamily: template.cssFontFamily,
    fontAssetId: template.fontAssetId,
    fontWeight: template.fontWeight,
    fontStyle: template.fontStyle,
    baseline: y0 + baselineOffset,
    isCustom: true,
  };

  state.document.blocks.push(block);
  state.selectedBlockId = block.id;
  state.pendingFocusBlockId = block.id;
  renderTextLayer();
  scheduleDraftSave();
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
  state.drag = null;
  document.body.classList.remove("dragging-block");
}

function beginResize(block, node, event) {
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
    resizing: false,
  };
}

function clearResizeState() {
  state.resize = null;
  document.body.classList.remove("resizing-block");
}

function handleDragMove(event) {
  if (state.resize || !state.drag || !state.document?.supportStatus?.supported) {
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

  const nextWidth = Math.min(
    page.width - resize.originX0,
    Math.max(CUSTOM_BLOCK_MIN_WIDTH, resize.originWidth + (dx / state.zoom)),
  );
  const nextHeight = Math.min(
    page.height - resize.originY0,
    Math.max(CUSTOM_BLOCK_MIN_HEIGHT, resize.originHeight + (dy / state.zoom)),
  );

  block.bbox.x1 = block.bbox.x0 + nextWidth;
  block.bbox.y1 = block.bbox.y0 + nextHeight;

  if (resize.node) {
    resize.node.style.width = `${nextWidth * state.zoom}px`;
    resize.node.style.height = `${nextHeight * state.zoom}px`;
  }
}

function handleDragEnd() {
  if (state.resize) {
    const wasResizing = state.resize.resizing;
    clearResizeState();
    if (wasResizing) {
      scheduleDraftSave();
    }
  }

  if (state.drag) {
    const wasDragging = state.drag.dragging;
    clearDragState();
    if (wasDragging) {
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
  state.selectedBlockId = null;
  state.pendingFocusBlockId = null;
  renderTextLayer();
  scheduleDraftSave();
  setStatus("Manuelles Feld entfernt.");
  return true;
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

  syncSelectedBlock();
}

function hasEditableDocument() {
  return Boolean(state.document?.supportStatus?.supported);
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
  el.prevButton.hidden = !hasDoc;
  el.nextButton.hidden = !hasDoc;
  el.zoomOutButton.hidden = !hasDoc;
  el.zoomFitButton.hidden = !hasDoc;
  el.zoomInButton.hidden = !hasDoc;
  el.templateLearningButton.hidden = true;
  el.templateNameInput.hidden = true;
  el.saveTemplateButton.hidden = true;
  el.prevButton.disabled = !hasDoc || state.currentPage <= 1;
  el.nextButton.disabled = !hasDoc || state.currentPage >= (state.document?.pageCount ?? 0);
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
}

function updateMeta() {
  if (!hasEditableDocument()) {
    el.meta.textContent = "Whiteboard aktiv.";
    return;
  }

  const selected = getSelectedBlock();
  const selectedText = selected ? ` | Auswahl: ${selected.originalText.slice(0, 60)}` : "";
  const templateText = state.document?.detectedTemplateId ? ` | Vorlage: ${state.document.detectedTemplateId}` : "";
  const documentClassText = state.document?.documentClass ? ` | Typ: ${state.document.documentClass}` : "";
  const supportModeText = state.document?.supportStatus?.supportMode ? ` | Modus: ${state.document.supportStatus.supportMode}` : "";
  const embeddedText = state.document?.embeddedSessionFound ? " | Eingebettete Sitzung" : "";
  el.meta.textContent =
    `${state.document.sourcePath} | Seite ${state.currentPage}/${state.document.pageCount} | Zoom ${Math.round(state.zoom * 100)}%${documentClassText}${supportModeText}${templateText}${embeddedText}${selectedText}`;
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

async function loadBackground() {
  const page = getCurrentPageModel();
  if (!page) {
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
  const backgroundUrl = `${state.serviceBaseUrl}/documents/${state.document.id}/pages/${state.currentPage}/background?width=${targetWidth}&ts=${Date.now()}`;
  el.pageContainer.hidden = true;
  el.backgroundImage.hidden = true;
  el.backgroundImage.onload = () => {
    el.backgroundImage.style.width = `${displayWidth}px`;
    el.backgroundImage.style.height = `${displayHeight}px`;
    el.pageContainer.style.width = `${displayWidth}px`;
    el.pageContainer.style.height = `${displayHeight}px`;
    el.textLayer.style.width = `${displayWidth}px`;
    el.textLayer.style.height = `${displayHeight}px`;
    el.backgroundImage.hidden = false;
    el.pageContainer.hidden = false;
    renderTextLayer();
  };
  el.backgroundImage.onerror = () => {
    el.backgroundImage.hidden = true;
    el.pageContainer.hidden = true;
    el.textLayer.innerHTML = "";
    setStatus("PDF-Hintergrund konnte nicht geladen werden.");
  };
  el.backgroundImage.src = backgroundUrl;
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

    const node = document.createElement("div");
    node.className = "text-block";
    if (block.isCheckbox) {
      node.classList.add("checkbox-block");
    }
    if (block.isCustom) {
      node.classList.add("custom-block");
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
    node.style.setProperty("--masked-text-color", block.color || "#000");

    if (block.isCheckbox) {
      const renderMark = block.currentText.trim() && !isUnchangedScanGeneratedCheckbox(block);
      node.title = "Ankreuzen";
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
      });
      node.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        toggleCheckbox(block);
      });
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
    editor.contentEditable = "plaintext-only";
    if (block.currentText) {
      editor.textContent = block.currentText;
    } else {
      ensureEditablePlaceholder(editor);
    }
    editor.style.fontFamily = block.cssFontFamily;
    editor.style.fontWeight = block.fontWeight || "400";
    editor.style.fontStyle = block.fontStyle || "normal";
    editor.style.fontSize = `${block.fontSize * scale}px`;
    editor.style.lineHeight = `${block.lineHeight * scale}px`;
    editor.style.color = block.color;
    editor.style.textAlign = block.align;

    editor.addEventListener("focus", () => {
      state.selectedBlockId = block.id;
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

    editor.addEventListener("mousedown", () => {
      state.selectedBlockId = block.id;
      syncSelectedBlock();
    });

    editor.addEventListener("keydown", (event) => {
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
      autoSizeTextBlock(node, editor, block, scale);
      scheduleDraftSave();
    });

    editor.addEventListener("blur", () => {
      block.currentText = readEditableText(editor);
      if (!block.currentText) {
        ensureEditablePlaceholder(editor);
      }
      if (isScanGeneratedTextField(block) && isScanFieldChanged(block)) {
        renderTextLayer();
      }
    });

    node.appendChild(editor);
    el.textLayer.appendChild(node);
    if (block.isCustom) {
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
    throw new Error("Bitte zuerst per Rechtsklick die benötigten Felder setzen.");
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
  state.lastExportPath = null;
  if (el.templateNameInput) {
    el.templateNameInput.value = getDefaultTemplateName();
  }
  injectFontFaces();
  showSupportStatus();

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
  setStatus("PDF geladen.");
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
      ? "Lernmodus aktiv. Felder per Rechtsklick setzen und danach die Vorlage speichern."
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
  await loadBackground();
  updateButtons();
});

el.zoomFitButton.addEventListener("click", async () => {
  state.zoom = state.fitZoom;
  await loadBackground();
  updateButtons();
});

el.zoomInButton.addEventListener("click", async () => {
  state.zoom = clamp(state.zoom * 1.15, 0.25, 4);
  await loadBackground();
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
  if (!state.document?.supportStatus?.supported || el.pageContainer.hidden) {
    return;
  }
  event.preventDefault();
  createCustomTextBlock(event);
});

el.pageContainer.addEventListener("click", (event) => {
  if (!state.document?.supportStatus?.supported || el.pageContainer.hidden || !isScanTemplateDocument()) {
    return;
  }

  const target = event.target;
  if (!(target instanceof Element) || target.closest(".text-block")) {
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

  if (state.drag || state.resize) {
    return;
  }

  clearSelection();
});

window.addEventListener("mousemove", handleResizeMove);
window.addEventListener("mousemove", handleDragMove);
window.addEventListener("mouseup", handleDragEnd);

window.addEventListener("keydown", (event) => {
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
