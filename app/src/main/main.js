const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require("electron");
const { spawn, spawnSync } = require("child_process");
const net = require("net");
const https = require("https");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");

let mainWindow = null;
let backendProcess = null;
let serviceBaseUrl = null;

const APP_DISPLAY_NAME = "PDF Editor";
const UPDATE_REPOSITORY = "matteohaudenschild/PDF-Editor";
const UPDATE_ASSET_NAME = `${APP_DISPLAY_NAME} Setup.exe`;
const UPDATE_API_URL = `https://api.github.com/repos/${UPDATE_REPOSITORY}/releases/latest`;
const UPDATE_TEMP_DIR_NAMES = [APP_DISPLAY_NAME, "PDF Desktop Editor"];

app.setName(APP_DISPLAY_NAME);

if (process.platform === "win32") {
  app.setAppUserModelId("com.sicherheitnord.pdf-editor");
}

function removeNativeWindowMenu(browserWindow) {
  if (!browserWindow) {
    return;
  }
  Menu.setApplicationMenu(null);
  if (typeof browserWindow.removeMenu === "function") {
    browserWindow.removeMenu();
  }
  browserWindow.setMenu(null);
  browserWindow.setAutoHideMenuBar(true);
  browserWindow.setMenuBarVisibility(false);
}

app.on("browser-window-created", (_event, browserWindow) => {
  removeNativeWindowMenu(browserWindow);
  browserWindow.once("ready-to-show", () => {
    removeNativeWindowMenu(browserWindow);
  });
});

function getProjectRoot() {
  if (app.isPackaged) {
    const asarRoot = path.join(process.resourcesPath, "app.asar");
    if (fs.existsSync(asarRoot)) {
      return asarRoot;
    }
    const unpackedRoot = path.join(process.resourcesPath, "app");
    if (fs.existsSync(unpackedRoot)) {
      return unpackedRoot;
    }
  }
  return path.resolve(__dirname, "..", "..", "..");
}

function getBackendRoot() {
  if (app.isPackaged) {
    const packagedBackendRoot = path.join(process.resourcesPath, "backend");
    if (fs.existsSync(packagedBackendRoot)) {
      return packagedBackendRoot;
    }
  }
  return path.join(getProjectRoot(), "backend");
}

function getAppIconPath() {
  return path.join(getProjectRoot(), "app", "assets", "pdf-editor-old.ico");
}

function isManagedUpdateTempFile(fileName) {
  const lowerName = String(fileName || "").toLowerCase();
  const setupNames = ["pdf editor setup", "pdf desktop editor setup"];
  return setupNames.some((setupName) => lowerName === `${setupName}.exe`
    || lowerName.startsWith(`${setupName} `))
    && (lowerName.endsWith(".exe") || lowerName.endsWith(".exe.download"));
}

function cleanupUpdateTempFiles({ keepPath = null } = {}) {
  const normalizedKeepPath = keepPath ? path.resolve(keepPath).toLowerCase() : null;
  for (const dirName of UPDATE_TEMP_DIR_NAMES) {
    const updateTempDir = path.join(app.getPath("temp"), dirName);
    try {
      if (!fs.existsSync(updateTempDir)) {
        continue;
      }
      for (const entry of fs.readdirSync(updateTempDir, { withFileTypes: true })) {
        if (!entry.isFile() || !isManagedUpdateTempFile(entry.name)) {
          continue;
        }
        const filePath = path.join(updateTempDir, entry.name);
        if (normalizedKeepPath && path.resolve(filePath).toLowerCase() === normalizedKeepPath) {
          continue;
        }
        fs.rmSync(filePath, { force: true });
      }
      if (fs.readdirSync(updateTempDir).length === 0) {
        fs.rmdirSync(updateTempDir);
      }
    } catch (error) {
      console.warn(`Update-Temp-Aufraeumen fehlgeschlagen: ${error.message}`);
    }
  }
}

function normalizeReleaseVersion(value) {
  return String(value || "").trim().replace(/^v/i, "");
}

function compareVersions(left, right) {
  const leftParts = normalizeReleaseVersion(left).split(/[.-]/).map((part) => Number.parseInt(part, 10) || 0);
  const rightParts = normalizeReleaseVersion(right).split(/[.-]/).map((part) => Number.parseInt(part, 10) || 0);
  const length = Math.max(leftParts.length, rightParts.length);
  for (let index = 0; index < length; index += 1) {
    const leftPart = leftParts[index] || 0;
    const rightPart = rightParts[index] || 0;
    if (leftPart > rightPart) {
      return 1;
    }
    if (leftPart < rightPart) {
      return -1;
    }
  }
  return 0;
}

function githubRequestBuffer(url, redirectCount = 0) {
  return new Promise((resolve, reject) => {
    const request = https.get(
      url,
      {
        headers: {
          "Accept": "application/vnd.github+json",
          "User-Agent": "PDF-Editor-Updater",
          "X-GitHub-Api-Version": "2022-11-28",
        },
      },
      (response) => {
        const location = response.headers.location;
        if (response.statusCode >= 300 && response.statusCode < 400 && location) {
          response.resume();
          if (redirectCount >= 5) {
            reject(new Error("Zu viele Weiterleitungen beim Update-Download."));
            return;
          }
          githubRequestBuffer(new URL(location, url).toString(), redirectCount + 1).then(resolve, reject);
          return;
        }

        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => {
          const body = Buffer.concat(chunks);
          if (response.statusCode < 200 || response.statusCode >= 300) {
            reject(new Error(`GitHub antwortete mit Status ${response.statusCode}: ${body.toString("utf8").slice(0, 300)}`));
            return;
          }
          resolve(body);
        });
      },
    );

    request.setTimeout(30000, () => {
      request.destroy(new Error("Update-Pruefung hat zu lange gedauert."));
    });
    request.on("error", reject);
  });
}

async function githubRequestJson(url) {
  const buffer = await githubRequestBuffer(url);
  return JSON.parse(buffer.toString("utf8"));
}

function findInstallerAsset(release) {
  const assets = Array.isArray(release?.assets) ? release.assets : [];
  const normalizeAssetName = (value) => String(value || "").toLowerCase().replace(/[\s._-]+/g, "");
  const expectedName = normalizeAssetName(UPDATE_ASSET_NAME);
  return assets.find((asset) => normalizeAssetName(asset?.name) === expectedName)
    || assets.find((asset) => String(asset?.name || "").toLowerCase().endsWith(".exe"));
}

function extractSha256(notes) {
  const match = String(notes || "").match(/\bSHA256\s*:\s*([a-f0-9]{64})\b/i);
  return match ? match[1].toUpperCase() : null;
}

function getFileSha256(filePath) {
  const hash = crypto.createHash("sha256");
  hash.update(fs.readFileSync(filePath));
  return hash.digest("hex").toUpperCase();
}

async function getLatestUpdateInfo() {
  const release = await githubRequestJson(UPDATE_API_URL);
  const currentVersion = app.getVersion();
  const latestVersion = normalizeReleaseVersion(release.tag_name || release.name);
  const asset = findInstallerAsset(release);
  const available = latestVersion
    ? compareVersions(currentVersion, latestVersion) < 0
    : false;

  return {
    available,
    currentVersion,
    latestVersion,
    releaseUrl: release.html_url || `https://github.com/${UPDATE_REPOSITORY}/releases/latest`,
    releaseName: release.name || release.tag_name || "",
    publishedAt: release.published_at || "",
    notes: release.body || "",
    assetName: asset?.name || null,
    assetSize: asset?.size || null,
    downloadUrl: asset?.browser_download_url || null,
    expectedSha256: extractSha256(release.body),
  };
}

function downloadFile(url, targetPath, redirectCount = 0) {
  const partialPath = `${targetPath}.download`;
  return new Promise((resolve, reject) => {
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });
    if (redirectCount === 0) {
      fs.rmSync(targetPath, { force: true });
      fs.rmSync(partialPath, { force: true });
    }
    const request = https.get(
      url,
      {
        headers: {
          "User-Agent": "PDF-Editor-Updater",
        },
      },
      (response) => {
        const location = response.headers.location;
        if (response.statusCode >= 300 && response.statusCode < 400 && location) {
          response.resume();
          if (redirectCount >= 5) {
            reject(new Error("Zu viele Weiterleitungen beim Update-Download."));
            return;
          }
          downloadFile(new URL(location, url).toString(), targetPath, redirectCount + 1).then(resolve, reject);
          return;
        }

        if (response.statusCode < 200 || response.statusCode >= 300) {
          response.resume();
          fs.rm(targetPath, { force: true }, () => {});
          fs.rm(partialPath, { force: true }, () => {});
          reject(new Error(`Update-Download fehlgeschlagen: HTTP ${response.statusCode}`));
          return;
        }

        const file = fs.createWriteStream(partialPath);
        response.pipe(file);
        file.on("finish", () => {
          file.close(() => {
            try {
              fs.renameSync(partialPath, targetPath);
              resolve(targetPath);
            } catch (error) {
              fs.rm(targetPath, { force: true }, () => {});
              fs.rm(partialPath, { force: true }, () => {});
              reject(error);
            }
          });
        });
        file.on("error", (error) => {
          request.destroy();
          fs.rm(targetPath, { force: true }, () => {});
          fs.rm(partialPath, { force: true }, () => {});
          reject(error);
        });
      },
    );

    request.setTimeout(600000, () => {
      request.destroy(new Error("Update-Download hat zu lange gedauert."));
    });
    request.on("error", (error) => {
      fs.rm(targetPath, { force: true }, () => {});
      fs.rm(partialPath, { force: true }, () => {});
      reject(error);
    });
  });
}

async function downloadAndInstallLatestUpdate() {
  const updateInfo = await getLatestUpdateInfo();
  if (!updateInfo.available) {
    throw new Error("Es ist keine neuere Version verfuegbar.");
  }
  if (!updateInfo.downloadUrl) {
    throw new Error("Im neuesten GitHub-Release wurde keine Setup-Datei gefunden.");
  }

  const safeVersion = updateInfo.latestVersion.replace(/[^0-9A-Za-z._-]/g, "_");
  const installerPath = path.join(app.getPath("temp"), APP_DISPLAY_NAME, `${APP_DISPLAY_NAME} Setup ${safeVersion}.exe`);
  cleanupUpdateTempFiles({ keepPath: installerPath });
  await downloadFile(updateInfo.downloadUrl, installerPath);

  if (updateInfo.expectedSha256) {
    const actualSha256 = getFileSha256(installerPath);
    if (actualSha256 !== updateInfo.expectedSha256) {
      fs.rmSync(installerPath, { force: true });
      throw new Error("Das Update wurde heruntergeladen, aber die Pruefsumme stimmt nicht.");
    }
  }

  const installer = spawn(installerPath, ["--auto-update", "--wait-pid", String(process.pid)], {
    detached: true,
    stdio: "ignore",
    windowsHide: false,
  });
  installer.unref();

  setTimeout(() => {
    stopBackend();
    app.quit();
  }, 500);

  return {
    started: true,
    installerPath,
    latestVersion: updateInfo.latestVersion,
  };
}

function getBackendPythonCandidates() {
  const backendRoot = getBackendRoot();
  const rawCandidates = [
    {
      label: "eingebettete Python-Laufzeit",
      command: path.join(backendRoot, "python-embed", "python.exe"),
      args: [],
    },
    {
      label: "mitgelieferte Python-Umgebung",
      command: path.join(backendRoot, ".venv", "Scripts", "python.exe"),
      args: [],
    },
    {
      label: "lokales Python 3.12",
      command: path.join(process.env.LOCALAPPDATA || "", "Programs", "Python", "Python312", "python.exe"),
      args: [],
    },
    {
      label: "Python Launcher (py -3.12)",
      command: "py",
      args: ["-3.12"],
    },
    {
      label: "python aus PATH",
      command: "python",
      args: [],
    },
  ];

  const seen = new Set();
  const candidates = [];
  for (const candidate of rawCandidates) {
    const identity = `${candidate.command}::${candidate.args.join(" ")}`;
    if (seen.has(identity)) {
      continue;
    }
    seen.add(identity);
    if (candidate.command === "py" || candidate.command === "python" || fs.existsSync(candidate.command)) {
      candidates.push(candidate);
    }
  }

  return candidates;
}

function buildPythonEnv(backendRoot) {
  const sitePackagesPath = path.join(backendRoot, ".venv", "Lib", "site-packages");
  const pythonPathEntries = [backendRoot];
  if (fs.existsSync(sitePackagesPath)) {
    pythonPathEntries.push(sitePackagesPath);
  }
  if (process.env.PYTHONPATH) {
    pythonPathEntries.push(process.env.PYTHONPATH);
  }

  return {
    ...process.env,
    PYTHONPATH: pythonPathEntries.join(path.delimiter),
  };
}

function resolveBackendPythonRuntime() {
  const backendRoot = getBackendRoot();
  const env = buildPythonEnv(backendRoot);
  const probeCode = [
    "import fastapi",
    "import uvicorn",
    "import pymupdf",
    "import pydantic",
    "print('ok')",
  ].join("; ");
  const probeFailures = [];

  for (const candidate of getBackendPythonCandidates()) {
    const probeResult = spawnSync(
      candidate.command,
      [...candidate.args, "-c", probeCode],
      {
        cwd: backendRoot,
        env,
        encoding: "utf8",
        timeout: 10000,
        windowsHide: true,
      },
    );

    if (!probeResult.error && probeResult.status === 0) {
      return {
        ...candidate,
        env,
      };
    }

    const reason = (
      probeResult.error?.message
      || probeResult.stderr
      || probeResult.stdout
      || `Exit-Code ${probeResult.status}`
    ).trim();
    probeFailures.push(`- ${candidate.label}: ${reason}`);
  }

  throw new Error(
    [
      "Keine nutzbare Python-Laufzeit für den lokalen PDF-Backend-Dienst gefunden.",
      "Die portable .venv ist auf anderen Windows-PCs oft nicht direkt portabel.",
      "Bitte die App über den Setup-Installer installieren, damit die eingebettete Python-Laufzeit automatisch geladen wird.",
      "",
      "Diagnose:",
      ...probeFailures,
    ].join("\n"),
  );
}

function getFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      const port = typeof address === "object" && address ? address.port : null;
      server.close(() => resolve(port));
    });
  });
}

async function waitForHealth(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(`${url}/healthz`);
      if (response.ok) {
        return;
      }
    } catch (error) {
      // ignore until timeout
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error("Lokaler Backend-Dienst wurde nicht rechtzeitig bereit.");
}

async function startBackend() {
  const port = await getFreePort();
  const backendRoot = getBackendRoot();
  const backendEntry = path.join(backendRoot, "server.py");
  const runtime = resolveBackendPythonRuntime();

  backendProcess = spawn(runtime.command, [...runtime.args, backendEntry, "--port", String(port)], {
    cwd: backendRoot,
    env: runtime.env,
    stdio: ["ignore", "pipe", "pipe"],
    windowsHide: true,
  });

  backendProcess.stdout.on("data", (chunk) => {
    process.stdout.write(`[backend] ${chunk}`);
  });

  backendProcess.stderr.on("data", (chunk) => {
    process.stderr.write(`[backend] ${chunk}`);
  });

  backendProcess.on("exit", (code) => {
    backendProcess = null;
    if (code !== 0) {
      console.error(`Backend exited with code ${code}`);
    }
  });

  serviceBaseUrl = `http://127.0.0.1:${port}`;
  await waitForHealth(serviceBaseUrl);
}

function stopBackend() {
  if (backendProcess) {
    backendProcess.kill();
    backendProcess = null;
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    title: APP_DISPLAY_NAME,
    width: 1400,
    height: 960,
    icon: getAppIconPath(),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  removeNativeWindowMenu(mainWindow);
  mainWindow.loadFile(path.join(getProjectRoot(), "app", "src", "renderer", "index.html"));
}

ipcMain.handle("service:get-base-url", () => serviceBaseUrl);
ipcMain.handle("app:get-version", () => app.getVersion());

ipcMain.handle("updates:check", async () => {
  try {
    return await getLatestUpdateInfo();
  } catch (error) {
    return {
      available: false,
      currentVersion: app.getVersion(),
      latestVersion: null,
      error: error instanceof Error ? error.message : String(error),
    };
  }
});

ipcMain.handle("updates:install", async () => {
  return downloadAndInstallLatestUpdate();
});

ipcMain.handle("dialog:open-pdf", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [{ name: "PDF", extensions: ["pdf"] }],
  });
  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
});

ipcMain.handle("dialog:save-pdf", async (_event, suggestedName = "whiteboard.pdf") => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: path.join(app.getPath("documents"), suggestedName),
    filters: [{ name: "PDF", extensions: ["pdf"] }],
  });
  if (result.canceled || !result.filePath) {
    return null;
  }
  return result.filePath;
});

ipcMain.handle("shell:reveal-file", async (_event, targetPath) => {
  if (!targetPath) {
    return false;
  }
  shell.showItemInFolder(targetPath);
  return true;
});

app.whenReady().then(async () => {
  try {
    Menu.setApplicationMenu(null);
    cleanupUpdateTempFiles();
    await startBackend();
    createWindow();

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  } catch (error) {
    dialog.showErrorBox("Start fehlgeschlagen", error instanceof Error ? error.message : String(error));
    stopBackend();
    app.quit();
  }
});

app.on("window-all-closed", () => {
  stopBackend();
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", () => {
  stopBackend();
});
