const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const { spawn, spawnSync } = require("child_process");
const net = require("net");
const path = require("path");
const fs = require("fs");

let mainWindow = null;
let backendProcess = null;
let serviceBaseUrl = null;

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
  return path.join(getProjectRoot(), "app", "assets", "pdf-editor-icon.ico");
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
    width: 1400,
    height: 960,
    icon: getAppIconPath(),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(getProjectRoot(), "app", "src", "renderer", "index.html"));
}

ipcMain.handle("service:get-base-url", () => serviceBaseUrl);

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
