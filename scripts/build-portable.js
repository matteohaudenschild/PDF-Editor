const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const projectRoot = path.resolve(__dirname, "..");
const electronDistRoot = path.join(projectRoot, "node_modules", "electron", "dist");
const outputRoot = path.join(projectRoot, "dist", "portable-win");
const resourcesAppRoot = path.join(outputRoot, "resources", "app");
const backendSourceRoot = path.join(projectRoot, "backend");
const appSourceRoot = path.join(projectRoot, "app");
const appDisplayName = "PDF Editor";
const appExecutableName = `${appDisplayName}.exe`;
const launcherOutputPath = path.join(projectRoot, appExecutableName);
const windowsIconPath = path.join(projectRoot, "app", "assets", "pdf-editor-old.ico");
const rootPackageJson = JSON.parse(fs.readFileSync(path.join(projectRoot, "package.json"), "utf8"));

function ensureExists(targetPath, description) {
  if (!fs.existsSync(targetPath)) {
    throw new Error(`${description} nicht gefunden: ${targetPath}`);
  }
}

function copyRecursive(sourcePath, targetPath) {
  const stats = fs.statSync(sourcePath);
  if (stats.isDirectory()) {
    fs.mkdirSync(targetPath, { recursive: true });
    for (const entry of fs.readdirSync(sourcePath)) {
      copyRecursive(path.join(sourcePath, entry), path.join(targetPath, entry));
    }
    return;
  }

  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.copyFileSync(sourcePath, targetPath);
}

function writeAppPackageJson() {
  const appPackageJson = {
    name: "pdf-desktop-editor",
    productName: appDisplayName,
    version: rootPackageJson.version,
    description: rootPackageJson.description,
    main: "app/src/main/main.js",
  };

  fs.writeFileSync(
    path.join(resourcesAppRoot, "package.json"),
    `${JSON.stringify(appPackageJson, null, 2)}\n`,
    "utf8",
  );
}

function findCSharpCompiler() {
  const candidates = [
    path.join(process.env.WINDIR || "C:\\Windows", "Microsoft.NET", "Framework64", "v4.0.30319", "csc.exe"),
    path.join(process.env.WINDIR || "C:\\Windows", "Microsoft.NET", "Framework", "v4.0.30319", "csc.exe"),
  ];

  return candidates.find((candidate) => fs.existsSync(candidate)) || null;
}

function getRceditExecutablePath() {
  const executableName = process.arch === "x64" || process.arch === "arm64"
    ? "rcedit-x64.exe"
    : "rcedit.exe";
  return path.join(projectRoot, "node_modules", "rcedit", "bin", executableName);
}

function applyWindowsExecutableIcon(executablePath) {
  ensureExists(windowsIconPath, "Windows-App-Icon");
  const rceditPath = getRceditExecutablePath();
  ensureExists(rceditPath, "rcedit");

  const result = spawnSync(
    rceditPath,
    [
      executablePath,
      "--set-icon",
      windowsIconPath,
      "--set-version-string",
      "FileDescription",
      appDisplayName,
      "--set-version-string",
      "ProductName",
      appDisplayName,
      "--set-version-string",
      "OriginalFilename",
      appExecutableName,
    ],
    {
      cwd: projectRoot,
      encoding: "utf8",
    },
  );

  if (result.status !== 0) {
    throw new Error(
      `Windows-Icon konnte nicht in die EXE geschrieben werden.\n${result.stdout || ""}\n${result.stderr || ""}`.trim(),
    );
  }
}

function buildRootLauncher() {
  const compilerPath = findCSharpCompiler();
  if (!compilerPath) {
    console.warn("Kein C#-Compiler gefunden. Root-Launcher-EXE wurde nicht erstellt.");
    return;
  }

  const sourceRoot = path.join(projectRoot, "scripts", ".generated");
  const sourcePath = path.join(sourceRoot, "root-launcher.cs");
  fs.mkdirSync(sourceRoot, { recursive: true });

  const launcherSource = `using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var targetPath = Path.Combine(baseDir, "dist", "portable-win", "${appExecutableName}");
        if (!File.Exists(targetPath))
        {
            MessageBox.Show(
                "Die portable App wurde nicht gefunden. Erwartet wurde:\\n" + targetPath,
                "${appDisplayName}",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = targetPath,
            WorkingDirectory = Path.GetDirectoryName(targetPath),
            UseShellExecute = true
        });
    }
}
`;

  fs.writeFileSync(sourcePath, launcherSource, "utf8");

  const compileResult = spawnSync(
    compilerPath,
    [
      "/nologo",
      "/target:winexe",
      "/reference:System.dll",
      "/reference:System.Windows.Forms.dll",
      "/reference:System.Drawing.dll",
      `/win32icon:${windowsIconPath}`,
      `/out:${launcherOutputPath}`,
      sourcePath,
    ],
    {
      cwd: projectRoot,
      encoding: "utf8",
    },
  );

  if (compileResult.status !== 0) {
    throw new Error(
      `Root-Launcher konnte nicht gebaut werden.\n${compileResult.stdout || ""}\n${compileResult.stderr || ""}`.trim(),
    );
  }
}

function buildPortable() {
  ensureExists(electronDistRoot, "Electron-Runtime");
  ensureExists(appSourceRoot, "App-Ordner");
  ensureExists(path.join(backendSourceRoot, ".venv"), "Backend-Venv");

  fs.rmSync(outputRoot, { recursive: true, force: true });
  copyRecursive(electronDistRoot, outputRoot);

  const electronExePath = path.join(outputRoot, "electron.exe");
  const portableExePath = path.join(outputRoot, appExecutableName);
  if (fs.existsSync(portableExePath)) {
    fs.rmSync(portableExePath, { force: true });
  }
  fs.renameSync(electronExePath, portableExePath);
  applyWindowsExecutableIcon(portableExePath);

  fs.mkdirSync(resourcesAppRoot, { recursive: true });
  writeAppPackageJson();
  copyRecursive(appSourceRoot, path.join(resourcesAppRoot, "app"));

  const backendTargetRoot = path.join(resourcesAppRoot, "backend");
  fs.mkdirSync(backendTargetRoot, { recursive: true });
  copyRecursive(path.join(backendSourceRoot, "server.py"), path.join(backendTargetRoot, "server.py"));
  copyRecursive(path.join(backendSourceRoot, "requirements.txt"), path.join(backendTargetRoot, "requirements.txt"));
  copyRecursive(path.join(backendSourceRoot, "pdf_editor_service"), path.join(backendTargetRoot, "pdf_editor_service"));
  const tessdataSourceRoot = path.join(backendSourceRoot, "tessdata");
  if (fs.existsSync(tessdataSourceRoot)) {
    copyRecursive(tessdataSourceRoot, path.join(backendTargetRoot, "tessdata"));
  }
  copyRecursive(path.join(backendSourceRoot, ".venv"), path.join(backendTargetRoot, ".venv"));
  buildRootLauncher();

  console.log(`Portable build bereit: ${portableExePath}`);
}

if (require.main === module) {
  buildPortable();
}

module.exports = {
  buildPortable,
  buildRootLauncher,
  copyRecursive,
  ensureExists,
  findCSharpCompiler,
};
