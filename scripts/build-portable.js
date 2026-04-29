const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const projectRoot = path.resolve(__dirname, "..");
const electronDistRoot = path.join(projectRoot, "node_modules", "electron", "dist");
const outputRoot = path.join(projectRoot, "dist", "portable-win");
const resourcesAppRoot = path.join(outputRoot, "resources", "app");
const backendSourceRoot = path.join(projectRoot, "backend");
const appSourceRoot = path.join(projectRoot, "app");
const launcherOutputPath = path.join(projectRoot, "PDF Desktop Editor.exe");
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
    productName: "PDF Desktop Editor",
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
        var targetPath = Path.Combine(baseDir, "dist", "portable-win", "PDF Desktop Editor.exe");
        if (!File.Exists(targetPath))
        {
            MessageBox.Show(
                "Die portable App wurde nicht gefunden. Erwartet wurde:\\n" + targetPath,
                "PDF Desktop Editor",
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
      `/win32icon:${path.join(projectRoot, "app", "assets", "pdf-editor-icon.ico")}`,
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
  const portableExePath = path.join(outputRoot, "PDF Desktop Editor.exe");
  if (fs.existsSync(portableExePath)) {
    fs.rmSync(portableExePath, { force: true });
  }
  fs.renameSync(electronExePath, portableExePath);

  fs.mkdirSync(resourcesAppRoot, { recursive: true });
  writeAppPackageJson();
  copyRecursive(appSourceRoot, path.join(resourcesAppRoot, "app"));

  const backendTargetRoot = path.join(resourcesAppRoot, "backend");
  fs.mkdirSync(backendTargetRoot, { recursive: true });
  copyRecursive(path.join(backendSourceRoot, "server.py"), path.join(backendTargetRoot, "server.py"));
  copyRecursive(path.join(backendSourceRoot, "requirements.txt"), path.join(backendTargetRoot, "requirements.txt"));
  copyRecursive(path.join(backendSourceRoot, "pdf_editor_service"), path.join(backendTargetRoot, "pdf_editor_service"));
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
