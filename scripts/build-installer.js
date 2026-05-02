const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const {
  buildPortable,
  copyRecursive,
  ensureExists,
  findCSharpCompiler,
} = require("./build-portable");

const projectRoot = path.resolve(__dirname, "..");
const distRoot = path.join(projectRoot, "dist");
const portableRoot = path.join(distRoot, "portable-win");
const appDisplayName = "PDF Editor";
const appExecutableName = `${appDisplayName}.exe`;
const installerFileName = `${appDisplayName} Setup.exe`;
const rootLauncherPath = path.join(projectRoot, appExecutableName);
const installerPayloadRoot = path.join(distRoot, "installer-payload");
const installerPayloadZip = path.join(distRoot, "pdf-desktop-editor-setup-payload.zip");
const installerOutputPath = path.join(distRoot, installerFileName);
const installerResourceName = "PdfDesktopEditorPayload.zip";
const pythonEmbedUrl = "https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip";
const pythonEmbedZip = path.join(distRoot, "python-3.12.10-embed-amd64.zip");
const laptopCopyRoot = path.join(distRoot, "ZUM_LAPTOP_KOPIEREN");
const laptopCopyInstallerPath = path.join(laptopCopyRoot, installerFileName);
const laptopCopyInstructionsPath = path.join(laptopCopyRoot, "INSTALLATION.txt");
const laptopCopyZip = path.join(distRoot, "ZUM_LAPTOP_KOPIEREN.zip");

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: projectRoot,
    encoding: "utf8",
    stdio: "pipe",
    ...options,
  });

  if (result.status !== 0) {
    throw new Error(
      [
        `Befehl fehlgeschlagen: ${command} ${args.join(" ")}`,
        result.stdout?.trim(),
        result.stderr?.trim(),
      ].filter(Boolean).join("\n"),
    );
  }

  return result;
}

function preparePortableBundle() {
  if (process.env.SKIP_PORTABLE_REBUILD === "1") {
    ensureExists(portableRoot, "Portable-App");
    ensureExists(rootLauncherPath, "Root-Launcher");
    return;
  }

  buildPortable();
  ensureExists(portableRoot, "Portable-App");
  ensureExists(rootLauncherPath, "Root-Launcher");
}

function configureEmbeddedPython(embeddedPythonRoot) {
  const pthFile = fs.readdirSync(embeddedPythonRoot).find((entry) => /^python.*\._pth$/i.test(entry));
  if (!pthFile) {
    throw new Error(`Python-Embed-Konfiguration nicht gefunden: ${embeddedPythonRoot}`);
  }

  const pthContent = [
    "python312.zip",
    ".",
    "..",
    "..\\.venv\\Lib\\site-packages",
    "import site",
    "",
  ].join("\r\n");
  fs.writeFileSync(path.join(embeddedPythonRoot, pthFile), pthContent, "utf8");
}

function probeEmbeddedPython(embeddedPythonRoot, backendRoot) {
  const pythonExe = path.join(embeddedPythonRoot, "python.exe");
  if (!fs.existsSync(pythonExe)) {
    return false;
  }

  const result = spawnSync(
    pythonExe,
    ["-c", "import fastapi, uvicorn, pymupdf, pydantic; print('ok')"],
    {
      cwd: backendRoot,
      encoding: "utf8",
      stdio: "pipe",
      timeout: 15000,
      windowsHide: true,
    },
  );
  return !result.error && result.status === 0;
}

function ensurePythonEmbedZip() {
  if (fs.existsSync(pythonEmbedZip)) {
    return;
  }

  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$ProgressPreference = 'SilentlyContinue'",
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12",
    `$url = '${pythonEmbedUrl}'`,
    `$out = '${pythonEmbedZip.replace(/'/g, "''")}'`,
    "Invoke-WebRequest -Uri $url -OutFile $out",
  ].join("; ");
  run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]);
  ensureExists(pythonEmbedZip, "Python-Embed-ZIP");
}

function ensurePortableEmbeddedPython() {
  const backendRoot = path.join(portableRoot, "resources", "app", "backend");
  const embeddedPythonRoot = path.join(backendRoot, "python-embed");

  if (fs.existsSync(embeddedPythonRoot)) {
    configureEmbeddedPython(embeddedPythonRoot);
    if (probeEmbeddedPython(embeddedPythonRoot, backendRoot)) {
      return;
    }
  }

  ensurePythonEmbedZip();
  fs.rmSync(embeddedPythonRoot, { recursive: true, force: true });
  fs.mkdirSync(embeddedPythonRoot, { recursive: true });

  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$zip = '${pythonEmbedZip.replace(/'/g, "''")}'`,
    `$target = '${embeddedPythonRoot.replace(/'/g, "''")}'`,
    "Expand-Archive -Path $zip -DestinationPath $target -Force",
  ].join("; ");
  run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]);
  configureEmbeddedPython(embeddedPythonRoot);

  if (!probeEmbeddedPython(embeddedPythonRoot, backendRoot)) {
    throw new Error("Die eingebettete Python-Laufzeit konnte nicht gegen die mitgelieferten Backend-Pakete geprueft werden.");
  }
}

function stageInstallerPayload() {
  fs.rmSync(installerPayloadRoot, { recursive: true, force: true });
  fs.mkdirSync(installerPayloadRoot, { recursive: true });

  copyRecursive(rootLauncherPath, path.join(installerPayloadRoot, appExecutableName));
  copyRecursive(portableRoot, path.join(installerPayloadRoot, "dist", "portable-win"));
}

function buildPayloadZip() {
  fs.rmSync(installerPayloadZip, { force: true });

  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$source = '${installerPayloadRoot.replace(/'/g, "''")}'`,
    `$zip = '${installerPayloadZip.replace(/'/g, "''")}'`,
    "if (Test-Path $zip) { Remove-Item $zip -Force }",
    "Compress-Archive -Path (Join-Path $source '*') -DestinationPath $zip -CompressionLevel Optimal",
  ].join("; ");

  run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]);
  ensureExists(installerPayloadZip, "Installer-Payload-ZIP");
}

function buildInstallerSource() {
  const sourceRoot = path.join(projectRoot, "scripts", ".generated");
  const sourcePath = path.join(sourceRoot, "setup-bootstrap.cs");
  fs.mkdirSync(sourceRoot, { recursive: true });

  const installerSource = `using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        var autoUpdate = HasFlag(args, "--auto-update") || HasFlag(args, "/auto-update");
        var waitPid = GetWaitPid(args);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new InstallerForm(autoUpdate, waitPid));
    }

    private static bool HasFlag(string[] args, string flag)
    {
        foreach (var arg in args)
        {
            if (string.Equals(arg, flag, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    private static int? GetWaitPid(string[] args)
    {
        for (var index = 0; index < args.Length - 1; index++)
        {
            int pid;
            if (string.Equals(args[index], "--wait-pid", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(args[index + 1], out pid))
            {
                return pid;
            }
        }
        return null;
    }
}

internal sealed class InstallerForm : Form
{
    private const string AppName = "${appDisplayName}";
    private const string AppExecutableName = "${appExecutableName}";
    private const string PythonEmbedUrl = "https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip";
    private const string PayloadResourceName = "${installerResourceName}";

    private readonly Label _titleLabel;
    private readonly Label _statusLabel;
    private readonly ProgressBar _progressBar;

    private readonly bool _autoUpdate;
    private readonly int? _waitPid;
    private bool _allowClose;
    private bool _started;

    public InstallerForm(bool autoUpdate, int? waitPid)
    {
        _autoUpdate = autoUpdate;
        _waitPid = waitPid;

        Text = AppName + (_autoUpdate ? " Update" : " Setup");
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(520, 156);
        Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

        _titleLabel = new Label
        {
            Left = 18,
            Top = 18,
            Width = 484,
            Height = 26,
            Text = AppName + (_autoUpdate ? " wird aktualisiert" : " wird installiert"),
            Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point),
        };
        Controls.Add(_titleLabel);

        _statusLabel = new Label
        {
            Left = 18,
            Top = 60,
            Width = 484,
            Height = 34,
            Text = "Vorbereitung...",
        };
        Controls.Add(_statusLabel);

        _progressBar = new ProgressBar
        {
            Left = 18,
            Top = 104,
            Width = 484,
            Height = 18,
            Style = ProgressBarStyle.Marquee,
        };
        Controls.Add(_progressBar);
    }

    protected override async void OnShown(EventArgs e)
    {
        base.OnShown(e);
        if (_started)
        {
            return;
        }

        _started = true;
        var installDir = GetInstallDirectory();
        if (!_autoUpdate)
        {
            var decision = MessageBox.Show(
                this,
                "Die App wird nach\\n\\n" + installDir + "\\n\\ninstalliert.\\n"
                + "Die benoetigte Python-Laufzeit ist im Installer enthalten. Falls sie fehlt oder beschaedigt ist, versucht der Installer sie ersatzweise von python.org zu laden.\\n\\nFortfahren?",
                AppName + " Setup",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information
            );

            if (decision != DialogResult.OK)
            {
                _allowClose = true;
                Close();
                return;
            }
        }

        try
        {
            if (_autoUpdate)
            {
                await WaitForPreviousAppAsync();
                CleanupUpdateTempFiles(GetCurrentInstallerPath());
            }

            await InstallAsync(installDir);
            _allowClose = true;

            if (_autoUpdate)
            {
                CleanupUpdateTempFiles(GetCurrentInstallerPath());
                LaunchApp(installDir);
                ScheduleCurrentInstallerDelete();
                Close();
                return;
            }

            var launchNow = MessageBox.Show(
                this,
                "Installation abgeschlossen. Soll " + AppName + " jetzt gestartet werden?",
                AppName + " Setup",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );

            if (launchNow == DialogResult.Yes)
            {
                LaunchApp(installDir);
            }

            Close();
        }
        catch (Exception ex)
        {
            _allowClose = true;
            MessageBox.Show(
                this,
                ex.Message,
                AppName + " Setup",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            Close();
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
            return;
        }

        base.OnFormClosing(e);
    }

    private async Task WaitForPreviousAppAsync()
    {
        if (!_waitPid.HasValue)
        {
            await Task.Delay(1200);
            return;
        }

        try
        {
            var process = Process.GetProcessById(_waitPid.Value);
            SetMarqueeStatus("Alte App wird geschlossen...");
            var started = DateTime.UtcNow;
            while (!process.HasExited && DateTime.UtcNow - started < TimeSpan.FromSeconds(45))
            {
                await Task.Delay(250);
                process.Refresh();
            }
        }
        catch
        {
            System.Threading.Thread.Sleep(600);
        }
    }

    private static string GetCurrentInstallerPath()
    {
        try
        {
            return Application.ExecutablePath;
        }
        catch
        {
            return Assembly.GetExecutingAssembly().Location;
        }
    }

    private static bool IsManagedUpdateTempFile(string fileName)
    {
        var lowerName = (fileName ?? string.Empty).ToLowerInvariant();
        var isSetupName = lowerName == "pdf editor setup.exe"
            || lowerName == "pdf desktop editor setup.exe"
            || lowerName.StartsWith("pdf editor setup ")
            || lowerName.StartsWith("pdf desktop editor setup ");
        return isSetupName && (lowerName.EndsWith(".exe") || lowerName.EndsWith(".exe.download"));
    }

    private static void CleanupUpdateTempFiles(string keepPath)
    {
        var normalizedKeepPath = string.IsNullOrWhiteSpace(keepPath)
            ? string.Empty
            : Path.GetFullPath(keepPath);
        var updateTempDirs = new[]
        {
            Path.Combine(Path.GetTempPath(), AppName),
            Path.Combine(Path.GetTempPath(), "PDF Desktop Editor"),
        };

        foreach (var updateTempDir in updateTempDirs)
        {
            try
            {
                if (!Directory.Exists(updateTempDir))
                {
                    continue;
                }

                foreach (var filePath in Directory.GetFiles(updateTempDir))
                {
                    if (!IsManagedUpdateTempFile(Path.GetFileName(filePath)))
                    {
                        continue;
                    }
                    if (!string.IsNullOrEmpty(normalizedKeepPath)
                        && string.Equals(Path.GetFullPath(filePath), normalizedKeepPath, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    TryDeleteFile(filePath);
                }

                if (Directory.GetFileSystemEntries(updateTempDir).Length == 0)
                {
                    Directory.Delete(updateTempDir);
                }
            }
            catch
            {
                // Update-Temp-Aufraeumen ist best effort.
            }
        }
    }

    private static void TryDeleteFile(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch
        {
            // Best effort.
        }
    }

    private void ScheduleCurrentInstallerDelete()
    {
        var selfPath = GetCurrentInstallerPath();
        if (!_autoUpdate || string.IsNullOrWhiteSpace(selfPath) || !File.Exists(selfPath))
        {
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/C ping 127.0.0.1 -n 3 > nul & del /F /Q \\"" + selfPath + "\\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                WindowStyle = ProcessWindowStyle.Hidden,
            });
        }
        catch
        {
            // Best effort.
        }
    }

    private async Task InstallAsync(string installDir)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "pdf-desktop-editor-setup-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            SetMarqueeStatus("App-Dateien werden entpackt...");
            ExtractPayload(tempDir);

            SetMarqueeStatus("App-Dateien werden installiert...");
            CopyDirectory(tempDir, installDir);

            var backendRoot = Path.Combine(installDir, "dist", "portable-win", "resources", "app", "backend");
            var embeddedPythonRoot = Path.Combine(backendRoot, "python-embed");
            await EnsureEmbeddedPythonAsync(backendRoot, embeddedPythonRoot, tempDir);

            SetMarqueeStatus("Desktop-Verknuepfung wird erstellt...");
            CreateDesktopShortcut(installDir);

            SetMarqueeStatus("Installation abgeschlossen.");
        }
        finally
        {
            try
            {
                Directory.Delete(tempDir, true);
            }
            catch
            {
                // Temp-Aufraeumen ist best effort.
            }
        }
    }

    private async Task EnsureEmbeddedPythonAsync(string backendRoot, string embeddedPythonRoot, string tempDir)
    {
        if (ProbePythonRuntime(embeddedPythonRoot, backendRoot))
        {
            SetMarqueeStatus("Eingebettete Python-Laufzeit ist bereits vorhanden.");
            return;
        }

        Directory.CreateDirectory(embeddedPythonRoot);
        var downloadPath = Path.Combine(tempDir, "python-embed.zip");

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        using (var client = new WebClient())
        {
            client.DownloadProgressChanged += (sender, args) =>
            {
                SetProgressStatus(
                    "Python-Laufzeit wird heruntergeladen...",
                    Math.Max(0, Math.Min(100, args.ProgressPercentage))
                );
            };

            await client.DownloadFileTaskAsync(new Uri(PythonEmbedUrl), downloadPath);
        }

        SetMarqueeStatus("Python-Laufzeit wird eingerichtet...");
        if (Directory.Exists(embeddedPythonRoot))
        {
            Directory.Delete(embeddedPythonRoot, true);
        }
        Directory.CreateDirectory(embeddedPythonRoot);
        ZipFile.ExtractToDirectory(downloadPath, embeddedPythonRoot);
        ConfigureEmbeddedPython(embeddedPythonRoot);

        if (!ProbePythonRuntime(embeddedPythonRoot, backendRoot))
        {
            throw new InvalidOperationException(
                "Die eingebettete Python-Laufzeit konnte eingerichtet werden, aber der Backend-Dienst startet noch nicht.\\n"
                + "Bitte pruefen, ob der Download von python.org vollstaendig war oder ob Windows das Starten blockiert."
            );
        }
    }

    private void ConfigureEmbeddedPython(string embeddedPythonRoot)
    {
        var pthFiles = Directory.GetFiles(embeddedPythonRoot, "python*._pth");
        if (pthFiles.Length == 0)
        {
            throw new FileNotFoundException("Die Python-Embed-Konfiguration wurde nicht gefunden.");
        }

        var pthContent = string.Join(
            Environment.NewLine,
            new[]
            {
                "python312.zip",
                ".",
                "..",
                "..\\\\.venv\\\\Lib\\\\site-packages",
                "import site",
                "",
            }
        );
        File.WriteAllText(pthFiles[0], pthContent);
    }

    private bool ProbePythonRuntime(string embeddedPythonRoot, string backendRoot)
    {
        var pythonExe = Path.Combine(embeddedPythonRoot, "python.exe");
        if (!File.Exists(pythonExe))
        {
            return false;
        }

        var probe = new ProcessStartInfo
        {
            FileName = pythonExe,
            Arguments = "-c \\"import fastapi, uvicorn, pymupdf, pydantic; print('ok')\\"",
            WorkingDirectory = backendRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        using (var process = Process.Start(probe))
        {
            if (process == null)
            {
                return false;
            }

            if (!process.WaitForExit(15000))
            {
                try
                {
                    process.Kill();
                }
                catch
                {
                    // ignore
                }
                return false;
            }

            return process.ExitCode == 0;
        }
    }

    private void ExtractPayload(string targetDirectory)
    {
        using (var payloadStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(PayloadResourceName))
        {
            if (payloadStream == null)
            {
                throw new InvalidOperationException("Das Installer-Payload konnte nicht geladen werden.");
            }

            using (var archive = new ZipArchive(payloadStream, ZipArchiveMode.Read))
            {
                archive.ExtractToDirectory(targetDirectory);
            }
        }
    }

    private static void CopyDirectory(string sourceDir, string targetDir)
    {
        Directory.CreateDirectory(targetDir);

        foreach (var directory in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
        {
            var relative = directory.Substring(sourceDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            Directory.CreateDirectory(Path.Combine(targetDir, relative));
        }

        foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
        {
            var relative = file.Substring(sourceDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var targetPath = Path.Combine(targetDir, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(targetPath) ?? targetDir);
            CopyFileWithRetry(file, targetPath);
        }
    }

    private static void CopyFileWithRetry(string sourcePath, string targetPath)
    {
        Exception lastError = null;
        for (var attempt = 0; attempt < 80; attempt++)
        {
            try
            {
                File.Copy(sourcePath, targetPath, true);
                return;
            }
            catch (IOException ex)
            {
                lastError = ex;
                System.Threading.Thread.Sleep(250);
            }
            catch (UnauthorizedAccessException ex)
            {
                lastError = ex;
                System.Threading.Thread.Sleep(250);
            }
        }

        throw new IOException("Eine App-Datei konnte nicht ersetzt werden: " + targetPath, lastError);
    }

    private static string GetInstallDirectory()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs",
            AppName
        );
    }

    private static void LaunchApp(string installDir)
    {
        var launcherPath = Path.Combine(installDir, AppExecutableName);
        if (!File.Exists(launcherPath))
        {
            throw new FileNotFoundException("Die installierte App wurde nicht gefunden.", launcherPath);
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = launcherPath,
            WorkingDirectory = installDir,
            UseShellExecute = true,
        });
    }

    private static void CreateDesktopShortcut(string installDir)
    {
        var launcherPath = Path.Combine(installDir, AppExecutableName);
        if (!File.Exists(launcherPath))
        {
            throw new FileNotFoundException("Die installierte App wurde nicht gefunden.", launcherPath);
        }

        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var shortcutPath = Path.Combine(desktopPath, AppName + ".lnk");
        var shellType = Type.GetTypeFromProgID("WScript.Shell");
        if (shellType == null)
        {
            throw new InvalidOperationException("Die Desktop-Verknuepfung konnte nicht erstellt werden.");
        }

        object shell = null;
        object shortcut = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            shortcut = shellType.InvokeMember(
                "CreateShortcut",
                BindingFlags.InvokeMethod,
                null,
                shell,
                new object[] { shortcutPath }
            );

            var shortcutType = shortcut.GetType();
            shortcutType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, new object[] { launcherPath });
            shortcutType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortcut, new object[] { installDir });
            shortcutType.InvokeMember("Description", BindingFlags.SetProperty, null, shortcut, new object[] { AppName });
            shortcutType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, new object[] { launcherPath + ",0" });
            shortcutType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
        }
        finally
        {
            if (shortcut != null && Marshal.IsComObject(shortcut))
            {
                Marshal.ReleaseComObject(shortcut);
            }
            if (shell != null && Marshal.IsComObject(shell))
            {
                Marshal.ReleaseComObject(shell);
            }
        }
    }

    private void SetMarqueeStatus(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(SetMarqueeStatus), message);
            return;
        }

        _statusLabel.Text = message;
        _progressBar.Style = ProgressBarStyle.Marquee;
    }

    private void SetProgressStatus(string message, int percentage)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string, int>(SetProgressStatus), message, percentage);
            return;
        }

        _statusLabel.Text = message + " " + percentage + "%";
        if (_progressBar.Style != ProgressBarStyle.Continuous)
        {
            _progressBar.Style = ProgressBarStyle.Continuous;
        }
        _progressBar.Value = Math.Max(0, Math.Min(100, percentage));
    }
}
`;

  fs.writeFileSync(sourcePath, installerSource, "utf8");
  return sourcePath;
}

function buildInstallerExe(sourcePath) {
  const compilerPath = findCSharpCompiler();
  if (!compilerPath) {
    throw new Error("Kein C#-Compiler gefunden. Der Setup-Installer konnte nicht gebaut werden.");
  }

  const iconPath = path.join(projectRoot, "app", "assets", "pdf-editor-old.ico");
  run(
    compilerPath,
    [
      "/nologo",
      "/target:winexe",
      "/reference:System.dll",
      "/reference:System.Windows.Forms.dll",
      "/reference:System.Drawing.dll",
      "/reference:System.Net.Http.dll",
      "/reference:System.IO.Compression.dll",
      "/reference:System.IO.Compression.FileSystem.dll",
      `/win32icon:${iconPath}`,
      `/resource:${installerPayloadZip},${installerResourceName}`,
      `/out:${installerOutputPath}`,
      sourcePath,
    ],
  );
}

function refreshLaptopCopyFolder() {
  fs.mkdirSync(laptopCopyRoot, { recursive: true });
  copyRecursive(installerOutputPath, laptopCopyInstallerPath);
  fs.writeFileSync(
    laptopCopyInstructionsPath,
    [
      `${appDisplayName} auf einem anderen Laptop installieren`,
      "",
      "1. Diese Datei auf den Laptop kopieren:",
      `   ${installerFileName}`,
      "",
      "2. Auf dem Laptop doppelklicken und die Installation bestaetigen.",
      "   Der Installer legt die App hier ab:",
      `   %LOCALAPPDATA%\\Programs\\${appDisplayName}`,
      "",
      "3. Danach erscheint auf dem Desktop eine Verknuepfung:",
      `   ${appDisplayName}`,
      "",
      "4. Der Laptop braucht normalerweise kein Internet mehr fuer die Installation.",
      "   Die offizielle eingebettete Python-Laufzeit ist im Installer enthalten.",
      "   Nur falls diese Laufzeit fehlt oder beschaedigt ist, versucht der Installer",
      "   sie ersatzweise von python.org herunterzuladen.",
      "",
      "5. Danach kann die App ohne Codex, ohne Node.js und ohne manuell",
      "   installiertes Python gestartet werden.",
      "",
      "6. Falls Windows SmartScreen warnt:",
      "   \"Weitere Informationen\" anklicken und dann \"Trotzdem ausfuehren\".",
      "",
      "7. Wichtig:",
      "   Nicht den Projektordner zippen. Nur die Setup-EXE weitergeben.",
      "",
    ].join("\r\n"),
    "utf8",
  );

  fs.rmSync(laptopCopyZip, { force: true });
  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$source = '${laptopCopyRoot.replace(/'/g, "''")}'`,
    `$zip = '${laptopCopyZip.replace(/'/g, "''")}'`,
    "if (Test-Path $zip) { Remove-Item $zip -Force }",
    "Compress-Archive -Path (Join-Path $source '*') -DestinationPath $zip -CompressionLevel Optimal",
  ].join("; ");
  run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]);
}

function buildInstaller() {
  preparePortableBundle();
  ensurePortableEmbeddedPython();
  stageInstallerPayload();
  buildPayloadZip();
  const sourcePath = buildInstallerSource();
  buildInstallerExe(sourcePath);
  refreshLaptopCopyFolder();
  console.log(`Setup-Installer bereit: ${installerOutputPath}`);
  console.log(`Laptop-Kopierordner bereit: ${laptopCopyRoot}`);
}

buildInstaller();
