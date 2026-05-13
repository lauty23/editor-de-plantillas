using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;

const string AppName = "Editor de Plantillas";
const string LegacyProcessName = "WinUI3TemplateEditor";
const string AppProcessName = "EditorDePlantillas";
const string BadCopiedProcessName = "Editor de Plantillas";
const string PayloadExeName = "EditorDePlantillas.exe";

try
{
    var silent = args.Any(arg =>
        arg.Equals("--silent", StringComparison.OrdinalIgnoreCase)
        || arg.Equals("/silent", StringComparison.OrdinalIgnoreCase)
        || arg.Equals("-silent", StringComparison.OrdinalIgnoreCase));
    var installDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WinUI3TemplateEditor");
    var localAppData = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
    var resolvedTarget = Path.GetFullPath(installDir);
    if (!resolvedTarget.StartsWith(localAppData, StringComparison.OrdinalIgnoreCase))
    {
        throw new InvalidOperationException($"Ruta de instalacion no valida: {resolvedTarget}");
    }

    foreach (var processName in new[] { LegacyProcessName, AppProcessName, BadCopiedProcessName })
    {
        foreach (var process in Process.GetProcessesByName(processName))
        {
            try
            {
                process.Kill();
                process.WaitForExit(3000);
            }
            catch
            {
                // Best effort: continuing lets the installer report any locked file cleanly.
            }
        }
    }

    if (Directory.Exists(installDir))
    {
        Directory.Delete(installDir, recursive: true);
    }
    Directory.CreateDirectory(installDir);

    var tempZip = Path.Combine(Path.GetTempPath(), $"WinUI3TemplateEditor_{Guid.NewGuid():N}.zip");
    try
    {
        ExtractEmbeddedZip(tempZip);
        ZipFile.ExtractToDirectory(tempZip, installDir, overwriteFiles: true);
    }
    finally
    {
        if (File.Exists(tempZip))
        {
            File.Delete(tempZip);
        }
    }

    var payloadExePath = Path.Combine(installDir, PayloadExeName);
    if (!File.Exists(payloadExePath))
    {
        throw new FileNotFoundException("No se encontro el ejecutable instalado.", payloadExePath);
    }

    var desktopShortcut = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), $"{AppName}.lnk");
    var startShortcut = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs", $"{AppName}.lnk");
    CreateShortcut(desktopShortcut, payloadExePath, installDir);
    CreateShortcut(startShortcut, payloadExePath, installDir);

    if (!silent)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = payloadExePath,
            WorkingDirectory = installDir,
            UseShellExecute = true
        });

        MessageBox(
            IntPtr.Zero,
            $"Editor de Plantillas instalado correctamente.\n\nEjecutable:\n{payloadExePath}\n\nAcceso directo:\n{desktopShortcut}",
            AppName,
            0x40);
    }
}
catch (Exception ex)
{
    Environment.ExitCode = 1;
    MessageBox(IntPtr.Zero, $"No se pudo instalar la aplicacion:\n\n{ex.Message}", AppName, 0x10);
}

static void ExtractEmbeddedZip(string destination)
{
    var assembly = Assembly.GetExecutingAssembly();
    var resourceName = assembly.GetManifestResourceNames()
        .FirstOrDefault(name => name.EndsWith("WinUI3TemplateEditor.zip", StringComparison.OrdinalIgnoreCase));
    if (resourceName is null)
    {
        throw new InvalidOperationException("No se encontro el ZIP embebido en el instalador.");
    }

    using var resource = assembly.GetManifestResourceStream(resourceName)
        ?? throw new InvalidOperationException("No se pudo abrir el ZIP embebido.");
    using var output = File.Create(destination);
    resource.CopyTo(output);
}

static void CreateShortcut(string shortcutPath, string targetPath, string workingDirectory)
{
    Directory.CreateDirectory(Path.GetDirectoryName(shortcutPath)!);
    var shellType = Type.GetTypeFromProgID("WScript.Shell")
        ?? throw new InvalidOperationException("No se pudo abrir WScript.Shell para crear accesos directos.");
    dynamic shell = Activator.CreateInstance(shellType)!;
    dynamic shortcut = shell.CreateShortcut(shortcutPath);
    shortcut.TargetPath = targetPath;
    shortcut.WorkingDirectory = workingDirectory;
    shortcut.IconLocation = $"{targetPath},0";
    shortcut.Save();
}

[DllImport("user32.dll", CharSet = CharSet.Unicode)]
static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
