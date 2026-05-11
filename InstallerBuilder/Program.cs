using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;

const string AppName = "Editor de Plantillas";
const string ProcessName = "WinUI3TemplateEditor";
const string ExeName = "WinUI3TemplateEditor.exe";

try
{
    var installDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WinUI3TemplateEditor");
    var localAppData = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
    var resolvedTarget = Path.GetFullPath(installDir);
    if (!resolvedTarget.StartsWith(localAppData, StringComparison.OrdinalIgnoreCase))
    {
        throw new InvalidOperationException($"Ruta de instalacion no valida: {resolvedTarget}");
    }

    foreach (var process in Process.GetProcessesByName(ProcessName))
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

    var exePath = Path.Combine(installDir, ExeName);
    CreateShortcut(
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), $"{AppName}.lnk"),
        exePath,
        installDir);
    CreateShortcut(
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs", $"{AppName}.lnk"),
        exePath,
        installDir);

    Process.Start(new ProcessStartInfo
    {
        FileName = exePath,
        WorkingDirectory = installDir,
        UseShellExecute = true
    });

    MessageBox(IntPtr.Zero, "Editor de Plantillas instalado correctamente.", AppName, 0x40);
}
catch (Exception ex)
{
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
