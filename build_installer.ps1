$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$dotnet = Join-Path $projectRoot ".dotnet\dotnet.exe"
$appZip = Join-Path $projectRoot "dist\WinUI3TemplateEditor.zip"
$payloadZip = Join-Path $projectRoot "installer\payload\WinUI3TemplateEditor.zip"
$installerProject = Join-Path $projectRoot "InstallerBuilder\WinUI3TemplateEditorInstaller.csproj"
$publishDir = Join-Path $projectRoot "dist\InstallerSingle"
$setupExe = Join-Path $projectRoot "dist\EditorDePlantillasSetup.exe"

Get-Process | Where-Object { $_.ProcessName -like "WinUI3TemplateEditor*" } | Stop-Process -Force -ErrorAction SilentlyContinue

powershell -ExecutionPolicy Bypass -File (Join-Path $projectRoot "build_winui3.ps1")

if (Test-Path -LiteralPath $appZip) {
    Remove-Item -LiteralPath $appZip -Force
}
Compress-Archive -Path (Join-Path $projectRoot "dist\WinUI3TemplateEditor\*") -DestinationPath $appZip

New-Item -ItemType Directory -Path (Split-Path -Parent $payloadZip) -Force | Out-Null
Copy-Item -LiteralPath $appZip -Destination $payloadZip -Force

if (Test-Path -LiteralPath $publishDir) {
    Remove-Item -LiteralPath $publishDir -Recurse -Force
}

& $dotnet publish $installerProject `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:EnableCompressionInSingleFile=true `
    -o $publishDir

Copy-Item -LiteralPath (Join-Path $publishDir "WinUI3TemplateEditorInstaller.exe") -Destination $setupExe -Force

Write-Host ""
Write-Host "Instalador listo:"
Write-Host $setupExe
