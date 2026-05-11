$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$dotnet = Join-Path $projectRoot ".dotnet\dotnet.exe"
$project = Join-Path $projectRoot "WinUI3TemplateEditor\WinUI3TemplateEditor.csproj"
$output = Join-Path $projectRoot "dist\WinUI3TemplateEditor"

if (-not (Test-Path $dotnet)) {
    $installer = Join-Path $projectRoot "dotnet-install.ps1"
    if (-not (Test-Path $installer)) {
        Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile $installer
    }
    powershell -ExecutionPolicy Bypass -File $installer -Channel 8.0 -InstallDir (Join-Path $projectRoot ".dotnet")
}

if (Test-Path $output) {
    Remove-Item $output -Recurse -Force
}

$versionArgs = @()
if (-not [string]::IsNullOrWhiteSpace($env:APP_VERSION)) {
    $versionArgs += "-p:Version=$env:APP_VERSION"
}

& $dotnet publish $project `
    -c Release `
    -p:Platform=x64 `
    -p:RuntimeIdentifier=win-x64 `
    -o $output `
    @versionArgs

Write-Host ""
Write-Host "Listo:"
Write-Host (Join-Path $output "WinUI3TemplateEditor.exe")
