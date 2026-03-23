$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $root
$distDir = Join-Path $projectRoot "dist"
$generatedDir = Join-Path $root ".generated"
$installerTemplatePath = Join-Path $root "install-control-center.ps1"
$generatedInstallerPath = Join-Path $generatedDir "install-control-center.generated.ps1"
$buildControlCenterPath = Join-Path $root "build-control-center.ps1"
$iconPath = Join-Path $root "assets/lobster-icon.ico"
$ps2exeScript = Join-Path $projectRoot ".tools/ps2exe/pkg/ps2exe.ps1"
$installerExePath = Join-Path $distDir "OpenClaw-Control-Center-Installer.exe"
$desktopInstallerPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "OpenClaw-Control-Center-Installer.exe"

function Copy-BestEffort {
  param(
    [string]$SourcePath,
    [string]$DestinationPath
  )

  try {
    Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
  } catch {
    Write-Warning ("Failed to copy {0} to {1}: {2}" -f $SourcePath, $DestinationPath, $_.Exception.Message)
  }
}

if (-not (Test-Path $buildControlCenterPath)) {
  throw "Missing control center build script: $buildControlCenterPath"
}

if (-not (Test-Path $ps2exeScript)) {
  throw "Missing PS2EXE script: $ps2exeScript"
}

New-Item -ItemType Directory -Path $generatedDir -Force | Out-Null
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

& $buildControlCenterPath

$appExe = Get-ChildItem -Path $distDir -Filter "*.exe" |
  Where-Object { $_.Name -like "OpenClaw*" -and $_.Name -notlike "*Installer*" } |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1

if ($null -eq $appExe) {
  throw "Missing app EXE for installer build in $distDir"
}

$template = Get-Content -Path $installerTemplatePath -Raw
$appBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($appExe.FullName))
$generated = $template.Replace("__EMBEDDED_APP_BASE64__", $appBase64)
Set-Content -Path $generatedInstallerPath -Value $generated -Encoding UTF8

. $ps2exeScript
Invoke-ps2exe -inputFile $generatedInstallerPath `
  -outputFile $installerExePath `
  -iconFile $iconPath `
  -noConsole `
  -STA `
  -DPIAware `
  -title "OpenClaw Control Center Installer" `
  -description "Install OpenClaw Control Center for the current user" `
  -company "WEIZHEN Local Tools" `
  -product "OpenClaw Control Center Installer" `
  -copyright "Copyright (c) 2026" `
  -version "2026.3.23.1"

Copy-BestEffort -SourcePath $installerExePath -DestinationPath $desktopInstallerPath
Write-Output $installerExePath
