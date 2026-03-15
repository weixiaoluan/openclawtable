$ErrorActionPreference = "Stop"

$meta = @'
{
  "exe_name": "OpenClaw \u63a7\u5236\u4e2d\u5fc3",
  "title": "OpenClaw \u63a7\u5236\u4e2d\u5fc3",
  "description": "OpenClaw \u9759\u9ed8\u8fd0\u884c\u3001\u5907\u4efd\u6062\u590d\u4e0e\u66f4\u65b0\u63a7\u5236\u9762\u677f"
}
'@ | ConvertFrom-Json

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $root
$assetsDir = Join-Path $root "assets"
$distDir = Join-Path $projectRoot "dist"
$runtimeDir = Join-Path $projectRoot "runtime"
$scriptPath = Join-Path $root "OpenClaw-Control-Center.ps1"
$generatedDir = Join-Path $root ".generated"
$generatedScriptPath = Join-Path $generatedDir "OpenClaw-Control-Center.generated.ps1"
$repoToolkitPath = Join-Path $runtimeDir "openclaw-toolkit.ps1"
$userToolkitPath = Join-Path ([Environment]::GetFolderPath("UserProfile")) ".openclaw\openclaw-toolkit.ps1"
$toolkitPath = if (Test-Path $repoToolkitPath) { $repoToolkitPath } else { $userToolkitPath }
$iconPath = Join-Path $assetsDir "lobster-icon.ico"
$outputExe = Join-Path $distDir ($meta.exe_name + ".exe")
$desktopRoot = [Environment]::GetFolderPath("Desktop")
$desktopExe = Join-Path $desktopRoot ($meta.exe_name + ".exe")
$userDesktopExe = Join-Path ([Environment]::GetFolderPath("UserProfile")) ("Desktop\" + $meta.exe_name + ".exe")
$legacyOutputExe = Join-Path $distDir "OpenClaw Control Center.exe"
$legacyDesktopExe = Join-Path $desktopRoot "OpenClaw Control Center.exe"
$ps2exeScript = Join-Path $projectRoot ".tools\ps2exe\pkg\ps2exe.ps1"
$pythonExe = "python"

if (-not (Test-Path $ps2exeScript)) {
  throw "Missing PS2EXE script: $ps2exeScript"
}

if (-not (Test-Path $toolkitPath)) {
  throw "Missing toolkit script: $toolkitPath"
}

New-Item -ItemType Directory -Force -Path $assetsDir | Out-Null
New-Item -ItemType Directory -Force -Path $distDir | Out-Null
New-Item -ItemType Directory -Force -Path $generatedDir | Out-Null
New-Item -ItemType Directory -Force -Path $runtimeDir | Out-Null

$knownExePaths = @($outputExe, $desktopExe, $userDesktopExe, $legacyOutputExe, $legacyDesktopExe) |
  Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
  Select-Object -Unique

Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
  Where-Object { $_.ExecutablePath -and ($knownExePaths -contains $_.ExecutablePath) } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

& $pythonExe (Join-Path $root "generate_lobster_assets.py")

$panelTemplate = Get-Content $scriptPath -Raw
$toolkitBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-Content $toolkitPath -Raw)))
$generatedPanel = $panelTemplate.Replace("__EMBEDDED_TOOLKIT_BASE64__", $toolkitBase64)
Set-Content -Path $generatedScriptPath -Value $generatedPanel -Encoding UTF8

. $ps2exeScript
Invoke-ps2exe -inputFile $generatedScriptPath `
  -outputFile $outputExe `
  -iconFile $iconPath `
  -noConsole `
  -STA `
  -DPIAware `
  -title $meta.title `
  -description $meta.description `
  -company "WEIZHEN Local Tools" `
  -product $meta.title `
  -copyright "Copyright (c) 2026" `
  -version "2026.3.15.1"

Copy-Item $outputExe $desktopExe -Force
if ($userDesktopExe -ne $desktopExe) {
  $userDesktopDir = Split-Path -Parent $userDesktopExe
  if (Test-Path $userDesktopDir) {
    Copy-Item $outputExe $userDesktopExe -Force
  }
}

foreach ($oldFile in @($legacyOutputExe, $legacyDesktopExe)) {
  if ((Test-Path $oldFile) -and ($oldFile -ne $outputExe) -and ($oldFile -ne $desktopExe)) {
    Remove-Item $oldFile -Force -ErrorAction SilentlyContinue
  }
}

Write-Output $outputExe
