Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

function New-UnicodeString {
  param([int[]]$CodePoints)
  return (-join ($CodePoints | ForEach-Object { [char]$_ }))
}

$appName = "OpenClaw " + (New-UnicodeString @(0x63A7, 0x5236, 0x4E2D, 0x5FC3))
$installRoot = Join-Path $env:LOCALAPPDATA "OpenClawControlCenter"
$installExe = Join-Path $installRoot ($appName + ".exe")
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\OpenClaw"
$startMenuShortcut = Join-Path $startMenuDir ($appName + ".lnk")
$versionFile = Join-Path $installRoot "version.txt"
$embeddedAppBase64 = @'
__EMBEDDED_APP_BASE64__
'@

function New-Shortcut {
  param(
    [string]$ShortcutPath,
    [string]$TargetPath,
    [string]$WorkingDirectory
  )

  $shortcutDir = Split-Path -Parent $ShortcutPath
  New-Item -ItemType Directory -Path $shortcutDir -Force | Out-Null

  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($ShortcutPath)
  $shortcut.TargetPath = $TargetPath
  $shortcut.WorkingDirectory = $WorkingDirectory
  $shortcut.IconLocation = $TargetPath
  $shortcut.Save()
}

function Get-DesktopShortcutPaths {
  param([string]$ShortcutName)

  $desktopDirs = @(
    [Environment]::GetFolderPath("Desktop"),
    [Environment]::GetFolderPath("DesktopDirectory"),
    (Join-Path $env:USERPROFILE "Desktop"),
    (Join-Path $env:PUBLIC "Desktop")
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

  $uniqueDirs = New-Object System.Collections.Generic.List[string]
  foreach ($desktopDir in $desktopDirs) {
    try {
      $fullDir = [System.IO.Path]::GetFullPath($desktopDir)
      if (-not $uniqueDirs.Contains($fullDir)) {
        $uniqueDirs.Add($fullDir)
      }
    } catch {
    }
  }

  foreach ($desktopDir in $uniqueDirs) {
    Join-Path $desktopDir $ShortcutName
  }
}

function Ensure-DesktopShortcut {
  param(
    [string]$ShortcutName,
    [string]$TargetPath,
    [string]$WorkingDirectory
  )

  $errors = New-Object System.Collections.Generic.List[string]
  foreach ($shortcutPath in (Get-DesktopShortcutPaths -ShortcutName $ShortcutName)) {
    try {
      New-Shortcut -ShortcutPath $shortcutPath -TargetPath $TargetPath -WorkingDirectory $WorkingDirectory
      return $shortcutPath
    } catch {
      $errors.Add($_.Exception.Message)
    }
  }

  if ($errors.Count -gt 0) {
    throw ("Failed to create desktop shortcut: " + ($errors -join " | "))
  }

  throw "No valid desktop directory found."
}

function Stop-InstalledProcess {
  param([string]$ExecutablePath)

  Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.ExecutablePath -eq $ExecutablePath } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}

New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
Stop-InstalledProcess -ExecutablePath $installExe

$bytes = [Convert]::FromBase64String(($embeddedAppBase64 -replace '\s+', ''))
[System.IO.File]::WriteAllBytes($installExe, $bytes)
Set-Content -Path $versionFile -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Encoding UTF8

$desktopShortcutPath = Ensure-DesktopShortcut -ShortcutName ($appName + ".lnk") -TargetPath $installExe -WorkingDirectory $installRoot
New-Shortcut -ShortcutPath $startMenuShortcut -TargetPath $installExe -WorkingDirectory $installRoot

$installDoneText = New-UnicodeString @(0x5B89, 0x88C5, 0x5B8C, 0x6210, 0x3002)
$locationLabel = New-UnicodeString @(0x7A0B, 0x5E8F, 0x4F4D, 0x7F6E, 0xFF1A)
$desktopCreatedText = New-UnicodeString @(0x5DF2, 0x5728, 0x684C, 0x9762, 0x521B, 0x5EFA, 0x5FEB, 0x6377, 0x65B9, 0x5F0F, 0x3002)
$shortcutLabel = New-UnicodeString @(0x5FEB, 0x6377, 0x65B9, 0x5F0F, 0xFF1A)

[System.Windows.Forms.MessageBox]::Show(
  ($installDoneText + "`r`n" + $desktopCreatedText + "`r`n`r`n" + $locationLabel + $installExe + "`r`n`r`n" + $shortcutLabel + $desktopShortcutPath),
  $appName,
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

Start-Process -FilePath $installExe -WorkingDirectory $installRoot
