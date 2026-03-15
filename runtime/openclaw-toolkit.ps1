Set-StrictMode -Version 3.0

try { Add-Type -AssemblyName System.IO.Compression } catch {}
try { Add-Type -AssemblyName System.IO.Compression.FileSystem } catch {}

$script:OpenClawToolkitSelfPath = $PSCommandPath
$script:OpenClawToolkitSelfContent = if ($PSCommandPath -and (Test-Path $PSCommandPath)) { Get-Content $PSCommandPath -Raw } else { $null }

function Get-OpenClawUserHome {
  return [Environment]::GetFolderPath("UserProfile")
}

function Get-OpenClawDefaultPrefix {
  return (Join-Path (Get-OpenClawUserHome) ".npm-global-user")
}

function Get-OpenClawDefaultCachePath {
  return (Join-Path $env:LOCALAPPDATA "npm-cache")
}

function Get-OpenClawNodeExe {
  try {
    return (Get-Command node.exe -ErrorAction Stop).Source
  } catch {
    $fallback = Join-Path ${env:ProgramFiles} "nodejs\node.exe"
    if (Test-Path $fallback) {
      return $fallback
    }
  }

  return $null
}

function Get-OpenClawNpmCli {
  $nodeExe = Get-OpenClawNodeExe
  if (-not $nodeExe) {
    return $null
  }

  $npmCli = Join-Path (Split-Path -Parent $nodeExe) "node_modules\npm\bin\npm-cli.js"
  if (Test-Path $npmCli) {
    return $npmCli
  }

  return $null
}

function Get-OpenClawContext {
  $userHome = Get-OpenClawUserHome
  $openClawHome = Join-Path $userHome ".openclaw"
  $tempLogDir = Join-Path $env:TEMP "openclaw"
  $defaultPrefix = Get-OpenClawDefaultPrefix
  $systemDrive = if ($env:SystemDrive) { $env:SystemDrive } else { "C:" }

  [pscustomobject]@{
    UserHome = $userHome
    Home = $openClawHome
    PauseFile = Join-Path $openClawHome ".paused"
    UpgradeLockFile = Join-Path $openClawHome ".upgrade-lock"
    LauncherVbs = Join-Path $openClawHome "run-hidden.vbs"
    GatewayCmd = Join-Path $openClawHome "gateway.cmd"
    WatchdogScript = Join-Path $openClawHome "gateway-watchdog.ps1"
    SupervisorScript = Join-Path $openClawHome "gateway-supervisor.ps1"
    StartupLauncher = Join-Path ([Environment]::GetFolderPath("Startup")) "OpenClaw Gateway.vbs"
    UpgradeScript = Join-Path $openClawHome "upgrade-openclaw.ps1"
    ToolkitScript = Join-Path $openClawHome "openclaw-toolkit.ps1"
    TempLogDir = $tempLogDir
    WatchdogLog = Join-Path $tempLogDir "gateway-watchdog.log"
    SupervisorLog = Join-Path $tempLogDir "gateway-supervisor.log"
    UpgradeLog = Join-Path $tempLogDir "upgrade-openclaw.log"
    UpgradeStdoutLog = Join-Path $tempLogDir "upgrade-npm.stdout.log"
    UpgradeStderrLog = Join-Path $tempLogDir "upgrade-npm.stderr.log"
    GatewayLogDir = Join-Path $systemDrive "tmp\openclaw"
    AuditRoot = Join-Path $openClawHome "logs\discord-audit"
    BackupRoot = Join-Path $openClawHome "backups"
    NpmrcPath = Join-Path $userHome ".npmrc"
    DefaultPrefix = $defaultPrefix
    KeepaliveTask = "OpenClaw Gateway Keepalive"
    GatewayTask = "OpenClaw Gateway"
    HealthzUrl = "http://127.0.0.1:18789/healthz"
    Desktop = [Environment]::GetFolderPath("Desktop")
    DefaultCachePath = Get-OpenClawDefaultCachePath
  }
}

function Write-OpenClawTextFile {
  param(
    [string]$Path,
    [string]$Content,
    [string]$Encoding = "ASCII"
  )

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }

  $existing = $null
  if (Test-Path $Path) {
    try {
      $existing = Get-Content $Path -Raw -ErrorAction Stop
    } catch {
      $existing = $null
    }
  }

  if ($existing -ceq $Content) {
    return
  }

  Set-Content -Path $Path -Value $Content -Encoding $Encoding
}

function Get-OpenClawRunHiddenVbsContent {
@'
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
pauseFile = fso.BuildPath(baseDir, ".paused")
upgradeLockFile = fso.BuildPath(baseDir, ".upgrade-lock")
watchdogScript = fso.BuildPath(baseDir, "gateway-watchdog.ps1")

If fso.FileExists(pauseFile) Then
  WScript.Quit 0
End If
If fso.FileExists(upgradeLockFile) Then
  WScript.Quit 0
End If
cmd = "powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File """ & watchdogScript & """"
shell.Run cmd, 0, False
'@
}

function Get-OpenClawGatewayCmdContent {
@'
@echo off
setlocal
rem OpenClaw Gateway (silent launcher)
set "BASE_DIR=%~dp0"
set "PAUSE_FILE=%BASE_DIR%.paused"
set "UPGRADE_LOCK=%BASE_DIR%.upgrade-lock"
set "LAUNCHER=%BASE_DIR%run-hidden.vbs"

if exist "%PAUSE_FILE%" exit /b 0
if exist "%UPGRADE_LOCK%" exit /b 0
if not exist "%LAUNCHER%" exit /b 1

start "" /b wscript.exe "%LAUNCHER%"
exit /b 0
'@
}

function Get-OpenClawGatewayWatchdogContent {
@'
$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $baseDir
$pauseFile = Join-Path $baseDir ".paused"
$upgradeLockFile = Join-Path $baseDir ".upgrade-lock"

if (Test-Path $pauseFile) {
  exit 0
}

if (Test-Path $upgradeLockFile) {
  exit 0
}

$logDir = Join-Path $env:TEMP "openclaw"
$logFile = Join-Path $logDir "gateway-watchdog.log"
$maxLogBytes = 1MB
$supervisorScript = Join-Path $baseDir "gateway-supervisor.ps1"
$port = 18789
$healthzUrl = "http://127.0.0.1:$port/healthz"

if (-not (Test-Path $logDir)) {
  New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Rotate-Log {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    return
  }

  $item = Get-Item $Path
  if ($item.Length -lt $maxLogBytes) {
    return
  }

  $archive = "$Path.1"
  if (Test-Path $archive) {
    Remove-Item $archive -Force -ErrorAction SilentlyContinue
  }
  Move-Item -Path $Path -Destination $archive -Force
}

function Write-Log {
  param([string]$Message)
  Rotate-Log -Path $logFile
  $line = "[{0}] {1}" -f (Get-Date -Format "yyyy/MM/dd ddd HH:mm:ss.ff"), $Message
  Add-Content -Path $logFile -Value $line
}

function Get-SupervisorProcesses {
  Get-CimInstance Win32_Process |
    Where-Object {
      $_.Name -like "powershell*" -and
      $_.CommandLine -like "*gateway-supervisor.ps1*"
    }
}

function Get-GatewayProcesses {
  Get-CimInstance Win32_Process |
    Where-Object {
      $_.Name -eq "node.exe" -and
      (
        $_.CommandLine -like "*node_modules\openclaw\dist\index.js*gateway*" -or
        $_.CommandLine -like "*node_modules\openclaw\openclaw.mjs*gateway*"
      )
    }
}

function Test-GatewayProbe {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $healthzUrl -TimeoutSec 3
    return ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300)
  } catch {
    return $false
  }
}

function Start-Supervisor {
  Start-Process -FilePath "powershell.exe" -ArgumentList @("-WindowStyle", "Hidden", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $supervisorScript) -WorkingDirectory $baseDir -WindowStyle Hidden | Out-Null
}

$mutex = New-Object System.Threading.Mutex($false, "Global\OpenClawGatewayWatchdog")
$hasMutex = $false

try {
  $hasMutex = $mutex.WaitOne(0, $false)
  if (-not $hasMutex) {
    exit 0
  }

  $supervisors = @(Get-SupervisorProcesses)
  if ($supervisors.Count -gt 1) {
    Write-Log "found $($supervisors.Count) supervisors; keeping newest and stopping duplicates"
    $ordered = $supervisors | Sort-Object CreationDate -Descending
    foreach ($proc in $ordered | Select-Object -Skip 1) {
      Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue
    }
    $supervisors = @($ordered | Select-Object -First 1)
  }

  $healthy = Test-GatewayProbe
  if ($healthy) {
    if ($supervisors.Count -eq 0) {
      Write-Log "probe healthy but supervisor missing; starting supervisor to adopt existing gateway"
      Start-Supervisor
      exit 0
    }
    Write-Log "probe healthy; no action needed"
    exit 0
  }

  if ($supervisors.Count -eq 0) {
    Write-Log "supervisor missing and probe unhealthy; starting supervisor"
    Start-Supervisor
    exit 0
  }

  $gatewayChildren = @(Get-GatewayProcesses)
  if ($gatewayChildren.Count -eq 0) {
    Write-Log "supervisor present but gateway child missing; recycling supervisor"
    Stop-Process -Id $supervisors[0].ProcessId -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Supervisor
    exit 0
  }

  Write-Log "supervisor present and gateway child exists; leaving recovery to supervisor"
}
finally {
  if ($hasMutex) {
    [void]$mutex.ReleaseMutex()
  }
  $mutex.Dispose()
}
'@
}

function Get-OpenClawGatewaySupervisorContent {
@'
$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $baseDir
$pauseFile = Join-Path $baseDir ".paused"
$upgradeLockFile = Join-Path $baseDir ".upgrade-lock"

if (Test-Path $pauseFile) {
  exit 0
}

if (Test-Path $upgradeLockFile) {
  exit 0
}

$userHome = [Environment]::GetFolderPath("UserProfile")
$npmrcPath = Join-Path $userHome ".npmrc"
$defaultPrefix = Join-Path $userHome ".npm-global-user"

function Get-PrefixPath {
  if (Test-Path $npmrcPath) {
    $line = Get-Content $npmrcPath | Where-Object { $_ -match '^prefix=' } | Select-Object -First 1
    if ($line) {
      return ($line -replace '^prefix=', '').Trim()
    }
  }

  return $defaultPrefix
}

function Get-NodeExePath {
  try {
    return (Get-Command node.exe -ErrorAction Stop).Source
  } catch {
    $fallback = Join-Path ${env:ProgramFiles} "nodejs\node.exe"
    if (Test-Path $fallback) {
      return $fallback
    }
  }

  throw "Unable to resolve node.exe from PATH."
}

$env:TMPDIR = $env:TEMP
$env:NO_PROXY = "127.0.0.1,localhost,::1"
$env:no_proxy = "127.0.0.1,localhost,::1"
$env:NODE_OPTIONS = "--dns-result-order=ipv4first"
Remove-Item Env:HTTP_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:HTTPS_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:http_proxy -ErrorAction SilentlyContinue
Remove-Item Env:https_proxy -ErrorAction SilentlyContinue
$env:OPENCLAW_GATEWAY_PORT = "18789"
$env:OPENCLAW_SYSTEMD_UNIT = "openclaw-gateway.service"
$env:OPENCLAW_WINDOWS_TASK_NAME = "OpenClaw Gateway"
$env:OPENCLAW_SERVICE_MARKER = "openclaw"
$env:OPENCLAW_SERVICE_KIND = "gateway"
$env:OPENCLAW_SERVICE_VERSION = "2026.3.13"

$nodeExe = Get-NodeExePath
$openclawEntry = Join-Path (Get-PrefixPath) "node_modules\openclaw\openclaw.mjs"
$port = 18789
$healthzUrl = "http://127.0.0.1:$port/healthz"
$probeIntervalSeconds = 10
$startupGraceSeconds = 60
$failureThreshold = 6
$restartDelaySeconds = 5
$logDir = Join-Path $env:TMPDIR "openclaw"
$logFile = Join-Path $logDir "gateway-supervisor.log"
$maxLogBytes = 2MB

if (-not (Test-Path $logDir)) {
  New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Rotate-Log {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    return
  }

  $item = Get-Item $Path
  if ($item.Length -lt $maxLogBytes) {
    return
  }

  $archive = "$Path.1"
  if (Test-Path $archive) {
    Remove-Item $archive -Force -ErrorAction SilentlyContinue
  }
  Move-Item -Path $Path -Destination $archive -Force
}

function Write-Log {
  param([string]$Message)
  Rotate-Log -Path $logFile
  $line = "[{0}] {1}" -f (Get-Date -Format "yyyy/MM/dd ddd HH:mm:ss.ff"), $Message
  Add-Content -Path $logFile -Value $line
}

function Get-GatewayProcessInfo {
  Get-CimInstance Win32_Process |
    Where-Object {
      $_.Name -eq "node.exe" -and
      (
        $_.CommandLine -like "*node_modules\openclaw\dist\index.js*gateway*" -or
        $_.CommandLine -like "*node_modules\openclaw\openclaw.mjs*gateway*"
      )
    }
}

function Test-GatewayProbe {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $healthzUrl -TimeoutSec 3
    return ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300)
  } catch {
    return $false
  }
}

function Start-GatewayProcess {
  Write-Log "starting gateway child on port $port"
  Start-Process -FilePath $nodeExe -ArgumentList @($openclawEntry, "gateway", "--port", "$port") -WorkingDirectory $baseDir -WindowStyle Hidden -PassThru
}

function Stop-GatewayProcess {
  param([int]$ProcessId)
  try {
    Stop-Process -Id $ProcessId -Force -ErrorAction Stop
    Write-Log "stopped gateway child pid=$ProcessId"
  } catch {
    Write-Log "stop process pid=$ProcessId failed: $($_.Exception.Message)"
  }
}

$mutex = New-Object System.Threading.Mutex($false, "Global\OpenClawGatewaySupervisor")
$hasMutex = $false

try {
  $hasMutex = $mutex.WaitOne(0, $false)
  if (-not $hasMutex) {
    Write-Log "another supervisor instance already owns the mutex; exiting"
    exit 0
  }

  Write-Log "supervisor active in $baseDir"

  while ($true) {
    try {
      if (Test-Path $upgradeLockFile) {
        Write-Log "upgrade lock detected; supervisor exiting"
        exit 0
      }

      $gatewayProcesses = @(Get-GatewayProcessInfo)

      if ($gatewayProcesses.Count -gt 1) {
        Write-Log "found $($gatewayProcesses.Count) gateway children; stopping duplicates"
        foreach ($duplicate in $gatewayProcesses) {
          Stop-GatewayProcess -ProcessId $duplicate.ProcessId
        }
        Start-Sleep -Seconds 2
        $gatewayProcesses = @()
      }

      if ($gatewayProcesses.Count -eq 1) {
        $proc = Get-Process -Id $gatewayProcesses[0].ProcessId -ErrorAction SilentlyContinue
        if ($null -eq $proc) {
          $gatewayProcesses = @()
        } else {
          Write-Log "adopting existing gateway child pid=$($proc.Id)"
        }
      }

      if ($gatewayProcesses.Count -eq 0) {
        $proc = Start-GatewayProcess
        Write-Log "gateway child started pid=$($proc.Id)"
      }

      $startupDeadline = (Get-Date).AddSeconds($startupGraceSeconds)
      $consecutiveFailures = 0

      while ($true) {
        Start-Sleep -Seconds $probeIntervalSeconds

        if (Test-Path $upgradeLockFile) {
          Write-Log "upgrade lock detected during probe loop; stopping child pid=$($proc.Id)"
          if (-not $proc.HasExited) {
            Stop-GatewayProcess -ProcessId $proc.Id
          }
          exit 0
        }

        $proc.Refresh()
        if ($proc.HasExited) {
          Write-Log "gateway child exited pid=$($proc.Id) code=$($proc.ExitCode)"
          break
        }

        $healthy = Test-GatewayProbe
        if ($healthy) {
          if ($consecutiveFailures -gt 0) {
            Write-Log "gateway probe recovered after $consecutiveFailures failures"
          }
          $consecutiveFailures = 0
          continue
        }

        if ((Get-Date) -lt $startupDeadline) {
          Write-Log "gateway probe failed during startup grace pid=$($proc.Id)"
          continue
        }

        $consecutiveFailures += 1
        Write-Log "gateway probe failed count=$consecutiveFailures pid=$($proc.Id)"

        if ($consecutiveFailures -ge $failureThreshold) {
          Write-Log "gateway deemed unhealthy; recycling pid=$($proc.Id)"
          Stop-GatewayProcess -ProcessId $proc.Id
          break
        }
      }

      Write-Log "waiting $restartDelaySeconds seconds before next launch cycle"
      Start-Sleep -Seconds $restartDelaySeconds
    } catch {
      Write-Log "supervisor loop exception: $($_.Exception.Message)"
      Start-Sleep -Seconds $restartDelaySeconds
    }
  }
}
finally {
  if ($hasMutex) {
    [void]$mutex.ReleaseMutex()
  }
  $mutex.Dispose()
}
'@
}

function Get-OpenClawUpgradeScriptContent {
@'
param(
  [string]$TargetVersion
)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$upgradeLock = Join-Path $baseDir ".upgrade-lock"
$logDir = Join-Path $env:TEMP "openclaw"
$logFile = Join-Path $logDir "upgrade-openclaw.log"
$userHome = [Environment]::GetFolderPath("UserProfile")
$nodeExe = $null
$npmCli = $null
$targetPrefix = Join-Path $userHome ".npm-global-user"
$targetModule = Join-Path $targetPrefix "node_modules\openclaw"
$npmrcPath = Join-Path $userHome ".npmrc"
$cachePath = Join-Path $env:LOCALAPPDATA "npm-cache"
$launcher = Join-Path $baseDir "run-hidden.vbs"
$restartAfterUpgrade = $false

if (-not (Test-Path $logDir)) {
  New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
  param([string]$Message)
  Add-Content -Path $logFile -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $Message)
}

function Get-LatestOpenClawVersion {
  try {
    $latest = Invoke-RestMethod -UseBasicParsing -Uri "https://registry.npmjs.org/openclaw/latest" -TimeoutSec 20
    if ($latest.version) {
      return [string]$latest.version
    }
  } catch {
    Write-Log "latest version lookup failed: $($_.Exception.Message)"
  }

  throw "Unable to resolve latest OpenClaw version"
}

function Resolve-NodeRuntime {
  try {
    $resolvedNode = (Get-Command node.exe -ErrorAction Stop).Source
  } catch {
    $fallback = Join-Path ${env:ProgramFiles} "nodejs\node.exe"
    if (Test-Path $fallback) {
      $resolvedNode = $fallback
    }
  }

  if (-not $resolvedNode) {
    throw "Unable to resolve node.exe from PATH."
  }

  $resolvedNpmCli = Join-Path (Split-Path -Parent $resolvedNode) "node_modules\npm\bin\npm-cli.js"
  if (-not (Test-Path $resolvedNpmCli)) {
    throw "Unable to resolve npm-cli.js near $resolvedNode"
  }

  return [pscustomobject]@{
    NodeExe = $resolvedNode
    NpmCli = $resolvedNpmCli
  }
}

function Invoke-Npm {
  param([string[]]$NpmArgs)
  $joined = ($NpmArgs -join " ")
  Write-Log "npm $joined"
  $stdoutFile = Join-Path $logDir "upgrade-npm.stdout.log"
  $stderrFile = Join-Path $logDir "upgrade-npm.stderr.log"
  Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue
  $proc = Start-Process -FilePath $nodeExe -ArgumentList (@($npmCli) + $NpmArgs) -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile
  foreach ($file in @($stdoutFile, $stderrFile)) {
    if (Test-Path $file) {
      Get-Content $file | ForEach-Object { Write-Log $_ }
    }
  }
  if ($proc.ExitCode -ne 0) {
    throw "npm failed: $joined"
  }
}

try {
  $runtime = Resolve-NodeRuntime
  $nodeExe = $runtime.NodeExe
  $npmCli = $runtime.NpmCli

  if ([string]::IsNullOrWhiteSpace($TargetVersion)) {
    $TargetVersion = Get-LatestOpenClawVersion
  }
  Write-Log "target version: $TargetVersion"

  Set-Content -Path $upgradeLock -Value ("upgrade started at {0}" -f (Get-Date -Format "s")) -Encoding ASCII
  Write-Log "upgrade lock created"

  Get-CimInstance Win32_Process |
    Where-Object {
      ($_.Name -like "powershell*" -and ($_.CommandLine -like "*gateway-watchdog.ps1*" -or $_.CommandLine -like "*gateway-supervisor.ps1*")) -or
      ($_.Name -eq "node.exe" -and $_.CommandLine -like "*openclaw*gateway*")
    } |
    ForEach-Object {
      Write-Log "stopping pid=$($_.ProcessId) name=$($_.Name)"
      Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

  Start-Sleep -Seconds 2

  if (Test-Path $npmrcPath) {
    $line = Get-Content $npmrcPath | Where-Object { $_ -match '^cache=' } | Select-Object -First 1
    if ($line) {
      $cachePath = ($line -replace '^cache=', '').Trim()
    }
  }

  New-Item -ItemType Directory -Path $targetPrefix -Force | Out-Null
  New-Item -ItemType Directory -Path $cachePath -Force | Out-Null

  Set-Content -Path $npmrcPath -Value @(
    "prefix=$targetPrefix",
    "cache=$cachePath"
  ) -Encoding ASCII
  Write-Log "npm prefix set to $targetPrefix"

  foreach ($path in @(
    $targetModule,
    (Join-Path $targetPrefix "openclaw"),
    (Join-Path $targetPrefix "openclaw.cmd"),
    (Join-Path $targetPrefix "openclaw.ps1"),
    (Join-Path $targetPrefix "openclaw.disabled"),
    (Join-Path $targetPrefix "openclaw.cmd.disabled")
  )) {
    if (Test-Path $path) {
      Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
      Write-Log "removed previous target file $path"
    }
  }

  Get-ChildItem -Path (Join-Path $targetPrefix "node_modules") -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like ".openclaw-*" } |
    ForEach-Object {
      Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
      Write-Log "removed stale temp dir $($_.FullName)"
    }

  Invoke-Npm -NpmArgs @("install","-g","openclaw@$TargetVersion","--prefix",$targetPrefix,"--force","--loglevel","verbose")

  $version = & (Join-Path $targetPrefix "openclaw.cmd") --version 2>&1
  $versionText = ($version | Out-String).Trim()
  Write-Log "installed version: $versionText"
  if ($versionText -notmatch [regex]::Escape($TargetVersion)) {
    throw "installed version mismatch: $versionText"
  }

  $restartAfterUpgrade = $true
}
finally {
  if (Test-Path $upgradeLock) {
    Remove-Item $upgradeLock -Force -ErrorAction SilentlyContinue
    Write-Log "upgrade lock removed"
  }

  if ($restartAfterUpgrade -and (Test-Path $launcher)) {
    Start-Process -FilePath "wscript.exe" -ArgumentList @("""$launcher""") -WindowStyle Hidden
    Write-Log "restart requested through $launcher"
  }
}
'@
}

function Ensure-OpenClawRuntimeFiles {
  $context = Get-OpenClawContext

  New-Item -ItemType Directory -Path $context.Home -Force | Out-Null
  New-Item -ItemType Directory -Path $context.TempLogDir -Force | Out-Null
  New-Item -ItemType Directory -Path $context.BackupRoot -Force | Out-Null

  Write-OpenClawTextFile -Path $context.LauncherVbs -Content (Get-OpenClawRunHiddenVbsContent) -Encoding "ASCII"
  Write-OpenClawTextFile -Path $context.GatewayCmd -Content (Get-OpenClawGatewayCmdContent) -Encoding "ASCII"
  Write-OpenClawTextFile -Path $context.WatchdogScript -Content (Get-OpenClawGatewayWatchdogContent) -Encoding "ASCII"
  Write-OpenClawTextFile -Path $context.SupervisorScript -Content (Get-OpenClawGatewaySupervisorContent) -Encoding "ASCII"
  Write-OpenClawTextFile -Path $context.UpgradeScript -Content (Get-OpenClawUpgradeScriptContent) -Encoding "ASCII"

  if ($script:OpenClawToolkitSelfContent) {
    Write-OpenClawTextFile -Path $context.ToolkitScript -Content $script:OpenClawToolkitSelfContent -Encoding "UTF8"
  }
}

function Ensure-OpenClawKeepaliveTask {
  $context = Get-OpenClawContext
  Ensure-OpenClawRuntimeFiles

  $launcherArgument = ('wscript.exe "{0}"' -f $context.LauncherVbs)
  $existingTask = $null
  try {
    $existingTask = Get-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction Stop
  } catch {}

  if ($existingTask) {
    try {
      Enable-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction SilentlyContinue | Out-Null
    } catch {}
    return $true
  }

  try {
    $create = Start-Process -FilePath "schtasks.exe" -ArgumentList @(
      "/Create",
      "/TN", $context.KeepaliveTask,
      "/TR", $launcherArgument,
      "/SC", "MINUTE",
      "/MO", "1",
      "/F"
    ) -WindowStyle Hidden -Wait -PassThru

    return ($create.ExitCode -eq 0)
  } catch {
    return $false
  }
}

function Add-OpenClawZipEntryFromBytes {
  param(
    [System.IO.Compression.ZipArchive]$Archive,
    [string]$EntryName,
    [byte[]]$ContentBytes
  )

  $normalized = ($EntryName -replace '\\', '/').TrimStart('/')
  $entry = $Archive.CreateEntry($normalized, [System.IO.Compression.CompressionLevel]::Fastest)
  $stream = $entry.Open()
  try {
    $stream.Write($ContentBytes, 0, $ContentBytes.Length)
  } finally {
    $stream.Dispose()
  }
}

function Add-OpenClawZipEntryFromFile {
  param(
    [System.IO.Compression.ZipArchive]$Archive,
    [string]$EntryName,
    [string]$SourcePath
  )

  $normalized = ($EntryName -replace '\\', '/').TrimStart('/')
  $entry = $Archive.CreateEntry($normalized, [System.IO.Compression.CompressionLevel]::Fastest)
  $input = [System.IO.File]::Open($SourcePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
  try {
    $output = $entry.Open()
    try {
      $input.CopyTo($output)
    } finally {
      $output.Dispose()
    }
  } finally {
    $input.Dispose()
  }
}

function Add-OpenClawDirectoryToZip {
  param(
    [System.IO.Compression.ZipArchive]$Archive,
    [string]$SourcePath,
    [string]$EntryRoot,
    [scriptblock]$ExcludeItem
  )

  if (-not (Test-Path $SourcePath)) {
    return
  }

  $items = @(Get-ChildItem $SourcePath -Force -ErrorAction SilentlyContinue)
  if ($items.Count -eq 0) {
    $dirEntry = (($EntryRoot -replace '\\', '/').Trim('/')) + '/'
    if ($dirEntry -ne "/") {
      $Archive.CreateEntry($dirEntry) | Out-Null
    }
    return
  }

  foreach ($item in $items) {
    if ($ExcludeItem -and (& $ExcludeItem $item)) {
      continue
    }

    $entryName = [System.IO.Path]::Combine($EntryRoot, $item.Name)
    if ($item.PSIsContainer) {
      Add-OpenClawDirectoryToZip -Archive $Archive -SourcePath $item.FullName -EntryRoot $entryName -ExcludeItem $ExcludeItem
    } else {
      Add-OpenClawZipEntryFromFile -Archive $Archive -EntryName $entryName -SourcePath $item.FullName
    }
  }
}

function Get-OpenClawCachePath {
  $context = Get-OpenClawContext
  if (Test-Path $context.NpmrcPath) {
    $line = Get-Content $context.NpmrcPath | Where-Object { $_ -match '^cache=' } | Select-Object -First 1
    if ($line) {
      return ($line -replace '^cache=', '').Trim()
    }
  }

  return $context.DefaultCachePath
}

function Write-OpenClawNpmConfig {
  param(
    [string]$Prefix = $(Get-OpenClawDefaultPrefix),
    [string]$CachePath = $(Get-OpenClawCachePath)
  )

  $context = Get-OpenClawContext
  New-Item -ItemType Directory -Path $Prefix -Force | Out-Null
  New-Item -ItemType Directory -Path $CachePath -Force | Out-Null

  $npmrcLines = @(
    "prefix=$Prefix",
    "cache=$CachePath"
  )
  Set-Content -Path $context.NpmrcPath -Value $npmrcLines -Encoding ASCII
}

function Ensure-OpenClawPrefixPath {
  param(
    [string]$Prefix = $(Get-OpenClawPrefix)
  )

  $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
  $parts = New-Object System.Collections.Generic.List[string]
  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

  foreach ($part in ($userPath -split ';')) {
    $trimmed = $part.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
      continue
    }

    if ($seen.Add($trimmed)) {
      $parts.Add($trimmed)
    }
  }

  if ($seen.Contains($Prefix)) {
    $filteredParts = New-Object System.Collections.Generic.List[string]
    foreach ($existingPart in $parts) {
      if ($existingPart -ne $Prefix) {
        $filteredParts.Add($existingPart)
      }
    }
    $parts = $filteredParts
  }

  $parts.Insert(0, $Prefix)
  [Environment]::SetEnvironmentVariable("Path", ($parts -join ';'), "User")
}

function Get-OpenClawPrefix {
  $context = Get-OpenClawContext
  if (Test-Path $context.NpmrcPath) {
    $line = Get-Content $context.NpmrcPath | Where-Object { $_ -match '^prefix=' } | Select-Object -First 1
    if ($line) {
      return ($line -replace '^prefix=', '').Trim()
    }
  }

  return $context.DefaultPrefix
}

function Get-OpenClawCommandPath {
  $prefix = Get-OpenClawPrefix
  $cmd = Join-Path $prefix "openclaw.cmd"
  if (Test-Path $cmd) {
    return $cmd
  }

  return $null
}

function Get-OpenClawCurrentVersion {
  $cmd = Get-OpenClawCommandPath
  if (-not $cmd) {
    return $null
  }

  try {
    $raw = & $cmd --version 2>$null | Out-String
    if ($raw -match '(\d+\.\d+\.\d+)') {
      return $matches[1]
    }
  } catch {
    return $null
  }

  return $null
}

function Get-OpenClawLatestVersion {
  try {
    $latest = Invoke-RestMethod -UseBasicParsing -Uri "https://registry.npmjs.org/openclaw/latest" -TimeoutSec 20
    if ($latest.version) {
      return [string]$latest.version
    }
  } catch {
    return $null
  }

  return $null
}

function Compare-OpenClawVersion {
  param(
    [string]$Left,
    [string]$Right
  )

  if ([string]::IsNullOrWhiteSpace($Left) -and [string]::IsNullOrWhiteSpace($Right)) { return 0 }
  if ([string]::IsNullOrWhiteSpace($Left)) { return -1 }
  if ([string]::IsNullOrWhiteSpace($Right)) { return 1 }

  $leftParts = $Left.Split('.') | ForEach-Object { [int]$_ }
  $rightParts = $Right.Split('.') | ForEach-Object { [int]$_ }
  $count = [Math]::Max($leftParts.Count, $rightParts.Count)

  for ($i = 0; $i -lt $count; $i++) {
    $l = if ($i -lt $leftParts.Count) { $leftParts[$i] } else { 0 }
    $r = if ($i -lt $rightParts.Count) { $rightParts[$i] } else { 0 }
    if ($l -gt $r) { return 1 }
    if ($l -lt $r) { return -1 }
  }

  return 0
}

function Test-OpenClawGatewayHealthy {
  $context = Get-OpenClawContext
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $context.HealthzUrl -TimeoutSec 5
    return ($response.StatusCode -eq 200 -and $response.Content -match '"ok":true')
  } catch {
    return $false
  }
}

function Get-OpenClawProcesses {
  Get-CimInstance Win32_Process -Filter "Name = 'wscript.exe' OR Name = 'node.exe' OR Name = 'powershell.exe' OR Name = 'pwsh.exe'" |
    Where-Object {
      ($_.Name -eq "wscript.exe" -and $_.CommandLine -like "*run-hidden.vbs*") -or
      ($_.Name -like "powershell*" -and ($_.CommandLine -like "*gateway-watchdog.ps1*" -or $_.CommandLine -like "*gateway-supervisor.ps1*" -or $_.CommandLine -like "*upgrade-openclaw.ps1*")) -or
      ($_.Name -eq "node.exe" -and $_.CommandLine -like "*openclaw*gateway*")
    }
}

function Get-OpenClawLatestAuditRun {
  $context = Get-OpenClawContext
  if (-not (Test-Path $context.AuditRoot)) {
    return $null
  }

  return Get-ChildItem $context.AuditRoot -Directory |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
}

function Get-OpenClawAutoStartEnabled {
  $context = Get-OpenClawContext
  return (Test-Path $context.StartupLauncher)
}

function Set-OpenClawAutoStart {
  param(
    [bool]$Enabled
  )

  $context = Get-OpenClawContext

  if ($Enabled) {
    Ensure-OpenClawRuntimeFiles
    $launcher = $context.LauncherVbs
    $content = @"
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
launcher = "$launcher"

If fso.FileExists(launcher) Then
  shell.Run "wscript.exe """ & launcher & """", 0, False
End If
"@

    New-Item -ItemType Directory -Path (Split-Path -Parent $context.StartupLauncher) -Force | Out-Null
    Set-Content -Path $context.StartupLauncher -Value $content -Encoding ASCII

    [void](Ensure-OpenClawKeepaliveTask)
  } else {
    Remove-Item $context.StartupLauncher -Force -ErrorAction SilentlyContinue
  }

  return Get-OpenClawStatus
}

function Get-OpenClawRecentErrors {
  param(
    [int]$MaxItems = 12,
    [int]$TailLines = 200
  )

  $context = Get-OpenClawContext
  $files = @()
  $pattern = 'error|failed|exception|timeout|denied|unhealthy|unauthorized|1006|abnormal closure|closed'

  foreach ($path in @(
    $context.WatchdogLog,
    "$($context.WatchdogLog).1",
    $context.SupervisorLog,
    "$($context.SupervisorLog).1",
    $context.UpgradeLog,
    $context.UpgradeStdoutLog,
    $context.UpgradeStderrLog
  )) {
    if (Test-Path $path) {
      $files += Get-Item $path
    }
  }

  if (Test-Path $context.GatewayLogDir) {
    Get-ChildItem $context.GatewayLogDir -Filter *.log -File -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 3 |
      ForEach-Object { $files += $_ }
  }

  $latestAudit = Get-OpenClawLatestAuditRun
  if ($latestAudit) {
    Get-ChildItem $latestAudit.FullName -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Extension -in @(".log", ".txt", ".json", ".jsonl", ".md") } |
      ForEach-Object { $files += $_ }
  }

  $recentItems = @()
  foreach ($file in ($files | Sort-Object FullName -Unique)) {
    $lines = @(Get-Content $file.FullName -Tail $TailLines -ErrorAction SilentlyContinue)
    if ($lines.Count -eq 0) {
      continue
    }

    for ($i = 0; $i -lt $lines.Count; $i++) {
      $line = [string]$lines[$i]
      if ($line -match $pattern) {
        $message = $line.Trim()
        if ($message.StartsWith("{")) {
          try {
            $json = $message | ConvertFrom-Json -ErrorAction Stop
            if ($json.PSObject.Properties.Name -contains "message") {
              $message = [string]$json.message
            } elseif ($json.PSObject.Properties.Name -contains "msg") {
              $message = [string]$json.msg
            } elseif ($json.PSObject.Properties.Name -contains "1") {
              $message = [string]$json."1"
            }
          } catch {}
        }

        $entry = [pscustomobject]@{
          File = $file.Name
          Path = $file.FullName
          LastWriteTime = $file.LastWriteTime
          TailIndex = $i
          Message = $message
        }
        $recentItems += $entry
      }
    }
  }

  return $recentItems |
    Sort-Object LastWriteTime, TailIndex -Descending |
    Select-Object -First $MaxItems
}

function Get-OpenClawRecentErrorSummary {
  param(
    [int]$MaxItems = 12
  )

  $items = @(Get-OpenClawRecentErrors -MaxItems $MaxItems)
  if ($items.Count -eq 0) {
    return "No recent error lines matched in the latest OpenClaw logs."
  }

  return ($items | ForEach-Object {
    $message = $_.Message
    if ($message.Length -gt 220) {
      $message = $message.Substring(0, 217) + "..."
    }

    "[{0}] {1}" -f $_.File, $message
  }) -join [Environment]::NewLine
}

function Get-OpenClawBackups {
  $context = Get-OpenClawContext
  if (-not (Test-Path $context.BackupRoot)) {
    return @()
  }

  return @(Get-ChildItem $context.BackupRoot -File -Filter *.ocbackup.zip -ErrorAction SilentlyContinue |
    Where-Object { $_.Length -gt 0 } |
    Sort-Object LastWriteTime -Descending |
    ForEach-Object {
      [pscustomobject]@{
        Name = $_.BaseName
        FileName = $_.Name
        FullName = $_.FullName
        LastWriteTime = $_.LastWriteTime
        Length = $_.Length
      }
    })
}

function Get-OpenClawStatus {
  param(
    [switch]$IncludeLatestVersion
  )

  $context = Get-OpenClawContext
  $latestAudit = Get-OpenClawLatestAuditRun
  $keepaliveTask = $null
  $backups = @(Get-OpenClawBackups)

  try { $keepaliveTask = Get-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction Stop } catch {}

  $paused = Test-Path $context.PauseFile
  $upgradeLocked = Test-Path $context.UpgradeLockFile
  $processes = @(Get-OpenClawProcesses)

  $currentVersion = Get-OpenClawCurrentVersion
  $latestVersion = if ($IncludeLatestVersion) { Get-OpenClawLatestVersion } else { $null }
  $healthy = $false

  if (-not $paused -and -not $upgradeLocked -and $processes.Count -gt 0) {
    $healthy = Test-OpenClawGatewayHealthy
  }

  $mode = "Stopped"
  if ($upgradeLocked) {
    $mode = "Updating"
  } elseif ($paused) {
    $mode = "Paused"
  } elseif ($healthy) {
    $mode = "Running"
  } elseif ($processes.Count -gt 0) {
    $mode = "Starting"
  }

  [pscustomobject]@{
    CurrentVersion = $currentVersion
    LatestVersion = $latestVersion
    UpdateAvailable = ($IncludeLatestVersion -and (Compare-OpenClawVersion $latestVersion $currentVersion) -gt 0)
    GatewayHealthy = $healthy
    Mode = $mode
    Paused = $paused
    UpgradeLocked = $upgradeLocked
    KeepaliveTaskState = if ($keepaliveTask) { [string]$keepaliveTask.State } else { "Missing" }
    AutoStartEnabled = Get-OpenClawAutoStartEnabled
    ProcessCount = $processes.Count
    Prefix = Get-OpenClawPrefix
    CommandPath = Get-OpenClawCommandPath
    LatestAuditPath = if ($latestAudit) { $latestAudit.FullName } else { $null }
    BackupRoot = $context.BackupRoot
    BackupCount = $backups.Count
    LatestBackupPath = if ($backups.Count -gt 0) { $backups[0].FullName } else { $null }
    WatchdogLog = $context.WatchdogLog
    SupervisorLog = $context.SupervisorLog
    UpgradeLog = $context.UpgradeLog
    GatewayLogDir = $context.GatewayLogDir
    StartupLauncherPath = $context.StartupLauncher
  }
}

function Start-OpenClawSilent {
  $context = Get-OpenClawContext

  if (Test-Path $context.UpgradeLockFile) {
    throw "OpenClaw is currently updating. Please wait for the update to finish."
  }

  Ensure-OpenClawRuntimeFiles
  if (Test-Path $context.PauseFile) {
    Remove-Item $context.PauseFile -Force -ErrorAction SilentlyContinue
  }

  Write-OpenClawNpmConfig
  Ensure-OpenClawPrefixPath
  [void](Ensure-OpenClawKeepaliveTask)

  try {
    Enable-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction SilentlyContinue | Out-Null
  } catch {}

  try {
    Start-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction SilentlyContinue
  } catch {}

  if (Test-Path $context.LauncherVbs) {
    Start-Process -FilePath "wscript.exe" -ArgumentList @("""$($context.LauncherVbs)""") -WindowStyle Hidden | Out-Null
  }

  Start-Sleep -Seconds 1
  return Get-OpenClawStatus
}

function Stop-OpenClawSilent {
  $context = Get-OpenClawContext
  New-Item -ItemType Directory -Path $context.Home -Force | Out-Null
  Set-Content -Path $context.PauseFile -Value ("paused at {0}" -f (Get-Date -Format "s")) -Encoding ASCII

  Get-OpenClawProcesses | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
  }

  Start-Sleep -Seconds 1
  return Get-OpenClawStatus
}

function Start-OpenClawUpdate {
  param(
    [string]$TargetVersion
  )

  $context = Get-OpenClawContext
  Ensure-OpenClawRuntimeFiles
  if (Test-Path $context.UpgradeLockFile) {
    throw "An OpenClaw update is already running."
  }

  $args = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $context.UpgradeScript)
  if (-not [string]::IsNullOrWhiteSpace($TargetVersion)) {
    $args += @("-TargetVersion", $TargetVersion)
  }

  return Start-Process -FilePath "powershell.exe" -ArgumentList $args -WindowStyle Hidden -PassThru
}

function Get-SafeOpenClawBackupName {
  param([string]$BackupName)

  if ([string]::IsNullOrWhiteSpace($BackupName)) {
    return "backup"
  }

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() + @(' ')
  $pattern = '[{0}]+' -f ([regex]::Escape((-join $invalidChars)))
  $safe = ($BackupName -replace $pattern, '-').Trim('-')
  if ([string]::IsNullOrWhiteSpace($safe)) {
    return "backup"
  }

  return $safe
}

function Test-OpenClawBackupExcludedItem {
  param([string]$Name)

  foreach ($pattern in @(
      "backups",
      "logs",
      "desktop-backup",
      "delivery-queue",
      "workspace",
      "workspace-*",
      "backup-main-fix-*",
      ".paused",
      ".upgrade-lock",
      "gateway-run.out.log",
      "gateway-run.err.log",
      "gateway-test.log"
    )) {
    if ($Name -like $pattern) {
      return $true
    }
  }

  return $false
}

function New-OpenClawBackup {
  param(
    [string]$BackupName
  )

  $context = Get-OpenClawContext
  Ensure-OpenClawRuntimeFiles
  $safeName = Get-SafeOpenClawBackupName $BackupName
  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $version = Get-OpenClawCurrentVersion
  if ([string]::IsNullOrWhiteSpace($version)) {
    $version = "unknown"
  }

  New-Item -ItemType Directory -Path $context.BackupRoot -Force | Out-Null
  $archiveName = "{0}--{1}--{2}.ocbackup.zip" -f $timestamp, $safeName, $version
  $archivePath = Join-Path $context.BackupRoot $archiveName

  $prefix = Get-OpenClawPrefix
  $manifest = [pscustomobject]@{
    backupSchemaVersion = 1
    backupName = $BackupName
    safeName = $safeName
    createdAt = (Get-Date).ToString("s")
    currentVersion = $version
    sourceUser = $env:USERNAME
    sourceComputer = $env:COMPUTERNAME
    sourcePrefix = $prefix
  }

  if (Test-Path $archivePath) {
    Remove-Item $archivePath -Force -ErrorAction SilentlyContinue
  }

  try {
    $archiveStream = [System.IO.File]::Open($archivePath, [System.IO.FileMode]::CreateNew)
    try {
      $archive = New-Object System.IO.Compression.ZipArchive($archiveStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
      try {
        Add-OpenClawZipEntryFromBytes -Archive $archive -EntryName "manifest.json" -ContentBytes ([System.Text.Encoding]::UTF8.GetBytes(($manifest | ConvertTo-Json -Depth 5)))

        if (Test-Path $context.Home) {
          $excludeHomeItem = {
            param($Item)
            Test-OpenClawBackupExcludedItem -Name $Item.Name
          }
          Add-OpenClawDirectoryToZip -Archive $archive -SourcePath $context.Home -EntryRoot "payload\.openclaw" -ExcludeItem $excludeHomeItem
        }

        foreach ($itemName in @("openclaw", "openclaw.cmd", "openclaw.ps1", "node_modules\openclaw")) {
          $source = Join-Path $prefix $itemName
          if (-not (Test-Path $source)) {
            continue
          }

          $entryRoot = [System.IO.Path]::Combine("payload\.npm-global-user", $itemName)
          if ((Get-Item $source).PSIsContainer) {
            Add-OpenClawDirectoryToZip -Archive $archive -SourcePath $source -EntryRoot $entryRoot -ExcludeItem $null
          } else {
            Add-OpenClawZipEntryFromFile -Archive $archive -EntryName $entryRoot -SourcePath $source
          }
        }

        if (Test-Path $context.NpmrcPath) {
          Add-OpenClawZipEntryFromFile -Archive $archive -EntryName "payload\.npmrc" -SourcePath $context.NpmrcPath
        }
      } finally {
        $archive.Dispose()
      }
    } finally {
      if ($archiveStream) {
        $archiveStream.Dispose()
      }
    }
  } catch {
    Remove-Item $archivePath -Force -ErrorAction SilentlyContinue
    throw
  }

  return $archivePath
}

function Restore-OpenClawBackup {
  param(
    [string]$BackupPath
  )

  if ([string]::IsNullOrWhiteSpace($BackupPath) -or -not (Test-Path $BackupPath)) {
    throw "Backup file not found: $BackupPath"
  }

  $context = Get-OpenClawContext
  $systemDrive = if ($env:SystemDrive) { $env:SystemDrive } else { "C:" }
  $restoreRoot = Join-Path $systemDrive ("oclrestore\{0}" -f ([guid]::NewGuid().ToString("N").Substring(0, 8)))
  $payloadRoot = Join-Path $restoreRoot "payload"
  $payloadHome = Join-Path $payloadRoot ".openclaw"
  $payloadPrefix = Join-Path $payloadRoot ".npm-global-user"
  $payloadNpmrc = Join-Path $payloadRoot ".npmrc"
  $targetPrefix = $context.DefaultPrefix

  Remove-Item $restoreRoot -Recurse -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Path $restoreRoot -Force | Out-Null
  [System.IO.Compression.ZipFile]::ExtractToDirectory($BackupPath, $restoreRoot)

  Stop-OpenClawSilent | Out-Null

  New-Item -ItemType Directory -Path $context.Home -Force | Out-Null
  New-Item -ItemType Directory -Path $targetPrefix -Force | Out-Null

  foreach ($item in Get-ChildItem $context.Home -Force -ErrorAction SilentlyContinue) {
    if ($item.Name -eq "backups") {
      continue
    }
    Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
  }

  foreach ($path in @(
    (Join-Path $targetPrefix "openclaw"),
    (Join-Path $targetPrefix "openclaw.cmd"),
    (Join-Path $targetPrefix "openclaw.ps1"),
    (Join-Path $targetPrefix "node_modules\openclaw"),
    (Join-Path $targetPrefix "node_modules\node_modules")
  )) {
    if (Test-Path $path) {
      Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  if (Test-Path $payloadHome) {
    foreach ($item in Get-ChildItem $payloadHome -Force) {
      if ($item.Name -in @(
          "openclaw-toolkit.ps1",
          "gateway-watchdog.ps1",
          "gateway-supervisor.ps1",
          "run-hidden.vbs",
          "gateway.cmd",
          "upgrade-openclaw.ps1",
          ".paused",
          ".upgrade-lock"
        )) {
        continue
      }
      Copy-Item $item.FullName (Join-Path $context.Home $item.Name) -Recurse -Force
    }
  }

  if (Test-Path $payloadPrefix) {
    foreach ($itemName in @("openclaw", "openclaw.cmd", "openclaw.ps1")) {
      $source = Join-Path $payloadPrefix $itemName
      if (Test-Path $source) {
        Copy-Item $source (Join-Path $targetPrefix $itemName) -Recurse -Force
      }
    }

    $payloadModule = Join-Path $payloadPrefix "node_modules\openclaw"
    if (Test-Path $payloadModule) {
      $targetNodeModules = Join-Path $targetPrefix "node_modules"
      New-Item -ItemType Directory -Path $targetNodeModules -Force | Out-Null
      Copy-Item $payloadModule (Join-Path $targetNodeModules "openclaw") -Recurse -Force
    }
  }

  $cachePath = Get-OpenClawCachePath
  if (Test-Path $payloadNpmrc) {
    try {
      $line = Get-Content $payloadNpmrc | Where-Object { $_ -match '^cache=' } | Select-Object -First 1
      if ($line) {
        $cachePath = ($line -replace '^cache=', '').Trim()
      }
    } catch {}
  }

  Write-OpenClawNpmConfig -Prefix $targetPrefix -CachePath $cachePath
  Ensure-OpenClawPrefixPath -Prefix $targetPrefix
  Ensure-OpenClawRuntimeFiles
  [void](Ensure-OpenClawKeepaliveTask)

  Remove-Item $restoreRoot -Recurse -Force -ErrorAction SilentlyContinue
  return Stop-OpenClawSilent
}

function Export-OpenClawErrorLogs {
  param(
    [string]$OutputZipPath
  )

  $context = Get-OpenClawContext
  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  if ([string]::IsNullOrWhiteSpace($OutputZipPath)) {
    $OutputZipPath = Join-Path $context.Desktop ("OpenClaw-ErrorLogs-{0}.zip" -f $timestamp)
  }

  $stagingRoot = Join-Path $env:TEMP ("openclaw-export-{0}" -f $timestamp)
  $stagingLogs = Join-Path $stagingRoot "logs"
  $summaryFile = Join-Path $stagingRoot "errors-summary.txt"
  $statusFile = Join-Path $stagingRoot "status.json"

  Remove-Item $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Path $stagingLogs -Force | Out-Null

  $copyTargets = @()
  foreach ($file in @(
    $context.WatchdogLog,
    "$($context.WatchdogLog).1",
    $context.SupervisorLog,
    "$($context.SupervisorLog).1",
    $context.UpgradeLog,
    $context.UpgradeStdoutLog,
    $context.UpgradeStderrLog,
    $context.NpmrcPath,
    (Join-Path $context.Home "openclaw.json")
  )) {
    if (Test-Path $file) {
      $copyTargets += Get-Item $file
    }
  }

  if (Test-Path $context.GatewayLogDir) {
    $copyTargets += Get-ChildItem $context.GatewayLogDir -Filter *.log -File -ErrorAction SilentlyContinue
  }

  if (Test-Path $context.AuditRoot) {
    $copyTargets += Get-ChildItem $context.AuditRoot -Directory |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 3 |
      ForEach-Object { Get-ChildItem $_.FullName -File -ErrorAction SilentlyContinue }
  }

  foreach ($item in $copyTargets | Sort-Object FullName -Unique) {
    $relative = $item.FullName.Replace(':', '')
    $target = Join-Path $stagingLogs $relative
    $targetDir = Split-Path -Parent $target
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    Copy-Item $item.FullName $target -Force
  }

  $status = Get-OpenClawStatus -IncludeLatestVersion
  $status | ConvertTo-Json -Depth 5 | Set-Content -Path $statusFile -Encoding UTF8

  $patterns = @(
    'error',
    'failed',
    'exception',
    'timeout',
    'denied',
    'unhealthy',
    'unauthorized',
    '1006',
    'closed'
  )

  $summaryLines = New-Object System.Collections.Generic.List[string]
  $summaryLines.Add("OpenClaw error summary")
  $summaryLines.Add(("Generated: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")))
  $summaryLines.Add("")

  $logFiles = Get-ChildItem $stagingLogs -Recurse -File -ErrorAction SilentlyContinue
  foreach ($logFile in $logFiles) {
    foreach ($match in Select-String -Path $logFile.FullName -Pattern $patterns -SimpleMatch -ErrorAction SilentlyContinue) {
      $summaryLines.Add(("{0}:{1}: {2}" -f $logFile.FullName.Replace($stagingLogs, 'logs'), $match.LineNumber, $match.Line.Trim()))
    }
  }

  if ($summaryLines.Count -le 3) {
    $summaryLines.Add("No obvious error lines were matched. Raw logs are still included.")
  }

  $summaryLines | Set-Content -Path $summaryFile -Encoding UTF8

  if (Test-Path $OutputZipPath) {
    Remove-Item $OutputZipPath -Force -ErrorAction SilentlyContinue
  }

  Compress-Archive -Path (Join-Path $stagingRoot "*") -DestinationPath $OutputZipPath -Force
  Remove-Item $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue

  return $OutputZipPath
}

