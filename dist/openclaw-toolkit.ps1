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

function Get-OpenClawConfigPath {
  return (Join-Path (Join-Path (Get-OpenClawUserHome) ".openclaw") "openclaw.json")
}

function Get-OpenClawConfiguredGatewayPort {
  $defaultPort = 18789

  $envPortRaw = [string]$env:OPENCLAW_GATEWAY_PORT
  if (-not [string]::IsNullOrWhiteSpace($envPortRaw)) {
    $envPort = 0
    if ([int]::TryParse($envPortRaw.Trim(), [ref]$envPort) -and $envPort -ge 1 -and $envPort -le 65535) {
      return $envPort
    }
  }

  $configPath = Get-OpenClawConfigPath
  if (Test-Path $configPath) {
    try {
      $config = Get-Content $configPath -Raw | ConvertFrom-Json -ErrorAction Stop
      $portValue = $config.gateway.port
      $configPort = 0
      if ($null -ne $portValue -and [int]::TryParse([string]$portValue, [ref]$configPort) -and $configPort -ge 1 -and $configPort -le 65535) {
        return $configPort
      }
    } catch {}
  }

  return $defaultPort
}

function Get-OpenClawGatewayLogDirectories {
  $systemDrive = if ($env:SystemDrive) { $env:SystemDrive } else { "C:" }
  $dirs = New-Object System.Collections.Generic.List[string]

  foreach ($candidate in @(
    (Join-Path $env:LOCALAPPDATA "Temp\openclaw"),
    (Join-Path $env:TEMP "openclaw"),
    (Join-Path $systemDrive "tmp\openclaw")
  )) {
    if ([string]::IsNullOrWhiteSpace($candidate)) {
      continue
    }

    if (-not $dirs.Contains($candidate)) {
      $dirs.Add($candidate)
    }
  }

  return @($dirs)
}

function Get-OpenClawGatewayLockPaths {
  $configPath = Get-OpenClawConfigPath
  $sha256 = [System.Security.Cryptography.SHA256]::Create()
  try {
    $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($configPath))
  } finally {
    $sha256.Dispose()
  }

  $hash = ([System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLowerInvariant()).Substring(0, 8)
  $paths = New-Object System.Collections.Generic.List[string]

  foreach ($dir in @(Get-OpenClawGatewayLogDirectories)) {
    if ([string]::IsNullOrWhiteSpace($dir)) {
      continue
    }

    $path = Join-Path $dir ("gateway.{0}.lock" -f $hash)
    if (-not $paths.Contains($path)) {
      $paths.Add($path)
    }
  }

  return @($paths)
}

function Convert-OpenClawCimDateTime {
  param(
    $Value
  )

  if ($Value -is [datetime]) {
    return [datetime]$Value
  }

  if ($Value -is [string] -and -not [string]::IsNullOrWhiteSpace($Value)) {
    try {
      return [System.Management.ManagementDateTimeConverter]::ToDateTime($Value)
    } catch {}
  }

  return $null
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
  $defaultPrefix = Get-OpenClawDefaultPrefix
  $systemDrive = if ($env:SystemDrive) { $env:SystemDrive } else { "C:" }
  $gatewayPort = Get-OpenClawConfiguredGatewayPort
  $gatewayLogDirs = Get-OpenClawGatewayLogDirectories
  $tempLogDir = $gatewayLogDirs[0]

  [pscustomobject]@{
    UserHome = $userHome
    Home = $openClawHome
    PauseFile = Join-Path $openClawHome ".paused"
    DesiredRunUntilFile = Join-Path $openClawHome ".desired-running-until"
    UpgradeLockFile = Join-Path $openClawHome ".upgrade-lock"
    LauncherVbs = Join-Path $openClawHome "run-hidden.vbs"
    WatchdogLauncherVbs = Join-Path $openClawHome "watchdog-launcher.vbs"
    GatewayLauncherVbs = Join-Path $openClawHome "gateway-launcher.vbs"
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
    GatewayLogDir = $gatewayLogDirs[0]
    GatewayLogDirs = $gatewayLogDirs
    AuditRoot = Join-Path $openClawHome "logs\discord-audit"
    BackupRoot = Join-Path $openClawHome "backups"
    NpmrcPath = Join-Path $userHome ".npmrc"
    DefaultPrefix = $defaultPrefix
    KeepaliveTask = "OpenClaw Gateway Keepalive"
    GatewayTask = "OpenClaw Gateway"
    GatewayPort = $gatewayPort
    HealthzUrl = ("http://127.0.0.1:{0}/healthz" -f $gatewayPort)
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

  try {
    Set-Content -Path $Path -Value $Content -Encoding $Encoding
  } catch [System.IO.IOException] {
    return
  }
}

function Get-OpenClawNodeExePath {
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

function Get-OpenClawEntryPath {
  $prefix = Get-OpenClawPrefix
  $entry = Join-Path $prefix "node_modules\openclaw\openclaw.mjs"
  if (Test-Path $entry) {
    return $entry
  }

  return $null
}

function Set-OpenClawDesiredRunWindow {
  param(
    [int]$Hours = 24
  )

  $context = Get-OpenClawContext
  New-Item -ItemType Directory -Path $context.Home -Force | Out-Null
  $deadline = (Get-Date).AddHours($Hours).ToString("s")
  Set-Content -Path $context.DesiredRunUntilFile -Value $deadline -Encoding ASCII
  return $deadline
}

function Clear-OpenClawDesiredRunWindow {
  $context = Get-OpenClawContext
  Remove-Item $context.DesiredRunUntilFile -Force -ErrorAction SilentlyContinue
}

function Get-OpenClawDesiredRunDeadline {
  $context = Get-OpenClawContext
  if (-not (Test-Path $context.DesiredRunUntilFile)) {
    return $null
  }

  try {
    $raw = (Get-Content $context.DesiredRunUntilFile -Raw -ErrorAction Stop).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) {
      return $null
    }

    return [datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture)
  } catch {
    return $null
  }
}

function Test-OpenClawDesiredRunActive {
  $deadline = Get-OpenClawDesiredRunDeadline
  if ($null -eq $deadline) {
    return $false
  }

  if ($deadline -le (Get-Date)) {
    Clear-OpenClawDesiredRunWindow
    return $false
  }

  return $true
}

function Get-OpenClawRunHiddenVbsContent {
@'
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
pauseFile = fso.BuildPath(baseDir, ".paused")
desiredRunFile = fso.BuildPath(baseDir, ".desired-running-until")
upgradeLockFile = fso.BuildPath(baseDir, ".upgrade-lock")
watchdogLauncher = fso.BuildPath(baseDir, "watchdog-launcher.vbs")

Function Pad2(value)
  Pad2 = Right("0" & CStr(value), 2)
End Function

Function ToIsoLocal(dt)
  ToIsoLocal = Year(dt) & "-" & Pad2(Month(dt)) & "-" & Pad2(Day(dt)) & "T" & Pad2(Hour(dt)) & ":" & Pad2(Minute(dt)) & ":" & Pad2(Second(dt))
End Function

If fso.FileExists(pauseFile) Then
  WScript.Quit 0
End If
If fso.FileExists(upgradeLockFile) Then
  WScript.Quit 0
End If
Set stream = fso.CreateTextFile(desiredRunFile, True, False)
stream.Write ToIsoLocal(DateAdd("h", 24, Now))
stream.Close
shell.Run "wscript.exe """ & watchdogLauncher & """", 0, False
'@
}

function Get-OpenClawWatchdogLauncherVbsContent {
@'
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
watchdogScript = fso.BuildPath(baseDir, "gateway-watchdog.ps1")
cmd = "powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File """ & watchdogScript & """"
shell.Run cmd, 0, False
'@
}

function Get-OpenClawGatewayLauncherVbsContent {
  $context = Get-OpenClawContext
  $nodeExe = Get-OpenClawNodeExePath
  $entry = Get-OpenClawEntryPath
  $serviceVersion = Get-OpenClawCurrentVersion
  if ([string]::IsNullOrWhiteSpace($entry)) {
    $entry = Join-Path (Get-OpenClawDefaultPrefix) "node_modules\openclaw\openclaw.mjs"
  }
  if ([string]::IsNullOrWhiteSpace($serviceVersion)) {
    $serviceVersion = "2026.3.13"
  }

@"
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("PROCESS")
env("HOME") = "$($context.UserHome)"
env("TMPDIR") = "$env:TEMP"
env("OPENCLAW_STATE_DIR") = "$($context.Home)"
env("OPENCLAW_CONFIG_PATH") = "$($context.Home)\openclaw.json"
env("OPENCLAW_GATEWAY_PORT") = "$($context.GatewayPort)"
env("OPENCLAW_WINDOWS_TASK_NAME") = "$($context.GatewayTask)"
env("OPENCLAW_SERVICE_MARKER") = "openclaw"
env("OPENCLAW_SERVICE_KIND") = "gateway"
env("OPENCLAW_SERVICE_VERSION") = "$serviceVersion"
shell.Run """" & "$nodeExe" & """" & " --disable-warning=ExperimentalWarning " & """" & "$entry" & """" & " gateway --port $($context.GatewayPort) --force", 0, False
"@
}

function Get-OpenClawGatewayCmdContent {
  $context = Get-OpenClawContext
  $nodeExe = Get-OpenClawNodeExePath
  $entry = Get-OpenClawEntryPath
  $serviceVersion = Get-OpenClawCurrentVersion
  if ([string]::IsNullOrWhiteSpace($entry)) {
    $entry = Join-Path (Get-OpenClawDefaultPrefix) "node_modules\openclaw\openclaw.mjs"
  }
  if ([string]::IsNullOrWhiteSpace($serviceVersion)) {
    $serviceVersion = "2026.3.13"
  }

@"
@echo off
rem OpenClaw Gateway
set "HOME=$($context.UserHome)"
set "TMPDIR=$env:TEMP"
set "OPENCLAW_STATE_DIR=$($context.Home)"
set "OPENCLAW_CONFIG_PATH=$($context.Home)\openclaw.json"
set "OPENCLAW_GATEWAY_PORT=$($context.GatewayPort)"
set "OPENCLAW_WINDOWS_TASK_NAME=$($context.GatewayTask)"
set "OPENCLAW_SERVICE_MARKER=openclaw"
set "OPENCLAW_SERVICE_KIND=gateway"
set "OPENCLAW_SERVICE_VERSION=$serviceVersion"
"$nodeExe" --disable-warning=ExperimentalWarning "$entry" gateway --port $($context.GatewayPort) --force
"@
}

function Get-OpenClawGatewayWatchdogContent {
  $port = Get-OpenClawConfiguredGatewayPort
  $content = @'
$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $baseDir
$pauseFile = Join-Path $baseDir ".paused"
$desiredRunUntilFile = Join-Path $baseDir ".desired-running-until"
$upgradeLockFile = Join-Path $baseDir ".upgrade-lock"

if (Test-Path $pauseFile) {
  exit 0
}

if (Test-Path $upgradeLockFile) {
  exit 0
}

$logDir = Join-Path $env:LOCALAPPDATA "Temp\openclaw"
$logFile = Join-Path $logDir "gateway-watchdog.log"
$maxLogBytes = 1MB
$gatewayTask = "OpenClaw Gateway"
$startupGraceSeconds = 900
$port = __OPENCLAW_PORT__
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

function Get-GatewayProcesses {
  Get-CimInstance Win32_Process |
    Where-Object {
      $_.Name -eq "node.exe" -and (
        $_.CommandLine -like "*node_modules\openclaw\dist\index.js*gateway*" -or
        $_.CommandLine -like "*node_modules\openclaw\openclaw.mjs*gateway*"
      )
    }
}

function Test-DesiredRunActive {
  if (-not (Test-Path $desiredRunUntilFile)) {
    return $false
  }

  try {
    $deadline = [datetime]::Parse((Get-Content $desiredRunUntilFile -Raw).Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
  } catch {
    Write-Log "desired run marker invalid; removing"
    Remove-Item $desiredRunUntilFile -Force -ErrorAction SilentlyContinue
    return $false
  }

  if ($deadline -le (Get-Date)) {
    Write-Log "desired run window expired at $deadline; removing marker"
    Remove-Item $desiredRunUntilFile -Force -ErrorAction SilentlyContinue
    return $false
  }

  return $true
}

function Test-GatewayProbe {
  $client = New-Object System.Net.Sockets.TcpClient
  $stream = $null
  $buffer = New-Object byte[] 4096
  $builder = New-Object System.Text.StringBuilder
  try {
    $async = $client.BeginConnect("127.0.0.1", $port, $null, $null)
    if (-not $async.AsyncWaitHandle.WaitOne(1500, $false)) {
      return $false
    }

    $client.EndConnect($async)
    $stream = $client.GetStream()
    $stream.ReadTimeout = 3000
    $stream.WriteTimeout = 3000

    $requestText = "GET /healthz HTTP/1.1`r`nHost: 127.0.0.1`r`nConnection: close`r`n`r`n"
    $requestBytes = [System.Text.Encoding]::ASCII.GetBytes($requestText)
    $stream.Write($requestBytes, 0, $requestBytes.Length)
    $stream.Flush()

    $deadline = (Get-Date).AddMilliseconds(3500)
    do {
      if (-not $stream.DataAvailable) {
        Start-Sleep -Milliseconds 100
        continue
      }

      $read = $stream.Read($buffer, 0, $buffer.Length)
      if ($read -le 0) {
        break
      }

      [void]$builder.Append([System.Text.Encoding]::UTF8.GetString($buffer, 0, $read))
      $responseText = $builder.ToString()
      if ($responseText -match 'HTTP/1\.[01] 200' -and ($responseText -match '"ok"\s*:\s*true' -or $responseText -match '"status"\s*:\s*"live"')) {
        return $true
      }

      if ($responseText -match 'HTTP/1\.[01] [45]\d\d') {
        return $false
      }
    } while ((Get-Date) -lt $deadline)

    $responseText = $builder.ToString()
    return ($responseText -match 'HTTP/1\.[01] 200' -and ($responseText -match '"ok"\s*:\s*true' -or $responseText -match '"status"\s*:\s*"live"'))
  } catch {
    return $false
  } finally {
    if ($stream) {
      $stream.Dispose()
    }
    $client.Dispose()
  }
}

function Stop-GatewayProcesses {
  param([System.Object[]]$Processes)

  foreach ($proc in @($Processes | Sort-Object ProcessId -Descending)) {
    try {
      Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
      Write-Log "stopped gateway process pid=$($proc.ProcessId) name=$($proc.Name)"
    } catch {
      Write-Log "stop process pid=$($proc.ProcessId) failed: $($_.Exception.Message)"
    }
  }
}

function Test-WithinStartupGrace {
  param([System.Object]$ProcessInfo)

  if ($null -eq $ProcessInfo) {
    return $false
  }

  try {
    $proc = Get-Process -Id $ProcessInfo.ProcessId -ErrorAction Stop
    $startedAt = $proc.StartTime
    return (((Get-Date) - $startedAt).TotalSeconds -lt $startupGraceSeconds)
  } catch {
    return $false
  }
}

function Get-GatewayLockPaths {
  $configPath = Join-Path $baseDir "openclaw.json"
  $sha256 = [System.Security.Cryptography.SHA256]::Create()
  try {
    $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($configPath))
  } finally {
    $sha256.Dispose()
  }

  $hash = ([System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLowerInvariant()).Substring(0, 8)
  $candidates = @(
    (Join-Path (Join-Path $env:LOCALAPPDATA "Temp\openclaw") ("gateway.{0}.lock" -f $hash)),
    (Join-Path (Join-Path $env:TEMP "openclaw") ("gateway.{0}.lock" -f $hash)),
    (Join-Path "C:\tmp\openclaw" ("gateway.{0}.lock" -f $hash))
  )

  return @($candidates | Select-Object -Unique)
}

function Clear-StaleGatewayLocks {
  param([System.Object[]]$GatewayProcesses)

  $activePids = @($GatewayProcesses | Select-Object -ExpandProperty ProcessId)
  foreach ($lockPath in @(Get-GatewayLockPaths)) {
    if (-not (Test-Path $lockPath)) {
      continue
    }

    $payload = $null
    try {
      $payload = Get-Content $lockPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch {}

    $ownerPid = 0
    if ($payload -and $null -ne $payload.pid) {
      [void][int]::TryParse([string]$payload.pid, [ref]$ownerPid)
    }

    $ownerProcess = $null
    if ($ownerPid -gt 0) {
      try {
        $ownerProcess = Get-CimInstance Win32_Process -Filter ("ProcessId = {0}" -f $ownerPid) -ErrorAction Stop
      } catch {
        $ownerProcess = $null
      }
    }

    $ownerIsGateway = ($ownerProcess -and $ownerProcess.Name -eq "node.exe" -and $ownerProcess.CommandLine -like "*openclaw*gateway*")
    $stale = ($ownerPid -le 0 -or -not $ownerProcess -or -not $ownerIsGateway -or ($activePids.Count -gt 0 -and $activePids -notcontains $ownerPid))
    if (-not $stale) {
      continue
    }

    try {
      Remove-Item $lockPath -Force -ErrorAction Stop
      Write-Log "removed stale gateway lock $lockPath ownerPid=$ownerPid"
    } catch {
      Write-Log "remove stale gateway lock failed for ${lockPath}: $($_.Exception.Message)"
    }
  }
}

function Start-GatewayTask {
  try {
    Start-ScheduledTask -TaskName $gatewayTask -ErrorAction Stop
    Write-Log "started scheduled task $gatewayTask"
    return $true
  } catch {
    try {
      & schtasks.exe @("/Run", "/TN", $gatewayTask) | Out-Null
      if ($LASTEXITCODE -eq 0) {
        Write-Log "started scheduled task $gatewayTask through schtasks.exe"
        return $true
      }
    } catch {}

    Write-Log "start scheduled task $gatewayTask failed: $($_.Exception.Message)"
    return $false
  }
}

$mutex = New-Object System.Threading.Mutex($false, "Global\OpenClawGatewayWatchdog")
$hasMutex = $false

try {
  $hasMutex = $mutex.WaitOne(0, $false)
  if (-not $hasMutex) {
    exit 0
  }

  if (-not (Test-DesiredRunActive)) {
    Write-Log "desired run window inactive; no action needed"
    exit 0
  }

  if (Test-GatewayProbe) {
    Write-Log "probe healthy; no action needed"
    exit 0
  }

  $gatewayProcesses = @(Get-GatewayProcesses)
  if ($gatewayProcesses.Count -gt 1) {
    Write-Log "found $($gatewayProcesses.Count) gateway processes; stopping duplicates before restart"
  } elseif ($gatewayProcesses.Count -eq 1) {
    if (Test-WithinStartupGrace -ProcessInfo $gatewayProcesses[0]) {
      Write-Log "gateway pid=$($gatewayProcesses[0].ProcessId) still within startup grace; no action needed"
      exit 0
    }
    Write-Log "probe unhealthy with gateway process pid=$($gatewayProcesses[0].ProcessId); restarting task"
  } else {
    Write-Log "probe unhealthy and gateway process missing; starting task"
  }

  try {
    Stop-ScheduledTask -TaskName $gatewayTask -ErrorAction SilentlyContinue | Out-Null
  } catch {}
  try {
    & schtasks.exe @("/End", "/TN", $gatewayTask) | Out-Null
  } catch {}

  if ($gatewayProcesses.Count -gt 0) {
    Stop-GatewayProcesses -Processes $gatewayProcesses
    Start-Sleep -Seconds 2
  }

  Clear-StaleGatewayLocks -GatewayProcesses @(Get-GatewayProcesses)
  [void](Start-GatewayTask)
}
finally {
  if ($hasMutex) {
    [void]$mutex.ReleaseMutex()
  }
  $mutex.Dispose()
}
'@
  return ($content -replace '__OPENCLAW_PORT__', [string]$port)
}

function Get-OpenClawGatewaySupervisorContent {
@'
$logDir = Join-Path $env:LOCALAPPDATA "Temp\openclaw"
$logFile = Join-Path $logDir "gateway-supervisor.log"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
Add-Content -Path $logFile -Value ("[{0}] supervisor deprecated; watchdog handles recovery directly" -f (Get-Date -Format "yyyy/MM/dd ddd HH:mm:ss.ff"))
exit 0
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
$logDir = Join-Path $env:LOCALAPPDATA "Temp\openclaw"
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
  $proc = Start-Process -FilePath $nodeExe `
    -ArgumentList (@("""$npmCli""") + $NpmArgs) `
    -Wait -PassThru -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutFile `
    -RedirectStandardError $stderrFile
  $exitCode = $proc.ExitCode

  $logLines = New-Object System.Collections.Generic.List[string]
  foreach ($file in @($stdoutFile, $stderrFile)) {
    if (Test-Path $file) {
      Get-Content $file | ForEach-Object {
        $logLines.Add($_)
        Write-Log $_
      }
    }
  }

  if ($exitCode -ne 0) {
    $summary = $logLines |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Last 1
    if ($summary) {
      throw "npm failed ($exitCode): $summary"
    }
    throw "npm failed ($exitCode): $joined"
  }
}

try {
  Set-Content -Path $upgradeLock -Value ("upgrade started at {0}" -f (Get-Date -Format "s")) -Encoding ASCII
  Write-Log "upgrade lock created"

  $runtime = Resolve-NodeRuntime
  $nodeExe = $runtime.NodeExe
  $npmCli = $runtime.NpmCli

  if ([string]::IsNullOrWhiteSpace($TargetVersion)) {
    $TargetVersion = Get-LatestOpenClawVersion
  }
  Write-Log "target version: $TargetVersion"

  Get-CimInstance Win32_Process |
    Where-Object {
      ($_.Name -like "powershell*" -and ($_.CommandLine -like "*gateway-watchdog.ps1*" -or $_.CommandLine -like "*gateway-supervisor.ps1*")) -or
      ($_.Name -eq "cmd.exe" -and $_.CommandLine -like "*gateway.cmd*") -or
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

  $installedPackageJson = Join-Path $targetModule "package.json"
  if (-not (Test-Path $installedPackageJson)) {
    throw "installed package.json not found: $installedPackageJson"
  }

  $installedPackage = Get-Content $installedPackageJson -Raw | ConvertFrom-Json
  $versionText = [string]$installedPackage.version
  Write-Log "installed version: $versionText"
  if ([string]::IsNullOrWhiteSpace($versionText) -or $versionText -notmatch [regex]::Escape($TargetVersion)) {
    throw "installed version mismatch: $versionText"
  }

  $restartAfterUpgrade = $true
} catch {
  Write-Log "upgrade failed: $($_.Exception.Message)"
  throw
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
  Write-OpenClawTextFile -Path $context.WatchdogLauncherVbs -Content (Get-OpenClawWatchdogLauncherVbsContent) -Encoding "ASCII"
  Write-OpenClawTextFile -Path $context.GatewayLauncherVbs -Content (Get-OpenClawGatewayLauncherVbsContent) -Encoding "ASCII"
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

  $watchdogArgument = ('wscript.exe "{0}"' -f $context.WatchdogLauncherVbs)

  try {
    & schtasks.exe @(
      "/Create",
      "/TN", $context.KeepaliveTask,
      "/TR", $watchdogArgument,
      "/SC", "MINUTE",
      "/MO", "1",
      "/F"
    ) | Out-Null

    return ($LASTEXITCODE -eq 0)
  } catch {
    return $false
  }
}

function Ensure-OpenClawGatewayTask {
  $context = Get-OpenClawContext
  Ensure-OpenClawRuntimeFiles

  try {
    & schtasks.exe @(
      "/Create",
      "/TN", $context.GatewayTask,
      "/TR", ('wscript.exe "{0}"' -f $context.GatewayLauncherVbs),
      "/SC", "ONCE",
      "/ST", "00:00",
      "/SD", "01/01/2000",
      "/F"
    ) | Out-Null

    return ($LASTEXITCODE -eq 0)
  } catch {
    return $false
  }
}

function Assert-OpenClawGatewayTask {
  $context = Get-OpenClawContext

  if (-not (Ensure-OpenClawGatewayTask)) {
    throw ("无法创建 Windows 网关任务 {0}。请检查任务计划程序服务与当前用户权限。" -f $context.GatewayTask)
  }

  try {
    $null = Get-ScheduledTask -TaskName $context.GatewayTask -ErrorAction Stop
  } catch {
    throw ("找不到 Windows 网关任务 {0}。请检查任务计划程序服务与当前用户权限。" -f $context.GatewayTask)
  }
}

function Assert-OpenClawKeepaliveTask {
  $context = Get-OpenClawContext

  if (-not (Ensure-OpenClawKeepaliveTask)) {
    throw ("无法创建 Windows 保活任务 {0}。请检查任务计划程序服务与当前用户权限。" -f $context.KeepaliveTask)
  }

  try {
    $null = Get-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction Stop
  } catch {
    throw ("找不到 Windows 保活任务 {0}。请检查任务计划程序服务与当前用户权限。" -f $context.KeepaliveTask)
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

function Get-OpenClawPackageJsonPath {
  $prefix = Get-OpenClawPrefix
  $packageJsonPath = Join-Path $prefix "node_modules/openclaw/package.json"
  if (Test-Path $packageJsonPath) {
    return $packageJsonPath
  }

  return $null
}

function Get-OpenClawCurrentVersion {
  $packageJsonPath = Get-OpenClawPackageJsonPath
  if ($packageJsonPath) {
    try {
      $packageJson = Get-Content $packageJsonPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      if ($packageJson.version) {
        return [string]$packageJson.version
      }
    } catch {}
  }

  $cmd = Get-OpenClawCommandPath
  if ($cmd) {
    try {
      $raw = & $cmd --version 2>$null | Out-String
      if ($raw -match '(\d+\.\d+\.\d+)') {
        return $matches[1]
      }
    } catch {
      return $null
    }
  }

  return $null
}

function Get-OpenClawTaskState {
  param(
    [string]$TaskName
  )

  if ([string]::IsNullOrWhiteSpace($TaskName)) {
    return $null
  }

  try {
    $raw = & schtasks.exe @("/Query", "/TN", $TaskName, "/FO", "CSV", "/NH") 2>$null | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($raw)) {
      return $null
    }

    $record = $raw | ConvertFrom-Csv -Header "TaskName", "NextRunTime", "Status" | Select-Object -First 1
    if ($record) {
      return [pscustomobject]@{
        TaskName = $TaskName
        State = ([string]$record.Status).Trim()
        NextRunTime = ([string]$record.NextRunTime).Trim()
      }
    }
  } catch {}

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
  $client = New-Object System.Net.Sockets.TcpClient
  $stream = $null
  $buffer = New-Object byte[] 4096
  $builder = New-Object System.Text.StringBuilder
  try {
    $async = $client.BeginConnect("127.0.0.1", $context.GatewayPort, $null, $null)
    if (-not $async.AsyncWaitHandle.WaitOne(1500, $false)) {
      return $false
    }

    $client.EndConnect($async)
    $stream = $client.GetStream()
    $stream.ReadTimeout = 3000
    $stream.WriteTimeout = 3000

    $requestText = "GET /healthz HTTP/1.1`r`nHost: 127.0.0.1`r`nConnection: close`r`n`r`n"
    $requestBytes = [System.Text.Encoding]::ASCII.GetBytes($requestText)
    $stream.Write($requestBytes, 0, $requestBytes.Length)
    $stream.Flush()

    $deadline = (Get-Date).AddMilliseconds(3500)
    do {
      if (-not $stream.DataAvailable) {
        Start-Sleep -Milliseconds 100
        continue
      }

      $read = $stream.Read($buffer, 0, $buffer.Length)
      if ($read -le 0) {
        break
      }

      [void]$builder.Append([System.Text.Encoding]::UTF8.GetString($buffer, 0, $read))
      $responseText = $builder.ToString()
      if ($responseText -match 'HTTP/1\.[01] 200' -and ($responseText -match '"ok"\s*:\s*true' -or $responseText -match '"status"\s*:\s*"live"')) {
        return $true
      }

      if ($responseText -match 'HTTP/1\.[01] [45]\d\d') {
        return $false
      }
    } while ((Get-Date) -lt $deadline)

    $responseText = $builder.ToString()
    return ($responseText -match 'HTTP/1\.[01] 200' -and ($responseText -match '"ok"\s*:\s*true' -or $responseText -match '"status"\s*:\s*"live"'))
  } catch {
    return $false
  } finally {
    if ($stream) {
      $stream.Dispose()
    }
    $client.Dispose()
  }
}

function Test-OpenClawGatewayPortListening {
  $context = Get-OpenClawContext
  $client = New-Object System.Net.Sockets.TcpClient
  try {
    $async = $client.BeginConnect("127.0.0.1", $context.GatewayPort, $null, $null)
    if (-not $async.AsyncWaitHandle.WaitOne(1500, $false)) {
      return $false
    }

    $client.EndConnect($async)
    return $true
  } catch {
    return $false
  } finally {
    $client.Dispose()
  }
}

function Get-OpenClawProcesses {
  Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" |
    Where-Object {
      ($_.Name -eq "node.exe" -and $_.CommandLine -like "*openclaw*gateway*" -and $_.CommandLine -notlike "*--help*")
    }
}

function Get-OpenClawNewestGatewayProcess {
  return @(Get-OpenClawProcesses |
    Sort-Object { Convert-OpenClawCimDateTime $_.CreationDate } -Descending |
    Select-Object -First 1)
}

function Test-OpenClawGatewayStartupGrace {
  param(
    [int]$GraceSeconds = 900
  )

  $process = @(Get-OpenClawNewestGatewayProcess | Select-Object -First 1)
  if ($process.Count -eq 0) {
    return $false
  }

  $startedAt = Convert-OpenClawCimDateTime $process[0].CreationDate
  if ($null -eq $startedAt) {
    return $false
  }

  return (((Get-Date) - $startedAt).TotalSeconds -lt $GraceSeconds)
}

function Clear-OpenClawStaleGatewayLocks {
  param(
    [switch]$OnlyIfUnhealthy
  )

  $healthy = Test-OpenClawGatewayHealthy
  if ($OnlyIfUnhealthy -and $healthy) {
    return @()
  }

  $activeGatewayPids = @((Get-OpenClawProcesses | Select-Object -ExpandProperty ProcessId))
  $removed = New-Object System.Collections.Generic.List[string]

  foreach ($lockPath in @(Get-OpenClawGatewayLockPaths)) {
    if (-not (Test-Path $lockPath)) {
      continue
    }

    $payload = $null
    try {
      $payload = Get-Content $lockPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch {}

    $ownerPid = 0
    if ($payload -and $null -ne $payload.pid) {
      [void][int]::TryParse([string]$payload.pid, [ref]$ownerPid)
    }

    $ownerProcess = $null
    if ($ownerPid -gt 0) {
      try {
        $ownerProcess = Get-CimInstance Win32_Process -Filter ("ProcessId = {0}" -f $ownerPid) -ErrorAction Stop
      } catch {
        $ownerProcess = $null
      }
    }

    $ownerIsGateway = ($ownerProcess -and $ownerProcess.Name -eq "node.exe" -and $ownerProcess.CommandLine -like "*openclaw*gateway*")
    $shouldRemove = $false

    if ($ownerPid -le 0) {
      $shouldRemove = $true
    } elseif (-not $ownerProcess) {
      $shouldRemove = $true
    } elseif (-not $ownerIsGateway) {
      $shouldRemove = $true
    } elseif ($activeGatewayPids.Count -gt 0 -and $activeGatewayPids -notcontains $ownerPid) {
      $shouldRemove = $true
    }

    if ($shouldRemove) {
      try {
        Remove-Item $lockPath -Force -ErrorAction Stop
        $removed.Add($lockPath) | Out-Null
      } catch {}
    }
  }

  return @($removed)
}

function Get-OpenClawGatewayLogFiles {
  $context = Get-OpenClawContext
  $files = @()

  foreach ($dir in @($context.GatewayLogDirs)) {
    if (-not [string]::IsNullOrWhiteSpace($dir) -and (Test-Path $dir)) {
      $files += Get-ChildItem $dir -Filter *.log -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 5
    }
  }

  return @($files | Sort-Object FullName -Unique)
}

function Get-OpenClawLatestFailure {
  $context = Get-OpenClawContext
  $files = @()
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

  $files += Get-OpenClawGatewayLogFiles
  $patterns = @(
    'Gateway failed to start:',
    'gateway already running',
    'lock timeout after \d+ms',
    'gateway deemed unhealthy',
    'supervisor loop exception',
    'gateway child exited .* code=',
    'upgrade failed:',
    'Cannot find module',
    'another gateway instance is already listening',
    'failed to bind',
    'Unable to resolve node.exe'
  )

  foreach ($file in ($files | Sort-Object -Property LastWriteTime, FullName -Descending -Unique)) {
    $lines = @(Get-Content $file.FullName -Tail 160 -ErrorAction SilentlyContinue)
    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
      $line = [string]$lines[$i]
      foreach ($pattern in $patterns) {
        if ($line -imatch $pattern) {
          return [pscustomobject]@{
            File = $file.Name
            Path = $file.FullName
            LastWriteTime = $file.LastWriteTime
            Message = $line.Trim()
          }
        }
      }
    }
  }

  return $null
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
    Assert-OpenClawGatewayTask
    Assert-OpenClawKeepaliveTask
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

  $files += Get-OpenClawGatewayLogFiles | Select-Object -First 3

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
  $gatewayTask = $null
  $backups = @(Get-OpenClawBackups)
  $commandPath = Get-OpenClawCommandPath
  $homeExists = Test-Path $context.Home

  $keepaliveTask = Get-OpenClawTaskState -TaskName $context.KeepaliveTask
  $gatewayTask = Get-OpenClawTaskState -TaskName $context.GatewayTask

  $paused = Test-Path $context.PauseFile
  $upgradeLocked = Test-Path $context.UpgradeLockFile
  $desiredRunActive = Test-OpenClawDesiredRunActive
  $processes = @(Get-OpenClawProcesses)
  $newestProcess = @(Get-OpenClawNewestGatewayProcess | Select-Object -First 1)
  $portListening = $false

  $currentVersion = Get-OpenClawCurrentVersion
  $latestVersion = if ($IncludeLatestVersion) { Get-OpenClawLatestVersion } else { $null }
  $healthy = $false
  $recentFailure = $null
  $withinStartupGrace = $false

  if (-not $paused -and -not $upgradeLocked -and $processes.Count -gt 0) {
    $portListening = Test-OpenClawGatewayPortListening
    $healthy = Test-OpenClawGatewayHealthy
  }

  if (-not $healthy) {
    $recentFailure = Get-OpenClawLatestFailure
  }

  if ($processes.Count -gt 0 -or ($gatewayTask -and [string]$gatewayTask.State -eq "Running")) {
    $withinStartupGrace = Test-OpenClawGatewayStartupGrace
  }

  if ($recentFailure -and (((Get-Date) - $recentFailure.LastWriteTime).TotalMinutes -gt 20)) {
    $recentFailure = $null
  }

  if ($recentFailure -and $newestProcess.Count -gt 0) {
    $newestStartedAt = Convert-OpenClawCimDateTime $newestProcess[0].CreationDate
    if ($newestStartedAt -and $recentFailure.LastWriteTime -lt $newestStartedAt) {
      $recentFailure = $null
    }
  }

  $mode = "Stopped"
  if ($upgradeLocked) {
    $mode = "Updating"
  } elseif ($paused) {
    $mode = "Paused"
  } elseif ($healthy) {
    $mode = "Running"
  } elseif ($processes.Count -gt 0 -and -not $withinStartupGrace) {
    $mode = "Failed"
  } elseif ($recentFailure) {
    $mode = "Failed"
  } elseif ($processes.Count -gt 0 -or ($gatewayTask -and [string]$gatewayTask.State -eq "Running") -or $desiredRunActive) {
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
    GatewayTaskState = if ($gatewayTask) { [string]$gatewayTask.State } else { "Missing" }
    AutoStartEnabled = Get-OpenClawAutoStartEnabled
    IsInstalled = -not [string]::IsNullOrWhiteSpace($commandPath)
    HasExistingConfig = $homeExists
    DesiredRunActive = $desiredRunActive
    ProcessCount = $processes.Count
    GatewayPort = $context.GatewayPort
    GatewayPortListening = $portListening
    FailureReason = if ($recentFailure) { [string]$recentFailure.Message } else { $null }
    Prefix = Get-OpenClawPrefix
    CommandPath = $commandPath
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
  Get-CimInstance Win32_Process -Filter "Name = 'wscript.exe' OR Name = 'powershell.exe' OR Name = 'pwsh.exe'" |
    Where-Object {
      ($_.Name -eq "wscript.exe" -and $_.CommandLine -like "*run-hidden.vbs*") -or
      ($_.Name -like "powershell*" -and ($_.CommandLine -like "*gateway-watchdog.ps1*" -or $_.CommandLine -like "*gateway-supervisor.ps1*"))
    } |
    ForEach-Object {
      Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

  Write-OpenClawNpmConfig
  Ensure-OpenClawPrefixPath
  Assert-OpenClawGatewayTask
  Assert-OpenClawKeepaliveTask
  [void](Set-OpenClawDesiredRunWindow)

  try {
    Enable-ScheduledTask -TaskName $context.GatewayTask -ErrorAction SilentlyContinue | Out-Null
  } catch {}

  try {
    Enable-ScheduledTask -TaskName $context.KeepaliveTask -ErrorAction SilentlyContinue | Out-Null
  } catch {}

  [void](Clear-OpenClawStaleGatewayLocks -OnlyIfUnhealthy)
  try {
    Start-ScheduledTask -TaskName $context.GatewayTask -ErrorAction SilentlyContinue
  } catch {}

  $deadline = (Get-Date).AddSeconds(20)
  do {
    Start-Sleep -Seconds 1
    $status = Get-OpenClawStatus
    if ($status.GatewayHealthy -or $status.Mode -eq "Failed") {
      return $status
    }
  } while ((Get-Date) -lt $deadline)

  return $status
}

function Stop-OpenClawSilent {
  $context = Get-OpenClawContext
  New-Item -ItemType Directory -Path $context.Home -Force | Out-Null
  Set-Content -Path $context.PauseFile -Value ("paused at {0}" -f (Get-Date -Format "s")) -Encoding ASCII
  Clear-OpenClawDesiredRunWindow

  try {
    Stop-ScheduledTask -TaskName $context.GatewayTask -ErrorAction SilentlyContinue | Out-Null
  } catch {}
  try {
    & schtasks.exe @("/End", "/TN", $context.GatewayTask) | Out-Null
  } catch {}

  Get-OpenClawProcesses | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
  }
  Get-CimInstance Win32_Process -Filter "Name = 'wscript.exe' OR Name = 'powershell.exe' OR Name = 'pwsh.exe'" |
    Where-Object {
      ($_.Name -eq "wscript.exe" -and $_.CommandLine -like "*run-hidden.vbs*") -or
      ($_.Name -like "powershell*" -and ($_.CommandLine -like "*gateway-watchdog.ps1*" -or $_.CommandLine -like "*gateway-supervisor.ps1*"))
    } |
    ForEach-Object {
      Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

  [void](Clear-OpenClawStaleGatewayLocks)
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

  $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $args -WindowStyle Hidden -PassThru
  $deadline = (Get-Date).AddSeconds(30)

  while ((Get-Date) -lt $deadline) {
    if (Test-Path $context.UpgradeLockFile) {
      return $proc
    }

    try {
      $proc.Refresh()
    } catch {}

    if ($proc.HasExited) {
      $message = "OpenClaw 安装/升级进程启动失败。"
      if (Test-Path $context.UpgradeLog) {
        $tail = Get-Content $context.UpgradeLog -Tail 20 -ErrorAction SilentlyContinue |
          Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
          Select-Object -Last 1
        if ($tail) {
          $message = "$message 最近日志：$tail"
        }
      }
      throw $message
    }

    Start-Sleep -Milliseconds 250
  }

  if (-not (Test-Path $context.UpgradeLockFile)) {
    throw "OpenClaw 安装/升级进程未进入执行状态。"
  }

  return $proc
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
      ".desired-running-until",
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
          "watchdog-launcher.vbs",
          "gateway-launcher.vbs",
          "gateway.cmd",
          "upgrade-openclaw.ps1",
          ".paused",
          ".desired-running-until",
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
  [void](Ensure-OpenClawGatewayTask)
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

  $copyTargets += Get-OpenClawGatewayLogFiles

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

