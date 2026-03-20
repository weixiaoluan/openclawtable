[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

trap {
  $message = $_.Exception.Message
  $line = $null
  try {
    if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) {
      $line = [string]$_.InvocationInfo.ScriptLineNumber
    }
  } catch {}

  $details = if ([string]::IsNullOrWhiteSpace($line)) {
    "启动失败：$message"
  } else {
    "启动失败：$message`n行号：$line"
  }

  try {
    [System.Windows.Forms.MessageBox]::Show($details, "OpenClaw 控制中心", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  } catch {
    Write-Error $details
  }

  exit 1
}

$locale = @'
{
  "window_title": "OpenClaw \u63a7\u5236\u4e2d\u5fc3",
  "missing_toolkit_title": "\u542f\u52a8\u5931\u8d25",
  "missing_toolkit_message": "\u7f3a\u5c11\u5de5\u5177\u811a\u672c\uff1a\n{0}",
  "app_name": "OpenClaw \u63a7\u5236\u4e2d\u5fc3",
  "sidebar_intro": "\u4e00\u4e2a\u5b89\u9759\u7684\u63a7\u5236\u9762\u677f\uff0c\u7528\u4e8e 24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c\u3001\u68c0\u67e5\u66f4\u65b0\u3001\u4e00\u952e\u5907\u4efd\u3001\u4e00\u952e\u6062\u590d\u3001\u5bfc\u51fa\u65e5\u5fd7\u548c\u5f3a\u5236\u9000\u51fa\u3002\u5e95\u5c42\u4ecd\u6cbf\u7528 watchdog \u4e0e supervisor \u94fe\u8def\u3002",
  "runtime_model_label": "\u8fd0\u884c\u6a21\u5f0f",
  "runtime_model_text": "24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c\u4f1a\u6cbf\u7528\u5f53\u524d\u4fdd\u6d3b\u903b\u8f91\u3002\u5f3a\u5236\u9000\u51fa\u4f1a\u5199\u5165 .paused\uff0c\u5e76\u6e05\u6389\u5f53\u524d\u540e\u53f0\u94fe\u8def\uff0c\u76f4\u5230\u4f60\u518d\u6b21\u542f\u52a8\u3002",
  "headline": "\u9759\u9ed8\u540e\u53f0\u63a7\u5236\u53f0",
  "subheadline": "\u5728\u4e00\u4e2a\u9762\u677f\u4e2d\u67e5\u770b\u7248\u672c\u3001\u7f51\u5173\u72b6\u6001\u3001\u5f00\u673a\u81ea\u542f\u3001\u5907\u4efd\u4e0e\u6062\u590d\u3001\u9519\u8bef\u6458\u8981\u548c\u65e5\u5fd7\u5bfc\u51fa\u3002",
  "current_mode_label": "\u5f53\u524d\u72b6\u6001",
  "auto_start_label": "\u5f00\u673a\u81ea\u542f",
  "current_version_label": "\u5f53\u524d\u7248\u672c",
  "latest_version_label": "\u6700\u65b0\u7248\u672c",
  "gateway_health_label": "\u7f51\u5173\u5065\u5eb7",
  "keepalive_label": "\u4fdd\u6d3b\u4efb\u52a1",
  "button_start_silent": "\u542f\u52a8 24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c",
  "button_enable_auto_start": "\u5f00\u542f\u5f00\u673a\u81ea\u542f",
  "button_disable_auto_start": "\u5173\u95ed\u5f00\u673a\u81ea\u542f",
  "button_check_updates": "\u68c0\u67e5\u66f4\u65b0",
  "button_create_backup": "\u4e00\u952e\u5907\u4efd",
  "button_restore_backup": "\u4e00\u952e\u6062\u590d",
  "button_export_logs": "\u5bfc\u51fa\u62a5\u9519\u65e5\u5fd7",
  "button_force_stop": "\u5f3a\u5236\u9000\u51fa\u9759\u9ed8\u8fd0\u884c",
  "button_refresh": "\u5237\u65b0\u72b6\u6001",
  "section_feedback_title": "\u72b6\u6001\u4e0e\u64cd\u4f5c\u53cd\u9988",
  "section_feedback_subtitle": "\u8fd9\u91cc\u4f1a\u663e\u793a\u6700\u8fd1\u4e00\u6b21\u68c0\u67e5\u3001\u5f53\u524d\u8fd0\u884c\u72b6\u6001\u548c\u6700\u65b0\u64cd\u4f5c\u7ed3\u679c\u3002",
  "busy_ready": "\u5c31\u7eea",
  "operation_default": "\u4f7f\u7528\u4e0a\u65b9\u6309\u94ae\u5373\u53ef\u542f\u52a8 24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c\u3001\u68c0\u67e5\u66f4\u65b0\u3001\u521b\u5efa\u8fc1\u79fb\u5907\u4efd\u3001\u9009\u62e9\u5386\u53f2\u5907\u4efd\u6062\u590d\u6216\u5f3a\u5236\u9000\u51fa\u540e\u53f0\u94fe\u8def\u3002",
  "section_logs_title": "\u65e5\u5fd7\u4e0e\u5bfc\u51fa",
  "label_watchdog": "\u770b\u95e8\u72d7\u65e5\u5fd7",
  "label_supervisor": "\u76d1\u7763\u5668\u65e5\u5fd7",
  "label_latest_audit": "\u6700\u8fd1\u5de1\u68c0",
  "section_runtime_title": "\u8fd0\u884c\u4fe1\u606f",
  "label_command": "OpenClaw \u547d\u4ee4",
  "label_prefix": "\u5b89\u88c5\u524d\u7f00",
  "label_gateway_log_dir": "\u7f51\u5173\u65e5\u5fd7\u76ee\u5f55",
  "section_recent_errors_title": "\u6700\u8fd1\u62a5\u9519\u6458\u8981",
  "section_recent_errors_subtitle": "\u5c55\u793a watchdog\u3001supervisor\u3001gateway \u548c audit \u4e2d\u6700\u8fd1\u547d\u4e2d\u7684\u5f02\u5e38\u884c\u3002",
  "status_not_installed": "\u672a\u5b89\u88c5",
  "status_click_check": "\u70b9\u51fb\u68c0\u67e5",
  "status_live": "\u6b63\u5e38",
  "status_waiting": "\u542f\u52a8\u4e2d",
  "status_failed_gateway": "\u5f02\u5e38",
  "status_stopped_gateway": "\u5df2\u505c\u6b62",
  "status_enabled": "\u5df2\u5f00\u542f",
  "status_disabled": "\u5df2\u5173\u95ed",
  "mode_running": "\u8fd0\u884c\u4e2d",
  "mode_starting": "\u542f\u52a8\u4e2d",
  "mode_failed": "\u542f\u52a8\u5931\u8d25",
  "mode_paused": "\u5df2\u6682\u505c",
  "mode_updating": "\u66f4\u65b0\u4e2d",
  "mode_stopped": "\u5df2\u505c\u6b62",
  "task_ready": "\u5c31\u7eea",
  "task_running": "\u8fd0\u884c\u4e2d",
  "task_disabled": "\u5df2\u7981\u7528",
  "task_queued": "\u961f\u5217\u4e2d",
  "task_missing": "\u7f3a\u5931",
  "msg_updating": "OpenClaw \u6b63\u5728\u66f4\u65b0\u3002\u66f4\u65b0\u671f\u95f4\u4fdd\u6d3b\u94fe\u8def\u4f1a\u81ea\u52a8\u6682\u505c\uff0c\u5b8c\u6210\u540e\u518d\u81ea\u52a8\u6062\u590d\u3002",
  "busy_updating": "\u66f4\u65b0\u4e2d",
  "msg_paused": "OpenClaw \u5f53\u524d\u5904\u4e8e\u6682\u505c\u72b6\u6001\u3002\u70b9\u51fb\u201c\u542f\u52a8 24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c\u201d\u5373\u53ef\u79fb\u9664 .paused \u5e76\u6062\u590d\u540e\u53f0\u8fd0\u884c\u3002",
  "busy_paused": "\u5df2\u6682\u505c",
  "msg_healthy": "OpenClaw \u8fd0\u884c\u6b63\u5e38\u3002\u4fdd\u6d3b\u4efb\u52a1\u4f1a\u5728\u540e\u53f0\u6301\u7eed\u68c0\u6d4b\u7f51\u5173\u5065\u5eb7\u5e76\u5728\u5fc5\u8981\u65f6\u81ea\u52a8\u6062\u590d\u3002",
  "busy_healthy": "\u8fd0\u884c\u6b63\u5e38",
  "msg_starting": "\u7f51\u5173\u4ecd\u5728\u542f\u52a8\u6216\u6062\u590d\u4e2d\u3002\u9759\u9ed8\u94fe\u8def\u5df2\u7ecf\u63a5\u7ba1\uff0c\u8bf7\u7a0d\u7b49\u5065\u5eb7\u72b6\u6001\u5207\u6362\u4e3a\u201c\u6b63\u5e38\u201d\u3002",
  "busy_starting": "\u542f\u52a8\u4e2d",
  "msg_failed_template": "OpenClaw \u542f\u52a8\u5931\u8d25\uff1a{0}",
  "msg_failed_generic": "\u540e\u53f0\u94fe\u8def\u672a\u80fd\u901a\u8fc7\u5065\u5eb7\u68c0\u67e5\uff0c\u8bf7\u5148\u5904\u7406\u62a5\u9519\u6458\u8981\u540e\u518d\u91cd\u8bd5\u3002",
  "busy_failed": "\u542f\u52a8\u5931\u8d25",
  "msg_refreshing": "\u6b63\u5728\u5237\u65b0\u72b6\u6001...",
  "busy_refreshing": "\u5237\u65b0\u4e2d",
  "msg_refresh_failed_prefix": "\u5237\u65b0\u5931\u8d25\uff1a",
  "msg_starting_silent": "\u6b63\u5728\u542f\u52a8 24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c...",
  "busy_starting_action": "\u542f\u52a8\u4e2d",
  "msg_start_triggered_template": "\u5df2\u89e6\u53d1\u9759\u9ed8\u8fd0\u884c\u94fe\u8def\u3002\u5f53\u524d\u72b6\u6001\uff1a{0}",
  "busy_triggered": "\u5df2\u89e6\u53d1",
  "msg_start_failed_prefix": "\u542f\u52a8\u5931\u8d25\uff1a",
  "dialog_force_stop_title": "\u5f3a\u5236\u9000\u51fa\u9759\u9ed8\u8fd0\u884c",
  "dialog_force_stop_message": "\u8fd9\u4f1a\u505c\u6b62\u5f53\u524d\u9759\u9ed8\u540e\u53f0\u94fe\u8def\uff0c\u5199\u5165 .paused\uff0c\u5e76\u7ed3\u675f\u73b0\u6709\u7684 OpenClaw \u540e\u53f0\u8fdb\u7a0b\u3002\u786e\u5b9a\u7ee7\u7eed\u5417\uff1f",
  "msg_stopping_silent": "\u6b63\u5728\u505c\u6b62\u9759\u9ed8\u540e\u53f0\u94fe\u8def...",
  "busy_stopping": "\u505c\u6b62\u4e2d",
  "msg_stopped_template": "\u9759\u9ed8\u8fd0\u884c\u5df2\u505c\u6b62\u3002\u5f53\u524d\u72b6\u6001\uff1a{0}",
  "busy_stopped": "\u5df2\u505c\u6b62",
  "msg_stop_failed_prefix": "\u505c\u6b62\u5931\u8d25\uff1a",
  "busy_stop_failed": "\u505c\u6b62\u5931\u8d25",
  "msg_checking_updates": "\u6b63\u5728\u68c0\u67e5 OpenClaw \u5b98\u65b9\u6700\u65b0\u7248\u672c...",
  "busy_checking": "\u68c0\u67e5\u4e2d",
  "msg_latest_version_unavailable": "\u6682\u65f6\u65e0\u6cd5\u83b7\u53d6\u5b98\u65b9\u6700\u65b0\u7248\u672c\uff0c\u8bf7\u68c0\u67e5\u7f51\u7edc\u540e\u518d\u8bd5\u3002",
  "busy_check_failed": "\u68c0\u67e5\u5931\u8d25",
  "msg_up_to_date_template": "\u5f53\u524d\u5df2\u662f\u6700\u65b0\u7248\u672c\uff1a{0}",
  "busy_up_to_date": "\u5df2\u6700\u65b0",
  "dialog_update_title": "OpenClaw \u66f4\u65b0",
  "dialog_update_prompt_template": "\u53d1\u73b0\u65b0\u7248\u672c\uff1a{0}\n\u5f53\u524d\u7248\u672c\uff1a{1}\n\u73b0\u5728\u5f00\u59cb\u66f4\u65b0\u5417\uff1f",
  "dialog_install_existing_prompt_template": "\u68c0\u6d4b\u5230\u672c\u5730\u5b58\u5728\u65e7\u914d\u7f6e\u6216\u5b89\u88c5\u6b8b\u7559\u3002\n\u5c06\u4fdd\u7559\u73b0\u6709\u914d\u7f6e\u5e76\u5b89\u88c5\u6700\u65b0\u7248\u672c\uff1a{0}\n\u73b0\u5728\u7ee7\u7eed\u5417\uff1f",
  "msg_update_started_template": "\u5df2\u5728\u540e\u53f0\u542f\u52a8\u66f4\u65b0\uff0c\u76ee\u6807\u7248\u672c\uff1a{0}\u3002\u66f4\u65b0\u9501\u4f1a\u5728\u8fc7\u7a0b\u4e2d\u9632\u6b62\u9759\u9ed8\u4fdd\u6d3b\u94fe\u8def\u6b7b\u5faa\u73af\u3002",
  "msg_install_started_template": "\u5df2\u5728\u540e\u53f0\u542f\u52a8 OpenClaw \u5b89\u88c5\uff0c\u76ee\u6807\u7248\u672c\uff1a{0}\u3002",
  "busy_update_pid_template": "\u66f4\u65b0 PID {0}",
  "msg_update_canceled": "\u5df2\u53d6\u6d88\u66f4\u65b0\u3002",
  "busy_canceled": "\u5df2\u53d6\u6d88",
  "msg_update_check_failed_prefix": "\u66f4\u65b0\u68c0\u67e5\u5931\u8d25\uff1a",
  "msg_collecting_logs": "\u6b63\u5728\u6536\u96c6\u62a5\u9519\u65e5\u5fd7...",
  "busy_exporting": "\u5bfc\u51fa\u4e2d",
  "msg_creating_backup": "\u6b63\u5728\u521b\u5efa OpenClaw \u8fc1\u79fb\u5907\u4efd...",
  "busy_backing_up": "\u5907\u4efd\u4e2d",
  "busy_backed_up": "\u5df2\u5907\u4efd",
  "msg_backup_created_template": "\u5907\u4efd\u5df2\u521b\u5efa\uff1a\n{0}",
  "msg_backup_failed_prefix": "\u5907\u4efd\u5931\u8d25\uff1a",
  "msg_backup_name_canceled": "\u5df2\u53d6\u6d88\u521b\u5efa\u5907\u4efd\u3002",
  "dialog_backup_name_title": "\u521b\u5efa OpenClaw \u5907\u4efd",
  "dialog_backup_name_prompt": "\u8bf7\u8f93\u5165\u8fd9\u6b21\u5907\u4efd\u7684\u540d\u79f0\uff08\u53ef\u4ee5\u81ea\u5b9a\u4e49\uff09\uff1a",
  "backup_default_name_template": "\u6211\u7684 OpenClaw \u5907\u4efd-{0}",
  "restore_filter": "OpenClaw \u5907\u4efd (*.ocbackup.zip;*.zip)|*.ocbackup.zip;*.zip",
  "dialog_restore_title": "\u6062\u590d OpenClaw \u5907\u4efd",
  "dialog_restore_prompt_template": "\u786e\u5b9a\u8981\u6062\u590d\u8fd9\u4e2a\u5907\u4efd\u5417\uff1f\n{0}\n\n\u6062\u590d\u4f1a\u8986\u76d6\u5f53\u524d OpenClaw \u914d\u7f6e\u548c\u5b89\u88c5\uff0c\u5e76\u5148\u505c\u6b62\u540e\u53f0\u9759\u9ed8\u8fd0\u884c\u3002",
  "msg_restore_canceled": "\u5df2\u53d6\u6d88\u6062\u590d\u3002",
  "msg_restoring_backup": "\u6b63\u5728\u6062\u590d OpenClaw \u5907\u4efd...",
  "busy_restoring": "\u6062\u590d\u4e2d",
  "busy_restored": "\u5df2\u6062\u590d",
  "msg_restore_completed_template": "\u5907\u4efd\u5df2\u6062\u590d\uff1a\n{0}\n\n\u5f53\u524d\u4f1a\u4fdd\u6301\u6682\u505c\u72b6\u6001\uff0c\u4f60\u53ef\u4ee5\u5728\u9762\u677f\u4e2d\u518d\u70b9\u201c\u542f\u52a8 24 \u5c0f\u65f6\u9759\u9ed8\u8fd0\u884c\u201d\u3002",
  "msg_restore_failed_prefix": "\u6062\u590d\u5931\u8d25\uff1a",
  "export_filename_template": "OpenClaw-\u62a5\u9519\u65e5\u5fd7-{0}.zip",
  "export_filter": "Zip \u538b\u7f29\u5305 (*.zip)|*.zip",
  "msg_log_export_canceled": "\u5df2\u53d6\u6d88\u5bfc\u51fa\u3002",
  "msg_exported_template": "\u62a5\u9519\u65e5\u5fd7\u5df2\u5bfc\u51fa\u5230\uff1a\n{0}",
  "busy_exported": "\u5df2\u5bfc\u51fa",
  "msg_export_failed_prefix": "\u5bfc\u51fa\u5931\u8d25\uff1a",
  "busy_export_failed": "\u5bfc\u51fa\u5931\u8d25",
  "msg_autostart_enabling": "\u6b63\u5728\u5f00\u542f\u5f00\u673a\u81ea\u542f...",
  "msg_autostart_disabling": "\u6b63\u5728\u5173\u95ed\u5f00\u673a\u81ea\u542f...",
  "msg_autostart_enabled": "\u5f00\u673a\u81ea\u542f\u5df2\u5f00\u542f\u3002\u4f60\u4e0b\u6b21\u767b\u5f55 Windows \u540e\uff0cOpenClaw \u4f1a\u6309\u540c\u4e00\u5957\u9759\u9ed8\u94fe\u8def\u81ea\u52a8\u542f\u52a8\u3002",
  "msg_autostart_disabled": "\u5f00\u673a\u81ea\u542f\u5df2\u5173\u95ed\u3002\u5f53\u524d\u8fd0\u884c\u72b6\u6001\u4e0d\u53d8\uff0c\u4f46\u4ee5\u540e\u767b\u5f55 Windows \u65f6\u4e0d\u4f1a\u518d\u81ea\u52a8\u542f\u52a8 OpenClaw\u3002",
  "busy_updated": "\u5df2\u66f4\u65b0",
  "msg_autostart_failed_prefix": "\u81ea\u542f\u8bbe\u7f6e\u5931\u8d25\uff1a"
}
'@ | ConvertFrom-Json

function Get-LocText {
  param([string]$Key)

  if ($locale.PSObject.Properties.Name -contains $Key) {
    return [string]$locale.$Key
  }

  return $Key
}

function Format-Loc {
  param(
    [string]$Key,
    [object[]]$FormatArgs
  )

  if ($null -eq $FormatArgs) {
    $FormatArgs = @()
  }

  return [string]::Format([System.Globalization.CultureInfo]::CurrentCulture, (Get-LocText $Key), ([object[]]$FormatArgs))
}

$embeddedToolkitBase64 = "__EMBEDDED_TOOLKIT_BASE64__"

function Get-PreferredDirectoryPath {
  param(
    [string[]]$Candidates
  )

  foreach ($candidate in @($Candidates)) {
    if ([string]::IsNullOrWhiteSpace($candidate)) {
      continue
    }

    try {
      $resolved = [IO.Path]::GetFullPath($candidate)
      if (-not [string]::IsNullOrWhiteSpace($resolved)) {
        return $resolved.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
      }
    } catch {
      continue
    }
  }

  return $null
}

function Get-ControlCenterUserHome {
  return Get-PreferredDirectoryPath @(
    [Environment]::GetFolderPath("UserProfile"),
    $env:USERPROFILE,
    $env:HOMEDRIVE + $env:HOMEPATH
  )
}

function Get-ControlCenterDesktopDirectory {
  return Get-PreferredDirectoryPath @(
    [Environment]::GetFolderPath("Desktop"),
    $(if (-not [string]::IsNullOrWhiteSpace((Get-ControlCenterUserHome))) { Join-Path (Get-ControlCenterUserHome) "Desktop" } else { $null }),
    (Get-ControlCenterBaseDirectory)
  )
}

function Get-ControlCenterCacheDirectory {
  return Get-PreferredDirectoryPath @(
    $(if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { Join-Path $env:LOCALAPPDATA "OpenClawControlCenter" } else { $null }),
    $(if (-not [string]::IsNullOrWhiteSpace($env:TEMP)) { Join-Path $env:TEMP "OpenClawControlCenter" } else { $null }),
    $(if (-not [string]::IsNullOrWhiteSpace((Get-ControlCenterBaseDirectory))) { Join-Path (Get-ControlCenterBaseDirectory) ".cache" } else { $null })
  )
}

function Get-ControlCenterBaseDirectory {
  $candidates = @(
    [System.AppContext]::BaseDirectory,
    [System.AppDomain]::CurrentDomain.BaseDirectory,
    $(if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { $null }),
    $PSScriptRoot
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

  foreach ($candidate in $candidates) {
    try {
      $resolved = [IO.Path]::GetFullPath($candidate)
      if (-not [string]::IsNullOrWhiteSpace($resolved)) {
        return $resolved.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
      }
    } catch {
      continue
    }
  }

  return $null
}

function Get-ToolkitCandidatePaths {
  $baseDir = Get-ControlCenterBaseDirectory
  $candidates = New-Object System.Collections.Generic.List[string]

  if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
    $parentDir = Split-Path -Parent $baseDir
    foreach ($candidate in @(
      (Join-Path $baseDir "openclaw-toolkit.ps1"),
      (Join-Path $baseDir "runtime\openclaw-toolkit.ps1"),
      $(if ($parentDir) { Join-Path $parentDir "runtime\openclaw-toolkit.ps1" } else { $null })
    )) {
      if (-not [string]::IsNullOrWhiteSpace($candidate) -and -not $candidates.Contains($candidate)) {
        $candidates.Add($candidate)
      }
    }
  }

  $userHome = Get-ControlCenterUserHome
  if (-not [string]::IsNullOrWhiteSpace($userHome)) {
    $defaultToolkitPath = Join-Path (Join-Path $userHome ".openclaw") "openclaw-toolkit.ps1"
    if (-not $candidates.Contains($defaultToolkitPath)) {
      $candidates.Add($defaultToolkitPath)
    }
  }

  return $candidates
}

function Get-EmbeddedToolkitContent {
  if ($embeddedToolkitBase64 -and $embeddedToolkitBase64 -ne "__EMBEDDED_TOOLKIT_BASE64__") {
    try {
      return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($embeddedToolkitBase64))
    } catch {
      return $null
    }
  }

  return $null
}

function Resolve-ToolkitPath {
  $toolkitCandidates = Get-ToolkitCandidatePaths
  foreach ($candidate in $toolkitCandidates) {
    if (Test-Path $candidate) {
      return $candidate
    }
  }

  if ($toolkitCandidates.Count -gt 0) {
    return $toolkitCandidates[$toolkitCandidates.Count - 1]
  }

  $userHome = Get-ControlCenterUserHome
  if (-not [string]::IsNullOrWhiteSpace($userHome)) {
    return (Join-Path (Join-Path $userHome ".openclaw") "openclaw-toolkit.ps1")
  }

  return $null
}

$toolkitContent = Get-EmbeddedToolkitContent
$toolkitPath = $null

if (-not [string]::IsNullOrWhiteSpace($toolkitContent)) {
  $script:OpenClawToolkitSelfPath = $null
  $script:OpenClawToolkitSelfContent = $toolkitContent
  . ([scriptblock]::Create($toolkitContent))
} else {
  $toolkitPath = Resolve-ToolkitPath
}

if (-not $toolkitPath) {
  $toolkitPath = Resolve-ToolkitPath
}

if (-not $toolkitContent -and -not (Test-Path $toolkitPath)) {
  [System.Windows.MessageBox]::Show(
    (Format-Loc "missing_toolkit_message" @($toolkitPath)),
    (Get-LocText "missing_toolkit_title"),
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Error
  ) | Out-Null
  exit 1
}

if (-not $toolkitContent) {
  . $toolkitPath
}

try {
  Ensure-OpenClawRuntimeFiles
  if (Get-Command Get-OpenClawContext -ErrorAction SilentlyContinue) {
    $context = Get-OpenClawContext
    if ($context -and $context.ToolkitScript -and (Test-Path $context.ToolkitScript)) {
      $toolkitPath = $context.ToolkitScript
    }
  }
} catch {
  [System.Windows.MessageBox]::Show(
    ("初始化运行时脚本失败：`n{0}" -f $_.Exception.Message),
    (Get-LocText "missing_toolkit_title"),
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Error
  ) | Out-Null
  exit 1
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="__WINDOW_TITLE__"
        Width="1180"
        Height="760"
        MinWidth="1080"
        MinHeight="700"
        ResizeMode="CanResize"
        WindowStartupLocation="CenterScreen"
        Background="#F4EEE7"
        FontFamily="Microsoft YaHei UI">
  <Grid Margin="24">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="320" />
      <ColumnDefinition Width="24" />
      <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>

    <Border Grid.Column="0" CornerRadius="34" Background="#171717" Padding="28">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto" />
          <RowDefinition Height="22" />
          <RowDefinition Height="*" />
          <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <StackPanel>
          <Border Width="184" Height="184" CornerRadius="46" Background="#F6E9DB" Padding="18">
            <Viewbox Stretch="Uniform">
              <Canvas Width="200" Height="200">
                <Path Stroke="#161616" StrokeThickness="5.5" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M92,44 C84,30 74,22 62,18" />
                <Path Stroke="#161616" StrokeThickness="5.5" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M108,44 C116,30 126,22 138,18" />
                <Path Fill="#161616" Data="M54,84 C34,71 28,48 40,36 C53,24 76,34 80,56 C83,69 73,83 54,84 Z" />
                <Path Fill="#161616" Data="M146,84 C166,71 172,48 160,36 C147,24 124,34 120,56 C117,69 127,83 146,84 Z" />
                <Path Stroke="#161616" StrokeThickness="10" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M74,96 C62,102 54,110 48,120" />
                <Path Stroke="#161616" StrokeThickness="10" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M126,96 C138,102 146,110 152,120" />
                <Path Stroke="#161616" StrokeThickness="8" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M82,128 L44,144" />
                <Path Stroke="#161616" StrokeThickness="8" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M86,150 L50,166" />
                <Path Stroke="#161616" StrokeThickness="8" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M118,128 L156,144" />
                <Path Stroke="#161616" StrokeThickness="8" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M114,150 L150,166" />
                <Ellipse Width="70" Height="92" Fill="#FF6B35" Canvas.Left="65" Canvas.Top="50" />
                <Ellipse Width="48" Height="62" Fill="#FF936B" Canvas.Left="76" Canvas.Top="64" />
                <Path Stroke="#FFF8EF" StrokeThickness="6" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M79,78 C93,72 107,72 121,78" />
                <Path Stroke="#FFF8EF" StrokeThickness="6" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M76,98 C92,92 108,92 124,98" />
                <Path Stroke="#FFF8EF" StrokeThickness="6" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M80,118 C94,112 106,112 120,118" />
                <Path Fill="#FF6B35" Data="M78,138 C88,148 112,148 122,138 L116,156 C111,167 89,167 84,156 Z" />
                <Path Fill="#FF936B" Data="M86,154 C94,161 106,161 114,154 L109,168 C105,175 95,175 91,168 Z" />
                <Path Fill="#161616" Data="M94,170 L82,190 L100,185 L107,170 Z" />
                <Path Fill="#161616" Data="M106,170 L118,190 L100,185 L93,170 Z" />
              </Canvas>
            </Viewbox>
          </Border>

          <TextBlock Margin="0,24,0,0" Width="260" FontSize="24" FontWeight="Bold" Foreground="#FFF7EE" TextWrapping="Wrap" Text="__APP_NAME__" />
          <TextBlock Margin="0,12,0,0" FontSize="15" Foreground="#D8CABE" TextWrapping="Wrap" Text="__SIDEBAR_INTRO__" />
        </StackPanel>

        <Border Grid.Row="3" CornerRadius="22" Background="#242424" Padding="18">
          <StackPanel>
            <TextBlock FontSize="12" Foreground="#C9B39E" Text="__RUNTIME_MODEL_LABEL__" />
            <TextBlock Margin="0,8,0,0" FontSize="14" Foreground="#FFF7EE" TextWrapping="Wrap" Text="__RUNTIME_MODEL_TEXT__" />
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <Grid Grid.Column="2">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />
        <RowDefinition Height="24" />
        <RowDefinition Height="Auto" />
        <RowDefinition Height="24" />
        <RowDefinition Height="*" />
      </Grid.RowDefinitions>

      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*" />
          <ColumnDefinition Width="300" />
        </Grid.ColumnDefinitions>

        <StackPanel>
          <TextBlock FontSize="34" FontWeight="Bold" Foreground="#171717" Text="__HEADLINE__" />
          <TextBlock Margin="0,10,0,0" FontSize="15" Foreground="#5D5750" Text="__SUBHEADLINE__" TextWrapping="Wrap" />
        </StackPanel>

        <StackPanel Grid.Column="1" HorizontalAlignment="Right">
          <Border CornerRadius="24" Background="#FCE4D7" Padding="18,14">
            <StackPanel>
              <TextBlock FontSize="12" Foreground="#8C5A40" Text="__CURRENT_MODE_LABEL__" />
              <TextBlock x:Name="ModeText" Margin="0,6,0,0" FontSize="24" FontWeight="Bold" Foreground="#171717" Text="-" />
            </StackPanel>
          </Border>
          <Border Margin="0,14,0,0" CornerRadius="24" Background="#EAF4EF" Padding="18,14">
            <StackPanel>
              <TextBlock FontSize="12" Foreground="#476B5A" Text="__AUTO_START_LABEL__" />
              <TextBlock x:Name="AutoStartText" Margin="0,6,0,0" FontSize="24" FontWeight="Bold" Foreground="#171717" Text="-" />
            </StackPanel>
          </Border>
        </StackPanel>
      </Grid>

      <UniformGrid Grid.Row="2" Columns="4" Rows="1">
        <Border Margin="0,0,18,0" CornerRadius="26" Background="#FFFFFF" Padding="20">
          <StackPanel>
            <TextBlock FontSize="12" Foreground="#7A746E" Text="__CURRENT_VERSION_LABEL__" />
            <TextBlock x:Name="CurrentVersionText" Margin="0,10,0,0" FontSize="22" FontWeight="Bold" Foreground="#171717" Text="-" />
          </StackPanel>
        </Border>
        <Border Margin="0,0,18,0" CornerRadius="26" Background="#FFF5EE" Padding="20">
          <StackPanel>
            <TextBlock FontSize="12" Foreground="#9A5E3E" Text="__LATEST_VERSION_LABEL__" />
            <TextBlock x:Name="LatestVersionText" Margin="0,10,0,0" FontSize="22" FontWeight="Bold" Foreground="#171717" Text="-" />
          </StackPanel>
        </Border>
        <Border Margin="0,0,18,0" CornerRadius="26" Background="#ECF8F3" Padding="20">
          <StackPanel>
            <TextBlock FontSize="12" Foreground="#3F7E67" Text="__GATEWAY_HEALTH_LABEL__" />
            <TextBlock x:Name="GatewayHealthText" Margin="0,10,0,0" FontSize="22" FontWeight="Bold" Foreground="#171717" Text="-" />
          </StackPanel>
        </Border>
        <Border CornerRadius="26" Background="#F2F0FF" Padding="20">
          <StackPanel>
            <TextBlock FontSize="12" Foreground="#5F5A90" Text="__KEEPALIVE_LABEL__" />
            <TextBlock x:Name="KeepaliveText" Margin="0,10,0,0" FontSize="22" FontWeight="Bold" Foreground="#171717" Text="-" />
          </StackPanel>
        </Border>
      </UniformGrid>

      <Grid Grid.Row="4">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto" />
          <RowDefinition Height="20" />
          <RowDefinition Height="*" />
        </Grid.RowDefinitions>

        <WrapPanel>
          <Button x:Name="StartSilentButton" Width="220" Height="56" Margin="0,0,14,14" Background="#171717" Foreground="#FFF7EE" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_START_SILENT__" />
          <Button x:Name="AutoStartButton" Width="180" Height="56" Margin="0,0,14,14" Background="#1E7A5A" Foreground="White" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_ENABLE_AUTO_START__" />
          <Button x:Name="CheckUpdateButton" Width="150" Height="56" Margin="0,0,14,14" Background="#FF6B35" Foreground="White" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_CHECK_UPDATES__" />
          <Button x:Name="BackupButton" Width="150" Height="56" Margin="0,0,14,14" Background="#244B72" Foreground="White" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_CREATE_BACKUP__" />
          <Button x:Name="RestoreButton" Width="150" Height="56" Margin="0,0,14,14" Background="#5D5FEF" Foreground="White" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_RESTORE_BACKUP__" />
          <Button x:Name="ExportLogsButton" Width="170" Height="56" Margin="0,0,14,14" Background="#FFFFFF" Foreground="#171717" BorderBrush="#DDD4CB" BorderThickness="1" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_EXPORT_LOGS__" />
          <Button x:Name="ForceStopButton" Width="190" Height="56" Margin="0,0,14,14" Background="#B42318" Foreground="White" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_FORCE_STOP__" />
          <Button x:Name="RefreshButton" Width="140" Height="56" Margin="0,0,14,14" Background="#EFE7DD" Foreground="#171717" BorderThickness="0" FontSize="15" FontWeight="SemiBold" Content="__BUTTON_REFRESH__" />
        </WrapPanel>

        <Border Grid.Row="2" CornerRadius="28" Background="#FFFFFF" Padding="22">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto" />
              <RowDefinition Height="18" />
              <RowDefinition Height="Auto" />
              <RowDefinition Height="18" />
              <RowDefinition Height="Auto" />
              <RowDefinition Height="18" />
              <RowDefinition Height="*" />
            </Grid.RowDefinitions>

            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="260" />
              </Grid.ColumnDefinitions>
              <StackPanel>
                <TextBlock FontSize="24" FontWeight="Bold" Foreground="#171717" Text="__SECTION_FEEDBACK_TITLE__" />
                <TextBlock Margin="0,8,0,0" FontSize="14" Foreground="#6B645E" Text="__SECTION_FEEDBACK_SUBTITLE__" TextWrapping="Wrap" />
              </StackPanel>
              <Border Grid.Column="1" HorizontalAlignment="Right" CornerRadius="18" Background="#F6EFE6" Padding="16,12">
                <TextBlock x:Name="BusyText" FontSize="13" Foreground="#8A5E42" Text="__BUSY_READY__" />
              </Border>
            </Grid>

            <Border Grid.Row="2" CornerRadius="20" Background="#F7F3EE" Padding="18">
              <TextBlock x:Name="OperationText" FontSize="15" Foreground="#1F1F1F" TextWrapping="Wrap" Text="__OPERATION_DEFAULT__" />
            </Border>

            <Grid Grid.Row="4">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="24" />
                <ColumnDefinition Width="*" />
              </Grid.ColumnDefinitions>

              <Border CornerRadius="20" Background="#FCFBF9" Padding="18">
                <StackPanel>
                  <TextBlock FontSize="13" Foreground="#7A746E" Text="__SECTION_LOGS_TITLE__" />
                  <TextBlock Margin="0,10,0,0" FontSize="14" Foreground="#171717" Text="__LABEL_WATCHDOG__" />
                  <TextBlock x:Name="WatchdogLogText" Margin="0,4,0,0" FontSize="12" Foreground="#6B645E" TextWrapping="Wrap" />
                  <TextBlock Margin="0,12,0,0" FontSize="14" Foreground="#171717" Text="__LABEL_SUPERVISOR__" />
                  <TextBlock x:Name="SupervisorLogText" Margin="0,4,0,0" FontSize="12" Foreground="#6B645E" TextWrapping="Wrap" />
                  <TextBlock Margin="0,12,0,0" FontSize="14" Foreground="#171717" Text="__LABEL_LATEST_AUDIT__" />
                  <TextBlock x:Name="AuditPathText" Margin="0,4,0,0" FontSize="12" Foreground="#6B645E" TextWrapping="Wrap" />
                </StackPanel>
              </Border>

              <Border Grid.Column="2" CornerRadius="20" Background="#FCFBF9" Padding="18">
                <StackPanel>
                  <TextBlock FontSize="13" Foreground="#7A746E" Text="__SECTION_RUNTIME_TITLE__" />
                  <TextBlock Margin="0,10,0,0" FontSize="14" Foreground="#171717" Text="__LABEL_COMMAND__" />
                  <TextBlock x:Name="CommandPathText" Margin="0,4,0,0" FontSize="12" Foreground="#6B645E" TextWrapping="Wrap" />
                  <TextBlock Margin="0,12,0,0" FontSize="14" Foreground="#171717" Text="__LABEL_PREFIX__" />
                  <TextBlock x:Name="PrefixPathText" Margin="0,4,0,0" FontSize="12" Foreground="#6B645E" TextWrapping="Wrap" />
                  <TextBlock Margin="0,12,0,0" FontSize="14" Foreground="#171717" Text="__LABEL_GATEWAY_LOG_DIR__" />
                  <TextBlock x:Name="GatewayLogDirText" Margin="0,4,0,0" FontSize="12" Foreground="#6B645E" TextWrapping="Wrap" />
                </StackPanel>
              </Border>
            </Grid>

            <Border Grid.Row="6" CornerRadius="20" Background="#FCFBF9" Padding="18">
              <StackPanel>
                <TextBlock FontSize="13" Foreground="#7A746E" Text="__SECTION_RECENT_ERRORS_TITLE__" />
                <TextBlock Margin="0,8,0,0" FontSize="14" Foreground="#171717" Text="__SECTION_RECENT_ERRORS_SUBTITLE__" TextWrapping="Wrap" />
                <TextBox x:Name="RecentErrorsText" Margin="0,14,0,0" Background="Transparent" BorderThickness="0" IsReadOnly="True"
                         AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas"
                         FontSize="12" Foreground="#5E5852" MinHeight="156" />
              </StackPanel>
            </Border>
          </Grid>
        </Border>
      </Grid>
    </Grid>
  </Grid>
</Window>
"@

$tokenMap = @{
  "__WINDOW_TITLE__" = Get-LocText "window_title"
  "__APP_NAME__" = Get-LocText "app_name"
  "__SIDEBAR_INTRO__" = Get-LocText "sidebar_intro"
  "__RUNTIME_MODEL_LABEL__" = Get-LocText "runtime_model_label"
  "__RUNTIME_MODEL_TEXT__" = Get-LocText "runtime_model_text"
  "__HEADLINE__" = Get-LocText "headline"
  "__SUBHEADLINE__" = Get-LocText "subheadline"
  "__CURRENT_MODE_LABEL__" = Get-LocText "current_mode_label"
  "__AUTO_START_LABEL__" = Get-LocText "auto_start_label"
  "__CURRENT_VERSION_LABEL__" = Get-LocText "current_version_label"
  "__LATEST_VERSION_LABEL__" = Get-LocText "latest_version_label"
  "__GATEWAY_HEALTH_LABEL__" = Get-LocText "gateway_health_label"
  "__KEEPALIVE_LABEL__" = Get-LocText "keepalive_label"
  "__BUTTON_START_SILENT__" = Get-LocText "button_start_silent"
  "__BUTTON_ENABLE_AUTO_START__" = Get-LocText "button_enable_auto_start"
  "__BUTTON_CHECK_UPDATES__" = Get-LocText "button_check_updates"
  "__BUTTON_CREATE_BACKUP__" = Get-LocText "button_create_backup"
  "__BUTTON_RESTORE_BACKUP__" = Get-LocText "button_restore_backup"
  "__BUTTON_EXPORT_LOGS__" = Get-LocText "button_export_logs"
  "__BUTTON_FORCE_STOP__" = Get-LocText "button_force_stop"
  "__BUTTON_REFRESH__" = Get-LocText "button_refresh"
  "__SECTION_FEEDBACK_TITLE__" = Get-LocText "section_feedback_title"
  "__SECTION_FEEDBACK_SUBTITLE__" = Get-LocText "section_feedback_subtitle"
  "__BUSY_READY__" = Get-LocText "busy_ready"
  "__OPERATION_DEFAULT__" = Get-LocText "operation_default"
  "__SECTION_LOGS_TITLE__" = Get-LocText "section_logs_title"
  "__LABEL_WATCHDOG__" = Get-LocText "label_watchdog"
  "__LABEL_SUPERVISOR__" = Get-LocText "label_supervisor"
  "__LABEL_LATEST_AUDIT__" = Get-LocText "label_latest_audit"
  "__SECTION_RUNTIME_TITLE__" = Get-LocText "section_runtime_title"
  "__LABEL_COMMAND__" = Get-LocText "label_command"
  "__LABEL_PREFIX__" = Get-LocText "label_prefix"
  "__LABEL_GATEWAY_LOG_DIR__" = Get-LocText "label_gateway_log_dir"
  "__SECTION_RECENT_ERRORS_TITLE__" = Get-LocText "section_recent_errors_title"
  "__SECTION_RECENT_ERRORS_SUBTITLE__" = Get-LocText "section_recent_errors_subtitle"
}

foreach ($token in $tokenMap.Keys) {
  $xaml = $xaml.Replace($token, [string]$tokenMap[$token])
}

[xml]$xamlXml = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xamlXml
$window = [Windows.Markup.XamlReader]::Load($reader)
$window.Title = Get-LocText "window_title"

$names = @(
  "ModeText",
  "AutoStartText",
  "CurrentVersionText",
  "LatestVersionText",
  "GatewayHealthText",
  "KeepaliveText",
  "BusyText",
  "OperationText",
  "WatchdogLogText",
  "SupervisorLogText",
  "AuditPathText",
  "CommandPathText",
  "PrefixPathText",
  "GatewayLogDirText",
  "RecentErrorsText",
  "StartSilentButton",
  "AutoStartButton",
  "CheckUpdateButton",
  "BackupButton",
  "RestoreButton",
  "ExportLogsButton",
  "ForceStopButton",
  "RefreshButton"
)

$ui = @{}
foreach ($name in $names) {
  $ui[$name] = $window.FindName($name)
}

$script:LatestVersionCache = $null
$script:LatestVersionCheckedAt = $null
$script:LastStatus = $null
$script:StickyMessageUntil = [datetime]::MinValue
$script:PendingOperations = New-Object System.Collections.ArrayList
$script:IsRefreshRunning = $false
$script:RefreshQueued = $false
$script:RefreshQueuedIncludeLatest = $false
$script:ActiveActionCount = 0
$script:ToolkitPathLiteral = $toolkitPath.Replace("'", "''")

function Convert-ModeText {
  param([string]$Mode)

  switch ($Mode) {
    "Running" { return Get-LocText "mode_running" }
    "Starting" { return Get-LocText "mode_starting" }
    "Failed" { return Get-LocText "mode_failed" }
    "Paused" { return Get-LocText "mode_paused" }
    "Updating" { return Get-LocText "mode_updating" }
    "Stopped" { return Get-LocText "mode_stopped" }
    default { return $Mode }
  }
}

function Convert-TaskStateText {
  param([string]$State)

  switch ($State) {
    "Ready" { return Get-LocText "task_ready" }
    "Running" { return Get-LocText "task_running" }
    "Disabled" { return Get-LocText "task_disabled" }
    "Queued" { return Get-LocText "task_queued" }
    "Missing" { return Get-LocText "task_missing" }
    default { return $State }
  }
}

function Convert-GatewayHealthText {
  param($Status)

  if ($Status.GatewayHealthy) {
    return Get-LocText "status_live"
  }

  if ($Status.Mode -eq "Failed") {
    return Get-LocText "status_failed_gateway"
  }

  if ($Status.Paused -or $Status.Mode -eq "Paused" -or $Status.Mode -eq "Stopped") {
    return Get-LocText "status_stopped_gateway"
  }

  if ($Status.ProcessCount -gt 0 -or $Status.Mode -eq "Starting" -or $Status.Mode -eq "Updating") {
    return Get-LocText "status_waiting"
  }

  return Get-LocText "status_stopped_gateway"
}

function Set-OperationMessage {
  param(
    [string]$Message,
    [string]$BusyLabel = $(Get-LocText "busy_ready")
  )

  $ui.OperationText.Text = $Message
  $ui.BusyText.Text = $BusyLabel
}

function Set-UserOperationMessage {
  param(
    [string]$Message,
    [string]$BusyLabel,
    [int]$StickySeconds = 10
  )

  Set-OperationMessage -Message $Message -BusyLabel $BusyLabel
  $script:StickyMessageUntil = (Get-Date).AddSeconds($StickySeconds)
}

function Set-ButtonsEnabled {
  param([bool]$Enabled)

  foreach ($name in @("StartSilentButton", "AutoStartButton", "CheckUpdateButton", "BackupButton", "RestoreButton", "ExportLogsButton", "ForceStopButton", "RefreshButton")) {
    $ui[$name].IsEnabled = $Enabled
  }
}

function Update-ActionButtonState {
  Set-ButtonsEnabled -Enabled ($script:ActiveActionCount -eq 0)
}

function Set-StatusText {
  param($Status)

  $ui.ModeText.Text = Convert-ModeText $Status.Mode
  $ui.AutoStartText.Text = if ($Status.AutoStartEnabled) { Get-LocText "status_enabled" } else { Get-LocText "status_disabled" }

  if ($Status.CurrentVersion) {
    $ui.CurrentVersionText.Text = $Status.CurrentVersion
  } else {
    $ui.CurrentVersionText.Text = Get-LocText "status_not_installed"
  }

  if ($Status.LatestVersion) {
    $ui.LatestVersionText.Text = $Status.LatestVersion
  } else {
    $ui.LatestVersionText.Text = Get-LocText "status_click_check"
  }

  $ui.GatewayHealthText.Text = Convert-GatewayHealthText $Status

  $ui.KeepaliveText.Text = Convert-TaskStateText $Status.KeepaliveTaskState
  $ui.WatchdogLogText.Text = $Status.WatchdogLog
  $ui.SupervisorLogText.Text = $Status.SupervisorLog
  $ui.AuditPathText.Text = if ($Status.LatestAuditPath) { $Status.LatestAuditPath } else { Get-LocText "task_missing" }
  $ui.CommandPathText.Text = if ($Status.CommandPath) { $Status.CommandPath } else { Get-LocText "status_not_installed" }
  $ui.PrefixPathText.Text = $Status.Prefix
  $ui.GatewayLogDirText.Text = $Status.GatewayLogDir
  $ui.AutoStartButton.Content = if ($Status.AutoStartEnabled) { Get-LocText "button_disable_auto_start" } else { Get-LocText "button_enable_auto_start" }
}

function Set-DefaultOperationMessage {
  param(
    $Status,
    [switch]$Force
  )

  if ($script:ActiveActionCount -gt 0) {
    return
  }

  if (-not $Force -and (Get-Date) -lt $script:StickyMessageUntil) {
    return
  }

  if ($status.UpgradeLocked) {
    Set-OperationMessage -Message (Get-LocText "msg_updating") -BusyLabel (Get-LocText "busy_updating")
  } elseif ($status.Paused) {
    Set-OperationMessage -Message (Get-LocText "msg_paused") -BusyLabel (Get-LocText "busy_paused")
  } elseif ($status.GatewayHealthy) {
    Set-OperationMessage -Message (Get-LocText "msg_healthy") -BusyLabel (Get-LocText "busy_healthy")
  } elseif ($status.Mode -eq "Failed") {
    $failureReason = if ($status.FailureReason) { [string]$status.FailureReason } else { Get-LocText "msg_failed_generic" }
    Set-OperationMessage -Message ((Get-LocText "msg_failed_template") -f $failureReason) -BusyLabel (Get-LocText "busy_failed")
  } else {
    Set-OperationMessage -Message (Get-LocText "msg_starting") -BusyLabel (Get-LocText "busy_starting")
  }
}

function Apply-StatusPayload {
  param(
    $Payload,
    [switch]$ForceDefaultMessage
  )

  $status = $Payload.Status
  if (-not $status) {
    return $null
  }

  if ($status.LatestVersion) {
    $script:LatestVersionCache = $status.LatestVersion
    $script:LatestVersionCheckedAt = Get-Date
  } elseif ($script:LatestVersionCache) {
    $status.LatestVersion = $script:LatestVersionCache
    $status.UpdateAvailable = ((Compare-OpenClawVersion $script:LatestVersionCache $status.CurrentVersion) -gt 0)
  }

  if (-not ($status.PSObject.Properties.Name -contains "IsInstalled")) {
    $installedDefault = -not [string]::IsNullOrWhiteSpace([string]$status.CommandPath)
    Add-Member -InputObject $status -NotePropertyName "IsInstalled" -NotePropertyValue $installedDefault -Force
  }

  if (-not ($status.PSObject.Properties.Name -contains "HasExistingConfig")) {
    $userHome = Get-ControlCenterUserHome
    $openClawHome = if (-not [string]::IsNullOrWhiteSpace($userHome)) { Join-Path $userHome ".openclaw" } else { $null }
    Add-Member -InputObject $status -NotePropertyName "HasExistingConfig" -NotePropertyValue ($openClawHome -and (Test-Path $openClawHome)) -Force
  }

  $script:LastStatus = $status
  Set-StatusText -Status $status

  if ($Payload.PSObject.Properties.Name -contains "RecentErrors") {
    $ui.RecentErrorsText.Text = [string]$Payload.RecentErrors
  }

  Set-DefaultOperationMessage -Status $status -Force:$ForceDefaultMessage
  return $status
}

function Complete-RefreshIfQueued {
  if (-not $script:RefreshQueued) {
    return
  }

  $includeLatestVersion = $script:RefreshQueuedIncludeLatest
  $script:RefreshQueued = $false
  $script:RefreshQueuedIncludeLatest = $false
  Request-PanelRefresh -IncludeLatestVersion:$includeLatestVersion
}

function Start-PanelAsyncOperation {
  param(
    [string]$Name,
    [string]$Body,
    [object[]]$Arguments = @(),
    [scriptblock]$OnSuccess,
    [scriptblock]$OnError,
    [switch]$IsAction,
    [switch]$IsRefresh
  )

  $scriptText = @"
param(`$Payload)
`$ErrorActionPreference = 'Stop'
. '$($script:ToolkitPathLiteral)'
$Body
"@

  $ps = [PowerShell]::Create()
  $null = $ps.AddScript($scriptText).AddArgument($Arguments)
  $handle = $ps.BeginInvoke()

  if ($IsAction) {
    $script:ActiveActionCount++
    Update-ActionButtonState
  }

  if ($IsRefresh) {
    $script:IsRefreshRunning = $true
  }

  [void]$script:PendingOperations.Add([pscustomobject]@{
      Name = $Name
      PowerShell = $ps
      Handle = $handle
      OnSuccess = $OnSuccess
      OnError = $OnError
      IsAction = [bool]$IsAction
      IsRefresh = [bool]$IsRefresh
    })
}

function Process-PendingOperations {
  for ($i = $script:PendingOperations.Count - 1; $i -ge 0; $i--) {
    $operation = $script:PendingOperations[$i]
    if (-not $operation.Handle.IsCompleted) {
      continue
    }

    $errorMessage = $null
    $result = $null

    try {
      $results = @($operation.PowerShell.EndInvoke($operation.Handle))
      if ($operation.PowerShell.HadErrors) {
        $firstError = $operation.PowerShell.Streams.Error | Select-Object -First 1
        if ($firstError) {
          throw [System.InvalidOperationException]::new($firstError.ToString())
        }
      }

      if ($results.Count -eq 1) {
        $result = $results[0]
      } elseif ($results.Count -gt 1) {
        $result = $results
      }
    } catch {
      $errorMessage = $_.Exception.Message
    } finally {
      try {
        $operation.PowerShell.Dispose()
      } catch {}

      $script:PendingOperations.RemoveAt($i)

      if ($operation.IsRefresh) {
        $script:IsRefreshRunning = $false
      }
    }

    try {
      if ($errorMessage) {
        if ($operation.OnError) {
          & $operation.OnError $errorMessage
        }
      } elseif ($operation.OnSuccess) {
        & $operation.OnSuccess $result
      }
    } finally {
      if ($operation.IsAction) {
        $script:ActiveActionCount = [Math]::Max(0, ($script:ActiveActionCount - 1))
        Update-ActionButtonState
      }

      if ($operation.IsRefresh) {
        Complete-RefreshIfQueued
      }
    }
  }
}

function Request-PanelRefresh {
  param(
    [switch]$IncludeLatestVersion,
    [switch]$ShowBusyMessage,
    [switch]$ForceDefaultMessage
  )

  if ($script:IsRefreshRunning) {
    $script:RefreshQueued = $true
    if ($IncludeLatestVersion) {
      $script:RefreshQueuedIncludeLatest = $true
    }
    return
  }

  $shouldRefreshLatest = $IncludeLatestVersion -or
    -not $script:LatestVersionCache -or
    -not $script:LatestVersionCheckedAt -or
    (((Get-Date) - $script:LatestVersionCheckedAt).TotalMinutes -ge 30)

  if ($ShowBusyMessage -and $script:ActiveActionCount -eq 0) {
    Set-OperationMessage -Message (Get-LocText "msg_refreshing") -BusyLabel (Get-LocText "busy_refreshing")
  }

  $body = @'
$includeLatest = [bool]$Payload[0]
$status = Get-OpenClawStatus -IncludeLatestVersion:$includeLatest
$errors = Get-OpenClawRecentErrorSummary -MaxItems 12
[pscustomobject]@{
  Status = $status
  RecentErrors = $errors
}
'@

  $forceDefaultMessageValue = [bool]$ForceDefaultMessage
  $refreshSuccess = {
    param($payload)
    Apply-StatusPayload -Payload $payload -ForceDefaultMessage:$forceDefaultMessageValue | Out-Null
  }.GetNewClosure()

  $refreshError = {
    param($message)
    if ($script:ActiveActionCount -eq 0) {
      Set-OperationMessage -Message ((Get-LocText "msg_refresh_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_check_failed")
    }
  }.GetNewClosure()

  Start-PanelAsyncOperation -Name "refresh" -Body $body -Arguments @($shouldRefreshLatest) -IsRefresh `
    -OnSuccess $refreshSuccess `
    -OnError $refreshError
}

function Start-StatusAction {
  param(
    [string]$Name,
    [string]$BusyMessage,
    [string]$BusyLabel,
    [string]$Body,
    [object[]]$Arguments = @(),
    [scriptblock]$OnSuccess,
    [scriptblock]$OnError
  )

  Set-UserOperationMessage -Message $BusyMessage -BusyLabel $BusyLabel -StickySeconds 30

  $wrappedError = {
    param($message)
    if ($OnError) {
      & $OnError $message
    } else {
      Set-UserOperationMessage -Message $message -BusyLabel (Get-LocText "busy_check_failed")
    }
  }.GetNewClosure()

  Start-PanelAsyncOperation -Name $Name -Body $Body -Arguments $Arguments -IsAction `
    -OnSuccess $OnSuccess `
    -OnError $wrappedError
}

$ui.RefreshButton.Add_Click({
  Request-PanelRefresh -IncludeLatestVersion -ShowBusyMessage -ForceDefaultMessage
})

$ui.StartSilentButton.Add_Click({
  $body = @'
$status = Start-OpenClawSilent
$errors = Get-OpenClawRecentErrorSummary -MaxItems 12
[pscustomobject]@{
  Status = $status
  RecentErrors = $errors
}
'@

  Start-StatusAction -Name "start-silent" `
    -BusyMessage (Get-LocText "msg_starting_silent") `
    -BusyLabel (Get-LocText "busy_starting_action") `
    -Body $body `
    -OnSuccess {
      param($payload)
      $status = Apply-StatusPayload -Payload $payload
      if ($status) {
        Set-UserOperationMessage -Message (Format-Loc "msg_start_triggered_template" @((Convert-ModeText $status.Mode))) -BusyLabel (Get-LocText "busy_triggered")
      }
      Request-PanelRefresh
    } `
    -OnError {
      param($message)
      Set-UserOperationMessage -Message ((Get-LocText "msg_start_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_check_failed")
    }
})

$ui.AutoStartButton.Add_Click({
  $enable = $true
  if ($script:LastStatus) {
    $enable = -not $script:LastStatus.AutoStartEnabled
  }

  $body = @'
$enable = [bool]$Payload[0]
$status = Set-OpenClawAutoStart -Enabled $enable
$errors = Get-OpenClawRecentErrorSummary -MaxItems 12
[pscustomobject]@{
  Status = $status
  RecentErrors = $errors
  Enabled = $enable
}
'@

  Start-StatusAction -Name "toggle-autostart" `
    -BusyMessage ($(if ($enable) { Get-LocText "msg_autostart_enabling" } else { Get-LocText "msg_autostart_disabling" })) `
    -BusyLabel (Get-LocText "busy_updating") `
    -Body $body `
    -Arguments @($enable) `
    -OnSuccess {
      param($payload)
      $status = Apply-StatusPayload -Payload $payload
      if ($payload.Enabled) {
        Set-UserOperationMessage -Message (Get-LocText "msg_autostart_enabled") -BusyLabel (Get-LocText "busy_updated")
      } else {
        Set-UserOperationMessage -Message (Get-LocText "msg_autostart_disabled") -BusyLabel (Get-LocText "busy_updated")
      }
      if (-not $status) {
        Request-PanelRefresh
      }
    } `
    -OnError {
      param($message)
      Set-UserOperationMessage -Message ((Get-LocText "msg_autostart_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_check_failed")
    }
})

$ui.ForceStopButton.Add_Click({
  $result = [System.Windows.MessageBox]::Show(
    (Get-LocText "dialog_force_stop_message"),
    (Get-LocText "dialog_force_stop_title"),
    [System.Windows.MessageBoxButton]::YesNo,
    [System.Windows.MessageBoxImage]::Warning
  )

  if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
    return
  }

  $body = @'
$status = Stop-OpenClawSilent
$errors = Get-OpenClawRecentErrorSummary -MaxItems 12
[pscustomobject]@{
  Status = $status
  RecentErrors = $errors
}
'@

  Start-StatusAction -Name "force-stop" `
    -BusyMessage (Get-LocText "msg_stopping_silent") `
    -BusyLabel (Get-LocText "busy_stopping") `
    -Body $body `
    -OnSuccess {
      param($payload)
      $status = Apply-StatusPayload -Payload $payload -ForceDefaultMessage
      if ($status) {
        Set-UserOperationMessage -Message (Format-Loc "msg_stopped_template" @((Convert-ModeText $status.Mode))) -BusyLabel (Get-LocText "busy_stopped")
      }
    } `
    -OnError {
      param($message)
      Set-UserOperationMessage -Message ((Get-LocText "msg_stop_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_stop_failed")
    }
})

$ui.CheckUpdateButton.Add_Click({
  $checkBody = @'
$status = Get-OpenClawStatus -IncludeLatestVersion
$errors = Get-OpenClawRecentErrorSummary -MaxItems 12
[pscustomobject]@{
  Status = $status
  RecentErrors = $errors
}
'@

  Start-StatusAction -Name "check-updates" `
    -BusyMessage (Get-LocText "msg_checking_updates") `
    -BusyLabel (Get-LocText "busy_checking") `
    -Body $checkBody `
    -OnSuccess {
      param($payload)
      $status = Apply-StatusPayload -Payload $payload
      if (-not $status) {
        return
      }

      if (-not $status.LatestVersion) {
        Set-UserOperationMessage -Message (Get-LocText "msg_latest_version_unavailable") -BusyLabel (Get-LocText "busy_check_failed")
        return
      }

      $hasInstalledVersion = -not [string]::IsNullOrWhiteSpace($status.CurrentVersion)
      $hasInstalledOrResidualState = [bool]$status.IsInstalled -or [bool]$status.HasExistingConfig
      $requiresConfirmation = $hasInstalledOrResidualState
      $isFreshInstall = -not $hasInstalledOrResidualState

      if ($hasInstalledVersion -and -not $status.UpdateAvailable) {
        Set-UserOperationMessage -Message (Format-Loc "msg_up_to_date_template" @($status.CurrentVersion)) -BusyLabel (Get-LocText "busy_up_to_date")
        return
      }

      if ($requiresConfirmation) {
        $promptMessage = if ($hasInstalledVersion) {
          Format-Loc "dialog_update_prompt_template" @($status.LatestVersion, $status.CurrentVersion)
        } else {
          Format-Loc "dialog_install_existing_prompt_template" @($status.LatestVersion)
        }

        $prompt = [System.Windows.MessageBox]::Show(
          $promptMessage,
          (Get-LocText "dialog_update_title"),
          [System.Windows.MessageBoxButton]::YesNo,
          [System.Windows.MessageBoxImage]::Question
        )

        if ($prompt -ne [System.Windows.MessageBoxResult]::Yes) {
          Set-UserOperationMessage -Message (Get-LocText "msg_update_canceled") -BusyLabel (Get-LocText "busy_canceled")
          return
        }
      }

      $updateBody = @'
$targetVersion = [string]$Payload[0]
$proc = Start-OpenClawUpdate -TargetVersion $targetVersion
[pscustomobject]@{
  TargetVersion = $targetVersion
  ProcessId = $proc.Id
}
'@

      $updateStartedSuccess = {
        param($updatePayload)
        if ($isFreshInstall) {
          Set-UserOperationMessage -Message (Format-Loc "msg_install_started_template" @($updatePayload.TargetVersion)) -BusyLabel (Format-Loc "busy_update_pid_template" @($updatePayload.ProcessId))
        } else {
          Set-UserOperationMessage -Message (Format-Loc "msg_update_started_template" @($updatePayload.TargetVersion)) -BusyLabel (Format-Loc "busy_update_pid_template" @($updatePayload.ProcessId))
        }
        Request-PanelRefresh -IncludeLatestVersion
      }.GetNewClosure()

      Start-StatusAction -Name "start-update" `
        -BusyMessage (Get-LocText "msg_updating") `
        -BusyLabel (Get-LocText "busy_updating") `
        -Body $updateBody `
        -Arguments @($status.LatestVersion) `
        -OnSuccess $updateStartedSuccess `
        -OnError {
          param($message)
          Set-UserOperationMessage -Message ((Get-LocText "msg_update_check_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_check_failed")
        }
    } `
    -OnError {
      param($message)
      Set-UserOperationMessage -Message ((Get-LocText "msg_update_check_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_check_failed")
    }
})

$ui.BackupButton.Add_Click({
  $defaultName = Format-Loc "backup_default_name_template" @((Get-Date -Format "yyyyMMdd-HHmm"))
  $backupName = [Microsoft.VisualBasic.Interaction]::InputBox(
    (Get-LocText "dialog_backup_name_prompt"),
    (Get-LocText "dialog_backup_name_title"),
    $defaultName
  )

  if ([string]::IsNullOrWhiteSpace($backupName)) {
    Set-UserOperationMessage -Message (Get-LocText "msg_backup_name_canceled") -BusyLabel (Get-LocText "busy_canceled")
    return
  }

  $body = @'
$backupPath = New-OpenClawBackup -BackupName ([string]$Payload[0])
$status = Get-OpenClawStatus
[pscustomobject]@{
  Status = $status
  BackupPath = $backupPath
}
'@

  Start-StatusAction -Name "create-backup" `
    -BusyMessage (Get-LocText "msg_creating_backup") `
    -BusyLabel (Get-LocText "busy_backing_up") `
    -Body $body `
    -Arguments @($backupName) `
    -OnSuccess {
      param($payload)
      if ($payload.PSObject.Properties.Name -contains "Status") {
        Apply-StatusPayload -Payload $payload | Out-Null
      }
      Set-UserOperationMessage -Message (Format-Loc "msg_backup_created_template" @($payload.BackupPath)) -BusyLabel (Get-LocText "busy_backed_up")
    } `
    -OnError {
      param($message)
      Set-UserOperationMessage -Message ((Get-LocText "msg_backup_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_check_failed")
    }
})

$ui.RestoreButton.Add_Click({
  try {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = Get-LocText "restore_filter"
    $dialog.InitialDirectory = if ($script:LastStatus -and $script:LastStatus.BackupRoot -and (Test-Path $script:LastStatus.BackupRoot)) {
      $script:LastStatus.BackupRoot
    } else {
      $userHome = Get-ControlCenterUserHome
      if (-not [string]::IsNullOrWhiteSpace($userHome)) { Join-Path $userHome ".openclaw\backups" } else { Get-ControlCenterBaseDirectory }
    }

    if (-not $dialog.ShowDialog()) {
      Set-UserOperationMessage -Message (Get-LocText "msg_restore_canceled") -BusyLabel (Get-LocText "busy_canceled")
      return
    }

    $confirm = [System.Windows.MessageBox]::Show(
      (Format-Loc "dialog_restore_prompt_template" @($dialog.FileName)),
      (Get-LocText "dialog_restore_title"),
      [System.Windows.MessageBoxButton]::YesNo,
      [System.Windows.MessageBoxImage]::Warning
    )

    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
      Set-UserOperationMessage -Message (Get-LocText "msg_restore_canceled") -BusyLabel (Get-LocText "busy_canceled")
      return
    }

    $body = @'
$status = Restore-OpenClawBackup -BackupPath ([string]$Payload[0])
[pscustomobject]@{
  Status = $status
  BackupPath = [string]$Payload[0]
}
'@

    Start-StatusAction -Name "restore-backup" `
      -BusyMessage (Get-LocText "msg_restoring_backup") `
      -BusyLabel (Get-LocText "busy_restoring") `
      -Body $body `
      -Arguments @($dialog.FileName) `
      -OnSuccess {
        param($payload)
        Apply-StatusPayload -Payload $payload -ForceDefaultMessage | Out-Null
        Set-UserOperationMessage -Message (Format-Loc "msg_restore_completed_template" @($payload.BackupPath)) -BusyLabel (Get-LocText "busy_restored")
      } `
      -OnError {
        param($message)
        Set-UserOperationMessage -Message ((Get-LocText "msg_restore_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_stop_failed")
      }
  }
  catch {
    Set-UserOperationMessage -Message ((Get-LocText "msg_restore_failed_prefix") + $_.Exception.Message) -BusyLabel (Get-LocText "busy_stop_failed")
  }
})

$ui.ExportLogsButton.Add_Click({
  try {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.FileName = (Format-Loc "export_filename_template" @((Get-Date -Format "yyyyMMdd-HHmmss")))
    $dialog.Filter = Get-LocText "export_filter"
    $dialog.InitialDirectory = Get-ControlCenterDesktopDirectory

    if (-not $dialog.ShowDialog()) {
      Set-UserOperationMessage -Message (Get-LocText "msg_log_export_canceled") -BusyLabel (Get-LocText "busy_canceled")
      return
    }

    $body = @'
$zipPath = Export-OpenClawErrorLogs -OutputZipPath ([string]$Payload[0])
[pscustomobject]@{
  ZipPath = $zipPath
}
'@

    Start-StatusAction -Name "export-logs" `
      -BusyMessage (Get-LocText "msg_collecting_logs") `
      -BusyLabel (Get-LocText "busy_exporting") `
      -Body $body `
      -Arguments @($dialog.FileName) `
      -OnSuccess {
        param($payload)
        Set-UserOperationMessage -Message (Format-Loc "msg_exported_template" @($payload.ZipPath)) -BusyLabel (Get-LocText "busy_exported")
      } `
      -OnError {
        param($message)
        Set-UserOperationMessage -Message ((Get-LocText "msg_export_failed_prefix") + $message) -BusyLabel (Get-LocText "busy_export_failed")
      }
  }
  catch {
    Set-UserOperationMessage -Message ((Get-LocText "msg_export_failed_prefix") + $_.Exception.Message) -BusyLabel (Get-LocText "busy_export_failed")
  }
})

$pollTimer = New-Object System.Windows.Threading.DispatcherTimer
$pollTimer.Interval = [TimeSpan]::FromMilliseconds(200)
$pollTimer.Add_Tick({
  Process-PendingOperations
})

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(15)
$timer.Add_Tick({
  Request-PanelRefresh
})

$window.Add_SourceInitialized({
  $pollTimer.Start()
  Set-OperationMessage -Message (Get-LocText "msg_refreshing") -BusyLabel (Get-LocText "busy_refreshing")
  Request-PanelRefresh -IncludeLatestVersion -ForceDefaultMessage
  $timer.Start()
})

$window.Add_Closing({
  $timer.Stop()
  $pollTimer.Stop()

  foreach ($operation in @($script:PendingOperations)) {
    try {
      $operation.PowerShell.Stop()
    } catch {}

    try {
      $operation.PowerShell.Dispose()
    } catch {}
  }
})

$null = $window.ShowDialog()
