# OpenClaw 控制中心

这是一个面向 Windows 的 OpenClaw 静默运行控制面板项目。

当前仓库包含：

- `panel/`
  控制面板源码、图标生成脚本、打包脚本。
- `runtime/`
  面板运行与打包依赖的 OpenClaw toolkit 源码。
- `dist/`
  已构建的 `OpenClaw 控制中心.exe`。
- `.tools/ps2exe/`
  本地打包 EXE 所需的 PS2EXE 工具。

## 功能

- 24 小时静默运行控制
- 开机自启开关
- 检查 OpenClaw 更新
- 一键备份，可自定义名称并保留多个版本
- 一键恢复，可选择历史备份恢复
- 导出报错日志
- 强制退出后台静默运行

## 重新构建 EXE

在 PowerShell 中执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\panel\build-control-center.ps1
```

构建结果会输出到：

- `dist/OpenClaw 控制中心.exe`

并同步复制到当前用户桌面。
