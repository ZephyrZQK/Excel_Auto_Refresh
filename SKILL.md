---
name: excel-pivot-auto-refresh
description: Build, maintain, or package a bilingual local Windows desktop app that lets non-technical users schedule automatic refreshes for Excel workbooks, pivot tables, Power Query connections, and external data. Use when users need a privacy-preserving local .exe for scheduled Excel refreshes, especially when workbooks contain confidential data and must not be uploaded. / 构建、维护或打包一个中英双语的 Windows 本地桌面应用，让非技术用户可以定时自动刷新 Excel 工作簿、数据透视表、Power Query 和外部数据连接；适用于含涉密数据、不能上传文件的本地自动刷新场景。
---

# Excel Pivot Auto Refresh / Excel 数据透视表自动刷新

## Purpose / 目的

Use this skill to create or maintain a Windows-only, single-file desktop app for scheduled Excel refresh tasks. The app is designed for non-technical users: they choose Excel files, pick daily refresh times, enable or pause tasks, and view recent status without seeing scripts or command lines.

使用此技能来创建或维护一个仅适用于 Windows 的单文件桌面应用，用于定时刷新 Excel。该应用面向非技术用户：用户只需选择 Excel 文件、设置每日刷新时间、启用或暂停任务，并查看最近运行状态，无需接触脚本或命令行。

## Core Product Rules / 核心产品规则

- Keep all workbook content on the user's machine. / 所有工作簿内容都保留在用户本机。
- Package the app as one `.exe` for distribution. / 应用必须打包成单个 `.exe` 文件用于分发。
- Support multiple workbook refresh tasks per user. / 支持每个用户配置多个工作簿刷新任务。
- Use Windows Task Scheduler so refreshes run after the app is closed. / 使用 Windows 任务计划程序，确保应用关闭后仍可定时运行。
- Configure missed tasks to run as soon as possible after the computer starts or the user logs in. / 如果电脑关机导致错过任务，应在电脑启动或用户登录后尽快补跑。
- Do not unlock sheets, protect sheets, fill formulas, or modify workbook business logic. / 不解除工作表保护、不重新保护工作表、不填充公式、不修改工作簿业务逻辑。
- Refresh Excel through the locally installed Microsoft Excel desktop app. / 通过本机安装的 Microsoft Excel 桌面版执行刷新。

## Bundled Resources / 内置资源

- Read `references/PRD.md` before changing product behavior, UI scope, packaging, scheduling, or privacy messaging. / 在修改产品行为、界面范围、打包方式、定时逻辑或隐私文案前，先阅读 `references/PRD.md`。
- Use `scripts/excel_auto_refresh_app.py` as the reference implementation. / 使用 `scripts/excel_auto_refresh_app.py` 作为参考实现。
- Use `scripts/build.ps1` to build a single-file `.exe` with PyInstaller. / 使用 `scripts/build.ps1` 通过 PyInstaller 构建单文件 `.exe`。
- Use `scripts/requirements.txt` for Python dependencies. / 使用 `scripts/requirements.txt` 安装 Python 依赖。

## Implementation Workflow / 实施流程

1. Confirm the target machine is Windows and has Microsoft Excel desktop installed. / 确认目标电脑是 Windows，并已安装 Microsoft Excel 桌面版。
2. Install Python dependencies from `scripts/requirements.txt` in a virtual environment. / 在虚拟环境中安装 `scripts/requirements.txt` 中的 Python 依赖。
3. Run the app during development. / 开发阶段运行应用：

```powershell
python scripts/excel_auto_refresh_app.py
```

4. Build the distributable. / 构建可分发文件：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build.ps1
```

5. Test the resulting `dist\ExcelAutoRefresh.exe`. / 测试生成的 `dist\ExcelAutoRefresh.exe`：
   - Add two workbook tasks. / 添加两个工作簿任务。
   - Run one task immediately. / 立即运行其中一个任务。
   - Confirm Windows Task Scheduler contains one task per enabled workbook. / 确认 Windows 任务计划程序中每个启用的工作簿都有对应任务。
   - Confirm the app can pause, resume, edit, and delete tasks. / 确认应用可以暂停、恢复、编辑和删除任务。
   - Confirm logs are written under `%APPDATA%\ExcelAutoRefresh\logs`. / 确认日志写入 `%APPDATA%\ExcelAutoRefresh\logs`。

## Scheduler Notes / 定时任务说明

Each configured workbook has a stable task id and a Windows scheduled task named:

每个已配置的工作簿都有一个稳定的任务 ID，并对应一个 Windows 计划任务，命名格式如下：

```text
ExcelAutoRefresh_<task-id>
```

The scheduled task runs the same executable in background mode:

计划任务会以后台模式运行同一个可执行文件：

```text
ExcelAutoRefresh.exe --run-task <task-id>
```

When developing from source, the app creates a task that runs:

从源码开发时，应用创建的计划任务会运行：

```text
python excel_auto_refresh_app.py --run-task <task-id>
```

## Refresh Behavior / 刷新行为

The background task should:

后台任务应执行以下操作：

1. Open the workbook with Excel COM automation. / 使用 Excel COM 自动化打开工作簿。
2. Call `RefreshAll`. / 调用 `RefreshAll`。
3. Wait for async queries when Excel exposes `CalculateUntilAsyncQueriesDone`. / 如果 Excel 支持 `CalculateUntilAsyncQueriesDone`，等待异步查询完成。
4. Refresh pivot caches and pivot tables. / 刷新数据透视缓存和数据透视表。
5. Calculate the workbook. / 重新计算工作簿。
6. Save and close the workbook. / 保存并关闭工作簿。
7. Write a local success or failure log. / 写入本地成功或失败日志。

Do not inspect or export cell-level business data unless a future user explicitly asks for diagnostics that require it.

除非未来用户明确要求进行必须读取内容的诊断，否则不要检查或导出单元格级业务数据。
