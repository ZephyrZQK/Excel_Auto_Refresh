# PRD: Excel Auto Refresh / Excel 自动刷新

## 1. Product Summary / 产品概述

Excel Auto Refresh is a Windows desktop app for non-technical users who need scheduled local refreshes for Excel workbooks, data connections, Power Query outputs, and pivot tables. Users configure one or more Excel files, choose daily refresh times, and let Windows run the refresh automatically.

Excel Auto Refresh 是一个 Windows 桌面应用，面向需要在本地定时刷新 Excel 工作簿、数据连接、Power Query 输出和数据透视表的非技术用户。用户可以配置一个或多个 Excel 文件，选择每日刷新时间，然后由 Windows 自动执行刷新。

The product must reduce anxiety for users who do not write code and for teams whose workbooks contain confidential data. The app must run locally, package as a single `.exe`, and avoid uploading, copying, or parsing workbook business content.

产品必须降低非技术用户的使用压力，并适用于工作簿包含涉密数据的团队。应用必须在本地运行，打包为单个 `.exe`，并避免上传、复制或解析工作簿中的业务内容。

## 2. Goals / 目标

- Let users configure multiple Excel workbook refresh tasks without code. / 让用户无需代码即可配置多个 Excel 工作簿刷新任务。
- Package the tool as one Windows `.exe`. / 将工具打包为单个 Windows `.exe`。
- Run scheduled refreshes even when the main app is closed. / 即使主应用关闭，也能执行定时刷新。
- Run missed scheduled refreshes as soon as possible after the computer becomes available. / 如果错过定时任务，在电脑可用后尽快补跑。
- Keep workbook contents on the local machine. / 工作簿内容始终保留在本机。
- Provide clear success/failure status and readable local logs. / 提供清晰的成功/失败状态和易读的本地日志。

## 3. Non-Goals / 非目标

- No cloud service, server, or workbook upload. / 不提供云服务、服务器或工作簿上传功能。
- No workbook content analysis. / 不分析工作簿内容。
- No sheet protection management. / 不管理工作表保护。
- No formula fill-down behavior. / 不执行公式下拉填充。
- No cross-platform support in the first version. / 第一版不支持跨平台。
- No support for Excel Online only environments. / 不支持仅有 Excel Online 的环境。

## 4. Target Users / 目标用户

- Business analysts and account teams who own recurring Excel reports. / 负责周期性 Excel 报告的业务分析师和客户团队。
- Managers who receive Excel reports but do not write VBA or scripts. / 接收 Excel 报告但不编写 VBA 或脚本的管理者。
- Teams handling confidential data where upload-based automation is not acceptable. / 处理涉密数据且不能接受上传式自动化的团队。

## 5. User Experience / 用户体验

### Main Window / 主窗口

The main window displays a task list:

主窗口展示任务列表：

- Workbook name or path / 工作簿名或路径
- Daily refresh time / 每日刷新时间
- Enabled/paused status / 启用或暂停状态
- Last run time / 上次运行时间
- Last result / 上次运行结果
- Actions: Run Now, Edit, Pause/Enable, Delete / 操作：立即运行、编辑、暂停/启用、删除

The window includes a short privacy reassurance:

窗口中包含简短的隐私提示：

> Files stay on this computer. The app only opens Excel locally to refresh and save workbooks.
>
> 文件保留在本机。应用只会在本地打开 Excel 来刷新并保存工作簿。

### Add/Edit Task / 添加或编辑任务

The user sees a small form:

用户看到一个简单表单：

- Select Excel file / 选择 Excel 文件
- Set daily refresh time / 设置每日刷新时间
- Enable this task / 启用此任务
- Save or Cancel / 保存或取消

Supported extensions: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`.

支持的扩展名：`.xlsx`、`.xlsm`、`.xlsb`、`.xls`。

### Failure Experience / 失败体验

The main window shows a simple failure status. Detailed technical information is behind a "View Logs" action and stored locally.

主窗口只显示简单的失败状态。详细技术信息通过“查看日志”入口查看，并存储在本地。

## 6. Functional Requirements / 功能需求

### Task Management / 任务管理

- Create multiple tasks. / 创建多个任务。
- Edit file path and daily time. / 编辑文件路径和每日时间。
- Pause and resume tasks. / 暂停和恢复任务。
- Delete tasks. / 删除任务。
- Run a task immediately for testing. / 立即运行任务用于测试。
- Persist tasks in `%APPDATA%\ExcelAutoRefresh\config.json`. / 将任务持久化到 `%APPDATA%\ExcelAutoRefresh\config.json`。

### Scheduling / 定时

- Use Windows Task Scheduler. / 使用 Windows 任务计划程序。
- Create one Windows scheduled task per enabled workbook task. / 每个启用的工作簿任务创建一个 Windows 计划任务。
- Use the same executable for scheduled background execution. / 使用同一个可执行文件进行后台定时执行。
- Configure missed runs to start when available. / 配置错过的任务在可用时启动。
- Remove the scheduled task when a workbook task is deleted. / 删除工作簿任务时同步移除计划任务。
- Disable or remove the scheduled task when a workbook task is paused. / 暂停工作簿任务时禁用或移除计划任务。

### Refresh Execution / 刷新执行

- Open the workbook using the locally installed Excel desktop app. / 使用本机安装的 Excel 桌面版打开工作簿。
- Run `RefreshAll`. / 执行 `RefreshAll`。
- Wait for asynchronous queries if supported. / 如果支持，则等待异步查询完成。
- Refresh pivot caches and pivot tables. / 刷新数据透视缓存和数据透视表。
- Calculate the workbook. / 重新计算工作簿。
- Save and close the workbook. / 保存并关闭工作簿。
- Log success or failure locally. / 在本地记录成功或失败日志。

### Privacy / 隐私

- Do not upload workbooks. / 不上传工作簿。
- Do not copy workbooks elsewhere. / 不将工作簿复制到其他位置。
- Do not store workbook contents in logs. / 不在日志中保存工作簿内容。
- Logs may include file path, task id, timestamps, and exception messages. / 日志可以包含文件路径、任务 ID、时间戳和异常信息。

## 7. Technical Requirements / 技术需求

- Language: Python. / 语言：Python。
- UI: Tkinter. / 界面：Tkinter。
- Excel automation: `pywin32`. / Excel 自动化：`pywin32`。
- Packaging: PyInstaller single-file `.exe`. / 打包：PyInstaller 单文件 `.exe`。
- OS: Windows. / 操作系统：Windows。
- Excel dependency: Installed Microsoft Excel desktop app. / Excel 依赖：已安装 Microsoft Excel 桌面版。

## 8. Data Storage / 数据存储

Configuration path:

配置路径：

```text
%APPDATA%\ExcelAutoRefresh\config.json
```

Log folder:

日志文件夹：

```text
%APPDATA%\ExcelAutoRefresh\logs
```

Task fields:

任务字段：

- `id`
- `file_path`
- `time`
- `enabled`
- `created_at`
- `updated_at`
- `last_run_at`
- `last_status`
- `last_message`

## 9. Success Metrics / 成功指标

- A non-technical user can add a task and run it immediately without instructions from a developer. / 非技术用户无需开发人员指导即可添加任务并立即运行。
- Scheduled refreshes continue after the app is closed. / 应用关闭后定时刷新仍可继续。
- Missed refreshes run after startup or login. / 错过的刷新可在启动或登录后补跑。
- Users can understand success/failure from the main screen. / 用户可以从主界面理解成功或失败状态。
- Confidential workbook content never leaves the local machine. / 涉密工作簿内容不会离开本机。

## 10. Acceptance Criteria / 验收标准

- The app can be built into one `.exe`. / 应用可以构建为单个 `.exe`。
- The app can add at least three workbook tasks. / 应用至少可以添加三个工作簿任务。
- Each enabled task appears in Windows Task Scheduler. / 每个启用的任务都会出现在 Windows 任务计划程序中。
- "Run Now" refreshes the chosen workbook and updates status. / “立即运行”可以刷新所选工作簿并更新状态。
- Pausing a task prevents scheduled execution. / 暂停任务后不会再定时执行。
- Deleting a task removes its scheduled task. / 删除任务会移除对应计划任务。
- A failed refresh leaves Excel closed where possible and writes a local log. / 刷新失败时尽可能关闭 Excel，并写入本地日志。
- The app does not perform sheet unlock/protect or formula fill operations. / 应用不执行工作表解锁/保护或公式填充操作。
