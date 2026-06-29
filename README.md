# Excel Auto Refresh / Excel 自动刷新

A local Windows desktop app for scheduling automatic Excel refresh tasks. It is designed for non-technical users who need to refresh workbooks, Power Query connections, external data, and pivot tables without writing VBA or scripts.

这是一个 Windows 本地桌面应用，用于定时自动刷新 Excel。它面向非技术用户，适合在不编写 VBA 或脚本的情况下刷新工作簿、Power Query、外部数据连接和数据透视表。

## Key Features / 主要功能

- Multiple Excel refresh tasks / 支持多个 Excel 刷新任务
- Daily scheduled refreshes / 支持每日定时刷新
- Run missed tasks when the computer becomes available / 电脑可用后补跑错过的任务
- Local-only workflow with no file upload / 仅本机运行，不上传文件
- Bilingual English and Chinese UI / 中英双语界面
- Single-file Windows `.exe` build / 可构建为单个 Windows `.exe`

## Privacy / 隐私说明

Files stay on the user's computer. The app only opens Microsoft Excel locally to refresh and save workbooks.

文件保留在用户电脑本机。应用只会在本地调用 Microsoft Excel 打开、刷新并保存工作簿。

The app does not:

应用不会：

- Upload workbooks / 上传工作簿
- Copy workbook contents elsewhere / 将工作簿内容复制到其他位置
- Unlock or protect sheets / 解除或重新保护工作表
- Fill formulas / 填充公式
- Parse business data from cells / 解析单元格中的业务数据

## Project Structure / 项目结构

```text
.
├── README.md
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   └── PRD.md
└── scripts/
    ├── build.ps1
    ├── excel_auto_refresh_app.py
    └── requirements.txt
```

## Run From Source / 从源码运行

```powershell
cd scripts
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe excel_auto_refresh_app.py
```

## Build EXE / 构建 EXE

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build.ps1
```

The output will be:

输出文件为：

```text
scripts\dist\ExcelAutoRefresh.exe
```

`scripts\dist` is generated locally and is ignored by git by default.

`scripts\dist` 是本地构建生成目录，默认不会提交到 git。

## Requirements / 环境要求

- Windows
- Microsoft Excel desktop app
- Python 3.10+ for development or rebuilding

- Windows 操作系统
- Microsoft Excel 桌面版
- 如果需要开发或重新构建，需要 Python 3.10+

## Notes / 说明

The app uses Windows Task Scheduler so refreshes continue even after the main window is closed.

应用使用 Windows 任务计划程序，因此主窗口关闭后，定时刷新仍可继续执行。
