import argparse
import ctypes
import json
import os
import subprocess
import sys
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from tkinter import (
    BooleanVar,
    Button,
    Checkbutton,
    END,
    Entry,
    Frame,
    Label,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
)
from tkinter import ttk

APP_NAME = "ExcelAutoRefresh"
TASK_PREFIX = "ExcelAutoRefresh_"
SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xlsb", ".xls")


def text(en: str, zh: str) -> str:
    return f"{en} / {zh}"


STATUS_LABELS = {
    "Enabled": text("Enabled", "已启用"),
    "Paused": text("Paused", "已暂停"),
    "Success": text("Success", "成功"),
    "Failed": text("Failed", "失败"),
    "Never run": text("Never run", "从未运行"),
}


def app_dir() -> Path:
    root = os.environ.get("APPDATA")
    if not root:
        root = str(Path.home())
    path = Path(root) / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    (path / "logs").mkdir(parents=True, exist_ok=True)
    return path


CONFIG_PATH = app_dir() / "config.json"
LOG_DIR = app_dir() / "logs"


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return {"tasks": []}
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if "tasks" not in data or not isinstance(data["tasks"], list):
            return {"tasks": []}
        return data
    except Exception:
        return {"tasks": []}


def save_config(config: dict) -> None:
    tmp = CONFIG_PATH.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    tmp.replace(CONFIG_PATH)


def find_task(config: dict, task_id: str) -> dict | None:
    for task in config.get("tasks", []):
        if task.get("id") == task_id:
            return task
    return None


def task_name(task_id: str) -> str:
    return f"{TASK_PREFIX}{task_id}"


def quote_ps(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def executable_command(task_id: str) -> str:
    if getattr(sys, "frozen", False):
        exe = sys.executable
        return f'"{exe}" --run-task {task_id}'
    python = sys.executable
    script = Path(__file__).resolve()
    return f'"{python}" "{script}" --run-task {task_id}'


def run_powershell(script: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )


def register_scheduled_task(task: dict) -> tuple[bool, str]:
    delete_scheduled_task(task.get("id", ""))
    if not task.get("enabled", True):
        return True, text("Task is paused.", "任务已暂停。")

    name = task_name(task["id"])
    time_text = task["time"]
    command = executable_command(task["id"])
    ps = f"""
$Action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument '/c {command}'
$Trigger = New-ScheduledTaskTrigger -Daily -At {quote_ps(time_text)}
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName {quote_ps(name)} -Action $Action -Trigger $Trigger -Settings $Settings -Description 'Refresh an Excel workbook locally with Excel Auto Refresh.' -Force | Out-Null
"""
    result = run_powershell(ps)
    if result.returncode != 0:
        return False, (result.stderr or result.stdout or text("Failed to create scheduled task.", "创建计划任务失败。")).strip()
    return True, text("Scheduled.", "已创建计划任务。")


def delete_scheduled_task(task_id: str) -> None:
    if not task_id:
        return
    name = task_name(task_id)
    ps = f"Unregister-ScheduledTask -TaskName {quote_ps(name)} -Confirm:$false -ErrorAction SilentlyContinue"
    run_powershell(ps)


def validate_time(value: str) -> bool:
    try:
        datetime.strptime(value, "%H:%M")
        return True
    except ValueError:
        return False


def log(task_id: str, message: str) -> None:
    path = LOG_DIR / f"{task_id}.log"
    with path.open("a", encoding="utf-8") as f:
        f.write(f"[{now_text()}] {message}\n")


def set_task_result(task_id: str, status: str, message: str) -> None:
    config = load_config()
    task = find_task(config, task_id)
    if task:
        task["last_run_at"] = now_text()
        task["last_status"] = status
        task["last_message"] = message[:500]
        task["updated_at"] = now_text()
        save_config(config)


def refresh_workbook(task: dict) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(text("Missing pywin32. Rebuild the app with requirements installed.", "缺少 pywin32。请安装依赖后重新构建应用。")) from exc

    file_path = task["file_path"]
    if not Path(file_path).exists():
        raise FileNotFoundError(f"{text('Workbook not found', '找不到工作簿')}: {file_path}")

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        workbook = excel.Workbooks.Open(
            file_path,
            UpdateLinks=3,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )

        workbook.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        try:
            for cache in workbook.PivotCaches():
                cache.Refresh()
        except Exception:
            pass

        for worksheet in workbook.Worksheets:
            try:
                for pivot_table in worksheet.PivotTables():
                    pivot_table.RefreshTable()
            except Exception:
                continue

        try:
            excel.CalculateFullRebuild()
        except Exception:
            excel.Calculate()

        workbook.Save()
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def run_task(task_id: str) -> int:
    config = load_config()
    task = find_task(config, task_id)
    if not task:
        log(task_id, "FAILED: Task is missing from config.")
        return 1
    if not task.get("enabled", True):
        log(task_id, "SKIPPED: Task is paused.")
        return 0
    try:
        log(task_id, f"START: {task.get('file_path')}")
        refresh_workbook(task)
        set_task_result(task_id, "Success", text("Workbook refreshed successfully.", "工作簿刷新成功。"))
        log(task_id, "SUCCESS: Workbook refreshed successfully.")
        return 0
    except Exception as exc:
        detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
        set_task_result(task_id, "Failed", detail)
        log(task_id, f"FAILED: {detail}")
        log(task_id, traceback.format_exc())
        return 1


class TaskDialog:
    def __init__(self, parent: Tk, task: dict | None = None):
        self.result = None
        self.window = Toplevel(parent)
        self.window.title(text("Excel refresh task", "Excel 刷新任务"))
        self.window.resizable(False, False)
        self.window.grab_set()

        self.path_var = StringVar(value=(task or {}).get("file_path", ""))
        self.time_var = StringVar(value=(task or {}).get("time", "09:00"))
        self.enabled_var = BooleanVar(value=(task or {}).get("enabled", True))

        Label(self.window, text=text("Excel file", "Excel 文件")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        path_row = Frame(self.window)
        path_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12)
        Entry(path_row, textvariable=self.path_var, width=54).pack(side="left", fill="x", expand=True)
        Button(path_row, text=text("Browse", "浏览"), command=self.browse).pack(side="left", padx=(8, 0))

        Label(self.window, text=text("Daily refresh time (24-hour HH:MM)", "每日刷新时间（24小时制 HH:MM）")).grid(row=2, column=0, sticky="w", padx=12, pady=(12, 4))
        Entry(self.window, textvariable=self.time_var, width=12).grid(row=3, column=0, sticky="w", padx=12)

        Checkbutton(self.window, text=text("Enable this task", "启用此任务"), variable=self.enabled_var).grid(row=4, column=0, sticky="w", padx=12, pady=12)

        buttons = Frame(self.window)
        buttons.grid(row=5, column=0, columnspan=2, sticky="e", padx=12, pady=(0, 12))
        Button(buttons, text=text("Cancel", "取消"), command=self.window.destroy).pack(side="right")
        Button(buttons, text=text("Save", "保存"), command=self.save).pack(side="right", padx=(0, 8))

    def browse(self) -> None:
        path = filedialog.askopenfilename(
            title=text("Choose an Excel workbook", "选择一个 Excel 工作簿"),
            filetypes=[(text("Excel workbooks", "Excel 工作簿"), "*.xlsx *.xlsm *.xlsb *.xls"), (text("All files", "所有文件"), "*.*")],
        )
        if path:
            self.path_var.set(path)

    def save(self) -> None:
        path = self.path_var.get().strip().strip('"')
        time_value = self.time_var.get().strip()
        if not path:
            messagebox.showerror(text("Missing file", "缺少文件"), text("Please choose an Excel file.", "请选择一个 Excel 文件。"))
            return
        if Path(path).suffix.lower() not in SUPPORTED_EXTENSIONS:
            messagebox.showerror(text("Unsupported file", "不支持的文件"), text("Please choose an Excel workbook file.", "请选择一个 Excel 工作簿文件。"))
            return
        if not Path(path).exists():
            messagebox.showerror(text("File not found", "找不到文件"), text("The selected file does not exist.", "所选文件不存在。"))
            return
        if not validate_time(time_value):
            messagebox.showerror(text("Invalid time", "时间格式无效"), text("Use 24-hour time in HH:MM format, for example 09:30.", "请使用 24 小时制 HH:MM 格式，例如 09:30。"))
            return
        self.result = {
            "file_path": path,
            "time": time_value,
            "enabled": bool(self.enabled_var.get()),
        }
        self.window.destroy()


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(text("Excel Auto Refresh", "Excel 自动刷新"))
        self.root.geometry("1240x760")
        self.root.minsize(1080, 660)
        self.config = load_config()
        self.configure_style()

        self.root.configure(bg="#f4f6f8")

        shell = Frame(self.root, bg="#f4f6f8")
        shell.pack(fill="both", expand=True)

        self.sidebar = Frame(shell, bg="#fbfcfd", width=220, highlightbackground="#e5e9ef", highlightthickness=1)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = Frame(self.sidebar, bg="#fbfcfd")
        brand.pack(fill="x", padx=18, pady=(18, 22))
        Label(brand, text="▣", bg="#fbfcfd", fg="#138a55", font=("Segoe UI Symbol", 17, "bold")).pack(side="left")
        Label(
            brand,
            text=text("Excel Auto Refresh", "Excel 自动刷新"),
            bg="#fbfcfd",
            fg="#162033",
            font=("Segoe UI", 12, "bold"),
            anchor="w",
        ).pack(side="left", padx=(10, 0))

        self.make_nav_button(text("Task list", "任务列表"), "◷", active=True).pack(fill="x", padx=16, pady=(0, 10))
        self.make_nav_button(text("Run logs", "运行日志"), "▤", command=self.open_logs).pack(fill="x", padx=16, pady=10)
        self.make_nav_button(text("Settings", "设置"), "⚙", command=self.show_settings).pack(fill="x", padx=16, pady=10)
        self.make_nav_button(text("About", "关于"), "ⓘ", command=self.show_about).pack(fill="x", padx=16, pady=10)

        sidebar_spacer = Frame(self.sidebar, bg="#fbfcfd")
        sidebar_spacer.pack(fill="both", expand=True)

        safety_card = Frame(self.sidebar, bg="#edf8f1", highlightbackground="#d9efe3", highlightthickness=1)
        safety_card.pack(fill="x", padx=18, pady=(0, 26))
        Label(safety_card, text="◇", bg="#edf8f1", fg="#118352", font=("Segoe UI Symbol", 18, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        Label(safety_card, text="本地运行 · 安全可靠", bg="#edf8f1", fg="#102033", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16)
        Label(
            safety_card,
            text="应用仅在本机打开 Excel\n不会上传文件或脚本\n所有文件保留在本机",
            bg="#edf8f1",
            fg="#607081",
            justify="left",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=16, pady=(8, 14))

        footer = Frame(self.sidebar, bg="#fbfcfd")
        footer.pack(fill="x", padx=18, pady=(0, 18))
        Label(footer, text="v1.0.0", bg="#fbfcfd", fg="#667085", font=("Segoe UI", 9)).pack(side="left")
        Label(footer, text="●", bg="#fbfcfd", fg="#20b15a", font=("Segoe UI", 9)).pack(side="left", padx=(48, 6))
        Label(footer, text="就绪", bg="#fbfcfd", fg="#667085", font=("Segoe UI", 9)).pack(side="left")

        main = Frame(shell, bg="#f4f6f8")
        main.pack(side="left", fill="both", expand=True, padx=(0, 0))

        content = Frame(main, bg="#ffffff", highlightbackground="#edf0f4", highlightthickness=1)
        content.pack(fill="both", expand=True, padx=0, pady=0)

        hero = Frame(content, bg="#ffffff")
        hero.pack(fill="x", padx=26, pady=(30, 18))
        Label(hero, text="▣↻", bg="#ffffff", fg="#13905a", font=("Segoe UI Symbol", 34, "bold")).pack(side="left", padx=(0, 18))

        title_block = Frame(hero, bg="#ffffff")
        title_block.pack(side="left", fill="x", expand=True)
        Label(title_block, text=text("Excel Auto Refresh", "Excel 自动刷新"), bg="#ffffff", fg="#101828", font=("Segoe UI", 24, "bold")).pack(anchor="w")
        Label(
            title_block,
            text="无需脚本或上传文件，即可在本机定时刷新 Excel",
            bg="#ffffff",
            fg="#667085",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(6, 0))

        Label(
            hero,
            text=text("▭  Local only", "仅本机运行"),
            bg="#edf8f1",
            fg="#117a4b",
            font=("Segoe UI", 10, "bold"),
            padx=18,
            pady=9,
        ).pack(side="right", padx=(18, 0))

        notice = Frame(content, bg="#f0f8f4")
        notice.pack(fill="x", padx=26, pady=(0, 24))
        Label(notice, text="◇", bg="#f0f8f4", fg="#118352", font=("Segoe UI Symbol", 19, "bold")).pack(side="left", padx=(22, 12), pady=15)
        Label(
            notice,
            text="文件保留在本机。 应用只会在本地打开 Excel 来刷新并保存工作簿。",
            bg="#f0f8f4",
            fg="#10804f",
            font=("Segoe UI", 12, "bold"),
        ).pack(side="left")

        table_card = Frame(content, bg="#ffffff", highlightbackground="#dfe5ec", highlightthickness=1)
        table_card.pack(fill="both", expand=True, padx=26, pady=(0, 22))

        columns = ("file", "time", "state", "last_run", "result", "actions")
        self.tree = ttk.Treeview(table_card, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("file", text=text("Workbook", "工作簿"))
        self.tree.heading("time", text=text("Time", "时间"))
        self.tree.heading("state", text=text("Status", "状态"))
        self.tree.heading("last_run", text=text("Last run", "上次运行"))
        self.tree.heading("result", text=text("Result", "结果"))
        self.tree.heading("actions", text=text("Actions", "操作"))
        self.tree.column("file", width=310, minwidth=220, stretch=True)
        self.tree.column("time", width=120, minwidth=90, anchor="center", stretch=False)
        self.tree.column("state", width=145, minwidth=120, anchor="center", stretch=False)
        self.tree.column("last_run", width=170, minwidth=140, anchor="center", stretch=False)
        self.tree.column("result", width=145, minwidth=110, anchor="center", stretch=False)
        self.tree.column("actions", width=150, minwidth=130, anchor="center", stretch=False)
        self.tree.bind("<Double-Button-1>", lambda _event: self.edit_task())

        self.y_scroll = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.y_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.y_scroll.grid(row=0, column=1, sticky="ns")
        table_card.columnconfigure(0, weight=1)
        table_card.rowconfigure(0, weight=1)

        self.empty_frame = Frame(table_card, bg="#ffffff")
        Label(self.empty_frame, text="▤◷", bg="#ffffff", fg="#0f8554", font=("Segoe UI Symbol", 42, "bold")).pack(pady=(46, 8))
        Label(self.empty_frame, text="暂无刷新任务", bg="#ffffff", fg="#101828", font=("Segoe UI", 16, "bold")).pack()
        Label(
            self.empty_frame,
            text="点击下方“添加任务”来选择 Excel 文件并设置每日刷新时间。",
            bg="#ffffff",
            fg="#667085",
            font=("Segoe UI", 10),
        ).pack(pady=(8, 18))
        self.make_action_button(self.empty_frame, "＋  添加任务", self.add_task, primary=True).pack()

        self.actions = Frame(content, bg="#ffffff")
        self.actions.pack(fill="x", padx=26, pady=(0, 26))
        self.make_action_button(self.actions, "＋  添加任务", self.add_task, primary=True).pack(side="left")

        task_actions = Frame(self.actions, bg="#ffffff")
        task_actions.pack(side="left", padx=(12, 0))
        self.make_action_button(task_actions, "▷  立即运行", self.run_now).pack(side="left", padx=(0, 8))
        self.make_action_button(task_actions, "✎  编辑", self.edit_task).pack(side="left", padx=(0, 8))
        self.make_action_button(task_actions, "Ⅱ  暂停", self.toggle_task).pack(side="left", padx=(0, 8))
        self.make_action_button(task_actions, "▷  启用", self.toggle_task).pack(side="left", padx=(0, 8))
        self.make_action_button(task_actions, "▢  删除", self.delete_task, danger=True).pack(side="left", padx=(0, 8))

        Frame(self.actions, bg="#e3e8ef", width=1, height=38).pack(side="left", padx=(20, 14))
        utility_actions = Frame(self.actions, bg="#ffffff")
        utility_actions.pack(side="right")
        self.make_action_button(utility_actions, "▤  日志", self.open_logs).pack(side="left", padx=(0, 10))
        self.make_action_button(utility_actions, "⟳  刷新", self.reload, outline=True).pack(side="left")

        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        self.refresh_list()

    def configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        self.root.configure(bg="#f4f6f8")
        style.configure("Treeview", rowheight=34, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff", borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), foreground="#445066", background="#ffffff", padding=(8, 12))
        style.map("Treeview", background=[("selected", "#e9f6ee")], foreground=[("selected", "#101828")])

    def make_nav_button(self, label: str, icon: str, command=None, active: bool = False) -> Button:
        bg = "#edf7f1" if active else "#fbfcfd"
        fg = "#0f8554" if active else "#2f3a4d"
        return Button(
            self.sidebar,
            text=f"{icon}   {label}",
            command=command,
            anchor="w",
            bg=bg,
            fg=fg,
            activebackground="#edf7f1",
            activeforeground="#0f8554",
            relief="flat",
            bd=0,
            padx=16,
            pady=13,
            font=("Segoe UI", 11, "bold" if active else "normal"),
            cursor="hand2",
        )

    def make_action_button(self, parent, label: str, command, primary: bool = False, danger: bool = False, outline: bool = False) -> Button:
        if primary:
            bg, fg, border = "#0f8a55", "#ffffff", "#0f8a55"
            active_bg = "#0b7446"
        elif danger:
            bg, fg, border = "#ffffff", "#d93a3a", "#e5e9ef"
            active_bg = "#fff2f2"
        elif outline:
            bg, fg, border = "#ffffff", "#118352", "#a8d8bf"
            active_bg = "#edf8f1"
        else:
            bg, fg, border = "#ffffff", "#175b3d", "#e1e6ed"
            active_bg = "#f4faf7"
        return Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=border,
            padx=17,
            pady=10,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )

    def show_settings(self) -> None:
        messagebox.showinfo("设置 / Settings", "当前版本暂无额外设置。\nSettings will be added in a later version.")

    def show_about(self) -> None:
        messagebox.showinfo(
            "关于 / About",
            "Excel Auto Refresh v1.0.0\n\n本地 Excel 自动刷新工具。\nFiles stay on this computer.",
        )

    def reload(self) -> None:
        self.config = load_config()
        self.refresh_list()

    def selected_task(self) -> dict | None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo(text("Choose a task", "选择任务"), text("Please choose a task first.", "请先选择一个任务。"))
            return None
        task_id = selection[0]
        task = find_task(self.config, task_id)
        if not task:
            messagebox.showinfo(text("Task missing", "任务不存在"), text("Please refresh the view and try again.", "请刷新视图后重试。"))
        return task

    def refresh_list(self) -> None:
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)
        for task in self.config.get("tasks", []):
            status = STATUS_LABELS["Enabled"] if task.get("enabled", True) else STATUS_LABELS["Paused"]
            last_status = STATUS_LABELS.get(task.get("last_status", "Never run"), task.get("last_status", "Never run"))
            last_run = task.get("last_run_at", "-")
            name = Path(task.get("file_path", "")).name or task.get("file_path", "")
            self.tree.insert(
                "",
                END,
                iid=task["id"],
                values=(name, task.get("time", "--:--"), status, last_run or "-", last_status, "Use buttons below / 使用下方按钮"),
            )
        if self.config.get("tasks"):
            self.empty_frame.place_forget()
            self.tree.grid()
            self.y_scroll.grid()
        else:
            self.tree.grid_remove()
            self.y_scroll.grid_remove()
            self.empty_frame.place(relx=0.5, rely=0.5, anchor="center")

    def add_task(self) -> None:
        dialog = TaskDialog(self.root)
        self.root.wait_window(dialog.window)
        if not dialog.result:
            return
        task = {
            "id": uuid.uuid4().hex[:12],
            "created_at": now_text(),
            "updated_at": now_text(),
            "last_run_at": "",
            "last_status": "Never run",
            "last_message": "",
            **dialog.result,
        }
        ok, message = register_scheduled_task(task)
        if not ok:
            messagebox.showerror(text("Schedule failed", "计划任务失败"), message)
            return
        self.config["tasks"].append(task)
        save_config(self.config)
        self.refresh_list()

    def edit_task(self) -> None:
        task = self.selected_task()
        if not task:
            return
        dialog = TaskDialog(self.root, task)
        self.root.wait_window(dialog.window)
        if not dialog.result:
            return
        task.update(dialog.result)
        task["updated_at"] = now_text()
        ok, message = register_scheduled_task(task)
        if not ok:
            messagebox.showerror(text("Schedule failed", "计划任务失败"), message)
            return
        save_config(self.config)
        self.refresh_list()

    def toggle_task(self) -> None:
        task = self.selected_task()
        if not task:
            return
        task["enabled"] = not task.get("enabled", True)
        task["updated_at"] = now_text()
        ok, message = register_scheduled_task(task)
        if not ok:
            messagebox.showerror(text("Schedule failed", "计划任务失败"), message)
            return
        save_config(self.config)
        self.refresh_list()

    def delete_task(self) -> None:
        task = self.selected_task()
        if not task:
            return
        if not messagebox.askyesno(text("Delete task", "删除任务"), text("Delete this refresh task?", "要删除这个刷新任务吗？")):
            return
        delete_scheduled_task(task["id"])
        self.config["tasks"] = [item for item in self.config["tasks"] if item["id"] != task["id"]]
        save_config(self.config)
        self.refresh_list()

    def run_now(self) -> None:
        task = self.selected_task()
        if not task:
            return
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        code = run_task(task["id"])
        self.root.config(cursor="")
        self.reload()
        if code == 0:
            messagebox.showinfo(text("Refresh complete", "刷新完成"), text("Workbook refreshed successfully.", "工作簿刷新成功。"))
        else:
            messagebox.showerror(text("Refresh failed", "刷新失败"), text("Refresh failed. Use View logs for details.", "刷新失败。请点击“查看日志”了解详情。"))

    def open_logs(self) -> None:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(LOG_DIR)

    def run(self) -> None:
        if not is_windows():
            messagebox.showerror(text("Windows required", "需要 Windows"), text("This app requires Windows and Microsoft Excel desktop.", "此应用需要 Windows 和 Microsoft Excel 桌面版。"))
            return
        self.root.mainloop()


def is_windows() -> bool:
    return sys.platform.startswith("win")


def hide_console_when_frozen() -> None:
    if getattr(sys, "frozen", False) and is_windows():
        try:
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
        except Exception:
            pass


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-task", help="Run a configured task by id.")
    args = parser.parse_args()

    if args.run_task:
        return run_task(args.run_task)

    hide_console_when_frozen()
    App().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
