"""加华油站小助手的桌面界面。"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from calendar import monthcalendar, monthrange
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "加华油站小助手"
STATUSES = (
    "全部状态",
    "金额不符",
    "升数不符",
    "时间先后异常",
    "无对应记录",
    "待确认",
)
HELP_TEXT = """加华油站小助手使用说明

1. 分别选择包含加油、发票数据的工作簿和发票审核单（选择 invoiceList_*.xlsx 原文件，不选 ~$ 开头的临时文件）。输入工作簿会自动备份后直接标注；审核单中被红冲或作废的发票不参与配对和标色。工具按表头识别切页，优先使用带“资料区间”的发票页。发票备注中的站点会出现在“审核站点”，支持多选和全选；仅核对勾选站点。

2. 选择工作簿后，开始、结束日期自动设为加油明细“交易时间”的最早、最晚日期；也可自由选择其他日期。可勾选“指定时刻”选择时、分、秒；仅选日期时，开始取当天 00:00:00，结束包含当天全部时间。点击“全部加油日期”可恢复默认范围。所选时间段没有可核对记录时会提示。

3. 界面“加油明细切页-交易时间”取自加油原表，“发票-交易时间”取自发票备注；“开票时间”另行展示。发票交易时间晚于加油明细交易时间不限间隔；早于开始时间不超过 10 分钟也视为正常，超过 10 分钟则显示“时间先后异常”。

4. 加油表“交易金额”已扣除优惠，直接与发票“合计金额（含税）”按原始数值严格比较，不再重复扣减优惠。加油量与发票数量也必须完全一致；优惠金额仅供查看，不参与核对。

5. 同站同枪同油品同升数的重复候选默认按时间顺序配对；仅在加油后60分钟内金额对应唯一、两侧均无歧义且不减少配对数时，自动纠正交叉配对并注明依据。其余情况保持原顺序；每笔记录只使用一次，重复票号、同票多行或其他证据不足的情况标为待确认。

6. 金额只要有差异，即使相差 0.01 元，也判为“金额不符”。付款类型为“记账卡”时，加油行仍标浅黄底；其他加油失败行和未配对发票在原发票切页及油站切页均标浅红底。加油明细末尾新增“失败原因”“关联发票号”，找不到对应发票时票号留空。

7. 备注缺少结束时间时，允许使用开票时间辅助判断并注明来源，日志和预览中的加油结束时间保持空白；备注时间存在但格式无效时标为待确认。备注交易时间、油枪号及时间来源在日志和预览中展示。

8. 点击“打开已标注原文件”直接查看输入工作簿，其中新增所选油站切页，未配对发票整行标红；同目录生成避重名的核对前备份。结果工作簿仍另存到输出目录，保留所选站点及“匹配成功”“匹配失败”“发票未配对”切页，并在发票原切页“备注”右侧新增“交易时间”。结果工作簿可从输出目录打开。重复核对已标注原文件会刷新旧标注和油站切页，不重复添加列；时间段外记录不新增标注。

9. 与结果工作簿同名的 .log 文件保存在输出目录的 logs 文件夹，只保存匹配失败的记录，包括对应位置、金额差、升数差、时间差、失败原因和配对依据。新核对成功后，删除同一输入文件的旧核对日志，保留其他文件。成功记录在“匹配成功”切页中按每对显示，包含油枪号、加油量、油品、加油时间、发票号码等信息。双击预览行可查看完整内容。

10. 程序“匹配失败”切页只显示未通过的加油明细，失败原因单独列出；输入加油原始行号可快速查找，成功与失败列表均按交易时间从早到晚显示，缺失时间的排在最后。点击“对应记录位置”可打开结果 Excel。“仅显示加油异常”和状态下拉框可组合筛选失败记录，状态下拉框只列出加油明细出现的失败状态。失败与成功页均以加油明细总笔数为分母；失败页另标明当前显示的行数。Excel 记账卡整行标黄，其他加油侧异常原行标红；结果工作簿“匹配失败”页也只保留加油侧，日志仍记录两侧失败。“金额差”＝实际发票金额－交易金额；“升数差”＝发票数量－加油量；时间差保留正负号。
"""


def application_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def settings_path() -> Path:
    base = Path(os.environ["LOCALAPPDATA"]) if os.environ.get("LOCALAPPDATA") else Path.home() / ".config"
    return base / APP_TITLE / "settings.json"


def _load_last_path(key: str) -> str:
    try:
        saved = json.loads(settings_path().read_text(encoding="utf-8"))
        path = Path(saved[key])
        return str(path) if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$") else ""
    except (OSError, ValueError, TypeError, KeyError):
        return ""


def load_last_input_path() -> str:
    return _load_last_path("last_input_path")


def load_last_invoice_audit_path() -> str:
    return _load_last_path("last_invoice_audit_path")


def _save_last_path(key: str, path: str) -> None:
    try:
        destination = settings_path()
        destination.parent.mkdir(parents=True, exist_ok=True)
        try:
            saved = json.loads(destination.read_text(encoding="utf-8"))
            if not isinstance(saved, dict):
                saved = {}
        except (OSError, ValueError):
            saved = {}
        saved[key] = str(Path(path).resolve())
        destination.write_text(json.dumps(saved, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def save_last_input_path(path: str) -> None:
    _save_last_path("last_input_path", path)


def save_last_invoice_audit_path(path: str) -> None:
    _save_last_path("last_invoice_audit_path", path)


def application_asset_path(filename: str) -> Path:
    resources = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return resources / "assets" / filename


def application_icon_path() -> Path:
    return application_asset_path("gawa.ico")


def set_application_id() -> None:
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GAWA.FuelInvoice.Reconciler")
        except (AttributeError, OSError):
            pass


def set_window_icon(window: tk.Tk | tk.Toplevel, *, default: bool = False) -> None:
    icon = application_icon_path()
    if icon.is_file():
        if default:
            try:
                window.iconbitmap(default=str(icon))
            except tk.TclError:
                pass
        try:
            window.iconbitmap(str(icon))
        except tk.TclError:
            pass


class DateTimePicker(tk.Toplevel):
    def __init__(self, parent: tk.Tk, variable: tk.StringVar, label: str, anchor: ttk.Button) -> None:
        super().__init__(parent)
        self.variable = variable
        self.title(f"选择{label}")
        self.resizable(False, False)
        self.transient(parent)
        set_window_icon(self)
        current = variable.get()
        chosen = datetime.fromisoformat(current) if current else datetime.combine(date.today(), datetime.min.time())
        selected_day = chosen.date()
        self.selected_date = selected_day
        self.year = tk.StringVar(value=str(selected_day.year))
        self.month = tk.StringVar(value=f"{selected_day.month:02d}")
        self.precise = tk.BooleanVar(value=bool(current and " " in current))
        self.hour = tk.StringVar(value=f"{chosen.hour:02d}")
        self.minute = tk.StringVar(value=f"{chosen.minute:02d}")
        self.second = tk.StringVar(value=f"{chosen.second:02d}")

        panel = ttk.Frame(self, padding=14)
        panel.pack(fill="both", expand=True)
        header = ttk.Frame(panel)
        header.pack(fill="x", pady=(0, 8))
        ttk.Button(header, text="‹", width=3, command=lambda: self._shift_month(-1)).pack(side="left")
        year = ttk.Spinbox(header, textvariable=self.year, from_=1, to=9999, state="readonly", width=6, command=self._change_month)
        year.pack(side="left", padx=(10, 4))
        ttk.Label(header, text="年").pack(side="left")
        month = ttk.Combobox(header, textvariable=self.month, values=tuple(f"{value:02d}" for value in range(1, 13)), state="readonly", width=4)
        self.month_combo = month
        month.pack(side="left", padx=(8, 4))
        ttk.Label(header, text="月").pack(side="left")
        ttk.Button(header, text="›", width=3, command=lambda: self._shift_month(1)).pack(side="right")
        month.bind("<<ComboboxSelected>>", self._change_month)

        self.calendar = ttk.Frame(panel)
        self.calendar.pack(fill="x")
        self._draw_calendar()
        time_row = ttk.Frame(panel)
        time_row.pack(fill="x", pady=(10, 8))
        ttk.Checkbutton(time_row, text="指定时刻", variable=self.precise, command=self._toggle_time).pack(side="left", padx=(0, 8))
        self.time_controls: list[ttk.Combobox] = []
        for index, (variable_part, limit) in enumerate(((self.hour, 24), (self.minute, 60), (self.second, 60))):
            if index:
                ttk.Label(time_row, text=":").pack(side="left")
            control = ttk.Combobox(time_row, textvariable=variable_part, values=tuple(f"{value:02d}" for value in range(limit)), state="readonly", width=3)
            control.pack(side="left")
            self.time_controls.append(control)
        self._toggle_time()

        actions = ttk.Frame(panel)
        actions.pack(fill="x")
        ttk.Button(actions, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(actions, text="确定", style="Primary.TButton", command=self._apply).pack(side="right", padx=(0, 8))
        self.update_idletasks()
        x_position = min(anchor.winfo_rootx(), self.winfo_screenwidth() - self.winfo_reqwidth() - 8)
        y_position = anchor.winfo_rooty() + anchor.winfo_height() + 4
        if y_position + self.winfo_reqheight() > self.winfo_screenheight():
            y_position = anchor.winfo_rooty() - self.winfo_reqheight() - 4
        self.geometry(f"+{max(0, x_position)}+{max(0, y_position)}")
        self.grab_set()
        self.focus_set()

    def _draw_calendar(self) -> None:
        for child in self.calendar.winfo_children():
            child.destroy()
        for column, weekday in enumerate(("一", "二", "三", "四", "五", "六", "日")):
            ttk.Label(self.calendar, text=weekday, anchor="center", width=4).grid(row=0, column=column, pady=(0, 4))
        for row, week in enumerate(monthcalendar(int(self.year.get()), int(self.month.get())), start=1):
            for column, day in enumerate(week):
                if day:
                    selected = day == self.selected_date.day and int(self.month.get()) == self.selected_date.month and int(self.year.get()) == self.selected_date.year
                    ttk.Button(self.calendar, text=str(day), width=3, style="CalendarSelected.TButton" if selected else "Calendar.TButton", command=lambda value=day: self._select_day(value)).grid(row=row, column=column, padx=1, pady=1)

    def _select_day(self, day: int) -> None:
        self.selected_date = date(int(self.year.get()), int(self.month.get()), day)
        self._draw_calendar()

    def _change_month(self, _event: object = None) -> None:
        year, month = int(self.year.get()), int(self.month.get())
        self.selected_date = date(year, month, min(self.selected_date.day, monthrange(year, month)[1]))
        self._draw_calendar()

    def _shift_month(self, step: int) -> None:
        year, month = int(self.year.get()), int(self.month.get()) + step
        if month == 0:
            year, month = year - 1, 12
        elif month == 13:
            year, month = year + 1, 1
        candidate = date(year, month, 1) if 1 <= year <= 9999 else None
        if candidate:
            self.year.set(str(year))
            self.month.set(f"{month:02d}")
            self._change_month()

    def _toggle_time(self) -> None:
        for control in self.time_controls:
            control.configure(state="readonly" if self.precise.get() else "disabled")

    def _apply(self) -> None:
        value = self.selected_date.isoformat()
        if self.precise.get():
            value += f" {self.hour.get()}:{self.minute.get()}:{self.second.get()}"
        self.variable.set(value)
        self.destroy()


class ReconciliationApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        set_window_icon(root, default=True)
        root.geometry("1120x800")
        root.minsize(1000, 720)
        self.base_directory = application_directory()
        self.input_path = tk.StringVar(value=load_last_input_path())
        self.invoice_audit_path = tk.StringVar(value=load_last_invoice_audit_path())
        self.output_directory = tk.StringVar(value=str(self.base_directory / "outputs"))
        self.start_text = tk.StringVar()
        self.end_text = tk.StringVar()
        self.worker_count = tk.StringVar(value="20")
        self.rules_text = tk.StringVar()
        self.red_only = tk.BooleanVar(value=False)
        self.status_filter = tk.StringVar(value=STATUSES[0])
        self.row_search = tk.StringVar()
        self.status_text = tk.StringVar(value="请选择工作簿，核实设置后开始核对。")
        self.summary_text = tk.StringVar(value="尚未核对。结果将分别统计加油记录、发票记录和配对数。")
        self.preview_text = tk.StringVar(value="异常 0 / 0 笔（当前显示 0 行）")
        self.success_text = tk.StringVar(value="匹配成功 0 / 0 笔")
        self.progress_text = tk.StringVar(value="核对进度：尚未开始")
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.results: list[dict] = []
        self.output_path: Path | None = None
        self.annotated_input_path: Path | None = None
        self.input_backup_path: Path | None = None
        self.log_path: Path | None = None
        self.help_window: tk.Toplevel | None = None
        self.setting_controls: list[ttk.Widget] = []
        self.station_vars: dict[str, tk.BooleanVar] = {}
        self.checkbox_images: tuple[tk.PhotoImage, tk.PhotoImage] | None = None
        self.station_path = ""
        self.station_scan_error = ""
        self.invoice_interval: tuple[str, str] | None = None
        self.fuel_date_range: tuple[str, str] | None = None
        self.time_buttons: list[ttk.Button] = []
        self.station_scan_id = 0
        self.search_refresh_id = None
        self.station_text = tk.StringVar(value="选择工作簿后自动识别发票备注中的加油站点")
        self._build()
        self.input_path.trace_add("write", self._schedule_station_scan)
        if self.input_path.get():
            self._schedule_station_scan()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.after(150, self._poll_events)

    def _build(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        page = "#F2F5FA"
        blue = "#1F4E79"
        text = "#273B53"
        self.root.configure(background=page)
        self.root.option_add("*Font", ("Microsoft YaHei UI", 10))
        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#FFFFFF", foreground=text)
        style.configure("Page.TFrame", background=page)
        style.configure("Page.TLabel", background=page, foreground=text)
        style.configure("Rule.TLabel", background=page, foreground=blue)
        style.configure("Footer.TLabel", background=page, foreground="#63778A")
        style.configure("Card.TLabelframe", background="#FFFFFF", bordercolor="#D7E3F0", borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background="#FFFFFF", foreground=blue, font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=text, padding=6, bordercolor="#CAD9E9", lightcolor="#FFFFFF", darkcolor="#FFFFFF")
        style.map("TEntry", bordercolor=[("focus", "#7FA8D0")], foreground=[("disabled", "#8A98A8")])
        style.configure("TCombobox", fieldbackground="#FFFFFF", background="#EEF4FC", foreground=text, padding=6, arrowcolor=blue, bordercolor="#CAD9E9")
        style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], selectbackground=[("readonly", "#FFFFFF")], selectforeground=[("readonly", text)])
        style.layout("TButton", [("Button.border", {"sticky": "nswe", "children": [("Button.padding", {"sticky": "nswe", "children": [("Button.label", {"sticky": "nswe"})]})]})])
        style.configure("TButton", background="#DCEBFC", foreground="#123D68", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 7), borderwidth=2, relief="raised", bordercolor="#7FA8D0", lightcolor="#FFFFFF", darkcolor="#A8C5E5")
        style.map("TButton", background=[("disabled", "#EEF2F6"), ("pressed", "#AFCDED"), ("active", "#C8E0FA")], foreground=[("disabled", "#8A98A8")], bordercolor=[("active", "#437FBC"), ("pressed", "#275D95")], relief=[("pressed", "sunken"), ("!pressed", "raised")])
        style.configure("Primary.TButton", background="#215FDB", foreground="#FFFFFF", bordercolor="#174BAF", lightcolor="#6F9AF0", darkcolor="#103C93")
        style.map("Primary.TButton", background=[("disabled", "#B7C7D8"), ("pressed", "#1243AC"), ("active", "#174FC4")], foreground=[("disabled", "#F2F5FA"), ("!disabled", "#FFFFFF")], bordercolor=[("active", "#103C93")])
        style.configure("Calendar.TButton", padding=(4, 4), font=("Microsoft YaHei UI", 9))
        style.configure("CalendarSelected.TButton", padding=(4, 4), background="#215FDB", foreground="#FFFFFF", bordercolor="#174BAF")
        style.map("CalendarSelected.TButton", background=[("pressed", "#1243AC"), ("active", "#174FC4")], foreground=[("!disabled", "#FFFFFF")])
        style.configure("TCheckbutton", background="#FFFFFF", foreground=text)
        style.map("TCheckbutton", background=[("active", "#FFFFFF")])
        unchecked = tk.PhotoImage(master=self.root, width=18, height=18)
        unchecked.put("#FFFFFF", to=(0, 0, 18, 18))
        unchecked.put("#7FA8D0", to=(1, 1, 17, 17))
        unchecked.put("#FFFFFF", to=(2, 2, 16, 16))
        checked = unchecked.copy()
        for x, y in ((4, 9), (5, 10), (6, 11), (7, 12), (8, 11), (9, 10), (10, 9), (11, 8), (12, 7), (13, 6)):
            checked.put("#215FDB", to=(x, y, x + 2, y + 2))
        self.checkbox_images = (unchecked, checked)
        style.element_create("Checkmark.indicator", "image", unchecked, ("selected", checked))
        style.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
            ("Checkmark.indicator", {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [("Checkbutton.label", {"sticky": "nswe"})]}),
        ]})])
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=text, rowheight=30, borderwidth=0)
        style.map("Treeview", background=[("selected", "#F3AEB9")], foreground=[("selected", text)])
        style.map("Success.Treeview", background=[("selected", "#DDEBFA")], foreground=[("selected", blue)])
        style.configure("Treeview.Heading", background="#EEF4FC", foreground=blue, font=("Microsoft YaHei UI", 10, "bold"), padding=(8, 8), relief="flat")
        style.map("Treeview.Heading", background=[("active", "#DFEBFA")])
        for scrollbar in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            style.configure(scrollbar, background="#D7E4F2", troughcolor="#F1F5FA", bordercolor="#F1F5FA", arrowcolor=blue)
        style.configure("Horizontal.TProgressbar", background="#4E86BA", troughcolor="#EEF4FC", bordercolor="#D7E3F0")
        outer = ttk.Frame(self.root, padding=20, style="Page.TFrame")
        outer.pack(fill="both", expand=True)
        form = ttk.LabelFrame(outer, text="核对设置", padding=(14, 10), style="Card.TLabelframe")
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        for row, label, variable, chooser in (
            (0, "数据工作簿", self.input_path, self._choose_input),
            (1, "发票审核单", self.invoice_audit_path, self._choose_invoice_audit),
            (2, "输出目录", self.output_directory, self._choose_output),
        ):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            entry = ttk.Entry(form, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            button = ttk.Button(form, text="浏览…", command=chooser)
            button.grid(row=row, column=2, padx=(8, 0), pady=5)
            self.setting_controls.extend((entry, button))
        ttk.Label(form, text="核对时间段").grid(row=3, column=0, sticky="nw", padx=(0, 12), pady=5)
        options = ttk.Frame(form)
        options.grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)
        for column, label, variable in ((0, "开始时间", self.start_text), (2, "结束时间", self.end_text)):
            ttk.Label(options, text=label).grid(row=0, column=column, sticky="w", padx=(0 if column == 0 else 20, 8))
            display = tk.StringVar(value=f"{variable.get() or '不限'}  ▾")
            variable.trace_add("write", lambda *_args, source=variable, target=display: target.set(f"{source.get() or '不限'}  ▾"))
            button = ttk.Button(options, textvariable=display, state="disabled")
            button.configure(command=lambda anchor=button, source=variable, title=label: self._choose_datetime(source, title, anchor))
            button.grid(row=0, column=column + 1, sticky="ew")
            self.setting_controls.append(button)
            self.time_buttons.append(button)
        options.columnconfigure(1, weight=1)
        options.columnconfigure(3, weight=1)
        ttk.Label(options, text="点击日期选择；可选具体时刻，未指定时刻则按整天核对。", foreground="#63778A").grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.all_time_button = ttk.Button(options, text="全部加油日期", command=self._clear_period, state="disabled")
        self.all_time_button.grid(row=1, column=3, sticky="e", pady=(6, 0))
        self.setting_controls.append(self.all_time_button)
        workers = ttk.Frame(form)
        workers.grid(row=4, column=1, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Label(workers, text="并发工作数").pack(side="left", padx=(0, 8))
        self.worker_combo = ttk.Combobox(workers, textvariable=self.worker_count, values=tuple(str(count) for count in range(1, 33)), state="readonly", width=5)
        self.worker_combo.pack(side="left")
        self.setting_controls.append(self.worker_combo)
        station_frame = ttk.Frame(form)
        station_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Label(station_frame, text="审核站点").pack(side="left", padx=(0, 12))
        select_all = ttk.Button(station_frame, text="全选", command=lambda: self._set_all_stations(True))
        select_all.pack(side="left")
        clear_all = ttk.Button(station_frame, text="取消全选", command=lambda: self._set_all_stations(False))
        clear_all.pack(side="left", padx=(6, 12))
        self.setting_controls.extend((select_all, clear_all))
        ttk.Label(station_frame, textvariable=self.station_text).pack(side="left")
        station_box = ttk.Frame(form)
        station_box.grid(row=6, column=0, columnspan=3, sticky="ew")
        self.station_canvas = tk.Canvas(station_box, height=54, background="#FFFFFF", highlightthickness=1, highlightbackground="#D7E3F0")
        self.station_canvas.pack(side="left", fill="x", expand=True)
        station_scroll = ttk.Scrollbar(station_box, orient="vertical", command=self.station_canvas.yview)
        station_scroll.pack(side="right", fill="y")
        self.station_canvas.configure(yscrollcommand=station_scroll.set)
        self.station_items = ttk.Frame(self.station_canvas)
        self.station_canvas.create_window((0, 0), window=self.station_items, anchor="nw")
        self.station_items.bind("<Configure>", lambda _event: self.station_canvas.configure(scrollregion=self.station_canvas.bbox("all")))
        actions = ttk.Frame(form)
        actions.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        self.run_button = ttk.Button(actions, text="开始核对", command=self._start_run, style="Primary.TButton")
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(actions, text="打开已标注原文件", command=self._open_result, state="disabled")
        self.open_button.pack(side="left", padx=(8, 0))
        self.log_button = ttk.Button(actions, text="打开核对日志", command=self._open_log, state="disabled")
        self.log_button.pack(side="left", padx=(8, 0))
        self.folder_button = ttk.Button(actions, text="打开输出目录", command=self._open_output)
        self.folder_button.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="使用帮助", command=self._show_help).pack(side="right")
        progress_row = ttk.Frame(form)
        progress_row.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        ttk.Label(progress_row, textvariable=self.progress_text).pack(side="left", padx=(0, 12))
        self.progress = ttk.Progressbar(progress_row, mode="determinate", maximum=1)
        self.progress.pack(side="left", fill="x", expand=True)

        rules = ttk.Label(
            outer,
            textvariable=self.rules_text,
            style="Rule.TLabel",
            justify="left",
        )
        rules.pack(anchor="w", pady=(10, 8))
        self._update_rules()
        ttk.Label(outer, textvariable=self.summary_text, justify="left", style="Page.TLabel").pack(anchor="w", pady=(0, 8))

        footer = ttk.Frame(outer, style="Page.TFrame")
        footer.pack(side="bottom", fill="x", pady=(8, 0))
        status_label = ttk.Label(footer, textvariable=self.status_text, wraplength=1040, style="Page.TLabel")
        status_label.pack(anchor="w", fill="x")
        ttk.Label(footer, text="原文件自动备份后直接标注；新增油站切页，未配对发票在原页和油站页标红。", style="Footer.TLabel").pack(anchor="w", pady=(4, 0))
        footer.bind("<Configure>", lambda event: status_label.configure(wraplength=max(1, event.width)))

        pages = ttk.Notebook(outer)
        pages.pack(fill="both", expand=True)
        result_frame = ttk.Frame(pages, padding=10)
        success_frame = ttk.Frame(pages, padding=10)
        pages.add(result_frame, text="匹配失败")
        pages.add(success_frame, text="匹配成功")
        filters = ttk.Frame(result_frame)
        filters.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Checkbutton(filters, text="仅显示加油异常", variable=self.red_only, command=self._refresh_preview).pack(side="left")
        ttk.Label(filters, text="核对状态").pack(side="left", padx=(20, 8))
        self.status_combo = ttk.Combobox(filters, textvariable=self.status_filter, values=(STATUSES[0],), state="readonly", width=22)
        self.status_combo.pack(side="left")
        self.status_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())
        ttk.Label(filters, text="查找原始行号").pack(side="left", padx=(16, 8))
        ttk.Entry(filters, textvariable=self.row_search, width=10).pack(side="left")
        self.row_search.trace_add("write", self._schedule_preview_refresh)
        ttk.Label(filters, textvariable=self.preview_text).pack(side="right")

        columns = ("source", "row", "counterpart", "fuel_start_time", "fuel_end_time", "invoice_time", "expected", "actual", "amount_diff", "volume_diff", "time_diff", "status", "payment_type", "failure_reason", "basis")
        headings = ("来源", "行号", "对应记录位置 ↗", "加油明细切页-交易时间", "发票-交易时间", "开票时间", "应开票金额", "实际发票金额", "金额差（元）", "升数差", "时间差（分）", "核对状态", "付款类型", "失败原因", "配对依据")
        widths = (105, 55, 150, 215, 170, 170, 100, 100, 100, 85, 100, 160, 100, 260, 360)
        self.result_table = ttk.Treeview(result_frame, columns=columns, show="headings", selectmode="browse", height=7)
        for column, heading, width in zip(columns, headings, widths):
            self.result_table.heading(column, text=heading)
            self.result_table.column(column, width=width, minwidth=45, anchor="w" if column in {"source", "counterpart", "status", "failure_reason", "basis"} else "center", stretch=False)
        self.result_table.tag_configure("red", background="#FFC7CE")
        self.result_table.tag_configure("yellow", background="#FFF2CC")
        self.result_table.grid(row=1, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_table.yview)
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_table.xview)
        horizontal.grid(row=2, column=0, sticky="ew")
        self.result_table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.result_table.bind("<Double-1>", self._show_row_details)
        self.result_table.bind("<ButtonRelease-1>", self._open_counterpart)
        self.result_table.bind("<Motion>", self._update_link_cursor)
        result_frame.rowconfigure(1, weight=1)
        result_frame.columnconfigure(0, weight=1)

        ttk.Label(success_frame, textvariable=self.success_text).grid(row=0, column=0, columnspan=2, sticky="e", pady=(0, 8))
        success_columns = ("source", "row", "counterpart", "station", "gun", "oil", "fuel_volume", "invoice_volume", "fuel_start_time", "fuel_end_time", "invoice_time", "expected", "actual", "invoice_number", "card")
        success_headings = ("加油工作表", "行号", "发票位置 ↗", "油站", "油枪号", "油品", "加油量", "发票数量", "加油明细切页-交易时间", "发票-交易时间", "开票时间", "交易金额", "发票金额", "发票号码", "卡号")
        success_widths = (110, 55, 150, 160, 75, 90, 95, 95, 215, 170, 170, 100, 100, 190, 145)
        self.success_table = ttk.Treeview(success_frame, columns=success_columns, show="headings", selectmode="browse", height=7, style="Success.Treeview")
        for column, heading, width in zip(success_columns, success_headings, success_widths):
            self.success_table.heading(column, text=heading)
            self.success_table.column(column, width=width, minwidth=45, anchor="center" if column in {"row", "gun", "fuel_volume", "invoice_volume", "expected", "actual"} else "w", stretch=False)
        self.success_table.grid(row=1, column=0, sticky="nsew")
        success_vertical = ttk.Scrollbar(success_frame, orient="vertical", command=self.success_table.yview)
        success_vertical.grid(row=1, column=1, sticky="ns")
        success_horizontal = ttk.Scrollbar(success_frame, orient="horizontal", command=self.success_table.xview)
        success_horizontal.grid(row=2, column=0, sticky="ew")
        self.success_table.configure(yscrollcommand=success_vertical.set, xscrollcommand=success_horizontal.set)
        self.success_table.bind("<Double-1>", self._show_row_details)
        self.success_table.bind("<ButtonRelease-1>", self._open_counterpart)
        self.success_table.bind("<Motion>", self._update_link_cursor)
        success_frame.rowconfigure(1, weight=1)
        success_frame.columnconfigure(0, weight=1)

    def _choose_datetime(self, variable: tk.StringVar, label: str, anchor: ttk.Button) -> None:
        if not self.fuel_date_range:
            messagebox.showerror("尚未读取日期", "请先选择数据工作簿，等待加油明细默认日期识别完成。", parent=self.root)
            return
        DateTimePicker(self.root, variable, label, anchor)

    def _clear_period(self) -> None:
        self.start_text.set(self.fuel_date_range[0] if self.fuel_date_range else "")
        self.end_text.set(self.fuel_date_range[1] if self.fuel_date_range else "")

    def _set_date_controls_enabled(self, enabled: bool) -> None:
        for button in (*self.time_buttons, self.all_time_button):
            button.configure(state="normal" if enabled else "disabled")

    def _choose_input(self) -> None:
        current = Path(self.input_path.get()).expanduser()
        selected = filedialog.askopenfilename(parent=self.root, title="选择加油与发票工作簿", initialdir=str(current.parent if current.is_file() else self.base_directory), filetypes=(("Excel 工作簿", "*.xlsx"),))
        if selected:
            self.input_path.set(selected)
            save_last_input_path(selected)

    def _choose_invoice_audit(self) -> None:
        current = Path(self.invoice_audit_path.get()).expanduser()
        source = Path(self.input_path.get()).expanduser()
        initial = current.parent if current.is_file() else source.parent if source.is_file() else self.base_directory
        selected = filedialog.askopenfilename(parent=self.root, title="选择发票审核单原文件", initialdir=str(initial), filetypes=(("Excel 工作簿", "*.xlsx"),))
        if selected:
            if Path(selected).name.startswith("~$"):
                messagebox.showerror("设置有误", "请选择发票审核单原文件，不要选择以 ~$ 开头的 Excel 临时锁定文件。", parent=self.root)
                return
            self.invoice_audit_path.set(selected)
            save_last_invoice_audit_path(selected)

    def _schedule_station_scan(self, *_args) -> None:
        self.station_scan_id += 1
        scan_id = self.station_scan_id
        self.station_path = ""
        self.station_scan_error = ""
        self.invoice_interval = None
        self.fuel_date_range = None
        self.start_text.set("")
        self.end_text.set("")
        self._set_date_controls_enabled(False)
        for child in self.station_items.winfo_children():
            child.destroy()
        self.station_vars.clear()
        path = self.input_path.get().strip()
        self.station_text.set("正在识别站点…" if path else "选择工作簿后自动识别发票备注中的加油站点")
        if path:
            self.root.after(500, lambda: self._begin_station_scan(scan_id, path))

    def _begin_station_scan(self, scan_id: int, path: str) -> None:
        if scan_id != self.station_scan_id or not Path(path).is_file():
            return
        threading.Thread(target=self._scan_stations_worker, args=(scan_id, path), daemon=True).start()

    def _scan_stations_worker(self, scan_id: int, path: str) -> None:
        try:
            from reconciliation import scan_workbook_metadata

            names, interval, fuel_range = scan_workbook_metadata(path)
            if fuel_range is None:
                raise ValueError("加油明细切页未找到有效交易时间，无法确定默认日期范围")
            self.events.put(("stations", (scan_id, path, names, interval, fuel_range, "")))
        except Exception as exc:
            self.events.put(("stations", (scan_id, path, [], None, None, str(exc))))

    def _set_all_stations(self, selected: bool) -> None:
        for variable in self.station_vars.values():
            variable.set(selected)

    def _schedule_preview_refresh(self, *_args) -> None:
        if self.search_refresh_id is not None:
            self.root.after_cancel(self.search_refresh_id)
        self.search_refresh_id = self.root.after(250, self._apply_search)

    def _apply_search(self) -> None:
        self.search_refresh_id = None
        self._refresh_preview()

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(parent=self.root, title="选择结果输出目录", initialdir=self.output_directory.get() or str(self.base_directory))
        if selected:
            self.output_directory.set(selected)

    def _update_rules(self, *_args) -> None:
        self.rules_text.set(
            "核对规则：交易金额已扣优惠，直接与发票金额比较；加油量与发票数量一致；不限制结束时间晚于开始时间的间隔，早于开始时间不超过 10 分钟也视为正常。\n"
            "发票备注交易时间为结束，结合油枪号和油品配对；缺失时用开票时间辅助。\n"
            "Excel 付款类型为记账卡的加油行标黄，其他加油异常行标红并写入失败原因；同名 .log 只记录失败。"
        )

    def _start_run(self) -> None:
        if self.busy:
            return
        # Read every Tk variable on the UI thread before starting background work.
        parameters = {
            "input_path": self.input_path.get().strip(),
            "invoice_audit_path": self.invoice_audit_path.get().strip(),
            "output_dir": self.output_directory.get().strip(),
            "date_text": "",
            "start_text": self.start_text.get().strip(),
            "end_text": self.end_text.get().strip(),
            "red_only": False,
            "worker_count": self.worker_count.get(),
        }
        try:
            source = Path(parameters["input_path"])
            if not parameters["input_path"] or not source.is_file():
                raise ValueError("请选择存在的数据工作簿。")
            if source.suffix.lower() != ".xlsx":
                raise ValueError("请选择 .xlsx 格式的 Excel 工作簿。")
            audit = Path(parameters["invoice_audit_path"])
            if not parameters["invoice_audit_path"] or not audit.is_file() or audit.suffix.lower() != ".xlsx":
                raise ValueError("请选择存在的 .xlsx 发票审核单。")
            if audit.name.startswith("~$"):
                raise ValueError("请选择发票审核单原文件，不要选择以 ~$ 开头的 Excel 临时锁定文件。")
            if audit.resolve() == source.resolve():
                raise ValueError("发票审核单与数据工作簿不能是同一个文件。")
            if not parameters["output_dir"]:
                raise ValueError("请填写输出目录。")
            if self.station_path != parameters["input_path"]:
                if self.station_scan_error:
                    raise ValueError(f"站点识别失败：{self.station_scan_error}")
                messagebox.showinfo("请稍候", "正在读取工作簿并识别加油站点，完成后即可开始核对。", parent=self.root)
                return
            if self.station_vars and not any(variable.get() for variable in self.station_vars.values()):
                raise ValueError("请至少勾选一个审核站点。")
            parameters["selected_stations"] = [name for name, variable in self.station_vars.items() if variable.get()] if self.station_vars else None
            parameters["worker_count"] = int(parameters["worker_count"])
            if not 1 <= parameters["worker_count"] <= 32:
                raise ValueError("并发工作数须为 1—32 的整数。")
            from reconciliation import parse_period
            parse_period(date_text="", start_text=parameters["start_text"], end_text=parameters["end_text"])
        except ValueError as exc:
            messagebox.showerror("设置有误", str(exc), parent=self.root)
            return
        save_last_input_path(parameters["input_path"])
        save_last_invoice_audit_path(parameters["invoice_audit_path"])
        self.busy = True
        self.output_path = None
        self.annotated_input_path = None
        self.input_backup_path = None
        self.log_path = None
        self.open_button.configure(state="disabled")
        self.log_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.folder_button.configure(state="disabled")
        for control in self.setting_controls:
            control.configure(state="disabled")
        self.results = []
        self._update_status_options()
        self._refresh_preview()
        self.summary_text.set("正在读取并核对两侧记录…")
        self.status_text.set("正在核对，请稍候。")
        self.progress.configure(mode="indeterminate", maximum=100, value=0)
        self.progress.start(15)
        self.progress_text.set("核对进度：正在读取工作簿…")
        threading.Thread(target=self._run_worker, args=(parameters,), name="reconciliation-worker", daemon=True).start()

    def _run_worker(self, parameters: dict) -> None:
        try:
            from reconciliation import run_reconciliation

            result = run_reconciliation(
                **parameters,
                progress=lambda message: self.events.put(("progress", str(message))),
                match_progress=lambda stage, done, total: self.events.put(("match_progress", (stage, done, total))),
                workflow_progress=lambda message, percent: self.events.put(("workflow_progress", (message, percent))),
            )
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc) or type(exc).__name__))

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    self.status_text.set(payload)
                elif event == "stations":
                    scan_id, path, names, interval, fuel_range, error = payload
                    if scan_id == self.station_scan_id and path == self.input_path.get().strip():
                        self.station_path = path if not error else ""
                        self.station_scan_error = error
                        self.invoice_interval = interval if not error else None
                        self.fuel_date_range = fuel_range if not error else None
                        if fuel_range:
                            self.start_text.set(fuel_range[0])
                            self.end_text.set(fuel_range[1])
                        self._set_date_controls_enabled(bool(fuel_range) and not self.busy)
                        date_text = f"，加油日期 {fuel_range[0]} 至 {fuel_range[1]}" if fuel_range else ""
                        interval_text = f"，发票资料区间 {interval[0]} 至 {interval[1]}" if interval else ""
                        self.station_text.set(f"已识别 {len(names)} 个站点{date_text}{interval_text}，请勾选审核范围" if names else (f"识别失败：{error}" if error else "未从发票备注中识别站点，默认核对全部"))
                        for index, name in enumerate(names):
                            variable = tk.BooleanVar(value=True)
                            self.station_vars[name] = variable
                            ttk.Checkbutton(self.station_items, text=name, variable=variable).grid(row=index // 4, column=index % 4, sticky="w", padx=(0, 16))
                elif event == "match_progress":
                    stage, done, total = payload
                    self.progress.stop()
                    self.progress.configure(mode="determinate", maximum=100,
                                            value=25 * (stage - 1 + done / max(1, total)))
                    self.progress_text.set(f"核对进度：第 {stage}/3 轮，{done}/{total}")
                elif event == "workflow_progress":
                    message, percent = payload
                    self.progress.configure(mode="determinate", maximum=100, value=percent)
                    self.progress_text.set(f"核对进度：{message}")
                elif event == "done":
                    self.progress.stop()
                    self._finish_run()
                    self.output_path = Path(payload["output_path"])
                    self.annotated_input_path = Path(payload["annotated_input_path"])
                    self.input_backup_path = Path(payload["input_backup_path"])
                    self.log_path = Path(payload["log_path"])
                    self.results = payload.get("results", [])
                    self.progress.configure(mode="determinate", maximum=100, value=100)
                    self.progress_text.set(f"核对进度：{payload['total_fuel']}/{payload['total_fuel']}（完成）")
                    self._update_status_options()
                    self.open_button.configure(state="normal")
                    self.log_button.configure(state="normal")
                    self.summary_text.set(
                        f"加油记录 {payload['total_fuel']} 条　发票记录 {payload['total_invoice']} 条　对应 {payload['paired']} 对　"
                        f"已匹配 {payload['matched']} 对，其余需核查　已过滤无效发票 {payload.get('filtered_invoice_count', 0)} 条　核对耗时：{payload['elapsed_text']}\n"
                        f"原文件已标注：加油异常 {payload['red_fuel']} 条，未配对发票 {payload['red_invoice']} 条；记账卡仍标黄。"
                    )
                    status = f"已标注原文件：{self.annotated_input_path}\n核对前备份：{self.input_backup_path}\n结果工作簿：{self.output_path}\n核对日志：{self.log_path}"
                    if payload.get("log_cleanup_warning"):
                        status += f"\n{payload['log_cleanup_warning']}"
                    self.status_text.set(status)
                    self._refresh_preview()
                elif event == "error":
                    self.progress.stop()
                    self.progress.configure(mode="determinate", maximum=100, value=0)
                    self._finish_run()
                    self.progress_text.set("核对进度：已中止")
                    self.summary_text.set("本次核对未完成，请检查提示后重试。")
                    self.status_text.set(f"核对失败：{payload}")
                    messagebox.showerror("核对未完成", payload, parent=self.root)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_events)

    def _finish_run(self) -> None:
        self.busy = False
        self.run_button.configure(state="normal")
        self.folder_button.configure(state="normal")
        for control in self.setting_controls:
            control.configure(state="readonly" if control is self.worker_combo else "normal")
        self._set_date_controls_enabled(self.fuel_date_range is not None)

    def _update_status_options(self) -> None:
        present = {result["status"] for result in self.results if result.get("kind") == "fuel" and result.get("status") and result["status"] != "已匹配"}
        available = (STATUSES[0], *(status for status in STATUSES[1:] if status in present), *(sorted(present - set(STATUSES))))
        self.status_combo.configure(values=available)
        if self.status_filter.get() not in available:
            self.status_filter.set(STATUSES[0])

    def _refresh_preview(self) -> None:
        selected_status = self.status_filter.get()
        red_only = self.red_only.get()
        row_search = self.row_search.get().strip()
        for table in (self.result_table, self.success_table):
            children = table.get_children()
            if children:
                table.delete(*children)
        visible = 0
        successful = 0
        fuel_rows = [(index, result) for index, result in enumerate(self.results) if result.get("kind") == "fuel"]
        time_key = lambda item: (not item[1].get("fuel_start_time"), item[1].get("fuel_start_time") or "", int(item[1].get("row") or 0))
        success_rows = sorted((item for item in fuel_rows if item[1].get("status") == "已匹配"),
                              key=time_key)
        failure_rows = sorted((item for item in fuel_rows if item[1].get("status") != "已匹配"),
                              key=time_key)
        for index, result in success_rows:
            values = tuple("" if result.get(key) is None else str(result.get(key, "")).replace("\n", "；") for key in self.success_table["columns"])
            self.success_table.insert("", "end", iid=str(index), values=values)
            successful += 1
        for index, result in failure_rows:
            if red_only and not result.get("red"):
                continue
            if selected_status != STATUSES[0] and result.get("status") != selected_status:
                continue
            if row_search and str(result.get("row", "")) != row_search:
                continue
            # Full multiline content stays in self.results for the detail window.
            values = tuple("" if result.get(key) is None else str(result.get(key, "")).replace("\n", "；") for key in self.result_table["columns"])
            self.result_table.insert("", "end", iid=str(index), values=values, tags=("yellow" if result.get("payment_type") == "记账卡" else "red",))
            visible += 1
        total = len(fuel_rows)
        self.preview_text.set(f"异常 {total - successful} / {total} 笔（当前显示 {visible} 行）")
        self.success_text.set(f"匹配成功 {successful} / {total} 笔")

    def _is_counterpart_cell(self, event) -> bool:
        table = getattr(event, "widget", self.result_table)
        item = table.identify_row(event.y)
        return bool(item and table.identify_column(event.x) == "#3" and self.results[int(item)].get("counterpart"))

    def _update_link_cursor(self, event) -> None:
        table = getattr(event, "widget", self.result_table)
        table.configure(cursor="hand2" if self._is_counterpart_cell(event) else "")

    def _open_counterpart(self, event) -> None:
        if self._is_counterpart_cell(event):
            self._open_result()

    def _show_row_details(self, event=None) -> None:
        if event is not None and self._is_counterpart_cell(event):
            return
        table = getattr(event, "widget", self.result_table) if event is not None else self.result_table
        selected = table.selection()
        if not selected:
            return
        result = self.results[int(selected[0])]
        lines = [f"{table.heading(key, 'text')}：\n{'' if result.get(key) is None else result.get(key, '')}" for key in table["columns"]]
        window = tk.Toplevel(self.root)
        window.title("记录核对详情")
        set_window_icon(window)
        window.geometry("850x680")
        window.minsize(620, 440)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=16)
        frame.pack(fill="both", expand=True)
        text_widget = tk.Text(frame, wrap="word", padx=12, pady=12, relief="flat")
        text_widget.insert("1.0", "\n\n".join(lines))
        text_widget.configure(state="disabled")
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=scrollbar.set)
        ttk.Button(frame, text="关闭", command=window.destroy).grid(row=1, column=0, columnspan=2, pady=(12, 0))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _open_path(self, path: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法打开", f"请手动打开：\n{path}\n\n{exc}", parent=self.root)

    def _open_result(self) -> None:
        if self.annotated_input_path is not None:
            if self.annotated_input_path.is_file():
                self._open_path(self.annotated_input_path)
            else:
                messagebox.showerror("原文件不存在", f"已标注原文件可能已被移动或删除：\n{self.annotated_input_path}", parent=self.root)

    def _open_log(self) -> None:
        if self.log_path is not None:
            if self.log_path.is_file():
                self._open_path(self.log_path)
            else:
                messagebox.showerror("核对日志不存在", f"日志文件可能已被移动或删除：\n{self.log_path}", parent=self.root)

    def _open_output(self) -> None:
        path = self.output_path.parent if self.output_path else Path(self.output_directory.get().strip() or self.base_directory / "outputs")
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("无法打开输出目录", str(exc), parent=self.root)
            return
        self._open_path(path)

    def _show_help(self) -> None:
        if self.help_window is not None and self.help_window.winfo_exists():
            self.help_window.lift()
            self.help_window.focus_set()
            return
        window = tk.Toplevel(self.root)
        self.help_window = window
        window.title("使用帮助")
        set_window_icon(window)
        window.geometry("760x610")
        window.minsize(600, 440)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=16)
        frame.pack(fill="both", expand=True)
        text_widget = tk.Text(frame, wrap="word", padx=12, pady=12, relief="flat")
        text_widget.insert("1.0", HELP_TEXT)
        text_widget.configure(state="disabled")
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=scrollbar.set)
        ttk.Button(frame, text="关闭", command=window.destroy).grid(row=1, column=0, columnspan=2, pady=(12, 0))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _close(self) -> None:
        if self.busy and not messagebox.askyesno("正在核对", "核对仍在进行。现在关闭将中止本次处理，是否关闭？", parent=self.root):
            return
        self.root.destroy()


def main(*, smoke: bool = False) -> int:
    set_application_id()
    root = tk.Tk()
    ReconciliationApp(root)
    if smoke:
        root.after(1000, root.destroy)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
