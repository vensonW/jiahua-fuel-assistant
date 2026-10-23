"""加华加油与发票核对助手的桌面界面。"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "加华加油与发票核对助手"
STATUSES = (
    "全部状态",
    "已匹配",
    "金额不符",
    "升数不符",
    "时间超限",
    "无对应记录",
    "待确认",
)
HELP_TEXT = """加华加油与发票核对助手使用说明

1. 选择包含加油数据和发票数据的 Excel 工作簿。工具识别前两页，不需要手动复制数据。

2. 开始时间、结束时间默认均为 2026-10-11。可输入 YYYY-MM-DD 或 YYYY-MM-DD HH:MM[:SS]；仅输入日期时，开始取当天 00:00:00，结束包含当天全部时间。两个输入都留空可核对全部，只填写一侧时请补齐另一侧。开始时间不得晚于结束时间。

3. 加油表“交易时间”是加油开始时间；发票备注“交易时间”是加油结束时间；发票“开票时间”是发票开具时间。优先根据备注油枪号和结束时间，结合开始时间、油品、升数及油站寻找对应记录。正常时间差＝结束时间－开始时间，必须在 0—30 分钟之间，包含恰好 30 分钟；时间倒置标为待确认。

4. 应开票金额＝交易金额－优惠金额，与发票“合计金额（含税）”按原始数值严格比较，不先四舍五入。优惠真正为空时按 0 计算并注明；字段无效或公式未取得数值时标为待确认。加油升数与发票升数也必须完全一致。

5. 同枪号、同油品、同升数的重复候选按双方时间顺序一一配对，注明“按时间顺序推定”，不能为使金额相等而调换顺序。每笔记录只使用一次；重复票号、同票多行或其他证据不足的情况标为待确认。

6. 金额只要有差异，即使相差 0.01 元，也判为“金额不符”并整行标浅红底。多项问题全部写入配对依据，存在金额差时主状态优先显示金额不符。其他不符、无对应记录、待确认也标浅红底；保留疑似对应位置及差异供核查。

7. 备注缺少结束时间时，允许使用开票时间辅助判断并注明来源，加油结束时间保持空白；备注时间存在但格式无效时标为待确认。发票页最右侧“加油完成时间”继续严格取备注时间，缺失时留空。

8. 结果另存到输出目录，重名时自动追加序号。前两页原表右侧追加结果，不新增异常切页，保留原始字段、其他工作表、合计行、完整卡号和发票号码。时间段之外的记录保留并注明未参与核对。

9. 前两页各有“匹配明细”列，仅“已匹配”记录填写，双方显示相同的枪号、卡号、油品、完整票号、三个时间、双方升数、各项金额及原始位置。预览可横向滚动，双击记录可查看完整多行内容；异常记录的匹配明细留空。

10. “仅显示红底”和状态下拉框可组合筛选预览，不会减少导出记录。结果工作簿支持按核对状态或浅红底颜色筛选。“金额差”＝实际发票金额－应开票金额；“升数差”＝发票升数－加油升数；时间差保留正负号。统计中的“对应”和“已匹配”按一对加油与发票计数，双方红底行数分别统计。
"""


def application_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


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


class ReconciliationApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        set_window_icon(root, default=True)
        root.geometry("1120x800")
        root.minsize(1000, 720)
        self.base_directory = application_directory()
        sample = self.base_directory / "中南加华加油和开票数据核对.xlsx"
        self.input_path = tk.StringVar(value=str(sample) if sample.exists() else "")
        self.output_directory = tk.StringVar(value=str(self.base_directory / "outputs"))
        self.start_text = tk.StringVar(value="2026-10-11")
        self.end_text = tk.StringVar(value="2026-10-11")
        self.red_only = tk.BooleanVar(value=False)
        self.status_filter = tk.StringVar(value=STATUSES[0])
        self.status_text = tk.StringVar(value="请选择工作簿，核实设置后开始核对。")
        self.summary_text = tk.StringVar(value="尚未核对。结果将分别统计加油记录、发票记录和配对数。")
        self.preview_text = tk.StringVar(value="结果预览：0 行")
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.results: list[dict] = []
        self.output_path: Path | None = None
        self.help_window: tk.Toplevel | None = None
        self.brand_logo: tk.PhotoImage | None = None
        self.setting_controls: list[ttk.Widget] = []
        self._build()
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
        style.configure("Header.TFrame", background="#DDE8F6")
        style.configure("Header.TLabel", background="#DDE8F6", foreground=blue)
        style.configure("Card.TLabelframe", background="#FFFFFF", bordercolor="#D7E3F0", borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background="#FFFFFF", foreground=blue, font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=text, padding=6, bordercolor="#CAD9E9", lightcolor="#FFFFFF", darkcolor="#FFFFFF")
        style.map("TEntry", bordercolor=[("focus", "#7FA8D0")], foreground=[("disabled", "#8A98A8")])
        style.configure("TCombobox", fieldbackground="#FFFFFF", background="#EEF4FC", foreground=text, padding=6, arrowcolor=blue, bordercolor="#CAD9E9")
        style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], selectbackground=[("readonly", "#FFFFFF")], selectforeground=[("readonly", text)])
        style.configure("TButton", background="#E8F1FB", foreground=blue, padding=(14, 7), bordercolor="#C7D9EC", lightcolor="#E8F1FB", darkcolor="#E8F1FB")
        style.map("TButton", background=[("disabled", "#EEF2F6"), ("pressed", "#C5DAF1"), ("active", "#D9E8F8")], foreground=[("disabled", "#8A98A8")])
        style.configure("Primary.TButton", background="#215FDB", foreground="#FFFFFF", font=("Microsoft YaHei UI", 10, "bold"), bordercolor="#215FDB", lightcolor="#215FDB", darkcolor="#215FDB")
        style.map("Primary.TButton", background=[("disabled", "#B7C7D8"), ("pressed", "#1243AC"), ("active", "#174FC4")], foreground=[("disabled", "#F2F5FA"), ("!disabled", "#FFFFFF")])
        style.configure("TCheckbutton", background="#FFFFFF", foreground=text)
        style.map("TCheckbutton", background=[("active", "#FFFFFF")])
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=text, rowheight=30, borderwidth=0)
        style.map("Treeview", background=[("selected", "#DDEBFA")], foreground=[("selected", blue)])
        style.configure("Treeview.Heading", background="#EEF4FC", foreground=blue, font=("Microsoft YaHei UI", 10, "bold"), padding=(8, 8), relief="flat")
        style.map("Treeview.Heading", background=[("active", "#DFEBFA")])
        for scrollbar in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            style.configure(scrollbar, background="#D7E4F2", troughcolor="#F1F5FA", bordercolor="#F1F5FA", arrowcolor=blue)
        style.configure("Horizontal.TProgressbar", background="#4E86BA", troughcolor="#EEF4FC", bordercolor="#D7E3F0")
        outer = ttk.Frame(self.root, padding=20, style="Page.TFrame")
        outer.pack(fill="both", expand=True)
        header = ttk.Frame(outer, padding=(16, 10), style="Header.TFrame")
        header.pack(fill="x", pady=(0, 16))
        logo_path = application_asset_path("gawa.png")
        if logo_path.is_file():
            try:
                logo = tk.PhotoImage(file=str(logo_path))
                factor = max(1, (max(logo.width(), logo.height()) + 47) // 48)
                self.brand_logo = logo.subsample(factor, factor) if factor > 1 else logo
                ttk.Label(header, image=self.brand_logo, style="Header.TLabel").pack(side="left", padx=(0, 12))
            except tk.TclError:
                pass
        ttk.Label(header, text="加油与发票数据核对", font=("Microsoft YaHei UI", 18, "bold"), style="Header.TLabel").pack(side="left")
        ttk.Button(header, text="使用帮助", command=self._show_help).pack(side="right")

        form = ttk.LabelFrame(outer, text="核对设置", padding=(14, 10), style="Card.TLabelframe")
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        for row, label, variable, chooser in (
            (0, "数据工作簿", self.input_path, self._choose_input),
            (1, "输出目录", self.output_directory, self._choose_output),
        ):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            entry = ttk.Entry(form, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            button = ttk.Button(form, text="浏览…", command=chooser)
            button.grid(row=row, column=2, padx=(8, 0), pady=5)
            self.setting_controls.extend((entry, button))
        ttk.Label(form, text="核对时间段").grid(row=2, column=0, sticky="nw", padx=(0, 12), pady=5)
        options = ttk.Frame(form)
        options.grid(row=2, column=1, columnspan=2, sticky="ew", pady=5)
        for column, label, variable in ((0, "开始时间", self.start_text), (2, "结束时间", self.end_text)):
            ttk.Label(options, text=label).grid(row=0, column=column, sticky="w", padx=(0 if column == 0 else 20, 8))
            entry = ttk.Entry(options, textvariable=variable, width=23)
            entry.grid(row=0, column=column + 1, sticky="ew")
            self.setting_controls.append(entry)
        options.columnconfigure(1, weight=1)
        options.columnconfigure(3, weight=1)
        ttk.Label(options, text="格式：YYYY-MM-DD 或 YYYY-MM-DD HH:MM[:SS]；仅填日期包含整天，两侧都留空核对全部。", foreground="#63778A").grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        actions = ttk.Frame(form)
        actions.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        self.run_button = ttk.Button(actions, text="开始核对", command=self._start_run, style="Primary.TButton")
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(actions, text="打开结果工作簿", command=self._open_result, state="disabled")
        self.open_button.pack(side="left", padx=(8, 0))
        self.folder_button = ttk.Button(actions, text="打开输出目录", command=self._open_output)
        self.folder_button.pack(side="left", padx=(8, 0))
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=140)
        self.progress.pack(side="right", padx=(12, 0))

        rules = ttk.Label(
            outer,
            text="核对规则：应开票＝交易金额－优惠；金额零差额、升数一致；结束－开始在 0—30 分钟内（含边界）。\n"
                 "发票备注交易时间为结束，结合油枪号和油品配对；缺失时用开票时间辅助，金额有差异均判金额不符并标红。",
            style="Rule.TLabel",
            justify="left",
        )
        rules.pack(anchor="w", pady=(10, 8))
        ttk.Label(outer, textvariable=self.summary_text, justify="left", style="Page.TLabel").pack(anchor="w", pady=(0, 8))

        footer = ttk.Frame(outer, style="Page.TFrame")
        footer.pack(side="bottom", fill="x", pady=(8, 0))
        status_label = ttk.Label(footer, textvariable=self.status_text, wraplength=1040, style="Page.TLabel")
        status_label.pack(anchor="w", fill="x")
        ttk.Label(footer, text="原文件保留；结果另存。预览筛选不影响导出记录。双击预览行可查看完整依据。", style="Footer.TLabel").pack(anchor="w", pady=(4, 0))
        footer.bind("<Configure>", lambda event: status_label.configure(wraplength=max(1, event.width)))

        result_frame = ttk.LabelFrame(outer, text="原表两侧记录预览", padding=10, style="Card.TLabelframe")
        result_frame.pack(fill="both", expand=True)
        filters = ttk.Frame(result_frame)
        filters.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Checkbutton(filters, text="仅显示红底", variable=self.red_only, command=self._refresh_preview).pack(side="left")
        ttk.Label(filters, text="核对状态").pack(side="left", padx=(20, 8))
        combo = ttk.Combobox(filters, textvariable=self.status_filter, values=STATUSES, state="readonly", width=22)
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())
        ttk.Label(filters, textvariable=self.preview_text).pack(side="right")

        columns = ("source", "row", "counterpart", "fuel_start_time", "fuel_end_time", "invoice_time", "expected", "actual", "amount_diff", "volume_diff", "time_diff", "status", "basis", "match_details")
        headings = ("来源", "行号", "对应记录位置", "加油开始时间", "加油结束时间", "开票时间", "应开票金额", "实际发票金额", "金额差（元）", "升数差", "时间差（分）", "核对状态", "配对依据", "匹配明细")
        widths = (105, 55, 140, 170, 170, 170, 100, 100, 100, 85, 100, 160, 360, 420)
        self.result_table = ttk.Treeview(result_frame, columns=columns, show="headings", selectmode="browse", height=7)
        for column, heading, width in zip(columns, headings, widths):
            self.result_table.heading(column, text=heading)
            self.result_table.column(column, width=width, minwidth=45, anchor="w" if column in {"source", "counterpart", "status", "basis", "match_details"} else "center", stretch=False)
        self.result_table.tag_configure("red", background="#FFC7CE")
        self.result_table.grid(row=1, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_table.yview)
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_table.xview)
        horizontal.grid(row=2, column=0, sticky="ew")
        self.result_table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.result_table.bind("<Double-1>", self._show_row_details)
        result_frame.rowconfigure(1, weight=1)
        result_frame.columnconfigure(0, weight=1)

    def _choose_input(self) -> None:
        current = Path(self.input_path.get()).expanduser()
        selected = filedialog.askopenfilename(parent=self.root, title="选择加油与发票工作簿", initialdir=str(current.parent if current.is_file() else self.base_directory), filetypes=(("Excel 工作簿", "*.xlsx"),))
        if selected:
            self.input_path.set(selected)

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(parent=self.root, title="选择结果输出目录", initialdir=self.output_directory.get() or str(self.base_directory))
        if selected:
            self.output_directory.set(selected)

    def _start_run(self) -> None:
        if self.busy:
            return
        # Read every Tk variable on the UI thread before starting background work.
        parameters = {
            "input_path": self.input_path.get().strip(),
            "output_dir": self.output_directory.get().strip(),
            "date_text": "",
            "start_text": self.start_text.get().strip(),
            "end_text": self.end_text.get().strip(),
            "minutes_tolerance": "30",
            "red_only": False,
        }
        try:
            source = Path(parameters["input_path"])
            if not parameters["input_path"] or not source.is_file():
                raise ValueError("请选择存在的数据工作簿。")
            if source.suffix.lower() != ".xlsx":
                raise ValueError("请选择 .xlsx 格式的 Excel 工作簿。")
            if not parameters["output_dir"]:
                raise ValueError("请填写输出目录。")
            from reconciliation import parse_period

            parse_period(date_text="", start_text=parameters["start_text"], end_text=parameters["end_text"])
        except ValueError as exc:
            messagebox.showerror("设置有误", str(exc), parent=self.root)
            return
        self.busy = True
        self.output_path = None
        self.open_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.folder_button.configure(state="disabled")
        for control in self.setting_controls:
            control.configure(state="disabled")
        self.results = []
        self._refresh_preview()
        self.summary_text.set("正在读取并核对两侧记录…")
        self.status_text.set("正在核对，请稍候。")
        self.progress.start(12)
        threading.Thread(target=self._run_worker, args=(parameters,), name="reconciliation-worker", daemon=True).start()

    def _run_worker(self, parameters: dict) -> None:
        try:
            from reconciliation import run_reconciliation

            result = run_reconciliation(**parameters, progress=lambda message: self.events.put(("progress", str(message))))
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc) or type(exc).__name__))

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    self.status_text.set(payload)
                elif event == "done":
                    self._finish_run()
                    self.output_path = Path(payload["output_path"])
                    self.results = payload.get("results", [])
                    self.open_button.configure(state="normal")
                    self.summary_text.set(
                        f"加油记录 {payload['total_fuel']} 条　发票记录 {payload['total_invoice']} 条　对应 {payload['paired']} 对　"
                        f"已匹配 {payload['matched']} 对，其余需核查\n"
                        f"红底：加油侧 {payload['red_fuel']} 条，发票侧 {payload['red_invoice']} 条"
                    )
                    self.status_text.set(f"核对完成，结果已保存：{self.output_path}")
                    self._refresh_preview()
                elif event == "error":
                    self._finish_run()
                    self.summary_text.set("本次核对未完成，请检查提示后重试。")
                    self.status_text.set(f"核对失败：{payload}")
                    messagebox.showerror("核对未完成", payload, parent=self.root)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_events)

    def _finish_run(self) -> None:
        self.busy = False
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.folder_button.configure(state="normal")
        for control in self.setting_controls:
            control.configure(state="normal")

    def _refresh_preview(self) -> None:
        selected_status = self.status_filter.get()
        red_only = self.red_only.get()
        children = self.result_table.get_children()
        if children:
            self.result_table.delete(*children)
        visible = 0
        for index, result in enumerate(self.results):
            if red_only and not result.get("red"):
                continue
            if selected_status != STATUSES[0] and result.get("status") != selected_status:
                continue
            # Full multiline content stays in self.results for the detail window.
            values = tuple("" if result.get(key) is None else str(result.get(key, "")).replace("\n", "；") for key in self.result_table["columns"])
            self.result_table.insert("", "end", iid=str(index), values=values, tags=("red",) if result.get("red") else ())
            visible += 1
        self.preview_text.set(f"显示 {visible} / {len(self.results)} 行")

    def _show_row_details(self, _event=None) -> None:
        selected = self.result_table.selection()
        if not selected:
            return
        result = self.results[int(selected[0])]
        lines = [f"{self.result_table.heading(key, 'text')}：\n{'' if result.get(key) is None else result.get(key, '')}" for key in self.result_table["columns"]]
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
        if self.output_path is not None:
            if self.output_path.is_file():
                self._open_path(self.output_path)
            else:
                messagebox.showerror("结果文件不存在", f"结果文件可能已被移动或删除：\n{self.output_path}", parent=self.root)

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
