"""GUI and command-line entry point for the packaged application."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def _write_message(text: str, *, error: bool = False) -> None:
    stream = sys.stderr if error else sys.stdout
    if stream is not None:
        try:
            stream.write(text + "\n")
            stream.flush()
            return
        except (OSError, UnicodeError):
            pass
    if error:
        import tkinter as tk
        from tkinter import messagebox
        from app import APP_TITLE, set_application_id, set_window_icon

        set_application_id()
        window = tk.Tk()
        set_window_icon(window, default=True)
        window.withdraw()
        messagebox.showerror(APP_TITLE, text, parent=window)
        window.destroy()


def _write_progress(text: str) -> None:
    if sys.stderr is not None:
        try:
            sys.stderr.write(text + "\n")
            sys.stderr.flush()
        except (OSError, UnicodeError):
            pass


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="加华油站小助手；无参数启动桌面界面。")
    parser.add_argument("--input", help="包含加油和发票记录的 xlsx 工作簿，按表头识别工作表")
    parser.add_argument("--invoice-audit", help="发票审核单 xlsx 原文件；过滤其中被红冲或无效的发票")
    parser.add_argument("--output-dir", help="输出目录，默认程序所在目录下的 outputs")
    parser.add_argument("--date", help="兼容单日 YYYY-MM-DD；传入空字符串核对全部；不与 --start/--end 混用")
    parser.add_argument("--start", help="开始日期或时间 YYYY-MM-DD [HH:MM[:SS]]；与 --end 配合使用")
    parser.add_argument("--end", help="结束日期或时间 YYYY-MM-DD [HH:MM[:SS]]；只填日期包含整天；两侧都传空字符串核对全部")
    parser.add_argument("--red-only", action="store_true", help="工作簿打开时默认只显示加油异常记录，仍保留完整数据")
    parser.add_argument("--workers", type=int, default=20, help="并发工作数，1—32，默认 20")
    parser.add_argument("--summary-json", help="将本次核对摘要保存为 UTF-8 JSON 文件")
    parser.add_argument("--smoke-gui", action="store_true", help=argparse.SUPPRESS)
    command_arguments = sys.argv[1:] if argv is None else argv
    args = parser.parse_args(command_arguments)
    period_requested = args.start is not None or args.end is not None
    if args.date is not None and period_requested:
        parser.error("--date 不能与 --start/--end 同时使用。")

    if args.smoke_gui or not args.input:
        if not args.smoke_gui and command_arguments:
            parser.error("命令行核对需要 --input 工作簿路径。")
        from app import main as gui_main

        return gui_main(smoke=args.smoke_gui)

    try:
        from app import application_directory
        from reconciliation import run_reconciliation

        result = run_reconciliation(
            input_path=args.input,
            invoice_audit_path=args.invoice_audit,
            output_dir=args.output_dir or application_directory() / "outputs",
            date_text=args.date if args.date is not None else ("" if period_requested else "2026-10-11"),
            start_text=args.start,
            end_text=args.end,
            red_only=args.red_only,
            worker_count=args.workers,
            progress=lambda message: _write_progress(str(message)),
        )
        summary = {key: value for key, value in result.items() if key != "results"}
        text = json.dumps(summary, ensure_ascii=False, indent=2, default=str)
        if args.summary_json:
            destination = Path(args.summary_json).expanduser().resolve()
            destination.parent.mkdir(parents=True, exist_ok=True)
            with destination.open("w", encoding="utf-8", newline="\n") as output:
                output.write(text + "\n")
        _write_message(text)
        return 0
    except Exception as exc:
        _write_message(f"核对未完成：{str(exc) or type(exc).__name__}", error=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
