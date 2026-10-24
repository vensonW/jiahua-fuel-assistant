"""Exercise our own GUI event loop, background worker and preview controls."""
from pathlib import Path
import sys
import time
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import ReconciliationApp, application_icon_path, set_application_id


project = Path(__file__).resolve().parent.parent
set_application_id()
root = tk.Tk()
app = ReconciliationApp(root)
source = project / "中南加华加油和开票数据核对.xlsx"
app.input_path.set(str(source))
app.invoice_audit_path.set(str(project / "invoiceList_2026-10-24.xlsx"))
scan_deadline = time.monotonic() + 30
while app.station_path != str(source) and time.monotonic() < scan_deadline:
    root.update()
    time.sleep(0.02)
assert app.station_path == str(source), app.station_text.get()
qa = project / "outputs" / "qa" / "original_marking"
qa.mkdir(parents=True, exist_ok=True)
app.output_directory.set(str(qa / "gui_run"))
app.start_text.set("2026-10-07 18:00:00")
app.end_text.set("2026-10-07 19:00:00")
app._start_run()
deadline = time.monotonic() + 120
while app.busy and time.monotonic() < deadline:
    root.update()
    time.sleep(0.02)
assert not app.busy and app.output_path and app.output_path.is_file(), app.status_text.get()
assert app.log_path and app.log_path.is_file()
assert app.log_path == app.output_path.parent / "logs" / (app.output_path.stem + ".log")
assert str(app.log_button["state"]) != "disabled"
failed = [result for result in app.results if result["kind"] == "fuel" and result["status"] != "已匹配"]
assert len(app.result_table.get_children()) == len(failed)
assert len(app.success_table.get_children()) == sum(result["kind"] == "fuel" and result["status"] == "已匹配" for result in app.results)
assert {"gun", "fuel_volume", "fuel_start_time", "fuel_end_time"} <= set(app.success_table["columns"])
assert {"fuel_start_time", "fuel_end_time", "invoice_time", "failure_reason", "payment_type"} <= set(app.result_table["columns"])
assert all(item.get("fuel_completion_time") for item in app.results)
assert sum(bool(item.get("match_details")) for item in app.results) == 2 * len(app.success_table.get_children())
assert all(app.results[int(item)]["status"] != "已匹配" for item in app.result_table.get_children())
assert all(app.results[int(item)]["failure_reason"] for item in app.result_table.get_children())
assert application_icon_path().is_file()
app.red_only.set(True)
app._refresh_preview()
assert len(app.result_table.get_children()) == sum(result["red"] for result in app.results)
app.red_only.set(False)
app.status_filter.set("金额不符")
app._refresh_preview()
assert len(app.result_table.get_children()) == sum(result["status"] == "金额不符" for result in app.results)
app.status_filter.set("全部状态")
app.red_only.set(True)
app._refresh_preview()
root.update()
from PIL import ImageGrab
import ctypes
from ctypes import wintypes
hwnd = ctypes.windll.user32.GetAncestor(root.winfo_id(), 2)
ctypes.windll.user32.SendMessageW.restype = ctypes.c_void_p
icon_handle = ctypes.windll.user32.SendMessageW(hwnd, 0x007F, 0, 0)
assert icon_handle, "Window small icon was not installed"
rectangle = wintypes.RECT()
ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rectangle))
image = ImageGrab.grab(bbox=(rectangle.left, rectangle.top, rectangle.right, rectangle.bottom))
image.save(qa / "gui.png")
failed_index = next((index for index, result in enumerate(app.results) if result["status"] != "已匹配"), None)
app.red_only.set(False)
app._refresh_preview()
table = app.result_table if failed_index is not None else app.success_table
selected_index = failed_index if failed_index is not None else next(index for index, result in enumerate(app.results) if result["kind"] == "fuel" and result["status"] == "已匹配")
table.selection_set(str(selected_index))
app._show_row_details() if table is app.result_table else app._show_row_details(type("Event", (), {"widget": table, "x": 0, "y": 0})())
root.update()
details = [window for window in root.winfo_children() if isinstance(window, tk.Toplevel)][-1]
details_hwnd = ctypes.windll.user32.GetAncestor(details.winfo_id(), 2)
assert ctypes.windll.user32.SendMessageW(details_hwnd, 0x007F, 0, 0)
details.destroy()
print("GUI failure/success tabs, invoice transaction time, log button and fuel red rows: passed")
root.destroy()
