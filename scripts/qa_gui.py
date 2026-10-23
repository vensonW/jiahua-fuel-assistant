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
qa = project / "outputs" / "qa" / "latest_period"
qa.mkdir(parents=True, exist_ok=True)
app.output_directory.set(str(qa / "gui_run"))
app.start_text.set("2026-10-11 21:28:03")
app.end_text.set("2026-10-11 23:59:59")
app._start_run()
deadline = time.monotonic() + 30
while app.busy and time.monotonic() < deadline:
    root.update()
    time.sleep(0.02)
assert not app.busy and app.output_path and app.output_path.is_file(), app.status_text.get()
assert len(app.result_table.get_children()) == 194
assert {"fuel_start_time", "fuel_end_time", "invoice_time", "match_details"} <= set(app.result_table["columns"])
assert all(item.get("fuel_completion_time") for item in app.results)
assert sum(bool(item.get("match_details")) for item in app.results) == 30
assert application_icon_path().is_file()
app.red_only.set(True)
app._refresh_preview()
assert len(app.result_table.get_children()) == 164
app.red_only.set(False)
app.status_filter.set("已匹配")
app._refresh_preview()
assert len(app.result_table.get_children()) == 30
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
matched_index = next(index for index, result in enumerate(app.results) if result["status"] == "已匹配")
app.red_only.set(False)
app._refresh_preview()
app.result_table.selection_set(str(matched_index))
app._show_row_details()
root.update()
details = [window for window in root.winfo_children() if isinstance(window, tk.Toplevel)][-1]
details_hwnd = ctypes.windll.user32.GetAncestor(details.winfo_id(), 2)
assert ctypes.windll.user32.SendMessageW(details_hwnd, 0x007F, 0, 0)
details.destroy()
print("GUI time range, brand icons, 194-row preview, 164 red rows, 30 successful details: passed")
root.destroy()
