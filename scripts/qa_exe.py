"""End-to-end validation of the delivered single-file Windows executable."""
from pathlib import Path
import argparse
import hashlib
import json
import subprocess

from PIL import Image
import win32con
import win32gui
import win32ui


root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument("--icon-only", action="store_true", help="Only inspect the embedded EXE icon")
args = parser.parse_args()
exe = root / "加华加油与发票核对助手.exe"
qa = root / "outputs" / "qa" / ("red_exe_icon" if args.icon_only else "original_marking")
qa.mkdir(parents=True, exist_ok=True)
summary_path = qa / "exe_summary.json"
arguments = [str(exe), "--input", str(root / "中南加华加油和开票数据核对.xlsx"),
             "--output-dir", str(root / "outputs"), "--start", "2026-10-11", "--end", "2026-10-11",
             "--summary-json", str(summary_path)]
summary = None
if not args.icon_only:
    subprocess.run(arguments, check=True, timeout=45, creationflags=subprocess.CREATE_NO_WINDOW)
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    assert summary["paired"] == 97 and summary["matched"] == 15
    assert summary["red_fuel"] == 82 and summary["red_invoice"] == 0
    log_path = Path(summary["log_path"])
    assert log_path == Path(summary["output_path"]).with_suffix(".log")
    log_text = log_path.read_text(encoding="utf-8")
    assert "加油开始时间" in log_text and "匹配明细" in log_text
    assert summary["period_start"] == "2026-10-11 00:00:00"
    assert summary["period_end"] == "2026-10-11 23:59:59.999999"
    subprocess.run([str(exe), "--smoke-gui"], check=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW)

large, small = win32gui.ExtractIconEx(str(exe), 0)
screen = win32gui.GetDC(0)
screen_dc = win32ui.CreateDCFromHandle(screen)
memory_dc = screen_dc.CreateCompatibleDC()
bitmap = win32ui.CreateBitmap()
bitmap.CreateCompatibleBitmap(screen_dc, 32, 32)
memory_dc.SelectObject(bitmap)
memory_dc.FillSolidRect((0, 0, 32, 32), 0xFFFFFF)
win32gui.DrawIconEx(memory_dc.GetSafeHdc(), 0, 0, large[0], 32, 32, 0, None, win32con.DI_NORMAL)
preview = Image.frombuffer("RGB", (32, 32), bitmap.GetBitmapBits(True), "raw", "BGRX", 0, 1)
preview.save(qa / "exe_icon.png")
pixels = list(preview.getdata())
assert sum(r > g + 50 and r > b + 50 for r, g, b in pixels) > 200
assert sum(r > 220 and g > 220 and b > 220 for r, g, b in pixels) > 20
memory_dc.DeleteDC()
win32gui.DeleteObject(bitmap.GetHandle())
win32gui.ReleaseDC(0, screen)
for handle in large + small:
    win32gui.DestroyIcon(handle)
evidence = dict(icon_only=args.icon_only, exe_cli_exit=None if args.icon_only else 0,
                native_exe_gui_exit=None if args.icon_only else 0, red_gawa_exe_icon=True,
                exe_sha256=hashlib.sha256(exe.read_bytes()).hexdigest(),
                source_sha256=hashlib.sha256((root / "中南加华加油和开票数据核对.xlsx").read_bytes()).hexdigest(),
                summary=summary)
(qa / "final_validation.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
print(json.dumps(evidence, ensure_ascii=False))
