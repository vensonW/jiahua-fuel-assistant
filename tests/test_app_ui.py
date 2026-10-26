"""Window title, button styling, and date/time selection behavior."""

from __future__ import annotations

import sys
import tempfile
import tkinter as tk
import unittest
from datetime import date
from pathlib import Path
from tkinter import ttk
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app import APP_TITLE, DateTimePicker, ReconciliationApp, load_last_input_path  # noqa: E402


class AppUiTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="gawa-ui-")
        self.addCleanup(self.temporary.cleanup)
        self.saved_settings = Path(self.temporary.name) / "settings.json"
        settings_patch = patch("app.settings_path", return_value=self.saved_settings)
        settings_patch.start()
        self.addCleanup(settings_patch.stop)
        availability_patch = patch("reconciliation.check_input_file_available")
        self.availability_check = availability_patch.start()
        self.addCleanup(availability_patch.stop)
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk is unavailable: {exc}")
        self.app = ReconciliationApp(self.root)
        self.root.update_idletasks()

    def tearDown(self):
        if hasattr(self, "root"):
            self.root.destroy()

    def test_title_buttons_and_date_controls(self):
        self.assertEqual(self.root.title(), APP_TITLE)
        self.assertEqual(APP_TITLE, "加华油站小助手")
        self.assertNotIn("Button.focus", str(ttk.Style(self.root).layout("TButton")))
        self.assertTrue(all(isinstance(control, ttk.Button) for control in self.app.setting_controls[4:7]))
        self.assertFalse(hasattr(self.app, "invoice_audit_path"))
        self.assertEqual(self.app.worker_count.get(), "20")
        self.assertEqual(self.app.open_button["text"], "打开已标注原文件")
        self.assertEqual(self.app.result_button["text"], "打开结果工作簿")
        self.assertEqual(str(self.app.result_button["state"]), "disabled")
        self.assertFalse(hasattr(self.app, "minutes_entry"))
        self.assertIn("早于开始时间不超过 10 分钟", self.app.rules_text.get())
        self.assertEqual(str(self.app.worker_combo["state"]), "readonly")
        self.assertEqual(self.app.result_table.heading("fuel_start_time")["text"], "加油明细切页-交易时间")
        self.assertEqual(self.app.result_table.heading("fuel_end_time")["text"], "发票-交易时间")
        self.assertIn("Checkmark.indicator", str(ttk.Style(self.root).layout("TCheckbutton")))
        unchecked, checked = self.app.checkbox_images
        self.assertNotEqual(unchecked.get(8, 12), checked.get(8, 12))

    def test_run_has_no_time_interval_setting(self):
        self.app.input_path.set("sample.xlsx")
        self.app.station_path = "sample.xlsx"
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread:
            self.app._start_run()
        self.assertNotIn("minutes_tolerance", thread.call_args.kwargs["args"][0])
        self.assertNotIn("invoice_audit_path", thread.call_args.kwargs["args"][0])
        self.assertEqual(str(self.app.progress["mode"]), "indeterminate")
        self.app.progress.stop()
        self.app._finish_run()

    def test_date_buttons_are_disabled_until_fuel_dates_are_loaded(self):
        controls = (*self.app.time_buttons, self.app.all_time_button)
        self.assertTrue(all(str(button["state"]) == "disabled" for button in controls))
        self.app.input_path.set("sample.xlsx")
        scan_id = self.app.station_scan_id
        self.assertTrue(all(str(button["state"]) == "disabled" for button in controls))
        self.app.events.put(("stations", (scan_id, "sample.xlsx", ["加华样例站"], None,
                                          ("2026-10-01", "2026-10-07"), "")))
        self.app._poll_events()
        self.assertTrue(all(str(button["state"]) == "normal" for button in controls))
        self.app.input_path.set("another.xlsx")
        self.assertTrue(all(str(button["state"]) == "disabled" for button in controls))
        self.app.events.put(("stations", (self.app.station_scan_id, "another.xlsx", [], None, None, "读取失败")))
        self.app._poll_events()
        self.assertTrue(all(str(button["state"]) == "disabled" for button in controls))
        self.app._finish_run()
        self.assertTrue(all(str(button["state"]) == "disabled" for button in controls))

    def test_last_selected_workbook_is_restored(self):
        workbook = Path(self.temporary.name) / "核对数据.xlsx"
        workbook.touch()
        with patch("app.filedialog.askopenfilename", return_value=str(workbook)):
            self.app._choose_input()
        self.assertEqual(load_last_input_path(), str(workbook))
        self.assertEqual(self.app.input_path.get(), str(workbook))
        workbook.unlink()
        self.assertEqual(load_last_input_path(), "")

    def test_gui_runs_without_separate_invoice_audit_file(self):
        self.app.input_path.set("sample.xlsx")
        self.app.station_path = "sample.xlsx"
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread, patch("app.messagebox.showerror") as error:
            self.app._start_run()
        thread.assert_called_once()
        error.assert_not_called()
        self.app._finish_run()

    def test_open_input_warns_before_starting_worker(self):
        from reconciliation import ReconciliationError

        self.app.input_path.set("sample.xlsx")
        self.app.station_path = "sample.xlsx"
        self.availability_check.side_effect = ReconciliationError("源文件已被 Excel/WPS 打开")
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread, \
                patch("app.messagebox.showwarning") as warning, patch.object(self.app, "_close") as close:
            self.app._start_run()
        thread.assert_not_called()
        warning.assert_called_once()
        self.assertEqual(warning.call_args.args[0], "请关闭源文件")
        close.assert_called_once()
        self.assertFalse(self.app.busy)

    def test_station_scan_in_progress_shows_reminder_instead_of_error(self):
        self.app.input_path.set("sample.xlsx")
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread, \
                patch("app.messagebox.showinfo") as reminder, patch("app.messagebox.showerror") as error:
            self.app._start_run()
        thread.assert_not_called()
        error.assert_not_called()
        self.assertEqual(reminder.call_args.args[0], "请稍候")
        self.assertIn("完成后即可开始核对", reminder.call_args.args[1])

    def test_station_scan_failure_still_reports_actual_problem(self):
        self.app.input_path.set("sample.xlsx")
        self.app.station_scan_error = "工作表缺少交易时间"
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread, \
                patch("app.messagebox.showinfo") as reminder, patch("app.messagebox.showerror") as error:
            self.app._start_run()
        thread.assert_not_called()
        reminder.assert_not_called()
        self.assertIn("工作表缺少交易时间", error.call_args.args[1])

    def test_manually_entered_valid_workbook_is_saved_on_run(self):
        workbook = Path(self.temporary.name) / "手动输入.xlsx"
        workbook.touch()
        self.app.input_path.set(str(workbook))
        self.app.station_path = str(workbook)
        with patch("app.threading.Thread"):
            self.app._start_run()
        self.assertEqual(load_last_input_path(), str(workbook))
        self.app._finish_run()

    def test_calendar_selects_date_and_optional_time(self):
        anchor = self.app.setting_controls[4]
        picker = DateTimePicker(self.root, self.app.start_text, "开始时间", anchor)
        picker._select_day(12)
        picker._apply()
        self.assertEqual(self.app.start_text.get(), "2026-10-12")

        picker = DateTimePicker(self.root, self.app.start_text, "开始时间", anchor)
        picker.precise.set(True)
        picker._toggle_time()
        picker.hour.set("21")
        picker.minute.set("28")
        picker.second.set("35")
        picker._apply()
        self.assertEqual(self.app.start_text.get(), "2026-10-12 21:28:35")

        self.app._clear_period()
        self.assertEqual((self.app.start_text.get(), self.app.end_text.get()), ("", ""))

    def test_end_picker_defaults_to_end_of_day(self):
        variable = tk.StringVar(value="2026-10-07")
        picker = DateTimePicker(self.root, variable, "结束时间", self.app.setting_controls[4])
        self.assertEqual((picker.hour.get(), picker.minute.get(), picker.second.get()), ("23", "59", "59"))
        picker.precise.set(True)
        picker._toggle_time()
        picker._apply()
        self.assertEqual(variable.get(), "2026-10-07 23:59:59")

    def test_calendar_allows_days_outside_fuel_date_range(self):
        variable = tk.StringVar(value="2026-10-03")
        self.app.fuel_date_range = ("2026-10-03", "2026-10-05")
        picker = DateTimePicker(self.root, variable, "开始时间", self.app.setting_controls[4])
        self.assertEqual(picker.month_combo["values"], tuple(f"{month:02d}" for month in range(1, 13)))
        picker._select_day(2)
        self.assertEqual(picker.selected_date, date(2026, 10, 2))
        picker._shift_month(-1)
        self.assertEqual(picker.month.get(), "09")
        picker._apply()
        self.assertEqual(variable.get(), "2026-09-02")

    def test_status_filter_lists_only_present_failures(self):
        self.assertEqual(self.app.status_combo["values"], ("全部状态",))
        self.app.results = [
            {"kind": "fuel", "status": "已匹配"},
            {"kind": "fuel", "status": "金额不符"},
            {"kind": "fuel", "status": "时间先后异常"},
            {"kind": "invoice", "status": "待确认"},
        ]
        self.app._update_status_options()
        self.assertEqual(self.app.status_combo["values"], ("全部状态", "金额不符", "时间先后异常"))
        self.app.status_filter.set("时间先后异常")
        self.app.results = [{"kind": "invoice", "status": "待确认"}, {"kind": "fuel", "status": "升数不符"}]
        self.app._update_status_options()
        self.assertEqual(self.app.status_combo["values"], ("全部状态", "升数不符"))
        self.assertEqual(self.app.status_filter.get(), "全部状态")

    def test_match_progress_updates_determinate_bar(self):
        self.app.events.put(("match_progress", (1, 5, 97)))
        self.app._poll_events()
        self.assertEqual(self.app.progress_text.get(), "核对进度：第 1/3 轮，5/97")
        self.assertAlmostEqual(float(self.app.progress["value"]), 25 * 5 / 97)
        self.assertEqual(int(float(self.app.progress["maximum"])), 100)
        self.app.events.put(("match_progress", (3, 97, 97)))
        self.app._poll_events()
        self.assertEqual(int(float(self.app.progress["value"])), 75)
        self.app.events.put(("workflow_progress", ("正在检查结果工作簿…", 90)))
        self.app._poll_events()
        self.assertEqual(int(float(self.app.progress["value"])), 90)
        self.assertEqual(self.app.progress_text.get(), "核对进度：正在检查结果工作簿…")

    def test_completion_summary_shows_elapsed_time(self):
        self.app.events.put(("done", {
            "output_path": "result.xlsx", "log_path": "result.log", "results": [],
            "annotated_input_path": "input.xlsx", "input_backup_path": "input_backup.xlsx", "red_invoice": 0,
            "total_fuel": 97, "total_invoice": 97, "paired": 90,
            "matched": 85, "red_fuel": 12, "elapsed_text": "2分34.56秒",
        }))
        self.app._poll_events()
        self.assertIn("核对耗时：2分34.56秒", self.app.summary_text.get())
        self.assertEqual(str(self.app.result_button["state"]), "normal")
        with patch.object(self.app, "_open_path") as opened, patch.object(Path, "is_file", return_value=True):
            self.app._open_result()
            self.app._open_result_workbook()
        self.assertEqual([call.args[0] for call in opened.call_args_list], [Path("input.xlsx"), Path("result.xlsx")])

    def test_completion_summary_shows_phase_times(self):
        self.app.events.put(("done", {
            "output_path": "result.xlsx", "log_path": "result.log", "results": [],
            "annotated_input_path": "input.xlsx", "input_backup_path": "input_backup.xlsx",
            "total_fuel": 1, "total_invoice": 1, "paired": 1, "matched": 1,
            "red_fuel": 0, "red_invoice": 0, "elapsed_text": "8分41秒",
            "phase_seconds": {"result_build": 65.0, "result_verify": 92.0,
                              "result_verify_input_load": 9.0, "result_verify_output_load": 11.0,
                              "result_verify_checks": 72.0,
                              "source_annotation": 37.0, "reversal_sheet_copy": 0.2,
                              "source_verify": 113.0, "source_verify_annotated_load": 10.0,
                              "source_verify_original_load": 12.0, "source_verify_checks": 91.0},
        }))
        self.app._poll_events()
        self.assertIn("源文件标注 37.00 秒", self.app.summary_text.get())
        self.assertIn("结果加载原件 9.00、加载结果 11.00、检查 72.00 秒", self.app.summary_text.get())

    def test_row_search_refresh_is_debounced(self):
        self.app.row_search.set("1698")
        first_refresh = self.app.search_refresh_id
        self.app.row_search.set("16980")
        self.assertIsNotNone(first_refresh)
        self.assertIsNotNone(self.app.search_refresh_id)
        self.assertNotEqual(first_refresh, self.app.search_refresh_id)

    def test_station_checkboxes_support_multiple_and_select_all(self):
        self.app.station_scan_id = 7
        self.app.input_path.set("sample.xlsx")
        self.app.station_scan_id = 7
        self.app.events.put(("stations", (7, "sample.xlsx", ["加华明珠加油站", "加华南屏加油站"], ("2026-10-01", "2026-10-07"), ("2026-10-02", "2026-10-06"), "")))
        self.app._poll_events()
        self.assertEqual(len(self.app.station_vars), 2)
        self.assertEqual(self.app.invoice_interval, ("2026-10-01", "2026-10-07"))
        self.assertEqual(self.app.fuel_date_range, ("2026-10-02", "2026-10-06"))
        self.assertEqual((self.app.start_text.get(), self.app.end_text.get()), ("2026-10-02", "2026-10-06 23:59:59"))
        self.app.start_text.set("2026-10-03")
        self.app._clear_period()
        self.assertEqual((self.app.start_text.get(), self.app.end_text.get()), ("2026-10-02", "2026-10-06 23:59:59"))
        self.app._set_all_stations(False)
        self.assertFalse(any(variable.get() for variable in self.app.station_vars.values()))
        self.app.station_vars["加华明珠加油站"].set(True)
        self.assertEqual([name for name, variable in self.app.station_vars.items() if variable.get()], ["加华明珠加油站"])
        self.app._set_all_stations(True)
        self.assertTrue(all(variable.get() for variable in self.app.station_vars.values()))
        self.assertTrue(all(widget.instate(("selected",)) for widget in self.app.station_items.winfo_children()))

    def test_disjoint_selected_time_can_start_reconciliation(self):
        self.app.input_path.set("sample.xlsx")
        self.app.station_path = "sample.xlsx"
        self.app.fuel_date_range = ("2026-10-01", "2026-10-07")
        self.app.invoice_interval = ("2026-10-01", "2026-10-07")
        self.app.start_text.set("2026-10-11")
        self.app.end_text.set("2026-10-11")
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread, patch("app.messagebox.showerror") as error:
            self.app._start_run()
        thread.assert_called_once()
        error.assert_not_called()
        self.assertEqual(thread.call_args.kwargs["args"][0]["start_text"], "2026-10-11")
        self.app._finish_run()

    def test_out_of_fuel_date_range_can_start_reconciliation(self):
        self.app.input_path.set("sample.xlsx")
        self.app.station_path = "sample.xlsx"
        self.app.fuel_date_range = ("2026-10-02", "2026-10-06")
        self.app.start_text.set("2026-10-01")
        self.app.end_text.set("2026-10-06")
        with patch.object(Path, "is_file", return_value=True), patch("app.threading.Thread") as thread, patch("app.messagebox.showerror") as error:
            self.app._start_run()
        thread.assert_called_once()
        error.assert_not_called()
        self.assertEqual(thread.call_args.kwargs["args"][0]["start_text"], "2026-10-01")
        self.app._finish_run()


if __name__ == "__main__":
    unittest.main()
