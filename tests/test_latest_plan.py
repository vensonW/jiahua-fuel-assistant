"""Latest-plan and inclusive date/time-range regression contracts."""

from __future__ import annotations

import inspect
import sys
import tempfile
import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook


sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import test_reconciliation as fixtures  # noqa: E402
from reconciliation import run_reconciliation  # noqa: E402


class LatestPlanTests(unittest.TestCase):
    """Use independent test cases; importing fixtures does not rerun old tests."""

    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="gawa-latest-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.input_path = self.directory / "原始数据.xlsx"
        self.output_directory = self.directory / "outputs"

    def make_workbook(self, fuels=None, invoices=None, footer=False):
        fixtures.ReconciliationContractTests.make_workbook(self, fuels, invoices, footer)

    def run_case(self, **options):
        summary = run_reconciliation(self.input_path, self.output_directory, **options)
        self.assertTrue(Path(summary["output_path"]).is_file())
        return summary

    @staticmethod
    def fuel_rows(summary):
        return [row for row in summary["results"] if row["source"] == "加油明细记录"]

    @staticmethod
    def invoice_rows(summary):
        return [row for row in summary["results"] if row["source"] == "发票"]

    def first_fuel(self, summary):
        return self.fuel_rows(summary)[0]

    @staticmethod
    def counterpart_row(result):
        return fixtures.ReconciliationContractTests.counterpart_row(result)

    def open_result(self, summary):
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        return workbook

    def appended_cell(self, sheet, header, row, label, original_columns):
        columns = [
            cell.column for cell in sheet[header]
            if cell.column > original_columns and cell.value == label
        ]
        self.assertEqual(len(columns), 1, f"{sheet.title}: appended {label!r}")
        return sheet.cell(row, columns[0])

    def assert_red_row(self, sheet, row):
        fixtures.ReconciliationContractTests.assert_red_row(self, sheet, row)

    def test_range_api_preserves_legacy_single_date_parameter(self):
        parameters = inspect.signature(run_reconciliation).parameters
        self.assertIn("date_text", parameters)
        self.assertIsNone(parameters["start_text"].default)
        self.assertIsNone(parameters["end_text"].default)
        self.make_workbook()
        self.assertEqual(self.first_fuel(self.run_case(date_text="2026-10-11"))["status"], "已匹配")

    def test_explicit_range_overrides_legacy_default_date(self):
        self.make_workbook(
            [fixtures.fuel(交易时间=datetime(2026, 10, 1, 10))],
            [fixtures.invoice(备注="交易时间：2026-10-01 10:05:00 油枪号：4", 开票时间="2026-10-01 10:06:00")],
        )
        summary = self.run_case(start_text="2026-10-01", end_text="2026-10-07")
        self.assertEqual(summary["matched"], 1)
        self.assertEqual(self.first_fuel(summary)["status"], "已匹配")
        self.assertEqual(len(summary["results"]), 2)

    def test_both_empty_range_endpoints_include_all_dates(self):
        self.make_workbook(
            [fixtures.fuel(交易时间=datetime(2027, 1, 1, 10))],
            [fixtures.invoice(备注="交易时间：2027-01-01 10:05:00 油枪号：4", 开票时间="2027-01-01 10:06:00")],
        )
        summary = self.run_case(start_text="", end_text="")
        self.assertEqual(summary["matched"], 1)
        self.assertTrue(all(row["status"] == "已匹配" for row in summary["results"]))

    def test_invalid_or_incomplete_range_is_rejected_without_writing(self):
        self.make_workbook()
        original = self.input_path.read_bytes()
        scenarios = [
            ("2026-10-01", ""), ("", "2026-10-07"),
            ("2026-10-01", None), (None, "2026-10-07"),
            ("2026-10-07", "2026-10-01"),
            ("2026-10-11 10:00:01", "2026-10-11 10:00:00"),
            ("2026-02-30", "2026-10-07"),
            ("2026-10-11 10:xx:00", "2026-10-12"),
            ("2026-10-11 24:00", "2026-10-12"),
            ("2026-10-11", "2026-10-11 11:61"),
            ("2026-10-11 trailing", "2026-10-12"),
        ]
        for start, end in scenarios:
            with self.subTest(start=start, end=end):
                with self.assertRaises(ValueError):
                    self.run_case(start_text=start, end_text=end)
                self.assertEqual(self.input_path.read_bytes(), original)
                self.assertFalse(list(self.output_directory.glob("*.xlsx")))

    def test_date_only_range_covers_entire_last_day(self):
        self.make_workbook(
            [fixtures.fuel(交易时间=datetime(2026, 10, 7, 23, 59, 59, 500000))],
            [fixtures.invoice(备注="交易时间：2026-10-07 23:59:59.500000 油枪号：4", 开票时间="2026-10-08 00:00:30")],
        )
        summary = self.run_case(start_text="2026-10-01", end_text="2026-10-07")
        self.assertEqual(summary["matched"], 1)
        self.assertEqual(self.first_fuel(summary)["status"], "已匹配")

    def test_datetime_range_includes_both_exact_endpoints(self):
        self.make_workbook(
            [fixtures.fuel(交易时间=datetime(2026, 10, 11, 10)), fixtures.fuel(枪号=9, 交易时间=datetime(2026, 10, 11, 10, 30), 流水号="F002")],
            [fixtures.invoice(备注="交易时间：2026-10-11 10:00:00 油枪号：4"), fixtures.invoice(备注="交易时间：2026-10-11 10:30:00 油枪号：9", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case(start_text="2026-10-11 10:00", end_text="2026-10-11 10:30:00")
        self.assertEqual(summary["matched"], 2)
        self.assertTrue(all(row["status"] == "已匹配" for row in summary["results"]))

    def test_each_side_uses_its_own_time_for_range_membership(self):
        self.make_workbook(
            [fixtures.fuel(交易时间=datetime(2026, 10, 11, 23, 50))],
            [fixtures.invoice(备注="交易时间：2026-10-12 00:10:00 油枪号：4", 开票时间="2026-10-11 23:59:00")],
        )
        summary = self.run_case(start_text="2026-10-11", end_text="2026-10-11")
        self.assertEqual(summary["paired"], 0)
        self.assertEqual(len(summary["results"]), 2)
        self.assertEqual(self.invoice_rows(summary)[0]["status"], "待确认")
        self.assertTrue(all(row["red"] for row in summary["results"]))
        self.assertEqual(self.first_fuel(summary)["status"], "无对应记录")
        included = self.run_case(start_text="2026-10-11", end_text="2026-10-12")
        self.assertEqual(included["matched"], 1)

    def test_range_crosses_month_and_year_and_retains_excluded_results(self):
        self.make_workbook(
            [fixtures.fuel(交易时间=datetime(2026, 12, 31, 23, 50)), fixtures.fuel(枪号=9, 交易时间=datetime(2027, 1, 2, 10), 流水号="F002")],
            [fixtures.invoice(备注="交易时间：2027-01-01 00:10:00 油枪号：4", 开票时间="2027-01-01 00:11:00"), fixtures.invoice(备注="交易时间：2027-01-02 10:01:00 油枪号：9", 开票时间="2027-01-02 10:02:00", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case(start_text="2026-12-31", end_text="2027-01-01")
        self.assertEqual(summary["matched"], 1)
        self.assertEqual(len(summary["results"]), 4)
        self.assertEqual(self.fuel_rows(summary)[1]["status"], "待确认")
        self.assertEqual(self.invoice_rows(summary)[1]["status"], "待确认")
        self.assertEqual(Decimal(summary["expected_total"]), Decimal("100"))
        self.assertEqual(Decimal(summary["actual_total"]), Decimal("100"))

    def test_invalid_remark_time_never_falls_back_to_valid_invoice_time(self):
        self.make_workbook(invoices=[fixtures.invoice(备注="交易时间：2026-10-11 10:xx:00 油枪号：4", 开票时间="2026-10-11 10:05:00")])
        summary = self.run_case(start_text="2026-10-11", end_text="2026-10-11")
        self.assertEqual(summary["matched"], 0)
        result = self.invoice_rows(summary)[0]
        self.assertEqual(result["status"], "待确认")
        self.assertTrue(result["red"])
        workbook = self.open_result(summary)
        sheet = workbook["发票"]
        self.assertIsNone(sheet.cell(3, sheet.max_column).value)

    def test_discount_formula_without_cache_is_not_described_as_empty_zero(self):
        self.make_workbook([fixtures.fuel(优惠金额="=5+5")])
        summary = self.run_case()
        result = self.first_fuel(summary)
        self.assertEqual(result["status"], "待确认")
        self.assertIsNone(result["expected"])
        self.assertIn("公式缺少计算缓存", result["basis"])
        self.assertNotIn("空优惠按0计算", result["basis"])
        self.assertTrue(result["red"])
        self.assertEqual(summary["matched"], 0)

    def test_signed_start_to_end_time_zero_thirty_and_negative_second(self):
        cases = [
            ("10:00:00", "已匹配", Decimal("0")),
            ("10:30:00", "已匹配", Decimal("30")),
            ("10:30:01", "时间超限", Decimal("30.016667")),
            ("09:59:59", "待确认", Decimal("-0.016667")),
        ]
        for stamp, status, difference in cases:
            with self.subTest(stamp=stamp):
                self.make_workbook(invoices=[fixtures.invoice(备注=f"交易时间：2026-10-11 {stamp} 油枪号：4")])
                summary = self.run_case()
                result = self.first_fuel(summary)
                self.assertTrue(result["counterpart"])
                self.assertEqual(result["status"], status)
                self.assertEqual(Decimal(result["time_diff"]), difference)
                self.assertEqual(result["red"], status != "已匹配")
                if status != "已匹配":
                    workbook = self.open_result(summary)
                    self.assert_red_row(workbook.worksheets[0], 5)
                    self.assert_red_row(workbook.worksheets[1], 3)

    def test_amount_difference_has_priority_over_reversed_time(self):
        self.make_workbook(invoices=[fixtures.invoice(备注="交易时间：2026-10-11 09:59:59 油枪号：4", **{"合计金额（含税）": 100.01})])
        summary = self.run_case()
        result = self.first_fuel(summary)
        self.assertTrue(result["counterpart"])
        self.assertEqual(result["status"], "金额不符")
        self.assertEqual(Decimal(result["amount_diff"]), Decimal("0.01"))
        self.assertLess(Decimal(result["time_diff"]), 0)
        self.assertTrue(any(word in result["basis"] for word in ("早于", "倒置", "先后", "负")))
        self.assertTrue(result["red"])

    def test_three_time_fields_and_last_completion_column_are_exported(self):
        self.make_workbook(invoices=[fixtures.invoice(备注="交易时间：2026-10-11 10:05:00 油枪号：4", 开票时间="2026-10-11 11:00:00")])
        summary = self.run_case()
        workbook = self.open_result(summary)
        for sheet, header, row, original_columns in [
            (workbook.worksheets[0], 4, 5, len(fixtures.FUEL_HEADERS)),
            (workbook.worksheets[1], 2, 3, len(fixtures.INVOICE_HEADERS)),
        ]:
            for label, expected in [
                ("加油开始时间", datetime(2026, 10, 11, 10)),
                ("加油结束时间", datetime(2026, 10, 11, 10, 5)),
                ("开票时间", datetime(2026, 10, 11, 11)),
            ]:
                cell = self.appended_cell(sheet, header, row, label, original_columns)
                self.assertEqual(cell.value, expected)
                self.assertIn("hh:mm:ss", cell.number_format)
        invoice_sheet = workbook.worksheets[1]
        self.assertEqual(invoice_sheet.cell(2, invoice_sheet.max_column).value, "加油完成时间")
        self.assertEqual(invoice_sheet.cell(3, invoice_sheet.max_column).value, datetime(2026, 10, 11, 10, 5))

    def test_fallback_invoice_time_does_not_become_fueling_end_time(self):
        self.make_workbook(invoices=[fixtures.invoice(备注="油枪号：4", 开票时间="2026-10-11 10:05:00")])
        summary = self.run_case(start_text="2026-10-11 10:00", end_text="2026-10-11 10:05")
        self.assertEqual(summary["matched"], 1)
        workbook = self.open_result(summary)
        for sheet, header, row, original_columns in [
            (workbook.worksheets[0], 4, 5, len(fixtures.FUEL_HEADERS)),
            (workbook.worksheets[1], 2, 3, len(fixtures.INVOICE_HEADERS)),
        ]:
            self.assertIsNone(self.appended_cell(sheet, header, row, "加油结束时间", original_columns).value)
            self.assertEqual(self.appended_cell(sheet, header, row, "开票时间", original_columns).value, datetime(2026, 10, 11, 10, 5))
        invoice_sheet = workbook.worksheets[1]
        self.assertIsNone(invoice_sheet.cell(3, invoice_sheet.max_column).value)
        self.assertIn("开票时间", self.first_fuel(summary)["basis"])

    def test_success_detail_is_identical_complete_wrapped_and_text_safe(self):
        card = "0001234567890123456789012345"
        number = "0026447200000093907128"
        self.make_workbook(
            [fixtures.fuel(卡号=card, 交易金额=290.19, 优惠金额=9.81, 加油量=38.4)],
            [fixtures.invoice(数电票号码=number, 数量=38.4, 备注="交易时间：2026-10-11 10:05:00 油枪号：4", 开票时间="2026-10-11 11:00:00", **{"合计金额（含税）": 280.38})],
        )
        summary = self.run_case()
        self.assertEqual(summary["matched"], 1)
        workbook = self.open_result(summary)
        details = []
        for sheet, header, row, original_columns in [
            (workbook.worksheets[0], 4, 5, len(fixtures.FUEL_HEADERS)),
            (workbook.worksheets[1], 2, 3, len(fixtures.INVOICE_HEADERS)),
        ]:
            cell = self.appended_cell(sheet, header, row, "匹配明细", original_columns)
            details.append(cell.value)
            self.assertIsInstance(cell.value, str)
            self.assertTrue(cell.alignment.wrap_text)
            self.assertGreater(sheet.column_dimensions[cell.column_letter].width, 30)
            self.assertGreater(sheet.row_dimensions[row].height or 0, 50)
        self.assertEqual(details[0], details[1])
        detail = details[0]
        self.assertIn("\n", detail)
        for text in [card, number, "加油明细记录", "发票", "95", "38.4", "290.19", "9.81", "280.38", "2026-10-11 10:00:00", "2026-10-11 10:05:00", "2026-10-11 11:00:00"]:
            self.assertIn(text, detail)
        for category in ("枪", "卡号", "油品", "发票", "开始", "结束", "开票", "加油升数", "发票升数", "交易金额", "优惠金额", "应开", "实际"):
            self.assertIn(category, detail)
        self.assertRegex(detail, r"加油明细记录[^\n]*5")
        self.assertRegex(detail, r"发票[^\n]*3")
        self.assertEqual(workbook.worksheets[0]["C5"].value, card)
        self.assertEqual(workbook.worksheets[1]["E3"].value, number)
        self.assertEqual(workbook.sheetnames, ["加油明细记录", "发票", "Sheet1"])

    def test_every_non_success_status_leaves_match_detail_blank(self):
        scenarios = [
            ([fixtures.fuel()], [fixtures.invoice(**{"合计金额（含税）": 100.01})]),
            ([fixtures.fuel()], [fixtures.invoice(数量=11)]),
            ([fixtures.fuel()], [fixtures.invoice(备注="交易时间：2026-10-11 10:30:01 油枪号：4")]),
            ([fixtures.fuel()], [fixtures.invoice(备注="交易时间：2026-10-11 09:59:59 油枪号：4")]),
            ([fixtures.fuel()], []),
            ([], [fixtures.invoice()]),
            ([fixtures.fuel()], [fixtures.invoice(备注="交易时间：格式有误 油枪号：4")]),
        ]
        for fuels, invoices in scenarios:
            with self.subTest(fuels=fuels, invoices=invoices):
                self.make_workbook(fuels, invoices)
                summary = self.run_case()
                self.assertEqual(summary["matched"], 0)
                workbook = self.open_result(summary)
                for result in summary["results"]:
                    sheet = workbook[result["source"]]
                    is_fuel = result["source"] == "加油明细记录"
                    cell = self.appended_cell(sheet, 4 if is_fuel else 2, result["row"], "匹配明细", len(fixtures.FUEL_HEADERS) if is_fuel else len(fixtures.INVOICE_HEADERS))
                    self.assertIsNone(cell.value)
                    self.assertTrue(result["red"])
                    self.assert_red_row(sheet, result["row"])

    def test_station_brand_prefix_can_match_single_station(self):
        self.make_workbook(
            [fixtures.fuel(油站名称="九洲港加油站")],
            [fixtures.invoice(备注="加油站点：加华九洲港加油站 交易时间：2026-10-11 10:05:00 油枪号：4")],
        )
        self.assertEqual(self.run_case()["matched"], 1)

    def test_multiple_stations_never_cross_pair_same_gun_time_and_amount(self):
        self.make_workbook(
            [fixtures.fuel(油站名称="九洲港加油站"), fixtures.fuel(油站名称="南屏加油站", 流水号="F002")],
            [fixtures.invoice(备注="加油站点：加华南屏加油站 交易时间：2026-10-11 10:05:00 油枪号：4"), fixtures.invoice(备注="加油站点：加华九洲港加油站 交易时间：2026-10-11 10:05:00 油枪号：4", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case()
        fuels = self.fuel_rows(summary)
        self.assertEqual(summary["matched"], 2)
        self.assertEqual(self.counterpart_row(fuels[0]), 4)
        self.assertEqual(self.counterpart_row(fuels[1]), 3)
        self.assertEqual(len({row["counterpart"] for row in fuels}), 2)

    def test_station_mismatch_cannot_be_success(self):
        self.make_workbook(
            [fixtures.fuel(油站名称="九洲港加油站")],
            [fixtures.invoice(备注="加油站点：加华南屏加油站 交易时间：2026-10-11 10:05:00 油枪号：4")],
        )
        summary = self.run_case()
        self.assertEqual(summary["matched"], 0)
        self.assertTrue(all(row["red"] for row in summary["results"]))

    def test_missing_station_evidence_is_not_success_in_multi_station_file(self):
        self.make_workbook(
            [fixtures.fuel(油站名称="九洲港加油站"), fixtures.fuel(油站名称="南屏加油站", 流水号="F002")],
            [fixtures.invoice(), fixtures.invoice(数电票号码="26447200000093907129")],
        )
        summary = self.run_case()
        self.assertEqual(summary["matched"], 0)
        self.assertEqual(len(summary["results"]), 4)
        self.assertTrue(all(row["red"] for row in summary["results"]))

    def test_duplicate_pairs_explicitly_mark_time_order_inference(self):
        self.make_workbook(
            [fixtures.fuel(), fixtures.fuel(交易时间=datetime(2026, 10, 11, 10, 5), 交易金额=99, 流水号="F002")],
            [fixtures.invoice(备注="交易时间：2026-10-11 10:01:00 油枪号：4", **{"合计金额（含税）": 99}), fixtures.invoice(备注="交易时间：2026-10-11 10:06:00 油枪号：4", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case()
        fuels = self.fuel_rows(summary)
        self.assertEqual(summary["paired"], 2)
        self.assertEqual(summary["matched"], 0)
        self.assertEqual(self.counterpart_row(fuels[0]), 3)
        self.assertEqual(self.counterpart_row(fuels[1]), 4)
        self.assertEqual(len({row["counterpart"] for row in fuels}), 2)
        for row in summary["results"]:
            self.assertIn("按时间顺序推定", row["basis"])
            self.assertEqual(row["status"], "金额不符")

    def test_duplicate_chronology_is_preserved_across_thirty_minute_boundary(self):
        self.make_workbook(
            [fixtures.fuel(), fixtures.fuel(交易时间=datetime(2026, 10, 11, 10, 20), 流水号="F002")],
            [fixtures.invoice(备注="交易时间：2026-10-11 10:31:00 油枪号：4"), fixtures.invoice(备注="交易时间：2026-10-11 10:40:00 油枪号：4", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case()
        fuels = self.fuel_rows(summary)
        self.assertEqual(summary["paired"], 2)
        self.assertEqual(self.counterpart_row(fuels[0]), 3)
        self.assertEqual(self.counterpart_row(fuels[1]), 4)
        self.assertEqual(fuels[0]["status"], "时间超限")
        self.assertEqual(fuels[1]["status"], "已匹配")
        self.assertEqual(Decimal(fuels[0]["time_diff"]), Decimal("31"))
        self.assertEqual(Decimal(fuels[1]["time_diff"]), Decimal("20"))


if __name__ == "__main__":
    unittest.main()
