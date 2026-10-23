"""Regression contracts for fuel/invoice reconciliation; all workbooks are temporary."""

from __future__ import annotations

import hashlib
import inspect
import re
import sys
import tempfile
import unittest
import zipfile
from copy import copy
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.utils import range_boundaries


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reconciliation import run_reconciliation  # noqa: E402
import reconciliation  # noqa: E402


FUEL_HEADERS = [
    "油站名称", "枪号", "卡号", "交易时间", "油品", "加油量", "单价", "交易金额",
    "优惠金额", "券金额", "卡余额", "卡类型", "付款方式", "油枪累计", "加油员",
    "车牌号", "客户编号", "流水号", "加油班次", "操作员",
]
INVOICE_HEADERS = [
    "发票种类", "红票标志", "发票代码", "发票号码", "数电票号码", "开票时间",
    "合计金额（含税）", "合计金额（不含税）", "合计税额", "备注", "开票人",
    "特定业务", "开票状态", "商品名称", "商品编码", "规格型号", "单位", "数量",
    "单价（含税）", "金额（含税）", "金额（不含税）", "税率", "税额", "开票单号",
    "创建时间", "交付手机", "交付邮箱", "部门门店", "开票方式", "流水号", "PDF地址",
]
SUCCESS = {"已匹配"}
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def fuel(**changes):
    row = {
        "油站名称": "加华样例站", "枪号": 4, "卡号": "001234",
        "交易时间": datetime(2026, 10, 11, 10, 0), "油品": 95, "加油量": 10,
        "单价": 10, "交易金额": 100, "优惠金额": 0, "流水号": "F001",
    }
    row.update(changes)
    return row


def invoice(**changes):
    row = {
        "发票种类": "数电发票（普通发票）", "红票标志": "蓝票", "数电票号码": "26447200000093907128",
        "开票时间": "2026-10-11 10:00:30", "合计金额（含税）": 100,
        "备注": "交易时间：2026-10-11 10:00:00   油枪号：4", "开票状态": "开票完成",
        "商品名称": "*汽油*95#车用汽油", "单位": "升", "数量": 10,
        "金额（含税）": 100, "开票单号": "I001", "流水号": "INV001",
    }
    row.update(changes)
    return row


class ReconciliationContractTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="gawa-contract-")
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.input_path = self.directory / "原始数据.xlsx"
        self.output_directory = self.directory / "outputs"

    def make_workbook(self, fuels=None, invoices=None, footer=False):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "加油明细记录"
        sheet.append(["加油明细记录"])
        sheet.append(["数据来源：样例"])
        sheet.append(["班次：2026-10-11"])
        sheet.append(FUEL_HEADERS)
        for row in fuels if fuels is not None else [fuel()]:
            sheet.append([row.get(header) for header in FUEL_HEADERS])
        if footer:
            sheet.append(["合计", None, None, None, None, 10, None, 100])
            sheet.append(["打印时间：2026-10-11 23:59:59"])
        sheet = workbook.create_sheet("发票")
        sheet.append(["开票区间：2026-10-11 00:00:00~2026-10-11 23:59:59"])
        sheet.append(INVOICE_HEADERS)
        for row in invoices if invoices is not None else [invoice()]:
            sheet.append([row.get(header) for header in INVOICE_HEADERS])
        if footer:
            sheet.append(["合计", None, None, None, None, None, 100])
            sheet.append(["打印时间：2026-10-11 23:59:59"])
        workbook.create_sheet("Sheet1")["B2"] = "原有备注保留"
        workbook.save(self.input_path)
        workbook.close()

    def run_case(self, **options):
        summary = run_reconciliation(self.input_path, self.output_directory, **options)
        self.assertTrue(Path(summary["output_path"]).is_file())
        self.assertEqual(Path(summary["output_path"]).parent.resolve(), self.output_directory.resolve())
        log_path = Path(summary["log_path"])
        self.assertEqual(log_path, Path(summary["output_path"]).with_suffix(".log"))
        self.assertTrue(log_path.is_file())
        log_path.read_text(encoding="utf-8")
        return summary

    @staticmethod
    def fuel_results(summary):
        return [row for row in summary["results"] if str(row["source"]) in {"加油", "加油明细记录", "fuel"}]

    def one_fuel_result(self, summary):
        rows = self.fuel_results(summary)
        self.assertEqual(len(rows), 1)
        return rows[0]

    @staticmethod
    def decimal(value):
        return Decimal(str(value))

    @staticmethod
    def counterpart_row(result):
        text = str(result["counterpart"])
        suffix = text.rsplit("!", 1)[-1]
        numbers = re.findall(r"\d+", suffix)
        return int(numbers[-1]) if numbers else None

    def assert_red_row(self, sheet, row):
        for cell in sheet[row]:
            self.assertEqual(cell.fill.fill_type, "solid", f"{sheet.title}!{cell.coordinate}")
            self.assertEqual(cell.fill.fgColor.type, "rgb", cell.coordinate)
            self.assertEqual(cell.fill.fgColor.rgb[-6:].upper(), "FFC7CE", cell.coordinate)

    def test_discount_and_invoice_total_are_used(self):
        self.make_workbook([fuel(交易金额=110, 优惠金额=10)], [invoice(**{"金额（含税）": 110})])
        summary = self.run_case()
        result = self.one_fuel_result(summary)
        self.assertEqual(result["status"], "已匹配")
        self.assertEqual(self.decimal(result["expected"]), Decimal("100"))
        self.assertEqual(self.decimal(result["actual"]), Decimal("100"))
        self.assertFalse(result["red"])
        self.assertEqual(summary["matched"], 1)

    def test_amount_comparison_is_exact_and_preserves_fractional_cents(self):
        for difference in ("0", "0.01", "-0.01", "0.0001", "-0.0001", "0.06", "-0.06", "0.0601", "-0.0601"):
            with self.subTest(difference=difference):
                delta = Decimal(difference)
                amount = float(Decimal("100") + delta)
                status = "已匹配" if delta == 0 else "金额不符"
                self.make_workbook(invoices=[invoice(**{"合计金额（含税）": amount})])
                summary = self.run_case()
                result = self.one_fuel_result(summary)
                self.assertEqual(result["status"], status)
                self.assertEqual(result["red"], delta != 0)
                self.assertEqual(self.decimal(result["amount_diff"]), delta)

    def test_configuration_cannot_bypass_exact_amount_comparison(self):
        parameters = inspect.signature(run_reconciliation).parameters
        self.assertEqual(parameters["amount_tolerance"].default, "0")
        self.assertEqual(parameters["minutes_tolerance"].default, "30")
        self.make_workbook()
        original_bytes = self.input_path.read_bytes()
        for tolerance in ("0.01", "0.0001", "0.06", "-0.01"):
            with self.subTest(tolerance=tolerance):
                with self.assertRaises(ValueError):
                    run_reconciliation(self.input_path, self.output_directory, amount_tolerance=tolerance)
        self.assertEqual(self.input_path.read_bytes(), original_bytes)

    def test_time_tolerance_includes_thirty_minutes_but_not_one_more_second(self):
        for stamp, status in [("10:30:00", "已匹配"), ("10:30:01", "时间超限")]:
            with self.subTest(stamp=stamp):
                self.make_workbook(invoices=[invoice(备注=f"交易时间：2026-10-11 {stamp} 油枪号：4")])
                result = self.one_fuel_result(self.run_case())
                self.assertEqual(result["status"], status)
                self.assertEqual(result["red"], status != "已匹配")

    def test_remark_transaction_time_takes_precedence(self):
        self.make_workbook(invoices=[invoice(开票时间="2026-10-11 15:00:00")])
        result = self.one_fuel_result(self.run_case())
        self.assertEqual(result["status"], "已匹配")
        self.assertEqual(self.decimal(result["time_diff"]), 0)
        self.assertIn("备注", result["basis"])

    def test_missing_remark_time_falls_back_to_invoice_time(self):
        for remark in ("油枪号：4", "交易时间：   油枪号：4"):
            with self.subTest(remark=remark):
                self.make_workbook(invoices=[invoice(备注=remark, 开票时间="2026-10-11 10:05:00")])
                result = self.one_fuel_result(self.run_case())
                self.assertEqual(result["status"], "已匹配")
                self.assertEqual(abs(self.decimal(result["time_diff"])), 5)
                self.assertIn("开票时间", result["basis"])

    def test_remark_completion_time_is_retained_without_an_extra_invoice_column(self):
        self.make_workbook(invoices=[invoice(开票时间="2026-10-11 15:00:00")])
        summary = self.run_case()
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        sheet = workbook["发票"]
        self.assertEqual(sheet.max_column, len(INVOICE_HEADERS))
        self.assertEqual([cell.value for cell in sheet[2]], INVOICE_HEADERS)
        self.assertEqual(workbook["加油明细记录"].max_column, len(FUEL_HEADERS))
        self.assertEqual(sheet["J3"].value, invoice()["备注"])
        for result in summary["results"]:
            self.assertEqual(result["fuel_end_time"], "2026-10-11 10:00:00")
        self.assertEqual(self.one_fuel_result(summary)["status"], "已匹配")

    def test_completion_time_stays_blank_when_matching_falls_back_to_invoice_time(self):
        for remark in ("油枪号：4", "交易时间：   油枪号：4"):
            with self.subTest(remark=remark):
                self.make_workbook(invoices=[invoice(备注=remark, 开票时间="2026-10-11 10:05:00")])
                summary = self.run_case()
                workbook = load_workbook(summary["output_path"])
                try:
                    sheet = workbook["发票"]
                    self.assertEqual(sheet.max_column, len(INVOICE_HEADERS))
                    self.assertEqual([cell.value for cell in sheet[2]], INVOICE_HEADERS)
                    self.assertTrue(all(result["fuel_end_time"] is None for result in summary["results"]))
                finally:
                    workbook.close()
                self.assertEqual(self.one_fuel_result(summary)["status"], "已匹配")
                self.assertEqual(self.decimal(self.one_fuel_result(summary)["time_diff"]), Decimal("5"))

    def test_completion_time_is_retained_for_invalid_amount_and_out_of_scope_invoices(self):
        self.make_workbook(invoices=[
            invoice(备注="交易时间：2026-10-11 10:07:00 油枪号：4", **{"合计金额（含税）": "金额未知"}),
            invoice(备注="交易时间：2026-10-12 11:08:09 油枪号：4", 开票时间="2026-10-12 11:08:39", 数电票号码="26447200000093907129"),
        ])
        summary = self.run_case()
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        sheet = workbook["发票"]
        self.assertEqual(sheet.max_column, len(INVOICE_HEADERS))
        invoices = [result for result in summary["results"] if result["source"] == "发票"]
        for result, stamp in zip(invoices, ["2026-10-11 10:07:00", "2026-10-12 11:08:09"]):
            self.assertEqual(result["fuel_end_time"], stamp)
            self.assertFalse(result["red"])
            self.assertTrue(all(cell.fill.fill_type is None for cell in sheet[result["row"]]))

    def test_same_time_multiple_guns_pair_to_corresponding_gun(self):
        self.make_workbook(
            [fuel(), fuel(枪号=9, 流水号="F002")],
            [invoice(备注="交易时间：2026-10-11 10:00:00 油枪号：9"), invoice(数电票号码="26447200000093907129")],
        )
        summary = self.run_case()
        rows = {result["row"]: result for result in self.fuel_results(summary)}
        self.assertEqual(self.counterpart_row(rows[5]), 4)
        self.assertEqual(self.counterpart_row(rows[6]), 3)
        self.assertEqual(summary["paired"], 2)
        self.assertEqual(summary["matched"], 2)
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        sheet = workbook["发票"]
        self.assertEqual(sheet.max_column, len(INVOICE_HEADERS))
        for result in summary["results"]:
            self.assertEqual(result["fuel_end_time"], "2026-10-11 10:00:00")

    def test_duplicate_candidates_pair_chronologically_and_only_once(self):
        self.make_workbook(
            [fuel(交易时间=datetime(2026, 10, 11, 10, 5), 流水号="F002"), fuel()],
            [invoice(), invoice(备注="交易时间：2026-10-11 10:06:00 油枪号：4", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case()
        rows = {result["row"]: result for result in self.fuel_results(summary)}
        self.assertEqual(set(rows), {5, 6})
        self.assertEqual(self.counterpart_row(rows[5]), 4)
        self.assertEqual(self.counterpart_row(rows[6]), 3)
        self.assertEqual(len({str(result["counterpart"]) for result in rows.values()}), 2)
        self.assertTrue(all(result["status"] == "已匹配" for result in rows.values()))
        self.assertTrue(all(any(word in result["basis"] for word in ("顺序", "重复", "时间")) for result in rows.values()))
        self.assertEqual(summary["paired"], 2)
        self.assertEqual(summary["total_fuel"], 2)
        self.assertEqual(summary["total_invoice"], 2)
        self.assertEqual(len(summary["results"]), 4)

    def test_duplicate_time_order_takes_precedence_over_equal_amounts(self):
        self.make_workbook(
            [fuel(), fuel(交易时间=datetime(2026, 10, 11, 10, 5), 交易金额=99, 流水号="F002")],
            [
                invoice(备注="交易时间：2026-10-11 10:01:00 油枪号：4", **{"合计金额（含税）": 99}),
                invoice(备注="交易时间：2026-10-11 10:06:00 油枪号：4", 数电票号码="26447200000093907129"),
            ],
        )
        summary = self.run_case()
        rows = {result["row"]: result for result in self.fuel_results(summary)}
        self.assertEqual(self.counterpart_row(rows[5]), 3)
        self.assertEqual(self.counterpart_row(rows[6]), 4)
        self.assertEqual(self.decimal(rows[5]["amount_diff"]), Decimal("-1"))
        self.assertEqual(self.decimal(rows[6]["amount_diff"]), Decimal("1"))
        self.assertTrue(all(result["status"] == "金额不符" for result in rows.values()))
        self.assertTrue(all(result["red"] for result in rows.values()))
        self.assertEqual(summary["paired"], 2)
        self.assertEqual(summary["matched"], 0)

    def test_any_amount_difference_has_actual_red_fill_and_filter_headers(self):
        self.make_workbook(invoices=[invoice(**{"合计金额（含税）": 99.99})])
        summary = self.run_case()
        self.assertEqual(self.one_fuel_result(summary)["status"], "金额不符")
        self.assertEqual(summary["matched"], 0)
        self.assertEqual(summary["red_fuel"], 1)
        self.assertEqual(summary["red_invoice"], 0)
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        sheet = workbook.worksheets[0]
        self.assert_red_row(sheet, 5)
        self.assertEqual(sheet.max_column, len(FUEL_HEADERS))
        self.assertTrue(sheet.auto_filter.ref)
        self.assertEqual(range_boundaries(sheet.auto_filter.ref), (1, 4, len(FUEL_HEADERS), 5))
        self.assertEqual([cell.value for cell in sheet[4]], FUEL_HEADERS)
        self.assertFalse(workbook.worksheets[1].auto_filter.ref)
        self.assertTrue(all(cell.fill.fill_type is None for cell in workbook.worksheets[1][3]))
        log = Path(summary["log_path"]).read_text(encoding="utf-8")
        for expected in ("金额不符", "金额差", "99.99", "-0.01", "加油明细记录", "发票"):
            self.assertIn(expected, log)

    def test_single_mismatch_and_multiple_mismatch_statuses(self):
        scenarios = [
            ({"合计金额（含税）": 101}, "金额不符"),
            ({"数量": 10.0001}, "升数不符"),
            ({"备注": "交易时间：2026-10-11 10:31:00 油枪号：4"}, "时间超限"),
            ({"合计金额（含税）": 101, "数量": 11}, "金额不符"),
        ]
        for changes, status in scenarios:
            with self.subTest(status=status):
                self.make_workbook(invoices=[invoice(**changes)])
                result = self.one_fuel_result(self.run_case())
                self.assertEqual(result["status"], status)
                self.assertTrue(result["red"])

    def test_nonzero_amount_difference_is_primary_for_confirmed_or_tentative_correspondence(self):
        scenarios = [({"数量": 11}, "升数不符"),
                     ({"备注": "交易时间：2026-10-11 10:00:00"}, "枪号缺失")]
        for changes, other_issue in scenarios:
            for difference in (Decimal("0.0001"), Decimal("-0.01"), Decimal("1")):
                with self.subTest(difference=difference, other_issue=other_issue):
                    changes = dict(changes, **{"合计金额（含税）": float(Decimal("100") + difference)})
                    self.make_workbook(invoices=[invoice(**changes)])
                    result = self.one_fuel_result(self.run_case())
                    self.assertTrue(result["counterpart"])
                    self.assertEqual(result["status"], "金额不符")
                    self.assertEqual(self.decimal(result["amount_diff"]), difference)
                    self.assertIn(other_issue, result["basis"])
                    self.assertTrue(result["red"])

    def test_unreadable_amount_requires_review_instead_of_amount_comparison(self):
        cases = [([fuel(交易金额="无法读取")], [invoice()], "加油明细记录"),
                 ([fuel()], [invoice(**{"合计金额（含税）": "无法读取"})], "发票")]
        for fuels, invoices, invalid_source in cases:
            with self.subTest(invalid_source=invalid_source):
                self.make_workbook(fuels, invoices)
                summary = self.run_case()
                invalid = next(result for result in summary["results"] if result["source"] == invalid_source)
                self.assertEqual(invalid["status"], "待确认")
                self.assertIsNone(invalid["amount_diff"])
                self.assertEqual(invalid["red"], invalid_source == "加油明细记录")
                self.assertEqual(summary["matched"], 0)

    def test_missing_counterparts_receive_results_on_both_sides(self):
        self.make_workbook([fuel()], [])
        summary = self.run_case()
        self.assertEqual(self.one_fuel_result(summary)["status"], "无对应记录")
        self.assertTrue(self.one_fuel_result(summary)["red"])
        self.make_workbook([], [invoice()])
        summary = self.run_case()
        self.assertEqual(summary["total_invoice"], 1)
        self.assertEqual(len(summary["results"]), 1)
        self.assertEqual(summary["results"][0]["status"], "无对应记录")
        self.assertFalse(summary["results"][0]["red"])

    def test_conflicting_gun_or_oil_cannot_be_automatically_matched(self):
        for changes in [{"备注": "交易时间：2026-10-11 10:00:00 油枪号：9"}, {"商品名称": "*汽油*92#车用汽油"}]:
            with self.subTest(changes=changes):
                self.make_workbook(invoices=[invoice(**changes)])
                result = self.one_fuel_result(self.run_case())
                self.assertNotIn(result["status"], SUCCESS)
                self.assertTrue(result["red"])

    def test_duplicate_invoice_numbers_and_same_invoice_multiple_lines_require_review(self):
        for volume in (10, 5):
            with self.subTest(invoice_line_volume=volume):
                self.make_workbook(
                    [fuel()],
                    [invoice(数量=volume), invoice(数量=volume, 流水号="INV002")],
                )
                summary = self.run_case()
                self.assertEqual(summary["total_invoice"], 2)
                self.assertEqual(len(summary["results"]), 3)
                self.assertEqual(summary["matched"], 0)
                self.assertEqual(summary["paired"], 0)
                invoice_rows = [result for result in summary["results"] if result not in self.fuel_results(summary)]
                self.assertTrue(all(result["status"] == "待确认" for result in invoice_rows))
                self.assertTrue(all(not result["red"] for result in invoice_rows))
                self.assertTrue(all(not result["counterpart"] for result in invoice_rows))

    def test_original_cells_sheet_count_and_full_invoice_number_are_preserved(self):
        self.make_workbook()
        original_bytes = self.input_path.read_bytes()
        before = load_workbook(self.input_path)
        original_values = {sheet.title: list(sheet.values) for sheet in before}
        before.close()
        summary = self.run_case()
        self.assertEqual(hashlib.sha256(original_bytes).digest(), hashlib.sha256(self.input_path.read_bytes()).digest())
        after = load_workbook(summary["output_path"])
        self.addCleanup(after.close)
        self.assertEqual(after.sheetnames, list(original_values))
        for name, rows in original_values.items():
            self.assertEqual(after[name].max_column, len(rows[0]))
            self.assertEqual(after[name].max_row, len(rows))
            for row_index, row in enumerate(rows, start=1):
                for column, value in enumerate(row, start=1):
                    self.assertEqual(after[name].cell(row_index, column).value, value, (name, row_index, column))
                    self.assertIsNone(after[name].cell(row_index, column).comment)
        self.assertEqual(after["发票"]["E3"].value, "26447200000093907128")
        self.assertIsInstance(after["发票"]["E3"].value, str)

    def test_original_comments_dimensions_and_invoice_presentation_are_preserved(self):
        self.make_workbook(invoices=[invoice(**{"合计金额（含税）": 100.01})])
        workbook = load_workbook(self.input_path)
        for sheet, row in [(workbook.worksheets[0], 5), (workbook.worksheets[1], 3)]:
            sheet.cell(row, 1).comment = Comment("原始人工备注，不要覆盖", "原作者")
            sheet.cell(row, 4).comment = Comment("另一条原始备注", "复核员")
            sheet.row_dimensions[row].height = 19.5
            sheet.row_dimensions[1].height = 28
            sheet.column_dimensions["A"].width = 18.25
            sheet.column_dimensions["J"].width = 36.75
        sheet = workbook["发票"]
        sheet.auto_filter.ref = "A2:AE3"
        sheet.auto_filter.add_filter_column(1, ["蓝票"])
        sheet.freeze_panes = "C3"
        sheet.sheet_view.zoomScale = 85
        sheet["G3"].fill = PatternFill("solid", fgColor="FFF2CC")
        workbook.save(self.input_path)
        workbook.close()
        before = load_workbook(self.input_path)
        self.addCleanup(before.close)
        summary = self.run_case(red_only=True)
        after = load_workbook(summary["output_path"])
        self.addCleanup(after.close)
        self.assertEqual(after.sheetnames, before.sheetnames)
        for original, result in zip(before, after):
            self.assertEqual((result.max_row, result.max_column), (original.max_row, original.max_column))
            self.assertEqual(list(result.values), list(original.values))
            self.assertEqual(
                {key: dict(value) for key, value in result.column_dimensions.items()},
                {key: dict(value) for key, value in original.column_dimensions.items()},
            )
            for row, dimension in original.row_dimensions.items():
                self.assertEqual(result.row_dimensions[row].height, dimension.height)
            for row in original:
                for cell in row:
                    actual = result[cell.coordinate]
                    self.assertEqual(actual.comment, cell.comment, cell.coordinate)
                    if original.title == "发票":
                        self.assertEqual(copy(actual.fill), copy(cell.fill), cell.coordinate)
            if original.title == "发票":
                self.assertEqual(result.auto_filter, original.auto_filter)
                self.assertEqual(result.freeze_panes, original.freeze_panes)
                self.assertEqual(result.sheet_view.zoomScale, original.sheet_view.zoomScale)
                self.assertEqual(
                    {key: dict(value) for key, value in result.row_dimensions.items()},
                    {key: dict(value) for key, value in original.row_dimensions.items()},
                )
        self.assert_red_row(after["加油明细记录"], 5)
        self.assertEqual(summary["red_invoice"], 0)

    def test_companion_log_contains_every_record_without_adding_workbook_information(self):
        self.make_workbook(
            [fuel(), fuel(交易时间=datetime(2026, 10, 11, 12), 流水号="F002")],
            [invoice(), invoice(备注="交易时间：2026-10-11 12:00:00 油枪号：4", 数电票号码="26447200000093907129", **{"合计金额（含税）": 100.01})],
        )
        summary = self.run_case()
        log = Path(summary["log_path"]).read_text(encoding="utf-8")
        self.assertEqual(len(summary["results"]), 4)
        for result in summary["results"]:
            for text in (result["source"], str(result["row"]), result["status"], result["counterpart"], result["invoice_number"]):
                self.assertIn(text, log)
        for text in ("金额差", "升数差", "时间差", "2026-10-11 10:00:00", "0.01", "100.01"):
            self.assertIn(text, log)
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        self.assertEqual(workbook.sheetnames, ["加油明细记录", "发票", "Sheet1"])
        self.assertEqual(workbook.worksheets[0].max_column, len(FUEL_HEADERS))
        self.assertEqual(workbook.worksheets[1].max_column, len(INVOICE_HEADERS))
        self.assertTrue(all(cell.comment is None for sheet in workbook for row in sheet for cell in row))

    def test_totals_and_print_footer_are_not_counted_as_records(self):
        self.make_workbook(footer=True)
        summary = self.run_case()
        self.assertEqual(summary["total_fuel"], 1)
        self.assertEqual(summary["total_invoice"], 1)
        self.assertEqual(len(summary["results"]), 2)
        self.assertEqual(summary["status_counts"].get("已匹配"), 1)
        self.assertEqual(sum(summary["status_counts"].values()), 1)

    def test_new_outputs_never_overwrite_an_existing_result(self):
        self.make_workbook()
        first = self.run_case()
        old_path = Path(first["output_path"])
        old_bytes = old_path.read_bytes()
        second = self.run_case()
        self.assertNotEqual(old_path.resolve(), Path(second["output_path"]).resolve())
        self.assertEqual(old_path.read_bytes(), old_bytes)

    def test_existing_companion_log_is_preserved_and_entire_pair_uses_new_name(self):
        self.make_workbook()
        first = self.run_case()
        first_workbook = Path(first["output_path"])
        first_log = Path(first["log_path"])
        original_log = first_log.read_bytes()
        first_workbook.unlink()
        second = self.run_case()
        self.assertNotEqual(Path(second["output_path"]), first_workbook)
        self.assertNotEqual(Path(second["log_path"]), first_log)
        self.assertFalse(first_workbook.exists(), "旧日志旁不能留下孤立的工作簿占位文件")
        self.assertEqual(first_log.read_bytes(), original_log)
        self.assertEqual(
            set(self.output_directory.iterdir()),
            {first_log, Path(second["output_path"]), Path(second["log_path"])},
        )

    def test_failed_final_log_replace_removes_undelivered_pair_and_temporary_files(self):
        self.make_workbook()
        self.output_directory.mkdir()
        preserved = {}
        for name in ("已有文件.xlsx", "已有文件.log", "人工备注.txt"):
            path = self.output_directory / name
            preserved[path] = ("必须保留：" + name).encode("utf-8")
            path.write_bytes(preserved[path])
        original_input = self.input_path.read_bytes()
        original_replace = reconciliation._replace_file
        attempted_log_targets = []

        def replace_except_log(source, target):
            if Path(target).suffix == ".log":
                attempted_log_targets.append(Path(target))
                raise PermissionError("模拟日志被永久占用")
            return original_replace(source, target)

        with patch("reconciliation._replace_file", side_effect=replace_except_log):
            with self.assertRaises(reconciliation.ReconciliationError) as failure:
                run_reconciliation(self.input_path, self.output_directory)
        self.assertIsInstance(failure.exception.__cause__, PermissionError)
        self.assertEqual(len(attempted_log_targets), 1)
        self.assertEqual(set(self.output_directory.iterdir()), set(preserved))
        for path, content in preserved.items():
            self.assertEqual(path.read_bytes(), content)
        self.assertEqual(self.input_path.read_bytes(), original_input)

    def test_non_target_date_rows_are_still_annotated_for_review(self):
        self.make_workbook(
            [fuel(), fuel(交易时间=datetime(2026, 10, 12, 10), 流水号="F002")],
            [invoice(), invoice(备注="交易时间：2026-10-12 10:00:00 油枪号：4", 开票时间="2026-10-12 10:00:30", 数电票号码="26447200000093907129")],
        )
        summary = self.run_case(date_text="2026-10-11")
        self.assertEqual(len(summary["results"]), 4)
        rows = {result["row"]: result for result in self.fuel_results(summary)}
        self.assertEqual(rows[5]["status"], "已匹配")
        self.assertEqual(rows[6]["status"], "待确认")
        self.assertTrue(rows[6]["red"])

    def test_empty_discount_means_zero_but_invalid_discount_requires_review(self):
        for discount, status in [(None, "已匹配"), ("", "已匹配"), ("金额未知", "待确认"), ("=5+5", "待确认")]:
            with self.subTest(discount=discount):
                self.make_workbook([fuel(优惠金额=discount)])
                result = self.one_fuel_result(self.run_case())
                self.assertEqual(result["status"], status)

    def test_invalid_time_cannot_be_mistaken_for_a_successful_match(self):
        for invalid_time in ("时间未知", "2026-10-11 00:0x:00"):
            with self.subTest(invalid_time=invalid_time):
                self.make_workbook(
                    [fuel(交易时间=invalid_time)],
                    [invoice(备注="交易时间：2026-10-11 00:00:00 油枪号：4", 开票时间="2026-10-11 00:00:30")],
                )
                result = self.one_fuel_result(self.run_case())
                self.assertEqual(result["status"], "待确认")
                self.assertTrue(result["red"])
        self.make_workbook(invoices=[invoice(备注="油枪号：4", 开票时间="无法解析")])
        summary = self.run_case()
        self.assertTrue(all(result["status"] not in SUCCESS for result in summary["results"]))

    def test_red_only_sets_excel_color_filter_and_hides_non_red_rows(self):
        self.make_workbook(
            [fuel(), fuel(交易时间=datetime(2026, 10, 11, 12), 流水号="F002")],
            [invoice(), invoice(备注="交易时间：2026-10-11 12:00:00 油枪号：4", 数电票号码="26447200000093907129", **{"合计金额（含税）": 100.01})],
        )
        summary = self.run_case(red_only=True)
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        self.assertTrue(workbook.worksheets[0].row_dimensions[5].hidden)
        self.assertFalse(workbook.worksheets[0].row_dimensions[6].hidden)
        self.assertFalse(workbook.worksheets[1].row_dimensions[3].hidden)
        self.assertFalse(workbook.worksheets[1].row_dimensions[4].hidden)
        self.assertEqual(range_boundaries(workbook.worksheets[0].auto_filter.ref)[2], len(FUEL_HEADERS))
        self.assertFalse(workbook.worksheets[1].auto_filter.ref)
        with zipfile.ZipFile(summary["output_path"]) as archive:
            styles = ElementTree.fromstring(archive.read("xl/styles.xml"))
            dxfs = styles.find("s:dxfs", NS)
            self.assertIsNotNone(dxfs)
            document = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            column = document.find("s:autoFilter/s:filterColumn", NS)
            self.assertEqual(column.attrib["colId"], "0")
            color_filter = column.find("s:colorFilter", NS)
            self.assertIsNotNone(color_filter)
            dxf = list(dxfs)[int(color_filter.attrib["dxfId"])]
            color = dxf.find("s:fill/s:patternFill/s:fgColor", NS)
            self.assertIsNotNone(color)
            self.assertEqual(color.attrib["rgb"][-6:].upper(), "FFC7CE")
            invoice_document = ElementTree.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            self.assertIsNone(invoice_document.find("s:autoFilter", NS))


if __name__ == "__main__":
    unittest.main()
