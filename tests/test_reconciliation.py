"""Regression contracts for fuel/invoice reconciliation; all workbooks are temporary."""

from __future__ import annotations

import hashlib
import inspect
import re
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree

from openpyxl import Workbook, load_workbook
from openpyxl.utils import range_boundaries


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reconciliation import run_reconciliation  # noqa: E402


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

    def test_invoice_last_column_contains_remark_completion_time_as_excel_datetime(self):
        self.make_workbook(invoices=[invoice(开票时间="2026-10-11 15:00:00")])
        summary = self.run_case()
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        sheet = workbook["发票"]
        self.assertEqual(sheet.cell(2, sheet.max_column).value, "加油完成时间")
        cell = sheet.cell(3, sheet.max_column)
        self.assertEqual(cell.value, datetime(2026, 10, 11, 10))
        self.assertIsInstance(cell.value, datetime)
        self.assertEqual(cell.number_format, "yyyy-mm-dd hh:mm:ss")
        self.assertNotIn("加油完成时间", [cell.value for cell in workbook["加油明细记录"][4]])
        self.assertNotIn("加油完成时间", [cell.value for row in workbook["Sheet1"] for cell in row])
        self.assertEqual(self.one_fuel_result(summary)["status"], "已匹配")

    def test_completion_time_stays_blank_when_matching_falls_back_to_invoice_time(self):
        for remark in ("油枪号：4", "交易时间：   油枪号：4"):
            with self.subTest(remark=remark):
                self.make_workbook(invoices=[invoice(备注=remark, 开票时间="2026-10-11 10:05:00")])
                summary = self.run_case()
                workbook = load_workbook(summary["output_path"])
                try:
                    sheet = workbook["发票"]
                    self.assertEqual(sheet.cell(2, sheet.max_column).value, "加油完成时间")
                    self.assertIsNone(sheet.cell(3, sheet.max_column).value)
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
        self.assertEqual(sheet.cell(2, sheet.max_column).value, "加油完成时间")
        for row, stamp in [(3, datetime(2026, 10, 11, 10, 7)), (4, datetime(2026, 10, 12, 11, 8, 9))]:
            self.assertEqual(sheet.cell(row, sheet.max_column).value, stamp)
            self.assert_red_row(sheet, row)

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
        self.assertEqual(sheet.cell(2, sheet.max_column).value, "加油完成时间")
        self.assertEqual(sheet.cell(3, sheet.max_column).value, datetime(2026, 10, 11, 10))
        self.assertEqual(sheet.cell(4, sheet.max_column).value, datetime(2026, 10, 11, 10))

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
        self.assertEqual(summary["red_invoice"], 1)
        workbook = load_workbook(summary["output_path"])
        self.addCleanup(workbook.close)
        for sheet, header, row in [(workbook.worksheets[0], 4, 5), (workbook.worksheets[1], 2, 3)]:
            self.assert_red_row(sheet, row)
            self.assertTrue(sheet.auto_filter.ref)
            self.assertEqual(sheet.auto_filter.ref.split(":")[0], f"A{header}")
            self.assertEqual(range_boundaries(sheet.auto_filter.ref)[2], sheet.max_column)
            self.assertIn("核对状态", [cell.value for cell in sheet[header]])

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
                self.assertTrue(invalid["red"])
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
        self.assertTrue(summary["results"][0]["red"])

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
                self.assertTrue(all(result["red"] for result in invoice_rows))
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
            for row_index, row in enumerate(rows, start=1):
                for column, value in enumerate(row, start=1):
                    self.assertEqual(after[name].cell(row_index, column).value, value, (name, row_index, column))
        self.assertEqual(after["发票"]["E3"].value, "26447200000093907128")
        self.assertIsInstance(after["发票"]["E3"].value, str)

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
        self.assertTrue(workbook.worksheets[1].row_dimensions[3].hidden)
        self.assertFalse(workbook.worksheets[1].row_dimensions[4].hidden)
        for sheet in workbook.worksheets[:2]:
            self.assertEqual(range_boundaries(sheet.auto_filter.ref)[2], sheet.max_column)
        with zipfile.ZipFile(summary["output_path"]) as archive:
            styles = ElementTree.fromstring(archive.read("xl/styles.xml"))
            dxfs = styles.find("s:dxfs", NS)
            self.assertIsNotNone(dxfs)
            for index in (1, 2):
                document = ElementTree.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
                color_filter = document.find("s:autoFilter/s:filterColumn/s:colorFilter", NS)
                self.assertIsNotNone(color_filter)
                dxf = list(dxfs)[int(color_filter.attrib["dxfId"])]
                color = dxf.find("s:fill/s:patternFill/s:fgColor", NS)
                self.assertIsNotNone(color)
                self.assertEqual(color.attrib["rgb"][-6:].upper(), "FFC7CE")


if __name__ == "__main__":
    unittest.main()
