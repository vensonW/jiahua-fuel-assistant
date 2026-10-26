"""Preview filtering, failure highlighting, and counterpart link behavior."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app import ReconciliationApp  # noqa: E402


class PreviewTable:
    def __init__(self, columns):
        self.columns = columns
        self.rows = {}

    def __getitem__(self, key):
        return self.columns if key == "columns" else None

    def get_children(self):
        return tuple(self.rows)

    def delete(self, *items):
        for item in items:
            del self.rows[item]

    def insert(self, _parent, _position, *, iid, values, tags=()):
        self.rows[iid] = (values, tags)

    def identify_row(self, position):
        return str(position) if str(position) in self.rows else ""

    def identify_column(self, position):
        return "#3" if position == 3 else "#2"


class AppPreviewTests(unittest.TestCase):
    def setUp(self):
        self.app = ReconciliationApp.__new__(ReconciliationApp)
        self.app.results = [
            dict(kind="fuel", source="加油", counterpart="'发票'!A3", status="已匹配", failure_reason="", red=False, gun="4", fuel_volume="10"),
            dict(kind="invoice", source="发票", counterpart="'加油'!A5", status="已匹配", failure_reason="", red=False),
            dict(kind="fuel", source="加油", counterpart="'发票'!A4", status="金额不符", failure_reason="金额不符", red=True),
            dict(kind="invoice", source="发票", counterpart="'加油'!A5", status="金额不符", failure_reason="金额不符", red=False),
        ]
        self.app.red_only = Mock()
        self.app.red_only.get.return_value = False
        self.app.status_filter = Mock()
        self.app.status_filter.get.return_value = "全部状态"
        self.app.row_search = Mock()
        self.app.row_search.get.return_value = ""
        self.app.preview_text = Mock()
        self.app.success_text = Mock()
        self.app.result_table = PreviewTable(("source", "row", "counterpart", "status", "payment_type", "failure_reason"))
        self.app.success_table = PreviewTable(("source", "row", "counterpart", "gun", "fuel_volume"))

    def test_preview_shows_only_fuel_failures(self):
        self.app._refresh_preview()
        self.assertEqual(set(self.app.result_table.rows), {"2"})
        self.assertEqual(self.app.result_table.rows["2"][1], ("red",))
        self.assertEqual(self.app.result_table.rows["2"][0][-1], "金额不符")
        self.assertEqual(set(self.app.success_table.rows), {"0"})
        self.assertEqual(self.app.success_table.rows["0"][0][-2:], ("4", "10"))
        self.app.preview_text.set.assert_called_with("异常 1 / 2 笔（当前显示 1 行）")
        self.app.success_text.set.assert_called_with("匹配成功 1 / 2 笔")

        self.app.red_only.get.return_value = True
        self.app._refresh_preview()
        self.assertEqual(set(self.app.result_table.rows), {"2"})
        self.assertEqual(set(self.app.success_table.rows), {"0"})
        self.app.preview_text.set.assert_called_with("异常 1 / 2 笔（当前显示 1 行）")

    def test_extra_invoice_does_not_change_fuel_totals(self):
        self.app.results.append(dict(kind="invoice", source="发票", counterpart="", status="无对应记录", failure_reason="无对应记录", red=False))
        self.app._refresh_preview()
        self.assertEqual(set(self.app.result_table.rows), {"2"})
        self.app.preview_text.set.assert_called_with("异常 1 / 2 笔（当前显示 1 行）")
        self.app.success_text.set.assert_called_with("匹配成功 1 / 2 笔")

    def test_account_card_failure_uses_yellow_tag(self):
        self.app.results[2]["card_type"] = "记账卡"
        self.app.results[2]["payment_type"] = "现金"
        self.app._refresh_preview()
        self.assertEqual(self.app.result_table.rows["2"][1], ("red",))
        self.app.results[2]["payment_type"] = "记账卡"
        self.app._refresh_preview()
        self.assertEqual(self.app.result_table.rows["2"][1], ("yellow",))
        self.assertEqual(self.app.result_table.rows["2"][0][-2], "记账卡")

    def test_failures_sort_by_transaction_time(self):
        self.app.results[2]["row"] = 552
        self.app.results[2]["fuel_start_time"] = "2026-10-07 19:41:05"
        self.app.results.append(dict(kind="fuel", source="加油", row=101, counterpart="", status="无对应记录",
                                     failure_reason="无对应记录", red=True, fuel_start_time="2026-10-07 20:49:20"))
        self.app.results.append(dict(kind="fuel", source="加油", row=50, counterpart="", status="待确认",
                                     failure_reason="交易时间无效", red=True))
        self.app._refresh_preview()
        self.assertEqual(list(self.app.result_table.rows), ["2", "4", "5"])
        self.app.preview_text.set.assert_called_with("异常 3 / 4 笔（当前显示 3 行）")

    def test_successes_sort_by_transaction_time(self):
        self.app.results[0]["row"] = 5
        self.app.results[0]["fuel_start_time"] = "2026-10-11 11:00:00"
        self.app.results.append(dict(kind="fuel", source="加油", row=9, counterpart="'发票'!A6",
                                     status="已匹配", failure_reason="", red=False,
                                     fuel_start_time="2026-10-11 10:00:00"))
        self.app._refresh_preview()
        self.assertEqual(list(self.app.success_table.rows), ["4", "0"])

    def test_original_row_search_finds_fuel_failure(self):
        self.app.results[2]["row"] = 1698
        self.app.results[3]["row"] = 7632
        self.app.row_search.get.return_value = "1698"
        self.app._refresh_preview()
        self.assertEqual(set(self.app.result_table.rows), {"2"})
        self.app.preview_text.set.assert_called_with("异常 1 / 2 笔（当前显示 1 行）")

    def test_counterpart_cell_opens_result_workbook(self):
        self.app._refresh_preview()
        self.app._open_result_workbook = Mock()
        self.app._open_counterpart(SimpleNamespace(x=3, y=2))
        self.app._open_result_workbook.assert_called_once_with()

        self.app._open_counterpart(SimpleNamespace(x=2, y=2))
        self.app._open_result_workbook.assert_called_once_with()

        self.app._open_counterpart(SimpleNamespace(widget=self.app.success_table, x=3, y=0))
        self.assertEqual(self.app._open_result_workbook.call_count, 2)


if __name__ == "__main__":
    unittest.main()
