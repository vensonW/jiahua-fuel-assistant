"""Read-only native Excel validation; never saves changes to the report."""
from pathlib import Path
import json
import sys

import pythoncom
import win32com.client
from openpyxl.utils import get_column_letter


def main():
    report = Path(sys.argv[1]).resolve()
    qa = report.parent / "qa" / "latest_period"
    qa.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(report), UpdateLinks=0, ReadOnly=True)
        evidence = {"version": excel.Version, "worksheets": workbook.Worksheets.Count, "sheets": []}
        for index, header, end, original_columns in [
            (1, 4, 101, 20), (2, 2, 99, 31)
        ]:
            sheet = workbook.Worksheets(index)
            status_column = original_columns + 8
            columns = {sheet.Cells(header, col).Text: col
                       for col in range(original_columns + 1, sheet.UsedRange.Columns.Count + 1)}
            for name in ("加油开始时间", "加油结束时间", "开票时间", "匹配明细"):
                assert name in columns, (sheet.Name, name)
            # Excel sees long identifiers as actual text and displays all digits.
            number_cell = sheet.Range("V5" if index == 1 else "E3")
            assert number_cell.Text == "26447200000093907128", (number_cell.Text, number_cell.Value2)
            row = header + 1
            for name, expected in (("加油开始时间", "2026-10-11 21:28:03"),
                                   ("加油结束时间", "2026-10-11 21:28:35"),
                                   ("开票时间", "2026-10-11 21:28:40")):
                assert sheet.Cells(row, columns[name]).Text == expected
                assert sheet.Cells(row, columns[name]).Value2 > 40000
            details = sheet.Cells(row, columns["匹配明细"])
            assert details.WrapText and details.RowHeight >= 170
            assert "811031001" in details.Value2 and "26447200000093907128" in details.Value2
            for row_number in range(header + 1, end + 1):
                assert bool(sheet.Cells(row_number, columns["匹配明细"]).Value2) == (sheet.Cells(row_number, status_column).Text == "已匹配")
            if index == 2:
                last_column = sheet.UsedRange.Columns.Count
                assert sheet.Cells(header, last_column).Text == "加油完成时间"
                assert sheet.Cells(row, last_column).Text == "2026-10-11 21:28:35"
                assert all(sheet.Cells(r, last_column).Text for r in range(3, 100))
            data_range = sheet.AutoFilter.Range
            data_range.AutoFilter(Field=status_column, Criteria1=13551615, Operator=8)
            red_visible = [row for row in range(header + 1, end + 1) if not sheet.Rows(row).Hidden]
            assert len(red_visible) == 82, (sheet.Name, len(red_visible))
            sheet.ShowAllData()
            data_range.AutoFilter(Field=status_column, Criteria1="金额不符")
            status_visible = [row for row in range(header + 1, end + 1) if not sheet.Rows(row).Hidden]
            assert red_visible == status_visible
            sheet.ShowAllData()
            detail_letter = get_column_letter(columns["匹配明细"])
            sheet.PageSetup.PrintArea = f"{detail_letter}{header}:{detail_letter}{row}"
            sheet.PageSetup.PrintTitleRows = ""
            sheet.PageSetup.Orientation = 2
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            sheet.ExportAsFixedFormat(0, str(qa / f"native_sheet_{index}.pdf"))
            if index == 2:
                sheet.PageSetup.PrintArea = f"{get_column_letter(columns['加油开始时间'])}{header}:{get_column_letter(columns['开票时间'])}{row}"
                sheet.ExportAsFixedFormat(0, str(qa / "native_three_times.pdf"))
            evidence["sheets"].append(dict(name=sheet.Name, red_visible=len(red_visible),
                                           status_visible=len(status_visible), invoice_text=number_cell.Text))
        (qa / "native_excel.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
        print(json.dumps(evidence, ensure_ascii=False))
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
