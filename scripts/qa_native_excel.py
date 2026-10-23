"""Read-only native Excel check of original columns and fuel-only marking."""
from pathlib import Path
from copy import copy
import json
import sys

import pythoncom
import win32com.client
from openpyxl import load_workbook


def main():
    report = Path(sys.argv[1]).resolve()
    root = Path(__file__).resolve().parents[1]
    qa = report.parent / "qa" / "original_marking"
    qa.mkdir(parents=True, exist_ok=True)
    original = load_workbook(root / "中南加华加油和开票数据核对.xlsx")
    result = load_workbook(report)
    try:
        assert original.sheetnames == result.sheetnames
        for before, after in zip(original.worksheets, result.worksheets):
            assert before.max_column == after.max_column
            for row in before:
                for cell in row:
                    saved = after.cell(cell.row, cell.column)
                    assert cell.value == saved.value
                    assert (cell.comment.text if cell.comment else None) == (saved.comment.text if saved.comment else None)
                    if before.title == original.worksheets[1].title:
                        assert copy(cell.fill) == copy(saved.fill)
                        assert copy(cell.font) == copy(saved.font)
                        assert cell.number_format == saved.number_format
            for key, dimension in before.row_dimensions.items():
                assert dimension.height == after.row_dimensions[key].height
            for key, dimension in before.column_dimensions.items():
                assert dimension.width == after.column_dimensions[key].width
        assert result.worksheets[0].max_column == 20
        assert result.worksheets[1].max_column == 31
    finally:
        original.close()
        result.close()
    log_path = report.with_suffix(".log")
    log = log_path.read_text(encoding="utf-8")
    assert "加油开始时间" in log and "26447200000093907128" in log and "匹配明细" in log
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(report), UpdateLinks=0, ReadOnly=True)
        fuel, invoice = workbook.Worksheets(1), workbook.Worksheets(2)
        assert fuel.UsedRange.Columns.Count == 20
        assert invoice.UsedRange.Columns.Count == 31
        assert invoice.Range("E3").Text == "26447200000093907128"
        red = 13551615
        fuel_red = [row for row in range(5, 102) if fuel.Cells(row, 1).Interior.Color == red]
        invoice_red = [row for row in range(3, 100) if invoice.Cells(row, 1).Interior.Color == red]
        assert len(fuel_red) == 82 and len(invoice_red) == 0
        fuel.AutoFilter.Range.AutoFilter(Field=1, Criteria1=red, Operator=8)
        visible = [row for row in range(5, 102) if not fuel.Rows(row).Hidden]
        assert visible == fuel_red
        fuel.ShowAllData()
        for sheet, area, name in [(fuel, "A4:I10", "fuel_marked"), (invoice, "A2:G7", "invoice_original")]:
            sheet.PageSetup.PrintArea = area
            sheet.PageSetup.PrintTitleRows = ""
            sheet.PageSetup.Orientation = 2
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            sheet.ExportAsFixedFormat(0, str(qa / (name + ".pdf")))
        evidence = dict(fuel_columns=20, invoice_columns=31, fuel_red=82, invoice_red=0,
                        fuel_color_filter_visible=82, original_values_and_invoice_format_preserved=True,
                        no_new_columns_or_comments=True, log_path=str(log_path))
        (qa / "native_excel.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
        print(json.dumps(evidence, ensure_ascii=False))
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
