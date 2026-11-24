# -*- coding: utf-8 -*-
"""Снимает оформление: обычный Excel, как после ввода в ячейки."""
import glob
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEF_FONT = Font(name="Calibri", size=11, bold=False, italic=False, color="000000")
DEF_AL = Alignment(horizontal="general", vertical="bottom", wrap_text=False)
NO_BORDER = Border()


def restyle_file(path):
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for rng in list(ws.merged_cells.ranges):
            try:
                ws.unmerge_cells(str(rng))
            except Exception:
                pass
        ws.sheet_view.showGridLines = True
        ws.freeze_panes = None
        ws.sheet_properties.tabColor = None
        for col in list(ws.column_dimensions.keys()):
            ws.column_dimensions[col].width = 8.43
        for r in list(ws.row_dimensions.keys()):
            ws.row_dimensions[r].height = None
        max_r, max_c = ws.max_row or 1, min(ws.max_column or 1, 40)
        for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
            for c in row:
                c.font = DEF_FONT
                c.fill = PatternFill()
                c.border = NO_BORDER
                c.alignment = DEF_AL
        try:
            ws.sheet_properties.pageSetUpPr.fitToPage = False
        except Exception:
            pass
    wb.save(path)
    print("plain", os.path.basename(path))


def main():
    for path in sorted(glob.glob(os.path.join(ROOT, "*.xlsx"))):
        restyle_file(path)


if __name__ == "__main__":
    main()
