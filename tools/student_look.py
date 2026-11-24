# -*- coding: utf-8 -*-
"""Приводит книги Excel к простому «студенческому» виду."""
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CAL = "Calibri"
BLUE = "5B9BD5"       # обычный синий Excel
BLUE2 = "2F5496"      # заголовок как в Word/Excel
YELLOW = "FFFF99"     # ячейки ввода
GREEN = "C6EFCE"      # результат
GRAY = "D9D9D9"
THIN = Border(
    left=Side(style="thin", color="808080"),
    right=Side(style="thin", color="808080"),
    top=Side(style="thin", color="808080"),
    bottom=Side(style="thin", color="808080"),
)
PRO_DARK = {"1F3864", "2E5C9A", "1F4E79", "2E75B6"}
PRO_GOLD = {"FFF2CC", "FFC000"}
PRO_PALE = {"EDF2F9", "D9E2F3", "D6E3F8", "DEEAF6", "F2F6FC"}


def rgb_of(fill):
    if fill is None or fill.fill_type != "solid":
        return None
    fg = fill.fgColor
    if fg is None or fg.rgb is None:
        return None
    return str(fg.rgb)[-6:].upper()


def restyle_ws(ws):
    ws.sheet_view.showGridLines = True
    ws.sheet_properties.tabColor = BLUE
    for row in ws.iter_rows():
        for c in row:
            if c.value is None and (c.fill is None or c.fill.fill_type is None):
                continue
            fnt = c.font
            size = fnt.size or 11
            color = None
            if fnt.color is not None and fnt.color.rgb and str(fnt.color.rgb)[-6:].upper() not in ("000000", "FFFFFF"):
                hexc = str(fnt.color.rgb)[-6:].upper()
                if hexc in PRO_DARK or hexc == "C00000":
                    color = BLUE2 if fnt.bold else "000000"
                else:
                    color = hexc
            # заголовки белым если синяя заливка
            fill_hex = rgb_of(c.fill)
            new_fill = None
            if fill_hex in PRO_DARK:
                new_fill = PatternFill("solid", fgColor=BLUE)
                color = "FFFFFF"
            elif fill_hex in PRO_GOLD:
                new_fill = PatternFill("solid", fgColor=YELLOW)
            elif fill_hex in PRO_PALE:
                new_fill = PatternFill("solid", fgColor="DEEBF7")
            elif fill_hex == "C00000":
                new_fill = PatternFill("solid", fgColor=GREEN)
            col = color or ("FFFFFF" if fill_hex in PRO_DARK else "000000")
            try:
                c.font = Font(name=CAL, size=min(float(size or 11), 16) if (size or 11) > 14 else (size or 11),
                              bold=bool(fnt.bold), italic=bool(fnt.italic), color=col)
            except Exception:
                c.font = Font(name=CAL, size=11, bold=bool(fnt.bold), italic=bool(fnt.italic))
            if new_fill is not None:
                c.fill = new_fill
            if c.border and c.border.left and c.border.left.style:
                c.border = THIN


def restyle_file(path):
    wb = load_workbook(path)
    for ws in wb.worksheets:
        restyle_ws(ws)
    wb.save(path)
    print("restyle", os.path.basename(path))


def main():
    for name in os.listdir(ROOT):
        if name.endswith(".xlsx") and not name.startswith("~$"):
            restyle_file(os.path.join(ROOT, name))


if __name__ == "__main__":
    main()
