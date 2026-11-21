# -*- coding: utf-8 -*-
"""Excel УСР №2, задача 2.4."""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
import usr2_data as D

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FN = os.path.join(ROOT, "Эконометрика Верич УСР 2.xlsx")

DARK = "1F3864"
MID = "2E5C9A"
PALE = "EDF2F9"
GOLD = "FFF2CC"
ACCENT = "C00000"
thin = Side(style="thin", color="9DB7E0")
med = Side(style="medium", color=DARK)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=Side(style="medium", color=MID))

F_TITLE = Font(name="Times New Roman", size=14, bold=True, color=DARK)
F_SUB = Font(name="Times New Roman", size=10.5, italic=True, color="404040")
F_SECT = Font(name="Times New Roman", size=12, bold=True, color=DARK)
F_HDR = Font(name="Times New Roman", size=10.5, bold=True, color="FFFFFF")
F_B = Font(name="Times New Roman", size=11, bold=True)
F_N = Font(name="Times New Roman", size=11)
F_RES = Font(name="Times New Roman", size=12, bold=True, color=ACCENT)
F_SMALL = Font(name="Times New Roman", size=9, italic=True, color="595959")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def title(ws, row, text, last="H"):
    ws.merge_cells(f"A{row}:{last}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 24


def subtitle(ws, row, text, last="H"):
    ws.merge_cells(f"A{row}:{last}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_SUB
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 18


def section(ws, row, text, last="H"):
    ws.merge_cells(f"A{row}:{last}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_SECT
    c.fill = PatternFill("solid", fgColor=PALE)
    c.border = BOTTOM
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20


def header(ws, row, values, start=1, h=30):
    for k, v in enumerate(values):
        c = ws.cell(row=row, column=start + k, value=v)
        c.font = F_HDR
        c.fill = PatternFill("solid", fgColor=DARK)
        c.alignment = CEN
        c.border = Border(left=Side(style="thin", color="FFFFFF"),
                          right=Side(style="thin", color="FFFFFF"),
                          top=med, bottom=med)
    ws.row_dimensions[row].height = h


def cells(ws, row, values, start=1, fmt=None, alt=False, bold=False):
    for k, v in enumerate(values):
        c = ws.cell(row=row, column=start + k, value=v)
        c.font = F_B if bold else F_N
        c.alignment = CEN
        if alt:
            c.fill = PatternFill("solid", fgColor=PALE)
        c.border = BOX
        if fmt:
            c.number_format = fmt


def result_block(ws, row, label, value, last="F"):
    ws.merge_cells(f"A{row}:D{row}")
    a = ws[f"A{row}"]
    a.value = label
    a.font = F_RES
    a.fill = PatternFill("solid", fgColor=GOLD)
    a.border = BOX
    a.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.merge_cells(f"E{row}:{last}{row}")
    e = ws[f"E{row}"]
    e.value = value
    e.font = F_RES
    e.fill = PatternFill("solid", fgColor=GOLD)
    e.border = BOX
    e.alignment = CEN
    ws.row_dimensions[row].height = 28


def widths(ws, m):
    for k, v in m.items():
        ws.column_dimensions[k].width = v


def note(ws, row, text, last="H", h=36):
    ws.merge_cells(f"A{row}:{last}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_SMALL
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[row].height = h


def sheet_task(wb):
    ws = wb.active
    ws.title = "Задание"
    widths(ws, {"A": 22, "B": 12, "C": 12, "D": 12, "E": 16, "F": 14, "G": 14, "H": 14})
    title(ws, 1, "Задача 2.4. Многоцелевая оптимизация производственной программы")
    subtitle(ws, 2, "Три вида продукции, три вида ресурсов. x3 ≥ 30. Критерии: прибыль (max), оптовая цена (max), трудоёмкость (min).")
    section(ws, 4, "1. Нормы расхода ресурсов и наличие")
    header(ws, 5, ["Ресурс", "П1", "П2", "П3", "Наличие"])
    names = ["Ресурс 1", "Ресурс 2", "Ресурс 3"]
    for i, n in enumerate(names):
        cells(ws, 6 + i, [n] + D.A[i] + [D.B[i]], alt=i % 2)
    section(ws, 10, "2. Показатели единицы продукции")
    header(ws, 11, ["Показатель", "П1", "П2", "П3", "Направление"])
    cells(ws, 12, ["Прибыль, ден. ед."] + D.C_PR + ["max"], alt=True)
    cells(ws, 13, ["Оптовая цена, ден. ед."] + D.C_PC + ["max"])
    cells(ws, 14, ["Трудоёмкость, чел.-ч"] + D.C_LB + ["min"], alt=True)
    section(ws, 16, "3. Модель (общие ограничения)")
    ws.merge_cells("A17:H20")
    ws["A17"].value = (
        "2x1 + 4x2 + 4x3 ≤ 400\n"
        "3x1 + 2x2 + 2x3 ≤ 300\n"
        "4x1 + 5x2 + 3x3 ≤ 500\n"
        "x1 ≥ 0,  x2 ≥ 0,  x3 ≥ 30"
    )
    ws["A17"].font = F_N
    ws["A17"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[17].height = 20
    ws.row_dimensions[18].height = 20
    ws.row_dimensions[19].height = 20
    ws.row_dimensions[20].height = 20
    section(ws, 22, "4. Требуется")
    note(ws, 23,
         "а) метод последовательных уступок: прибыль (уступка 580) → оптовая цена (уступка 1090) → трудоёмкость;\n"
         "б) метод ведущего критерия: max прибыль, оптовая цена ≥ 2400, трудоёмкость ≤ 1600;\n"
         "в) метод равных и наименьших относительных отклонений при x1 ≥ 20;\n"
         "г) метод минимакса при x3 ≥ 40.", last="H", h=70)
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def put_plan(ws, row, x, tag="B"):
    """x1,x2,x3 в B,C,D row; формулы критериев и ресурсов."""
    header(ws, row, ["Переменная", "x1 (П1)", "x2 (П2)", "x3 (П3)"])
    cells(ws, row + 1, ["План, ед.", x[0], x[1], x[2]], fmt="0.000")
    header(ws, row + 3, ["Показатель", "Формула", "Значение", "Ограничение"])
    # profit
    cells(ws, row + 4, ["Прибыль f1", "22x1+12x2+14x3",
                        f"=22*{tag}{row+1}+12*C{row+1}+14*D{row+1}", "max"], fmt="0.00")
    cells(ws, row + 5, ["Оптовая цена f2", "24x1+16x2+30x3",
                        f"=24*{tag}{row+1}+16*C{row+1}+30*D{row+1}", "max"], fmt="0.00", alt=True)
    cells(ws, row + 6, ["Трудоёмкость f3", "20x1+10x2+15x3",
                        f"=20*{tag}{row+1}+10*C{row+1}+15*D{row+1}", "min"], fmt="0.00")
    header(ws, row + 8, ["Ресурс", "Расход", "Наличие", "Остаток"])
    cells(ws, row + 9, ["Ресурс 1",
                       f"=2*{tag}{row+1}+4*C{row+1}+4*D{row+1}", 400,
                       f"=C{row+9}-B{row+9}"], fmt="0.00")
    cells(ws, row + 10, ["Ресурс 2",
                        f"=3*{tag}{row+1}+2*C{row+1}+2*D{row+1}", 300,
                        f"=C{row+10}-B{row+10}"], fmt="0.00", alt=True)
    cells(ws, row + 11, ["Ресурс 3",
                        f"=4*{tag}{row+1}+5*C{row+1}+3*D{row+1}", 500,
                        f"=C{row+11}-B{row+11}"], fmt="0.00")


def sheet_a(wb):
    ws = wb.create_sheet("2.4а")
    widths(ws, {"A": 28, "B": 16, "C": 16, "D": 16, "E": 16, "F": 14})
    title(ws, 1, "2.4 (а). Метод последовательных уступок")
    subtitle(ws, 2, "Важность: прибыль → оптовая цена → трудоёмкость. Уступки Δf1 = 580, Δf2 = 1090.")
    section(ws, 4, "Шаг 1. Максимум прибыли (ведущий критерий)")
    put_plan(ws, 5, D.A1_X)
    result_block(ws, 18, "f1* (максимум прибыли), ден. ед.", "=C9")
    ws["E18"].number_format = "0.00"
    section(ws, 20, "Шаг 2. Максимум оптовой цены при f1 ≥ f1* − 580 = 1600")
    put_plan(ws, 21, D.A2_X)
    result_block(ws, 34, "f2* при уступке по прибыли, ден. ед.", "=C26")
    ws["E34"].number_format = "0.00"
    section(ws, 36, "Шаг 3. Минимум трудоёмкости при f1 ≥ 1600 и f2 ≥ 3450 − 1090 = 2360")
    put_plan(ws, 37, D.A3_X)
    result_block(ws, 50, "Компромисс: прибыль f1, ден. ед.", "=C41")
    ws["E50"].number_format = "0.00"
    result_block(ws, 51, "Оптовая цена f2 / трудоёмкость f3", "=C42")
    ws["E51"].number_format = "0.00"
    note(ws, 53, "Компромисс: x1 = 21,25; x2 = 59,375; x3 = 30. Прибыль 1600, оптовая цена 2360, трудоёмкость 1468,75.", h=32)
    add_chart(ws, 5, 21, 37)


def add_chart(ws, r1, r2, r3):
    # skip if too crowded; put a small comparison table instead
    pass


def sheet_b(wb):
    ws = wb.create_sheet("2.4б")
    widths(ws, {"A": 28, "B": 16, "C": 16, "D": 16, "E": 16, "F": 14})
    title(ws, 1, "2.4 (б). Метод ведущего критерия")
    subtitle(ws, 2, "max f1 при f2 ≥ 2400 и f3 ≤ 1600, x3 ≥ 30.")
    section(ws, 4, "Оптимальный план")
    put_plan(ws, 5, D.B_X)
    result_block(ws, 18, "Максимум прибыли f1, ден. ед.", "=C9")
    ws["E18"].value = "=C9"
    result_block(ws, 19, "Оптовая цена f2 (≥ 2400)", "=C10")
    ws["E19"].value = "=C10"
    result_block(ws, 20, "Трудоёмкость f3 (≤ 1600)", "=C11")
    ws["E20"].value = "=C11"
    note(ws, 22, "План x = (30; 55; 30). Прибыль 1740 ден. ед., оптовая цена 2500, трудоёмкость 1600 чел.-ч.", h=32)


def sheet_c(wb):
    ws = wb.create_sheet("2.4в")
    widths(ws, {"A": 32, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16})
    title(ws, 1, "2.4 (в). Равные и наименьшие относительные отклонения")
    subtitle(ws, 2, "Дополнительно x1 ≥ 20. Сначала ищутся частные оптимумы f_k*, затем план с равными |f_k − f_k*| / |f_k*|.")
    section(ws, 4, "Частные оптимумы при x1 ≥ 20, x3 ≥ 30")
    header(ws, 5, ["Критерий", "x1", "x2", "x3", "f_k*"])
    cells(ws, 6, ["max f1 (прибыль)", 80, 0, 30, 2180], fmt="0.00")
    cells(ws, 7, ["max f2 (цена)", 50, 0, 75, 3450], fmt="0.00", alt=True)
    cells(ws, 8, ["min f3 (труд)", 20, 0, 30, 850], fmt="0.00")
    section(ws, 10, "Компромиссный план")
    put_plan(ws, 11, D.C_X)
    header(ws, 24, ["Критерий", "f_k", "f_k*", "Отклонение", "δ_k"])
    cells(ws, 25, ["f1", "=C15", 2180, "=C25-B25", "=(C25-B25)/C25"], fmt="0.0000")
    cells(ws, 26, ["f2", "=C16", 3450, "=C26-B26", "=(C26-B26)/C26"], fmt="0.0000", alt=True)
    cells(ws, 27, ["f3", "=C17", 850, "=B27-C27", "=(B27-C27)/C27"], fmt="0.0000")
    for r in (25, 26, 27):
        ws[f"B{r}"].number_format = "0.00"
        ws[f"C{r}"].number_format = "0.00"
        ws[f"D{r}"].number_format = "0.00"
        ws[f"E{r}"].number_format = "0.00%"
    result_block(ws, 29, "Общее относительное отклонение δ", "=E25")
    ws["E29"].value = "=E25"
    ws["E29"].number_format = "0.00%"
    note(ws, 31, "δ1 = δ2 = δ3 ≈ 42,14 %. План x ≈ (20,00; 25,10; 37,15).", h=28)


def sheet_d(wb):
    ws = wb.create_sheet("2.4г")
    widths(ws, {"A": 32, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16})
    title(ws, 1, "2.4 (г). Метод минимакса")
    subtitle(ws, 2, "x3 ≥ 40. Минимизируется наибольшее относительное отклонение y = max δ_k.")
    section(ws, 4, "Частные оптимумы при x3 ≥ 40")
    header(ws, 5, ["Критерий", "x1", "x2", "x3", "f_k*"])
    cells(ws, 6, ["max f1", 73.333333, 0, 40, 2173.333333], fmt="0.00")
    cells(ws, 7, ["max f2", 50, 0, 75, 3450], fmt="0.00", alt=True)
    cells(ws, 8, ["min f3", 0, 0, 40, 600], fmt="0.00")
    section(ws, 10, "Компромиссный план (min y)")
    put_plan(ws, 11, D.D_X)
    header(ws, 24, ["Критерий", "f_k", "f_k*", "δ_k"])
    cells(ws, 25, ["f1", "=C15", 2173.333333, "=(C26-B25)/C26"], fmt="0.0000")
    # C26 is wrong. Use column C as f*
    ws["C25"] = 2173.333333
    ws["C26"] = 3450
    ws["C27"] = 600
    ws["B25"] = "=C15"
    ws["B26"] = "=C16"
    ws["B27"] = "=C17"
    ws["A25"] = "f1"
    ws["A26"] = "f2"
    ws["A27"] = "f3"
    ws["D25"] = "=(C25-B25)/C25"
    ws["D26"] = "=(C26-B26)/C26"
    ws["D27"] = "=(B27-C27)/C27"
    for r in range(25, 28):
        for col in "BCD":
            ws[f"{col}{r}"].number_format = "0.0000" if col == "D" else "0.00"
            ws[f"{col}{r}"].font = F_N
            ws[f"{col}{r}"].alignment = CEN
            ws[f"{col}{r}"].border = BOX
        ws[f"A{r}"].font = F_N
        ws[f"A{r}"].alignment = CEN
        ws[f"A{r}"].border = BOX
        if r % 2 == 0:
            for col in "ABCD":
                ws[f"{col}{r}"].fill = PatternFill("solid", fgColor=PALE)
    result_block(ws, 29, "Минимакс y* = max δ_k", "=MAX(D25:D27)")
    ws["E29"].value = "=MAX(D25:D27)"
    ws["E29"].number_format = "0.00%"
    note(ws, 31, "y* ≈ 55,76 %. План x ≈ (0; 33,46; 40). Отклонения: δ1 = δ3 = y*, δ2 меньше.", h=32)


def sheet_compare(wb):
    ws = wb.create_sheet("Сводка")
    widths(ws, {"A": 28, "B": 12, "C": 12, "D": 12, "E": 14, "F": 14, "G": 14})
    title(ws, 1, "Сводка решений задачи 2.4", last="G")
    header(ws, 3, ["Метод", "x1", "x2", "x3", "f1 прибыль", "f2 цена", "f3 труд"])
    cells(ws, 4, ["а) уступки", 21.25, 59.375, 30, 1600, 2360, 1468.75], fmt="0.00")
    cells(ws, 5, ["б) ведущий", 30, 55, 30, 1740, 2500, 1600], fmt="0.00", alt=True)
    cells(ws, 6, ["в) равные отклонения", 20, 25.098855, 37.148803, 1261.27, 1996.05, 1208.22], fmt="0.00")
    cells(ws, 7, ["г) минимакс", 0, 33.456221, 40, 961.47, 1735.30, 934.56], fmt="0.00", alt=True)
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = None
    chart.y_axis.title = None
    chart.style = 10
    data = Reference(ws, min_col=5, min_row=3, max_col=7, max_row=7)
    cats = Reference(ws, min_col=1, min_row=4, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 18
    chart.height = 8
    ws.add_chart(chart, "A10")
    note(ws, 28, "Значения x и f_k согласованы с листами 2.4а–2.4г. Диаграмма — сравнение критериев по методам.", last="G", h=28)


def build():
    wb = openpyxl.Workbook()
    sheet_task(wb)
    sheet_a(wb)
    sheet_b(wb)
    sheet_c(wb)
    sheet_d(wb)
    sheet_compare(wb)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.leftMargin = 0.4
        ws.page_setup.rightMargin = 0.4
        ws.print_options.horizontalCentered = True
    wb.save(FN)
    print("сохранено:", FN)


if __name__ == "__main__":
    build()
