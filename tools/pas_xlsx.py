# -*- coding: utf-8 -*-
"""Оформление Excel УСР №1 (Пась А.М.)."""
import os
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.formula import ArrayFormula

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FN = os.path.join(ROOT, "Пась А.М. 24ИТ-3 УСР Свитина.xlsx")

DARK, MID, PALE, GOLD, ACCENT, LIGHT = "1F3864", "2E5C9A", "EDF2F9", "FFF2CC", "C00000", "D6E3F8"
thin = Side(style="thin", color="9DB7E0")
med = Side(style="medium", color=DARK)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=Side(style="medium", color=MID))
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
F_TITLE = Font(name="Times New Roman", size=14, bold=True, color=DARK)
F_SUB = Font(name="Times New Roman", size=10.5, italic=True, color="404040")
F_SECT = Font(name="Times New Roman", size=12, bold=True, color=DARK)
F_HDR = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
F_B = Font(name="Times New Roman", size=11, bold=True)
F_N = Font(name="Times New Roman", size=11)
F_RES = Font(name="Times New Roman", size=12, bold=True, color=ACCENT)
F_S = Font(name="Times New Roman", size=9, italic=True, color="595959")


def title(ws, r, t, last="L"):
    ws.merge_cells(f"A{r}:{last}{r}")
    c = ws[f"A{r}"]; c.value = t; c.font = F_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 24


def subtitle(ws, r, t, last="L"):
    ws.merge_cells(f"A{r}:{last}{r}")
    c = ws[f"A{r}"]; c.value = t; c.font = F_SUB
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 18


def section(ws, r, t, last="L"):
    ws.merge_cells(f"A{r}:{last}{r}")
    c = ws[f"A{r}"]; c.value = t; c.font = F_SECT
    c.fill = PatternFill("solid", fgColor=PALE); c.border = BOTTOM
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 20


def header(ws, r, vals, start=1, h=28):
    for k, v in enumerate(vals):
        c = ws.cell(row=r, column=start + k, value=v)
        c.font = F_HDR; c.fill = PatternFill("solid", fgColor=DARK)
        c.alignment = CEN
        c.border = Border(left=Side(style="thin", color="FFFFFF"),
                          right=Side(style="thin", color="FFFFFF"), top=med, bottom=med)
    ws.row_dimensions[r].height = h


def put(ws, r, c, v, bold=False, fmt=None, fill=None, font=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or (F_B if bold else F_N)
    cell.alignment = CEN
    cell.border = BOX
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    return cell


def widths(ws, m):
    for k, v in m.items():
        ws.column_dimensions[k].width = v


def result(ws, r, label, value, last="F"):
    ws.merge_cells(f"A{r}:D{r}")
    a = ws[f"A{r}"]; a.value = label; a.font = F_RES
    a.fill = PatternFill("solid", fgColor=GOLD); a.border = BOX
    a.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.merge_cells(f"E{r}:{last}{r}")
    e = ws[f"E{r}"]; e.value = value; e.font = F_RES
    e.fill = PatternFill("solid", fgColor=GOLD); e.border = BOX; e.alignment = CEN
    ws.row_dimensions[r].height = 26


def note(ws, r, t, last="L", h=36):
    ws.merge_cells(f"A{r}:{last}{r}")
    c = ws[f"A{r}"]; c.value = t; c.font = F_S
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[r].height = h


def page(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.leftMargin = 0.4
    ws.page_setup.rightMargin = 0.4
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False


def sheet_task(wb):
    ws = wb.active
    ws.title = "Задание"
    widths(ws, {"A": 28, "B": 14, "C": 14, "D": 14, "E": 16, "F": 18, "G": 16, "H": 14})
    title(ws, 1, "УСР №1. Модель межотраслевого баланса  —  задача 1.6")
    subtitle(ws, 2, "Пась А. М., группа 24ИТ-3. Преподаватель: Свитина.")
    section(ws, 4, "Исходные данные (межотраслевые потоки за отчётный период, млн руб.)")
    header(ws, 5, ["Отрасль-производитель", "1", "2", "3", "Yотч", "Xотч"])
    put(ws, 6, 1, "1"); put(ws, 6, 2, 30); put(ws, 6, 3, 10); put(ws, 6, 4, 15)
    put(ws, 6, 5, 90, fill=GOLD); put(ws, 6, 6, "=B6+C6+D6+E6", bold=True)
    put(ws, 7, 1, "2"); put(ws, 7, 2, 35); put(ws, 7, 3, 50); put(ws, 7, 4, 20)
    put(ws, 7, 5, 25, fill=GOLD); put(ws, 7, 6, "=B7+C7+D7+E7", bold=True)
    put(ws, 8, 1, "3"); put(ws, 8, 2, 15); put(ws, 8, 3, 25); put(ws, 8, 4, 30)
    put(ws, 8, 5, 60, fill=GOLD); put(ws, 8, 6, "=B8+C8+D8+E8", bold=True)
    for r in range(6, 9):
        ws[f"F{r}"].number_format = "0.00"
    section(ws, 10, "Параметры задания")
    header(ws, 11, ["Показатель", "Значение", "Смысл"])
    ws.merge_cells("C12:F12"); ws.merge_cells("C13:F13")
    ws.merge_cells("C14:F14"); ws.merge_cells("C15:F15")
    ws.merge_cells("C16:F16"); ws.merge_cells("C17:F17")
    put(ws, 12, 1, "Доля з/п в ВДС"); put(ws, 12, 2, 0.3, fmt="0%"); put(ws, 12, 3, "30 % валовой добавленной стоимости")
    put(ws, 13, 1, "Yпл отрасли 1"); put(ws, 13, 2, 120); put(ws, 13, 3, "плановое конечное использование")
    put(ws, 14, 1, "Yпл отрасли 2"); put(ws, 14, 2, 30); put(ws, 14, 3, "")
    put(ws, 15, 1, "Yпл отрасли 3"); put(ws, 15, 2, 60); put(ws, 15, 3, "")
    put(ws, 16, 1, "n1, r, эластичность з/п"); put(ws, 16, 2, "1; 2; 0,7")
    put(ws, 16, 3, "цена отрасли 1 ×2; рост з/п отстаёт на 30 %")
    put(ws, 17, 1, "n2, прирост з/п"); put(ws, 17, 2, "2; +30%")
    put(ws, 17, 3, "з/п во 2-й отрасли растёт на 30 %, в остальных неизменна")
    for r in range(12, 18):
        ws[f"C{r}"].alignment = LEFT
        ws[f"C{r}"].font = F_N
    section(ws, 19, "Требуется")
    note(ws, 20,
         "1) схема МОБ за отчётный период (промежуточное потребление и затраты, з/п, прочие элементы ВДС);\n"
         "2) матрица коэффициентов прямых затрат A;\n"
         "3) валовой выпуск Xпл и схема МОБ в плановом периоде при заданном Yпл;\n"
         "4) индексы цен, если цена продукции 1-й отрасли увеличивается в 2 раза (рост з/п отстаёт на 30 %);\n"
         "5) индексы цен, если з/п во 2-й отрасли увеличивается на 30 %, в остальных отраслях неизменна.",
         last="G", h=88)
    page(ws)


def mob_block(ws, r0, x_cells, y_cells, gold_y=False):
    """Заголовок таблицы МОБ начиная с r0 (header row). Возвращает строки данных 1,2,3."""
    header(ws, r0, ["Производитель \\ потребитель", "1", "2", "3",
                    "Пром. потребление", "Конечное использование Y", "Валовой выпуск X"])
    rows = []
    for i, lab in enumerate((1, 2, 3)):
        r = r0 + 1 + i
        put(ws, r, 1, lab, bold=True, fill=PALE if i % 2 else None)
        for j, xc in enumerate(x_cells[i]):
            put(ws, r, 2 + j, xc, fmt="0.00", fill=PALE if i % 2 else None)
        put(ws, r, 5, f"=B{r}+C{r}+D{r}", fmt="0.00")
        put(ws, r, 6, y_cells[i], fmt="0.00", fill=GOLD if gold_y else None)
        put(ws, r, 7, f"=E{r}+F{r}", fmt="0.00", bold=True)
        rows.append(r)
    r = r0 + 4
    put(ws, r, 1, "Промежуточные затраты", bold=True, fill=LIGHT)
    for col, L in enumerate("BCD", 2):
        put(ws, r, col, f"={L}{rows[0]}+{L}{rows[1]}+{L}{rows[2]}", fmt="0.00", fill=LIGHT, bold=True)
    put(ws, r, 5, f"=B{r}+C{r}+D{r}", fmt="0.00", fill=LIGHT)
    put(ws, r, 6, "–", fill=LIGHT); put(ws, r, 7, "–", fill=LIGHT)
    r_w = r0 + 5
    put(ws, r_w, 1, "Заработная плата (30 % ВДС)", fill=PALE)
    for col, L in enumerate("BCD", 2):
        put(ws, r_w, col, f"=0.3*{L}{r0+7}", fmt="0.00", fill=PALE)
    put(ws, r_w, 5, "–", fill=PALE); put(ws, r_w, 6, "–", fill=PALE); put(ws, r_w, 7, "–", fill=PALE)
    r_o = r0 + 6
    put(ws, r_o, 1, "Прочие элементы ВДС")
    for col, L in enumerate("BCD", 2):
        put(ws, r_o, col, f"={L}{r0+7}-{L}{r_w}", fmt="0.00")
    put(ws, r_o, 5, "–"); put(ws, r_o, 6, "–"); put(ws, r_o, 7, "–")
    r_z = r0 + 7
    put(ws, r_z, 1, "ВДС Z", bold=True, fill=GOLD)
    for col, L in enumerate("BCD", 2):
        put(ws, r_z, col, f"={L}{r0+8}-{L}{r}", fmt="0.00", fill=GOLD, bold=True)
    put(ws, r_z, 5, f"=B{r_z}+C{r_z}+D{r_z}", fmt="0.00", fill=GOLD, bold=True)
    put(ws, r_z, 6, "–", fill=GOLD)
    put(ws, r_z, 7, f"=E{r_z}", fmt="0.00", fill=GOLD, bold=True)
    r_x = r0 + 8
    put(ws, r_x, 1, "Валовые затраты X", bold=True, fill=LIGHT)
    for i, rr in enumerate(rows):
        put(ws, r_x, 2 + i, f"=G{rr}", fmt="0.00", fill=LIGHT, bold=True)
    put(ws, r_x, 5, f"=B{r_x}+C{r_x}+D{r_x}", fmt="0.00", fill=LIGHT)
    put(ws, r_x, 6, "–", fill=LIGHT); put(ws, r_x, 7, "–", fill=LIGHT)
    return rows, r, r_w, r_o, r_z, r_x


def sheet_mob(wb):
    ws = wb.create_sheet("1.6 МОБ")
    widths(ws, {"A": 32, "B": 14, "C": 14, "D": 14, "E": 20, "F": 24, "G": 18,
                "H": 3, "I": 12, "J": 12, "K": 12, "L": 12})
    title(ws, 1, "Задача 1.6. Схема МОБ, матрица A и плановый валовой выпуск")
    subtitle(ws, 2, "Все производные величины — формулами. Жёлтым выделены исходные Y и доля з/п.")
    section(ws, 4, "1. Схема МОБ за отчётный период")
    rows, r_int, r_w, r_o, r_z, r_x = mob_block(
        ws, 5,
        [[30, 10, 15], [35, 50, 20], [15, 25, 30]],
        [90, 25, 60], gold_y=True)

    section(ws, 16, "2. Матрица коэффициентов прямых затрат  a_ij = x_ij / X_j", last="K")
    header(ws, 17, ["A", "1", "2", "3"], start=1)
    # a_ij = x_ij / X_j  -> B6/B13 etc. data rows 6,7,8; X row 13
    for i, rr in enumerate((6, 7, 8)):
        put(ws, 18 + i, 1, i + 1, bold=True, fill=PALE if i % 2 else None)
        put(ws, 18 + i, 2, f"=B{rr}/$B$13", fmt="0.0000", fill=PALE if i % 2 else None)
        put(ws, 18 + i, 3, f"=C{rr}/$C$13", fmt="0.0000", fill=PALE if i % 2 else None)
        put(ws, 18 + i, 4, f"=D{rr}/$D$13", fmt="0.0000", fill=PALE if i % 2 else None)

    header(ws, 17, ["E − A", "1", "2", "3"], start=6)
    put(ws, 18, 6, 1, bold=True)
    put(ws, 18, 7, "=1-B18", fmt="0.0000"); put(ws, 18, 8, "=-C18", fmt="0.0000"); put(ws, 18, 9, "=-D18", fmt="0.0000")
    put(ws, 19, 6, 2, bold=True, fill=PALE)
    put(ws, 19, 7, "=-B19", fmt="0.0000", fill=PALE); put(ws, 19, 8, "=1-C19", fmt="0.0000", fill=PALE)
    put(ws, 19, 9, "=-D19", fmt="0.0000", fill=PALE)
    put(ws, 20, 6, 3, bold=True)
    put(ws, 20, 7, "=-B20", fmt="0.0000"); put(ws, 20, 8, "=-C20", fmt="0.0000"); put(ws, 20, 9, "=1-D20", fmt="0.0000")

    section(ws, 22, "3. Матрица полных затрат B = (E − A)⁻¹  и плановый выпуск Xпл = B · Yпл", last="K")
    header(ws, 23, ["B", "1", "2", "3", "Yпл", "Xпл = B·Yпл"])
    # MINVERSE as array
    ws["B24"] = ArrayFormula("B24:D26", "=MINVERSE(G18:I20)")
    for r in range(24, 27):
        for c in range(2, 5):
            ws.cell(row=r, column=c).number_format = "0.0000"
            ws.cell(row=r, column=c).font = F_N
            ws.cell(row=r, column=c).alignment = CEN
            ws.cell(row=r, column=c).border = BOX
        put(ws, r, 1, r - 23, bold=True, fill=PALE if (r - 24) % 2 else None)
    put(ws, 24, 5, "='Задание'!B13", fmt="0.00", fill=GOLD)
    put(ws, 25, 5, "='Задание'!B14", fmt="0.00", fill=GOLD)
    put(ws, 26, 5, "='Задание'!B15", fmt="0.00", fill=GOLD)
    put(ws, 24, 6, "=B24*$E$24+C24*$E$25+D24*$E$26", fmt="0.00", bold=True, fill=GOLD)
    put(ws, 25, 6, "=B25*$E$24+C25*$E$25+D25*$E$26", fmt="0.00", bold=True, fill=GOLD)
    put(ws, 26, 6, "=B26*$E$24+C26*$E$25+D26*$E$26", fmt="0.00", bold=True, fill=GOLD)

    section(ws, 28, "4. Схема МОБ в плановом периоде")
    # x_ij = a_ij * Xpl_j
    mob_block(
        ws, 29,
        [[f"=B18*$B$37", f"=C18*$C$37", f"=D18*$D$37"],
         [f"=B19*$B$37", f"=C19*$C$37", f"=D19*$D$37"],
         [f"=B20*$B$37", f"=C20*$C$37", f"=D20*$D$37"]],
        ["=E24", "=E25", "=E26"])
    # Xpl row of plan table is row 37 (29+8)
    # Need G30:G32 = Xpl. Currently G = E+F. Override G from F24:F26
    put(ws, 30, 7, "=F24", fmt="0.00", bold=True)
    put(ws, 31, 7, "=F25", fmt="0.00", bold=True)
    put(ws, 32, 7, "=F26", fmt="0.00", bold=True)

    note(ws, 40,
         "Проверка балансов: сумма валового выпуска равна сумме валовых затрат; сумма ВДС равна сумме конечного использования.",
         last="G", h=28)
    page(ws)


def sheet_prices(wb):
    ws = wb.create_sheet("Цены")
    widths(ws, {"A": 36, "B": 14, "C": 14, "D": 14, "E": 16, "F": 16, "G": 16, "H": 14})
    title(ws, 1, "Задача 1.6. Прогнозирование цен на основе МОБ")
    subtitle(ws, 2, "П. 4: цена 1-й отрасли ×2, эластичность з/п = 0,7.  П. 5: з/п 2-й отрасли +30 %.")

    section(ws, 4, "Отчётные абсолютные величины (из листа «1.6 МОБ»)")
    header(ws, 5, ["", "Отрасль 1", "Отрасль 2", "Отрасль 3"])
    put(ws, 6, 1, "Валовой выпуск X")
    put(ws, 6, 2, "='1.6 МОБ'!B13", fmt="0.00")
    put(ws, 6, 3, "='1.6 МОБ'!C13", fmt="0.00")
    put(ws, 6, 4, "='1.6 МОБ'!D13", fmt="0.00")
    put(ws, 7, 1, "Заработная плата")
    put(ws, 7, 2, "='1.6 МОБ'!B10", fmt="0.00")
    put(ws, 7, 3, "='1.6 МОБ'!C10", fmt="0.00")
    put(ws, 7, 4, "='1.6 МОБ'!D10", fmt="0.00")
    put(ws, 8, 1, "Прочие элементы ВДС")
    put(ws, 8, 2, "='1.6 МОБ'!B11", fmt="0.00")
    put(ws, 8, 3, "='1.6 МОБ'!C11", fmt="0.00")
    put(ws, 8, 4, "='1.6 МОБ'!D11", fmt="0.00")
    put(ws, 9, 1, "Поток из 1-й отрасли x_1j")
    put(ws, 9, 2, "='1.6 МОБ'!B6", fmt="0.00")
    put(ws, 9, 3, "='1.6 МОБ'!C6", fmt="0.00")
    put(ws, 9, 4, "='1.6 МОБ'!D6", fmt="0.00")
    put(ws, 10, 1, "Поток из 2-й отрасли x_2j")
    put(ws, 10, 2, "='1.6 МОБ'!B7", fmt="0.00")
    put(ws, 10, 3, "='1.6 МОБ'!C7", fmt="0.00")
    put(ws, 10, 4, "='1.6 МОБ'!D7", fmt="0.00")
    put(ws, 11, 1, "Поток из 3-й отрасли x_3j")
    put(ws, 11, 2, "='1.6 МОБ'!B8", fmt="0.00")
    put(ws, 11, 3, "='1.6 МОБ'!C8", fmt="0.00")
    put(ws, 11, 4, "='1.6 МОБ'!D8", fmt="0.00")

    section(ws, 13, "4. Цена продукции 1-й отрасли увеличивается в r = 2 раза; k_w = 0,7")
    note(ws, 14,
         "Баланс стоимости отрасли j:  p_j X_j = Σ_i p_i x_ij + з/п_j·k_w·p_j + проч_j·p_j. "
         "p1 = 2 задано; p2 и p3 находятся из уравнений для отраслей 2 и 3.",
         last="G", h=36)
    header(ws, 16, ["Индекс", "Отрасль 1", "Отрасль 2", "Отрасль 3"])
    put(ws, 17, 1, "p (индекс цен)", bold=True, fill=GOLD)
    put(ws, 17, 2, 2, fmt="0.0000", fill=GOLD, bold=True)
    # p2, p3 solved — keep as formulas via rearrangement:
    # p2*(X2 - x22 - wage2*kw - other2) - p3*x32 = p1*x12
    # -p2*x23 + p3*(X3 - x33 - wage3*kw - other3) = p1*x13
    put(ws, 18, 1, "X_j − x_jj − з/п_j·k_w − проч_j")
    put(ws, 18, 2, "=B6-B10-B7*0.7-B8", fmt="0.00")
    put(ws, 18, 3, "=C6-C10-C7*0.7-C8", fmt="0.00")
    put(ws, 18, 4, "=D6-D11-D7*0.7-D8", fmt="0.00")
    # Wait x_jj: industry 2 diagonal is C10 (x22), industry 3 D11 (x33)
    # I mixed rows. x11=B9, x22=C10, x33=D11
    # Let me set p2 p3 as computed values with explanation, plus SUMPRODUCT check.

    # Use closed form from numpy we trust:
    # p2=1.4307987 p3=1.43490756
    # Better Cramer's formulas:
    # A11 = X2 - x22 - w2*kw - o2 = C6-C10-C7*0.7-C8
    # A12 = -x32 = -C11
    # A21 = -x23 = -D10
    # A22 = X3 - x33 - w3*kw - o3 = D6-D11-D7*0.7-D8
    # b1 = p1*x12 = B17*C9
    # b2 = p1*x13 = B17*D9
    put(ws, 18, 1, "Коэф. при p2 (отрасль 2)")
    put(ws, 18, 2, "=C6-C10-C7*0.7-C8", fmt="0.00")  # only B needed
    # cleanup: rewrite 18-22 simply
    for col in range(2, 5):
        ws.cell(row=18, column=col).value = None
        ws.cell(row=18, column=col).border = BOX
    put(ws, 18, 2, "=C6-C10-C7*0.7-C8", fmt="0.000")
    put(ws, 18, 3, "–")
    put(ws, 18, 4, "–")
    put(ws, 19, 1, "Коэф. при p3 (отрасль 2)  = −x_32")
    put(ws, 19, 2, "=-C11", fmt="0.000")
    put(ws, 19, 3, "–"); put(ws, 19, 4, "–")
    put(ws, 20, 1, "Коэф. при p2 (отрасль 3)  = −x_23")
    put(ws, 20, 2, "=-D10", fmt="0.000")
    put(ws, 20, 3, "–"); put(ws, 20, 4, "–")
    put(ws, 21, 1, "Коэф. при p3 (отрасль 3)")
    put(ws, 21, 2, "=D6-D11-D7*0.7-D8", fmt="0.000")
    put(ws, 21, 3, "–"); put(ws, 21, 4, "–")
    put(ws, 22, 1, "Правая часть отрасли 2  p1·x_12")
    put(ws, 22, 2, "=$B$17*C9", fmt="0.000")
    put(ws, 22, 3, "–"); put(ws, 22, 4, "–")
    put(ws, 23, 1, "Правая часть отрасли 3  p1·x_13")
    put(ws, 23, 2, "=$B$17*D9", fmt="0.000")
    put(ws, 23, 3, "–"); put(ws, 23, 4, "–")
    put(ws, 17, 3, "=(B22*B21-B23*B19)/(B18*B21-B20*B19)", fmt="0.0000", fill=GOLD, bold=True)
    put(ws, 17, 4, "=(B18*B23-B20*B22)/(B18*B21-B20*B19)", fmt="0.0000", fill=GOLD, bold=True)

    result(ws, 25, "Индексы цен (п. 4): p1 ; p2 ; p3", '=TEXT(B17,"0.000")&" ; "&TEXT(C17,"0.000")&" ; "&TEXT(D17,"0.000")')

    section(ws, 27, "5. Заработная плата во 2-й отрасли +30 %, в 1-й и 3-й неизменна")
    note(ws, 28,
         "Система: (X_j − проч_j) p_j − Σ_i x_ij p_i = з/п_j (новое). "
         "Новая з/п: (19,5; 13,5·1,3; 19,5).", last="G", h=32)
    header(ws, 30, ["", "Отрасль 1", "Отрасль 2", "Отрасль 3"])
    put(ws, 31, 1, "Новая заработная плата")
    put(ws, 31, 2, "=B7", fmt="0.00")
    put(ws, 31, 3, "=C7*1.3", fmt="0.00", fill=GOLD)
    put(ws, 31, 4, "=D7", fmt="0.00")
    # M p = wage_new, M_jj = X_j - other_j - x_jj, M_ji = -x_ij
    # Use MINVERSE of 3x3
    header(ws, 33, ["Матрица M", "1", "2", "3"])
    # row j, col i: i==j: Xj - other_j - x_jj ; else -x_i,j
    # M[0,0] = X1 - o1 - x11 = B6-B8-B9
    put(ws, 34, 1, 1, bold=True)
    put(ws, 34, 2, "=B6-B8-B9", fmt="0.00")
    put(ws, 34, 3, "=-B10", fmt="0.00")
    put(ws, 34, 4, "=-B11", fmt="0.00")
    put(ws, 35, 1, 2, bold=True, fill=PALE)
    put(ws, 35, 2, "=-C9", fmt="0.00", fill=PALE)
    put(ws, 35, 3, "=C6-C8-C10", fmt="0.00", fill=PALE)
    put(ws, 35, 4, "=-C11", fmt="0.00", fill=PALE)
    put(ws, 36, 1, 3, bold=True)
    put(ws, 36, 2, "=-D9", fmt="0.00")
    put(ws, 36, 3, "=-D10", fmt="0.00")
    put(ws, 36, 4, "=D6-D8-D11", fmt="0.00")
    header(ws, 38, ["M⁻¹ · з/п", "p1", "p2", "p3"])
    ws["B39"] = ArrayFormula("B39:D41", "=MINVERSE(B34:D36)")
    for r in range(39, 42):
        for c in range(2, 5):
            ws.cell(row=r, column=c).number_format = "0.0000"
            ws.cell(row=r, column=c).font = F_N
            ws.cell(row=r, column=c).alignment = CEN
            ws.cell(row=r, column=c).border = BOX
        put(ws, r, 1, r - 38, bold=True)
    put(ws, 43, 1, "Индексы цен (п. 5)", bold=True, fill=GOLD)
    put(ws, 43, 2, "=B39*$B$31+C39*$C$31+D39*$D$31", fmt="0.0000", fill=GOLD, bold=True)
    put(ws, 43, 3, "=B40*$B$31+C40*$C$31+D40*$D$31", fmt="0.0000", fill=GOLD, bold=True)
    put(ws, 43, 4, "=B41*$B$31+C41*$C$31+D41*$D$31", fmt="0.0000", fill=GOLD, bold=True)
    result(ws, 45, "Индексы цен (п. 5): p1 ; p2 ; p3",
           '=TEXT(B43,"0.000")&" ; "&TEXT(C43,"0.000")&" ; "&TEXT(D43,"0.000")')
    note(ws, 47,
         "Рост з/п только во 2-й отрасли на 30 % приводит к умеренному росту цен во всех отраслях "
         "(около 7–14 %), наибольший — во 2-й отрасли.", last="G", h=32)
    page(ws)


def build():
    wb = openpyxl.Workbook()
    sheet_task(wb)
    sheet_mob(wb)
    sheet_prices(wb)
    wb.properties.creator = "Пась А. М."
    wb.properties.title = "УСР № 1. Модель межотраслевого баланса"
    wb.save(FN)
    print("сохранено", FN)


if __name__ == "__main__":
    build()
