# -*- coding: utf-8 -*-
"""Построение Excel-файла с решением УСР №5."""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference

import data as D

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FN = os.path.join(ROOT, "Эконометрика Верич УСР 5.xlsx")

DARK = "1F3864"
MID = "2E5C9A"
LIGHT = "D9E2F3"
PALE = "EDF2F9"
ACCENT = "C00000"
GOLD = "FFF2CC"

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
F_RED = Font(name="Times New Roman", size=11, bold=True, color=ACCENT)

CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFTT = Alignment(horizontal="left", vertical="top", wrap_text=True)


def title(ws, row, text, last_col="L", size=14):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = Font(name="Times New Roman", size=size, bold=True, color=DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 25


def subtitle(ws, row, text, last_col="L"):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_SUB
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 17


def section(ws, row, text, last_col="L"):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_SECT
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.fill = PatternFill("solid", fgColor=PALE)
    c.border = BOTTOM
    ws.row_dimensions[row].height = 21


def header(ws, row, values, start=1, h=32):
    for k, v in enumerate(values):
        c = ws.cell(row=row, column=start + k, value=v)
        c.font = F_HDR
        c.fill = PatternFill("solid", fgColor=DARK)
        c.alignment = CEN
        c.border = Border(left=Side(style="thin", color="FFFFFF"),
                          right=Side(style="thin", color="FFFFFF"),
                          top=med, bottom=med)
    ws.row_dimensions[row].height = h


def cells(ws, row, values, start=1, fmt=None, alt=False, bold=False, center=True):
    for k, v in enumerate(values):
        c = ws.cell(row=row, column=start + k, value=v)
        c.font = F_B if bold else F_N
        c.alignment = CEN if center else LEFT
        if alt:
            c.fill = PatternFill("solid", fgColor=PALE)
        c.border = BOX
        if fmt:
            c.number_format = fmt


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def result_block(ws, row, label, value, last_col="F", wide=True):
    ws.merge_cells(f"A{row}:D{row}")
    c = ws[f"A{row}"]
    c.value = label
    c.font = F_RES
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.fill = PatternFill("solid", fgColor=GOLD)
    c.border = BOX
    ws.merge_cells(f"E{row}:{last_col}{row}")
    v = ws[f"E{row}"]
    v.value = value
    v.font = F_RES
    v.alignment = CEN
    v.fill = PatternFill("solid", fgColor=GOLD)
    v.border = BOX
    ws.row_dimensions[row].height = 30


def note(ws, row, text, last_col="H", h1=2):
    ws.merge_cells(f"A{row}:{last_col}{row + h1 - 1}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = F_SMALL
    c.alignment = LEFTT


def style_chart(ch, w=16, h=9):
    ch.width = w
    ch.height = h
    ch.style = 2


# =====================================================================  ЗАДАНИЕ
def sheet_task(wb):
    ws = wb.active
    ws.title = "Задание"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = DARK
    widths(ws, {"A": 2, "B": 34, "C": 50, "D": 26, "E": 14})

    ws.merge_cells("B2:D2")
    c = ws["B2"]
    c.value = "УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»"
    c.font = Font(name="Times New Roman", size=11, bold=True, color="404040")
    c.alignment = Alignment(horizontal="center")

    ws.merge_cells("B4:D4")
    c = ws["B4"]
    c.value = "Эконометрика и исследование операций в экономике"
    c.font = Font(name="Times New Roman", size=12, italic=True, color="404040")
    c.alignment = Alignment(horizontal="center")

    ws.merge_cells("B6:D6")
    c = ws["B6"]
    c.value = "УСР № 5"
    c.font = Font(name="Times New Roman", size=22, bold=True, color=DARK)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[6].height = 30

    ws.merge_cells("B7:D7")
    c = ws["B7"]
    c.value = "Модели сетевого планирования и управления"
    c.font = Font(name="Times New Roman", size=14, bold=True, color=MID)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[7].height = 22

    ws.merge_cells("B9:D9")
    c = ws["B9"]
    c.value = "Решённые задачи: 5.1 (е), 5.3 (а), 5.5 (а), 5.8 (а), 5.16 (а)"
    c.font = Font(name="Times New Roman", size=11, bold=True, color=ACCENT)
    c.alignment = Alignment(horizontal="center")

    ws["B11"] = "Исполнитель:"
    ws["B11"].font = F_B
    ws["C11"] = "студент группы 24ИТ-3   Ильич Э. С."
    ws["C11"].font = F_N
    ws["B12"] = "Проверил:"
    ws["B12"].font = F_B
    ws["C12"] = "ассистент   Верич А. И."
    ws["C12"].font = F_N

    section(ws, 14, "Состав работы", last_col="D")
    header(ws, 15, ["Лист", "Содержание расчётов"], start=2, h=22)
    ws.merge_cells("C15:D15")
    rows = [
        ("5.1", "Сетевой график проекта (вариант е): параметры событий и работ, резервы "
                "времени, критический путь, линейный график (диаграмма Ганта)"),
        ("5.3", "Оптимизация проекта по времени при директивном сроке t₀ = 35 дн.: расчёт "
                "дополнительных средств x_ij, моментов начала и окончания работ"),
        ("5.5", "Минимизация срока выполнения проекта при B = 42 ден. ед.: распределение "
                "дополнительных средств по работам"),
        ("5.8", "Оптимизация проекта по стоимости: пошаговое сокращение продолжительности "
                "работ по коэффициентам дополнительных затрат (КДЗ)"),
        ("5.16", "Оптимизация проекта по ресурсам при R = 9 ед.: построение линейного графика "
                 "и эпюры потребления ресурса"),
    ]
    r = 16
    for name, text in rows:
        c = ws.cell(row=r, column=2, value=name)
        c.font = F_B
        c.alignment = CEN
        c.border = BOX
        c.fill = PatternFill("solid", fgColor=LIGHT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        c = ws.cell(row=r, column=3, value=text)
        c.font = F_N
        c.alignment = LEFT
        c.border = BOX
        ws.row_dimensions[r].height = 32
        r += 1

    note(ws, 22, "Значения x_ij в задачах 5.3 и 5.5 получены надстройкой «Поиск решения» "
                 "(минимизация целевой функции при заданных ограничениях). Все производные "
                 "величины на листах рассчитаны по формулам Excel.", last_col="D")


# =====================================================================  5.1
def sheet_51(wb):
    ws = wb.create_sheet("5.1")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = MID
    widths(ws, {"A": 13, "B": 21, "C": 12, "D": 11, "E": 11, "F": 11, "G": 11,
                "H": 10, "I": 10, "J": 10, "K": 15})
    title(ws, 1, "Задача 5.1 (е). Сетевой график, параметры событий и работ, линейный график")
    subtitle(ws, 2, "8 работ, 9 событий; фиктивные работы (3,4), (5,8), (6,8) — нулевой продолжительности")

    section(ws, 4, "1. Исходные данные", last_col="C")
    header(ws, 5, ["Работа", "Предшествующие работы", "Время, t_ij"], h=28)
    r = 6
    for k, (name, prev, t) in enumerate(D.T51_TABLE):
        cells(ws, r, [name, prev, t], alt=(k % 2 == 1))
        r += 1

    tp, tn, T, rows = D.params51()

    section(ws, 15, "2. Параметры событий (четырёхсекторная схема)", last_col="D")
    header(ws, 16, ["Событие i", "Ранний срок t_р(i)", "Поздний срок t_п(i)", "Резерв R(i)"], h=30)
    r = 17
    for i in range(1, 10):
        crit = (tn[i] - tp[i]) == 0
        cells(ws, r, [i, tp[i], tn[i], tn[i] - tp[i]], alt=((i - 1) % 2 == 1), bold=crit)
        if crit:
            for c in "ABCD":
                ws[f"{c}{r}"].font = F_RED
        r += 1

    section(ws, 27, "3. Параметры работ", last_col="K")
    header(ws, 28, ["Работа", "Код (i, j)", "t_ij", "t_рн", "t_ро", "t_пн", "t_по",
                    "R_п", "R_с", "R_н", "Критическая"], h=42)
    r = 29
    for k, q in enumerate(rows):
        crit = q["rp"] == 0
        cells(ws, r, [q["name"].split()[0], f"({q['i']}, {q['j']})", q["t"], q["es"], q["ef"],
                      q["ls"], q["lf"], q["rp"], q["rs"], q["rn"], "да" if crit else "–"],
              alt=(k % 2 == 1), bold=crit)
        if crit:
            for c in "ABCDEFGHIJK":
                ws[f"{c}{r}"].font = F_RED
        r += 1

    note(ws, 38, "Нумерация событий выполнена по алгоритму Фалкерсона. Работа (3,4) отражает зависимость "
                 "a5 и a6 от a2; работы (5,8) и (6,8) — зависимость a7 от a4 и a5.", last_col="K")

    section(ws, 40, "4. Результаты", last_col="K")
    result_block(ws, 41, "Минимальное время выполнения проекта T_кр, дн.", T, last_col="F")
    result_block(ws, 42, "Критический путь", "1 – 3 – 4 – 6 – 8 – 9", last_col="F")
    ws.merge_cells("G41:K42")
    c = ws["G41"]
    c.value = "Критические работы: a2 (1,3), a5 (4,6), a7 (8,9);\nT_кр = 5 + 6 + 5 = 16 дн."
    c.font = F_N
    c.alignment = LEFTT
    c.fill = PatternFill("solid", fgColor=PALE)
    c.border = BOX

    section(ws, 44, "5. Линейный график проекта", last_col="K")
    hr = 45
    header(ws, hr, ["Работа", "Начало", "Длительность", "Окончание"], h=26)
    r = hr + 1
    for q in rows:
        cells(ws, r, [q["name"].split()[0], q["es"], q["t"], q["ef"]])
        r += 1
    last = r - 1

    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.gapWidth = 50
    data = Reference(ws, min_col=2, max_col=3, min_row=hr, max_row=last)
    cats = Reference(ws, min_col=1, min_row=hr + 1, max_row=last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.noFill = True
    ch.series[0].graphicalProperties.line.noFill = True
    ch.series[1].graphicalProperties.solidFill = "2E75B6"
    ch.series[1].graphicalProperties.line.solidFill = DARK
    ch.title = "Диаграмма Ганта"
    ch.x_axis.title = "дни"
    ch.y_axis.scaling.orientation = "maxMin"
    ch.y_axis.majorGridlines = None
    ch.legend = None
    style_chart(ch, 15, 8.5)
    ws.add_chart(ch, "M5")


# =====================================================================  5.3 / 5.5
def sheet_opt(wb, name, title_text, subtitle_text, works, t, d, k, x, limit, s,
              result_label, objective_is_result, note_text):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = MID
    widths(ws, {"A": 22, "B": 15, "C": 15, "D": 13, "E": 14, "F": 10, "G": 14, "H": 15})
    title(ws, 1, title_text)
    subtitle(ws, 2, subtitle_text)

    n = len(works)
    section(ws, 4, "1. Исходные данные", last_col="E")
    header(ws, 5, ["Работа (i, j)", "t_ij", "d_ij", "k_ij", "x_max = (t – d)/k"], h=32)
    for idx in range(n):
        r = 6 + idx
        cells(ws, r, [f"({works[idx][0]}, {works[idx][1]})", t[idx], d[idx], k[idx],
                      f"=ROUND((B{r}-C{r})/D{r},2)"], alt=(idx % 2 == 1))
        ws[f"D{r}"].number_format = "0.00"

    base = 6 + n + 2
    section(ws, base - 1, "2. Решение: дополнительные средства и сроки выполнения работ", last_col="H")
    header(ws, base, ["Работа (i, j)", "x_ij, ден. ед.", "t′_ij = t – k·x",
                      "t_н (i)", "t_о", "d_ij", "t_о – t_н ≥ d", "s_j ≥ t_о"], h=46)

    row_ev = base + n + 3
    for idx in range(n):
        r = base + 1 + idx
        i_, j_ = works[idx]
        sj_cell = f"{get_column_letter(1 + j_)}{row_ev + 1}"
        cells(ws, r, [f"({i_}, {j_})", x[idx],
                      f"=ROUND(B{6+idx}-D{6+idx}*B{r},4)",
                      s[i_ - 1],
                      f"=D{r}+C{r}",
                      d[idx],
                      f'=IF(ROUND(E{r}-D{r},6)>=F{r},"да","нет")',
                      f'=IF({sj_cell}>=ROUND(E{r},6),"да","нет")'],
              alt=(idx % 2 == 1))
        for col in "CDE":
            ws[f"{col}{r}"].number_format = "0.00"
        ws[f"B{r}"].number_format = "General"
        ws[f"B{r}"].fill = PatternFill("solid", fgColor=GOLD)
        ws[f"D{r}"].fill = PatternFill("solid", fgColor=GOLD)

    section(ws, row_ev - 1, "3. Сроки свершения событий плана", last_col="H")
    header(ws, row_ev, ["Событие"] + [f"s{k_}" for k_ in range(1, len(s) + 1)], h=20)
    cells(ws, row_ev + 1, ["Срок, дн."] + list(s))
    for idx in range(len(s)):
        ws.cell(row=row_ev + 1, column=2 + idx).fill = PatternFill("solid", fgColor=GOLD)
        ws.cell(row=row_ev + 1, column=2 + idx).number_format = "0.00"

    r1 = row_ev + 3
    ws.merge_cells(f"A{r1}:B{r1}")
    c = ws[f"A{r1}"]
    c.value = "Целевая функция  f = Σ x_ij  →  min"
    c.font = F_N
    c.alignment = LEFT
    c.border = BOX
    ws[f"C{r1}"] = f"=SUM(B{base+1}:B{base+n})"
    ws[f"C{r1}"].font = F_B
    ws[f"C{r1}"].alignment = CEN
    ws[f"C{r1}"].border = BOX
    ws[f"C{r1}"].number_format = "General"
    ws[f"C{r1}"].fill = PatternFill("solid", fgColor=GOLD)

    r2 = r1 + 1
    ws.merge_cells(f"A{r2}:B{r2}")
    ws[f"A{r2}"] = ("Директивное время t₀, дн." if name == "5.3" else "Выделенные средства B, ден. ед.")
    ws[f"A{r2}"].font = F_N
    ws[f"A{r2}"].alignment = LEFT
    ws[f"A{r2}"].border = BOX
    ws[f"C{r2}"] = limit
    ws[f"C{r2}"].font = F_B
    ws[f"C{r2}"].alignment = CEN
    ws[f"C{r2}"].border = BOX
    ws[f"C{r2}"].number_format = "General"

    r3 = r2 + 1
    ws.merge_cells(f"A{r3}:B{r3}")
    ws[f"A{r3}"] = "Срок выполнения проекта, дн."
    ws[f"A{r3}"].font = F_N
    ws[f"A{r3}"].alignment = LEFT
    ws[f"A{r3}"].border = BOX
    ws[f"C{r3}"] = f"=MAX(E{base+1}:E{base+n})"
    ws[f"C{r3}"].font = F_B
    ws[f"C{r3}"].alignment = CEN
    ws[f"C{r3}"].border = BOX
    ws[f"C{r3}"].number_format = "General"

    r4 = r3 + 1
    ws.merge_cells(f"A{r4}:B{r4}")
    ws[f"A{r4}"] = "Проверка ограничения"
    ws[f"A{r4}"].font = F_N
    ws[f"A{r4}"].alignment = LEFT
    ws[f"A{r4}"].border = BOX
    ws[f"C{r4}"] = (f'=IF(C{r3}<=C{r2},"выполнено","не выполнено")' if name == "5.3"
                    else f'=IF(C{r1}<=C{r2},"выполнено","не выполнено")')
    ws[f"C{r4}"].font = F_B
    ws[f"C{r4}"].alignment = CEN
    ws[f"C{r4}"].border = BOX

    val = f"=C{r1}" if objective_is_result else f"=C{r3}"
    result_block(ws, r4 + 2, result_label, val, last_col="H")
    note(ws, r4 + 4, note_text, last_col="H")


# =====================================================================  5.8
def sheet_58(wb):
    ws = wb.create_sheet("5.8")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = MID
    widths(ws, {"A": 7, "B": 17, "C": 8, "D": 9, "E": 9, "F": 11, "G": 10, "H": 10, "I": 10,
                "J": 11, "K": 15.5, "L": 15.5})
    title(ws, 1, "Задача 5.8 (а). Оптимизация проекта по стоимости (минимизация удорожания)")
    subtitle(ws, 2, "Заданы нормальные D_ij и срочные d_ij продолжительности, стоимости c_ij и C_ij; "
                    "фиктивная работа (2,3) имеет нулевую продолжительность")

    section(ws, 4, "1. Исходные данные и КДЗ работ", last_col="L")
    header(ws, 5, ["Работа", "D_ij", "d_ij", "c_ij", "C_ij", "h_ij = (C–c)/(D–d)"], h=52)
    for idx, (i_, j_) in enumerate(D.W58):
        r = 6 + idx
        cells(ws, r, [f"({i_}, {j_})", D.DD58[idx], D.dd58[idx], D.cc58[idx], D.CC58[idx],
                      f"=ROUND((E{r}-D{r})/(B{r}-C{r}),2)"], alt=(idx % 2 == 1))
        ws[f"F{r}"].font = F_B
        ws[f"F{r}"].number_format = "0.00"

    section(ws, 13, "2. Пошаговая оптимизация (алгоритм КДЗ)", last_col="L")
    hdr = (["Шаг", "Сокращаемые работы", "Δ, дн."] +
           [f"t({W[0]},{W[1]})" for W in D.W58] +
           ["T_кр, дн.", "Прирост стоимости", "Стоимость проекта"])
    header(ws, 14, hdr, h=56)
    hidx = {W: f"F{6 + k}" for k, W in enumerate(D.W58)}

    for step, sel, delta, durs, T, dc in D.STEPS58:
        r = 15 + step
        cells(ws, r, [step, sel, delta] + list(durs), alt=(step % 2 == 1))
        col = {4 + k: get_column_letter(4 + k) for k in range(6)}
        a, b = col[4] + str(r), col[5] + str(r)      # (1,2), (1,3)
        c1, d1 = col[6] + str(r), col[7] + str(r)    # (2,4), (2,5)
        e1, f1 = col[8] + str(r), col[9] + str(r)    # (3,4), (4,5)
        # T_кр = max по полным путям
        ws.cell(row=r, column=10, value=f"=MAX({a}+{e1}+{f1},{a}+{c1}+{f1},{a}+{d1},{b}+{e1}+{f1})")
        # прирост стоимости
        if step == 0:
            ws.cell(row=r, column=11, value="–")
    # прирост стоимости (формулы) и накопленная стоимость
    for step, sel, delta, durs, T, dc in D.STEPS58:
        r = 15 + step
        for cidx in (3, 10, 11, 12):
            cc = ws.cell(row=r, column=cidx)
            cc.font = F_B if cidx in (10, 11) else F_N
            cc.alignment = CEN
            cc.border = BOX
            if cidx == 10:
                cc.fill = PatternFill("solid", fgColor=LIGHT)
            if cidx == 12:
                cc.number_format = "#,##0"
            if cidx == 11:
                cc.number_format = "#,##0"
        if step > 0:
            parts = [f"C{r}*{hidx[w]}" for w in D.W58 if f"({w[0]},{w[1]})" in sel]
            ws.cell(row=r, column=11, value="=" + "+".join(parts))
            ws.cell(row=r, column=11).font = F_B
            ws.cell(row=r, column=11).alignment = CEN
            ws.cell(row=r, column=11).border = BOX
            ws.cell(row=r, column=11).number_format = "#,##0"
        ws.cell(row=r, column=12,
                value=("=SUM($D$6:$D$11)" if step == 0 else f"=L{r-1}+K{r}"))
        ws.cell(row=r, column=12).font = F_N
        ws.cell(row=r, column=12).alignment = CEN
        ws.cell(row=r, column=12).border = BOX
        ws.cell(row=r, column=12).number_format = "#,##0"
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=1).alignment = CEN

    last = 15 + len(D.STEPS58) - 1
    section(ws, last + 2, "3. Результаты оптимизации", last_col="F")
    result_block(ws, last + 3, "Минимально возможное критическое время T_кр, дн.", f"=J{last}")
    result_block(ws, last + 4, "Прирост стоимости проекта, ден. ед.", f"=L{last}-L15")
    result_block(ws, last + 5, "Стоимость выполнения проекта, ден. ед.", f"=L{last}")
    note(ws, last + 7,
         "На каждом шаге сокращается работа (работы) с наименьшим КДЗ из числа критических. Если "
         "критические пути не имеют общих работ, сокращаются работы с минимальным КДЗ на каждом пути; "
         "если сумма КДЗ этих работ больше КДЗ общей работы — сокращается общая работа. Величина "
         "сокращения ограничена минимальной продолжительностью работы и моментом появления нового "
         "критического пути.", last_col="H")


# =====================================================================  5.16
def sheet_516(wb):
    ws = wb.create_sheet("5.16")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = MID
    last = "R"
    widths(ws, {"A": 17, "B": 9, "C": 9, "D": 10, "E": 10, "F": 10, "G": 8,
                "H": 7, "I": 7, "J": 7, "K": 7, "L": 7, "M": 7, "N": 7,
                "O": 7, "P": 7, "Q": 7, "R": 7})
    title(ws, 1, "Задача 5.16 (а). Оптимизация проекта по ресурсам (R = 9 ед.)", last_col=last)
    subtitle(ws, 2, "Работы не допускают перерыва в выполнении; каждой работе приписаны t_ij и r_ij",
             last_col=last)

    section(ws, 4, "1. Исходные данные и параметры сетевого графика", last_col="F")
    header(ws, 5, ["Работа (i, j)", "t_ij", "r_ij", "t_рн", "t_ро", "R_п"], h=28)
    tp, tn, T = D.net([(f"w{k}", w[1], w[2], w[3]) for k, w in enumerate(D.W516)], 5)
    n = len(D.W516)
    for k, w in enumerate(D.W516):
        rp = tn[w[2]] - tp[w[1]] - w[3]
        cells(ws, 6 + k, [w[0], w[3], w[4], tp[w[1]], tp[w[1]] + w[3], rp], alt=(k % 2 == 1))

    rT = 6 + n
    ws.merge_cells(f"A{rT}:B{rT}")
    ws[f"A{rT}"] = "T_кр, дн."
    ws[f"A{rT}"].font = F_B
    ws[f"A{rT}"].alignment = CEN
    ws[f"A{rT}"].border = BOX
    ws[f"C{rT}"] = T
    ws[f"C{rT}"].font = F_B
    ws[f"C{rT}"].alignment = CEN
    ws[f"C{rT}"].border = BOX
    ws.merge_cells(f"D{rT}:F{rT}")
    ws[f"D{rT}"] = "критический путь 1 – 3 – 4 – 5"
    ws[f"D{rT}"].font = F_SMALL
    ws[f"D{rT}"].alignment = LEFT

    r2 = rT + 2
    section(ws, r2, "2. Потребность в ресурсе при ранних сроках начала работ", last_col=last)
    header(ws, r2 + 1, ["День"] + list(range(1, 18)), h=20)
    cells(ws, r2 + 2, ["Ресурс, ед."] + [
        f"=SUMPRODUCT(($D$6:$D$11<{get_column_letter(2+k)}${r2+1})*"
        f"($D$6:$D$11+$B$6:$B$11>={get_column_letter(2+k)}${r2+1})*$C$6:$C$11)"
        for k in range(17)])
    cells(ws, r2 + 3, ["Максимум", f"=MAX(B{r2+2}:R{r2+2})"])
    ws.merge_cells(f"A{r2+4}:R{r2+4}")
    c = ws[f"A{r2+4}"]
    c.value = ("Максимальная потребность 12 ед. превышает R = 9 ед. — сдвиг работы (1,2) "
               "на 4 дня в пределах полного резерва")
    c.font = F_RED
    c.alignment = LEFT

    r3 = r2 + 6
    section(ws, r3, "3. Оптимизированный план (сдвиг работ в пределах резервов)", last_col=last)
    header(ws, r3 + 1, ["Работа (i, j)", "Начало", "Окончание", "t_ij", "r_ij", "Сдвиг, дн."], h=28)
    for k, w in enumerate(D.W516):
        st = D.ST516_OPT[k]
        rr = r3 + 2 + k
        cells(ws, rr, [w[0], st, f"=B{rr}+D{rr}", w[3], w[4],
                       (st - tp[w[1]]) if st - tp[w[1]] else "–"], alt=(k % 2 == 1))
        ws[f"B{rr}"].fill = PatternFill("solid", fgColor=GOLD)
    opt0, opt1 = r3 + 2, r3 + 1 + n

    r4 = opt1 + 2
    section(ws, r4, "4. Эпюра потребления ресурса после оптимизации", last_col=last)
    header(ws, r4 + 1, ["День"] + list(range(1, 18)), h=20)
    cells(ws, r4 + 2, ["Ресурс, ед."] + [
        f"=SUMPRODUCT(($B${opt0}:$B${opt1}<{get_column_letter(2+k)}${r4+1})*"
        f"($B${opt0}:$B${opt1}+$D${opt0}:$D${opt1}>={get_column_letter(2+k)}${r4+1})*$E${opt0}:$E${opt1})"
        for k in range(17)])
    for k in range(17):
        cell = ws.cell(row=r4 + 2, column=2 + k)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.font = F_B
    cells(ws, r4 + 3, ["Ограничение R"] + ["=$E$37"] * 17)
    for k in range(17):
        cell = ws.cell(row=r4 + 3, column=2 + k)
        cell.font = Font(name="Times New Roman", size=9, italic=True, color=ACCENT)

    r5 = r4 + 5
    section(ws, r5, "5. Результаты", last_col=last)
    # result_block uses E row; put R in E37 if r5+1==37 else fix R refs
    result_block(ws, r5 + 1, "Ограничение R, ед.", D.R516, last_col="F")
    result_block(ws, r5 + 2, "Максимальная потребность в ресурсе, ед.",
                 f"=MAX(B{r4+2}:R{r4+2})", last_col="F")
    result_block(ws, r5 + 3, "Время выполнения проекта T, дн.",
                 f"=MAX(C{opt0}:C{opt1})", last_col="F")
    # R constraint row should point to the R result cell
    r_cell = f"E{r5+1}"
    for k in range(17):
        ws.cell(row=r4 + 3, column=2 + k).value = f"=${r_cell}"
    note(ws, r5 + 5,
         "Сдвинута работа (1,2) на 4 дня (в пределах Rп = 6); работа (2,5) начинается в t = 10. "
         "Максимальная потребность равна 9 ед., срок проекта совпадает с критическим — 17 дней.",
         last_col="H")

    ch = BarChart()
    ch.type = "col"
    ch.gapWidth = 40
    data = Reference(ws, min_col=2, max_col=18, min_row=r4 + 2, max_row=r4 + 2)
    cats = Reference(ws, min_col=2, max_col=18, min_row=r4 + 1, max_row=r4 + 1)
    ch.add_data(data, from_rows=True, titles_from_data=False)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.solidFill = "8FAADC"
    ch.series[0].graphicalProperties.line.solidFill = DARK
    ch.title = "Эпюра потребления ресурса"
    ch.y_axis.title = "ед."
    ch.x_axis.title = "дни"
    ch.legend = None
    ch.y_axis.majorGridlines = None
    style_chart(ch, 15, 7.5)

    ln = LineChart()
    ln.add_data(Reference(ws, min_col=2, max_col=18, min_row=r4 + 3, max_row=r4 + 3),
                from_rows=True, titles_from_data=False)
    ln.series[0].graphicalProperties.line.solidFill = ACCENT
    ln.series[0].graphicalProperties.line.dashStyle = "dash"
    ln.series[0].graphicalProperties.line.width = 18000
    ln.series[0].marker = None
    ch += ln
    ws.add_chart(ch, "A44")


def main():
    wb = openpyxl.Workbook()
    sheet_task(wb)
    sheet_51(wb)
    sheet_opt(wb, "5.3",
              "Задача 5.3 (а). Оптимизация проекта по времени при t₀ = 35 дн.",
              "Проект из 6 событий и 10 работ; минимизируется сумма дополнительных средств",
              D.W53, D.T53, D.D53, D.K53, D.X53, D.T0_53, D.S53,
              "Минимальные дополнительные средства, ден. ед.", True,
              "Значения x_ij найдены надстройкой «Поиск решения»: минимизация f = Σx при ограничениях "
              "t_о ≤ 35; t_о – t_н ≥ d_ij; t_о – t_н = t_ij – k_ij·x_ij; 0 ≤ x_ij ≤ (t_ij – d_ij)/k_ij. "
              "Проект выполняется за 35 дней вместо 40; критический путь 1 – 4 – 5 – 6.")
    sheet_opt(wb, "5.5",
              "Задача 5.5 (а). Минимизация срока выполнения проекта при B = 42 ден. ед.",
              "Проект из 5 событий и 7 работ; распределение дополнительных средств x_ij",
              D.W55, D.T55, D.D55, D.K55, D.X55, D.B55, D.S55,
              "Минимальное время выполнения проекта, дн.", False,
              "Значения x_ij найдены надстройкой «Поиск решения»: минимизация t_о(5) при ограничениях "
              "Σx ≤ 42; t_о – t_н ≥ d_ij; t_о – t_н = t_ij – k_ij·x_ij; 0 ≤ x_ij ≤ (t_ij – d_ij)/k_ij. "
              "Срок сокращён с 40 до 30 дней при затратах 41 ден. ед. (B = 42).")
    sheet_58(wb)
    sheet_516(wb)

    wb.properties.creator = "Ильич Э. С."
    wb.properties.title = "УСР № 5. Модели сетевого планирования и управления"
    wb.active = 0
    wb.save(FN)
    print("сохранено:", FN)


if __name__ == "__main__":
    main()
