# -*- coding: utf-8 -*-
"""Excel УСР №3, задачи 3.4, 3.6, 3.12 (в учебнике 1.4, 1.6, 3.12)."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FN = os.path.join(ROOT, "Эконометрика Верич УСР 3.xlsx")

CAL = "Calibri"
BLUE = "5B9BD5"
YELLOW = "FFFF99"
GREEN = "C6EFCE"
THIN = Border(
    left=Side(style="thin", color="808080"),
    right=Side(style="thin", color="808080"),
    top=Side(style="thin", color="808080"),
    bottom=Side(style="thin", color="808080"),
)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def font(size=11, bold=False, color="000000", italic=False):
    return Font(name=CAL, size=size, bold=bold, color=color, italic=italic)


def fill(hexc):
    return PatternFill("solid", fgColor=hexc)


def cell(ws, r, c, v, bold=False, fmt=None, bg=None, size=11, align="center"):
    x = ws.cell(row=r, column=c, value=v)
    x.font = font(size, bold)
    x.alignment = CEN if align == "center" else LEFT
    x.border = THIN
    if fmt:
        x.number_format = fmt
    if bg:
        x.fill = fill(bg)
    return x


def hdr(ws, r, vals, start=1):
    for k, v in enumerate(vals):
        x = cell(ws, r, start + k, v, bold=True, bg=BLUE)
        x.font = font(11, True, "FFFFFF")


def w(ws, m):
    for k, v in m.items():
        ws.column_dimensions[k].width = v


def sheet_34(wb):
    ws = wb.active
    ws.title = "3.4"
    w(ws, {"A": 38, "B": 16, "C": 16, "D": 16, "E": 16, "F": 14})
    ws["A1"] = "Задача 3.4 (в учебнике — 1.4). Уровень инфляции"
    ws["A1"].font = font(16, True, "2F5496")
    ws.merge_cells("A1:E1")
    ws["A2"] = "Найти инфляцию i, до которой NPV >= 0. Ставка 15%, I0=150, CF=150 и 150."
    ws["A2"].font = font(10, False, "666666", True)
    ws.merge_cells("A2:E2")

    hdr(ws, 4, ["Показатель", "Значение"])
    cell(ws, 5, 1, "Начальные инвестиции I0", align="left")
    cell(ws, 5, 2, 150, bg=YELLOW, fmt="0")
    cell(ws, 6, 1, "CF1", align="left")
    cell(ws, 6, 2, 150, bg=YELLOW, fmt="0")
    cell(ws, 7, 1, "CF2", align="left")
    cell(ws, 7, 2, 150, bg=YELLOW, fmt="0")
    cell(ws, 8, 1, "Ставка дисконта r", align="left")
    cell(ws, 8, 2, 0.15, bg=YELLOW, fmt="0%")

    ws["A10"] = "Дисконт с учётом инфляции: d = (1+r)*(1+i) − 1"
    ws["A10"].font = font(10, italic=True)
    ws.merge_cells("A10:E10")

    hdr(ws, 12, ["i", "d", "NPV", "Проект"])
    # table of inflations + goal
    infs = [0, 0.05, 0.10, 0.20, 0.30, 0.40, 0.406986, 0.50]
    for k, i in enumerate(infs):
        r = 13 + k
        cell(ws, r, 1, i, fmt="0.00%", bg=YELLOW)
        cell(ws, r, 2, f"=(1+$B$8)*(1+A{r})-1", fmt="0.00%")
        cell(ws, r, 3, f"=-$B$5+$B$6/(1+B{r})+$B$7/(1+B{r})^2", fmt="0.00")
        cell(ws, r, 4, f'=IF(C{r}>=0,"эффективен","неэффективен")')

    ws["A22"] = "Критическая инфляция (NPV = 0) — подбором / из уравнения d² − d − 1 = 0, d=(1+r)(1+i)"
    ws["A22"].font = font(10, italic=True)
    ws.merge_cells("A22:E22")
    hdr(ws, 23, ["d*", "i* = d*/(1+r) − 1", "проверка NPV"])
    cell(ws, 24, 1, "=(1+SQRT(5))/2", fmt="0.0000", bg=GREEN, bold=True)
    cell(ws, 24, 2, "=A24/(1+B8)-1", fmt="0.00%", bg=GREEN, bold=True)
    cell(ws, 24, 3, "=-B5+B6/(1+(1+B8)*(1+B24)-1)+B7/(1+(1+B8)*(1+B24)-1)^2", fmt="0.00", bg=GREEN)
    ws["A26"] = "Вывод: проект эффективен при инфляции не выше i* ≈ 40,70 %."
    ws["A26"].font = font(11, True)
    ws.merge_cells("A26:E26")

    ch = BarChart()
    ch.type = "col"
    ch.title = "NPV при разной инфляции"
    data = Reference(ws, min_col=3, min_row=12, max_row=20)
    cats = Reference(ws, min_col=1, min_row=13, max_row=20)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.legend = None
    ch.y_axis.title = "NPV"
    ch.width = 14
    ch.height = 8
    ws.add_chart(ch, "F4")


def sheet_36(wb):
    ws = wb.create_sheet("3.6")
    w(ws, {"A": 36, "B": 14, "C": 14, "D": 16, "E": 18, "F": 16})
    ws["A1"] = "Задача 3.6 (в учебнике — 1.6). Инвестиции и инфляция"
    ws["A1"].font = font(16, True, "2F5496")
    ws.merge_cells("A1:E1")
    ws["A2"] = "I0=800, n=5 лет, CF=300 каждый год, r=12%. Инфляция 10%, 12%, 18%."
    ws["A2"].font = font(10, False, "666666", True)

    hdr(ws, 4, ["Параметр", "Значение"])
    cell(ws, 5, 1, "I0", align="left"); cell(ws, 5, 2, 800, bg=YELLOW)
    cell(ws, 6, 1, "CF (ежегодно)", align="left"); cell(ws, 6, 2, 300, bg=YELLOW)
    cell(ws, 7, 1, "Срок n, лет", align="left"); cell(ws, 7, 2, 5, bg=YELLOW)
    cell(ws, 8, 1, "Ставка r", align="left"); cell(ws, 8, 2, 0.12, bg=YELLOW, fmt="0%")

    hdr(ws, 10, ["Инфляция i", "d=(1+r)(1+i)−1", "NPV", "Вывод"])
    for k, i in enumerate((0.10, 0.12, 0.18)):
        r = 11 + k
        cell(ws, r, 1, i, fmt="0%", bg=YELLOW)
        cell(ws, r, 2, f"=(1+$B$8)*(1+A{r})-1", fmt="0.00%")
        # NPV = -I + CF * (1 - (1+d)^-n)/d
        cell(ws, r, 3,
             f"=-$B$5+$B$6*(1-(1+B{r})^(-$B$7))/B{r}", fmt="0.00")
        cell(ws, r, 4, f'=IF(C{r}>0;"принимаем";"отклоняем")')

    ws["A15"] = "NPV считала аннуитетом: CF * (1 − (1+d)^(−n)) / d − I0. Жёлтые ячейки — исходные данные."
    ws["A15"].font = font(10, italic=True)
    ws.merge_cells("A15:E15")
    ws["A17"] = "При 10% инфляции NPV > 0, при 12% почти 0 (чуть меньше), при 18% проект невыгоден."
    ws["A17"].font = font(11, True)
    ws.merge_cells("A17:E17")


def sheet_312(wb):
    ws = wb.create_sheet("3.12")
    w(ws, {chr(65+i): 12 for i in range(12)})
    ws.column_dimensions["A"].width = 22
    ws["A1"] = "Задача 3.12. Портфель из 6 проектов, ставка 15%, бюджет 100 на t=0"
    ws["A1"].font = font(16, True, "2F5496")
    ws.merge_cells("A1:L1")
    ws["A2"] = "Взаимоисключающие: 1 и 2; 3 и 5. Спрос на продукцию проектов 2,3,5 ≤ 550. x — 0/1, Solver."
    ws["A2"].font = font(10, False, "666666", True)
    ws.merge_cells("A2:L2")

    hdr(ws, 4, ["Проект", "t=0", "t=2", "t=3", "t=4", "NPV (15%)", "Продукция", "x (0/1)"])
    # data
    data = [
        (1, None, -90, 50, 100, 0),
        (2, -50, 40, 70, 110, 400),
        (3, -40, 30, 30, 30, 110),
        (4, None, -40, 20, 80, 0),
        (5, -60, 20, 70, 50, 150),
        (6, -100, 70, 140, 100, 0),
    ]
    for k, (n, c0, c2, c3, c4, q) in enumerate(data):
        r = 5 + k
        cell(ws, r, 1, n, bold=True)
        cell(ws, r, 2, c0 if c0 is not None else 0, fmt="0", bg=YELLOW)
        cell(ws, r, 3, c2, fmt="0", bg=YELLOW)
        cell(ws, r, 4, c3, fmt="0", bg=YELLOW)
        cell(ws, r, 5, c4, fmt="0", bg=YELLOW)
        # NPV = C0 + C2/1.15^2 + C3/1.15^3 + C4/1.15^4
        cell(ws, r, 6, f"=B{r}+C{r}/1.15^2+D{r}/1.15^3+E{r}/1.15^4", fmt="0.00")
        cell(ws, r, 7, q, fmt="0")
        # x
        xv = [1, 0, 0, 1, 0, 1][k]  # optimal 1,4,6
        cell(ws, r, 8, xv, bg=YELLOW, fmt="0")

    ws["A12"] = "Ограничения (для Поиска решения)"
    ws["A12"].font = font(12, True, "2F5496")
    hdr(ws, 13, ["Ограничение", "Формула", "Значение", "Предел"])
    cell(ws, 14, 1, "x1 + x2 ≤ 1", align="left")
    cell(ws, 14, 2, "=H5+H6")
    cell(ws, 14, 3, "=B14")
    cell(ws, 14, 4, 1)
    cell(ws, 15, 1, "x3 + x5 ≤ 1", align="left")
    cell(ws, 15, 2, "=H7+H9")
    cell(ws, 15, 3, "=B15")
    cell(ws, 15, 4, 1)
    cell(ws, 16, 1, "Инвестиции t=0 ≤ 100", align="left")
    cell(ws, 16, 2, "=-B6*H6-B7*H7-B9*H9-B10*H10")
    cell(ws, 16, 3, "=B16", fmt="0.00")
    cell(ws, 16, 4, 100)
    cell(ws, 17, 1, "Продукция ≤ 550", align="left")
    cell(ws, 17, 2, "=G6*H6+G7*H7+G9*H9")
    cell(ws, 17, 3, "=B17")
    cell(ws, 17, 4, 550)
    cell(ws, 18, 1, "Целевая: сумма NPV·x", align="left")
    cell(ws, 18, 2, "=SUMPRODUCT(F5:F10,H5:H10)", fmt="0.00", bg=GREEN, bold=True)
    cell(ws, 18, 3, "=B18", fmt="0.00", bg=GREEN, bold=True)
    cell(ws, 18, 4, "max", bg=GREEN)

    ws["A20"] = "Поиск решения: изменяемые H5:H10 (бинары), макс C18, ограничения C14:C17."
    ws["A20"].font = font(10, italic=True)
    ws.merge_cells("A20:H20")
    ws["A21"] = "Оптимум: проекты 1, 4 и 6. NPV ≈ 152,80 млн руб. (бюджет t=0 закрывает проект 6, 1 и 4 стартуют в году 2)."
    ws["A21"].font = font(11, True)
    ws.merge_cells("A21:H21")

    hdr(ws, 23, ["Проект", "В портфеле?", "NPV вклада"])
    for k in range(6):
        r = 24 + k
        cell(ws, r, 1, k + 1)
        cell(ws, r, 2, f'=IF(H{5+k}=1,"да","нет")')
        cell(ws, r, 3, f"=F{5+k}*H{5+k}", fmt="0.00")


def sheet_task(wb):
    ws = wb.create_sheet("Задание", 0)
    w(ws, {"A": 80})
    ws["A1"] = "УСР №3. Модели анализа инвестиционных проектов"
    ws["A1"].font = font(16, True, "2F5496")
    ws["A3"] = "В учебнике задачи темы 3 пронумерованы как 1.4, 1.6 и 3.12 (стр. 107–109). Считаем их как 3.4, 3.6, 3.12."
    ws["A3"].font = font(11)
    ws["A5"] = "3.4  I0=150, CF=150; 150, r=15%. Найти инфляцию, до которой проект эффективен."
    ws["A6"] = "3.6  I0=800, 5 лет по 300, r=12%, инфляция 10 / 12 / 18 %. Стоит ли инвестировать?"
    ws["A7"] = "3.12 Шесть проектов, 100 млн на начало 1-го года, r=15%. Исключающие пары 1–2 и 3–5, спрос ≤ 550."
    for r in range(5, 8):
        ws[f"A{r}"].font = font(11)


def build():
    wb = Workbook()
    sheet_34(wb)
    sheet_36(wb)
    sheet_312(wb)
    sheet_task(wb)
    wb.properties.creator = "Ильич Э. С."
    wb.properties.title = "УСР № 3"
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(FN)
    print("сохранено", FN)


if __name__ == "__main__":
    build()
