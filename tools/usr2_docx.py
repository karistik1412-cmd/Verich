# -*- coding: utf-8 -*-
"""Сборка отчёта по УСР №2 (Word)."""
import os
import re
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import usr2_data as D

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")
FN = os.path.join(ROOT, "Эконометрика УСР 2.docx")

FONT = "Times New Roman"
DARK = RGBColor(0x1F, 0x38, 0x64)
RED = RGBColor(0xC0, 0x00, 0x00)
GREY = RGBColor(0x59, 0x59, 0x59)
HDR_FILL = "DEEAF6"
ALT_FILL = "F2F6FC"

doc = Document()


# ----------------------------------------------------------------- утилиты
def setup_styles():
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(14)
    st._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st._element.rPr.rFonts.set(qn("w:cs"), FONT)
    pf = st.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # поля страницы
    s = doc.sections[0]
    s.page_width = Cm(21)
    s.page_height = Cm(29.7)
    s.left_margin = Cm(3)
    s.right_margin = Cm(1.5)
    s.top_margin = Cm(2)
    s.bottom_margin = Cm(2)
    s.footer_distance = Cm(1.0)


def run(p, text, size=14, bold=False, italic=False, color=None, sub=False, sup=False):
    r = p.add_run(text)
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if sub:
        r.font.subscript = True
    if sup:
        r.font.superscript = True
    if color is not None:
        r.font.color.rgb = color
    return r


TOKEN = re.compile(r"(_\{[^}]*\}|\^\{[^}]*\}|\*[^*]+\*)")
NBSP = "\u00a0"


def rich(p, text, size=14, bold=False, italic=False, color=None):
    """Разметка: _{sub}, ^{sup}, *курсив*."""
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith("_{"):
            run(p, part[2:-1], size, bold, italic, color, sub=True)
        elif part.startswith("^{"):
            run(p, part[2:-1], size, bold, italic, color, sup=True)
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            run(p, part[1:-1], size, bold, True, color)
        else:
            part = part.replace(" · ", NBSP + "·" + NBSP).replace(" = ", NBSP + "=" + NBSP)
            part = re.sub(r"\[[^\]]*\]", lambda m: m.group(0).replace(" ", NBSP), part)
            run(p, part, size, bold, italic, color)


def P(text="", indent=True, align=None, size=14, bold=False, italic=False, color=None,
      space_before=0, space_after=0, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(1.25) if indent else Cm(0)
    pf.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    pf.keep_with_next = keep
    if text:
        rich(p, text, size, bold, italic, color)
    return p


_FIRST_H1 = [True]


def H1(text, page_break=True):
    if len(doc.paragraphs) == 0 and len(doc.tables) == 0:
        page_break = False
    if page_break:
        doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.space_before = Pt(0)
    rich(p, text, 16, bold=True, color=DARK)
    return p


def H2(text, space_before=14, size=14):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(4)
    pf.keep_with_next = True
    rich(p, text, size, bold=True, color=DARK)
    return p


def FCAP(text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(4)
    pf.space_after = Pt(10)
    rich(p, text, 13, bold=False, italic=False)
    return p


def TCAP(text, keep=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(10)
    pf.space_after = Pt(4)
    pf.keep_with_next = keep
    rich(p, text, 13)
    return p


def formula(text, number=None, size=14):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.15
    if number is None:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rich(p, text, size, italic=True)
        return p
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ts = pf.tab_stops
    ts.add_tab_stop(Cm(8.0), WD_TAB_ALIGNMENT.CENTER)
    ts.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT)
    run(p, "\t", size)
    rich(p, text, size, italic=True)
    run(p, "\t(" + number + ")", size)
    return p


def figure(fname, width_cm, caption):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(os.path.join(ASSETS, fname), width=Cm(width_cm))
    FCAP(caption)


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def set_cell(cell, text, size=11, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, italic=False,
             color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.25
    pf.space_before = Pt(1)
    pf.space_after = Pt(1)
    pf.alignment = align
    rich(p, str(text), size, bold, italic, color)


def table(rows, widths, font=11, header_rows=1, align=None, highlight=None, bold_rows=(),
          caption=None, center_table=True):
    if caption:
        TCAP(caption)
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    keep = len(rows) <= (header_rows + 9)
    for r_i, row in enumerate(rows):
        tr = t.rows[r_i]
        if r_i < header_rows:
            trPr = tr._tr.get_or_add_trPr()
            th = OxmlElement("w:tblHeader")
            trPr.append(th)
        for c_i, val in enumerate(row):
            cell = t.cell(r_i, c_i)
            cell.width = Cm(widths[c_i])
            is_head = r_i < header_rows
            al = WD_ALIGN_PARAGRAPH.CENTER
            if align and not is_head:
                al = align[c_i]
            set_cell(cell, val, font, bold=is_head or (r_i in bold_rows), align=al)
            if is_head:
                shade(cell, HDR_FILL)
            elif highlight and (r_i, c_i) in highlight:
                shade(cell, highlight[(r_i, c_i)])
            elif r_i % 2 == 0:
                shade(cell, ALT_FILL)
            if keep and r_i < len(rows) - 1:
                for cp in cell.paragraphs:
                    cp.paragraph_format.keep_with_next = True
    for c_i, w in enumerate(widths):
        for row in t.rows:
            row.cells[c_i].width = Cm(w)
    return t


def bullets(items, size=14):
    for it in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(1.25)
        pf.left_indent = Cm(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        rich(p, "– " + it, size)


def add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run(p, "", 12)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    rPr.append(rf)
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    p._p.append(fld)



def title_page():
    def tp(text, size=14, bold=False, before=0, after=0, color=None,
           align=WD_ALIGN_PARAGRAPH.CENTER, indent=0.0, tabs=None):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(indent)
        pf.alignment = align
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = 1.0
        if tabs:
            for pos in tabs:
                pf.tab_stops.add_tab_stop(Cm(pos), WD_TAB_ALIGNMENT.RIGHT)
        rich(p, text, size, bold=bold, color=color)
        return p

    tp("МИНИСТЕРСТВО ОБРАЗОВАНИЯ РЕСПУБЛИКИ БЕЛАРУСЬ", 13, True, 0, 4)
    tp("УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»", 13, True, 0, 26)
    tp("Инженерный факультет", 14, False, 0, 0)
    tp("Кафедра информационных технологий и интеллектуальных систем", 13.5, False, 0, 0)
    tp("ОТЧЁТ", 22, True, 110, 8)
    tp("УСР №2", 16, True, 0, 30)
    tp("по дисциплине: «Эконометрика и исследование операций в экономике»", 14, False, 0, 8)
    tp("на тему: «Модели и методы многоцелевой оптимизации»", 14, False, 0, 12)
    tp("Задача 2.4 (а)–(г)", 14, True, 0, 0, DARK)

    R = WD_ALIGN_PARAGRAPH.RIGHT
    tp("Выполнил:", 14, False, 104, 0, None, R)
    tp("студент группы 24ИТ-3", 14, False, 0, 0, None, R)
    tp("Ильич Э. С.", 14, False, 0, 0, None, R)
    tp("Проверил:", 14, False, 22, 0, None, R)
    tp("ассистент", 14, False, 0, 0, None, R)
    tp("Верич А. И.", 14, False, 0, 0, None, R)
    tp("ПИНСК 2026", 14, True, 40, 0)


TOC_ITEMS = [
    ("ЗАДАЧИ", 0),
    ("Задача 2.4 (а). Метод последовательных уступок", 1),
    ("Задача 2.4 (б). Метод ведущего критерия", 1),
    ("Задача 2.4 (в). Равные относительные отклонения", 1),
    ("Задача 2.4 (г). Метод минимакса", 1),
    ("ПРИЛОЖЕНИЕ А", 0),
    ("ПРИЛОЖЕНИЕ Б", 0),
]


def toc(pages):
    H1("ОГЛАВЛЕНИЕ", page_break=True)
    pages = pages or {}
    for text, lvl in TOC_ITEMS:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0 if lvl == 0 else 0.75)
        pf.line_spacing = 1.5
        pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, 1)
        run(p, text, 14, bold=(lvl == 0))
        run(p, "\t" + str(pages.get(text, "")) if pages else "\t", 14, bold=(lvl == 0))


def intro_common():
    H1("ЗАДАЧИ", page_break=True)
    P("Решение выполнено в Excel (файл «Эконометрика Верич УСР 2.xlsx»), а не «в тетради». "
      "Ниже — условие и то, что получено на листах: жёлтые ячейки — ввод, формулы — расход ресурсов "
      "и критерии, для ЛП использован «Поиск решения».")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rich(p, "ЗАДАЧА 2.4", 14, bold=True, color=DARK)

    P("Для производства трёх видов продукции предприятие использует три вида ресурсов. "
      "Нормы расхода ресурсов на единицу продукции и наличие ресурсов приведены в таблице 2.1.")
    table(
        [["Ресурсы", "П_{1}", "П_{2}", "П_{3}", "Наличие"],
         ["Ресурс 1", "2", "4", "4", "400"],
         ["Ресурс 2", "3", "2", "2", "300"],
         ["Ресурс 3", "4", "5", "3", "500"]],
        [3.6, 2.6, 2.6, 2.6, 3.2], font=12,
        caption="Таблица 2.1 – Нормы расхода и наличие ресурсов")
    P("Единица продукции каждого вида характеризуется прибылью (22; 12; 14), оптовой ценой "
      "(24; 16; 30) и трудоёмкостью (20; 10; 15). Продукции третьего вида должно быть "
      "произведено не менее 30 ед.")
    P("Требуется найти решение задачи:")
    bullets([
        "методом последовательных уступок при важности критериев: прибыль – уступка 580, "
        "оптовая цена – уступка 1090, трудоёмкость;",
        "методом ведущего критерия: самый важный критерий – прибыль, нижняя граница оптовой "
        "цены 2400, трудоёмкость не более 1600;",
        "методом равных и наименьших относительных отклонений при x_{1} ≥ 20;",
        "методом минимакса при объёме производства продукции третьего вида не менее 40 ед.",
    ])
    H2("Как устроен Excel")
    P("На листе «Задание» в B6:D8 стоят нормы расхода, в E6:E8 – наличие ресурсов, "
      "в B12:D14 – прибыль, цена и трудоёмкость единицы. Переменные x1, x2, x3 на рабочих "
      "листах – жёлтые ячейки. Критерии считаются формулами вида")
    formula("=22·B6+12·C6+14·D6", None)
    P("расход ресурса 1: =2·B6+4·C6+4·D6 и так далее. Планы ищутся надстройкой "
      "«Поиск решения» (симплекс, неотрицательные x, x3 ≥ 30).")


def fmt(x, n=2):
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    s = f"{x:.{n}f}".rstrip("0").replace(".", ",")
    return s


def plan_row(name, x, f):
    return [name, fmt(x[0], 3), fmt(x[1], 3), fmt(x[2], 3),
            fmt(f[0], 2), fmt(f[1], 2), fmt(f[2], 2)]


def task24a():
    H2("ЗАДАЧА 2.4 (а)", space_before=16)
    P("Лист «2.4а». Три блока «Поиска решения» на одних и тех же ячейках B6:D6.")
    P("Шаг 1. Цель – ячейка C9 (прибыль =22·B6+12·C6+14·D6) → max. Ограничения: "
      "расход ресурсов ≤ E6:E8 листа «Задание», D6 ≥ 30. Solver даёт B6:D6 = 80; 0; 30, "
      "C9 = 2180. Уступка 580 → в модель добавляется C9 ≥ 1600.")
    P("Шаг 2. Цель – C10 (цена) → max при том же и C9 ≥ 1600. Получаем 50; 0; 75, "
      "C10 = 3450. Уступка 1090 → C10 ≥ 2360.")
    P("Шаг 3. Цель – C11 (труд) → min при C9 ≥ 1600 и C10 ≥ 2360. План в B6:D6:")
    table(
        [["Шаг", "x_{1}", "x_{2}", "x_{3}", "f_{1}", "f_{2}", "f_{3}"],
         plan_row("1. max f_{1}", D.A1_X, D.A1_F),
         plan_row("2. max f_{2}", D.A2_X, D.A2_F),
         plan_row("3. min f_{3}", D.A3_X, D.A3_F)],
        [3.4, 2.0, 2.2, 2.0, 2.2, 2.2, 2.2], font=11,
        caption="Таблица 2.2 – Шаги метода последовательных уступок")
    P("Компромиссное решение задачи 2.4 (а): произвести 21,25 ед. продукции П_{1}, "
      "59,375 ед. П_{2} и 30 ед. П_{3}.")


def task24b():
    H2("ЗАДАЧА 2.4 (б)")
    P("Лист «2.4б». В B6:D6 – переменные. C9 = прибыль (цель Поиска решения, max). "
      "Ограничения Solver: C10 ≥ 2400, C11 ≤ 1600, D6 ≥ 30, расход ресурсов ≤ наличия. "
      "После запуска B6:D6 = 30; 55; 30.")
    P("C9 = 1740, C10 = 2500, C11 = 1600 (ограничение по труду активное).")
    table(
        [["Показатель", "План", "Ограничение", "Выполнение"],
         ["x_{1}, x_{2}, x_{3}", "30; 55; 30", "x_{3} ≥ 30", "да"],
         ["Прибыль f_{1}", "1740", "max", "–"],
         ["Оптовая цена f_{2}", "2500", "≥ 2400", "да"],
         ["Трудоёмкость f_{3}", "1600", "≤ 1600", "да"]],
        [4.4, 3.6, 3.6, 3.0], font=12,
        caption="Таблица 2.3 – Решение методом ведущего критерия")
    P("Полученный план допустим. По прибыли он хуже частного максимума 2180, что является "
      "платой за выполнение ограничений по цене и трудоёмкости.")


def task24c():
    H2("ЗАДАЧА 2.4 (в)")
    P("Лист «2.4в». Сначала три запуска Поиска решения (max C для f1, max для f2, min для f3) "
      "при x1 ≥ 20 – частные оптимумы в B6:E8:")
    table(
        [["Критерий", "x_{1}", "x_{2}", "x_{3}", "f_k*"],
         ["max f_{1}", "80", "0", "30", "2180"],
         ["max f_{2}", "50", "0", "75", "3450"],
         ["min f_{3}", "20", "0", "30", "850"]],
        [3.6, 2.4, 2.4, 2.4, 3.0], font=12,
        caption="Таблица 2.4 – Частные оптимумы (x_{1} ≥ 20)")
    P("В компромиссном блоке в B12:D12 снова x. Отклонения считаются формулами "
      "δ1 = 1−f1/2180, δ2 = 1−f2/3450, δ3 = f3/850−1. Solver: min любое δ при "
      "δ1=δ2=δ3 и x1 ≥ 20. План: B12=20, C12≈25,10, D12≈37,15. Все δ ≈ 0,4214.")
    table(
        [["Критерий", "f_k", "f_k*", "δ_k"],
         ["f_{1} (max)", "1261,27", "2180", "0,4214"],
         ["f_{2} (max)", "1996,05", "3450", "0,4214"],
         ["f_{3} (min)", "1208,22", "850", "0,4214"]],
        [4.0, 3.4, 3.4, 3.4], font=12,
        caption="Таблица 2.5 – Относительные отклонения в компромиссной точке")
    P("Таким образом, ни один критерий не «жертвует» больше остальных: все три отклоняются "
      "от своего частного оптимума на одну и ту же долю 42,14 %.")


def task24d():
    H2("ЗАДАЧА 2.4 (г)")
    P("Лист «2.4г». Частные оптимумы при x3 ≥ 40 – три запуска Solver, результаты в B6:E8:")
    table(
        [["Критерий", "x_{1}", "x_{2}", "x_{3}", "f_k*"],
         ["max f_{1}", "73,333", "0", "40", "2173,33"],
         ["max f_{2}", "50", "0", "75", "3450"],
         ["min f_{3}", "0", "0", "40", "600"]],
        [3.6, 2.6, 2.4, 2.4, 3.0], font=12,
        caption="Таблица 2.6 – Частные оптимумы (x_{3} ≥ 40)")
    P("В блоке плана B12:D12 – x, отдельная ячейка y. Solver: y → min при "
      "δk ≤ y, D12 ≥ 40. Получается x = (0; 33,46; 40), y ≈ 0,5576. "
      "δ1 = δ3 = y, δ2 меньше y.")
    table(
        [["Критерий", "f_k", "f_k*", "δ_k"],
         ["f_{1} (max)", "961,47", "2173,33", "0,5576"],
         ["f_{2} (max)", "1735,30", "3450", "0,4970"],
         ["f_{3} (min)", "934,56", "600", "0,5576"]],
        [4.0, 3.4, 3.4, 3.4], font=12,
        caption="Таблица 2.7 – Отклонения в минимаксной точке")
    P("Метод минимакса гарантирует, что самое плохое (наибольшее) относительное отклонение "
      "сделано как можно меньшим. По сравнению с пунктом (в) допустимое множество уже "
      "(x_{3} ≥ 40 вместо 30), поэтому y* выше, чем общее δ пункта (в).")
    H2("Сводка решений")
    table(
        [["Метод", "x_{1}", "x_{2}", "x_{3}", "f_{1}", "f_{2}", "f_{3}"],
         ["а) уступки", "21,25", "59,375", "30", "1600", "2360", "1468,75"],
         ["б) ведущий", "30", "55", "30", "1740", "2500", "1600"],
         ["в) равные δ", "20,00", "25,10", "37,15", "1261,27", "1996,05", "1208,22"],
         ["г) минимакс", "0", "33,46", "40", "961,47", "1735,30", "934,56"]],
        [3.2, 2.0, 2.2, 2.0, 2.4, 2.4, 2.4], font=10.5,
        caption="Таблица 2.8 – Сравнение четырёх методов")
    P("Метод ведущего критерия даёт наибольшую прибыль среди компромиссных планов (1740), "
      "метод уступок – наименьшую трудоёмкость при сохранении оговорённых уступок, "
      "методы равных отклонений и минимакса выравнивают потери критериев, но снижают прибыль.")


def appendix_a():
    H1("ПРИЛОЖЕНИЕ А")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rich(p, "Решение задач в Excel", 14, bold=True)
    P("Расчёты выполнены в Microsoft Excel: целевые функции и расход ресурсов заданы "
      "формулами SUMPRODUCT; планы получены решением соответствующих задач ЛП "
      "(надстройка «Поиск решения»).")
    for fn, cap, w in [
        ("xl2_task.png", "Рисунок А.1 – Исходные данные задачи 2.4", 14.6),
        ("xl2_a.png", "Рисунок А.2 – Метод последовательных уступок", 10.8),
        ("xl2_b.png", "Рисунок А.3 – Метод ведущего критерия", 14.6),
        ("xl2_c.png", "Рисунок А.4 – Равные относительные отклонения", 13.4),
        ("xl2_d.png", "Рисунок А.5 – Метод минимакса", 13.4),
        ("xl2_sum.png", "Рисунок А.6 – Сводка решений", 14.6),
    ]:
        figure(fn, w, cap)


def appendix_b():
    H1("ПРИЛОЖЕНИЕ Б")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rich(p, "Конспект теоретических вопросов", 14, bold=True)

    H2("2.2 Метод линейной комбинации частных критериев")
    P("Пусть задан вектор весовых коэффициентов важности критериев "
      "α = (α_{1}, …, α_K), α_k > 0, Σ α_k = 1. Критерии предварительно нормируют "
      "(делят на свои оптимумы или заменяют относительными отклонениями). Многоцелевая "
      "задача сводится к одноцелевой:")
    formula("F(X) = Σ_{k=1}^{K} α_k f_k(X)  →  max", "2.21")
    P("при исходных ограничениях φ_i(X) ≤ (≥, =) b_i, X ≥ 0. Решение зависит от выбора весов: "
      "разные α дают разные точки области компромисса (Парето).")

    H2("2.3 Метод последовательных уступок")
    P("Критерии упорядочивают по важности. Сначала решают задачу по самому важному критерию "
      "и находят f_{1}*. Назначают уступку Δf_{1} и переходят ко второму критерию с "
      "дополнительным ограничением f_{1} ≥ f_{1}* − Δf_{1} (если f_{1} → max) либо "
      "f_{1} ≤ f_{1}* + Δf_{1} (если f_{1} → min). После нахождения f_{2}* назначают уступку "
      "Δf_{2} и решают задачу по третьему критерию с двумя дополнительными ограничениями, "
      "и так далее до наименее важного критерия.")
    P("Если полученные значения критериев не устраивают ЛПР, величины уступок изменяют "
      "и расчёт повторяют. Метод прозрачен содержательно: уступка – это та потеря по "
      "важному критерию, на которую согласны ради улучшения следующего.")

    H2("2.4 Метод ведущего критерия")
    P("Все критерии, кроме самого важного, переводятся в ограничения. Критерии минимизации "
      "умножают на −1. Обозначив через β_k нижние границы остальных критериев, модель имеет вид")
    formula("F(X) = f_{1}(X)  →  max", "2.22")
    formula("f_k(X) ≥ β_k,     k = 2, …, K", "2.23")
    P("при исходных ограничениях и X ≥ 0. Полученное решение может не принадлежать области "
      "компромиссов, поэтому его необходимо проверить на эффективность по Парето: не должно "
      "существовать допустимого плана, улучшающего хотя бы один критерий без ухудшения остальных.")

    H2("2.5 Метод равных и наименьших относительных отклонений")
    P("Ищут план, в котором относительные отклонения всех критериев от своих частных оптимумов "
      "равны и как можно меньше. Для критериев с одинаковым направлением оптимизации")
    formula("d_i f_i − d_j f_j = 0,     d_k = 1 / f_k*,", "2.24")
    P("для критериев с противоположными направлениями")
    formula("d_i f_i + d_j f_j = 2.", "2.25")
    P("Эти равенства добавляют к исходным ограничениям; в качестве целевой функции можно взять "
      "любой из частных критериев.")

    H2("2.6 Метод минимакса")
    P("Сначала находят частные оптимумы f_k*. Вводят переменную y – наибольшее относительное "
      "отклонение. Для максимизируемых критериев")
    formula("f_k(X) + f_k* · y ≥ f_k*,", "2.26")
    P("для минимизируемых")
    formula("f_k(X) − f_k* · y ≤ f_k*.", "2.27")
    P("Целевая функция расширенной задачи: y → min. Оптимум минимизирует «наихудшее» "
      "относительное отклонение, что соответствует принципу гарантированного результата.")


SECTIONS = [
    ("title", title_page),
    ("toc", lambda: toc(None)),
    ("intro", intro_common),
    ("task24a", task24a),
    ("task24b", task24b),
    ("task24c", task24c),
    ("task24d", task24d),
    ("appA", appendix_a),
    ("appB", appendix_b),
]


def build(pages=None, skip=(), out=FN):
    global doc
    doc = Document()
    setup_styles()
    add_page_number_footer(doc.sections[0])
    doc.sections[0].different_first_page_header_footer = True
    # toc with pages
    mapping = dict(SECTIONS)
    order = [k for k, _ in SECTIONS]
    for k in order:
        if k in skip:
            continue
        if k == "toc":
            toc(pages)
        else:
            mapping[k]()
    doc.save(out)
    return out


if __name__ == "__main__":
    build()
    print("saved", FN)
