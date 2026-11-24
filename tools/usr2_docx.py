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
import omml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")
FN = os.path.join(ROOT, "Эконометрика УСР 2.docx")

FONT = "Times New Roman"
DARK = RGBColor(0x00, 0x00, 0x00)
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
    rich(p, text, 16, bold=True)
    return p


def H2(text, space_before=14, size=14):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(4)
    pf.keep_with_next = True
    rich(p, text, size, bold=True)
    return p


def FCAP(text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(4)
    pf.space_after = Pt(10)
    rich(p, text, 12, bold=True)
    return p


def TCAP(text, keep=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(10)
    pf.space_after = Pt(4)
    pf.keep_with_next = keep
    rich(p, text, 12, bold=True)
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

    tp("МИНИСТЕРСТВО ОБРАЗОВАНИЯ РЕСПУБЛИКИ БЕЛАРУСЬ", 16, False, 0, 0)
    tp("УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»", 16, False, 0, 12)
    tp("Инженерный факультет", 14, False, 12, 0)
    tp("Кафедра информационных технологий и интеллектуальных систем", 14, False, 0, 0)
    tp("ОТЧЁТ", 16, True, 72, 6)
    tp("УСР №2", 16, True, 0, 18)
    tp("по дисциплине: «Эконометрика и исследование операций в экономике»", 16, False, 0, 6)
    tp("на тему: «Модели и методы многоцелевой оптимизации»", 16, False, 0, 0)

    L = WD_ALIGN_PARAGRAPH.LEFT
    tp("Выполнил:", 16, False, 80, 0, None, L)
    tp("студент", 16, False, 0, 0, None, L)
    tp("группы 24ИТ-3\tИльич Э.С.", 16, False, 0, 0, None, L, tabs=[16.5])
    tp("Проверил:", 16, False, 16, 0, None, L)
    tp("Ассистент\tВерич А.И.", 16, False, 0, 0, None, L, tabs=[16.5])
    tp("ПИНСК 2026", 16, False, 36, 0)


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
    H1("ЗАДАЧА", page_break=True)
    P("Для производства трех видов продукции предприятие использует три вида ресурсов. "
      "Данные представлены в табл. 1:")
    table(
        [["Ресурсы", "П_{1}", "П_{2}", "П_{3}", "Наличие"],
         ["Ресурс 1", "2", "4", "4", "400"],
         ["Ресурс 2", "3", "2", "2", "300"],
         ["Ресурс 3", "4", "5", "3", "500"]],
        [3.6, 2.6, 2.6, 2.6, 3.2], font=12,
        caption="Таблица 1")
    P("Единица продукции каждого вида характеризуется следующими показателями: прибылью – "
      "(22; 12; 14), оптовой ценой – (24; 16; 30), трудоемкостью – (20; 10; 15). Третьего вида "
      "продукции должно быть произведено не менее 30 единиц. Требуется найти решение задачи:")
    P("а) методом последовательных уступок при следующей важности критериев: прибыль – уступка 580, "
      "оптовая цена – уступка 1090, трудоемкость;")
    P("б) методом ведущего критерия: в качестве самого важного критерия считать прибыль, "
      "нижняя граница оптовой цены составляет 2400, трудоемкость не должна превышать 1600;")
    P("в) методом равных и наименьших относительных отклонений при объеме производства "
      "продукции первого вида не меньше 20 единиц;")
    P("г) методом минимакса при объеме производства продукции третьего вида не меньше 40 единиц.")
    H2("Решение:")
    P("Составим ЭММ задачи, где x_{1}, x_{2}, x_{3} – объёмы выпуска продукции П_{1}, П_{2}, П_{3}.")
    P("f_{1} – прибыль;  f_{2} – оптовая цена;  f_{3} – трудоёмкость.")
    omml.add_emm_24(doc)


def fmt(x, n=2):
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    s = f"{x:.{n}f}".rstrip("0").replace(".", ",")
    return s


def plan_row(name, x, f):
    return [name, fmt(x[0], 3), fmt(x[1], 3), fmt(x[2], 3),
            fmt(f[0], 2), fmt(f[1], 2), fmt(f[2], 2)]


def task24a():
    P("а) Решение задачи методом последовательных уступок.", indent=True, bold=True)
    P("Предпочтительность всех критериев определена. Находим оптимальное решение по "
      "критерию I – прибыли (рис. 1.1):")
    figure("xl2_a1.png", 14.6, "Рисунок 1.1 – Шаг 1. Максимум прибыли")
    formula("x_{1} = 80,   x_{2} = 0,   x_{3} = 30,   f_{1}* = 2180.", None)
    P("Далее, так как Δf_{1} = 580, решаем задачу по критерию II с дополнительным ограничением "
      "f_{1} ≥ 2180 − 580, то есть")
    formula("22x_{1} + 12x_{2} + 14x_{3} ≥ 1600.", None)
    P("Оптимальное решение по критерию II (рис. 1.2):")
    figure("xl2_a2.png", 14.6, "Рисунок 1.2 – Шаг 2. Максимум оптовой цены при уступке по прибыли")
    P("x_{1} = 50, x_{2} = 0, x_{3} = 75, f_{2}* = 3450. Уступка по цене 1090:")
    formula("24x_{1} + 16x_{2} + 30x_{3} ≥ 2360.", None)
    P("Решаем по критерию III с учётом ранее введённых ограничений (рис. 1.3):")
    figure("xl2_a3.png", 14.6, "Рисунок 1.3 – Шаг 3. Минимум трудоёмкости (компромисс)")
    P("Компромиссный план:")
    table(
        [["Шаг", "x_{1}", "x_{2}", "x_{3}", "f_{1}", "f_{2}", "f_{3}"],
         plan_row("1. max f_{1}", D.A1_X, D.A1_F),
         plan_row("2. max f_{2}", D.A2_X, D.A2_F),
         plan_row("3. min f_{3}", D.A3_X, D.A3_F)],
        [3.4, 2.0, 2.2, 2.0, 2.2, 2.2, 2.2], font=11,
        caption="Таблица 2")
    P("Компромиссное решение задачи 2.4 (а): произвести 21,25 ед. продукции П_{1}, "
      "59,375 ед. П_{2} и 30 ед. П_{3}.")


def task24b():
    P("б) Решение задачи методом ведущего критерия.", indent=True, bold=True)
    P("Все критерии, кроме самого важного, переводятся в разряд ограничений. ЭММ:")
    omml.add_eqarr(doc, [
        ["max ", ("f", "1"), " = 22", ("x", "1"), " + 12", ("x", "2"), " + 14", ("x", "3")],
    ])
    omml.add_brace_system(doc, [
        ["24", ("x", "1"), " + 16", ("x", "2"), " + 30", ("x", "3"), " ≥ 2400"],
        ["20", ("x", "1"), " + 10", ("x", "2"), " + 15", ("x", "3"), " ≤ 1600"],
        ["и ограничения (1)–(4)"],
    ])
    P("Результат (рис. 1.4):")
    figure("xl2_b1.png", 14.6, "Рисунок 1.4 – Ведущий критерий: план и целевые функции")
    figure("xl2_b2.png", 14.6, "Рисунок 1.5 – Ведущий критерий: ресурсы и значения f_k")
    formula("X* = (30; 55; 30),   f_{1} = 1740,   f_{2} = 2500,   f_{3} = 1600.", None)
    table(
        [["Показатель", "План", "Ограничение", "Выполнение"],
         ["x_{1}, x_{2}, x_{3}", "30; 55; 30", "x_{3} ≥ 30", "да"],
         ["Прибыль f_{1}", "1740", "max", "–"],
         ["Оптовая цена f_{2}", "2500", "≥ 2400", "да"],
         ["Трудоёмкость f_{3}", "1600", "≤ 1600", "да"]],
        [4.4, 3.6, 3.6, 3.0], font=12,
        caption="Таблица 3")
    P("Полученный план допустим. По прибыли он хуже частного максимума 2180, что является "
      "платой за выполнение ограничений по цене и трудоёмкости.")


def task24c():
    P("в) Решение задачи методом равных и наименьших относительных отклонений.", indent=True, bold=True)
    P("Сначала на листе «2.4в» решаем задачу по каждому критерию отдельно (x_{1} ≥ 20). "
      "Частные оптимумы (рис. 1.6):")
    figure("xl2_c1.png", 14.6, "Рисунок 1.6 – Частные оптимумы при x_{1} ≥ 20")
    table(
        [["Критерий", "x_{1}", "x_{2}", "x_{3}", "f_k*"],
         ["max f_{1}", "80", "0", "30", "2180"],
         ["max f_{2}", "50", "0", "75", "3450"],
         ["min f_{3}", "20", "0", "30", "850"]],
        [3.6, 2.4, 2.4, 2.4, 3.0], font=12,
        caption="Таблица 4")
    P("Обозначим d_k = 1 / f_k*. Условие равенства относительных отклонений:")
    formula("d_{1} f_{1} − d_{2} f_{2} = 0,     d_{1} f_{1} + d_{3} f_{3} = 2.", None)
    P("Решая ЛП с этими равенствами и x_{1} ≥ 20, получаем")
    formula("x_{1} = 20,   x_{2} ≈ 25,10,   x_{3} ≈ 37,15,   δ_{1} = δ_{2} = δ_{3} ≈ 0,4214.", None)
    table(
        [["Критерий", "f_k", "f_k*", "δ_k"],
         ["f_{1} (max)", "1261,27", "2180", "0,4214"],
         ["f_{2} (max)", "1996,05", "3450", "0,4214"],
         ["f_{3} (min)", "1208,22", "850", "0,4214"]],
        [4.0, 3.4, 3.4, 3.4], font=12,
        caption="Таблица 5")
    figure("xl2_c2.png", 14.6, "Рисунок 1.7 – Компромиссный план (равные отклонения)")
    figure("xl2_c3.png", 14.6, "Рисунок 1.8 – Относительные отклонения δ_k")
    P("Таким образом, ни один критерий не «жертвует» больше остальных: все три отклоняются "
      "от своего частного оптимума на одну и ту же долю 42,14 %.")


def task24d():
    P("г) Решение задачи методом минимакса.", indent=True, bold=True)
    P("При x_{3} ≥ 40 сначала находим частные оптимумы (рис. 1.9). Вводим y – наибольшее относительное "
      "отклонение. Для максимизируемых критериев")
    figure("xl2_d1.png", 14.6, "Рисунок 1.9 – Частные оптимумы при x_{3} ≥ 40")
    formula("f_k + f_k* · y ≥ f_k*,     k = 1, 2,", None)
    P("для минимизируемого")
    formula("f_{3} − f_{3}* · y ≤ f_{3}*,     y → min.", None)
    P("Оптимум:")
    table(
        [["Критерий", "x_{1}", "x_{2}", "x_{3}", "f_k*"],
         ["max f_{1}", "73,333", "0", "40", "2173,33"],
         ["max f_{2}", "50", "0", "75", "3450"],
         ["min f_{3}", "0", "0", "40", "600"]],
        [3.6, 2.6, 2.4, 2.4, 3.0], font=12,
        caption="Таблица 6")
    formula("x_{1} = 0,   x_{2} ≈ 33,46,   x_{3} = 40,   y* ≈ 0,5576.", None)
    P("δ_{1} = δ_{3} = y*, δ_{2} ≈ 0,4970 < y*.")
    table(
        [["Критерий", "f_k", "f_k*", "δ_k"],
         ["f_{1} (max)", "961,47", "2173,33", "0,5576"],
         ["f_{2} (max)", "1735,30", "3450", "0,4970"],
         ["f_{3} (min)", "934,56", "600", "0,5576"]],
        [4.0, 3.4, 3.4, 3.4], font=12,
        caption="Таблица 7")
    figure("xl2_d2.png", 14.6, "Рисунок 1.10 – Компромиссный план (минимакс)")
    figure("xl2_d3.png", 14.6, "Рисунок 1.11 – Отклонения δ_k и y*")
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
        caption="Таблица 8")
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
        ("xl2_task1.png", "Рисунок А.1 – Нормы расхода ресурсов", 14.6),
        ("xl2_task2.png", "Рисунок А.2 – Показатели единицы продукции", 14.6),
        ("xl2_a1.png", "Рисунок А.3 – Уступки, шаг 1 (max прибыль)", 14.6),
        ("xl2_a2.png", "Рисунок А.4 – Уступки, шаг 2 (max цена)", 14.6),
        ("xl2_a3.png", "Рисунок А.5 – Уступки, шаг 3 (min труд)", 14.6),
        ("xl2_b1.png", "Рисунок А.6 – Ведущий критерий: план", 14.6),
        ("xl2_b2.png", "Рисунок А.7 – Ведущий критерий: ресурсы", 14.6),
        ("xl2_c1.png", "Рисунок А.8 – Равные отклонения: частные оптимумы", 14.6),
        ("xl2_c2.png", "Рисунок А.9 – Равные отклонения: план", 14.6),
        ("xl2_c3.png", "Рисунок А.10 – Равные отклонения: δ_k", 14.6),
        ("xl2_d1.png", "Рисунок А.11 – Минимакс: частные оптимумы", 14.6),
        ("xl2_d2.png", "Рисунок А.12 – Минимакс: план", 14.6),
        ("xl2_d3.png", "Рисунок А.13 – Минимакс: y*", 14.6),
        ("xl2_sum.png", "Рисунок А.14 – Сводка решений", 14.6),
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
