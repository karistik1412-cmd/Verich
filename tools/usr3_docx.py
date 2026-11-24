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
import omml
import word_toc
# УСР 3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")
FN = os.path.join(ROOT, "Эконометрика УСР 3.docx")

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


def H1(text, page_break=True, bookmark=None):
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
    if bookmark:
        word_toc.bookmark(p, bookmark)
    return p


def H2(text, space_before=14, size=14, bookmark=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(4)
    pf.keep_with_next = True
    rich(p, text, size, bold=True)
    if bookmark:
        word_toc.bookmark(p, bookmark)
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

    tp("МИНИСТЕРСТВО ОБРАЗОВАНИЯ РЕСПУБЛИКИ БЕЛАРУСЬ", 16, False, 0, 0)
    tp("УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»", 16, False, 0, 12)
    tp("Инженерный факультет", 14, False, 12, 0)
    tp("Кафедра информационных технологий и интеллектуальных систем", 14, False, 0, 0)
    tp("ОТЧЁТ", 16, True, 72, 6)
    tp("УСР №3", 16, True, 0, 18)
    tp("по дисциплине: «Эконометрика и исследование операций в экономике»", 16, False, 0, 6)
    tp("на тему: «Модели анализа и оценки инвестиционных проектов»", 16, False, 0, 0)
    L = WD_ALIGN_PARAGRAPH.LEFT
    tp("Выполнил:", 16, False, 80, 0, None, L)
    tp("студент", 16, False, 0, 0, None, L)
    tp("группы 24ИТ-3\tИльич Э.С.", 16, False, 0, 0, None, L, 0.0, [16.5])
    tp("Проверил:", 16, False, 16, 0, None, L)
    tp("Ассистент\tВерич А.И.", 16, False, 0, 0, None, L, 0.0, [16.5])
    tp("ПИНСК 2026", 16, False, 36, 0)


TOC_ITEMS = [
    ("ЗАДАЧИ", 0, "toc_zad"),
    ("Задача 3.4. Предельный уровень инфляции", 1, "toc_34"),
    ("Задача 3.6. Инвестиции при инфляции", 1, "toc_36"),
    ("Задача 3.12. Оптимальный портфель", 1, "toc_312"),
    ("ПРИЛОЖЕНИЕ А", 0, "toc_a"),
]


def toc(pages=None):
    H1("ОГЛАВЛЕНИЕ", page_break=True)
    word_toc.fill_toc(doc, TOC_ITEMS, lambda p, t, s, bold=False: run(p, t, s, bold=bold))


def tasks():
    H1("ЗАДАЧА", page_break=True, bookmark="toc_zad")
    P("Решение записываем как на занятии. Excel – для подстановки в формулы NPV "
      "и для Поиска решения в задаче 3.12 (скрины в приложении).")

    H2("ЗАДАЧА 3.4", bookmark="toc_34")
    P("В учебнике – задача 1.4. I_{0} = 150, CF_{1} = CF_{2} = 150, r = 15 %. "
      "Найти предельный уровень инфляции, при котором NPV ≥ 0.")
    P("Ставка с учётом инфляции и NPV:")
    formula("d = (1+r)(1+i) − 1", "3.1")
    formula("NPV = −I_{0} + CF_{1}/(1+d) + CF_{2}/(1+d)^{2}", "3.2")
    P("При NPV = 0 и CF_{1} = CF_{2} = I_{0} = 150 получаем 1/(1+d) + 1/(1+d)^{2} = 1. "
      "Пусть φ = 1+d, тогда φ^{2} − φ − 1 = 0, φ* = (1+√5)/2,")
    formula("i* = φ*/(1+r) − 1 = (1+√5)/(2·1,15) − 1 ≈ 40,70 %.", "3.3")
    table(
        [["i", "NPV"],
         ["0 %", "93,86"],
         ["10 %", "62,31"],
         ["40 %", "1,04"],
         ["40,70 %", "0,00"],
         ["50 %", "−12,63"]],
        [4.0, 5.0], font=12,
        caption="Таблица 1")
    P("Проект эффективен при инфляции не выше 40,70 %.")

    H2("ЗАДАЧА 3.6", bookmark="toc_36")
    P("В учебнике – задача 1.6. I_{0} = 800, n = 5, ежегодный CF = 300, r = 12 %. "
      "Инфляция 10 %, 12 % и 18 %. NPV аннуитета:")
    formula("NPV = −I_{0} + CF · (1 − (1+d)^{−n}) / d", "3.4")
    P("d по (3.1). Для i = 10 %, 12 %, 18 %:")
    table(
        [["Инфляция", "d", "NPV", "Решение"],
         ["10 %", "23,20 %", "37,51", "принимаем"],
         ["12 %", "25,44 %", "−0,44", "отклоняем"],
         ["18 %", "32,16 %", "−98,53", "отклоняем"]],
        [3.2, 3.2, 3.2, 5.4], font=12,
        caption="Таблица 2")
    P("При инфляции 10 % проект ещё выгоден, при 12 % и 18 % – нет.")

    H2("ЗАДАЧА 3.12", bookmark="toc_312")
    P("Шесть проектов, r = 15 %, бюджет t = 0 не более 100 млн. Пары 1–2 и 3–5 взаимоисключающие, "
      "спрос на продукцию проектов 2, 3, 5 ≤ 550. NPV проекта:")
    formula("NPV = CF_{0} + CF_{2}/1,15^{2} + CF_{3}/1,15^{3} + CF_{4}/1,15^{4}", "3.5")
    omml.add_eqarr(doc, [["max Σ NPV_j x_j"]])
    omml.add_brace_system(doc, [
        ["x_j ∈ {0; 1}"],
        ["x_1 + x_2 ≤ 1,    x_3 + x_5 ≤ 1"],
        ["инвестиции t=0 ≤ 100"],
        ["продукция проектов 2, 3, 5 ≤ 550"],
    ])
    P("Оптимум:")
    table(
        [["Проект", "NPV, млн руб.", "x"],
         ["1", "22,00", "1"],
         ["2", "89,16", "0"],
         ["3", "19,56", "0"],
         ["4", "28,65", "1"],
         ["5", "29,74", "0"],
         ["6", "102,16", "1"],
         ["Итого NPV", "152,80", ""]],
        [4.0, 5.0, 3.0], font=12,
        caption="Таблица 3")
    P("Оптимальный портфель: проекты 1, 4 и 6, NPV ≈ 152,80 млн руб.")


def appendix_a():
    H1("ПРИЛОЖЕНИЕ А", bookmark="toc_a")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rich(p, "Файл Excel", 14, bold=True)
    P("«Эконометрика Верич УСР 3.xlsx»: лист «Задание» – формулировки; «3.4», «3.6», «3.12» – "
      "расчёты. Исходные данные выделены жёлтым, ответы – зелёным. Для 3.12 использована "
      "надстройка «Поиск решения».")


SECTIONS = [
    ("title", title_page),
    ("toc", None),
    ("tasks", tasks),
    ("appA", appendix_a),
]


def build(pages=None, skip=(), out=FN):
    global doc
    doc = Document()
    setup_styles()
    word_toc.enable_update_fields(doc)
    add_page_number_footer(doc.sections[0])
    doc.sections[0].different_first_page_header_footer = True
    mapping = dict(SECTIONS)
    for k, _ in SECTIONS:
        if k in skip:
            continue
        if k == "toc":
            toc(pages)
        else:
            mapping[k]()
    doc.core_properties.author = "Ильич Э. С."
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build(pages={
        "ЗАДАЧИ": 3,
        "Задача 3.4. Предельный уровень инфляции": 3,
        "Задача 3.6. Инвестиции при инфляции": 4,
        "Задача 3.12. Оптимальный портфель": 5,
        "ПРИЛОЖЕНИЕ А": 6,
    }))
