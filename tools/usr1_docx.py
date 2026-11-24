# -*- coding: utf-8 -*-
"""Сборка отчёта по УСР №1 (Word)."""
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
# данные задачи 1.6 зашиты в текст отчёта

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")
FN = os.path.join(ROOT, "Эконометрика УСР 1.docx")

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
    tp("УСР №1", 16, True, 0, 18)
    tp("по дисциплине: «Эконометрика и исследование операций в экономике»", 16, False, 0, 6)
    tp("на тему: «Модели межотраслевого баланса»", 16, False, 0, 0)
    L = WD_ALIGN_PARAGRAPH.LEFT
    tp("Выполнил:", 16, False, 80, 0, None, L)
    tp("студент", 16, False, 0, 0, None, L)
    tp("группы 24ИТ-3\tИльич Э.С.", 16, False, 0, 0, None, L, 0.0, [16.5])
    tp("Проверил:", 16, False, 16, 0, None, L)
    tp("Ассистент\tВерич А.И.", 16, False, 0, 0, None, L, 0.0, [16.5])
    tp("ПИНСК 2026", 16, False, 36, 0)


TOC_ITEMS = [
    ("ЗАДАЧИ", 0),
    ("Задача 1.6. Модель межотраслевого баланса", 1),
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


def f4(x):
    return f"{x:.4f}".replace(".", ",")


def f2(x):
    return f"{x:.2f}".replace(".", ",")


def task16():
    H1("ЗАДАЧА", page_break=True)
    P("Условие – ниже. Решение: схема МОБ и формулы баланса, числа считаем в Excel "
      "(файл «Эконометрика Верич УСР 1.xlsx», функции МОБР и МУМНОЖ).")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    rich(p, "ЗАДАЧА 1.6 (б)", 14, bold=True)

    P("Вариант б табл. 1.3 учебника: p = 30 % ВДС на з/п; n_{1} = 1, r = 2, p_{1} = 30 % "
      "(рост з/п отстаёт от цен, k_w = 0,7); n_{2} = 2, p_{2} = 30 %. "
      "Потоки и Y^{отч} – в таблице 1. Требуется:")
    bullets([
        "построить схему МОБ за отчётный период с выделением промежуточного потребления и затрат, "
        "заработной платы и прочих элементов ВДС;",
        "рассчитать матрицу коэффициентов прямых затрат A;",
        "определить плановый валовой выпуск X^{пл} и построить схему МОБ, если "
        "Y^{пл} = (120; 30; 60);",
        "найти индексы цен, если цена продукции 1-й отрасли увеличивается в 2 раза, а рост "
        "заработной платы отстаёт от роста цен на 30 % (коэффициент эластичности 0,7);",
        "найти индексы цен, если заработная плата во 2-й отрасли увеличивается на 30 %, "
        "а в остальных отраслях остаётся неизменной.",
    ])

    table(
        [["Производитель \\ потребитель", "1", "2", "3", "Y^{отч}"],
         ["1", "30", "10", "15", "90"],
         ["2", "35", "50", "20", "25"],
         ["3", "15", "25", "30", "60"]],
        [5.4, 2.4, 2.4, 2.4, 2.8], font=12,
        caption="Таблица 1")

    H2("Решение:")
    P("Валовой выпуск X_i = сумма потоков строки плюс Y_i. Получаем X^{отч} = (145; 130; 130). "
      "ВДС Z_j = X_j − промежуточные затраты. Зарплата – 30 % ВДС. Схема МОБ:")

    table(
        [["", "1", "2", "3", "Пром. потр.", "Y", "X"],
         ["1", "30", "10", "15", "55", "90", "145"],
         ["2", "35", "50", "20", "105", "25", "130"],
         ["3", "15", "25", "30", "70", "60", "130"],
         ["Пром. затраты", "80", "85", "65", "230", "–", "–"],
         ["з/п", "19,5", "13,5", "19,5", "–", "–", "–"],
         ["Прочие ВДС", "45,5", "31,5", "45,5", "–", "–", "–"],
         ["ВДС Z", "65", "45", "65", "175", "–", "175"],
         ["Валовые затраты X", "145", "130", "130", "405", "–", "–"]],
        [3.6, 1.8, 1.8, 1.8, 2.4, 1.6, 1.8], font=10,
        caption="Таблица 1.2 – Схема МОБ за отчётный период")

    P("Коэффициенты прямых затрат a_{ij} = x_{ij} / X_j. Матрица полных затрат B = (E − A)^{−1} "
      "считается в Excel функцией МОБР:")
    table(
        [["A", "1", "2", "3"],
         ["1", "0,2069", "0,0769", "0,1154"],
         ["2", "0,2414", "0,3846", "0,1538"],
         ["3", "0,1034", "0,1923", "0,2308"]],
        [3.0, 3.6, 3.6, 3.6], font=12,
        caption="Таблица 1.3 – Матрица прямых затрат A")

    P("Матрица полных затрат B = (E − A)^{−1} находится обращением:")
    table(
        [["B", "1", "2", "3"],
         ["1", "1,3707", "0,2513", "0,2559"],
         ["2", "0,6227", "1,8475", "0,4629"],
         ["3", "0,3400", "0,4957", "1,4501"]],
        [3.0, 3.6, 3.6, 3.6], font=12,
        caption="Таблица 1.4 – Матрица полных затрат B")

    P("Плановый выпуск X^{пл} = B Y^{пл}. В Excel (МУМНОЖ) X^{пл} ≈ (187,38; 157,92; 142,68). "
      "Плановые потоки x_{ij}^{пл} = a_{ij} X_j^{пл}:")

    table(
        [["", "1", "2", "3", "Пром. потр.", "Y^{пл}", "X^{пл}"],
         ["1", "38,77", "12,15", "16,46", "67,38", "120", "187,38"],
         ["2", "45,23", "60,74", "21,95", "127,92", "30", "157,92"],
         ["3", "19,38", "30,37", "32,93", "82,68", "60", "142,68"],
         ["Пром. затраты", "103,38", "103,25", "71,34", "277,97", "–", "–"],
         ["з/п", "25,20", "16,40", "21,40", "–", "–", "–"],
         ["Прочие ВДС", "58,80", "38,27", "49,94", "–", "–", "–"],
         ["ВДС Z", "84,00", "54,66", "71,34", "210,00", "–", "210"],
         ["Валовые затраты", "187,38", "157,92", "142,68", "488", "–", "–"]],
        [3.4, 1.9, 1.9, 1.9, 2.4, 1.8, 2.0], font=10,
        caption="Таблица 1.5 – Схема МОБ в плановом периоде")
    P("Сумма ВДС 210 совпадает с суммой конечного использования 120 + 30 + 60 = 210, "
      "что подтверждает согласованность баланса.")

    H2("Прогнозирование цен")
    P("Пункт 4. Задано p_{1} = 2, эластичность з/п k_w = 0,7. Из стоимостного баланса отраслей "
      "получаем систему относительно p_{2}, p_{3}. Решение в Excel (МОБР, МУМНОЖ): "
      "p_{2} ≈ 1,431, p_{3} ≈ 1,435.")
    P("Пункт 5. Зарплата 2-й отрасли · 1,3. Система на три индекса цен. "
      "p ≈ (1,087; 1,140; 1,075).")

    table(
        [["Сценарий", "p_{1}", "p_{2}", "p_{3}"],
         ["п. 4  (p_{1} = 2, k_w = 0,7)", "2,000", "1,431", "1,435"],
         ["п. 5  (з/п_{2} · 1,3)", "1,087", "1,140", "1,075"]],
        [6.4, 3.0, 3.0, 3.0], font=12,
        caption="Таблица 1.6 – Индексы цен")


def appendix_a():
    H1("ПРИЛОЖЕНИЕ А")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rich(p, "Расчёты в Excel", 14, bold=True)
    P("Файл «Эконометрика Верич УСР 1.xlsx» содержит три листа: «Задание» – исходные данные; "
      "«1.6 МОБ» – схемы отчётного и планового периодов, матрицы A и B, выпуск X^{пл} = B Y^{пл} "
      "(формулы MINVERSE и SUMPRODUCT); «Цены» – системы для пунктов 4 и 5. Все производные "
      "величины связаны формулами с исходными ячейками.")


def appendix_b():
    H1("ПРИЛОЖЕНИЕ Б")
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rich(p, "Конспект. Использование модели МОБ для прогнозирования цен", 14, bold=True)

    H2("1.3 Прогнозирование цен")
    P("Стоимостной состав продукции отрасли j записывается как сумма промежуточных затрат "
      "и валовой добавленной стоимости. В индексах цен p_i уравнение имеет вид")
    formula("p_j = Σ_{i=1}^{n} a_{ij} p_i + v_j,", "1.8")
    P("или в матричной форме pᵀ (E − A) = vᵀ, откуда pᵀ = vᵀ B. Здесь v_j – доля ВДС "
      "в цене продукции j-й отрасли.")
    P("Если цена продукции отрасли n_{1} увеличивается в r раз, а рост заработной платы "
      "отстаёт от роста цен (коэффициент эластичности k_w < 1), то в стоимостном балансе "
      "заработная плата умножается на k_w p_j, прочие элементы ВДС – на индекс цены своей "
      "отрасли. Индексы остальных отраслей находятся из полученной линейной системы.")
    P("Аналогично анализируется автономный рост заработной платы в одной отрасли: новая "
      "величина з/п подставляется в правую часть стоимостного баланса, индексы цен "
      "находятся как решение системы (1.7).")


SECTIONS = [
    ("title", title_page),
    ("task16", task16),
    ("appA", appendix_a),
    ("appB", appendix_b),
]


def build(pages=None, skip=(), out=FN):
    global doc
    doc = Document()
    setup_styles()
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
    doc.core_properties.title = "УСР № 1. Модель межотраслевого баланса"
    doc.save(out)
    return out


if __name__ == "__main__":
    pages = {
        "ЗАДАЧИ": 3,
        "Задача 1.6. Модель межотраслевого баланса": 3,
        "ПРИЛОЖЕНИЕ А": 7,
        "ПРИЛОЖЕНИЕ Б": 8,
    }
    print(build(pages=pages))
