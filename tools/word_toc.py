# -*- coding: utf-8 -*-
"""Оглавление Word: закладки и PAGEREF (номера страниц после открытия в Word)."""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

_BM_ID = [100]


def enable_update_fields(doc):
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        settings.append(el)


def bookmark(paragraph, name):
    _BM_ID[0] += 1
    bid = str(_BM_ID[0])
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bid)
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bid)
    p = paragraph._p
    p.insert(0, start)
    p.append(end)


def _pagerefs(paragraph, name, size, bold, font, run_fn):
    r = paragraph.add_run()
    r.font.name = font
    r.font.size = Pt(size)
    r.bold = bold
    el = r._r
    b = OxmlElement("w:fldChar")
    b.set(qn("w:fldCharType"), "begin")
    i = OxmlElement("w:instrText")
    i.set(qn("xml:space"), "preserve")
    i.text = f" PAGEREF {name} \\h "
    s = OxmlElement("w:fldChar")
    s.set(qn("w:fldCharType"), "separate")
    el.append(b)
    el.append(i)
    el.append(s)
    run_fn(paragraph, "  ", size, bold)
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), "end")
    r2 = paragraph.add_run()
    r2._r.append(e)


def fill_toc(doc, items, run_fn, font="Times New Roman"):
    """items: (title, level, bookmark_name)."""
    for text, lvl, bm in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0 if lvl == 0 else 0.75)
        pf.line_spacing = 1.5
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, 1)
        run_fn(p, text, 14, bold=(lvl == 0))
        run_fn(p, "\t", 14, bold=(lvl == 0))
        _pagerefs(p, bm, 14, lvl == 0, font, run_fn)
