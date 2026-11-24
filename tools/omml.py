# -*- coding: utf-8 -*-
"""Вставка формул Word (OMML) — система со скобкой, как в примере."""
from docx.oxml import parse_xml
from docx.oxml.ns import nsmap
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _t(s):
    s = (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    return (
        f'<m:r xmlns:m="{M}" xmlns:w="{W}">'
        f'<m:rPr><m:sty m:val="p"/></m:rPr>'
        f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="28"/></w:rPr>'
        f'<m:t xml:space="preserve">{s}</m:t></m:r>'
    )


def _sub(base, idx):
    return (
        f'<m:sSub xmlns:m="{M}">'
        f'<m:e>{_t(base)}</m:e>'
        f'<m:sub>{_t(idx)}</m:sub>'
        f'</m:sSub>'
    )


def _line_tokens(parts):
    """parts: str or ('x', '1') for subscript."""
    bits = []
    for p in parts:
        if isinstance(p, tuple):
            bits.append(_sub(p[0], p[1]))
        else:
            bits.append(_t(p))
    return "".join(bits)


def _eqarr(lines):
    body = "".join(f'<m:e>{_line_tokens(ln)}</m:e>' for ln in lines)
    return f'<m:eqArr xmlns:m="{M}"><m:eqArrPr></m:eqArrPr>{body}</m:eqArr>'


def _brace(lines):
    inner = _eqarr(lines)
    return (
        f'<m:d xmlns:m="{M}">'
        f'<m:dPr><m:begChr m:val="{{"/><m:endChr m:val=""/>'
        f'<m:grow m:val="1"/></m:dPr>'
        f'<m:e>{inner}</m:e></m:d>'
    )


def _para(inner):
    return parse_xml(
        f'<m:oMathPara xmlns:m="{M}">'
        f'<m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
        f'<m:oMath>{inner}</m:oMath></m:oMathPara>'
    )


def add_math(doc, inner_xml):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.15
    p._p.append(_para(inner_xml))
    return p


def add_emm_24(doc):
    """max/min критерии + система ограничений со скобкой, как на скрине."""
    add_eqarr(doc, [
        ["max ", ("f", "1"), " = 22", ("x", "1"), " + 12", ("x", "2"), " + 14", ("x", "3"), "     (I)"],
        ["max ", ("f", "2"), " = 24", ("x", "1"), " + 16", ("x", "2"), " + 30", ("x", "3"), "     (II)"],
        ["min ", ("f", "3"), " = 20", ("x", "1"), " + 10", ("x", "2"), " + 15", ("x", "3"), "     (III)"],
    ])
    add_brace_system(doc, [
        ["2", ("x", "1"), " + 4", ("x", "2"), " + 4", ("x", "3"), " ≤ 400     (1)"],
        ["3", ("x", "1"), " + 2", ("x", "2"), " + 2", ("x", "3"), " ≤ 300     (2)"],
        ["4", ("x", "1"), " + 5", ("x", "2"), " + 3", ("x", "3"), " ≤ 500     (3)"],
        [("x", "3"), " ≥ 30     (4)"],
    ])


def add_eqarr(doc, lines):
    add_math(doc, _eqarr(lines))


def add_brace_system(doc, lines):
    add_math(doc, _brace(lines))
