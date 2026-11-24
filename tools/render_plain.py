# -*- coding: utf-8 -*-
"""Скриншоты фрагментов листа: сетка Excel, без оформления."""
import ast
import os
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets")
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONTB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

HDR_BG = "#E6E6E6"
GRID = "#D4D4D4"
HDR_LINE = "#B4B4B4"
PAD = 8
RH = 22
CW0 = 36

_CELL = re.compile(r"\b([A-Z]+)(\d+)\b")
_RANGE = re.compile(r"\b([A-Z]+)(\d+):([A-Z]+)(\d+)\b")


def _font(n=11, bold=False):
    return ImageFont.truetype(FONTB if bold else FONT, n)


def _fmt_num(v):
    if isinstance(v, bool):
        return "ИСТИНА" if v else "ЛОЖЬ"
    if isinstance(v, float):
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        s = f"{v:.4f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    return str(v).replace("\n", " ").replace("\r", " ")


def _eval_formula(ws, formula, cache):
    expr = formula[1:].strip()

    def rng(m):
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        i1, i2 = column_index_from_string(c1), column_index_from_string(c2)
        vals = []
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(i1, i2), max(i1, i2) + 1):
                v = _value(ws, r, c, cache)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    vals.append(v)
        return str(vals) if vals else "[]"

    expr = _RANGE.sub(rng, expr)

    def cell(m):
        c, r = m.group(1), int(m.group(2))
        v = _value(ws, r, column_index_from_string(c), cache)
        if v is None:
            return "0"
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return repr(v)
        return "0"

    expr = _CELL.sub(cell, expr)
    expr = expr.replace("MAX", "max").replace("MIN", "min").replace("ABS", "abs")
    try:
        return ast.literal_eval(expr) if re.fullmatch(r"[-0-9.]+", expr) else eval(
            expr, {"__builtins__": {}}, {"max": max, "min": min, "abs": abs}
        )
    except Exception:
        return formula


def _value(ws, r, c, cache):
    key = (r, c)
    if key in cache:
        return cache[key]
    cache[key] = None
    v = ws.cell(r, c).value
    if isinstance(v, str) and v.startswith("="):
        v = _eval_formula(ws, v, cache)
    cache[key] = v
    return v


def _txt(ws, r, c, cache):
    v = _value(ws, r, c, cache)
    if v is None:
        return ""
    return _fmt_num(v)


def render_range(path, sheet, out_name, r1, r2, c1=1, c2=None):
    wb = load_workbook(path, data_only=False)
    ws = wb[sheet]
    if c2 is None:
        c2 = 1
        for row in ws.iter_rows(min_row=r1, max_row=r2):
            for cell in row:
                if cell.value is not None:
                    c2 = max(c2, cell.column)
    cache = {}
    f11 = _font(11)
    dummy = Image.new("RGB", (10, 10))
    dr = ImageDraw.Draw(dummy)

    colw = []
    for c in range(c1, c2 + 1):
        w = 56
        for r in range(r1, r2 + 1):
            t = _txt(ws, r, c, cache)
            if t:
                w = max(w, int(dr.textlength(t[:48], font=f11) + 16))
        colw.append(min(w, 260))

    n_rows = r2 - r1 + 1
    W = CW0 + sum(colw) + PAD
    H = RH + n_rows * 20 + PAD
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, RH], fill=HDR_BG)
    d.rectangle([0, 0, CW0, H], fill=HDR_BG)
    d.line([(0, RH), (W, RH)], fill=HDR_LINE)
    d.line([(CW0, 0), (CW0, H)], fill=HDR_LINE)
    x = CW0
    for i, w in enumerate(colw):
        letter = get_column_letter(c1 + i)
        tw = dr.textlength(letter, font=_font(10, True))
        d.text((x + (w - tw) / 2, 4), letter, font=_font(10, True), fill="#5A5A5A")
        x += w
        d.line([(x, 0), (x, H)], fill=GRID)
    for i, r in enumerate(range(r1, r2 + 1)):
        y = RH + i * 20
        d.line([(0, y), (W, y)], fill=GRID)
        lab = str(r)
        tw = dr.textlength(lab, font=_font(10))
        d.text(((CW0 - tw) / 2, y + 3), lab, font=_font(10), fill="#5A5A5A")

    x0 = CW0
    for j, c in enumerate(range(c1, c2 + 1)):
        y0 = RH
        for r in range(r1, r2 + 1):
            t = _txt(ws, r, c, cache)
            if t:
                d.text((x0 + 4, y0 + 3), t[:56], font=f11, fill="#000000")
            y0 += 20
        x0 += colw[j]

    d.rectangle([0, 0, W - 1, H - 1], outline="#A0A0A0")
    dest = os.path.join(OUT, out_name)
    img.save(dest, dpi=(160, 160))
    print(dest, W, H)
    return dest


def main():
    x2 = os.path.join(ROOT, "Эконометрика Верич УСР 2.xlsx")
    jobs = [
        (x2, "Задание", "xl2_task1.png", 4, 8, 1, 5),
        (x2, "Задание", "xl2_task2.png", 10, 14, 1, 5),
        (x2, "2.4а", "xl2_a1.png", 4, 18, 1, 5),
        (x2, "2.4а", "xl2_a2.png", 20, 34, 1, 5),
        (x2, "2.4а", "xl2_a3.png", 36, 51, 1, 5),
        (x2, "2.4б", "xl2_b1.png", 4, 11, 1, 4),
        (x2, "2.4б", "xl2_b2.png", 13, 20, 1, 5),
        (x2, "2.4в", "xl2_c1.png", 4, 8, 1, 5),
        (x2, "2.4в", "xl2_c2.png", 10, 22, 1, 4),
        (x2, "2.4в", "xl2_c3.png", 24, 29, 1, 5),
        (x2, "2.4г", "xl2_d1.png", 4, 8, 1, 5),
        (x2, "2.4г", "xl2_d2.png", 10, 22, 1, 4),
        (x2, "2.4г", "xl2_d3.png", 24, 29, 1, 5),
        (x2, "Сводка", "xl2_sum.png", 1, 7, 1, 7),
    ]
    for path, sheet, out, r1, r2, c1, c2 in jobs:
        render_range(path, sheet, out, r1, r2, c1, c2)


if __name__ == "__main__":
    main()
