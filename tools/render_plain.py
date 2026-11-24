# -*- coding: utf-8 -*-
"""Скриншоты листов: обычный Excel (сетка, Calibri, без оформления)."""
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets")
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONTB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

HDR_BG = "#E6E6E6"
GRID = "#D4D4D4"
HDR_LINE = "#B4B4B4"
SEL = "#217346"  # excel green tab not needed
PAD = 8
RH = 22  # header row height
CW0 = 36  # row-number column


def _font(n=11, bold=False):
    return ImageFont.truetype(FONTB if bold else FONT, n)


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "ИСТИНА" if v else "ЛОЖЬ"
    if isinstance(v, float):
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        s = f"{v:.4f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    s = str(v).replace("\n", " ").replace("\r", " ")
    return s


def render_sheet(path, sheet, out_name, max_rows=45, max_cols=12):
    wb = load_workbook(path, data_only=False)
    ws = wb[sheet]
    max_r, max_c = 1, 1
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                max_r = max(max_r, c.row)
                max_c = max(max_c, c.column)
    max_r = min(max_r, max_rows)
    max_c = min(max_c, max_cols)

    f11 = _font(11)
    dummy = Image.new("RGB", (10, 10))
    dr = ImageDraw.Draw(dummy)

    colw = []
    for c in range(1, max_c + 1):
        w = 64
        for r in range(1, max_r + 1):
            t = _txt(ws.cell(r, c).value)
            if t:
                w = max(w, int(dr.textlength(t[:48], font=f11) + 16))
        colw.append(min(w, 280))

    W = CW0 + sum(colw) + PAD
    H = RH + max_r * 20 + PAD
    img = Image.new("RGB", (W, H), "#FFFFFF")
    d = ImageDraw.Draw(img)

    # col headers
    x = CW0
    d.rectangle([0, 0, W, RH], fill=HDR_BG)
    d.rectangle([0, 0, CW0, H], fill=HDR_BG)
    d.line([(0, RH), (W, RH)], fill=HDR_LINE)
    d.line([(CW0, 0), (CW0, H)], fill=HDR_LINE)
    for i, w in enumerate(colw):
        letter = get_column_letter(i + 1)
        tw = dr.textlength(letter, font=_font(10, True))
        d.text((x + (w - tw) / 2, 4), letter, font=_font(10, True), fill="#5A5A5A")
        x += w
        d.line([(x, 0), (x, H)], fill=GRID)
    for r in range(1, max_r + 1):
        y = RH + (r - 1) * 20
        d.line([(0, y), (W, y)], fill=GRID)
        lab = str(r)
        tw = dr.textlength(lab, font=_font(10))
        d.text(((CW0 - tw) / 2, y + 3), lab, font=_font(10), fill="#5A5A5A")

    x0 = CW0
    for c in range(1, max_c + 1):
        y0 = RH
        for r in range(1, max_r + 1):
            t = _txt(ws.cell(r, c).value).replace("\n", " ")
            if t:
                d.text((x0 + 4, y0 + 3), t[:60], font=f11, fill="#000000")
            y0 += 20
        x0 += colw[c - 1]

    # outer
    d.rectangle([0, 0, W - 1, H - 1], outline="#A0A0A0")
    dest = os.path.join(OUT, out_name)
    img.save(dest, dpi=(160, 160))
    print(dest, W, H)
    return dest


def main():
    jobs = [
        (os.path.join(ROOT, "Эконометрика Верич УСР 2.xlsx"), [
            ("Задание", "xl2_task.png", 20, 8),
            ("2.4а", "xl2_a.png", 50, 8),
            ("2.4б", "xl2_b.png", 30, 8),
            ("2.4в", "xl2_c.png", 35, 8),
            ("2.4г", "xl2_d.png", 35, 8),
            ("Сводка", "xl2_sum.png", 12, 8),
        ]),
        (os.path.join(ROOT, "Эконометрика Верич УСР 5.xlsx"), [
            ("Задание", "xl_task.png", 20, 8),
            ("5.1", "xl_51.png", 48, 12),
            ("5.3", "xl_53.png", 40, 10),
            ("5.5", "xl_55.png", 40, 10),
            ("5.8", "xl_58.png", 40, 10),
            ("5.16", "xl_516.png", 40, 12),
        ]),
        (os.path.join(ROOT, "Эконометрика Верич УСР 1.xlsx"), [
            ("Задание", "xl1_task.png", 20, 8),
            ("1.6 МОБ", "xl1_mob.png", 25, 8),
            ("Цены", "xl1_cen.png", 40, 8),
        ]),
        (os.path.join(ROOT, "Эконометрика Верич УСР 3.xlsx"), [
            ("3.4", "xl3_34.png", 28, 8),
            ("3.6", "xl3_36.png", 20, 8),
            ("3.12", "xl3_312.png", 22, 10),
        ]),
    ]
    for path, sheets in jobs:
        if not os.path.exists(path):
            print("skip", path)
            continue
        for name, out, mr, mc in sheets:
            try:
                render_sheet(path, name, out, mr, mc)
            except Exception as e:
                print("fail", name, e)


if __name__ == "__main__":
    main()
