# -*- coding: utf-8 -*-
"""Рендер листов Excel в PNG (для приложения к отчёту)."""
import os
import re
import warnings
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from PIL import Image, ImageDraw, ImageFont
import formulas

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "Эконометрика Верич УСР 2.xlsx")
OUT = os.path.join(ROOT, "assets")
os.makedirs(OUT, exist_ok=True)

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
F_REG = FONT_DIR + "DejaVuSerif.ttf"
F_BOLD = FONT_DIR + "DejaVuSerif-Bold.ttf"

PX_PER_CHAR = 7.0        # ширина столбца Excel -> пиксели
PX_PER_PT = 1.34         # высота строки


def col_px(ws, col_letter):
    w = ws.column_dimensions[col_letter].width
    if w is None:
        w = 9.0
    return int(w * PX_PER_CHAR)


def row_px(ws, row):
    h = ws.row_dimensions[row].height
    if h is None:
        h = 14.5
    return int(h * PX_PER_PT)


def fmt_value(v, fmt):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "да" if v else "нет"
    if isinstance(v, str):
        return v
    if isinstance(v, (int, float)):
        if fmt in ("0.00", "0.0"):
            return f"{v:.2f}"
        if fmt and "#,##0" in fmt:
            return f"{v:,.0f}".replace(",", " ")
        if fmt and "%" in fmt:
            return f"{v * 100:.0f}%"
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        return f"{v:.3g}"
    return str(v)


def sheet_image(ws_name, values, out_name, scale=1.0, pad=14, max_col=None):
    wb = openpyxl.load_workbook(XLSX)
    ws = wb[ws_name]
    # границы используемой области
    max_r, max_c = 1, 1
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                max_r = max(max_r, c.row)
                max_c = max(max_c, c.column)
    for m in ws.merged_cells.ranges:
        max_r = max(max_r, m.max_row)
        max_c = max(max_c, m.max_col)
    if max_col:
        max_c = min(max_c, max_col)
    min_r, min_c = 1, 1

    colw = {c: col_px(ws, get_column_letter(c)) for c in range(min_c, max_c + 1)}
    rowh = {r: row_px(ws, r) for r in range(min_r, max_r + 1)}
    W = int(sum(colw.values()) * scale) + 2 * pad
    H = int(sum(rowh.values()) * scale) + 2 * pad
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    xs = {min_c: pad}
    for c in range(min_c, max_c + 1):
        xs[c + 1] = xs[c] + int(colw[c] * scale)
    ys = {min_r: pad}
    for r in range(min_r, max_r + 1):
        ys[r + 1] = ys[r] + int(rowh[r] * scale)

    merged = {}
    for m in ws.merged_cells.ranges:
        merged[(m.min_row, m.min_col)] = (m.max_row, m.max_col)

    def rect(r1, c1, r2, c2):
        return [xs[c1], ys[r1], xs[c2 + 1] - 1, ys[r2 + 1] - 1]

    # заливки
    for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c):
        for c in row:
            f = c.fill
            if f is not None and f.fill_type == "solid" and f.fgColor.rgb not in (None, "00000000"):
                rgb = f.fgColor.rgb[-6:]
                if rgb in ("FFFFFF", "000000") and f.fgColor.theme is not None:
                    continue
                r2, c2 = merged.get((c.row, c.column), (c.row, c.column))
                d.rectangle(rect(c.row, c.column, r2, c2), fill="#" + rgb)

    # границы и текст
    for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c):
        for c in row:
            v = values.get((ws_name, c.coordinate), c.value)
            txt = c.value if isinstance(c.value, str) and c.value.startswith("=") else ""
            txt = fmt_value(v, c.number_format)
            r2, c2 = merged.get((c.row, c.column), (c.row, c.column))
            b = c.border
            x1, y1, x2, y2 = rect(c.row, c.column, r2, c2)
            for side, coords in ((b.left, (x1, y1, x1, y2)), (b.right, (x2, y1, x2, y2)),
                                 (b.top, (x1, y1, x2, y1)), (b.bottom, (x1, y2, x2, y2))):
                if side is not None and side.style:
                    col = "#" + (side.color.rgb[-6:] if side.color and side.color.rgb else "808080")
                    w = 2 if side.style in ("medium", "thick") else 1
                    d.line(coords, fill=col, width=w)
            if not txt:
                continue
            size = max(7, int((c.font.size or 11) * 1.05 * scale))
            bold = bool(c.font.bold)
            font = ImageFont.truetype(F_BOLD if bold else F_REG, size)
            col = "#" + (c.font.color.rgb[-6:] if c.font.color and c.font.color.rgb
                         and c.font.color.rgb != "FF000000" else "000000")
            al = c.alignment
            # перенос по словам
            maxw = (x2 - x1) - 6
            words, lines, cur = txt.split(), [], ""
            for wd in words:
                test = (cur + " " + wd).strip()
                if d.textlength(test, font=font) <= maxw or not cur:
                    cur = test
                else:
                    lines.append(cur)
                    cur = wd
            if cur:
                lines.append(cur)
            lh = size + 2
            total_h = lh * len(lines)
            if al.vertical == "top":
                ty = y1 + 3
            elif al.vertical == "bottom":
                ty = y2 - total_h - 2
            else:
                ty = y1 + ((y2 - y1) - total_h) / 2
            for i, ln in enumerate(lines):
                tw = d.textlength(ln, font=font)
                if al.horizontal == "center":
                    tx = x1 + ((x2 - x1) - tw) / 2
                elif al.horizontal == "right":
                    tx = x2 - tw - 4
                else:
                    tx = x1 + 4
                d.text((tx, ty + i * lh), ln, font=font, fill=col)
    img.save(os.path.join(OUT, out_name), dpi=(200, 200))
    return os.path.join(OUT, out_name), W, H


def computed_values():
    xl = formulas.ExcelModel().loads(XLSX).finish()
    sol = xl.calculate()
    out = {}
    for k, v in sol.items():
        m = re.match(r"'\[.*\](\w[\w\.]*)'!([A-Z]+\d+)$", k)
        if not m:
            continue
        val = v.value
        try:
            val = val[0][0]
        except Exception:
            pass
        if isinstance(val, float) and abs(val - round(val)) < 1e-9:
            val = int(round(val))
        out[(m.group(1), m.group(2))] = val
    return out


def main():
    vals = computed_values()
    for name, out, scale in [("Задание", "xl2_task.png", 1.0), ("2.4а", "xl2_a.png", 1.0),
                             ("2.4б", "xl2_b.png", 1.0), ("2.4в", "xl2_c.png", 1.0),
                             ("2.4г", "xl2_d.png", 1.0), ("Сводка", "xl2_sum.png", 1.0)]:
        p, w, h = sheet_image(name, vals, out, scale=scale)
        print(p, w, h)


if __name__ == "__main__":
    main()
