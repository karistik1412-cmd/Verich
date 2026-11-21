# -*- coding: utf-8 -*-
"""Рисунки для отчёта: сетевые графики, графики Ганта, эпюры ресурсов."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyArrowPatch, Rectangle
import matplotlib.font_manager as fm

import data as D

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family": "DejaVu Serif",
    "font.size": 11,
    "mathtext.fontset": "dejavuserif",
})
NODE_EDGE = "#1F3864"
NODE_FILL = "#FFFFFF"
CRIT = "#C00000"
BAR = "#2E75B6"
BAR2 = "#A6A6A6"


def net_fig(pos, arcs, n, fname, crit=(), dummies=(), labels=None, figsize=(9, 4.4),
            title=None, node_r=0.26, label_off=0.17, fontsize=10.5):
    fig, ax = plt.subplots(figsize=figsize)
    for (i, j, txt) in arcs:
        x1, y1 = pos[i]
        x2, y2 = pos[j]
        rad = 0.0
        if (i, j) in dummies:
            rad = 0.0
        style = "--" if (i, j) in dummies else "-"
        col = CRIT if (i, j) in crit else NODE_EDGE
        ar = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=13,
                             lw=1.5, color=col, linestyle=style,
                             shrinkA=node_r * 72, shrinkB=node_r * 72,
                             connectionstyle=f"arc3,rad={rad}")
        ax.add_patch(ar)
        if txt:
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            dx, dy = x2 - x1, y2 - y1
            ln = max((dx ** 2 + dy ** 2) ** 0.5, 1e-6)
            # сдвиг подписи по нормали
            ax.text(mx - dy / ln * label_off, my + dx / ln * label_off, txt,
                    ha="center", va="center", fontsize=fontsize, color=col,
                    bbox=dict(boxstyle="round,pad=0.14", fc="white", ec="none", alpha=0.9))
    for k, (x, y) in pos.items():
        ax.add_patch(Circle((x, y), node_r, facecolor=NODE_FILL, edgecolor=NODE_EDGE, lw=1.6, zorder=3))
        ax.text(x, y, str(k), ha="center", va="center", fontsize=12, zorder=4)
    ax.set_aspect("equal")
    ax.axis("off")
    if title:
        ax.set_title(title, fontsize=12)
    xs = [p[0] for p in pos.values()]
    ys = [p[1] for p in pos.values()]
    ax.set_xlim(min(xs) - 0.55, max(xs) + 0.55)
    ax.set_ylim(min(ys) - 0.55, max(ys) + 0.55)
    fig.tight_layout(pad=0.2)
    fig.savefig(os.path.join(OUT, fname), dpi=220, facecolor="white")
    plt.close(fig)


def gantt(rows, T, fname, title=None, figsize=(9, 4.2), res=None, res_max=None,
          xlabel="t, дни", colors=None):
    """rows: (label, start, dur, critical_or_None)"""
    n = len(rows)
    nrows = 2 if res is not None else 1
    ratios = [3.2, 1.3] if res is not None else [1]
    fig, axes = plt.subplots(nrows, 1, figsize=figsize,
                             gridspec_kw={"height_ratios": ratios}, squeeze=False)
    ax = axes[0][0]
    for k, (lab, st, du, crit) in enumerate(rows):
        y = n - k
        c = CRIT if crit else BAR
        ax.add_patch(Rectangle((st, y - 0.32), du, 0.64, facecolor=c, edgecolor=NODE_EDGE, lw=0.8))
        ax.text(st + du / 2, y, f"{du:g}", ha="center", va="center", color="white", fontsize=9)
    ax.set_yticks([n - k for k in range(n)])
    ax.set_yticklabels([r[0] for r in rows], fontsize=10)
    ax.set_xlim(0, T)
    ax.set_ylim(0.3, n + 0.7)
    ax.set_xticks(range(0, int(T) + 1))
    ax.grid(axis="x", color="#D9D9D9", lw=0.7)
    ax.set_axisbelow(True)
    ax.set_xlabel(xlabel, fontsize=9)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    if title:
        ax.set_title(title, fontsize=11)
    if res is not None:
        ax2 = axes[1][0]
        ax2.bar(range(len(res)), res, width=0.72, color="#8FAADC", edgecolor=NODE_EDGE, lw=0.7)
        ax2.axhline(res_max, color=CRIT, lw=1.3, ls="--")
        ax2.text(len(res) - 0.4, res_max + 0.4, f"R = {res_max}", color=CRIT, ha="right", fontsize=9)
        ax2.set_xticks(range(len(res)))
        ax2.set_xticklabels([str(i + 1) for i in range(len(res))], fontsize=8)
        ax2.set_ylabel("ресурс, ед.", fontsize=9)
        ax2.set_xlabel("дни", fontsize=9)
        for s in ("top", "right"):
            ax2.spines[s].set_visible(False)
        ax2.set_ylim(0, max(max(res), res_max) + 2)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, fname), dpi=220, facecolor="white")
    plt.close(fig)


def main():
    # ---------- 5.1 (е)
    pos51 = {1: (0, 1.2), 2: (1.6, 2.4), 3: (1.6, 0.15), 4: (3.1, 1.25),
             5: (3.1, 3.05), 6: (4.7, 1.85), 7: (4.7, 0.45), 8: (6.2, 1.55), 9: (7.6, 1.15)}
    arcs51 = [(1, 2, "1"), (1, 3, "5"), (2, 4, "3"), (2, 5, "2"), (3, 4, ""),
              (4, 6, "6"), (4, 7, "5"), (5, 8, ""), (6, 8, ""), (7, 9, "3"), (8, 9, "5")]
    net_fig(pos51, arcs51, 9, "net51.png", crit={(1, 3), (3, 4), (4, 6), (6, 8), (8, 9)},
            dummies={(3, 4), (5, 8), (6, 8)}, figsize=(9.2, 4.6))

    tp, tn, T, rows = D.params51()
    gantt([(r["name"], r["es"], r["t"], r["rp"] == 0) for r in rows], T, "gantt51.png",
          title="", figsize=(9, 3.6))

    # ---------- 5.3
    pos53 = {1: (0, 1.5), 2: (2.1, 2.9), 3: (2.1, 0.15), 4: (4.2, 1.5),
             5: (6.3, 2.9), 6: (8.4, 1.5)}
    arcs53 = [(1, 2, "9"), (1, 3, "12"), (1, 4, "18"), (2, 4, "8"), (2, 5, "12"),
              (3, 4, "5"), (3, 6, "12"), (5, 6, "12"), (4, 5, "10"), (4, 6, "13")]
    net_fig(pos53, arcs53, 6, "net53.png",
            crit={(1, 4), (4, 5), (5, 6)}, figsize=(9.2, 4.2))

    # ---------- 5.5
    pos55 = {1: (0, 1.5), 2: (2.1, 0.3), 3: (4.2, 1.5), 4: (6.3, 2.7), 5: (8.4, 1.5)}
    arcs55 = [(1, 2, "10"), (1, 3, "18"), (1, 4, "16"), (2, 3, "12"),
              (3, 4, "7"), (3, 5, "13"), (4, 5, "11")]
    net_fig(pos55, arcs55, 5, "net55.png", crit={(1, 2), (2, 3), (3, 4), (4, 5)},
            figsize=(9.2, 4.0))

    # ---------- 5.8
    pos58 = {1: (0, 1.5), 2: (2.2, 2.6), 3: (2.2, 0.3), 4: (4.8, 1.5), 5: (7.4, 2.6)}
    arcs58 = [(1, 2, "14, 8"), (1, 3, "20, 12"), (2, 4, "14, 8"), (2, 5, "10, 6"),
              (3, 4, "12, 7"), (4, 5, "6, 4"), (2, 3, "")]
    net_fig(pos58, arcs58, 5, "net58.png", crit={(1, 3), (3, 4), (4, 5)},
            dummies={(2, 3)}, figsize=(9.2, 4.0))

    # ---------- 5.16 (а)
    pos516 = {1: (0, 1.5), 2: (2.5, 2.9), 3: (2.5, 0.1), 4: (5.0, 1.5), 5: (7.6, 1.5)}
    arcs516 = [(1, 2, "6, 4"), (1, 4, "4, 3"), (1, 3, "6, 5"), (2, 5, "5, 3"),
               (3, 4, "4, 4"), (4, 5, "7, 6")]
    net_fig(pos516, arcs516, 5, "net516.png", crit={(1, 3), (3, 4), (4, 5)},
            figsize=(9.2, 4.2))

    crit_idx = {1, 3, 5}  # (1,3), (3,4), (4,5)
    rows_es = [(f"{w[0]} ({w[3]}; {w[4]})", D.ST516_ES[k], w[3], k in crit_idx)
               for k, w in enumerate(D.W516)]
    gantt(rows_es, 17, "gantt516_a.png", res=D.PROF_ES, res_max=9, figsize=(9.4, 5.4))
    rows_opt = [(f"{w[0]} ({w[3]}; {w[4]})", D.ST516_OPT[k], w[3], k in crit_idx)
                for k, w in enumerate(D.W516)]
    gantt(rows_opt, 17, "gantt516_b.png", res=D.PROF_OPT, res_max=9, figsize=(9.4, 5.4))
    print("готово:", sorted(os.listdir(OUT)))


if __name__ == "__main__":
    main()
