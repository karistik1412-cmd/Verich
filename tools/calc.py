# -*- coding: utf-8 -*-
"""Расчёты к отчёту УСР №5 (тема 5, сетевые модели)."""
import itertools, math
import numpy as np
from scipy.optimize import linprog

# ---------------------------------------------------------------- 5.1 (е)
# работы: (имя, начало, конец, длительность)
G51 = [
    ("a1", 1, 2, 1),
    ("a2", 1, 3, 5),
    ("a3", 2, 4, 3),
    ("a4", 2, 5, 2),
    ("a5", 4, 6, 6),
    ("a6", 4, 7, 5),
    ("a7", 8, 9, 5),
    ("a8", 7, 9, 3),
]
DUMMY51 = [("d1", 3, 4, 0), ("d2", 5, 8, 0), ("d3", 6, 8, 0)]
N51 = 9


def net_params(nodes, arcs, crit_only=False):
    """t_p, t_п событий (arcs = (i,j,t))."""
    tp = {i: 0 for i in range(1, nodes + 1)}
    order = sorted(tp)
    for j in order:
        inc = [a for a in arcs if a[1] == j]
        if inc:
            tp[j] = max(tp[a[0]] + a[2] for a in inc)
    T = tp[nodes]
    tn = {i: T for i in range(1, nodes + 1)}
    for i in reversed(order):
        out = [a for a in arcs if a[0] == i]
        if out:
            tn[i] = min(tn[a[1]] - a[2] for a in out)
    return tp, tn, T


def work_params(arcs):
    nodes = max(max(a[2] for a in arcs), max(a[1] for a in arcs))
    tp, tn, T = net_params(nodes, [(a[1], a[2], a[3]) for a in arcs])
    res = []
    for (name, i, j, t) in arcs:
        rp = tn[j] - tp[i] - t          # полный резерв
        rs = tp[j] - tp[i] - t          # свободный резерв
        rn = max(0, tp[j] - tn[i] - t)  # независимый резерв
        res.append(dict(name=name, i=i, j=j, t=t,
                        es=tp[i], ef=tp[i] + t,
                        ls=tn[j] - t, lf=tn[j],
                        rp=rp, rs=rs, rn=rn))
    return tp, tn, T, res


def report51():
    arcs = [(n, i, j, t) for (n, i, j, t) in G51] + [(n, i, j, t) for (n, i, j, t) in DUMMY51]
    tp, tn, T, res = work_params(arcs)
    print("=== 5.1(е) ===")
    print("t_p:", tp)
    print("t_п:", tn)
    print("T кр =", T)
    for r in res:
        print(f"{r['name']:>3} ({r['i']},{r['j']}) t={r['t']:>2}  "
              f"{r['es']:>3} {r['ef']:>3} {r['ls']:>3} {r['lf']:>3}  "
              f"Rп={r['rp']:>3} Rс={r['rs']:>3} Rн={r['rn']:>3}")
    return tp, tn, T, res


# ---------------------------------------------------------------- 5.3 (а)
W53 = [(1, 2), (1, 3), (1, 4), (2, 4), (2, 5), (3, 4), (3, 6), (4, 5), (4, 6), (5, 6)]
T53 = [9, 12, 18, 8, 12, 5, 12, 10, 13, 12]
D53 = [7, 10, 15, 6, 10, 3, 8, 7, 12, 10]
K53 = [0.05, 0.20, 0.25, 0.08, 0.15, 0.10, 0.06, 0.05, 0.10, 0.50]
T0_53 = 35


def lp_task53():
    n = len(W53)
    # переменные: x1..xn, t_н событий y1..y6 (y1=0), t_о событий z?  используем
    # y_i - время свершения события i (раннее), z - фиктивные? Проще:
    # переменные: x(10), s_1..s_6 (времена свершения событий)
    # min sum x
    c = np.zeros(n + 6)
    c[:n] = 1
    A, b = [], []
    # s_j - s_i + k*x >= t  ->  -s_j + s_i - k*x <= -t
    for idx, (i, j) in enumerate(W53):
        row = np.zeros(n + 6)
        row[n + i - 1] += 1
        row[n + j - 1] -= 1
        row[idx] -= K53[idx]
        A.append(row); b.append(-T53[idx])
    # s_6 <= t0
    row = np.zeros(n + 6); row[n + 5] = 1
    A.append(row); b.append(T0_53)
    # x <= (t-d)/k
    bounds = []
    for idx in range(n):
        row = np.zeros(n + 6); row[idx] = 1
        A.append(row); b.append((T53[idx] - D53[idx]) / K53[idx])
        bounds.append((0, None))
    for i in range(6):
        bounds.append((0, None))
    res = linprog(c, A_ub=np.array(A), b_ub=np.array(b), bounds=bounds, method="highs")
    x = res.x[:n]
    s = res.x[n:]
    dur = [T53[i] - K53[i] * x[i] for i in range(n)]
    print("\n=== 5.3(а) ===")
    print("f min =", res.fun)
    print("события:", np.round(s, 4))
    for i, (a, bb) in enumerate(W53):
        print(f"({a},{bb}): t={T53[i]:>2} d={D53[i]:>2} k={K53[i]:>4} "
              f"x={x[i]:8.4f} -> t'={dur[i]:7.4f}  "
              f"tн={s[a-1]:7.3f} tо={s[bb-1]:7.3f}")
    return x, s, dur


# ---------------------------------------------------------------- 5.5 (а)
W55 = [(1, 2), (1, 3), (1, 4), (2, 3), (3, 4), (3, 5), (4, 5)]
T55 = [10, 18, 16, 12, 7, 13, 11]
D55 = [7, 14, 12, 10, 5, 9, 8]
K55 = [0.50, 0.10, 0.25, 0.40, 0.20, 0.15, 0.30]
B55 = 42


def lp_task55():
    n = len(W55)
    c = np.zeros(n + 6)   # x(7), s1..s5, T
    c[-1] = 1
    A, b = [], []
    # s_j - s_i + k*x >= t
    for idx, (i, j) in enumerate(W55):
        row = np.zeros(n + 6)
        row[n + i - 1] += 1
        row[n + j - 1] -= 1
        row[idx] -= K55[idx]
        A.append(row); b.append(-T55[idx])
    # s5 = T ;  s1 = 0
    row = np.zeros(n + 6); row[n + 4] = 1; row[-1] = -1
    A.append(row); b.append(0)
    # budget
    row = np.zeros(n + 6); row[:n] = 1
    A.append(row); b.append(B55)
    bounds = [(0, (T55[i] - D55[i]) / K55[i]) for i in range(n)] + [(0, None)] * 5 + [(0, None)]
    bounds[n] = (0, 0)
    res = linprog(c, A_ub=np.array(A), b_ub=np.array(b), bounds=bounds, method="highs")
    x = res.x[:n]; s = res.x[n:n + 5]; T = res.x[-1]
    dur = [T55[i] - K55[i] * x[i] for i in range(n)]
    print("\n=== 5.5(а) ===")
    print("T min =", T, " sum x =", sum(x))
    for i, (a, bb) in enumerate(W55):
        print(f"({a},{bb}): t={T55[i]:>2} d={D55[i]:>2} k={K55[i]:>4} "
              f"x={x[i]:8.4f} -> t'={dur[i]:7.4f}  "
              f"tн={s[a-1]:7.3f} tо={s[bb-1]:7.3f}")
    return x, s, T, dur


# ---------------------------------------------------------------- 5.8 (а)
W58 = [(1, 2), (1, 3), (2, 4), (2, 5), (3, 4), (4, 5)]
DD58 = [14, 20, 14, 10, 12, 6]
dd58 = [8, 12, 8, 6, 7, 4]
cc58 = [160, 120, 35, 60, 72, 110]
CC58 = [190, 176, 95, 84, 112, 144]
FICT = [(2, 3, 0)]


def arcs58(dur):
    a = [(W58[i][0], W58[i][1], dur[i]) for i in range(len(W58))]
    return a + FICT


def paths58():
    """все полные пути 1..5"""
    edges = {(1, 2), (1, 3), (2, 4), (2, 5), (3, 4), (4, 5), (2, 3)}
    out = []
    def dfs(node, path):
        if node == 5:
            out.append(path[:]); return
        for (a, b) in sorted(edges):
            if a == node:
                path.append((a, b)); dfs(b, path); path.pop()
    dfs(1, [])
    return out


def report58():
    print("\n=== 5.8(а) ===")
    h = [(CC58[i] - cc58[i]) / (DD58[i] - dd58[i]) for i in range(6)]
    print("КДЗ:", h)
    tp, tn, T = net_params(5, arcs58(DD58))
    print("t_p:", tp, "t_п:", tn, "Tкр=", T)
    # полные резервы
    for i, (a, b) in enumerate(W58):
        print(f"({a},{b}) D={DD58[i]} d={dd58[i]} c={cc58[i]} C={CC58[i]} h={h[i]} "
              f"Rп={tn[b]-tp[a]-DD58[i]}")
    print("пути:", [(p, sum(dict(((W58[i][0], W58[i][1]), DD58[i]) for i in range(6)).get(e, 0) for e in p)) for p in paths58()])
    # LP: минимум стоимости при T = Tmin(min по d)
    tp2, tn2, Tmin = net_params(5, arcs58(dd58))
    print("минимально возможное T =", Tmin)
    n = 6
    c = np.zeros(n + 5)
    for i in range(n):
        c[i] = -h[i]          # cost = sum(c_i + h_i*(D_i - t_i)) -> min => max sum h*t
    A, b = [], []
    for i, (a_, b_) in enumerate(W58):
        row = np.zeros(n + 5); row[n + a_ - 1] += 1; row[n + b_ - 1] -= 1; row[i] += 1
        A.append(row); b.append(0)
    row = np.zeros(n + 5); row[n + 1] -= 1; row[n + 2] += 1
    A.append(row); b.append(0)
    row = np.zeros(n + 5); row[n + 4] = 1
    A.append(row); b.append(Tmin)
    bounds = [(dd58[i], DD58[i]) for i in range(n)] + [(0, None)] * 5
    bounds[n] = (0, 0)
    res = linprog(c, A_ub=np.array(A), b_ub=np.array(b), bounds=bounds, method="highs")
    t = res.x[:n]; s = res.x[n:]
    cost = sum(cc58[i] + h[i] * (DD58[i] - t[i]) for i in range(n))
    print("опт. продолжительности:", np.round(t, 4))
    print("события:", np.round(s, 3))
    print("стоимость:", cost, " прирост:", cost - sum(cc58))
    return h, t, s, Tmin, cost


if __name__ == "__main__":
    report51()
    lp_task53()
    lp_task55()
    report58()
