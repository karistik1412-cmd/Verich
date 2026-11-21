# -*- coding: utf-8 -*-
"""5.8 (а) — алгоритм КДЗ; 5.15 (а) — оптимизация по ресурсам (проверка перебором)."""
import itertools

# ============================================================ 5.8 (а)
W58 = [(1, 2), (1, 3), (2, 4), (2, 5), (3, 4), (4, 5)]
D58 = [14, 20, 14, 10, 12, 6]
d58 = [8, 12, 8, 6, 7, 4]
c58 = [160, 120, 35, 60, 72, 110]
C58 = [190, 176, 95, 84, 112, 144]
H58 = [(C58[i] - c58[i]) / (D58[i] - d58[i]) for i in range(6)]
EDGES = [(1, 2), (1, 3), (2, 3), (2, 4), (2, 5), (3, 4), (4, 5)]
PATHS58 = [[(1, 2), (2, 3), (3, 4), (4, 5)], [(1, 2), (2, 4), (4, 5)],
           [(1, 2), (2, 5)], [(1, 3), (3, 4), (4, 5)]]


def tp_tn(nodes, arcs):
    tp = [0] * (nodes + 1)
    for j in range(2, nodes + 1):
        inc = [a for a in arcs if a[1] == j]
        tp[j] = max(tp[a[0]] + a[2] for a in inc)
    T = tp[nodes]
    tn = [T] * (nodes + 1)
    for i in range(nodes - 1, 0, -1):
        out = [a for a in arcs if a[0] == i]
        tn[i] = min(tn[a[1]] - a[2] for a in out)
    return tp, tn, T


def arcs58(dur):
    return [(W58[i][0], W58[i][1], dur[i]) for i in range(6)] + [(2, 3, 0)]


def plen(path, dur):
    s = 0
    for e in path:
        if e == (2, 3):
            continue
        s += dur[W58.index(e)]
    return s


def greedy58(verbose=True):
    dur = list(D58)
    cost = sum(c58)
    steps = []
    for step in range(20):
        tp, tn, T = tp_tn(5, arcs58(dur))
        rp = [tn[b] - tp[a] - dur[i] for i, (a, b) in enumerate(W58)]
        crit = [i for i in range(6) if rp[i] == 0]
        cpaths = [p for p in PATHS58 if plen(p, dur) == T]
        if verbose:
            print(f"\nшаг {step}: T={T} dur={dur} Rп={rp} C={cost}")
            print("   критические пути:", [[e for e in p if e != (2, 3)] for p in cpaths])
        fully = [p for p in cpaths if all(dur[W58.index(e)] == d58[W58.index(e)]
                                          for e in p if e != (2, 3))]
        if fully:
            if verbose:
                print("   критический путь полностью сокращён — СТОП")
            break
        # работы с минимальным КДЗ на каждом критическом пути
        perm = []
        for p in cpaths:
            real = [e for e in p if e != (2, 3) and dur[W58.index(e)] > d58[W58.index(e)]]
            perm.append(min(real, key=lambda e: H58[W58.index(e)]))
        common = [e for e in W58 if e != (2, 3) and dur[W58.index(e)] > d58[W58.index(e)]
                  and all(e in p for p in cpaths)]
        if not common:
            sel = sorted(set(perm))
            reason = "критические пути не имеют общих работ: на каждом выбираем работу с минимальным КДЗ"
        else:
            hc = min(common, key=lambda e: H58[W58.index(e)])
            ssum = sum(H58[W58.index(e)] for e in set(perm))
            if ssum > H58[W58.index(hc)] + 1e-12:
                sel = [hc]
                reason = (f"сумма КДЗ {ssum:g} > КДЗ общей работы {H58[W58.index(hc)]:g} -> "
                          f"сокращаем общую работу ({hc[0]},{hc[1]})")
            elif abs(ssum - H58[W58.index(hc)]) < 1e-12:
                sel = sorted(set(perm) | {hc})
                reason = "сумма КДЗ работ равна КДЗ общей работы -> сокращаем все"
            else:
                sel = sorted(set(perm))
                reason = (f"сумма КДЗ {ssum:g} < КДЗ общей работы {H58[W58.index(hc)]:g}")
        # величина сокращения
        delta = min(dur[W58.index(e)] - d58[W58.index(e)] for e in sel)
        # пока не появится новый критический путь
        for p in PATHS58:
            if plen(p, dur) == T:
                continue
            red = sum(1 for e in sel if e in p)
            slack = T - plen(p, dur)
            if red < len(sel):
                delta = min(delta, slack)
        delta = int(delta)
        for e in sel:
            dur[W58.index(e)] -= delta
            cost += delta * H58[W58.index(e)]
        steps.append(dict(step=step + 1, sel=[f"({e[0]},{e[1]})" for e in sel], delta=delta,
                          T=None, reason=reason, dur=list(dur), cost=cost))
        if verbose:
            print(f"   сокращаем {[f'({e[0]},{e[1]})' for e in sel]} на {delta} ({reason})")
    tp, tn, T = tp_tn(5, arcs58(dur))
    if verbose:
        print(f"\nИТОГ 5.8: T_min = {T}; продолжительности {dur}; "
              f"прирост стоимости {cost - sum(c58)}; общая стоимость {cost}")
        print("t_p:", tp, "t_п:", tn)
    return T, dur, cost, steps


# ============================================================ 5.15 (а)
W515 = [(1, 2, 6, 4), (1, 3, 4, 5), (1, 4, 6, 3), (2, 3, 2, 4), (2, 5, 3, 3),
        (3, 5, 4, 2), (4, 5, 3, 4)]
R515 = 9
IDX = {(w[0], w[1]): k for k, w in enumerate(W515)}


def profile(starts, T=None):
    if T is None:
        T = max(starts[k] + W515[k][2] for k in range(len(W515)))
    out = []
    for x in range(T):
        out.append(sum(w[3] for k, w in enumerate(W515) if starts[k] <= x < starts[k] + w[2]))
    return out


def feasible(starts):
    """проверка предшествования"""
    for k, (a, b, t, r) in enumerate(W515):
        for m, (a2, b2, t2, r2) in enumerate(W515):
            if b == a2 and starts[m] < starts[k] + t:
                return False
    return True


def brute_force_515():
    n = len(W515)
    best = None
    for T in range(12, 20):
        for combo in itertools.product(range(T), repeat=n):
            st = list(combo)
            if max(st[k] + W515[k][2] for k in range(n)) > T:
                continue
            if not feasible(st):
                continue
            if max(profile(st)) <= R515:
                best = (T, st)
                print("минимум T =", T, "старты:", st, "профиль:", profile(st))
                return best
    return best


if __name__ == "__main__":
    print("==== 5.8 ====")
    greedy58()
    print("\n==== 5.15: перебор ====")
    brute_force_515()
