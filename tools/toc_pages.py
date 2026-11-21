# -*- coding: utf-8 -*-
"""Двухпроходный расчёт номеров страниц для оглавления + рендер QA-страниц."""
import os
import sys
import pymupdf
from spire.doc import Document, FileFormat
import build_docx as B

HEADINGS = [
    ("ЗАДАЧИ", "ЗАДАЧИ", "first"),
    ("Задача 5.1 (е). Сетевой график проекта", "ЗАДАЧА 5.1 (е)", "first"),
    ("Задача 5.3 (а). Оптимизация по времени", "ЗАДАЧА 5.3 (а)", "first"),
    ("Задача 5.5 (а). Минимизация срока", "ЗАДАЧА 5.5 (а)", "first"),
    ("Задача 5.8 (а). Оптимизация по стоимости", "ЗАДАЧА 5.8 (а)", "first"),
    ("Задача 5.15 (а). Оптимизация по ресурсам", "ЗАДАЧА 5.15 (а)", "first"),
    ("ПРИЛОЖЕНИЕ А", "ПРИЛОЖЕНИЕ А", "last"),
    ("ПРИЛОЖЕНИЕ Б", "ПРИЛОЖЕНИЕ Б", "last"),
]
SECTION_OF = {
    "ЗАДАЧИ": "task51",
    "Задача 5.1 (е). Сетевой график проекта": "task51",
    "Задача 5.3 (а). Оптимизация по времени": "task53",
    "Задача 5.5 (а). Минимизация срока": "task55",
    "Задача 5.8 (а). Оптимизация по стоимости": "task58",
    "Задача 5.15 (а). Оптимизация по ресурсам": "task515",
    "ПРИЛОЖЕНИЕ А": "appA",
    "ПРИЛОЖЕНИЕ Б": "appB",
}
ORDER = [k for k, _ in B.SECTIONS]


def to_pdf(docx_path, pdf_path="/tmp/qa.pdf"):
    d = Document()
    d.LoadFromFile(docx_path)
    d.SaveToFile(pdf_path, FileFormat.PDF)
    d.Close()
    return pdf_path


def page_texts(pdf_path):
    doc = pymupdf.open(pdf_path)
    return [p.get_text() for p in doc]


def locate(texts, needle, mode="first", skip_pages=()):
    hits = [i + 1 for i, t in enumerate(texts)
            if needle in t and (i + 1) not in skip_pages]
    if not hits:
        return None
    return hits[0] if mode == "first" else hits[-1]


def resolve_pages(verbose=True):
    res = {}
    skip = []
    for rnd in range(8):
        path = B.build(pages=None, skip=tuple(skip), out="/tmp/qa.docx")
        texts = page_texts(to_pdf(path))
        tocpg = {i + 1 for i, t in enumerate(texts)
                 if t.count("..") > 3 or "ОГЛАВЛЕНИЕ" in t}
        rel = {k: locate(texts, n, m, tocpg) for k, n, m in HEADINGS}
        offset = 0
        anchors = [k for k, _, _ in HEADINGS if k in res and rel[k]]
        if anchors:
            a = anchors[-1]
            offset = res[a] - rel[a]
        elif skip:
            raise RuntimeError("нет якоря для пересчёта смещения")
        added = {k: v + offset for k, v in rel.items() if v and k not in res}
        res.update(added)
        n_res = len(res)
        if verbose:
            print(f"проход {rnd}: стр. {len(texts)}, смещ. {offset}, +{added}")
        if n_res == len(HEADINGS):
            break
        # держим якорем секцию последнего найденного заголовка
        last_anchor = [k for k, _, _ in HEADINGS if k in res][-1]
        sec = SECTION_OF[last_anchor]
        new_skip = ORDER[: ORDER.index(sec)]
        if new_skip == skip or not added:
            # сдвигаем окно ещё на одну секцию вперёд
            idx = ORDER.index(sec) + 1
            if idx >= len(ORDER):
                break
            new_skip = ORDER[:idx]
        skip = new_skip
    return res


if __name__ == "__main__":
    pages = resolve_pages()
    print(pages)
    out = B.build(pages=pages)
    print("saved", out)
