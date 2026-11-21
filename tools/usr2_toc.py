# -*- coding: utf-8 -*-
import os, sys
import pymupdf
from spire.doc import Document, FileFormat
import usr2_docx as B

HEADINGS = [
    ("ЗАДАЧИ", "ЗАДАЧИ", "first"),
    ("Задача 2.4 (а). Метод последовательных уступок", "ЗАДАЧА 2.4 (а)", "first"),
    ("Задача 2.4 (б). Метод ведущего критерия", "ЗАДАЧА 2.4 (б)", "first"),
    ("Задача 2.4 (в). Равные относительные отклонения", "ЗАДАЧА 2.4 (в)", "first"),
    ("Задача 2.4 (г). Метод минимакса", "ЗАДАЧА 2.4 (г)", "first"),
    ("ПРИЛОЖЕНИЕ А", "ПРИЛОЖЕНИЕ А", "last"),
    ("ПРИЛОЖЕНИЕ Б", "ПРИЛОЖЕНИЕ Б", "last"),
]
SECTION_OF = {
    "ЗАДАЧИ": "intro",
    "Задача 2.4 (а). Метод последовательных уступок": "task24a",
    "Задача 2.4 (б). Метод ведущего критерия": "task24b",
    "Задача 2.4 (в). Равные относительные отклонения": "task24c",
    "Задача 2.4 (г). Метод минимакса": "task24d",
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
            raise RuntimeError("нет якоря")
        added = {k: v + offset for k, v in rel.items() if v and k not in res}
        res.update(added)
        if verbose:
            print(f"проход {rnd}: стр. {len(texts)}, смещ. {offset}, +{added}")
        if len(res) == len(HEADINGS):
            break
        last_anchor = [k for k, _, _ in HEADINGS if k in res][-1]
        sec = SECTION_OF[last_anchor]
        new_skip = ORDER[: ORDER.index(sec)]
        if new_skip == skip or not added:
            break
        skip = new_skip
    return res


if __name__ == "__main__":
    pages = resolve_pages()
    print(pages)
    B.build(pages=pages)
    print("saved", B.FN)
