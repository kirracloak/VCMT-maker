# app.py
import io
import re
import datetime
import streamlit as st
from typing import List, Dict, Iterable, Tuple
from docx import Document

# ---------- helpers ----------
def normalise_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def unique_preserve(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for it in items:
        it_n = normalise_space(it)
        if it_n and it_n not in seen:
            seen.add(it_n)
            out.append(it_n)
    return out

def all_doc_text_lines(doc: Document) -> Iterable[str]:
    # paragraphs
    for p in doc.paragraphs:
        if p.text:
            yield p.text
    # table cells
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                parts = []
                for p in cell.paragraphs:
                    if p.text:
                        if not parts or normalise_space(parts[-1]) != normalise_space(p.text):
                            parts.append(p.text)
                if parts:
                    yield " ".join(parts)

def load_docx(file_bytes: bytes) -> Document:
    return Document(io.BytesIO(file_bytes))

def extract_units_from_doc(doc: Document) -> Dict[str, Dict]:
    paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    full_text = "\n".join(paras)
    codes = sorted(set(re.findall(r"\b[A-Z]{3,}[A-Z0-9]{2,}\b", full_text)))
    units: Dict[str, Dict] = {}
    for code in codes:
        name = ""
        for i, line in enumerate(paras):
            if code in line:
                m = re.search(rf"{code}\s*[-:]\s*(.+)", line)
                if m:
                    name = normalise_space(m.group(1))
                elif i + 1 < len(paras) and len(paras[i + 1].split()) >= 3:
                    name = paras[i + 1]
                break
        units[code] = {"code": code, "name": name}
    return units

def list_tables_info(doc: Document) -> List[str]:
    out = []
    for i, t in enumerate(doc.tables):
        header = []
        if len(t.rows) > 0:
            header = [normalise_space(c.text) for c in t.rows[0].cells]
        out.append(f"Table {i}: {len(t.rows)} rows x {len(t.columns)} cols | header: {' | '.join(header[:6]) if header else '(none)'}")
    return out

def find_part1_table_index(doc: Document) -> int:
    # simple heuristic: first table with >= 4 columns
    for i, t in enumerate(doc.tables):
        if len(t.columns) >= 4:
            return i
    return -1

def get_section_text_for_unit(doc: Document, unit_code: str) -> Tuple[str, List[str]]:
    # heuristic extraction of application excerpt and performance evidence bullets
    paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    idxs = [i for i, p in enumerate(paras) if unit_code in p]
    app_excerpt = ""
    perf_bullets: List[str] = []

    if idxs:
        start = max(0, idxs[0] - 6)
        window = paras[start:start + 20]
    else:
        window = paras[:40]

    for i, line in enumerate(window):
        if re.search(r"Application Statement|Application of the unit|Application of skill", line, re.IGNORECASE):
            for j in range(i + 1, min(i + 4, len(window))):
                if len(window[j].split()) >= 4:
                    app_excerpt = window[j]
                    break
            break
    if not app_excerpt:
        for line in window:
            if len(line.split()) >= 6 and "performance" not in line.lower():
                app_excerpt = line
                break

    all_paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    for i, p in enumerate(all_paras):
        if re.search(r"Performance Evidence", p
