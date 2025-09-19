import io, re, streamlit as st
from typing import List, Dict, Iterable, Optional
from docx import Document

# ---------- helpers ----------
def normalise_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def all_doc_text_lines(doc: Document) -> Iterable[str]:
    # paragraphs
    for p in doc.paragraphs:
        if p.text:
            yield p.text
    # table cells
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text:
                        yield p.text

def load_docx(file_bytes: bytes) -> Document:
    return Document(io.BytesIO(file_bytes))

def extract_units_from_doc(doc: Document) -> Dict[str, Dict]:
    # scan both paragraphs and table cells
    paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    full_text = "\n".join(paras)
    # very general pattern, e.g. BSBWHS311, SITXWHS005
    codes = sorted(set(re.findall(r"\b[A-Z]{3,}[A-Z0-9]{2,}\b", full_text)))
    units: Dict[str, Dict] = {}
    for code in codes:
        # try to find a nearby name
        name = ""
        for i, line in enumerate(paras):
            if code in line:
                m = re.search(rf"{code}\s*[-:]\s*(.+)", line)
                if m:
                    name = normalise_space(m.group(1))
                else:
                   
