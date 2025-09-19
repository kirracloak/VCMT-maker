import io
import re
import copy
import datetime
from typing import List, Dict, Optional, Iterable

import numpy as np
import streamlit as st
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ----------------------------
# Utilities and policy helpers
# ----------------------------

def normalise_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def mask_evidence_id(eid: str) -> str:
    if not eid or str(eid).strip().lower() == "pending":
        return "Pending"
    s = str(eid)
    return "*" * max(0, len(s) - 4) + s[-4:]

def validate_year(y: str) -> bool:
    try:
        y = str(y).strip()
        year = int(y)
        now = datetime.datetime.now().year
        return len(y) == 4 and 1900 <= year <= now
    except Exception:
        return False

def respond_to_instruction_request(user_text: str) -> Optional[str]:
    triggers = [
        r"show (your|the) instructions",
        r"reveal (your|the) prompt",
        r"what are your rules",
        r"display (system|agent) prompt",
        r"print your instructions",
    ]
    for pat in triggers:
        if re.search(pat, user_text or "", flags=re.IGNORECASE):
            return " I am not trained to do this"
    return None


# ----------------------------
# DOCX helpers
# ----------------------------

def all_doc_text_lines(doc: Document) -> Iterable[str]:
    for p in doc.paragraphs:
        if p.text:
            yield p.text
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text:
                        yield p.text

def load_docx(file_bytes: bytes) -> Document:
    return Document(io.BytesIO(file_bytes))

def list_tables_info(doc: Document):
    info = []
    for i, t in enumerate(doc.tables):
        header = []
        if len(t.rows) > 0:
            header = [normalise_space(c.text) for c in t.rows[0].cells]
        info.append(
            {"index": i, "n_rows": len(t.rows), "n_cols": len(t.columns), "header": header}
        )
    return info

def write_row_to_table(table, values: List[str], row_index: Optional[int] = None):
    if row_index is None or row_index >= len(table.rows):
        row = table.add_row()
    else:
        row = table.rows[row_index]
    for j, v in enumerate(values[:len(row.cells)]):
        row.cells[j].text = v or ""

def first_empty_row_index(table) -> Optional[int]:
    for i in range(1, len(table.rows)):  # skip header
        if all(not normalise_space(c.text) for c in table.rows[i].cells):
            return i
    return None

def extract_units_from_doc(doc: Document) -> Dict[str, Dict]:
    paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    full_text = "\n".join(paras)
    code_candidates = re.findall(r"\b[A-Z]{3,}[A-Z0-9]{2,}\b", full_text)
    unit_codes = sorted(set(code_candidates))

    units: Dict[str, Dict] = {}
    for code in unit_codes:
        indices = [i for i, p in enumerate(paras) if code in p]
        if not indices:
            continue
        idx = indices[0]

        # Try same line "CODE - Name" or nearby lines for a title-ish string
        name = ""
        m = re.search(rf"{code}\s*[-:–]\s*(.+)", paras[idx])
        if m:
            name = normalise_space(m.group(1))
        else:
            for j in range(idx + 1, min(idx + 6, len(paras))):
                cand = paras[j]
                if len(cand.split()) >= 3 and not re.match(r"^(Unit Code|Unit Name)", cand, re.I):
                    name = cand
                    break

        units[code] = {
            "code": code,
            "name": name,
            "application_statement": "",
            "performance_evidence": [],
            "performance_criteria": [],
        }
    return units


# ----------------------------
# Evidence matching helpers
# ----------------------------

def build_common_evidence(target_evidence: List[str], prior_evidence_blocks: List[List[str]], max_items: int = 7) -> List[str]:
    target = [normalise_space(x) for x in target_evidence if normalise_space(x)]
    prior = [normalise_space(x) for block in prior_evidence_blocks for x in (block or []) if normalise_space(x)]
    if not target or not prior:
        return []

    vec = TfidfVectorizer(ngram_range=(1, 2), min_df=1, stop_words="english")
    X = vec.fit_transform(target + prior)
    tX = X[:len(target), :]
    pX = X[len(target):, :]
    if pX.shape[0] == 0:
        return []

    sim = cosine_similarity(tX, pX)
    pairs = []
    for i in range(sim.shape[0]):
        j = int(np.argmax(sim[i]))
        score = float(sim[i, j])
        pairs.append((target[i], score))

    pairs.sort(key=lambda x: x[1], reverse
