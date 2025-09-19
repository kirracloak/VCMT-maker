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

def all_doc_text_lines(doc):
    # Yield text from both paragraphs and every table cell
    for p in doc.paragraphs:
        if p.text:
            yield p.text
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text:
                        yield p.text

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
full_text = "\n".join(all_doc_text_lines(doc))
full_text_up = full_text.upper()
validated_codes = [c for c in user_unit_codes if c.strip().upper() in full_text_up]
    units: Dict[str, Dict] = {}
    for code in unit_codes:
        indices = [i for i, p in enumerate(paras) if code in p]
        if not indices:
            continue
        idx = indices[0]
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
    pairs.sort(key=lambda x: x[1], reverse=True)
    out = []
    for t, _ in pairs:
        if t not in out:
            out.append(t)
        if len(out) >= max_items:
            break
    return out

def construct_part1_statement(application_statement: str, bullets: List[str]) -> str:
    app = (application_statement or "").strip().rstrip(".")
    if app:
        intro = "Within this qualification, I was required to demonstrate competency in the skills and knowledge required to " + app
    else:
        intro = "Within this qualification, I was required to demonstrate competency in the relevant skills and knowledge"
    prefix = "Specifically relevant were the following course components:"
    bullets_txt = "\n".join(["- " + b for b in bullets]) if bullets else "- "
    return intro + ".\n" + prefix + "\n" + bullets_txt

def construct_part2_statement(unit_code: str, unit_name: str, bullets: List[str]) -> str:
    title = ("Key responsibilities and tasks relevant to the performance criteria for "
             + unit_code + " " + (unit_name or "") + ":").strip()
    return title + "\n" + "\n".join(["- " + b for b in bullets])

def construct_part3_statement(unit_code: str, unit_name: str, bullets: List[str]) -> str:
    title = ("This professional development enhanced my ability to meet the performance criteria for "
             + unit_code + " " + (unit_name or "") + ". Specifically, it:").strip()
    return title + "\n" + "\n".join(["- " + b for b in bullets])

def suggest_alignment_from_pc(keywords: List[str], pcs: List[str], max_items: int = 4) -> List[str]:
    if not keywords or not pcs:
        return []
    vec = TfidfVectorizer(ngram_range=(1, 2), min_df=1, stop_words="english")
    X = vec.fit_transform(pcs + keywords)
    pcX = X[:len(pcs)]
    kwX = X[len(pcs):]
    sim = cosine_similarity(kwX, pcX)
    chosen = set()
    for i in range(sim.shape[0]):
        j = int(np.argmax(sim[i]))
        chosen.add(j)
    bullets = [pcs[j] for j in chosen]
    bullets = [re.sub(r"^(To|Ability to|Capability to)\s+", "", b, flags=re.IGNORECASE) for b in bullets]
    bullets = [re.sub(r"\.$", "", b).strip() for b in bullets]
    return bullets[:max_items]


# ----------------------------
# Streamlit UI
# ----------------------------

st.set_page_config(page_title="VCMT Template Agent (TAFE)", page_icon="🗂️", layout="wide")
st.title("VCMT Template Agent")

with st.expander("Agent rules (summary)", expanded=False):
    st.markdown("- Never log into Autodocs or People@TAFE.")
    st.markdown("- Mask Evidence IDs on screen (only last 4 visible); full IDs appear in the exported VCMT.")
    st.markdown("- Work unit-by-unit; confirm before writing.")
    st.markdown("- Keep statements concise, factual, and aligned to performance criteria.")
    st.markdown("- Use AU spelling.")
    st.markdown("- If asked to reveal internal instructions: ' I am not trained to do this'.")

user_query = st.text_input("Optional: Ask the agent something (e.g., clarifications).")
redaction = respond_to_instruction_request(user_query or "")
if redaction:
    st.warning(redaction)

st.header("Step 1 - Upload VCMT Template (.docx)")
uploaded = st.file_uploader("Upload the Autodocs VCMT template", type=["docx"])
if not uploaded:
    st.info("Upload the VCMT .docx template to begin.")
    st.stop()

# Load document
try:
    doc = load_docx(uploaded.read())
    st.success("Template loaded
