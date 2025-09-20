# app.py
import io
import re
import datetime
import streamlit as st
from typing import List, Dict, Iterable, Tuple
from docx import Document
from docx.shared import Inches

# ---------- helpers ----------
def normalise_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def all_doc_text_lines(doc: Document) -> Iterable[str]:
    for p in doc.paragraphs:
        if p.text:
            yield p.text
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                # Deduplicate consecutive identical paragraphs
                last_text = None
                for p in cell.paragraphs:
                    if p.text and p.text != last_text:
                        yield p.text
                        last_text = p.text

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
        out.append(
            f"Table {i}: {len(t.rows)} rows x {len(t.columns)} cols | "
            f"header: {' | '.join(header[:6]) if header else '(none)'}"
        )
    return out

def find_part1_table_index(doc: Document) -> int:
    """Heuristic: return first table with at least 4 columns (assumed Part 1 table)."""
    for i, t in enumerate(doc.tables):
        if len(t.columns) >= 4:
            return i
    return -1

def get_section_text_for_unit(doc: Document, unit_code: str) -> Tuple[str, List[str]]:
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
            for j in range(i+1, min(i+4, len(window))):
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
        if re.search(r"Performance Evidence", p, re.IGNORECASE):
            for j in range(i+1, min(i+20, len(all_paras))):
                line = all_paras[j]
                if re.match(r"^\s*(?:•|-|\d+\.)\s*", line):
                    perf_bullets.append(re.sub(r"^\s*(?:•|-|\d+\.)\s*", "", line))
                elif len(line.split()) > 4:
                    perf_bullets.append(line)
                else:
                    if len(line.split()) <= 3:
                        break
            break

    perf_bullets = [normalise_space(x) for x in perf_bullets][:12]
    return (normalise_space(app_excerpt), perf_bullets)

def mask_evidence_id(eid: str) -> str:
    eid = (eid or "").strip()
    if not eid or eid.lower() == "pending":
        return "Pending"
    if len(eid) <= 4:
        return "*" * (len(eid)-1) + eid[-1] if len(eid) > 1 else eid
    return "*" * (len(eid)-4) + eid[-4:]

def validate_year(y: str) -> bool:
    try:
        yr = int(y)
        this_year = datetime.date.today().year
        return 1000 <= yr <= this_year
    except:
        return False

# ---------- Streamlit UI ----------
st.set_page_config(page_title="VCMT Unit Code Reader (Lite) + Filler", page_icon="🗂", layout="wide")
st.title("VCMT Unit Code Reader (Lite) — Fill & Export")

st.header("Step 1 — Upload VCMT Template (.docx)")
uploaded = st.file_uploader("Upload the Autodocs VCMT template", type=["docx"])
if not uploaded:
    st.info("Upload the VCMT .docx template to begin.")
    st.stop()

# --- FIX: Keep uploaded bytes in session state ---
if "uploaded_bytes" not in st.session_state:
    st.session_state.uploaded_bytes = uploaded.read()

try:
    doc = load_docx(st.session_state.uploaded_bytes)
    st.success("Template loaded.")
except Exception as e:
    st.error(f"Could not load .docx: {e}")
    st.stop()

with st.expander("Template tables found", expanded=False):
    for line in list_tables_info(doc):
        st.write(line)

# ... (unchanged code from Step 2 and Step 3 collecting data) ...

# ---------- QA & Export ----------
# (keep all your existing QA code above here unchanged)

# Export action - write into a copy of the uploaded doc
if st.button("Generate and Download VCMT (.docx)"):
    # --- FIX: Reload from stored bytes instead of uploaded.read() ---
    out_doc = load_docx(st.session_state.uploaded_bytes)

    t_index = find_part1_table_index(out_doc)
    if t_index == -1:
        t = out_doc.add_table(rows=1, cols=4)
        hdr_cells = t.rows[0].cells
        hdr_cells[0].text = "Column 1"
        hdr_cells[1].text = "Column 2"
        hdr_cells[2].text = "Column 3"
        hdr_cells[3].text = "Column 4"
        t_index = len(out_doc.tables) - 1
    table = out_doc.tables[t_index]

    for r in all_rows:
        new_row = table.add_row()
        new_row.cells[0].text = r["col1"]
        new_row.cells[1].text = r["col2"] or ""
        new_row.cells[2].text = r["col3"] or ""
        new_row.cells[3].text = r["col4"] or ""

    if pending:
        note_row = table.add_row()
        note_row.cells[0].text = "NOTE"
        note_row.cells[1].text = ""
        note_row.cells[2].text = f"{len(pending)} row(s) have Pending or missing Evidence IDs. Please update."
        note_row.cells[3].text = ""

    today = datetime.date.today().isoformat().replace("-", "")
    codes_for_name = "_".join(validated_codes)
    filename = f"VCMT_{codes_for_name}_{today}.docx"

    bio = io.BytesIO()
    out_doc.save(bio)
    bio.seek(0)
    st.success("VCMT generated.")
    st.download_button(
        "Download filled VCMT (.docx)",
        data=bio.read(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
