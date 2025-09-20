import io
import re
import streamlit as st
from typing import List, Dict, Iterable
from docx import Document

# ---------- helpers ----------
def normalise_space(s: str) -> str:
    """Collapse multiple spaces into one and strip ends."""
    return re.sub(r"\s+", " ", (s or "")).strip()

def all_doc_text_lines(doc: Document) -> Iterable[str]:
    """Yield all text lines from paragraphs and table cells."""
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
    """Load a .docx file from raw bytes."""
    return Document(io.BytesIO(file_bytes))

def extract_units_from_doc(doc: Document) -> Dict[str, Dict]:
    """Extract unit codes and names from the document."""
    paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    full_text = "\n".join(paras)

    # e.g. BSBWHS311, SITXWHS005
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
    """Return basic info about tables in the document."""
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

# ---------- UI ----------
st.set_page_config(page_title="VCMT Unit Code Reader (Lite)", page_icon="🗂", layout="wide")
st.title("VCMT Unit Code Reader (Lite)")

st.header("Step 1 - Upload VCMT Template (.docx)")
uploaded = st.file_uploader("Upload the Autodocs VCMT template", type=["docx"])
if not uploaded:
    st.info("Upload the VCMT .docx template to begin.")
    st.stop()

try:
    doc = load_docx(uploaded.read())
    st.success("Template loaded.")
except Exception as e:
    st.error(f"Could not load .docx: {e}")
    st.stop()

with st.expander("Template tables found", expanded=False):
    for line in list_tables_info(doc):
        st.write(line)

st.header("Step 2 - Select Unit(s)")
extracted_units = extract_units_from_doc(doc)
discovered_codes = sorted(extracted_units.keys())
st.caption("Codes below are detected from both paragraphs and table cells.")

user_unit_codes = st.multiselect(
    "Choose unit code(s) to complete (enter manually if not visible).",
    options=discovered_codes,
    default=[]
)

manual_units_input = st.text_input(
    "Optional: add unit code(s) comma-separated (e.g., BSBWHS211, SITXWHS005)"
)
if manual_units_input.strip():
    user_unit_codes += [normalise_space(x) for x in manual_units_input.split(",") if normalise_space(x)]

# validate against ALL text (case-insensitive)
full_text = "\n".join(all_doc_text_lines(doc))
full_text_up = full_text.upper()
validated_codes = [c for c in user_unit_codes if c.strip().upper() in full_text_up]
invalid_codes = [c for c in set(user_unit_codes) - set(validated_codes)]

if invalid_codes:
    st.warning(
        "The following codes were not found in the document text and will be skipped: "
        + ", ".join(invalid_codes)
    )

if not validated_codes:
    st.info("Please select at least one valid unit code.")
    st.stop()

st.success("Validated unit codes: " + ", ".join(validated_codes))

st.header("Preview")
for code in validated_codes:
    meta = extracted_units.get(code.upper(), extracted_units.get(code, {}))
    st.write(f"- {code}  |  {meta.get('name','')}")
st.info(
    "This lite version proves code detection from tables is working. "
    "If you want, I can layer back the Part 1/2/3 writing features on top of this stable base."
)
