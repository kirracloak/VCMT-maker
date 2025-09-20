# app.py
import io
import re
import datetime
import streamlit as st
from typing import List, Dict, Iterable
from docx import Document

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
                for p in cell.paragraphs:
                    if p.text:
                        yield p.text

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

def find_part_tables(doc: Document):
    """Return indexes of Part1, Part2, Part3 tables by header text (case insensitive)."""
    p1 = p2 = p3 = -1
    for i, t in enumerate(doc.tables):
        if len(t.columns) < 4:
            continue
        header = " ".join([normalise_space(c.text) for c in t.rows[0].cells]).lower()
        if "qualification" in header or "units of competency" in header:
            if p1 == -1: p1 = i
        elif "industry" in header or "community experience" in header:
            if p2 == -1: p2 = i
        elif "professional development" in header:
            if p3 == -1: p3 = i
    return p1, p2, p3

def insert_into_table(table, values: List[str]):
    """
    Insert into the first free (blank) row if available, otherwise add a new row.
    values: list of strings matching table columns.
    """
    def row_is_blank(row):
        return all(not normalise_space(c.text) for c in row.cells)

    # Try to find first completely blank row (skip header row 0)
    target_row = None
    for r in table.rows[1:]:
        if row_is_blank(r):
            target_row = r
            break

    if not target_row:
        target_row = table.add_row()

    for i, v in enumerate(values):
        if i < len(target_row.cells):
            target_row.cells[i].text = v or ""

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
st.set_page_config(page_title="VCMT Unit Code Filler", page_icon="🗂", layout="wide")
st.title("VCMT Unit Code Reader — Fill & Export")

# --- Step 1 ---
st.header("Step 1 — Upload VCMT Template (.docx)")
uploaded = st.file_uploader("Upload the Autodocs VCMT template", type=["docx"])
if not uploaded:
    st.info("Upload the VCMT .docx template to begin.")
    st.stop()

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

# --- Step 2 ---
st.header("Step 2 — Detect & select unit(s)")
extracted_units = extract_units_from_doc(doc)
discovered_codes = sorted(extracted_units.keys())
user_unit_codes = st.multiselect("Choose unit code(s)", options=discovered_codes, default=[])

manual_units_input = st.text_input("Optional: add unit code(s) comma-separated")
if manual_units_input.strip():
    added = [normalise_space(x).upper() for x in manual_units_input.split(",") if normalise_space(x)]
    user_unit_codes = list(dict.fromkeys([u.upper() for u in user_unit_codes] + added))

full_text_up = "\n".join(all_doc_text_lines(doc)).upper()
validated_codes = [c for c in user_unit_codes if c.strip().upper() in full_text_up]

if not validated_codes:
    st.info("Please select at least one valid unit code.")
    st.stop()
validated_codes = [c.strip().upper() for c in validated_codes]
st.success("Validated unit codes: " + ", ".join(validated_codes))

if "units_data" not in st.session_state:
    st.session_state.units_data = {}
if "edit_unit" not in st.session_state:
    st.session_state.edit_unit = None

# --- Step 3 ---
st.header("Step 3 — Enter details for each unit")
for unit_code in validated_codes:
    unit_name = extracted_units.get(unit_code, {}).get("name", "")
    key = f"unit__{unit_code}"
    if key not in st.session_state.units_data:
        st.session_state.units_data[key] = {
            "unit_code": unit_code,
            "unit_name": unit_name,
            "part1": [],
            "part2": [],
            "part3": []
        }

    expanded_flag = (st.session_state.edit_unit == unit_code)
    with st.expander(f"{unit_code} — {unit_name}", expanded=expanded_flag or True):
        # --- Part 1 ---
        st.subheader("Part 1 — Qualifications / Units of Competency")
        if st.button(f"Add Qualification for {unit_code}", key=f"add_p1_{unit_code}"):
            st.session_state.units_data[key]["part1"].append({"qual_name":"", "year":"", "evidence_id":"", "generated_statement":""})
        for idx, entry in enumerate(st.session_state.units_data[key]["part1"]):
            entry["qual_name"] = st.text_input("Qualification name", key=f"{unit_code}_p1_name_{idx}", value=entry["qual_name"])
            entry["year"] = st.text_input("Year completed (YYYY)", key=f"{unit_code}_p1_year_{idx}", value=entry["year"])
            entry["evidence_id"] = st.text_input("Evidence ID", key=f"{unit_code}_p1_eid_{idx}", value=entry["evidence_id"])
            entry["generated_statement"] = f"Within this qualification, I was required to demonstrate competency in {unit_name}."

        # --- Part 2 ---
        st.subheader("Part 2 — Industry / Community Experience")
        if st.button(f"Add Experience for {unit_code}", key=f"add_p2_{unit_code}"):
            st.session_state.units_data[key]["part2"].append({"role_title":"", "employer":"", "years_worked":"", "evidence_id":"", "generated_statement":""})
        for idx, entry in enumerate(st.session_state.units_data[key]["part2"]):
            entry["role_title"] = st.text_input("Role title", key=f"{unit_code}_p2_role_{idx}", value=entry["role_title"])
            entry["employer"] = st.text_input("Employer", key=f"{unit_code}_p2_emp_{idx}", value=entry["employer"])
            entry["years_worked"] = st.text_input("Years worked (e.g., 2013–2015)", key=f"{unit_code}_p2_years_{idx}", value=entry["years_worked"])
            entry["evidence_id"] = st.text_input("Evidence ID", key=f"{unit_code}_p2_eid_{idx}", value=entry["evidence_id"])
            entry["generated_statement"] = f"Key responsibilities relevant to {unit_code} {unit_name}."

        # --- Part 3 ---
        st.subheader("Part 3 — Professional Development")
        if st.button(f"Add PD for {unit_code}", key=f"add_p3_{unit_code}"):
            st.session_state.units_data[key]["part3"].append({"pd_title":"", "year":"", "evidence_id":"", "generated_statement":""})
        for idx, entry in enumerate(st.session_state.units_data[key]["part3"]):
            entry["pd_title"] = st.text_input("PD title", key=f"{unit_code}_p3_title_{idx}", value=entry["pd_title"])
            entry["year"] = st.text_input("Year (YYYY)", key=f"{unit_code}_p3_year_{idx}", value=entry["year"])
            entry["evidence_id"] = st.text_input("Evidence ID", key=f"{unit_code}_p3_eid_{idx}", value=entry["evidence_id"])
            entry["generated_statement"] = f"This professional development enhanced my ability to meet criteria for {unit_code} {unit_name}."

# --- Step 4 ---
st.header("Step 4 — QA and Export")
rows_p1, rows_p2, rows_p3, row_to_unit = [], [], [], []
for data in st.session_state.units_data.values():
    ucode = data["unit_code"]
    for p in data["part1"]:
        rows_p1.append(p); row_to_unit.append(ucode)
    for r in data["part2"]:
        rows_p2.append(r); row_to_unit.append(ucode)
    for r in data["part3"]:
        rows_p3.append(r); row_to_unit.append(ucode)

st.subheader("QA Preview")
def qa_block(rows, partname):
    for i, r in enumerate(rows, 1):
        missing = not r["qual_name"] if "qual_name" in r else not r.get("role_title", r.get("pd_title"))
        pending = (not r["evidence_id"]) or r["evidence_id"].lower()=="pending"
        invalid_year = "year" in r and r["year"] and not validate_year(r["year"])
        color = "lightgreen"
        if missing or invalid_year: color="salmon"
        elif pending: color="khaki"
        label = f"{r.get('qual_name') or r.get('role_title') or r.get('pd_title')} | {r.get('year',r.get('years_worked',''))} | Evidence: {mask_evidence_id(r['evidence_id'])}"
        st.markdown(f"<div style='background-color:{color}; padding:6px; border-radius:5px;'>{partname}: {label}</div>", unsafe_allow_html=True)
        if st.button(f"Edit {partname} Row {i}", key=f"edit_{partname}_{i}"):
            st.session_state.edit_unit = row_to_unit[i-1]
            st.experimental_rerun()

if rows_p1: qa_block(rows_p1, "Part1")
if rows_p2: qa_block(rows_p2, "Part2")
if rows_p3: qa_block(rows_p3, "Part3")

if st.button("Generate and Download VCMT (.docx)"):
    out_doc = load_docx(st.session_state.uploaded_bytes)
    p1_idx, p2_idx, p3_idx = find_part_tables(out_doc)

    if p1_idx != -1:
        for r in rows_p1:
            insert_into_table(out_doc.tables[p1_idx], [
                r["qual_name"], r["year"], r["generated_statement"], r["evidence_id"]
            ])

    if p2_idx != -1:
        for r in rows_p2:
            insert_into_table(out_doc.tables[p2_idx], [
                r["role_title"] + (f' ({r["employer"]})' if r["employer"] else ""),
                r["years_worked"], r["generated_statement"], r["evidence_id"]
            ])

    if p3_idx != -1:
        for r in rows_p3:
            insert_into_table(out_doc.tables[p3_idx], [
                r["pd_title"], r["year"], r["generated_statement"], r["evidence_id"]
            ])

    today = datetime.date.today().isoformat().replace("-", "")
    filename = f"VCMT_{'_'.join(validated_codes)}_{today}.docx"
    bio = io.BytesIO()
    out_doc.save(bio)
    bio.seek(0)
    st.download_button("Download filled VCMT (.docx)", data=bio.read(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
