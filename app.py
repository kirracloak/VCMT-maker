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
    """Return indexes of Part1, Part2, Part3 tables or -1 if not found."""
    p1 = p2 = p3 = -1
    for i, t in enumerate(doc.tables):
        if len(t.columns) < 4:
            continue
        header = " ".join([normalise_space(c.text) for c in t.rows[0].cells]).lower()
        if p1 == -1:
            p1 = i
        if "industry" in header or "community experience" in header:
            if p2 == -1:
                p2 = i
        if "professional development" in header:
            if p3 == -1:
                p3 = i
    return p1, p2, p3

def mask_evidence_id(eid: str) -> str:
    eid = (eid or "").strip()
    if not eid or eid.lower() == "pending":
        return "Pending"
    if len(eid) <= 4:
        return "*" * (len(eid)-1) + eid[-1] if len(eid) > 1 else eid
    return "*" * (len(eid)-4) + eid[-4:]

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

    with st.expander(f"{unit_code} — {unit_name}", expanded=True):
        # --- Part 1: multiple qualifications ---
        st.subheader("Part 1 — Qualifications / Units of Competency")
        if st.button(f"Add Qualification for {unit_code}", key=f"add_p1_{unit_code}"):
            st.session_state.units_data[key]["part1"].append({"qual_name":"", "year":"", "evidence_id":"", "generated_statement":""})
        for idx, entry in enumerate(st.session_state.units_data[key]["part1"]):
            st.text_input("Qualification name", key=f"{unit_code}_p1_name_{idx}", value=entry["qual_name"])
            st.text_input("Year completed (YYYY)", key=f"{unit_code}_p1_year_{idx}", value=entry["year"])
            st.text_input("Evidence ID", key=f"{unit_code}_p1_eid_{idx}", value=entry["evidence_id"])
            entry["qual_name"] = st.session_state[f"{unit_code}_p1_name_{idx}"]
            entry["year"] = st.session_state[f"{unit_code}_p1_year_{idx}"]
            entry["evidence_id"] = st.session_state[f"{unit_code}_p1_eid_{idx}"]
            entry["generated_statement"] = f"Within this qualification, I was required to demonstrate competency in {unit_name}."

        # --- Part 2: multiple roles ---
        st.subheader("Part 2 — Industry / Community Experience")
        if st.button(f"Add Experience for {unit_code}", key=f"add_p2_{unit_code}"):
            st.session_state.units_data[key]["part2"].append({"role_title":"", "employer":"", "years_worked":"", "evidence_id":"", "generated_statement":""})
        for idx, entry in enumerate(st.session_state.units_data[key]["part2"]):
            st.text_input("Role title", key=f"{unit_code}_p2_role_{idx}", value=entry["role_title"])
            st.text_input("Employer", key=f"{unit_code}_p2_emp_{idx}", value=entry["employer"])
            st.text_input("Years worked (e.g., 2013–2015)", key=f"{unit_code}_p2_years_{idx}", value=entry["years_worked"])
            st.text_input("Evidence ID", key=f"{unit_code}_p2_eid_{idx}", value=entry["evidence_id"])
            entry["role_title"] = st.session_state[f"{unit_code}_p2_role_{idx}"]
            entry["employer"] = st.session_state[f"{unit_code}_p2_emp_{idx}"]
            entry["years_worked"] = st.session_state[f"{unit_code}_p2_years_{idx}"]
            entry["evidence_id"] = st.session_state[f"{unit_code}_p2_eid_{idx}"]
            entry["generated_statement"] = f"Key responsibilities relevant to {unit_code} {unit_name}."

        # --- Part 3: multiple PDs ---
        st.subheader("Part 3 — Professional Development")
        if st.button(f"Add PD for {unit_code}", key=f"add_p3_{unit_code}"):
            st.session_state.units_data[key]["part3"].append({"pd_title":"", "year":"", "evidence_id":"", "generated_statement":""})
        for idx, entry in enumerate(st.session_state.units_data[key]["part3"]):
            st.text_input("PD title", key=f"{unit_code}_p3_title_{idx}", value=entry["pd_title"])
            st.text_input("Year (YYYY)", key=f"{unit_code}_p3_year_{idx}", value=entry["year"])
            st.text_input("Evidence ID", key=f"{unit_code}_p3_eid_{idx}", value=entry["evidence_id"])
            entry["pd_title"] = st.session_state[f"{unit_code}_p3_title_{idx}"]
            entry["year"] = st.session_state[f"{unit_code}_p3_year_{idx}"]
            entry["evidence_id"] = st.session_state[f"{unit_code}_p3_eid_{idx}"]
            entry["generated_statement"] = f"This professional development enhanced my ability to meet criteria for {unit_code} {unit_name}."

# --- Step 4 ---
st.header("Step 4 — QA and Export")
rows_p1, rows_p2, rows_p3 = [], [], []
for data in st.session_state.units_data.values():
    ucode = data["unit_code"]
    for p in data["part1"]:
        rows_p1.append(p)
    for r in data["part2"]:
        rows_p2.append(r)
    for r in data["part3"]:
        rows_p3.append(r)

# QA display
st.subheader("QA Preview")
if rows_p1:
    st.write("**Part 1 Entries**")
    for i, r in enumerate(rows_p1, 1):
        st.write(f"{i}. {r['qual_name']} | {r['year']} | Evidence: {mask_evidence_id(r['evidence_id'])}")
if rows_p2:
    st.write("**Part 2 Entries**")
    for i, r in enumerate(rows_p2, 1):
        st.write(f"{i}. {r['role_title']} ({r['employer']}) | {r['years_worked']} | Evidence: {mask_evidence_id(r['evidence_id'])}")
if rows_p3:
    st.write("**Part 3 Entries**")
    for i, r in enumerate(rows_p3, 1):
        st.write(f"{i}. {r['pd_title']} | {r['year']} | Evidence: {mask_evidence_id(r['evidence_id'])}")

if st.button("Generate and Download VCMT (.docx)"):
    out_doc = load_docx(st.session_state.uploaded_bytes)
    p1_idx, p2_idx, p3_idx = find_part_tables(out_doc)

    if p1_idx != -1:
        table = out_doc.tables[p1_idx]
        for r in rows_p1:
            new_row = table.add_row()
            new_row.cells[0].text = r["qual_name"]
            new_row.cells[1].text = r["year"]
            new_row.cells[2].text = r["generated_statement"]
            new_row.cells[3].text = r["evidence_id"]

    if p2_idx != -1:
        table = out_doc.tables[p2_idx]
        for r in rows_p2:
            new_row = table.add_row()
            new_row.cells[0].text = r["role_title"] + (f' ({r["employer"]})' if r["employer"] else "")
            new_row.cells[1].text = r["years_worked"]
            new_row.cells[2].text = r["generated_statement"]
            new_row.cells[3].text = r["evidence_id"]

    if p3_idx != -1:
        table = out_doc.tables[p3_idx]
        for r in rows_p3:
            new_row = table.add_row()
            new_row.cells[0].text = r["pd_title"]
            new_row.cells[1].text = r["year"]
            new_row.cells[2].text = r["generated_statement"]
            new_row.cells[3].text = r["evidence_id"]

    today = datetime.date.today().isoformat().replace("-", "")
    filename = f"VCMT_{'_'.join(validated_codes)}_{today}.docx"
    bio = io.BytesIO()
    out_doc.save(bio)
    bio.seek(0)
    st.download_button("Download filled VCMT (.docx)", data=bio.read(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
