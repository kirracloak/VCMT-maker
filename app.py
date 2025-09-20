# app.py
import io
import re
import datetime
import streamlit as st
import pandas as pd
from typing import List, Dict, Iterable
from docx import Document

# ---------- helpers ----------
def normalise_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def all_doc_text_lines(doc: Document) -> Iterable[str]:
    """Yield all text lines from paragraphs and table cells (deduping consecutive repeats)."""
    for p in doc.paragraphs:
        if p.text:
            yield p.text
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
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

def list_tables_info(doc: Document):
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
    for i, t in enumerate(doc.tables):
        if len(t.columns) >= 4:
            return i
    return -1

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
st.caption("Codes detected from both paragraphs and table cells.")

user_unit_codes = st.multiselect(
    "Choose unit code(s) to complete (enter manually if not visible).",
    options=discovered_codes,
    default=[]
)

manual_units_input = st.text_input("Optional: add unit code(s) comma-separated")
if manual_units_input.strip():
    added = [normalise_space(x).upper() for x in manual_units_input.split(",") if normalise_space(x)]
    user_unit_codes = list(dict.fromkeys([u.upper() for u in user_unit_codes] + added))

full_text_up = "\n".join(all_doc_text_lines(doc)).upper()
validated_codes = [c for c in user_unit_codes if c.strip().upper() in full_text_up]
invalid_codes = [c for c in set(user_unit_codes) - set(validated_codes)]

if invalid_codes:
    st.warning("The following codes were not found in document text and will be skipped: " + ", ".join(invalid_codes))
if not validated_codes:
    st.info("Please select at least one valid unit code.")
    st.stop()

validated_codes = [c.strip().upper() for c in validated_codes]
st.success("Validated unit codes: " + ", ".join(validated_codes))

if "units_data" not in st.session_state:
    st.session_state.units_data = {}
if "edit_unit" not in st.session_state:
    st.session_state.edit_unit = None

part1_table_index = find_part1_table_index(doc)
if part1_table_index == -1:
    st.warning("No Part 1 table found. Rows will be appended to a new table instead.")

# --- Step 3 ---
st.header("Step 3 — For each unit, answer prompts")
for unit_code in validated_codes:
    unit_name = extracted_units.get(unit_code, {}).get("name", "")
    key = f"unit__{unit_code}"
    if key not in st.session_state.units_data:
        st.session_state.units_data[key] = {
            "unit_code": unit_code,
            "unit_name": unit_name,
            "part1": [],
            "part2": None,
            "part3": None,
            "confirmed": False
        }

    expanded_flag = (st.session_state.edit_unit == unit_code) or (not st.session_state.units_data[key]["confirmed"])
    with st.expander(f"Unit: {unit_code} — {unit_name}", expanded=expanded_flag):
        # Part 1
        st.markdown("### Part 1 — Qualifications / Units of Competency")
        qual_raw = st.text_area(f"Qualifications (one per line)", key=f"qual_raw_{unit_code}")
        if st.button(f"Parse Qualifications for {unit_code}", key=f"parse_q_{unit_code}"):
            parsed = [{"qual_name": q, "year": "", "evidence_id": ""} for q in qual_raw.splitlines() if normalise_space(q)]
            st.session_state.units_data[key]["part1"] = parsed

        if st.session_state.units_data[key]["part1"]:
            for idx, entry in enumerate(st.session_state.units_data[key]["part1"]):
                cols = st.columns([4,1,2])
                with cols[0]:
                    st.write(f"**{idx+1}. {entry['qual_name']}**")
                with cols[1]:
                    year_val = st.text_input(f"Year (YYYY)", key=f"{unit_code}_p1_year_{idx}", value=entry.get("year",""))
                with cols[2]:
                    eid_val = st.text_input(f"Evidence ID", key=f"{unit_code}_p1_eid_{idx}", value=entry.get("evidence_id",""))
                if year_val and not validate_year(year_val) and not (year_val.isdigit() and len(year_val)<=2):
                    st.error("Enter a valid year (YYYY) or number of years.")
                entry["year"] = year_val
                entry["evidence_id"] = eid_val
                entry["generated_statement"] = f"Within this qualification, I was required to demonstrate competency in {unit_name}."

        # Part 2
        st.markdown("### Part 2 — Industry / Community Experience")
        p2_role = st.text_input(f"Role title", key=f"{unit_code}_p2_role")
        p2_years = st.text_input(f"How many years?", key=f"{unit_code}_p2_years")
        p2_eid = st.text_input(f"Evidence ID (or Pending)", key=f"{unit_code}_p2_eid")
        p2_col3 = f"Key responsibilities relevant to {unit_code} {unit_name}."
        st.session_state.units_data[key]["part2"] = {"role_title": p2_role, "years_exp": p2_years, "evidence_id": p2_eid, "generated_statement": p2_col3}

        # Part 3
        st.markdown("### Part 3 — Professional Development")
        p3_title = st.text_input(f"PD title", key=f"{unit_code}_p3_title")
        p3_year = st.text_input(f"Year completed (YYYY)", key=f"{unit_code}_p3_year")
        p3_eid = st.text_input(f"Evidence ID (or Pending)", key=f"{unit_code}_p3_eid")
        p3_col3 = f"This professional development enhanced my ability to meet criteria for {unit_code} {unit_name}."
        st.session_state.units_data[key]["part3"] = {"pd_title": p3_title, "year": p3_year, "evidence_id": p3_eid, "generated_statement": p3_col3}

        if st.button(f"Mark {unit_code} Complete", key=f"confirm_{unit_code}"):
            st.session_state.units_data[key]["confirmed"] = True
            st.session_state.edit_unit = None
            st.success(f"{unit_code} marked complete.")

# --- Step 4 ---
st.header("Step 4 — QA and Export")
all_rows = []
row_to_unit = []
for data in st.session_state.units_data.values():
    ucode = data["unit_code"]
    # Part 1
    for p in data["part1"]:
        all_rows.append({"unit_code": ucode, "col1": p["qual_name"], "col2": p["year"], "col3": p["generated_statement"], "col4": p["evidence_id"]})
        row_to_unit.append(ucode)
    # Part 2
    if data["part2"]:
        r = data["part2"]
        all_rows.append({"unit_code": ucode, "col1": r["role_title"], "col2": r["years_exp"], "col3": r["generated_statement"], "col4": r["evidence_id"]})
        row_to_unit.append(ucode)
    # Part 3
    if data["part3"]:
        r = data["part3"]
        all_rows.append({"unit_code": ucode, "col1": r["pd_title"], "col2": r["year"], "col3": r["generated_statement"], "col4": r["evidence_id"]})
        row_to_unit.append(ucode)

if not all_rows:
    st.info("No rows collected yet.")
    st.stop()

# --- QA Preview with Edit buttons ---
st.subheader("QA Preview")
for i, row in enumerate(all_rows, start=1):
    unit = row_to_unit[i-1]
    valid_year = (not row["col2"]) or validate_year(row["col2"]) or (row["col2"].isdigit() and len(row["col2"])<=2)
    pending = (not row["col4"]) or row["col4"].strip().lower()=="pending"

    color = "lightgreen"
    if not valid_year:
        color = "salmon"
    elif pending:
        color = "khaki"

    st.markdown(
        f"<div style='background-color:{color}; padding:6px; border-radius:5px;'>"
        f"<b>Row {i}</b> — Unit {row['unit_code']} | {row['col1']} | {row['col2']} | Evidence: {mask_evidence_id(row['col4'])}"
        f"</div>",
        unsafe_allow_html=True
    )
    if st.button(f"Edit Row {i}", key=f"editrow_{i}"):
        st.session_state.edit_unit = unit
        st.experimental_rerun()

# --- Export ---
if st.button("Generate and Download VCMT (.docx)"):
    out_doc = load_docx(st.session_state.uploaded_bytes)
    t_index = find_part1_table_index(out_doc)
    if t_index == -1:
        t = out_doc.add_table(rows=1, cols=4)
        hdr = t.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Column 1","Column 2","Column 3","Column 4"
        t_index = len(out_doc.tables)-1
    table = out_doc.tables[t_index]
    for r in all_rows:
        new_row = table.add_row()
        new_row.cells[0].text = r["col1"]
        new_row.cells[1].text = r["col2"]
        new_row.cells[2].text = r["col3"]
        new_row.cells[3].text = r["col4"]
    today = datetime.date.today().isoformat().replace("-","")
    filename = f"VCMT_{'_'.join(validated_codes)}_{today}.docx"
    bio = io.BytesIO()
    out_doc.save(bio)
    bio.seek(0)
    st.download_button("Download filled VCMT (.docx)", data=bio.read(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
