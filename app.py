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
    for p in doc.paragraphs:
        if p.text:
            yield p.text
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
    """Heuristic: first table with >= 4 columns."""
    for i, t in enumerate(doc.tables):
        if len(t.columns) >= 4:
            return i
    return -1

def get_section_text_for_unit(doc: Document, unit_code: str) -> Tuple[str, List[str]]:
    """Heuristic extraction of Application Statement and Performance Evidence."""
    paras = [normalise_space(t) for t in all_doc_text_lines(doc) if normalise_space(t)]
    idxs = [i for i, p in enumerate(paras) if unit_code in p]
    app_excerpt = ""
    perf_bullets: List[str] = []

    window = paras[max(0, (idxs[0] - 6) if idxs else 0):][:20]
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

    perf_bullets = unique_preserve(perf_bullets)[:12]
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

# IMPORTANT: capture the bytes ONCE and reuse (avoid BadZipFile on second read)
template_bytes = uploaded.getvalue()
if "template_bytes" not in st.session_state:
    st.session_state["template_bytes"] = template_bytes
else:
    # keep the first uploaded file in the session for consistent reloads
    st.session_state["template_bytes"] = template_bytes

try:
    doc = load_docx(st.session_state["template_bytes"])
    st.success("Template loaded.")
except Exception as e:
    st.error(f"Could not load .docx: {e}")
    st.stop()

with st.expander("Template tables found", expanded=False):
    for line in list_tables_info(doc):
        st.write(line)

st.header("Step 2 — Detect & select unit(s)")
extracted_units = extract_units_from_doc(doc)
discovered_codes = sorted(extracted_units.keys())
st.caption("Codes detected from both paragraphs and table cells.")

user_unit_codes = st.multiselect(
    "Choose unit code(s) to complete (enter manually if not visible).",
    options=discovered_codes,
    default=[]
)

manual_units_input = st.text_input("Optional: add unit code(s) comma-separated (e.g., BSBWHS211, SITXWHS005)")
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

part1_table_index = find_part1_table_index(doc)
if part1_table_index == -1:
    st.warning("No table with >=4 columns found. The app will still collect data but cannot auto-insert rows into the template. A new 4-column table will be appended at export.")

# Loop through each unit and collect data
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

    st.subheader(f"Unit: {unit_code} — {unit_name}")
    with st.expander(f"Open prompts for {unit_code}", expanded=not st.session_state.units_data[key]["confirmed"]):
        app_excerpt, perf_bullets = get_section_text_for_unit(doc, unit_code)

        st.markdown("**Unit context (extracted from template)**")
        if app_excerpt:
            st.write(f"Application excerpt: {app_excerpt}")
        else:
            st.write("_No clear application statement found in template._")
        if perf_bullets:
            st.write("Performance Evidence (sample lines):")
            for b in perf_bullets[:8]:
                st.write(f"- {b}")
        else:
            st.write("_No Performance Evidence lines found._")

        # ---------- Part 1 ----------
        st.markdown("### Part 1 — Qualifications / Units of Competency")
        st.write("Enter qualifications/units that support competency. One per line. After entering, press 'Parse Qualifications' to convert into rows.")
        qual_raw = st.text_area(f"Qualifications for {unit_code} (format: one per line)", key=f"qual_raw_{unit_code}", value="")
        if st.button(f"Parse Qualifications for {unit_code}", key=f"parse_q_{unit_code}"):
            quals = [normalise_space(q) for q in qual_raw.splitlines() if normalise_space(q)]
            quals = list(dict.fromkeys(quals))
            parsed = [{"qual_name": q, "year": "", "evidence_id": ""} for q in quals]
            st.session_state.units_data[key]["part1"] = parsed

        if st.session_state.units_data[key]["part1"]:
            st.write("Now provide Year (YYYY) and People@TAFE Evidence ID (or 'Pending') for each qualification.")
            for idx, entry in enumerate(st.session_state.units_data[key]["part1"]):
                cols = st.columns([4,1,2])
                with cols[0]:
                    st.write(f"**{idx+1}. {entry['qual_name']}**")
                with cols[1]:
                    year_val = st.text_input(f"Year for #{idx+1}", key=f"{unit_code}_p1_year_{idx}", value=entry.get("year",""))
                with cols[2]:
                    eid_val = st.text_input(f"Evidence ID for #{idx+1}", key=f"{unit_code}_p1_eid_{idx}", value=entry.get("evidence_id",""))
                if year_val and not validate_year(year_val):
                    st.error("Enter a valid 4-digit year <= current year.")
                st.session_state.units_data[key]["part1"][idx]["year"] = year_val
                st.session_state.units_data[key]["part1"][idx]["evidence_id"] = eid_val

            # Generate Column 3
            for idx, entry in enumerate(st.session_state.units_data[key]["part1"]):
                bullets = perf_bullets or []
                matched = []
                qual_words = set(w.lower() for w in re.findall(r"\w{4,}", entry["qual_name"]))
                for b in bullets:
                    b_words = set(w.lower() for w in re.findall(r"\w{4,}", b))
                    if qual_words & b_words:
                        matched.append(b)
                if not matched and bullets:
                    matched = bullets[:3]
                matched = unique_preserve(matched)
                col3 = f"Within this qualification, I was required to demonstrate competency in the skills and knowledge required to {app_excerpt or unit_name}.\n\nSpecifically relevant were the following course components:\n"
                for m in matched:
                    col3 += f"• {m}\n"
                st.session_state.units_data[key]["part1"][idx]["generated_statement"] = col3

        # ---------- Part 2 ----------
        st.markdown("### Part 2 — Industry / Community Experience")
        p2 = st.session_state.units_data[key].get("part2") or {"role_title":"","years_exp":"","evidence_id":"","generated_statement":""}
        p2_role = st.text_input(f"Role title for {unit_code} (e.g., Plant Operator)", key=f"{unit_code}_p2_role", value=p2.get("role_title",""))
        p2_employer = st.text_input(f"Employer (optional)", key=f"{unit_code}_p2_employer", value=p2.get("employer",""))
        p2_years = st.text_input(f"How many years in this role?", key=f"{unit_code}_p2_years", value=p2.get("years_exp",""))
        p2_eid = st.text_input(f"People@TAFE Evidence ID for role (or Pending)", key=f"{unit_code}_p2_eid", value=p2.get("evidence_id",""))
        responsibilities = unique_preserve(perf_bullets[:6] if perf_bullets else [
            "Conduct pre-start checks and operate equipment under WHS procedures and SOPs.",
            "Identify hazards and apply risk controls using site procedures.",
            "Maintain records to meet compliance and traceability standards."
        ])
        p2_col3 = f"Key responsibilities and tasks relevant to the performance criteria for {unit_code} {unit_name}:\n"
        for r in responsibilities[:7]:
            p2_col3 += f"• {r}\n"
        role_full = p2_role
        if p2_employer and p2_employer not in (p2_role or ""):
            role_full = f"{role_full} ({p2_employer})"
        st.session_state.units_data[key]["part2"] = {"role_title": role_full, "years_exp": p2_years, "evidence_id": p2_eid, "generated_statement": p2_col3}

        # ---------- Part 3 ----------
        st.markdown("### Part 3 — Professional Development (PD)")
        p3_title = st.text_input(f"PD title for {unit_code}", key=f"{unit_code}_p3_title", value=(st.session_state.units_data[key].get("part3") or {}).get("pd_title",""))
        p3_year = st.text_input(f"Year completed (YYYY)", key=f"{unit_code}_p3_year", value=(st.session_state.units_data[key].get("part3") or {}).get("year",""))
        p3_eid = st.text_input(f"People@TAFE Evidence ID (or Pending)", key=f"{unit_code}_p3_eid", value=(st
