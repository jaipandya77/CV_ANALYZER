import io
import json
import re
import os
import time
import base64
import urllib.parse
from difflib import SequenceMatcher
from datetime import datetime
import streamlit as st
import PyPDF2
import openpyxl
from google import genai
from google.genai import types, errors

# Page configuration must be the first Streamlit command
st.set_page_config(page_title="CV Analyzer", page_icon="📄", layout="wide")

# ------------------------------------------------------------
# EXECUTIVE MIDNIGHT THEME (OPTIMIZED FOR 60 FPS SCROLLING)
# ------------------------------------------------------------
st.markdown("""<style>
:root {
    --bg-canvas: #0b0f19;
    --bg-card: #151c2c;
    --border-subtle: #243048;
    --border-accent: rgba(99, 102, 241, 0.35);
    --text-main: #f8fafc;
    --text-muted: #94a3b8;
    --brand-primary: #6366f1;
}

/* 1. Hardware-accelerated fixed canvas (zero repaint on scroll) */
.stApp {
    background-color: #0b0f19 !important;
    background-image: radial-gradient(circle at 15% 0%, rgba(99,102,241,.12), transparent 360px),
                      radial-gradient(circle at 85% 15%, rgba(139,92,246,.08), transparent 320px) !important;
    background-attachment: fixed !important;
    color: var(--text-main) !important;
}

[data-testid="stAppViewContainer"], [data-testid="stMainBlockContainer"] {
    background: transparent !important;
}

.block-container {
    max-width: 1240px;
    padding-top: 1.6rem;
    padding-bottom: 3.5rem;
}

h1, h2, h3, h4, h5, p, span, label {
    color: var(--text-main) !important;
    letter-spacing: -0.02em;
}

/* 2. Hero Banner */
.app-hero {
    position: relative;
    border: 1px solid var(--border-accent);
    border-radius: 20px;
    padding: 28px 32px;
    margin-bottom: 18px;
    background: linear-gradient(135deg, #172033 0%, #0f172a 100%);
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
}
.app-hero .eyebrow {
    display: inline-flex;
    align-items: center;
    padding: 5px 12px;
    border-radius: 999px;
    background: rgba(99, 102, 241, 0.18);
    border: 1px solid rgba(99, 102, 241, 0.3);
    font-size: .74rem;
    text-transform: uppercase;
    letter-spacing: .12em;
    font-weight: 800;
    color: #a5b4fc !important;
    margin-bottom: 10px;
}
.app-hero h1 {
    margin: 0;
    font-size: 2.35rem;
    font-weight: 800;
    color: #ffffff !important;
}
.app-hero p {
    margin: 8px 0 0;
    color: #94a3b8 !important;
    font-size: 0.98rem;
    line-height: 1.6;
    max-width: 820px;
}

/* 3. Workflow Stepper */
.workflow {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 12px;
    margin: 14px 0 24px;
}
.workflow-step {
    position: relative;
    border: 1px solid var(--border-subtle);
    border-radius: 13px;
    padding: 13px 15px;
    background: var(--bg-card);
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.18);
    font-size: .88rem;
    font-weight: 700;
    color: #e2e8f0 !important;
}
.workflow-step span {
    display: block;
    color: #818cf8 !important;
    font-size: .68rem;
    text-transform: uppercase;
    letter-spacing: .1em;
    margin-bottom: 2px;
}
.workflow-step:not(:last-child):after {
    content: "›";
    position: absolute;
    right: -9px;
    top: 50%;
    transform: translateY(-50%);
    color: #4f46e5;
    font-size: 1.35rem;
    z-index: 2;
}

/* 4. Dropzone */
[data-testid="stFileUploader"] section {
    background: var(--bg-card) !important;
    border: 1.5px dashed #4338ca !important;
    border-radius: 16px !important;
    padding: 24px !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #6366f1 !important;
    background: #182236 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] * {
    color: #cbd5e1 !important;
}
[data-testid="stFileUploader"] section button {
    background: linear-gradient(135deg, #4f46e5, #6366f1) !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 6px 16px !important;
    box-shadow: 0 2px 8px rgba(79, 70, 229, 0.25);
}
[data-testid="stFileUploader"] section button * {
    color: #ffffff !important;
    font-weight: 750 !important;
}
[data-testid="stFileUploaderFile"] {
    background: #1e293b !important;
    border: 1px solid #334155 !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploaderFile"] * {
    color: #f8fafc !important;
}

/* 5. Buttons */
button[kind="primary"], 
.stButton > button[kind="primary"], 
div.stButton > button:first-child,
.stDownloadButton > button {
    background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%) !important;
    border: none !important;
    border-radius: 12px !important;
    min-height: 2.9rem !important;
    box-shadow: 0 4px 14px rgba(79, 70, 229, 0.25) !important;
    transition: transform .12s ease !important;
}
button[kind="primary"]:hover, 
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button:hover {
    transform: translateY(-1px) !important;
}
button[kind="primary"] *, 
.stButton > button[kind="primary"] *, 
div.stButton > button:first-child *,
.stDownloadButton > button * {
    color: #ffffff !important;
    font-weight: 750 !important;
    font-size: 1rem !important;
}

/* 6. Form Controls */
[data-testid="stWidgetLabel"] label,
[data-testid="stWidgetLabel"] p {
    color: #e2e8f0 !important;
    font-weight: 700 !important;
    font-size: 0.92rem !important;
}
div[data-baseweb="select"] > div {
    background: var(--bg-card) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 11px !important;
    color: #f8fafc !important;
}
div[data-baseweb="select"] > div:focus-within {
    border-color: #6366f1 !important;
}
div[data-baseweb="select"] * {
    color: #f8fafc !important;
}
div[data-baseweb="popover"] ul,
div[data-baseweb="menu"] {
    background: #111827 !important;
    border: 1px solid #374151 !important;
}
div[data-baseweb="menu"] li {
    background: #111827 !important;
    color: #f8fafc !important;
}
div[data-baseweb="menu"] li:hover {
    background: #1f2937 !important;
    color: #818cf8 !important;
}

/* 7. Tags & Expanders */
[data-baseweb="tag"] {
    background: #1e293b !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
}
[data-baseweb="tag"] span, [data-baseweb="tag"] svg {
    color: #e2e8f0 !important;
}
[data-testid="stExpander"] {
    background: var(--bg-card) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 14px !important;
    margin-top: .4rem !important;
    margin-bottom: .8rem !important;
}
[data-testid="stExpander"] summary, [data-testid="stExpander"] summary * {
    color: #f8fafc !important;
    font-weight: 750 !important;
}

.candidate-chip {
    display: inline-flex;
    align-items: center;
    padding: 5px 12px;
    border-radius: 999px;
    background: rgba(99, 102, 241, 0.18);
    border: 1px solid rgba(99, 102, 241, 0.32);
    color: #a5b4fc !important;
    font-size: .72rem;
    font-weight: 800;
    letter-spacing: .08em;
    margin-bottom: 8px;
}

.app-footer {
    text-align: center;
    color: #64748b !important;
    font-size: .84rem;
    padding: 2.8rem 0 .5rem;
}
.app-footer strong {
    color: #94a3b8 !important;
}
hr {
    border-color: #1e293b !important;
}
@media (max-width: 800px) {
    .workflow { grid-template-columns: 1fr; }
    .workflow-step:not(:last-child):after { display: none; }
}
</style>""", unsafe_allow_html=True)

# ------------------------------------------------------------
# CREDENTIALS & DIRECTORY RESOLUTION
# ------------------------------------------------------------
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    st.error("Missing GEMINI_API_KEY. Add it to .streamlit/secrets.toml or your environment variables.")
    st.stop()

client = genai.Client(api_key=GEMINI_API_KEY)
MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
V7_PATH = os.path.join(BASE_DIR, "skillgroomers_master_data_v7.json")
V6_PATH = os.path.join(BASE_DIR, "skillgroomers_master_data_v6.json")
MASTER_DATA_FILE = V7_PATH if os.path.exists(V7_PATH) else V6_PATH


class GeminiTemporaryUnavailable(Exception):
    pass


def extract_text_from_pdf(file):
    pdf_bytes = file.getvalue() if hasattr(file, "getvalue") else file.read()
    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    text = []
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text.append(extracted)
    return "\n".join(text)


resume_schema = {
    "type": "object",
    "properties": {
        "candidate_name": {"type": "string"},
        "email": {"type": "string"},
        "phone": {"type": "string"},
        "alternate_phone": {"type": "string"},
        "date_of_birth": {"type": "string"},
        "gender": {"type": "string"},
        "current_city": {"type": "string"},
        "current_state": {"type": "string"},
        "native_city": {"type": "string"},
        "native_state": {"type": "string"},
        "core_role": {"type": "string"},
        "functional_area": {"type": "string"},
        "key_skills": {"type": "array", "items": {"type": "string"}},
        "role": {"type": "string"},
        "industry": {"type": "string"},
        "highest_qualification": {
            "type": "object",
            "properties": {
                "qualification_text": {"type": "string"},
                "degree": {"type": "string"},
                "course": {"type": "string"},
                "specialization": {"type": "string"},
                "institute": {"type": "string"},
                "year": {"type": "string"}
            },
            "required": ["qualification_text", "degree", "course", "specialization", "institute", "year"]
        },
        "work_experience": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "company": {"type": "string"},
                    "designation": {"type": "string"},
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"},
                    "country": {"type": "string"}
                },
                "required": ["company", "designation", "start_date", "end_date", "country"]
            }
        },
        "annual_salary": {"type": "string"},
        "certifications": {"type": "array", "items": {"type": "string"}}
    },
    "required": [
        "candidate_name", "email", "phone", "alternate_phone", "date_of_birth", "gender",
        "current_city", "current_state", "native_city", "native_state", "core_role",
        "functional_area", "key_skills", "role", "industry",
        "highest_qualification", "work_experience", "annual_salary", "certifications"
    ]
}


def extract_resume_data(raw_text, pdf_bytes=None):
    prompt = """
You are a precision factual data extraction engine for HR resumes.
The resume is the SOLE source of truth. Never invent, extrapolate, or estimate facts.
Accuracy takes absolute precedence over completeness. If a field is missing, return an empty string "" or an empty array [].
Do NOT calculate total experience or tenure—Python handles all math deterministically.

1. CANDIDATE PROFILE & CONTACT:
- candidate_name: Full official name in Title Case.
- phone / alternate_phone: Primary and secondary contact numbers (strip redundant label words).
- email: Valid email address.
- date_of_birth: YYYY-MM-DD or DD/MM/YYYY if explicitly present; else "".
- gender: "Male", "Female", or "" if not mentioned.

2. LOCATIONS:
- current_city / current_state: Candidate's present residential city/state.
- native_city / native_state: Candidate's permanent home address/native place.
- Note: Districts and localities are not states (e.g., Thane is in Maharashtra).

3. DOMAIN & CLASSIFICATION FACTS:
- role: Verbatim current or most recent official job designation as written on the resume. Do NOT summarize or shorten.
- core_role: Candidate's primary functional specialty (e.g., QA/QC, Safety, Planning, Civil Engineer, Billing, MEP, Rebar Detailing, Architecture).
- functional_area: Broader department (e.g., Quality Assurance / Quality Control, Project Management / Site Engineering, Safety / Health / Environment).
- industry: Primary industry (e.g., Construction, Real Estate, Infrastructure, Oil & Gas, IT).
- key_skills: Extract individual technical tools, methodologies, codes, and domain skills as separate ATOMIC array items (e.g., ["AutoCAD", "BBS", "Revit"], NOT ["AutoCAD and Revit"]). Strictly exclude soft skills (e.g., "hardworking", "team player", "punctual") and certifications.
- certifications: Specific professional credentials (e.g., PMP, NEBOSH, IOSH, LEED AP, Six Sigma).

4. HIGHEST QUALIFICATION:
- Extract the single HIGHEST COMPLETED academic qualification only.
- Strict completion rule: Ignore pursuing, ongoing, or incomplete degrees.
- Hierarchy: Doctorate > Masters > Bachelors > Diploma > HSC (12th) > SSC (10th).
- qualification_text: The verbatim phrase written in the CV (e.g., "Diploma in Civil Engineering", "B.Tech Mechanical").
- degree: Standard degree name (e.g., "Bachelor of Technology", "Diploma").
- course: Standard course abbreviation (e.g., "B.Tech", "Diploma", "B.E.").
- specialization: Specific academic branch only (e.g., "Civil Engineering", "Mechanical", "Electrical"). If no branch is mentioned (e.g., "Polytechnic Diploma"), return "".
- institute: Name of the college/university.
- year: 4-digit completion year (e.g., "2018").

5. WORK EXPERIENCE:
- Extract every professional position in chronological order, OLDEST FIRST.
- company: Clean official organization name (remove extra descriptions).
- designation: Exact designation held at that organization.
- start_date / end_date: Standardize to "YYYY-MM" or "YYYY". If the position is currently active, end_date must be "Present".
- country: Country of employment (default "India" if cities are in India).

6. COMPENSATION:
- annual_salary: The raw compensation string exactly as stated (e.g., "6.5 LPA", "8,50,000 INR"). If unstated, return "".
"""

    if pdf_bytes:
        contents = [types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt]
    else:
        contents = prompt + "\n\nRESUME TEXT:\n" + (raw_text or "")

    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=contents,
                config={
                    "response_mime_type": "application/json",
                    "response_schema": resume_schema,
                    "temperature": 0
                }
            )
            return json.loads(response.text)
        except errors.ServerError as error:
            status_code = getattr(error, "code", None) or getattr(error, "status_code", None)
            if status_code != 503 and "503" not in str(error).upper():
                raise
            if attempt == max_attempts - 1:
                raise GeminiTemporaryUnavailable("Gemini is temporarily busy. Please try again.") from error
            time.sleep(2 ** (attempt + 1))


# ---------- Deterministic Math & Parsing Helpers ----------

def parse_start_date(value):
    if not value or not isinstance(value, str):
        return datetime(1900, 1, 1)
    val = value.strip()
    for fmt in ("%Y-%m-%d", "%Y-%m", "%d-%m-%Y", "%d/%m/%Y", "%m/%Y", "%b %Y", "%B %Y"):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            pass
    match = re.search(r"\b(19\d{2}|20\d{2})\b", val)
    return datetime(int(match.group(1)), 1, 1) if match else datetime(1900, 1, 1)


def parse_end_date(value):
    if not value or not isinstance(value, str):
        return datetime.today()
    val = value.strip()
    if val.lower() in ["present", "current", "ongoing", "till date", "now"]:
        return datetime.today()
    for fmt in ("%Y-%m-%d", "%Y-%m", "%d-%m-%Y", "%d/%m/%Y", "%m/%Y", "%b %Y", "%B %Y"):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            pass
    match = re.search(r"\b(19\d{2}|20\d{2})\b", val)
    return datetime(int(match.group(1)), 12, 31) if match else datetime.today()


def sort_jobs(jobs):
    return sorted(jobs or [], key=lambda j: parse_start_date(j.get("start_date", "")))


def _month_index(year, month):
    return (year * 12) + (month - 1)


def _is_year_only_date(value):
    return bool(re.fullmatch(r"\d{4}", str(value or "").strip()))


def has_year_only_employment_dates(jobs):
    return any(_is_year_only_date(j.get("start_date")) or _is_year_only_date(j.get("end_date")) for j in (jobs or []))


def _job_month_range(job, as_of=None):
    as_of = as_of or datetime.today()
    s_val = str(job.get("start_date", "") or "").strip()
    e_val = str(job.get("end_date", "") or "").strip()
    start = parse_start_date(s_val)
    if start.year == 1900:
        return None
    if e_val.lower() in ["present", "current", "ongoing", "till date", "now"]:
        end = as_of
    elif _is_year_only_date(e_val):
        ey = int(e_val)
        end = datetime(ey - 1, 12, 1) if (_is_year_only_date(s_val) and ey > start.year) else datetime(ey, 12, 1)
    else:
        end = parse_end_date(e_val)
    return None if end < start else (_month_index(start.year, start.month), _month_index(end.year, end.month))


def total_experience_months(jobs, as_of=None):
    covered = set()
    for job in jobs or []:
        r = _job_month_range(job, as_of=as_of)
        if r:
            covered.update(range(r[0], r[1] + 1))
    return len(covered)


def total_experience_months_upper_estimate(jobs, as_of=None):
    covered = set()
    as_of = as_of or datetime.today()
    for job in jobs or []:
        start = parse_start_date(job.get("start_date", ""))
        if start.year == 1900:
            continue
        e_val = str(job.get("end_date", "") or "").strip().lower()
        end = as_of if e_val in ["present", "current", "ongoing", "till date", "now"] else parse_end_date(job.get("end_date", ""))
        if end >= start:
            covered.update(range(_month_index(start.year, start.month), _month_index(end.year, end.month) + 1))
    return len(covered)


def split_months(total_months):
    m = max(0, int(round(total_months)))
    return m // 12, m % 12


def format_months(total_months):
    y, m = split_months(total_months)
    parts = []
    if y: parts.append(f"{y} year{'s' if y != 1 else ''}")
    if m: parts.append(f"{m} month{'s' if m != 1 else ''}")
    return " ".join(parts) if parts else "0 years"


def experience_display(jobs, as_of=None):
    lower = total_experience_months(jobs, as_of=as_of)
    if not has_year_only_employment_dates(jobs):
        return format_months(lower)
    upper = max(lower, total_experience_months_upper_estimate(jobs, as_of=as_of))
    return f"Approx. {format_months(lower)}" if upper == lower else f"Approx. {format_months(lower)} – {format_months(upper)}"


def number_of_employers(jobs):
    return len({(j.get("company") or "").strip().lower() for j in (jobs or []) if (j.get("company") or "").strip()})


def compute_job_changes(jobs):
    sorted_j = sort_jobs(jobs or [])
    if len(sorted_j) < 2:
        return 0
    changes, prev = 0, None
    for j in sorted_j:
        curr = (j.get("company") or "").strip().lower()
        if curr and prev and curr != prev:
            changes += 1
        if curr:
            prev = curr
    return changes


def average_tenure_months(total_months, job_changes):
    return total_months if job_changes <= 0 else int(round(total_months / job_changes))


def get_current_and_previous_jobs(jobs):
    sorted_jobs = sort_jobs(jobs or [])
    if not sorted_jobs:
        return None, None
    current = next((j for j in reversed(sorted_jobs) if str(j.get("end_date", "")).strip().lower() in {"present", "current", "ongoing", "till date", "now"}), sorted_jobs[-1])
    cur_comp = normalize_master_text(current.get("company", ""))
    previous = next((j for j in reversed(sorted_jobs) if j is not current and normalize_master_text(j.get("company", "")) and normalize_master_text(j.get("company", "")) != cur_comp), None)
    return current, previous


def clean_key_skills(skills):
    if not skills:
        return []
    certs = ["pmp", "six sigma", "leed ap", "prince2", "scrum", "itil", "ccna", "ccnp", "aws", "azure", "gcp"]
    res, seen = [], set()
    for s in skills:
        if isinstance(s, str):
            c = s.strip().strip("'\"").strip()
            if c and c.lower() not in seen and not any(cert in c.lower() for cert in certs):
                seen.add(c.lower())
                res.append(c)
    return res


def parse_salary_lakhs(salary_str):
    if not salary_str:
        return None
    s = str(salary_str).strip()
    m_lakh = re.search(r"(\d+(?:\.\d+)?)\s*(?:lpa|lac|lacs|lakh|lakhs)", s, re.I)
    if m_lakh:
        return float(m_lakh.group(1))
    m_num = re.search(r"(\d+(?:\.\d+)?)", s.replace(",", "").strip())
    if m_num:
        num = float(m_num.group(1))
        return round(num / 100000.0, 2) if num >= 10000 else num if num < 100 else None
    return None


def reconcile_qualification_text(data):
    result = dict(data)
    qual = dict(result.get("highest_qualification", {}) or {})
    qn = normalize_master_text(str(qual.get("qualification_text", "") or ""))
    if qn:
        if qn.startswith("degree in ") and not any(t in qn for t in ["b e", "bachelor of engineering", "b tech", "btech"]):
            qual["degree"] = qual.get("qualification_text")
            qual["course"] = ""
            branch = re.sub(r"^degree\s+in\s+", "", qual["degree"], flags=re.I).strip()
            if branch:
                qual["specialization"] = branch
        elif "diploma" in qn:
            qual["degree"] = "Diploma"
            qual["course"] = "Diploma"
    result["highest_qualification"] = qual
    return result


def normalize_factual_education(data):
    result = dict(data)
    qual = dict(result.get("highest_qualification", {}) or {})
    d_raw = str(qual.get("degree", "") or "").strip()
    c_raw = str(qual.get("course", "") or "").strip()
    comb = normalize_master_text(str(qual.get("qualification_text", "") or f"{d_raw} {c_raw}"))

    d_full, c_short = d_raw, c_raw
    if any(t in comb for t in ["bachelor of engineering", "b e", "be civil", "be mechanical"]):
        d_full, c_short = "Bachelor of Engineering", "B.E."
    elif any(t in comb for t in ["bachelor of technology", "b tech", "btech"]):
        d_full, c_short = "Bachelor of Technology", "B.Tech"
    elif any(t in comb for t in ["master of technology", "m tech", "mtech"]):
        d_full, c_short = "Master of Technology", "M.Tech"
    elif any(t in comb for t in ["master of business administration", "mba"]):
        d_full, c_short = "Master of Business Administration", "MBA"
    elif "diploma" in comb:
        d_full, c_short = "Diploma", "Diploma"

    qual["degree"] = d_full
    qual["course"] = c_short
    result["highest_qualification"] = qual
    return result


# ------------------------------------------------------------
# TAXONOMY & STRICT CONFIRMATION MATCHING
# ------------------------------------------------------------
@st.cache_data
def load_skill_groomers_master_data():
    if not os.path.exists(MASTER_DATA_FILE):
        raise FileNotFoundError(f"Missing {MASTER_DATA_FILE} at {BASE_DIR}")
    with open(MASTER_DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


MASTER_DATA = load_skill_groomers_master_data()
CORE_ROLE_MASTER = MASTER_DATA["core_roles"]
KEYWORD_MASTER = MASTER_DATA["keywords"]
INDUSTRY_MASTER = MASTER_DATA["industries"]
LOCATION_MASTER = MASTER_DATA["locations"]
FUNCTIONAL_AREA_MASTER = MASTER_DATA["functional_area_records"]
EDUCATION_MASTER = MASTER_DATA["education_records"]


def normalize_master_text(value):
    v = str(value or "").strip().lower().replace("&", " and ").replace("–", "-").replace("—", "-")
    v = re.sub(r"[/,_\-]+", " ", v)
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s]+", " ", v)).strip()


def tokens(value):
    return {t for t in normalize_master_text(value).split() if len(t) > 1}


def text_similarity(left, right):
    l_n, r_n = normalize_master_text(left), normalize_master_text(right)
    if not l_n or not r_n: return 0.0
    if l_n == r_n: return 1.0
    if l_n in r_n or r_n in l_n:
        return max(0.86, min(len(l_n), len(r_n)) / max(len(l_n), len(r_n)))
    t_l, t_r = tokens(l_n), tokens(r_n)
    t_score = len(t_l & t_r) / len(t_l | t_r) if (t_l and t_r) else 0.0
    return (0.58 * SequenceMatcher(None, l_n, r_n).ratio()) + (0.42 * t_score)


def best_named_match(value, names, minimum_score=0.82):
    best_n, best_s = "", 0.0
    for name in names:
        s = text_similarity(value, name)
        if s > best_s:
            best_n, best_s = name, s
    return (best_n, best_s) if best_s >= minimum_score else ("", best_s)


def map_core_role(data, source_text=""):
    raw = str(data.get("core_role", "") or "").strip()
    extracted_role = str(data.get("role", "") or "").strip()
    
    # 1. Gather designations with strict recency (Current job first, never oldest)
    jobs = sort_jobs(data.get("work_experience", []))
    cur_job, prev_job = get_current_and_previous_jobs(jobs)
    
    cur_title = cur_job.get("designation", "") if cur_job else ""
    prev_title = prev_job.get("designation", "") if prev_job else ""
    recent_titles = [j.get("designation", "") for j in reversed(jobs[-4:])] if jobs else []

    title_candidates = []
    if cur_title: title_candidates.append(cur_title)
    if extracted_role and extracted_role not in title_candidates: title_candidates.append(extracted_role)
    if raw and raw not in title_candidates: title_candidates.append(raw)
    if prev_title and prev_title not in title_candidates: title_candidates.append(prev_title)
    for t in recent_titles:
        if t and t not in title_candidates: title_candidates.append(t)

    norm_master_map = {normalize_master_text(r): r for r in CORE_ROLE_MASTER}

    # Step A: Direct exact match against any normalized master role
    for t in title_candidates:
        t_n = normalize_master_text(t)
        if t_n in norm_master_map:
            return {"name": norm_master_map[t_n], "score": 1.0, "matched": True, "review_required": False}

    # Step B: Domain Rule Mapping (Specific titles first, followed by broad disciplines)
    domain_rules = [
        # Quantity Surveying & Cost Estimation (handles database typo "Quantity Suveyor")
        (["csa estimator"], "CSA Estimator"),
        (["quantity estimator"], "Quantity Estimator"),
        (["quantity estimation"], "Quantity Estimation"),
        (["quantity surveyor", "quantity suveyor", "qs engineer", "senior qs", "jr qs", "sr qs", "qs"], "Quantity Suveyor"),
        (["estimator", "estimation engineer", "cost estimation", "cost estimator"], "Estimator"),

        # Rebar Detailing
        (["rebar detailer", "rebar draftsman", "rebar detailing", "rebar lead", "rc detailer", "bar bending detailer"], "Rebar Detailing"),

        # Safety, Health & Environment
        (["corporate safety manager"], "corporate safety manager"),
        (["deputy safety manager"], "Deputy safety manager"),
        (["safety manager", "hse manager", "ehs manager"], "safety manager"),
        (["safety officer", "hse officer", "ehs officer", "safety engineer"], "Safety Officer"),
        (["scaffolding inspector"], "scaffolding Inspector"),
        (["safety", "hse", "ehs", "fire and safety", "fire safety"], "Safety"),

        # Planning, Scheduling & Controls
        (["manager planning and coordination", "manager planning"], "Manager Planning & Coordination"),
        (["planning engineer", "planning manager", "project planner", "lead planner", "scheduling engineer", "p6 planner"], "Planning"),
        (["pmo analyst", "pmo lead", "pmo manager", "pmo"], "PMO"),

        # Contracts, Tendering & Procurement
        (["contracts and procurement", "tendering and contracts", "procurement engineer"], "Contracts & Procurement"),
        (["contracts engineer", "contracts manager", "tendering engineer", "contracts"], "Contracts"),
        (["billing engineer", "client billing", "subcontractor billing", "billing manager", "billing"], "Billing"),

        # Quality Assurance / Quality Control
        (["qa qc manager", "qaqc manager", "quality manager", "head quality"], "QA/QC Manager"),
        (["quality engineer", "sr quality engineer", "qa engineer", "qc engineer"], "Quality Engineer"),
        (["quality control", "qc inspector", "quality inspector"], "Quality control"),
        (["qa qc", "qaqc", "quality assurance"], "QAQC"),

        # MEP / Building Services
        (["mep engineer", "mep coordinator", "mep manager", "mep site", "hvac engineer", "plumbing engineer", "electrical mep", "mep"], "MEP"),

        # Finishing, Fit-outs & Interiors
        (["finishing supervisor", "fit out supervisor", "interior supervisor"], "Finishing Supervisor"),
        (["finishing engineer", "fit out engineer", "finishing", "fitout", "interior fit out"], "Finishing"),

        # Site Engineering & Execution
        (["civil supervisor", "site supervisor civil", "general supervisor"], "civil supervisor"),
        (["site engineer civil", "civil site engineer", "site engineer"], "Site Engineer"),
        (["civil engineer", "senior civil engineer", "project civil engineer"], "Civil Engineer"),
        (["execution engineer", "execution manager", "site execution", "execution"], "Execution"),

        # Project & Construction Management
        (["project director"], "Project Director"),
        (["project manager infra", "infra project manager"], "Project Manager - Infra"),
        (["project manager", "senior project manager", "assistant project manager"], "Project Manager"),
        (["construction manager", "construction lead"], "Construction Manager"),
        (["project coordinator"], "Project coordinator"),
        (["project engineer", "senior project engineer"], "Project Engineer"),
        (["project management", "project or construction management"], "Project Management"),

        # Architecture & Design
        (["senior architect", "lead architect"], "Senior Architect"),
        (["architect", "project architect", "junior architect", "architectural"], "Architect"),
        (["design engineer", "design manager", "structural design", "design"], "Design"),

        # Non-Civil, Plant & Corporate Functions
        (["geotechnical engineer", "geotech engineer", "soil engineer"], "Geotechnical"),
        (["plant and machinery", "plant and equipment", "pnm engineer", "pnm"], "plant & Machinery"),
        (["talent acquisition", "recruiter", "recruitment specialist"], "Talent Acquisition/ Recruitment"),
        (["am hr", "assistant manager hr"], "AM HR"),
        (["human resources", "hr executive", "hr manager", "hr"], "HR"),
        (["business development manager", "bdm"], "Business Development Manager"),
        (["business development", "bdr"], "Business Development"),
        (["accounts", "accountant", "senior accountant"], "accounts"),
        (["finance manager", "finance lead", "finance"], "Finance"),
        (["sales manager", "sales executive", "sales"], "sales"),
    ]

    for title in title_candidates:
        title_norm = normalize_master_text(title)
        for triggers, target in domain_rules:
            for tr in triggers:
                if re.search(r"\b" + re.escape(tr) + r"\b", title_norm):
                    if target in CORE_ROLE_MASTER:
                        return {"name": target, "score": 0.99, "matched": True, "review_required": False}

    # Step C: Fallback fuzzy matching on raw extracted core role
    if raw:
        name, score = best_named_match(raw, CORE_ROLE_MASTER, minimum_score=0.78)
        if name:
            return {"name": name, "score": round(score, 3), "matched": True, "review_required": False}

    return {"name": "", "score": 0.0, "matched": False, "review_required": True}


def map_functional_area(data, source_text=""):
    raw_fa = str(data.get("functional_area", "") or "")
    raw_role = str(data.get("role", "") or "")
    core_role = str(data.get("core_role", "") or "")
    
    # 1. Gather all title and discipline evidence
    jobs = sort_jobs(data.get("work_experience", []))
    cur_job, prev_job = get_current_and_previous_jobs(jobs)
    
    cur_title = cur_job.get("designation", "") if cur_job else ""
    evidence = normalize_master_text(f"{raw_fa} {raw_role} {core_role} {cur_title} {source_text}")
    title_n = normalize_master_text(f"{cur_title} {raw_role}")

    # Helper: Search FUNCTIONAL_AREA_MASTER for records matching specific keyword criteria
    def find_master_records(keywords):
        matches = []
        for r in FUNCTIONAL_AREA_MASTER:
            comb = normalize_master_text(f"{r.get('functional_area','')} {r.get('sub_functional_area','')} {r.get('role','')}")
            if any(k in comb for k in keywords):
                matches.append(r)
        return matches

    # ------------------------------------------------------------
    # RULE 1: QA / QC & QUALITY ASSURANCE / CONTROL
    # ------------------------------------------------------------
    qa_triggers = ["qa qc", "qaqc", "quality assurance", "quality control", "quality manager", "quality head", "project quality head", "quality engineer", "qc inspector"]
    if any(q in title_n or q in normalize_master_text(raw_fa) for q in qa_triggers):
        qa_recs = find_master_records(["qa qc", "qaqc", "quality"])
        if qa_recs:
            is_manager = any(m in title_n for m in ["manager", "head", "lead", "director", "chief", "senior manager", "sr manager"])
            if is_manager:
                # Prioritize Manager / Head level Quality record
                mgr_rec = next((r for r in qa_recs if any(m in normalize_master_text(r.get("role", "")) for m in ["manager", "head", "lead"])), None)
                if mgr_rec:
                    return {**mgr_rec, "functionalAreaId": mgr_rec.get("role_id"), "matched": True, "review_required": False}
            # Fallback to closest Quality record
            best_qa = next((r for r in qa_recs if "quality" in normalize_master_text(r.get("role", ""))), qa_recs[0])
            return {**best_qa, "functionalAreaId": best_qa.get("role_id"), "matched": True, "review_required": False}

    # ------------------------------------------------------------
    # RULE 2: SAFETY / HEALTH / ENVIRONMENT (HSE)
    # ------------------------------------------------------------
    if any(t in title_n or t in normalize_master_text(raw_fa) for t in ["safety officer", "safety engineer", "hse", "ehs", "safety manager", "fire safety"]):
        rec = next((r for r in FUNCTIONAL_AREA_MASTER if r.get("role_id") == 49), None)
        if not rec:
            safety_recs = find_master_records(["safety", "hse", "ehs"])
            rec = safety_recs[0] if safety_recs else None
        if rec:
            return {**rec, "functionalAreaId": rec.get("role_id", 49), "matched": True, "review_required": False}

    # ------------------------------------------------------------
    # RULE 3: PLANNING, SCHEDULING & REBAR DETAILING
    # ------------------------------------------------------------
    if any(t in title_n for t in ["rebar detailer", "planning engineer", "planning manager", "project planner", "scheduler", "p6"]):
        rec = next((r for r in FUNCTIONAL_AREA_MASTER if r.get("role_id") == 66), None)
        if rec:
            return {**rec, "functionalAreaId": 66, "matched": True, "review_required": False}

    # ------------------------------------------------------------
    # RULE 4: BILLING, CONTRACTS & QUANTITY SURVEYING
    # ------------------------------------------------------------
    if any(t in title_n for t in ["billing engineer", "billing", "quantity surveyor", "quantity suveyor", "estimation", "tendering"]):
        billing_recs = find_master_records(["billing", "quantity", "contracts"])
        if billing_recs:
            return {**billing_recs[0], "functionalAreaId": billing_recs[0].get("role_id"), "matched": True, "review_required": False}

    # ------------------------------------------------------------
    # RULE 5: MEP / HVAC / ELECTRICAL / MECHANICAL
    # ------------------------------------------------------------
    if "electrical" in title_n:
        rec = next((r for r in FUNCTIONAL_AREA_MASTER if r.get("role_id") == 60), None)
        if rec: return {**rec, "functionalAreaId": 60, "matched": True, "review_required": False}
    if any(t in title_n for t in ["mep", "hvac", "plumbing", "mechanical"]):
        rec = next((r for r in FUNCTIONAL_AREA_MASTER if r.get("role_id") == 62), None)
        if rec: return {**rec, "functionalAreaId": 62, "matched": True, "review_required": False}

    # ------------------------------------------------------------
    # RULE 6: CIVIL SITE ENGINEERING & PROJECT MANAGEMENT
    # ------------------------------------------------------------
    if any(t in title_n for t in ["project manager", "site engineer", "civil engineer"]) and any(t in evidence for t in ["civil", "construction"]):
        rec = next((r for r in FUNCTIONAL_AREA_MASTER if r.get("role_id") == 59), None)
        if rec:
            return {**rec, "functionalAreaId": 59, "matched": True, "review_required": False}

    # ------------------------------------------------------------
    # RULE 7: FALLBACK FUZZY & TOKEN-OVERLAP MATCHING
    # ------------------------------------------------------------
    best, best_s = None, 0.0
    for r in FUNCTIONAL_AREA_MASTER:
        r_fa = r.get("functional_area", "")
        r_role = r.get("role", "")
        
        # Check standard similarity
        s1 = text_similarity(raw_fa, r_fa)
        s2 = text_similarity(raw_role, r_role)
        s3 = text_similarity(cur_title, r_role)
        
        # Token containment check (avoids length penalties on long titles)
        tokens_input = tokens(f"{cur_title} {raw_role} {raw_fa}")
        tokens_master = tokens(f"{r_fa} {r_role}")
        t_overlap = len(tokens_input & tokens_master) / len(tokens_master) if tokens_master else 0.0
        
        s = max(s1, s2, s3, t_overlap)
        if s > best_s:
            best, best_s = r, s

    if best and best_s >= 0.75:
        return {**best, "functionalAreaId": best.get("role_id"), "matched": True, "review_required": False}

    return {"functionalAreaId": None, "functional_area": "", "sub_functional_area": "", "role": "", "matched": False, "review_required": True}


def map_industry(data):
    raw = str(data.get("industry", "") or "")
    evidence = normalize_master_text(f"{raw} {data.get('core_role','')} {data.get('functional_area','')}")
    if any(t in evidence for t in ["construction", "civil", "cement"]):
        rec = next((x for x in INDUSTRY_MASTER if x["id"] == 6), None)
        if rec: return {"industryId": 6, "name": rec["name"], "matched": True, "review_required": False}
    if any(t in evidence for t in ["real estate", "property"]):
        rec = next((x for x in INDUSTRY_MASTER if x["id"] == 16), None)
        if rec: return {"industryId": 16, "name": rec["name"], "matched": True, "review_required": False}

    name, s = best_named_match(raw, [i["name"] for i in INDUSTRY_MASTER if i.get("active", 1) == 1], minimum_score=0.82)
    rec = next((x for x in INDUSTRY_MASTER if x["name"] == name), None)
    if rec:
        return {"industryId": rec["id"], "name": rec["name"].strip(), "matched": True, "review_required": False}
    return {"industryId": None, "name": "", "matched": False, "review_required": True}


def map_keywords(data, source_text=""):
    raw_skills = data.get("key_skills", []) or []
    factual = clean_key_skills(raw_skills)

    # 1. Split compound skills (e.g. "AutoCAD, Revit" -> ["AutoCAD", "Revit"])
    candidate_skills = []
    for s in factual:
        parts = re.split(r"[,;/|]|\band\b", str(s), flags=re.I)
        for p in parts:
            p_clean = p.strip().strip("'\"").strip()
            if p_clean and p_clean not in candidate_skills:
                candidate_skills.append(p_clean)

    # 2. Normalized lookup table
    norm_to_master = {normalize_master_text(k): k.strip() for k in KEYWORD_MASTER if str(k).strip()}

    # 3. Domain alias dictionary
    skill_aliases = {
        "bar bending schedule": "BBS",
        "bar bending": "BBS",
        "bbs": "BBS",
        "primavera p6": "Primavera",
        "p6": "Primavera",
        "primavera": "Primavera",
        "microsoft project": "MSP",
        "ms project": "MSP",
        "msp": "MSP",
        "mivan formwork": "Mivan",
        "mivan shuttering": "Mivan",
        "mivan": "Mivan",
        "running account bills": "RA Bill",
        "running account bill": "RA Bill",
        "ra bills": "RA Bill",
        "ra billing": "RA Bill",
        "ra bill": "RA Bill",
        "client billing": "Billing",
        "subcontractor billing": "Billing",
        "quality assurance": "QAQC",
        "quality control": "QAQC",
        "qa qc": "QAQC",
        "qaqc": "QAQC",
        "autocad 2d": "AutoCAD",
        "autocad 3d": "AutoCAD",
        "autocad drafting": "AutoCAD",
        "cad": "AutoCAD",
        "autocad": "AutoCAD",
        "quantity takeoff": "Quantity Take-off",
        "quantity take off": "Quantity Take-off",
        "boq preparation": "BOQ",
        "bill of quantities": "BOQ",
        "boq": "BOQ",
        "rate analysis": "Rate Analysis",
        "estimation and costing": "Estimation",
        "cost estimation": "Estimation",
        "quantity survey": "Quantity Surveying",
        "quantity surveyor": "Quantity Surveying",
        "revit architecture": "Revit",
        "autodesk revit": "Revit",
        "revit": "Revit",
        "staad pro": "STAAD.Pro",
        "staad": "STAAD.Pro",
        "etabs": "ETABS",
        "ms office": "MS Office",
        "ms excel": "MS Excel",
        "advance excel": "MS Excel",
        "advanced excel": "MS Excel",
        "excel": "MS Excel",
    }

    selected = []

    def add_match(val):
        if val and val in KEYWORD_MASTER and val not in selected:
            selected.append(val)

    # 4. Resolve candidate skills against KEYWORD_MASTER
    for skill in candidate_skills:
        s_norm = normalize_master_text(skill)
        if not s_norm:
            continue

        # A. Direct match
        if s_norm in norm_to_master:
            add_match(norm_to_master[s_norm])
            continue

        # B. Alias match
        if s_norm in skill_aliases:
            target = skill_aliases[s_norm]
            target_norm = normalize_master_text(target)
            if target in KEYWORD_MASTER:
                add_match(target)
                continue
            elif target_norm in norm_to_master:
                add_match(norm_to_master[target_norm])
                continue

        # C. Phrase & word boundary match
        matched_kw = None
        for k_norm, k_orig in norm_to_master.items():
            if len(k_norm) < 3:
                continue
            if re.search(r"\b" + re.escape(k_norm) + r"\b", s_norm):
                matched_kw = k_orig
                break
            elif len(s_norm) >= 3 and re.search(r"\b" + re.escape(s_norm) + r"\b", k_norm):
                matched_kw = k_orig
                break

        if matched_kw:
            add_match(matched_kw)
            continue

        # D. High-confidence fuzzy match fallback (>= 0.86)
        best_n, best_s = best_named_match(skill, KEYWORD_MASTER, minimum_score=0.86)
        if best_n:
            add_match(best_n)

    return selected


def map_location(value, cat=None):
    v_n = normalize_master_text(value)
    if not v_n:
        return {"id": None, "name": "", "category": cat or "", "matched": False, "review_required": False}
    item = next((i for i in LOCATION_MASTER if i.get("active", 1) == 1 and (cat is None or i.get("category") == cat) and normalize_master_text(i["name"]) == v_n), None)
    if item:
        return {"id": item["id"], "name": item["name"].strip(), "category": item["category"].strip(), "matched": True, "review_required": False}
    return {"id": None, "name": "", "category": cat or "", "matched": False, "review_required": True}


def _match_course_family(comb):
    rules = [
        ([r"\bm\s*tech\b", r"\bmtech\b", r"\bmaster\s*(?:of\s*)?technology\b"], "Masters of Technology (M. Tech)"),
        ([r"\bm\s*sc\b", r"\bmsc\b", r"\bmaster\s*(?:of\s*)?science\b"], "Masters of Science (M.Sc)"),
        ([r"\bmba\b", r"\bpgdm\b", r"\bmaster\s*(?:of\s*)?business\s+administration\b"], "Masters of Business Administration (MBA / PGDM)"),
        ([r"\bmca\b", r"\bmaster\s*(?:of\s*)?computer\s+applications?\b"], "Masters of Computer Application (MCA)"),
        ([r"\bm\s*com\b", r"\bmcom\b", r"\bmaster\s*(?:of\s*)?commerce\b"], "Masters of Commerce (M.Com)"),
        ([r"\bb\s*e\b", r"\bbe\b", r"\bbachelor\s*(?:of\s*)?engineering\b", r"\bb\s*tech\b", r"\bbtech\b", r"\bbachelor\s*(?:of\s*)?technology\b"], "BE / B.Tech"),
        ([r"\bb\s*sc\b", r"\bbsc\b", r"\bbachelor\s*(?:of\s*)?science\b"], "Bachelor of Science (B.Sc)"),
        ([r"\bb\s*com\b", r"\bbcom\b", r"\bbachelor\s*(?:of\s*)?commerce\b"], "Bachelor of Commerce (B.Com)"),
        ([r"\bbba\b", r"\bbachelor\s*(?:of\s*)?business\s+administration\b"], "Bachelor of Business Administration (BBA)"),
        ([r"\bbca\b", r"\bbachelor\s*(?:of\s*)?computer\s+applications?\b"], "Bachelor of Computer Application (BCA)"),
        ([r"\bm\s*arch\b", r"\bmarch\b", r"\bmaster\s*(?:of\s*)?architecture\b"], "Masters of Arcitect (M. Arch.)"),
        ([r"\bb\s*arch\b", r"\bbarch\b", r"\bbachelor\s*(?:of\s*)?architecture\b"], "Bachor of Arcitect (B. Arch.)"),
        ([r"\bdiploma\b"], "Diploma"),
        ([r"\bhsc\b", r"\b12th\b"], "HSC"),
        ([r"\bssc\b", r"\b10th\b", r"\bmatriculation\b"], "SSC"),
        ([r"\biti\b"], "ITI"),
    ]
    for patterns, target in rules:
        if any(re.search(pat, comb) for pat in patterns):
            return target
    return ""


def map_education(data):
    qual = data.get("highest_qualification", {}) or {}
    degree = str(qual.get("degree", "") or "")
    course = str(qual.get("course", "") or "")
    spec = str(qual.get("specialization", "") or "").strip()
    year = qual.get("year", "")
    qtext = str(qual.get("qualification_text", "") or "").strip()

    comb = normalize_master_text(f"{degree} {course} {qtext}")
    spec_n = normalize_master_text(spec)

    pref_course = _match_course_family(comb)
    if not pref_course:
        return {
            "educationId": None,
            "qualification": "",
            "course": "",
            "specialization": "",
            "completionYear": str(year or ""),
            "matched": False,
            "course_matched": False,
            "review_required": True,
            "reason": "Course family not confirmed"
        }

    cands = [r for r in EDUCATION_MASTER if normalize_master_text(r.get("course", "")) == normalize_master_text(pref_course)]
    if not cands:
        return {
            "educationId": None,
            "qualification": "",
            "course": "",
            "specialization": "",
            "completionYear": str(year or ""),
            "matched": False,
            "course_matched": False,
            "review_required": True,
            "reason": "Course not found in master data"
        }

    cand_qualification = cands[0]["qualification"].strip()

    generic_words = {"polytechnic", "diploma", "engineering", "degree", "general", "polytechnique", "technical", "studies", "institute"}
    if not spec_n or spec_n in generic_words:
        return {
            "educationId": None,
            "qualification": cand_qualification,
            "course": pref_course,
            "specialization": "",
            "completionYear": str(year or ""),
            "matched": False,
            "course_matched": True,
            "review_required": True,
            "reason": f"Course '{pref_course}' confirmed, but specialization branch is missing in CV"
        }

    cleaned_spec = re.sub(r"\s+engineering$", "", spec_n).strip()

    exact = next((r for r in cands if normalize_master_text(r.get("specialization", "")) == spec_n), None)
    if not exact and cleaned_spec:
        exact = next((r for r in cands if normalize_master_text(r.get("specialization", "")) == cleaned_spec), None)

    if exact:
        return {
            "educationId": exact["education_id"],
            "qualification": exact["qualification"].strip(),
            "course": exact["course"].strip(),
            "specialization": exact["specialization"].strip(),
            "completionYear": str(year or ""),
            "matched": True,
            "course_matched": True,
            "review_required": False
        }

    return {
        "educationId": None,
        "qualification": cand_qualification,
        "course": pref_course,
        "specialization": "",
        "completionYear": str(year or ""),
        "matched": False,
        "course_matched": True,
        "review_required": True,
        "reason": f"Course '{pref_course}' confirmed, but specialization '{spec}' not found in master data"
    }


def build_skill_groomers_mapping(data, source_text=""):
    return {
        "core_role": map_core_role(data, source_text),
        "functional_area": map_functional_area(data, source_text),
        "industry": map_industry(data),
        "key_skills": map_keywords(data, source_text),
        "education": map_education(data),
        "current_city": map_location(data.get("current_city", ""), "Top Metropolitan Cities"),
        "current_state": map_location(data.get("current_state", ""), "States"),
        "native_city": map_location(data.get("native_city", ""), "Top Metropolitan Cities"),
        "native_state": map_location(data.get("native_state", ""), "States"),
    }


def apply_mapping_to_excel_data(data, mapping):
    excel_data = dict(data)
    excel_data["skill_groomers_ids"] = {
        "city": mapping["current_city"]["id"],
        "state": mapping["current_state"]["id"],
        "nativeCity": mapping["native_city"]["id"],
        "nativeState": mapping["native_state"]["id"],
        "functionalAreaId": mapping["functional_area"].get("functionalAreaId"),
        "industryId": mapping["industry"].get("industryId"),
        "educationId": mapping["education"].get("educationId"),
    }
    return excel_data


def populate_excel(data, sheet):
    sheet["C4"] = data.get("candidate_name", "")
    sheet["C5"] = data.get("email", "")
    sheet["C6"] = data.get("phone", "")
    sheet["C7"] = data.get("alternate_phone", "")
    sheet["C8"] = data.get("date_of_birth", "")
    sheet["C9"] = data.get("gender", "")
    sheet["C10"] = data.get("current_city", "")
    sheet["C11"] = data.get("current_state", "")
    sheet["C12"] = data.get("native_city", "")
    sheet["C13"] = data.get("native_state", "")
    sheet["C3"] = data.get("core_role", "")
    sheet["C14"] = ""
    sheet["C15"] = ", ".join(clean_key_skills(data.get("key_skills", [])))
    sheet["C16"] = data.get("functional_area", "")
    sheet["C17"] = data.get("role", "")
    sheet["C18"] = data.get("industry", "")

    jobs = sort_jobs(data.get("work_experience", []))
    cur, prev = get_current_and_previous_jobs(jobs)
    if cur:
        sheet["C20"] = cur.get("designation", "")
        sheet["C21"] = cur.get("company", "")
    if prev:
        sheet["C23"] = prev.get("designation", "")
        sheet["C24"] = prev.get("company", "")

    sheet["C26"] = experience_display(jobs)
    sheet["C27"] = datetime.today().strftime("%Y-%m-%d")
    sheet["C28"] = number_of_employers(jobs)
    sal = parse_salary_lakhs(data.get("annual_salary", ""))
    sheet["C29"] = sal if sal is not None else data.get("annual_salary", "")
    changes = compute_job_changes(jobs)
    sheet["C31"] = format_months(average_tenure_months(total_experience_months(jobs), changes))

    qual = data.get("highest_qualification", {})
    sheet["C33"] = qual.get("qualification_text", "") or qual.get("degree", "")
    sheet["C34"] = qual.get("year", "")
    sheet["C35"] = qual.get("course", "")
    sheet["C36"] = qual.get("specialization", "")


# ------------------------------------------------------------
# UI DISPLAY HELPERS
# ------------------------------------------------------------
def _info_grid(items, columns=3):
    cols = st.columns(columns)
    for idx, (label, val) in enumerate(items):
        with cols[idx % columns]:
            shown = val if val not in (None, "", [], {}) else "—"
            st.markdown(
                f"""
                <div style="border:1px solid #243048;border-radius:12px;padding:12px 14px;margin-bottom:12px;background:#151c2c;box-shadow:0 4px 16px rgba(0,0,0,0.25);min-height:74px;">
                    <div style="font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#818cf8;font-weight:750;margin-bottom:4px;">{label}</div>
                    <div style="font-size:14px;line-height:1.35;color:#f8fafc;font-weight:650;word-break:break-word;">{shown}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def _metric_card(label, val):
    shown = val if val not in (None, "", [], {}) else "—"
    st.markdown(
        f"""
        <div style="border:1px solid #243048;border-radius:14px;padding:13px 16px;background:linear-gradient(180deg,#151c2c,#101624);box-shadow:0 6px 20px rgba(0,0,0,0.3);min-height:76px;">
            <div style="font-size:12px;color:#94a3b8;font-weight:700;margin-bottom:4px;">{label}</div>
            <div style="font-size:21px;line-height:1.2;color:#ffffff;font-weight:800;">{shown}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _safe_index(options, value):
    try:
        return options.index(value)
    except ValueError:
        return 0


def _sg_date(val):
    if not val: return ""
    v = str(val).strip()
    if v.lower() in {"present", "current", "ongoing", "till date", "now"}:
        return datetime.today().strftime("%d/%m/%Y")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y-%m"):
        try: return datetime.strptime(v, fmt).strftime("%d/%m/%Y")
        except ValueError: continue
    return v


def _field_hint(cv_value, is_matched, unconfirmed_msg="Select closest master match"):
    """Displays a clean helper note under dropdowns showing raw CV wording."""
    if not cv_value:
        st.caption("ℹ️ *Not mentioned in CV*")
    elif not is_matched:
        st.markdown(
            f'<div style="font-size:12px;color:#fef08a;background:rgba(180,83,9,0.22);padding:6px 11px;border-radius:7px;margin:-4px 0 12px 0;border:1px solid #b45309;">'
            f'⚠️ <strong>CV states:</strong> "{cv_value}" — <em>{unconfirmed_msg}</em>'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div style="font-size:11px;color:#4ade80;margin:-4px 0 12px 2px;">'
            f'✓ Confirmed from CV: <em>"{cv_value}"</em>'
            f'</div>',
            unsafe_allow_html=True,
        )


SG_ADD_CANDIDATE_URL = "https://skillgroomers.projects-digitalgem.com/main/candidate/new"


def build_transfer_url(data, mapping, sel_core, sel_func, sel_ind, sel_kw, sel_city, sel_state, sel_native_city, sel_native_state, sel_edu):
    jobs = sort_jobs(data.get("work_experience", []))
    cur, prev = get_current_and_previous_jobs(jobs)
    tot_m = total_experience_months(jobs)
    y, m = split_months(tot_m)
    avg_y, avg_m = split_months(average_tenure_months(tot_m, compute_job_changes(jobs)))

    # Graceful partial education fallback: if recruiter didn't select an override, fill confirmed parts
    sg_edu = mapping.get("education", {}) or {}
    if sel_edu:
        edu_qual = sel_edu.get("qualification", "")
        edu_course = sel_edu.get("course", "")
        edu_spec = sel_edu.get("specialization", "")
    elif sg_edu.get("course_matched"):
        edu_qual = sg_edu.get("qualification", "")
        edu_course = sg_edu.get("course", "")
        edu_spec = sg_edu.get("specialization", "")
    else:
        edu_qual = ""
        edu_course = ""
        edu_spec = ""

    transfer = {
        "version": 1,
        "candidate": {
            "coreRole": sel_core if sel_core != "— Not selected —" else "",
            "fullName": str(data.get("candidate_name", "") or ""),
            "emailId": str(data.get("email", "") or ""),
            "mobileNumber": str(data.get("phone", "") or ""),
            "alternateNumber": str(data.get("alternate_phone", "") or ""),
            "dateOfBirth": _sg_date(data.get("date_of_birth")),
            "gender": str(data.get("gender", "") or ""),
            "currentCity": "" if sel_city == "— Not selected —" else sel_city,
            "currentState": "" if sel_state == "— Not selected —" else sel_state,
            "nativeCity": "" if sel_native_city == "— Not selected —" else sel_native_city,
            "nativeState": "" if sel_native_state == "— Not selected —" else sel_native_state,
            "keySkills": list(sel_kw or []),
            "functionalArea": sel_func.get("functional_area", "") if sel_func else "",
            "role": sel_func.get("role", "") if sel_func else "",
            "industry": "" if sel_ind == "— Not selected —" else sel_ind,
            "currentDesignation": cur.get("designation", "") if cur else "",
            "currentEmployer": cur.get("company", "") if cur else "",
            "startDate": _sg_date(cur.get("start_date")) if cur else "",
            "endDate": _sg_date(cur.get("end_date")) if cur else "",
            "previousDesignation": prev.get("designation", "") if prev else "",
            "previousEmployer": prev.get("company", "") if prev else "",
            "totalExperienceInYears": None if has_year_only_employment_dates(jobs) else y,
            "totalExperienceInMonths": None if has_year_only_employment_dates(jobs) else m,
            "totalExperienceAsOfDate": datetime.today().strftime("%d/%m/%Y"),
            "totalNumberOfJobs": str(number_of_employers(jobs)),
            "annualSalaryLakh": parse_salary_lakhs(data.get("annual_salary")),
            "averageTenureInYears": None if has_year_only_employment_dates(jobs) else avg_y,
            "averageTenureInMonths": None if has_year_only_employment_dates(jobs) else avg_m,
            "qualification": edu_qual,
            "course": edu_course,
            "specialization": edu_spec,
            "completionYear": str(data.get("highest_qualification", {}).get("year", "") or ""),
        }
    }
    raw = json.dumps(transfer, separators=(",", ":")).encode("utf-8")
    token = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    return f"{SG_ADD_CANDIDATE_URL}#cvfill={urllib.parse.quote(token)}"


def _render_review_card(filename, file_info):
    data = file_info["factual_data"]
    mapping = file_info["mapping"]
    candidate = file_info["candidate_name"]

    jobs = sort_jobs(data.get("work_experience", []))
    cur, prev = get_current_and_previous_jobs(jobs)
    tot_m = total_experience_months(jobs)
    factual_edu = data.get("highest_qualification", {}) or {}

    st.markdown(f'<div class="candidate-chip">QUICK REVIEW</div>', unsafe_allow_html=True)
    st.markdown(f'<h2 style="margin:0 0 6px 0;color:#ffffff;">{candidate}</h2>', unsafe_allow_html=True)

    st.markdown("### 1. Candidate Facts")
    m1, m2, m3, m4 = st.columns(4)
    with m1: _metric_card("Experience", experience_display(jobs))
    with m2: _metric_card("Avg. Tenure", format_months(average_tenure_months(tot_m, compute_job_changes(jobs))))
    with m3: _metric_card("Employers", number_of_employers(jobs))
    with m4: _metric_card("Job Changes", compute_job_changes(jobs))

    _info_grid([
        ("Candidate", data.get("candidate_name", "")),
        ("Email", data.get("email", "")),
        ("Mobile", data.get("phone", "")),
        ("Current Job", f"{cur.get('designation','')} — {cur.get('company','')}" if cur else "—"),
        ("Previous Job", f"{prev.get('designation','')} — {prev.get('company','')}" if prev else "—"),
        ("Education", f"{factual_edu.get('qualification_text','') or factual_edu.get('degree','')} ({factual_edu.get('year','')})"),
    ], columns=3)

    st.markdown("### 2. Skill Groomers Classification")
    sg_core = mapping.get("core_role", {}) or {}
    sg_func = mapping.get("functional_area", {}) or {}
    sg_ind = mapping.get("industry", {}) or {}
    sg_edu = mapping.get("education", {}) or {}

    f_path = " → ".join(x for x in [sg_func.get("functional_area"), sg_func.get("sub_functional_area"), sg_func.get("role")] if x) if sg_func.get("matched") else "— Not Confirmed (Left Blank) —"
    
    if sg_edu.get("matched"):
        e_path = f"{sg_edu.get('qualification')} → {sg_edu.get('course')} → {sg_edu.get('specialization')}"
    elif sg_edu.get("course_matched"):
        e_path = f"{sg_edu.get('qualification')} → {sg_edu.get('course')} → [⚠️ Select Specialization]"
    else:
        e_path = "— Not Confirmed (Left Blank) —"

    _info_grid([
        ("Core Role", sg_core.get("name", "") if sg_core.get("matched") else "— Not Confirmed —"),
        ("Industry", sg_ind.get("name", "") if sg_ind.get("matched") else "— Not Confirmed —"),
        ("Functional Area", f_path),
        ("Key Skills", ", ".join(mapping.get("key_skills", []) or [])),
        ("Education", e_path),
    ], columns=2)

    with st.expander("✏️ Edit Classifications (Manual Override)"):
        # 1. Core Role
        c_opts = ["— Not selected —"] + sorted({str(r).strip() for r in CORE_ROLE_MASTER if str(r).strip()}, key=str.lower)
        curr_core_val = sg_core.get("name", "") if sg_core.get("matched") else ""
        sel_core = st.selectbox("Core Role", c_opts, index=_safe_index(c_opts, curr_core_val), key=f"c_{filename}")
        _field_hint(data.get("core_role") or data.get("role"), sg_core.get("matched"))

        # 2. Functional Area
        fa_list = [f"{r['functional_area']} → {r['sub_functional_area']} → {r['role']}" for r in FUNCTIONAL_AREA_MASTER]
        fa_opts = ["— Not selected —"] + sorted(fa_list, key=str.lower)
        curr_fa = f"{sg_func.get('functional_area')} → {sg_func.get('sub_functional_area')} → {sg_func.get('role')}" if sg_func.get("matched") else ""
        sel_fa_lbl = st.selectbox("Functional Area", fa_opts, index=_safe_index(fa_opts, curr_fa), key=f"fa_{filename}")
        sel_func_rec = next((r for r in FUNCTIONAL_AREA_MASTER if f"{r['functional_area']} → {r['sub_functional_area']} → {r['role']}" == sel_fa_lbl), None)
        fa_cv_text = f"{data.get('functional_area', '')} (Role: {data.get('role', '')})".strip()
        _field_hint(fa_cv_text, sg_func.get("matched"))

        # 3. Industry
        ind_opts = ["— Not selected —"] + sorted({str(i['name']).strip() for i in INDUSTRY_MASTER if i.get('active', 1) == 1}, key=str.lower)
        curr_ind_val = sg_ind.get("name", "") if sg_ind.get("matched") else ""
        sel_ind = st.selectbox("Industry", ind_opts, index=_safe_index(ind_opts, curr_ind_val), key=f"ind_{filename}")
        _field_hint(data.get("industry"), sg_ind.get("matched"))

        # 4. Key Skills (Uncapped & Auto-selected)
        kw_opts = sorted({str(k).strip() for k in KEYWORD_MASTER if str(k).strip()}, key=str.lower)
        sel_kw = st.multiselect("Key Skills", kw_opts, default=mapping.get("key_skills", []), key=f"kw_{filename}")

        # Show all raw skills extracted directly from the CV PDF
        raw_cv_skills = clean_key_skills(data.get("key_skills", []) or [])
        if raw_cv_skills:
            st.markdown(
                f"""
                <div style="font-size:12px;color:#e2e8f0;background:#101624;padding:9px 13px;border-radius:9px;margin:5px 0 14px 0;border:1px solid #243048;line-height:1.65;">
                    <div style="color:#818cf8;font-weight:750;font-size:11px;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:3px;">
                        📄 All Key Skills Extracted From CV ({len(raw_cv_skills)}):
                    </div>
                    {" &nbsp;•&nbsp; ".join(f"<strong style='color:#f8fafc;'>{s}</strong>" for s in raw_cv_skills)}
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.caption("ℹ️ *No key skills extracted from CV*")

        # 5. Locations
        city_opts = ["— Not selected —"] + sorted({str(l['name']).strip() for l in LOCATION_MASTER if l.get('category') == 'Top Metropolitan Cities'}, key=str.lower)
        state_opts = ["— Not selected —"] + sorted({str(l['name']).strip() for l in LOCATION_MASTER if l.get('category') == 'States'}, key=str.lower)
        
        c1, c2 = st.columns(2)
        curr_city = mapping["current_city"].get("name", "") if mapping["current_city"].get("matched") else ""
        curr_state = mapping["current_state"].get("name", "") if mapping["current_state"].get("matched") else ""
        with c1: 
            sel_city = st.selectbox("Current City", city_opts, index=_safe_index(city_opts, curr_city), key=f"city_{filename}")
            _field_hint(data.get("current_city"), mapping["current_city"].get("matched"))
        with c2: 
            sel_state = st.selectbox("Current State", state_opts, index=_safe_index(state_opts, curr_state), key=f"state_{filename}")
            _field_hint(data.get("current_state"), mapping["current_state"].get("matched"))

        n1, n2 = st.columns(2)
        curr_native_city = mapping["native_city"].get("name", "") if mapping["native_city"].get("matched") else ""
        curr_native_state = mapping["native_state"].get("name", "") if mapping["native_state"].get("matched") else ""
        with n1: 
            sel_native_city = st.selectbox("Native City", city_opts, index=_safe_index(city_opts, curr_native_city), key=f"nc_{filename}")
            _field_hint(data.get("native_city"), mapping["native_city"].get("matched"))
        with n2: 
            sel_native_state = st.selectbox("Native State", state_opts, index=_safe_index(state_opts, curr_native_state), key=f"ns_{filename}")
            _field_hint(data.get("native_state"), mapping["native_state"].get("matched"))

        # 6. Education
        edu_list = [f"{e['qualification']} → {e['course']} → {e['specialization']}" for e in EDUCATION_MASTER]
        edu_opts = ["— Not selected —"] + sorted(edu_list, key=str.lower)
        curr_edu = f"{sg_edu.get('qualification')} → {sg_edu.get('course')} → {sg_edu.get('specialization')}" if sg_edu.get("matched") else ""
        sel_edu_lbl = st.selectbox("Education", edu_opts, index=_safe_index(edu_opts, curr_edu), key=f"edu_{filename}")
        sel_edu_rec = next((e for e in EDUCATION_MASTER if f"{e['qualification']} → {e['course']} → {e['specialization']}" == sel_edu_lbl), None)
        
        raw_edu = factual_edu.get("qualification_text") or f"{factual_edu.get('degree', '')} {factual_edu.get('specialization', '')}".strip()
        if sg_edu.get("matched"):
            _field_hint(raw_edu, True)
        elif sg_edu.get("course_matched"):
            _field_hint(
                raw_edu,
                False,
                unconfirmed_msg=f"Course '{sg_edu.get('course')}' confirmed. Qualification & Course will auto-fill in Skill Groomers (select specialization above if you wish to override now)."
            )
        else:
            _field_hint(raw_edu, False, unconfirmed_msg="Degree/course not recognized, please select manually")

    approve_key = f"approved_{filename}"
    if approve_key not in st.session_state:
        st.session_state[approve_key] = False

    st.markdown("### 3. Approval & Export")

    # Developer Debug Inspector Toggle
    dbg_key = f"debug_open_{filename}"
    if dbg_key not in st.session_state:
        st.session_state[dbg_key] = False

    col_app, col_dbg = st.columns([3, 1])
    with col_app:
        if st.button("✅ Approve Candidate", key=f"app_{filename}", use_container_width=True):
            st.session_state[approve_key] = True
            st.success("Candidate verified and approved for Skill Groomers.")
    with col_dbg:
        if st.button("🐞 Debug Info", key=f"dbg_btn_{filename}", use_container_width=True):
            st.session_state[dbg_key] = not st.session_state[dbg_key]

    # Resolve education strings for debugger & transfer payload
    if sel_edu_rec:
        deb_qual = sel_edu_rec.get("qualification", "")
        deb_course = sel_edu_rec.get("course", "")
        deb_spec = sel_edu_rec.get("specialization", "")
    elif sg_edu.get("course_matched"):
        deb_qual = sg_edu.get("qualification", "")
        deb_course = sg_edu.get("course", "")
        deb_spec = sg_edu.get("specialization", "")
    else:
        deb_qual, deb_course, deb_spec = "", "", ""

    if st.session_state[dbg_key]:
        st.markdown(
            """
            <div style="border: 1px dashed #6366f1; border-radius: 12px; padding: 10px 14px; margin: 10px 0; background: #0f172a;">
                <span style="color:#a5b4fc; font-weight:800; font-size:12px; text-transform:uppercase; letter-spacing:0.06em;">
                    🛠️ Developer Debug Inspector
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        tab_payload, tab_gemini, tab_mapping, tab_math = st.tabs([
            "1. Extension Payload",
            "2. Raw Gemini JSON",
            "3. Taxonomy Mapping",
            "4. Experience Math"
        ])

        with tab_payload:
            st.caption("Decoded JSON sent to the extension / bookmarklet:")
            y, m = split_months(tot_m)
            st.json({
                "coreRole": sel_core if sel_core != "— Not selected —" else "",
                "fullName": data.get("candidate_name", ""),
                "emailId": data.get("email", ""),
                "mobileNumber": data.get("phone", ""),
                "alternateNumber": data.get("alternate_phone", ""),
                "dateOfBirth": _sg_date(data.get("date_of_birth")),
                "gender": data.get("gender", ""),
                "currentCity": "" if sel_city == "— Not selected —" else sel_city,
                "currentState": "" if sel_state == "— Not selected —" else sel_state,
                "nativeCity": "" if sel_native_city == "— Not selected —" else sel_native_city,
                "nativeState": "" if sel_native_state == "— Not selected —" else sel_native_state,
                "keySkills": list(sel_kw or []),
                "functionalArea": sel_func_rec.get("functional_area", "") if sel_func_rec else "",
                "role": sel_func_rec.get("role", "") if sel_func_rec else "",
                "industry": "" if sel_ind == "— Not selected —" else sel_ind,
                "currentDesignation": cur.get("designation", "") if cur else "",
                "currentEmployer": cur.get("company", "") if cur else "",
                "totalExperienceInYears": None if has_year_only_employment_dates(jobs) else y,
                "totalExperienceInMonths": None if has_year_only_employment_dates(jobs) else m,
                "totalNumberOfJobs": str(number_of_employers(jobs)),
                "qualification": deb_qual,
                "course": deb_course,
                "specialization": deb_spec,
                "completionYear": str(data.get("highest_qualification", {}).get("year", "") or ""),
            })

        with tab_gemini:
            st.caption("Exact output from Gemini API before Python normalization:")
            st.json(data)

        with tab_mapping:
            st.caption("Taxonomy scores and database matching results:")
            st.json(mapping)

        with tab_math:
            st.caption("Experience calculation breakdown:")
            st.write({
                "Total Valid Months": tot_m,
                "Formatted Experience": experience_display(jobs),
                "Distinct Employers": number_of_employers(jobs),
                "Job Transitions": compute_job_changes(jobs),
                "Has Year-Only Dates": has_year_only_employment_dates(jobs),
                "Raw Jobs Count": len(jobs),
            })

    transfer_url = build_transfer_url(data, mapping, sel_core, sel_func_rec, sel_ind, sel_kw, sel_city, sel_state, sel_native_city, sel_native_state, sel_edu_rec)

    if st.session_state[approve_key]:
        st.markdown(
            f"""
            <a href="{transfer_url}" target="_blank"
               style="display:block;width:100%;text-align:center;padding:0.78rem 1.2rem;border-radius:12px;background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#ffffff;font-weight:750;text-decoration:none;margin-top:10px;box-shadow:0 8px 24px rgba(79,70,229,.35);">
               ⚡ Open Skill Groomers & Fill Candidate →
            </a>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.button("⚡ Open Skill Groomers & Fill Candidate →", key=f"disabled_sg_{filename}", disabled=True, use_container_width=True, help="Click Approve Candidate above first.")


# ------------------------------------------------------------
# MAIN STREAMLIT APP SHELL
# ------------------------------------------------------------
st.markdown(
    """
    <div class="app-hero">
        <div class="eyebrow">Skill Groomers • Recruitment Intelligence</div>
        <h1>CV Analyzer</h1>
        <p>Extract verified candidate facts, calculate deterministic experience metrics, and autofill directly into the recruitment platform.</p>
    </div>
    <div class="workflow">
        <div class="workflow-step"><span>Step 1</span>Upload CV</div>
        <div class="workflow-step"><span>Step 2</span>Analyse</div>
        <div class="workflow-step"><span>Step 3</span>Review Facts</div>
        <div class="workflow-step"><span>Step 4</span>Approve</div>
        <div class="workflow-step"><span>Step 5</span>Autofill Platform</div>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader("Upload candidate resumes (PDF)", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed", key="pdf_uploader")

if "generated_files" not in st.session_state:
    st.session_state.generated_files = {}

if st.button("Analyse Resumes", type="primary", use_container_width=True, key="analyse_resumes_btn"):
    if not uploaded_files:
        st.warning("Please upload at least one PDF resume.")
    else:
        st.session_state.generated_files = {}
        for uploaded_file in uploaded_files:
            try:
                with st.spinner(f"Analyzing {uploaded_file.name}..."):
                    pdf_bytes = uploaded_file.getvalue()
                    raw_text = extract_text_from_pdf(uploaded_file)
                    raw_gemini = extract_resume_data(raw_text, pdf_bytes=pdf_bytes)

                    raw_gemini = reconcile_qualification_text(raw_gemini)
                    data = normalize_factual_education(raw_gemini)
                    mapping = build_skill_groomers_mapping(data, source_text=raw_text)
                    excel_data = apply_mapping_to_excel_data(data, mapping)

                    template_path = os.path.join(BASE_DIR, "resume1.xlsx")
                    if os.path.exists(template_path):
                        wb = openpyxl.load_workbook(template_path)
                        populate_excel(excel_data, wb.active)
                        buf = io.BytesIO()
                        wb.save(buf)
                        excel_bytes = buf.getvalue()
                    else:
                        excel_bytes = b""

                    st.session_state.generated_files[uploaded_file.name] = {
                        "data": excel_bytes,
                        "candidate_name": data.get("candidate_name", uploaded_file.name),
                        "mapping": mapping,
                        "factual_data": data,
                    }
                st.success(f"✓ Ready: {uploaded_file.name}")
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")

if st.session_state.generated_files:
    st.divider()
    for filename, file_info in st.session_state.generated_files.items():
        with st.container():
            _render_review_card(filename, file_info)
            if file_info["data"]:
                st.download_button(
                    label=f"⬇️ Download Excel Report ({file_info['candidate_name']})",
                    data=file_info["data"],
                    file_name=f"{file_info['candidate_name']}_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{filename}",
                    use_container_width=True,
                )

st.markdown(
    """
    <div class="app-footer">
        CV Analyzer &nbsp;•&nbsp; Developed by <strong>Jai Pandya</strong>
    </div>
    """,
    unsafe_allow_html=True,
)
