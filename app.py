import io
import json
import re
import os
import base64
import urllib.parse
from difflib import SequenceMatcher
from datetime import datetime
import streamlit as st
import PyPDF2
import openpyxl
from google import genai
from google.genai import types

# Setup the Gemini API client using secure streamlit secrets
# Never hard-code API keys in source code.
# Put GEMINI_API_KEY in .streamlit/secrets.toml or in an environment variable.
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    st.error("Missing GEMINI_API_KEY. Add it to .streamlit/secrets.toml or your environment variables.")
    st.stop()

client = genai.Client(api_key=GEMINI_API_KEY)
MODEL = "gemini-3.5-flash-lite"

# Helper to read raw text content out of uploaded candidate PDF documents page-by-page.
# This text is useful for deterministic keyword/location rules, but some CVs
# contain dates/employer lines that are visually rendered yet missing from
# ordinary PDF text extraction. Gemini therefore also receives the original
# PDF bytes directly for factual extraction (see extract_resume_data).
def extract_text_from_pdf(file):
    pdf_bytes = file.getvalue() if hasattr(file, "getvalue") else file.read()
    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    text = []
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text.append(extracted)
    return "\n".join(text)

# Extended schema containing factual CV information needed
# for both the Excel report and future Skill Groomers integration.
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
        "key_skills": {
            "type": "array",
            "items": {"type": "string"}
        },
        "role": {"type": "string"},
        "industry": {"type": "string"},
        "highest_qualification": {
            "type": "object",
            "properties": {
                "degree": {"type": "string"},
                "course": {"type": "string"},
                "specialization": {"type": "string"},
                "institute": {"type": "string"},
                "year": {"type": "string"}
            },
            "required": ["degree", "course", "specialization", "institute", "year"]
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
    prompt = f"""
Extract structured HR information from the resume below.

The resume is the ONLY source of factual candidate information.
Return only information supported by the resume.
Never invent personal information, employers, dates, education, salary, skills, or other candidate facts.
Accuracy is more important than completeness.
If factual information is missing or uncertain, return an empty string "" for that field.

Do not calculate total experience, average tenure, number of employers, job changes, or employment gaps.
Python will perform those calculations.

============================================================
PERSONAL DETAILS
============================================================
Extract candidate_name, email, phone, alternate_phone, date_of_birth, and gender.
- date_of_birth: only if explicitly stated; prefer YYYY-MM-DD.
- gender: only if explicitly stated or unambiguous.
- If unavailable, return "".

============================================================
LOCATION
============================================================
Extract current_city, current_state, native_city, and native_state.
Do not treat current location as native location automatically.
Do not treat a job location as residential location automatically.
If uncertain, return "".

============================================================
PROFESSIONAL PROFILE
============================================================
core_role:
- Return a concise main professional/core role supported by the candidate's overall career history.
- Do NOT leave this blank when the resume contains enough professional experience to identify a main role.
- Prefer a broad, useful role rather than copying an unusually specific designation.
- Examples: Safety, Quantity Estimator, Planning Engineer, Civil Engineer, Project Manager, QA/QC, Recruitment, HR, MEP Engineer.
- If the resume genuinely provides no reliable professional-role evidence, return "".

functional_area:
- Return a concise human-readable functional area supported by the candidate's work history, responsibilities, role, and skills.
- Do NOT leave this blank when the resume provides enough professional evidence to determine a functional area.
- This is for the Excel report; it does NOT need to exactly match the Skill Groomers website hierarchy yet.
- Examples: Safety / Health / Environment, Project Management, Quantity Surveying / Estimation, Quality Assurance / Quality Control, Human Resources, Recruitment, MEP Engineering.
- Do not invent an unrelated area. If there is genuinely insufficient evidence, return "".

key_skills:
- Return the genuine important technical/professional skills supported by the resume.
- This factual list is for the Excel/report and is NOT limited to Skill Groomers' 4-keyword limit.
- Prioritize meaningful technical/professional skills and avoid weak duplicates.
- Do not invent skills.
- Do not include certifications, company names, job titles by themselves, hobbies, or generic personality traits.
- Skill Groomers will separately select a maximum of 4 supported master keywords.

role:
- Extract the candidate's current or most recent professional role/designation.
- Prefer wording supported by the resume.
- Keep this as the candidate's factual current/most recent designation.
- Do not replace it with the functional_area value.
- Exact Skill Groomers website classification will be mapped separately from master data.
- Do not exaggerate seniority.

industry:
- Primary industry based on recent/relevant employment history.
- Examples: Construction, Real Estate, IT, Recruitment, Oil and Gas, Manufacturing, Infrastructure.

============================================================
HIGHEST QUALIFICATION
============================================================
Return the highest COMPLETED qualification only.
Do not treat ongoing/pursuing education as completed.

Use these field meanings exactly:
- degree: FULL qualification name, e.g. "Bachelor of Engineering", "Bachelor of Science", "Master of Business Administration", "Diploma".
- course: qualification abbreviation/short course name, e.g. "B.E.", "B.Tech", "B.Sc", "MBA", "Diploma".
- specialization: field/branch only, e.g. "Civil Engineering", "Mechanical Engineering", "Physics", "Human Resources".
- institute: institution/university name.
- year: passing/completion year as YYYY.

Example: "B.E. Civil Engineering, 2018" should become:
degree = "Bachelor of Engineering"
course = "B.E."
specialization = "Civil Engineering"
year = "2018"

Do not put the specialization into the course field when the qualification abbreviation is available.
Do not duplicate the same specialization into both course and specialization.
Do not associate a nearby year or specialization with the wrong qualification.
If unavailable, return "".

============================================================
WORK EXPERIENCE
============================================================
Extract EVERY professional employment position in the resume.
Do not stop after recent jobs.
Include legitimate employment such as Graduate Engineer Trainee, Junior Engineer, Site Engineer, Engineer, Planning Engineer, Manager.
Do not include internships, academic/college projects, workshops, training programs, or volunteer work unless clearly represented as actual professional employment.
Return jobs in chronological order, OLDEST FIRST.
For every job extract company, designation, start_date, end_date, and country.
- start_date/end_date: YYYY-MM when month available; YYYY if only year available.
- Current employment end_date: "Present".
- country: actual country name only if supported; otherwise "".
Never deliberately reverse start and end dates.

============================================================
SALARY
============================================================
annual_salary:
- Extract current annual salary/CTC only if explicitly stated.
- Preserve useful salary information such as "₹8.5 LPA" or "850000".
- Do not estimate salary.
- If unavailable, return "".

============================================================
CERTIFICATIONS
============================================================
Return professional certifications as individual strings.
Do not count ordinary training, seminars, workshops, or memberships unless clearly presented as certifications.
If none are present, return [].

============================================================
FINAL ACCURACY RULES
============================================================
Never invent candidate personal details, employers, dates, designations, countries, qualifications, salary, certifications, or skills.
core_role, functional_area, and industry may summarize strongly supported professional history, but must not introduce experience the candidate does not have.
Do not leave core_role or functional_area blank merely because the exact wording is absent: infer a concise professional category when the resume's employment history, responsibilities, and skills clearly support it.
Exact Skill Groomers website Functional Area/Sub Functional Area/Role IDs will still be mapped separately from master data.

RESUME SOURCE:
Use the attached PDF as the primary source when one is provided.
The fallback extracted text is included only when no PDF bytes are available.
"""

    # Passing the actual PDF to Gemini preserves visually rendered dates,
    # employer names and multi-column layout that text extractors can omit.
    # If PDF input is unavailable, keep the previous text-only path.
    if pdf_bytes:
        contents = [
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt,
        ]
    else:
        contents = prompt + "\n\nRESUME TEXT:\n" + (raw_text or "")

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


# ---------- Helper functions for date parsing and experience calculations ----------

def parse_start_date(value):
    """Parse YYYY-MM or YYYY. Unknown dates sort to the beginning."""
    if not value or not isinstance(value, str):
        return datetime(1900, 1, 1)

    val = value.strip()

    for fmt in ("%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            pass

    match = re.fullmatch(r"(\d{4})", val)
    if match:
        return datetime(int(match.group(1)), 1, 1)

    return datetime(1900, 1, 1)


def parse_end_date(value):
    """Parse an employment end date. Present/current resolves to today."""
    if not value or not isinstance(value, str):
        return datetime.today()

    val = value.strip()
    lowered = val.lower()

    if lowered in ["present", "current", "ongoing", "till date", "now"]:
        return datetime.today()

    for fmt in ("%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            pass

    match = re.fullmatch(r"(\d{4})", val)
    if match:
        return datetime(int(match.group(1)), 12, 31)

    return datetime.today()


def get_job_months(job):
    """
    Calculate a job duration directly in whole months.

    This avoids converting years -> decimals -> months. For current jobs,
    the current partial month is counted once it has begun.
    """
    try:
        start = parse_start_date(job.get("start_date", ""))
        end_value = job.get("end_date", "")
        end = parse_end_date(end_value)

        if start.year == 1900 or end < start:
            return 0

        months = (end.year - start.year) * 12 + (end.month - start.month)

        # If exact day information is available, count the additional month
        # only after reaching the start day. For YYYY-MM input both dates use
        # day 1, so the completed month difference remains deterministic.
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(job.get("start_date", "")).strip()) and \
           re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(end_value).strip()):
            if end.day >= start.day:
                months += 1

        # For an ongoing job, include the current partial month.
        if isinstance(end_value, str) and end_value.strip().lower() in [
            "present", "current", "ongoing", "till date", "now"
        ]:
            months += 1

        return max(0, months)
    except Exception:
        return 0


def sort_jobs(jobs):
    return sorted(jobs, key=lambda j: parse_start_date(j.get("start_date", "")))


def _month_index(year, month):
    """Convert a calendar year/month into a sortable integer index."""
    return (year * 12) + (month - 1)


def _job_month_range(job, as_of=None):
    """
    Return the inclusive calendar-month range covered by one job.

    CVs normally provide employment dates only to month precision (YYYY-MM).
    Therefore July 2017 to December 2018 means every calendar month from
    2017-07 through 2018-12. Present/current jobs run through the current
    calendar month.
    """
    as_of = as_of or datetime.today()
    start = parse_start_date(job.get("start_date", ""))
    if start.year == 1900:
        return None

    end_value = str(job.get("end_date", "") or "").strip().lower()
    if end_value in ["present", "current", "ongoing", "till date", "now"]:
        end = as_of
    else:
        end = parse_end_date(job.get("end_date", ""))

    if end < start:
        return None

    return _month_index(start.year, start.month), _month_index(end.year, end.month)


def total_experience_months(jobs, as_of=None):
    """
    Calculate total professional experience as the UNION of calendar months.

    This deliberately avoids summing each job independently, because concurrent
    roles would otherwise double-count the same month. Gaps are naturally
    excluded. Each covered calendar month is counted once.
    """
    covered_months = set()
    for job in jobs:
        month_range = _job_month_range(job, as_of=as_of)
        if not month_range:
            continue
        start_idx, end_idx = month_range
        covered_months.update(range(start_idx, end_idx + 1))
    return len(covered_months)


def split_months(total_months):
    total_months = max(0, int(round(total_months)))
    return total_months // 12, total_months % 12


def format_months(total_months):
    years, months = split_months(total_months)
    parts = []
    if years:
        parts.append(f"{years} year{'s' if years != 1 else ''}")
    if months:
        parts.append(f"{months} month{'s' if months != 1 else ''}")
    return " ".join(parts) if parts else "0 years"

def number_of_employers(jobs):
    companies = set()
    for job in jobs:
        company = job.get("company", "").strip()
        if company:
            companies.add(company.lower())
    return len(companies)

def compute_job_changes(jobs):
    """Count number of job changes (consecutive distinct employers)."""
    if len(jobs) < 2:
        return 0
    changes = 0
    prev_company = None
    for job in jobs:
        curr = job.get("company", "").strip().lower()
        if not curr:
            continue
        if prev_company is not None and curr != prev_company:
            changes += 1
        prev_company = curr
    return changes

def average_tenure_months(total_months, job_changes):
    """Average tenure = total experience / job changes, kept in months."""
    if job_changes <= 0:
        return total_months
    return int(round(total_months / job_changes))



def get_current_and_previous_jobs(jobs):
    """Return the current/latest job and the previous DISTINCT employer."""
    sorted_jobs = sort_jobs(jobs)
    if not sorted_jobs:
        return None, None

    present_terms = {"present", "current", "ongoing", "till date", "now"}
    current = None
    for job in reversed(sorted_jobs):
        if str(job.get("end_date", "") or "").strip().lower() in present_terms:
            current = job
            break
    if current is None:
        current = sorted_jobs[-1]

    current_company = normalize_master_text(current.get("company", ""))
    previous = None
    for job in reversed(sorted_jobs):
        if job is current:
            continue
        company = normalize_master_text(job.get("company", ""))
        if company and company != current_company:
            previous = job
            break

    return current, previous

# Optional: post-process key_skills to remove known certification keywords
def clean_key_skills(skills):
    if not skills:
        return []

    cert_keywords = ["pmp", "six sigma", "leed ap", "prince2", "scrum", "itil", "ccna", "ccnp", "aws", "azure", "gcp"]
    cleaned = []

    seen = set()

    for skill in skills:
        if not isinstance(skill, str):
            continue

        skill = skill.strip()
        if not skill:
            continue

        if any(cert in skill.lower() for cert in cert_keywords):
            continue

        normalized = skill.lower()
        if normalized in seen:
            continue

        seen.add(normalized)
        cleaned.append(skill)

    # Keep all genuine factual skills for Excel. Skill Groomers mapping
    # separately selects a maximum of 4 supported master keywords.
    return cleaned


def normalize_factual_education(data):
    """
    Normalize ONLY the presentation/semantics of the highest completed
    qualification. This does not use Skill Groomers categories and does not
    invent a new qualification.

    Example supported by CV extraction:
      degree="B.E.", course="Civil Engineering", specialization="Civil Engineering"
    becomes:
      degree="Bachelor of Engineering", course="B.E.",
      specialization="Civil Engineering".
    """
    result = dict(data)
    qual = dict(result.get("highest_qualification", {}) or {})

    degree_raw = str(qual.get("degree", "") or "").strip()
    course_raw = str(qual.get("course", "") or "").strip()
    spec_raw = str(qual.get("specialization", "") or "").strip()
    combined = normalize_master_text(f"{degree_raw} {course_raw}")

    def contains_any(*terms):
        return any(term in combined for term in terms)

    # Degree/course normalization is only performed when the extracted text
    # itself contains evidence for that qualification family.
    degree_full = degree_raw
    course_short = course_raw

    if contains_any("bachelor of engineering", "bachelor engineering", "b e", "be civil", "be mechanical", "be electrical"):
        degree_full, course_short = "Bachelor of Engineering", "B.E."
    elif contains_any("bachelor of technology", "bachelor technology", "b tech", "btech"):
        degree_full, course_short = "Bachelor of Technology", "B.Tech"
    elif contains_any("master of engineering", "master engineering", "m e"):
        degree_full, course_short = "Master of Engineering", "M.E."
    elif contains_any("master of technology", "master technology", "m tech", "mtech"):
        degree_full, course_short = "Master of Technology", "M.Tech"
    elif contains_any("bachelor of science", "bachelor science", "b sc", "bsc"):
        degree_full, course_short = "Bachelor of Science", "B.Sc"
    elif contains_any("master of science", "master science", "m sc", "msc"):
        degree_full, course_short = "Master of Science", "M.Sc"
    elif contains_any("master of business administration", "mba"):
        degree_full, course_short = "Master of Business Administration", "MBA"
    elif contains_any("bachelor of commerce", "b com", "bcom"):
        degree_full, course_short = "Bachelor of Commerce", "B.Com"
    elif contains_any("master of commerce", "m com", "mcom"):
        degree_full, course_short = "Master of Commerce", "M.Com"
    elif contains_any("bachelor of business administration", "bba"):
        degree_full, course_short = "Bachelor of Business Administration", "BBA"
    elif contains_any("bachelor of computer applications", "bca"):
        degree_full, course_short = "Bachelor of Computer Applications", "BCA"
    elif contains_any("master of computer applications", "mca"):
        degree_full, course_short = "Master of Computer Applications", "MCA"
    elif "diploma" in combined:
        degree_full, course_short = "Diploma", "Diploma"

    # If Gemini placed a clear branch/field in course while a qualification
    # family is known, preserve that field as specialization instead.
    course_n = normalize_master_text(course_raw)
    spec_n = normalize_master_text(spec_raw)
    field_terms = [
        "civil", "mechanical", "electrical", "electronics", "computer",
        "information technology", "telecom", "production", "physics",
        "chemistry", "biology", "mathematics", "commerce", "finance",
        "human resources", "marketing", "architecture", "agriculture",
    ]
    if course_short != course_raw and not spec_raw and any(term in course_n for term in field_terms):
        spec_raw = course_raw
        spec_n = course_n

    # If course and specialization contain the same field, keep it only as
    # specialization once a true qualification short name is known.
    if course_short != course_raw and spec_raw and course_n == spec_n:
        pass  # course_short already holds the qualification abbreviation.

    qual["degree"] = degree_full
    qual["course"] = course_short
    qual["specialization"] = spec_raw
    result["highest_qualification"] = qual
    return result


# ============================================================
# SKILL GROOMERS MASTER-DATA MAPPER — V6
# ============================================================
# The website master values below are loaded from the complete snapshots
# collected from:
#   /corerole
#   /keywords
#   /functionalArea
#   /industry
#   /education
#   /location
#
# Gemini extracts factual / human-readable CV information.
# Python then maps that evidence to exact Skill Groomers values and IDs.
# The original website spelling is preserved in all returned mappings.

MASTER_DATA_FILE = "skillgroomers_master_data_v6.json"


def load_skill_groomers_master_data():
    if not os.path.exists(MASTER_DATA_FILE):
        raise FileNotFoundError(
            f"Missing {MASTER_DATA_FILE}. Keep it in the same folder as this app."
        )

    with open(MASTER_DATA_FILE, "r", encoding="utf-8") as file:
        master = json.load(file)

    required = [
        "core_roles",
        "keywords",
        "industries",
        "locations",
        "functional_area_records",
        "education_records",
    ]
    missing = [key for key in required if key not in master]
    if missing:
        raise ValueError(
            f"{MASTER_DATA_FILE} is missing required sections: {', '.join(missing)}"
        )
    return master


MASTER_DATA = load_skill_groomers_master_data()
CORE_ROLE_MASTER = MASTER_DATA["core_roles"]
KEYWORD_MASTER = MASTER_DATA["keywords"]
INDUSTRY_MASTER = MASTER_DATA["industries"]
LOCATION_MASTER = MASTER_DATA["locations"]
FUNCTIONAL_AREA_MASTER = MASTER_DATA["functional_area_records"]
EDUCATION_MASTER = MASTER_DATA["education_records"]


def normalize_master_text(value):
    """Normalize only for comparison; never submit this normalized wording."""
    value = str(value or "").strip().lower()
    value = value.replace("&", " and ")
    value = value.replace("–", "-").replace("—", "-")
    value = re.sub(r"[/,_\-]+", " ", value)
    # Punctuation is ignored for matching, so B.E. == BE and B.Tech == B Tech.
    value = re.sub(r"[^a-z0-9\s]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def tokens(value):
    return {token for token in normalize_master_text(value).split() if len(token) > 1}


def text_similarity(left, right):
    left_n = normalize_master_text(left)
    right_n = normalize_master_text(right)

    if not left_n or not right_n:
        return 0.0

    if left_n == right_n:
        return 1.0

    if left_n in right_n or right_n in left_n:
        shorter = min(len(left_n), len(right_n))
        longer = max(len(left_n), len(right_n))
        containment = shorter / longer if longer else 0
        return max(0.86, containment)

    left_tokens = tokens(left_n)
    right_tokens = tokens(right_n)
    token_score = 0.0
    if left_tokens and right_tokens:
        intersection = len(left_tokens & right_tokens)
        union = len(left_tokens | right_tokens)
        token_score = intersection / union if union else 0.0

    sequence_score = SequenceMatcher(None, left_n, right_n).ratio()
    return (0.58 * sequence_score) + (0.42 * token_score)


def best_named_match(value, names, minimum_score=0.58):
    best_name = ""
    best_score = 0.0

    for name in names:
        score = text_similarity(value, name)
        if score > best_score:
            best_name = name
            best_score = score

    if best_score < minimum_score:
        return "", best_score

    return best_name, best_score


def _combined_professional_evidence(data):
    fields = [
        data.get("core_role", ""),
        data.get("functional_area", ""),
        data.get("role", ""),
        data.get("industry", ""),
    ]
    fields.extend(data.get("key_skills", []) or [])

    jobs = data.get("work_experience", []) or []
    for job in jobs[-3:]:
        fields.append(job.get("designation", ""))
        fields.append(job.get("company", ""))

    return " ".join(str(v) for v in fields if v)


def _designation_evidence(data):
    """Role-centric evidence used for deterministic primary-role decisions.

    Skills such as QA, Safety or Project Management can appear in many CVs as
    supporting capabilities. They should not override the candidate's actual
    profession unless the designation/core role also supports them.
    """
    jobs = sort_jobs(data.get("work_experience", []) or [])
    recent_designations = [str(j.get("designation", "") or "") for j in jobs[-3:]]
    return normalize_master_text(" ".join([
        str(data.get("role", "") or ""),
        str(data.get("core_role", "") or ""),
        *recent_designations,
    ]))


def map_core_role(data, source_text=""):
    """Return an exact Skill Groomers core-role name plus confidence."""
    raw_core = str(data.get("core_role", "") or "")
    role = str(data.get("role", "") or "")
    skills = data.get("key_skills", []) or []
    evidence = normalize_master_text(" ".join([raw_core, role, *skills]))
    designation_evidence = _designation_evidence(data)
    source_n = normalize_master_text(source_text or "")

    # Strong profession signals from current/recent designations should win
    # over generic supporting skills. This prevents a civil/fitout engineer
    # from becoming QAQC merely because Quality Assurance appears in skills.
    if "billing" in designation_evidence:
        target = "Billing"
        if target in CORE_ROLE_MASTER:
            return {
                "name": target, "score": 0.995, "matched": True,
                "review_required": False,
                "reason": "Current/recent designation is primarily Billing",
            }

    fitout_or_civil = any(t in designation_evidence for t in [
        "fitout engineer", "fit out engineer", "fitout executive",
        "executive engineer", "civil engineer", "site engineer"
    ])
    civil_context = any(t in (designation_evidence + " " + source_n) for t in [
        "civil", "construction", "interior", "fitout", "fit out", "site execution"
    ])
    if fitout_or_civil and civil_context:
        target = "Civil Engineer"
        if target in CORE_ROLE_MASTER:
            return {
                "name": target, "score": 0.99, "matched": True,
                "review_required": False,
                "reason": "Civil/fitout/site-engineering designation evidence",
            }

    # Exact website value from AI is accepted only after stronger designation
    # rules above have had a chance to correct overly broad classifications.
    for name in CORE_ROLE_MASTER:
        if normalize_master_text(raw_core) == normalize_master_text(name):
            return {
                "name": name,
                "score": 1.0,
                "matched": True,
                "review_required": False,
                "reason": "Exact master match",
            }

    # Domain rules based on confirmed production classifications.
    # Construction QA/QC profiles consistently use the exact SG core role QAQC.
    if any(term in designation_evidence for term in ["qaqc", "qa qc", "quality engineer", "quality control engineer", "quality assurance engineer"]):
        target = "QAQC"
        if target in CORE_ROLE_MASTER:
            return {
                "name": target,
                "score": 0.99,
                "matched": True,
                "review_required": False,
                "reason": "QA/QC professional evidence",
            }

    if any(term in designation_evidence for term in ["safety", "ehs", "hse"]):
        target = "Safety"
        if target in CORE_ROLE_MASTER:
            return {
                "name": target,
                "score": 0.98,
                "matched": True,
                "review_required": False,
                "reason": "Safety/EHS evidence",
            }

    # Quantity Surveyor profiles in the existing production data use Estimator.
    if any(term in designation_evidence for term in ["quantity surveyor", "quantity surveying", " qs "]):
        target = "Estimator"
        if target in CORE_ROLE_MASTER:
            return {
                "name": target,
                "score": 0.92,
                "matched": True,
                "review_required": True,
                "reason": "Quantity-surveying profile; production mapping commonly uses Estimator",
            }

    if "quantity estimator" in designation_evidence:
        target = "Quantity Estimator"
        if target in CORE_ROLE_MASTER:
            return {
                "name": target,
                "score": 0.96,
                "matched": True,
                "review_required": False,
                "reason": "Explicit quantity-estimator evidence",
            }

    candidates = [raw_core, role, *skills]
    best_name = ""
    best_score = 0.0
    for candidate in candidates:
        name, score = best_named_match(candidate, CORE_ROLE_MASTER, minimum_score=0.54)
        if score > best_score:
            best_name, best_score = name, score

    return {
        "name": best_name,
        "score": round(best_score, 3),
        "matched": bool(best_name),
        "review_required": bool(best_name and best_score < 0.82),
        "reason": "Best master similarity" if best_name else "No reliable master match",
    }


def _functional_record(record):
    return {
        "functionalAreaId": record.get("role_id"),
        "functional_area": record.get("functional_area", ""),
        "sub_functional_area": record.get("sub_functional_area", ""),
        "role": record.get("role", ""),
    }


def _find_functional_id(role_id):
    for record in FUNCTIONAL_AREA_MASTER:
        if record.get("role_id") == role_id:
            return record
    return None


def map_functional_area(data, source_text=""):
    """
    Map to one complete website hierarchy.
    The returned Functional Area, Sub Functional Area, Role and ID
    always come from the SAME website master record.
    """
    evidence = _combined_professional_evidence(data)
    evidence_n = normalize_master_text(evidence + " " + (source_text or ""))
    designation_n = _designation_evidence(data)

    # Explicit civil/fitout/site-engineering profiles map to the civil/site
    # hierarchy, not to the telecom-specific Project Manager record.
    fitout_civil_terms = [
        "fitout engineer", "fit out engineer", "fitout executive",
        "civil engineer", "site engineer", "executive engineer"
    ]
    if any(term in designation_n for term in fitout_civil_terms) and any(
        term in evidence_n for term in ["civil", "construction", "interior", "fitout", "site"]
    ):
        record = _find_functional_id(59)
        if record:
            return {
                **_functional_record(record),
                "score": 0.99,
                "matched": True,
                "review_required": False,
                "reason": "Civil/fitout/site-engineering profile",
            }

    # High-confidence company-specific semantic rules.
    # Safety/EHS -> confirmed ID 49.
    if any(term in evidence_n for term in ["safety", "ehs", "hse", "health safety environment"]):
        record = _find_functional_id(49)
        if record:
            return {
                **_functional_record(record),
                "score": 0.99,
                "matched": True,
                "review_required": False,
                "reason": "Safety/EHS profile",
            }

    # Construction planning / project-controls profiles -> ID 66.
    # This handles candidates whose factual role is Planning / Project Controls
    # but whose SG hierarchy belongs to construction/site engineering rather
    # than the telecom-specific Project Manager record (ID 55).
    planning_terms = [
        "planning", "project planning", "project controls", "project control",
        "project coordination", "planning manager", "planning engineer", "msp",
        "primavera", "critical path"
    ]
    construction_planning_terms = [
        "construction", "civil", "site", "residential", "building",
        "real estate", "contractor", "project management"
    ]
    if any(term in evidence_n for term in planning_terms) and any(
        term in evidence_n for term in construction_planning_terms
    ):
        record = _find_functional_id(66)
        if record:
            return {
                **_functional_record(record),
                "score": 0.97,
                "matched": True,
                "review_required": False,
                "reason": "Construction planning/project-controls profile",
            }

    # QA/QC. Construction/site QAQC profiles in the validated SG records
    # use Project Management / Site Engineers -> Construction... (ID 66).
    qa_terms = ["qaqc", "qa qc", "quality control", "quality assurance", "quality engineer"]
    construction_terms = [
        "construction", "civil", "site", "residential", "real estate",
        "building", "rcc", "concrete", "project management"
    ]
    if any(term in evidence_n for term in qa_terms):
        target_id = 66 if any(term in evidence_n for term in construction_terms) else 45
        record = _find_functional_id(target_id)
        if record:
            return {
                **_functional_record(record),
                "score": 0.97 if target_id == 66 else 0.91,
                "matched": True,
                "review_required": False,
                "reason": (
                    "Construction/site QAQC profile; aligned to validated SG classifications"
                    if target_id == 66 else "QA/QC evidence"
                ),
            }


    # Quantity-surveying / estimation profiles: existing production example
    # is classified under Site Engineers -> Construction ... (ID 66).
    quantity_terms = [
        "quantity survey", "quantity surveying", "quantity take off", "estimation",
        "estimator", "billing", "boq", "earthwork quantification", "costx", "agtek"
    ]
    raw_functional = normalize_master_text(data.get("functional_area", ""))
    raw_role = normalize_master_text(data.get("role", ""))
    raw_core = normalize_master_text(data.get("core_role", ""))

    quantity_is_primary = (
        any(term in raw_functional for term in quantity_terms)
        or any(term in raw_role for term in quantity_terms)
        or any(term in raw_core for term in quantity_terms)
    )

    if quantity_is_primary:
        record = _find_functional_id(66)
        if record:
            return {
                **_functional_record(record),
                "score": 0.90,
                "matched": True,
                "review_required": True,
                "reason": "Quantity surveying/estimation profile; aligned to existing production classification",
            }

    # MEP branches.
    if any(term in evidence_n for term in ["hvac", "plumbing", "firefighting", "fire protection", "mechanical mep"]):
        record = _find_functional_id(62)
        if record:
            return {
                **_functional_record(record),
                "score": 0.92,
                "matched": True,
                "review_required": False,
                "reason": "Mechanical/MEP evidence",
            }

    if any(term in evidence_n for term in ["electrical engineer", "electrical mep", "revit mep", "ht lt"]):
        record = _find_functional_id(60)
        if record:
            return {
                **_functional_record(record),
                "score": 0.92,
                "matched": True,
                "review_required": False,
                "reason": "Electrical engineering evidence",
            }

    # HR / Recruitment.
    if any(term in evidence_n for term in ["recruitment", "talent acquisition", "recruiter"]):
        record = _find_functional_id(13)
        if record:
            return {
                **_functional_record(record),
                "score": 0.94,
                "matched": True,
                "review_required": False,
                "reason": "Recruitment evidence",
            }

    if any(term in evidence_n for term in ["human resources", " hr ", "hrbp", "employee relation", "payroll"]):
        record = _find_functional_id(12)
        if record:
            return {
                **_functional_record(record),
                "score": 0.90,
                "matched": True,
                "review_required": False,
                "reason": "HR evidence",
            }

    # Generic scoring over every allowed website hierarchy.
    best = None
    best_score = 0.0

    raw_fields = [
        data.get("functional_area", ""),
        data.get("role", ""),
        data.get("core_role", ""),
        *(data.get("key_skills", []) or []),
    ]

    explicit_telecom = any(
        term in evidence_n
        for term in ["telecom", "telecommunication", "fiber optic", "fibre optic", "rf engineer", "telecom engineer"]
    )

    for record in FUNCTIONAL_AREA_MASTER:
        # Do not let generic "Project Management" similarity select a
        # telecom-only website role when the CV contains no telecom evidence.
        if record.get("role_id") in {55, 58} and not explicit_telecom:
            continue

        target_parts = [
            record.get("functional_area", ""),
            record.get("sub_functional_area", ""),
            record.get("role", ""),
        ]

        score = 0.0
        # Match each CV field against each hierarchy component.
        for raw in raw_fields:
            if not raw:
                continue
            local = max(text_similarity(raw, target) for target in target_parts)
            score = max(score, local)

        # Small bonus if several meaningful tokens occur in the combined evidence.
        target_tokens = tokens(" ".join(target_parts))
        evidence_tokens = tokens(evidence_n)
        overlap = len(target_tokens & evidence_tokens)
        if overlap >= 2:
            score = min(1.0, score + 0.08)
        elif overlap == 1:
            score = min(1.0, score + 0.03)

        if score > best_score:
            best, best_score = record, score

    if not best or best_score < 0.58:
        return {
            "functionalAreaId": None,
            "functional_area": "",
            "sub_functional_area": "",
            "role": "",
            "score": round(best_score, 3),
            "matched": False,
            "review_required": True,
            "reason": "No reliable hierarchy match",
        }

    return {
        **_functional_record(best),
        "score": round(best_score, 3),
        "matched": True,
        "review_required": best_score < 0.80,
        "reason": "Best complete-hierarchy master match",
    }


def map_industry(data):
    """Map extracted industry to exact website name and industryId."""
    raw = str(data.get("industry", "") or "")
    evidence = normalize_master_text(
        " ".join([
            raw,
            str(data.get("core_role", "") or ""),
            str(data.get("functional_area", "") or ""),
        ])
    )

    # Exact match first.
    for record in INDUSTRY_MASTER:
        if normalize_master_text(raw) == normalize_master_text(record["name"]):
            return {
                "industryId": record["id"],
                "name": record["name"],
                "score": 1.0,
                "matched": True,
                "review_required": False,
            }

    semantic_ids = []
    if any(t in evidence for t in ["construction", "civil", "cement", "metal", "building"]):
        semantic_ids.append(6)
    if any(t in evidence for t in ["real estate", "property"]):
        semantic_ids.insert(0, 16)
    if any(t in evidence for t in ["recruitment", "talent acquisition"]):
        semantic_ids.insert(0, 17)
    if any(t in evidence for t in ["oil and gas", "oil gas", "power", "infrastructure", "energy"]):
        semantic_ids.insert(0, 13)
    if any(t in evidence for t in ["architect", "interior design"]):
        semantic_ids.insert(0, 23)
    if any(t in evidence for t in ["hvac", "ventilation", "air conditioning"]):
        semantic_ids.insert(0, 39)

    if semantic_ids:
        wanted = semantic_ids[0]
        record = next((x for x in INDUSTRY_MASTER if x["id"] == wanted), None)
        if record:
            return {
                "industryId": record["id"],
                "name": record["name"],
                "score": 0.94,
                "matched": True,
                "review_required": False,
            }

    names = [item["name"] for item in INDUSTRY_MASTER if item.get("active", 1) == 1 and item["id"] != 58]
    name, score = best_named_match(raw, names, minimum_score=0.56)

    if name:
        record = next(item for item in INDUSTRY_MASTER if item["name"] == name)
        return {
            "industryId": record["id"],
            "name": record["name"],
            "score": round(score, 3),
            "matched": True,
            "review_required": score < 0.80,
        }

    return {
        "industryId": None,
        "name": raw,
        "score": round(score, 3),
        "matched": False,
        "review_required": True,
    }


def _keyword_exact_or_fuzzy(skill):
    skill_n = normalize_master_text(skill)
    if not skill_n:
        return "", 0.0

    # Exact normalized master option.
    for master in KEYWORD_MASTER:
        if normalize_master_text(master) == skill_n:
            return master, 1.0

    # Common aliases / normalizations.
    aliases = [
        (["autocad", "auto cad"], "Autocad"),
        (["quantity takeoff", "quantity take off"], "Quantity take-off"),
        (["quantity surveying", "quantity surveyor", "qs"], "QS"),
        (["quantity estimation"], "Quantity Estimation"),
        (["cost estimation", "cost estimating"], "cost estimation"),
        (["primavera p6", "primavera"], "Primavera"),
        (["microsoft project", "ms project", "msp"], "MSP"),
        (["quality assurance"], "QA"),
        (["quality control"], "Quality Control"),
        (["ehs", "hse"], "HSE"),
        (["revit mep"], " REVIT MEP"),
    ]
    for variants, exact in aliases:
        if any(normalize_master_text(v) == skill_n or normalize_master_text(v) in skill_n for v in variants):
            if exact in KEYWORD_MASTER:
                return exact, 0.95

    return best_named_match(skill, KEYWORD_MASTER, minimum_score=0.69)


def _skill_phrase_count(source_normalized, phrase):
    """Count a normalized phrase deterministically using token boundaries."""
    phrase_n = normalize_master_text(phrase)
    if not phrase_n or len(phrase_n) < 3:
        return 0
    pattern = r"(?<!\w)" + re.escape(phrase_n) + r"(?!\w)"
    return len(re.findall(pattern, source_normalized))


def map_keywords(data, source_text=""):
    """Return up to four deterministic exact Skill Groomers keyword names.

    The final four are chosen primarily from the PDF text itself, not from
    Gemini's changing top-skill shortlist. Gemini skills are only a fallback.
    This makes repeated runs of the same CV produce the same SG key skills.
    """
    source = source_text or _combined_professional_evidence(data)
    source_n = normalize_master_text(source)
    profile_n = normalize_master_text(" ".join([
        str(data.get("core_role", "") or ""),
        str(data.get("functional_area", "") or ""),
        str(data.get("role", "") or ""),
        source,
    ]))

    # Stable master-data order is the final tie-breaker.
    master_index = {name: i for i, name in enumerate(KEYWORD_MASTER)}
    scores = {}
    evidence = {}

    def add(name, score, why):
        if name not in master_index:
            return
        score = float(score)
        if score > scores.get(name, float("-inf")):
            scores[name] = score
            evidence[name] = why

    # 1) Direct phrases present in the actual PDF text.
    # Longer phrases receive a small preference over generic one-word terms.
    for name in KEYWORD_MASTER:
        name_n = normalize_master_text(name)
        if len(name_n) < 3:
            continue
        count = _skill_phrase_count(source_n, name)
        if count:
            token_bonus = min(len(name_n.split()), 4) * 5
            frequency_bonus = min(count, 5) * 2
            add(name, 100 + token_bonus + frequency_bonus,
                f"PDF phrase match ({count} occurrence{'s' if count != 1 else ''})")

    # 2) Deterministic aliases for common CV wording that differs from SG labels.
    # An alias is only used when its wording actually occurs in the PDF text.
    alias_rules = [
        (["qa qc", "qa/qc", "quality assurance quality control"], "QAQC"),
        (["quality assurance"], "QA"),
        (["quality control"], "Quality Control"),
        (["quality audit", "quality auditing"], "Quality Audit"),
        (["auto cad", "autocad"], "Autocad"),
        (["quantity surveying", "quantity surveyor"], "QS"),
        (["quantity take off", "quantity takeoff"], "Quantity take-off"),
        (["quantity estimation"], "Quantity Estimation"),
        (["cost estimation", "cost estimating"], "cost estimation"),
        (["primavera p6", "primavera"], "Primavera"),
        (["microsoft project", "ms project"], "MSP"),
        (["health safety environment", "health safety and environment", "hse"], "HSE"),
        (["environment health safety", "ehs"], "HSE"),
        (["talent acquisition"], "Talent Acquisition"),
        (["project planning"], "project planning"),
        (["project management"], "project management"),
        (["billing engineering", "billing"], "Billing"),
        (["ra bill", "ra bills", "running account bill", "running account bills"], "RA Bill"),
    ]
    for variants, master_name in alias_rules:
        for variant in variants:
            variant_n = normalize_master_text(variant)
            if variant_n and _skill_phrase_count(source_n, variant_n):
                add(master_name, 124 + min(len(variant_n.split()), 4) * 2,
                    f"PDF alias match: {variant}")
                break

    # 3) Role-aware boosts. These NEVER create a skill by themselves; they only
    # rank skills for which the PDF already supplied evidence above.
    role_priority_groups = []
    if any(t in profile_n for t in ["qa qc", "qaqc", "quality assurance", "quality control", "quality engineer"]):
        role_priority_groups.append([
            "QAQC", "Quality Control", "QA", "QC", "QMS", "Quality Audit", "Audit"
        ])
    if any(t in profile_n for t in ["planning engineer", "project planning", "project control", "planning manager"]):
        role_priority_groups.append([
            "Planning", "project planning", "Primavera", "MSP", "scheduling", "project coordination", "project management"
        ])
    if any(t in profile_n for t in ["billing engineer", "billing contract", "billing & contract", "billing and contract"]):
        role_priority_groups.append([
            "Billing", "RA Bill", "QS", "Quantity Estimation", "Quantity take-off", "cost estimation", "Estimation & Planning"
        ])
    elif any(t in profile_n for t in ["quantity survey", "estimator", "estimation"]):
        role_priority_groups.append([
            "QS", "Quantity take-off", "Quantity Estimation", "cost estimation", "Estimation & Planning", "Billing", "RA Bill"
        ])
    if any(t in profile_n for t in ["fitout", "fit out", "civil engineer", "site engineer"]):
        role_priority_groups.append([
            "Civil", "Billing", "Quality Control", "QA", "Autocad", "project planning", "cost estimation"
        ])
    if any(t in profile_n for t in ["safety", "hse", "ehs"]):
        role_priority_groups.append([
            "HSE", "safety", "safety manager", "HIRA", "Audit", "OH&S ", "HSE plan"
        ])
    if any(t in profile_n for t in ["recruitment", "talent acquisition", "recruiter"]):
        role_priority_groups.append([
            "Talent Acquisition", "Recruitment", "sourcing", "client coordination"
        ])

    for group in role_priority_groups:
        for rank, name in enumerate(group):
            if name in scores:
                scores[name] += max(0, 30 - (rank * 4))

    # 4) For the normal app flow, source_text is always available, so the final
    # Skill Groomers selection does NOT depend on Gemini's varying skill shortlist.
    # Keep the old factual-skill fallback only for legacy/manual calls where no
    # PDF source text was supplied.
    if not source_text:
        raw_skills = clean_key_skills(data.get("key_skills", []))
        for skill in raw_skills:
            candidate, similarity = _keyword_exact_or_fuzzy(skill)
            if candidate:
                add(candidate, 70 + (similarity * 10), f"Factual skill fallback: {skill}")

    ranked = sorted(
        scores,
        key=lambda name: (-scores[name], master_index.get(name, 10**9), normalize_master_text(name)),
    )

    # Avoid spending multiple SG slots on near-duplicate skill families.
    # This affects only the max-4 Skill Groomers mapping; the factual CV skill
    # list remains complete and is displayed separately to HR.
    skill_families = [
        {"qaqc", "qa", "qc", "quality control"},
    ]

    selected = []
    used_families = set()
    for name in ranked:
        name_n = normalize_master_text(name)
        family_index = next(
            (i for i, family in enumerate(skill_families) if name_n in family),
            None,
        )
        if family_index is not None and family_index in used_families:
            continue
        selected.append(name)
        if family_index is not None:
            used_families.add(family_index)
        if len(selected) == 4:
            break

    return selected


LOCATION_ALIASES = {
    "bangalore": "Bangalore / Bengaluru",
    "bengaluru": "Bangalore / Bengaluru",
    "delhi": "Delhi / NCR",
    "ncr": "Delhi / NCR",
    "hyderabad": "Hyderabad / Secunderabad",
    "secunderabad": "Hyderabad / Secunderabad",
    "odisha": "Odisha / Orissa",
    "orissa": "Odisha / Orissa",
    "uttarakhand": "Uttaranchal",
    "tamil nadu": "Tamilnadu",
    "usa": "US",
    "united states": "US",
    "united arab emirates": "UAE",
}


def map_location(value, expected_category=None):
    """
    HR rule: Skill Groomers location must be an EXACT master-data match.
    If the CV wording does not exactly match an active SG location after
    harmless case/spacing normalization, leave the SG ID blank and require
    review. No city aliases and no fuzzy substitution are used.
    """
    value_n = normalize_master_text(value)
    if not value_n:
        return {
            "id": None,
            "name": "",
            "category": expected_category or "",
            "matched": False,
            "review_required": False,
            "reason": "No factual location in CV",
        }

    candidates = [
        item for item in LOCATION_MASTER
        if item.get("active", 1) == 1
        and (expected_category is None or item.get("category") == expected_category)
    ]

    for item in candidates:
        if normalize_master_text(item["name"]) == value_n:
            return {
                "id": item["id"],
                "name": item["name"],
                "category": item["category"],
                "matched": True,
                "review_required": False,
                "reason": "Exact active Skill Groomers location match",
            }

    return {
        "id": None,
        "name": str(value or "").strip(),
        "category": expected_category or "",
        "matched": False,
        "review_required": True,
        "reason": "No exact Skill Groomers location match; HR review required",
    }


EDUCATION_ALIASES = {
    "be": "be btech",
    "b e": "be btech",
    "btech": "be btech",
    "b tech": "be btech",
    "bachelor engineering": "be btech",
    "bsc": "bachelor science",
    "b sc": "bachelor science",
    "msc": "masters science",
    "m sc": "masters science",
    "mtech": "masters technology",
    "m tech": "masters technology",
    "mba": "masters business administration",
    "bba": "bachelor business administration",
    "bcom": "bachelor commerce",
    "b com": "bachelor commerce",
    "mcom": "masters commerce",
    "m com": "masters commerce",
}


def _education_component_score(raw, master):
    raw_n = normalize_master_text(raw)
    master_n = normalize_master_text(master)
    if not raw_n or not master_n:
        return 0.0

    score = text_similarity(raw_n, master_n)

    for alias, canonical in EDUCATION_ALIASES.items():
        if alias in raw_n and canonical in master_n:
            score = max(score, 0.94)

    return score


def _preferred_education_course(degree, course):
    """Detect an explicit course family before considering specialization."""
    evidence = normalize_master_text(f"{degree} {course}")

    # Order matters: specific higher qualifications before generic tokens.
    rules = [
        (["m tech", "mtech", "master technology"], "M.Tech"),
        (["m sc", "msc", "master science"], "M.Sc"),
        (["mba", "pgdm", "master business administration"], "MBA / PGDM"),
        (["mca"], "MCA"),
        (["m com", "mcom"], "M.Com"),
        (["b e", "be civil", "be mechanical", "be electrical", "bachelor engineering", "bachelor of engineering", "b tech", "btech"], "BE / B.Tech"),
        (["b sc", "bsc", "bachelor science", "bachelor of science"], "Bachelor of Science (B.Sc)"),
        (["b com", "bcom"], "B.Com"),
        (["bba"], "BBA"),
        (["bca"], "BCA"),
        (["b arch", "barch", "bachelor architecture"], "Bachor of Arcitect (B. Arch.)"),
        (["diploma"], "Diploma"),
        (["hsc"], "HSC"),
        (["ssc", "matriculation"], "SSC"),
        (["iti"], "ITI"),
    ]
    for triggers, target in rules:
        if any(trigger in evidence for trigger in triggers):
            return target
    return ""


def map_education(data):
    """
    Map the highest completed qualification to the SG education hierarchy.

    V7 rule: identify the degree/course FIRST, then select specialization.
    This prevents a specialization such as Civil from incorrectly turning a
    BE Civil profile into Diploma Civil.
    """
    qual = data.get("highest_qualification", {}) or {}
    degree = str(qual.get("degree", "") or "")
    course = str(qual.get("course", "") or "")
    specialization = str(qual.get("specialization", "") or "")
    completion_year = qual.get("year", "")

    preferred_course = _preferred_education_course(degree, course)
    candidates = list(EDUCATION_MASTER)

    if preferred_course:
        preferred_n = normalize_master_text(preferred_course)
        candidates = [
            r for r in candidates
            if normalize_master_text(r.get("course", "")) == preferred_n
        ]

    if not candidates:
        candidates = list(EDUCATION_MASTER)

    # Prefer an exact/contained specialization inside the already identified
    # course family. This makes BE/B.Tech + Civil deterministically choose
    # Graduation -> BE/B.Tech -> Civil (ID 31), never Diploma -> Civil (ID 116).
    if preferred_course and specialization:
        raw_spec_n = normalize_master_text(specialization)
        exact_spec_candidates = []
        for record in candidates:
            master_spec_n = normalize_master_text(record.get("specialization", ""))
            if master_spec_n and (
                master_spec_n == raw_spec_n
                or master_spec_n in raw_spec_n
                or raw_spec_n in master_spec_n
            ):
                exact_spec_candidates.append(record)

        if len(exact_spec_candidates) == 1:
            exact = exact_spec_candidates[0]
            return {
                "educationId": exact["education_id"],
                "qualification": exact["qualification"],
                "course": exact["course"],
                "specialization": exact["specialization"],
                "completionYear": completion_year,
                "score": 0.99,
                "matched": True,
                "review_required": False,
                "reason": "Exact course-family and specialization match",
            }

    best = None
    best_score = 0.0
    for record in candidates:
        course_score = max(
            _education_component_score(degree, record["course"]),
            _education_component_score(course, record["course"]),
        )
        spec_score = _education_component_score(specialization, record["specialization"])

        # Once course family is explicit, specialization is the deciding factor.
        if preferred_course:
            score = (0.35 * course_score) + (0.65 * spec_score)
            if normalize_master_text(record["course"]) == normalize_master_text(preferred_course):
                score = min(1.0, score + 0.25)
        else:
            qual_score = _education_component_score(degree + " " + course, record["qualification"])
            score = (0.52 * course_score) + (0.38 * spec_score) + (0.10 * qual_score)

        spec_n = normalize_master_text(record["specialization"])
        raw_spec_n = normalize_master_text(specialization)
        if spec_n and raw_spec_n and (spec_n == raw_spec_n or spec_n in raw_spec_n):
            score = min(1.0, score + 0.15)

        if score > best_score:
            best, best_score = record, score

    if not best or best_score < 0.54:
        return {
            "educationId": None,
            "qualification": degree,
            "course": course,
            "specialization": specialization,
            "completionYear": completion_year,
            "score": round(best_score, 3),
            "matched": False,
            "review_required": True,
        }

    return {
        "educationId": best["education_id"],
        "qualification": best["qualification"],
        "course": best["course"],
        "specialization": best["specialization"],
        "completionYear": completion_year,
        "score": round(best_score, 3),
        "matched": True,
        "review_required": best_score < 0.80,
    }


# Deterministic city -> state relationships for active Skill Groomers metro-city options.
# These are used ONLY to fill a missing SG state. They do not overwrite the factual
# CV state stored in Excel/raw extraction. Ambiguous SG city labels such as Delhi / NCR
# are deliberately omitted.
SG_CITY_TO_STATE = {
    "Mumbai": "Maharashtra",
    "Pune": "Maharashtra",
    "Nagpur": "Maharashtra",
    "Ahmedabad": "Gujarat",
    "Rajkot": "Gujarat",
    "Bangalore / Bengaluru": "Karnataka",
    "Chennai": "Tamilnadu",
    "Hyderabad / Secunderabad": "Telangana",
    "Kolkata": "West Bengal",
    "Patna": "Bihar",
    "Bhopal": "Madhya Pradesh",
    "Indore": "Madhya Pradesh",
    "Kanpur": "Uttar Pradesh",
    "Lucknow": "Uttar Pradesh",
    "Madgaon": "Goa",
    "Ponda": "Goa",
    "Shimla": "Himachal Pradesh",
    "Chandigarh": "Union Territories",
}


def _infer_missing_state_from_city(city_mapping, state_mapping, factual_state=""):
    """Infer an SG state only when the CV state is missing and city is unambiguous.

    If the CV explicitly supplies a valid state that conflicts with the known city/state
    relationship, keep the explicit state but flag it for HR review.
    """
    if not isinstance(city_mapping, dict) or not city_mapping.get("matched"):
        return state_mapping

    inferred_state_name = SG_CITY_TO_STATE.get(city_mapping.get("name", ""))
    if not inferred_state_name:
        return state_mapping

    inferred = map_location(inferred_state_name, "States")
    if not inferred.get("matched"):
        return state_mapping

    factual_state = str(factual_state or "").strip()

    # Missing state in the CV: deterministic geographic inference is safe here.
    if not factual_state:
        inferred["reason"] = (
            f"State inferred deterministically from exact city match: "
            f"{city_mapping.get('name')} → {inferred_state_name}"
        )
        inferred["inferred_from_city"] = True
        inferred["review_required"] = False
        return inferred

    # If CV supplied a state and both city/state matched SG exactly, detect conflicts.
    if state_mapping.get("matched"):
        if state_mapping.get("id") != inferred.get("id"):
            result = dict(state_mapping)
            result["review_required"] = True
            result["location_conflict"] = True
            result["reason"] = (
                f"Location conflict: city {city_mapping.get('name')} maps to "
                f"{inferred_state_name}, but CV states {factual_state}. HR review required."
            )
            return result

    return state_mapping


def build_skill_groomers_mapping(data, source_text=""):
    """Build one exact, reviewable Skill Groomers website mapping."""
    current_city = map_location(data.get("current_city", ""), "Top Metropolitan Cities")
    current_state = map_location(data.get("current_state", ""), "States")
    native_city = map_location(data.get("native_city", ""), "Top Metropolitan Cities")
    native_state = map_location(data.get("native_state", ""), "States")

    current_state = _infer_missing_state_from_city(
        current_city, current_state, data.get("current_state", "")
    )
    native_state = _infer_missing_state_from_city(
        native_city, native_state, data.get("native_state", "")
    )

    mapping = {
        "core_role": map_core_role(data, source_text=source_text),
        "functional_area": map_functional_area(data, source_text=source_text),
        "industry": map_industry(data),
        "key_skills": map_keywords(data, source_text=source_text),
        "education": map_education(data),
        "current_city": current_city,
        "current_state": current_state,
        "native_city": native_city,
        "native_state": native_state,
    }

    review_sections = {
        "Core Role": mapping["core_role"],
        "Functional Area": mapping["functional_area"],
        "Industry": mapping["industry"],
        "Education": mapping["education"],
        "Current City": mapping["current_city"],
        "Current State": mapping["current_state"],
        "Native City": mapping["native_city"],
        "Native State": mapping["native_state"],
    }

    mapping["review_reasons"] = [
        {
            "section": label,
            "reason": section.get("reason", "Mapping requires HR review"),
        }
        for label, section in review_sections.items()
        if isinstance(section, dict) and section.get("review_required", False)
    ]
    mapping["review_required"] = bool(mapping["review_reasons"])

    return mapping


def apply_mapping_to_excel_data(data, mapping):
    """
    V7 ACCURACY-FIRST RULE:
    Excel keeps the factual/human-readable values extracted from the CV.
    Skill Groomers mapping is stored separately and must never overwrite
    education, employment, personal details, or other CV facts.
    """
    excel_data = dict(data)
    excel_data["skill_groomers_ids"] = {
        "city": mapping["current_city"]["id"],
        "state": mapping["current_state"]["id"],
        "nativeCity": mapping["native_city"]["id"],
        "nativeState": mapping["native_state"]["id"],
        "functionalAreaId": mapping["functional_area"]["functionalAreaId"],
        "industryId": mapping["industry"]["industryId"],
        "educationId": mapping["education"]["educationId"],
    }
    return excel_data


def _website_datetime(value):
    value = str(value or "").strip()
    if not value:
        return None
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return f"{value} 00:00:00"
    return value


def build_skill_groomers_payload_preview(final_data, mapping):
    """
    Build the candidate structure intended for Skill Groomers.
    Resume stays as a placeholder until the S3 upload step is added.
    """
    jobs = sort_jobs(final_data.get("work_experience", []))
    current_job, previous_job = get_current_and_previous_jobs(jobs)

    total_months = total_experience_months(jobs)
    total_years, remaining_months = split_months(total_months)

    job_changes = compute_job_changes(jobs)
    avg_months = average_tenure_months(total_months, job_changes)
    avg_years, avg_remaining_months = split_months(avg_months)

    education = mapping["education"]
    ids = final_data.get("skill_groomers_ids", {})

    key_skills = final_data.get("key_skills", []) or []
    salary = final_data.get("annual_salary", "")
    try:
        salary_value = float(str(salary).replace(",", "").strip()) if str(salary).strip() else None
    except ValueError:
        salary_value = None

    payload = {
        "fullName": final_data.get("candidate_name", ""),
        "emailId": final_data.get("email", ""),
        "mobileNumber": final_data.get("phone", ""),
        "alternateNumber": final_data.get("alternate_phone") or None,
        "dateOfBirth": _website_datetime(final_data.get("date_of_birth", "")),
        "city": mapping["current_city"].get("id"),
        "state": mapping["current_state"].get("id"),
        "nativeCity": mapping["native_city"].get("id"),
        "nativeState": mapping["native_state"].get("id"),
        "coreRole": mapping["core_role"].get("name", "") if mapping["core_role"].get("matched") else final_data.get("core_role", ""),
        "gender": final_data.get("gender", "") or None,
        "rating": None,
        "source": None,
        "communication": None,
        "employment": {
            "keySkill": ",".join(mapping.get("key_skills", []) or key_skills),
            "industryId": mapping["industry"].get("industryId"),
            "functionalAreaId": mapping["functional_area"].get("functionalAreaId"),
            "currentDesignation": current_job.get("designation", "") if current_job else "",
            "currentEmployer": current_job.get("company", "") if current_job else "",
            "startDate": _website_datetime(current_job.get("start_date", "")) if current_job else None,
            "endDate": (
                datetime.today().strftime("%Y-%m-%d 00:00:00")
                if current_job and str(current_job.get("end_date", "")).strip().lower()
                in ["present", "current", "ongoing", "till date", "now"]
                else _website_datetime(current_job.get("end_date", "")) if current_job else None
            ),
            "previousDesignation": previous_job.get("designation", "") if previous_job else "",
            "previousEmployer": previous_job.get("company", "") if previous_job else "",
            "totalExperienceInYears": total_years,
            "totalExperienceInMonths": remaining_months,
            "averageTenureInYears": avg_years,
            "averageTenureInMonths": avg_remaining_months,
            "totalExperienceAsOfDate": datetime.today().strftime("%Y-%m-%d 00:00:00"),
            "totalNumberOfJobs": str(number_of_employers(jobs)),
            "annualSalary": salary_value,
        },
        "educations": (
            [{
                "educationId": education["educationId"],
                "completionYear": str(education.get("completionYear", "") or ""),
            }]
            if education.get("matched") and education.get("educationId") is not None
            else []
        ),
        "resume": "<uploaded resume URL will be inserted here>",
    }

    return payload


# ---------- Populate the Excel template ----------

def populate_excel(data, sheet):
    # Personal Details
    sheet["C4"] = data.get("candidate_name", "")
    sheet["C5"] = data.get("email", "")
    sheet["C6"] = data.get("phone", "")
    sheet["C7"] = data.get("alternate_phone", "")
# Date of Birth
# Only use the DOB if it is explicitly mentioned in the CV
    dob = data.get("date_of_birth", "")

    sheet["C8"] = dob or ""
    sheet["C9"] = data.get("gender", "")
    sheet["C10"] = data.get("current_city", "")
    sheet["C11"] = data.get("current_state", "")
    sheet["C12"] = data.get("native_city", "")
    sheet["C13"] = data.get("native_state", "")

    # Core Role
    sheet["C14"] = data.get("core_role", "")

    # Employment / Skills
    raw_skills = data.get("key_skills", [])
    # Clean certifications from skills as a safety net
    cleaned_skills = clean_key_skills(raw_skills)
    sheet["C15"] = ", ".join(cleaned_skills)
    # Human-readable CV classification for Excel; SG mapping stays separate.
    sheet["C16"] = data.get("functional_area", "")
    sheet["C17"] = data.get("role", "")
    sheet["C18"] = data.get("industry", "")

    # Work experience
    jobs = sort_jobs(data.get("work_experience", []))
    current_job, prev_job = get_current_and_previous_jobs(jobs)

    # Current Employment
    if current_job:
        sheet["C20"] = current_job.get("designation", "")
        sheet["C21"] = current_job.get("company", "")
    else:
        sheet["C20"] = ""
        sheet["C21"] = ""

    # Previous Employment
    if prev_job:
        sheet["C23"] = prev_job.get("designation", "")
        sheet["C24"] = prev_job.get("company", "")
    else:
        sheet["C23"] = ""
        sheet["C24"] = ""

    # Total Experience - calculated directly in months
    total_months = total_experience_months(jobs)
    sheet["C26"] = format_months(total_months)
    sheet["C27"] = datetime.today().strftime("%Y-%m-%d")

    num_emp = number_of_employers(jobs)
    sheet["C28"] = num_emp

    # Salary is only present when explicitly stated in the CV.
    sheet["C29"] = data.get("annual_salary", "")

    # Average Tenure = total experience / job changes (HR-requested formula)
    job_changes = compute_job_changes(jobs)
    avg_months = average_tenure_months(total_months, job_changes)
    sheet["C31"] = format_months(avg_months)

    # Education
    qual = data.get("highest_qualification", {})
    sheet["C33"] = qual.get("degree", "")          # Qualification (full name)
    sheet["C34"] = qual.get("year", "")            # Year of Passing
    sheet["C35"] = qual.get("course", "")          # Course (acronym)
    sheet["C36"] = qual.get("specialization", "")  # Specialization


# ---------- HR Review UI helpers ----------

def _safe_index(options, value):
    """Return a safe selectbox index for value."""
    try:
        return options.index(value)
    except ValueError:
        return 0


def _active_location_options(category=None):
    records = [
        item for item in LOCATION_MASTER
        if item.get("active", 1) == 1
        and (category is None or item.get("category") == category)
    ]
    return ["— Not selected —"] + sorted(
        {str(item.get("name", "")).strip() for item in records if str(item.get("name", "")).strip()},
        key=str.lower
    )


def _location_id_from_name(name, category=None):
    if not name or name == "— Not selected —":
        return None
    for item in LOCATION_MASTER:
        if (
            item.get("active", 1) == 1
            and str(item.get("name", "")).strip() == name
            and (category is None or item.get("category") == category)
        ):
            return item.get("id")
    return None


def _core_role_options():
    return ["— Not selected —"] + sorted(
        {str(name).strip() for name in CORE_ROLE_MASTER if str(name).strip()},
        key=str.lower
    )


def _industry_options():
    return ["— Not selected —"] + sorted(
        {
            str(item.get("name", "")).strip()
            for item in INDUSTRY_MASTER
            if item.get("active", 1) == 1 and str(item.get("name", "")).strip()
        },
        key=str.lower
    )


def _industry_id_from_name(name):
    if not name or name == "— Not selected —":
        return None
    for item in INDUSTRY_MASTER:
        if item.get("active", 1) == 1 and str(item.get("name", "")).strip() == name:
            return item.get("id")
    return None


def _functional_role_options():
    options = []
    for item in FUNCTIONAL_AREA_MASTER:
        role = str(item.get("role", "")).strip()
        if not role:
            continue
        label = (
            f"{item.get('functional_area', '')} → "
            f"{item.get('sub_functional_area', '')} → {role}"
        )
        options.append((label, item))
    options.sort(key=lambda pair: pair[0].lower())
    return options


def _education_options():
    options = []
    for item in EDUCATION_MASTER:
        label = (
            f"{item.get('qualification', '')} → "
            f"{item.get('course', '')} → "
            f"{item.get('specialization', '')}"
        )
        options.append((label, item))
    options.sort(key=lambda pair: pair[0].lower())
    return options



def _display_value(label, value):
    """Premium compact read-only factual field for HR review."""
    shown = value if value not in (None, "") else "—"
    st.markdown(
        f'<div class="field-label">{label}</div>'
        f'<div class="field-value">{shown}</div>',
        unsafe_allow_html=True,
    )



SG_ADD_CANDIDATE_URL = "https://skillgroomers.projects-digitalgem.com/main/candidate/new"


def _sg_date_for_form(value):
    """Convert YYYY-MM / YYYY-MM-DD website values into DD/MM/YYYY for SG form."""
    value = str(value or "").strip()
    if not value:
        return ""
    if value.lower() in {"present", "current", "ongoing", "till date", "now"}:
        return datetime.today().strftime("%d/%m/%Y")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y-%m"):
        try:
            parsed = datetime.strptime(value, fmt)
            # Month-only CV dates are intentionally represented as first of month
            # because the SG date control requires a day.
            return parsed.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return value


def build_skill_groomers_form_transfer(
    data,
    reviewed_payload,
    selected_core,
    selected_func_item,
    selected_industry,
    selected_keywords,
    selected_city,
    selected_state,
    selected_native_city,
    selected_native_state,
    selected_edu_item,
):
    """
    Build browser-extension transfer data.
    Only HR-reviewed/factual values are included; no authentication data is exported.
    """
    employment = reviewed_payload.get("employment", {}) or {}
    jobs = sort_jobs(data.get("work_experience", []))
    current_job, _ = get_current_and_previous_jobs(jobs)

    functional_area = ""
    functional_role = ""
    if selected_func_item:
        functional_area = str(selected_func_item.get("functional_area", "") or "")
        functional_role = str(selected_func_item.get("role", "") or "")

    qualification = ""
    course = ""
    specialization = ""
    if selected_edu_item:
        qualification = str(selected_edu_item.get("qualification", "") or "")
        course = str(selected_edu_item.get("course", "") or "")
        specialization = str(selected_edu_item.get("specialization", "") or "")

    current_end = current_job.get("end_date", "") if current_job else ""

    return {
        "version": 1,
        "candidate": {
            "coreRole": selected_core if selected_core != "— Not selected —" else "",
            "fullName": str(reviewed_payload.get("fullName", "") or ""),
            "emailId": str(reviewed_payload.get("emailId", "") or ""),
            "mobileNumber": str(reviewed_payload.get("mobileNumber", "") or ""),
            "alternateNumber": str(reviewed_payload.get("alternateNumber", "") or ""),
            "dateOfBirth": _sg_date_for_form(data.get("date_of_birth", "")),
            "gender": str(data.get("gender", "") or ""),
            "currentCity": "" if selected_city == "— Not selected —" else selected_city,
            "currentState": "" if selected_state == "— Not selected —" else selected_state,
            "nativeCity": "" if selected_native_city == "— Not selected —" else selected_native_city,
            "nativeState": "" if selected_native_state == "— Not selected —" else selected_native_state,
            "keySkills": list(selected_keywords or [])[:4],
            "functionalArea": functional_area,
            "role": functional_role,
            "industry": "" if selected_industry == "— Not selected —" else selected_industry,
            "currentDesignation": str(employment.get("currentDesignation", "") or ""),
            "currentEmployer": str(employment.get("currentEmployer", "") or ""),
            "startDate": _sg_date_for_form(employment.get("startDate", "")),
            "endDate": _sg_date_for_form(current_end),
            "previousDesignation": str(employment.get("previousDesignation", "") or ""),
            "previousEmployer": str(employment.get("previousEmployer", "") or ""),
            "totalExperienceInYears": employment.get("totalExperienceInYears"),
            "totalExperienceInMonths": employment.get("totalExperienceInMonths"),
            "totalExperienceAsOfDate": _sg_date_for_form(
                employment.get("totalExperienceAsOfDate", "")
            ),
            "totalNumberOfJobs": str(employment.get("totalNumberOfJobs", "") or ""),
            "annualSalaryLakh": employment.get("annualSalary"),
            "averageTenureInYears": employment.get("averageTenureInYears"),
            "averageTenureInMonths": employment.get("averageTenureInMonths"),
            "qualification": qualification,
            "course": course,
            "specialization": specialization,
            "completionYear": str(
                ((reviewed_payload.get("educations") or [{}])[0]).get("completionYear", "")
                if reviewed_payload.get("educations")
                else ""
            ),
        }
    }


def skill_groomers_transfer_url(transfer):
    """
    Put the transfer JSON in the URL fragment.
    Fragments are not sent to the Skill Groomers server; the installed extension reads it locally.
    """
    raw = json.dumps(transfer, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    token = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    return f"{SG_ADD_CANDIDATE_URL}#cvfill={urllib.parse.quote(token)}"


def _render_review_card(filename, file_info):
    """
    Fast HR review workflow.
    Goal: most clean CVs can be reviewed in ~30–60 seconds.
    Confirmed CV facts stay read-only; Skill Groomers classifications are prefilled.
    HR only opens the edit section when something needs changing.
    """
    data = file_info["factual_data"]
    mapping = file_info["mapping"]
    payload = file_info["payload_preview"]
    candidate = file_info["candidate_name"]

    jobs = sort_jobs(data.get("work_experience", []))
    current_job, previous_job = get_current_and_previous_jobs(jobs)
    total_months = total_experience_months(jobs)
    job_changes = compute_job_changes(jobs)
    avg_months = average_tenure_months(total_months, job_changes)
    factual_edu = data.get("highest_qualification", {}) or {}

    st.markdown(
        f'<div class="candidate-chip">QUICK REVIEW</div>'
        f'<h2 style="margin:0 0 4px 0">{candidate}</h2>'
        f'<div class="review-subtitle">'
        f'Check the summary. Open Edit classifications only if something needs changing.'
        f'</div>',
        unsafe_allow_html=True,
    )

    if data.get("_employment_extraction_warning"):
        st.error(
            "Employment section detected, but no job history was extracted. "
            "Do not approve this candidate until the employment extraction is reviewed."
        )

    # ---------------------------------------------------------------
    # 1. FAST FACT CHECK
    # ---------------------------------------------------------------
    st.markdown("### 1. Candidate Check")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        _metric_card("Experience", format_months(total_months))
    with m2:
        _metric_card("Avg. Tenure", format_months(avg_months))
    with m3:
        _metric_card("Employers", number_of_employers(jobs))
    with m4:
        _metric_card("Job Changes", job_changes)

    _info_grid(
        [
            ("Candidate", data.get("candidate_name", "")),
            ("Email", data.get("email", "")),
            ("Mobile", data.get("phone", "")),
            (
                "Current Employment",
                " — ".join(
                    x for x in [
                        current_job.get("designation", "") if current_job else "",
                        current_job.get("company", "") if current_job else "",
                    ] if x
                ),
            ),
            (
                "Previous Employment",
                " — ".join(
                    x for x in [
                        previous_job.get("designation", "") if previous_job else "",
                        previous_job.get("company", "") if previous_job else "",
                    ] if x
                ),
            ),
            (
                "Education",
                " | ".join(
                    x for x in [
                        factual_edu.get("degree", ""),
                        factual_edu.get("specialization", ""),
                        str(factual_edu.get("year", "") or ""),
                    ] if x
                ),
            ),
        ],
        columns=3,
    )

    # ---------------------------------------------------------------
    # 2. SUGGESTED SG MAPPING SUMMARY
    # ---------------------------------------------------------------
    st.markdown("### 2. Suggested Skill Groomers Mapping")

    sg_core = mapping.get("core_role", {}) or {}
    sg_func = mapping.get("functional_area", {}) or {}
    sg_industry = mapping.get("industry", {}) or {}
    sg_edu = mapping.get("education", {}) or {}
    sg_city = mapping.get("current_city", {}) or {}
    sg_state = mapping.get("current_state", {}) or {}
    sg_native_city = mapping.get("native_city", {}) or {}
    sg_native_state = mapping.get("native_state", {}) or {}

    functional_path = " → ".join(
        x for x in [
            sg_func.get("functional_area", ""),
            sg_func.get("sub_functional_area", ""),
            sg_func.get("role", ""),
        ] if x
    )
    education_path = " → ".join(
        x for x in [
            sg_edu.get("qualification", ""),
            sg_edu.get("course", ""),
            sg_edu.get("specialization", ""),
        ] if x
    )
    current_location = ", ".join(
        x for x in [sg_city.get("name", ""), sg_state.get("name", "")] if x
    )
    native_location = ", ".join(
        x for x in [sg_native_city.get("name", ""), sg_native_state.get("name", "")] if x
    )
    mapped_keywords = mapping.get("key_skills", []) or []

    _info_grid(
        [
            ("Core Role", sg_core.get("name", "")),
            ("Industry", sg_industry.get("name", "")),
            ("Functional Area", functional_path),
            ("Key Skills", ", ".join(mapped_keywords[:4])),
            ("Current Location", current_location),
            ("Native Location", native_location),
            ("Education", education_path),
        ],
        columns=2,
    )

    review_reasons = mapping.get("review_reasons", []) or []
    if review_reasons:
        st.warning(
            f"{len(review_reasons)} item(s) need attention. "
            "Review the highlighted reason(s) below before approval."
        )
        for item in review_reasons:
            st.write(
                f"⚠️ **{item.get('section', 'Mapping')}** — "
                f"{item.get('reason', 'Review required')}"
            )
    else:
        st.success("No mapping conflicts or unresolved fields detected.")

    # ---------------------------------------------------------------
    # 3. EDIT ONLY WHEN NEEDED
    # ---------------------------------------------------------------
    with st.expander(
        "✏️ Edit Skill Groomers classifications",
        expanded=bool(review_reasons),
    ):
        st.caption(
            "The suggested values are already selected. "
            "Change only the fields that are incorrect or unresolved."
        )

        # Core Role
        _cv_source_note("CV says", data.get("core_role", "") or "Not provided")
        core_options = _core_role_options()
        current_core = sg_core.get("name", "")
        selected_core = st.selectbox(
            "Core Role",
            core_options,
            index=_safe_index(core_options, current_core),
            key=f"sg_core_{filename}",
        )

        # Functional Area
        functional_options = _functional_role_options()
        functional_labels = ["— Not selected —"] + [
            label for label, _ in functional_options
        ]
        current_fid = sg_func.get("functionalAreaId")
        current_func_label = "— Not selected —"
        for label, item in functional_options:
            if item.get("role_id") == current_fid:
                current_func_label = label
                break

        selected_func_label = st.selectbox(
            "Functional Area → Sub Functional Area → Role",
            functional_labels,
            index=_safe_index(functional_labels, current_func_label),
            key=f"sg_func_{filename}",
        )
        selected_func_item = next(
            (item for label, item in functional_options if label == selected_func_label),
            None,
        )

        # Industry
        _cv_source_note("CV says", data.get("industry", "") or "Not provided")
        industry_options = _industry_options()
        current_industry = sg_industry.get("name", "")
        selected_industry = st.selectbox(
            "Industry",
            industry_options,
            index=_safe_index(industry_options, current_industry),
            key=f"sg_industry_{filename}",
        )

        # Key Skills
        # Show every factual skill extracted from the CV immediately above the
        # separate Skill Groomers selector. The factual list is unrestricted;
        # only the SG website selection below is limited to four master keywords.
        factual_skills = clean_key_skills(data.get("key_skills", []) or [])
        st.markdown("**CV Key Skills**")
        if factual_skills:
            st.caption(" • ".join(str(skill) for skill in factual_skills))
        else:
            st.caption("Not provided in CV")

        keyword_options = sorted(
            {
                str(keyword).strip()
                for keyword in KEYWORD_MASTER
                if str(keyword).strip()
            },
            key=str.lower,
        )
        default_keywords = [
            keyword for keyword in mapped_keywords
            if keyword in keyword_options
        ][:4]
        selected_keywords = st.multiselect(
            "Key Skills (max 4)",
            keyword_options,
            default=default_keywords,
            max_selections=4,
            key=f"sg_keywords_{filename}",
        )

        # Location
        st.markdown("#### Location")
        current_cv_location = ", ".join(
            x for x in [data.get("current_city", ""), data.get("current_state", "")]
            if x
        ) or "Not provided"
        native_cv_location = ", ".join(
            x for x in [data.get("native_city", ""), data.get("native_state", "")]
            if x
        ) or "Not provided"
        _cv_source_note(
            "CV location",
            f"Current: {current_cv_location} | Native: {native_cv_location}",
        )

        city_options = _active_location_options("Top Metropolitan Cities")
        state_options = _active_location_options("States")
        l1, l2 = st.columns(2)

        with l1:
            mapped_city = sg_city.get("name", "") if sg_city.get("matched") else "— Not selected —"
            selected_city = st.selectbox(
                "Current City",
                city_options,
                index=_safe_index(city_options, mapped_city),
                key=f"sg_city_{filename}",
            )

        with l2:
            mapped_state = sg_state.get("name", "") if sg_state.get("matched") else "— Not selected —"
            selected_state = st.selectbox(
                "Current State",
                state_options,
                index=_safe_index(state_options, mapped_state),
                key=f"sg_state_{filename}",
            )

        n1, n2 = st.columns(2)
        with n1:
            mapped_native_city = (
                sg_native_city.get("name", "")
                if sg_native_city.get("matched")
                else "— Not selected —"
            )
            selected_native_city = st.selectbox(
                "Native City",
                city_options,
                index=_safe_index(city_options, mapped_native_city),
                key=f"sg_native_city_{filename}",
            )

        with n2:
            mapped_native_state = (
                sg_native_state.get("name", "")
                if sg_native_state.get("matched")
                else "— Not selected —"
            )
            selected_native_state = st.selectbox(
                "Native State",
                state_options,
                index=_safe_index(state_options, mapped_native_state),
                key=f"sg_native_state_{filename}",
            )

        # Education
        st.markdown("#### Education")
        _cv_source_note(
            "CV education",
            " | ".join([
                str(factual_edu.get("degree", "") or "—"),
                str(factual_edu.get("course", "") or "—"),
                str(factual_edu.get("specialization", "") or "—"),
                str(factual_edu.get("year", "") or "—"),
            ]),
        )

        education_options = _education_options()
        education_labels = ["— Not selected —"] + [
            label for label, _ in education_options
        ]
        current_eid = sg_edu.get("educationId")
        current_edu_label = "— Not selected —"
        for label, item in education_options:
            if item.get("education_id") == current_eid:
                current_edu_label = label
                break

        selected_edu_label = st.selectbox(
            "Qualification → Course → Specialization",
            education_labels,
            index=_safe_index(education_labels, current_edu_label),
            key=f"sg_edu_{filename}",
        )
        selected_edu_item = next(
            (item for label, item in education_options if label == selected_edu_label),
            None,
        )

        completion_year = str(factual_edu.get("year", "") or "")
        _display_value("Completion Year (locked from CV)", completion_year)

    # ---------------------------------------------------------------
    # 4. REVIEWED PAYLOAD
    # ---------------------------------------------------------------
    reviewed_payload = dict(payload)
    reviewed_payload["coreRole"] = (
        None if selected_core == "— Not selected —" else selected_core
    )
    reviewed_payload["city"] = _location_id_from_name(
        selected_city, "Top Metropolitan Cities"
    )
    reviewed_payload["state"] = _location_id_from_name(
        selected_state, "States"
    )
    reviewed_payload["nativeCity"] = _location_id_from_name(
        selected_native_city, "Top Metropolitan Cities"
    )
    reviewed_payload["nativeState"] = _location_id_from_name(
        selected_native_state, "States"
    )

    reviewed_payload["employment"] = dict(payload.get("employment", {}))
    reviewed_payload["employment"]["keySkill"] = ",".join(selected_keywords)
    reviewed_payload["employment"]["industryId"] = _industry_id_from_name(
        selected_industry
    )
    reviewed_payload["employment"]["functionalAreaId"] = (
        selected_func_item.get("role_id") if selected_func_item else None
    )
    reviewed_payload["educations"] = (
        [{
            "educationId": selected_edu_item.get("education_id"),
            "completionYear": completion_year,
        }]
        if selected_edu_item else []
    )
    file_info["reviewed_payload"] = reviewed_payload

    # ---------------------------------------------------------------
    # 5. ONE-STEP APPROVAL
    # ---------------------------------------------------------------
    st.markdown("### 3. Approve")

    approve_key = f"approved_{filename}"
    if approve_key not in st.session_state:
        st.session_state[approve_key] = False

    st.caption(
        "If the summary is correct, approval is one click. "
        "You only need to open the edit section when a value needs changing."
    )

    if st.button(
        "✅ Approve Candidate",
        key=f"approve_btn_{filename}",
        use_container_width=True,
    ):
        st.session_state[approve_key] = True
        st.success("Candidate approved for Skill Groomers submission.")

    if st.session_state[approve_key]:
        st.success("✅ Approved")
    else:
        st.info("Not approved yet.")

    with st.expander("Technical payload preview"):
        st.json(reviewed_payload)

    transfer = build_skill_groomers_form_transfer(
        data=data,
        reviewed_payload=reviewed_payload,
        selected_core=selected_core,
        selected_func_item=selected_func_item,
        selected_industry=selected_industry,
        selected_keywords=selected_keywords,
        selected_city=selected_city,
        selected_state=selected_state,
        selected_native_city=selected_native_city,
        selected_native_state=selected_native_state,
        selected_edu_item=selected_edu_item,
    )
    transfer_url = skill_groomers_transfer_url(transfer)

    if st.session_state[approve_key]:
        st.markdown(
            f"""
            <a href="{transfer_url}" target="_blank" rel="noopener noreferrer"
               style="
                   display:block;
                   width:100%;
                   text-align:center;
                   padding:0.72rem 1rem;
                   border-radius:0.55rem;
                   background:#433adb;
                   color:white;
                   font-weight:700;
                   text-decoration:none;
                   margin-top:0.55rem;
               ">
               ⚡ Open Skill Groomers & Fill Candidate
            </a>
            """,
            unsafe_allow_html=True,
        )
        st.caption(
            "Requires the Skill Groomers AutoFill Chrome extension. "
            "The extension fills the form; HR still reviews and clicks Skill Groomers' own Save button."
        )
    else:
        st.button(
            "⚡ Open Skill Groomers & Fill Candidate",
            key=f"export_sg_disabled_{filename}",
            use_container_width=True,
            disabled=True,
            help="Approve the candidate first.",
        )



# ---------- Compact review/debug display helpers ----------

def _clean_display(value):
    if value in (None, "", [], {}):
        return "—"
    return str(value)


def _info_grid(items, columns=3):
    """Display label/value pairs in a compact, consistent grid."""
    cols = st.columns(columns)
    for index, (label, value) in enumerate(items):
        with cols[index % columns]:
            st.markdown(
                f"""
                <div style="
                    border:1px solid #e5e7eb;
                    border-radius:10px;
                    padding:10px 12px;
                    margin-bottom:10px;
                    background:#ffffff;
                    min-height:72px;
                ">
                    <div style="
                        font-size:11px;
                        text-transform:uppercase;
                        letter-spacing:.04em;
                        color:#6b7280;
                        font-weight:700;
                        margin-bottom:4px;
                    ">{label}</div>
                    <div style="
                        font-size:14px;
                        line-height:1.35;
                        color:#111827;
                        font-weight:600;
                        word-break:break-word;
                    ">{_clean_display(value)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def _status_badge(label, ok=True):
    bg = "#ecfdf5" if ok else "#fff7ed"
    border = "#a7f3d0" if ok else "#fed7aa"
    color = "#047857" if ok else "#c2410c"
    st.markdown(
        f"""
        <div style="
            display:inline-block;
            padding:6px 10px;
            border-radius:999px;
            border:1px solid {border};
            background:{bg};
            color:{color};
            font-size:12px;
            font-weight:700;
            margin:2px 0 8px 0;
        ">{label}</div>
        """,
        unsafe_allow_html=True,
    )




def _metric_card(label, value):
    """Compact metric card that always shows the full value without Streamlit ellipsis."""
    shown = _clean_display(value)
    st.markdown(
        f"""
        <div style="
            border:1px solid #2f3440;
            border-radius:10px;
            padding:10px 12px;
            background:transparent;
            min-height:72px;
        ">
            <div style="
                font-size:12px;
                color:#cbd5e1;
                font-weight:600;
                margin-bottom:5px;
            ">{label}</div>
            <div style="
                font-size:18px;
                line-height:1.25;
                color:#ffffff;
                font-weight:600;
                white-space:normal;
                overflow:visible;
                word-break:normal;
            ">{shown}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )




def _cv_source_note(label, value):
    shown = _clean_display(value)
    st.markdown(
        f'<div class="review-source-note"><strong>{label}:</strong> {shown}</div>',
        unsafe_allow_html=True,
    )


# ---------- Streamlit UI ----------

st.set_page_config(page_title="CV Analyzer", page_icon="📄", layout="centered")
st.markdown("""<style>
.main-title { text-align: center; font-size: 38px; font-weight: 700; margin-bottom: 5px; }
.subtitle { text-align: center; color: #666; margin-bottom: 25px; }



/* Review-page status polish */
.review-source-note {
    color: #8b93a1;
    font-size: 0.82rem;
    line-height: 1.35;
    margin: 2px 0 7px 0;
}

/* Multiselect selections are choices, not errors */
[data-baseweb="tag"] {
    background-color: #374151 !important;
    color: #f8fafc !important;
}
[data-baseweb="tag"] span,
[data-baseweb="tag"] svg {
    color: #f8fafc !important;
}

/* Avoid using red as the default focus state for valid review controls */
div[data-baseweb="select"] > div:focus-within {
    border-color: #64748b !important;
    box-shadow: 0 0 0 1px #64748b !important;
}


/* Faster HR review: tighter vertical rhythm */
[data-testid="stExpander"] {
    margin-top: 0.35rem !important;
    margin-bottom: 0.65rem !important;
}

</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📄 CV Analyzer V8.5</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload CVs, verify the extracted facts, review Skill Groomers mapping, then download the HR report.</div>', unsafe_allow_html=True)
st.markdown(
    """
    <div class="app-hero">
        <div class="eyebrow">Skill Groomers • Recruitment Intelligence</div>
        <h1>CV Analyzer V8.5</h1>
        <p>Convert resumes into verified HR-ready candidate profiles with deterministic experience calculations and review-safe Skill Groomers mapping.</p>
        <div class="hero-meta">
            <span class="hero-pill">CV Fact Extraction</span>
            <span class="hero-pill">Experience Validation</span>
            <span class="hero-pill">Skill Groomers Mapping</span>
            <span class="hero-pill">HR Review Workflow</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="premium-card" style="margin-bottom:14px">
        <div class="section-kicker">Workflow</div>
        <div style="display:flex;gap:18px;flex-wrap:wrap;font-weight:700;color:#334155">
            <span>1&nbsp; Upload CV</span>
            <span>→</span>
            <span>2&nbsp; Analyze</span>
            <span>→</span>
            <span>3&nbsp; Review</span>
            <span>→</span>
            <span>4&nbsp; Approve</span>
            <span>→</span>
            <span>5&nbsp; Export / Submit</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.info("Upload one or more PDF CVs. Factual CV data stays separate from Skill Groomers mapping so HR can review classifications without changing source facts.")

st.sidebar.markdown("### Developer Tools")
developer_debug = st.sidebar.checkbox(
    "Show Developer Debug",
    value=False,
    help="Shows raw extraction, normalized data, full mapping details and API payload information.",
)

uploaded_files = st.file_uploader("Upload CV(s)", type=["pdf"], accept_multiple_files=True, help="Only PDF files are accepted.")
if "generated_files" not in st.session_state:
    st.session_state.generated_files = {}
if uploaded_files:
    st.success(f"Selected {len(uploaded_files)} file(s).")

if st.button("🔍 Analyse CVs", use_container_width=True):
    if not uploaded_files:
        st.warning("Please upload at least one PDF CV first.")
    else:
        st.session_state.generated_files = {}
        for uploaded_file in uploaded_files:
            st.write(f"--- Processing: **{uploaded_file.name}** ---")

            try:
                with st.spinner(f"Analysing {uploaded_file.name}..."):
                    pdf_bytes = uploaded_file.getvalue()
                    raw_text = extract_text_from_pdf(uploaded_file)
                    raw_gemini_data = extract_resume_data(raw_text, pdf_bytes=pdf_bytes)

                    # V8: normalize only factual presentation/field semantics first.
                    # This remains CV-derived data and is separate from SG classification.
                    data = normalize_factual_education(raw_gemini_data)

                    # Safety guard: a CV that visibly contains professional-experience
                    # language but yields no jobs should never silently show 0 years.
                    experience_words = normalize_master_text(raw_text)
                    extraction_warning = (
                        any(term in experience_words for term in [
                            "professional experience", "work experience", "employment",
                            "career experience", "experience"
                        ])
                        and not (data.get("work_experience") or [])
                    )
                    data["_employment_extraction_warning"] = extraction_warning

                    # Map factual CV extraction against the complete Skill Groomers master-data snapshot.
                    mapping = build_skill_groomers_mapping(data, source_text=raw_text)
                    excel_data = apply_mapping_to_excel_data(data, mapping)

                    if developer_debug:
                        # ----------------------------------------------------------
                        # TWO-LEVEL DEBUG / REVIEW VIEW
                        # Level 1 = HR-readable operational diagnostics.
                        # Level 2 = complete developer diagnostics.
                        # ----------------------------------------------------------
                        debug_jobs = sort_jobs(data.get("work_experience", []))
                        debug_total_months = total_experience_months(debug_jobs)
                        debug_job_changes = compute_job_changes(debug_jobs)
                        debug_avg_months = average_tenure_months(
                            debug_total_months, debug_job_changes
                        )
                        debug_current, debug_previous = get_current_and_previous_jobs(debug_jobs)
                        factual_education = data.get("highest_qualification", {}) or {}

                        sg_core = mapping.get("core_role", {}) or {}
                        sg_functional = mapping.get("functional_area", {}) or {}
                        sg_industry = mapping.get("industry", {}) or {}
                        sg_education = mapping.get("education", {}) or {}
                        sg_city = mapping.get("current_city", {}) or {}
                        sg_state = mapping.get("current_state", {}) or {}

                        # -------------------------
                        # LEVEL 1 — HR
                        # -------------------------
                        # Developer-only technical diagnostics
                        if developer_debug:
                            factual_debug = {
                                "candidate_name": data.get("candidate_name", ""),
                                "email": data.get("email", ""),
                                "phone": data.get("phone", ""),
                                "date_of_birth": data.get("date_of_birth", ""),
                                "gender": data.get("gender", ""),
                                "current_city": data.get("current_city", ""),
                                "current_state": data.get("current_state", ""),
                                "native_city": data.get("native_city", ""),
                                "native_state": data.get("native_state", ""),
                                "core_role": data.get("core_role", ""),
                                "key_skills": data.get("key_skills", []),
                                "functional_area": data.get("functional_area", ""),
                                "role": data.get("role", ""),
                                "industry": data.get("industry", ""),
                                "current_employment": debug_current,
                                "previous_distinct_employment": debug_previous,
                                "calculated_total_experience": format_months(debug_total_months),
                                "number_of_employers": number_of_employers(debug_jobs),
                                "job_changes": debug_job_changes,
                                "calculated_average_tenure": format_months(debug_avg_months),
                                "highest_completed_qualification": factual_education,
                            }

                            skill_groomers_debug = {
                                "core_role": sg_core.get("name", ""),
                                "core_role_confidence": sg_core.get("score"),
                                "functional_area": sg_functional.get("functional_area", ""),
                                "sub_functional_area": sg_functional.get("sub_functional_area", ""),
                                "role": sg_functional.get("role", ""),
                                "functional_area_id": sg_functional.get("functionalAreaId"),
                                "industry": sg_industry.get("name", ""),
                                "industry_id": sg_industry.get("industryId"),
                                "key_skills": mapping.get("key_skills", []),
                                "education_qualification": sg_education.get("qualification", ""),
                                "education_course": sg_education.get("course", ""),
                                "education_specialization": sg_education.get("specialization", ""),
                                "education_id": sg_education.get("educationId"),
                                "current_city": sg_city.get("name", ""),
                                "current_city_id": sg_city.get("id"),
                                "current_state": sg_state.get("name", ""),
                                "current_state_id": sg_state.get("id"),
                                "native_city": mapping.get("native_city", {}).get("name", ""),
                                "native_city_id": mapping.get("native_city", {}).get("id"),
                                "native_state": mapping.get("native_state", {}).get("name", ""),
                                "native_state_id": mapping.get("native_state", {}).get("id"),
                                "review_required": mapping.get("review_required", False),
                                "review_reasons": mapping.get("review_reasons", []),
                            }

                            st.subheader(f"🛠️ Developer Debug — {uploaded_file.name}")
                            st.caption(
                                "Technical pipeline view for extraction, normalization, mapping and payload troubleshooting."
                            )

                            dev_tab1, dev_tab2, dev_tab3, dev_tab4 = st.tabs(
                                [
                                    "Pipeline Summary",
                                    "Extraction",
                                    "Mapping",
                                    "API Payload",
                                ]
                            )

                            with dev_tab1:
                                d1, d2, d3, d4 = st.columns(4)
                                d1.metric("Experience", format_months(debug_total_months))
                                d2.metric("Avg. Tenure", format_months(debug_avg_months))
                                d3.metric("Employers", number_of_employers(debug_jobs))
                                d4.metric("Job Changes", debug_job_changes)

                                _info_grid(
                                    [
                                        ("Candidate", data.get("candidate_name", "")),
                                        ("Core Role", sg_core.get("name", "")),
                                        ("Functional Area ID", sg_functional.get("functionalAreaId")),
                                        ("Industry ID", sg_industry.get("industryId")),
                                        ("Education ID", sg_education.get("educationId")),
                                        ("Review Required", mapping.get("review_required", False)),
                                    ],
                                    columns=3,
                                )

                                if mapping.get("review_required"):
                                    st.warning("One or more mapping fields require review.")
                                    for item in mapping.get("review_reasons", []):
                                        st.write(
                                            f"• **{item.get('section', 'Mapping')}** — "
                                            f"{item.get('reason', 'Review required')}"
                                        )
                                else:
                                    st.success("No mapper review flags.")

                            with dev_tab2:
                                with st.expander("CV / Excel Facts", expanded=True):
                                    st.json(factual_debug)
                                with st.expander("Raw Gemini JSON"):
                                    st.json(raw_gemini_data)
                                with st.expander("Normalized Factual Data"):
                                    st.json(data)

                            with dev_tab3:
                                with st.expander("Skill Groomers Mapping Summary", expanded=True):
                                    st.json(skill_groomers_debug)
                                with st.expander("Full Skill Groomers Mapping Details"):
                                    st.json(mapping)

                                if sg_functional.get("matched"):
                                    st.caption(
                                        "Functional hierarchy: "
                                        f"{sg_functional.get('functional_area', '')} → "
                                        f"{sg_functional.get('sub_functional_area', '')} → "
                                        f"{sg_functional.get('role', '')} "
                                        f"(ID {sg_functional.get('functionalAreaId')})"
                                    )

                            with dev_tab4:
                                skill_groomers_payload = build_skill_groomers_payload_preview(
                                    excel_data, mapping
                                )
                                st.caption(
                                    "Payload preview only. The resume URL is inserted later during real submission."
                                )
                                st.json(skill_groomers_payload)

                    # Load template and populate sheet
                    workbook = openpyxl.load_workbook("resume1.xlsx")
                    sheet = workbook.active
                    populate_excel(excel_data, sheet)

                    # Save to memory buffer
                    excel_buffer = io.BytesIO()
                    workbook.save(excel_buffer)
                    excel_buffer.seek(0)


                st.success(f"✅ Finished analyzing {uploaded_file.name}!")
                skill_groomers_payload = build_skill_groomers_payload_preview(excel_data, mapping)
                st.session_state.generated_files[uploaded_file.name] = {
                    "data": excel_buffer.getvalue(),
                    "candidate_name": data.get("candidate_name", uploaded_file.name),
                    "mapping": mapping,
                    "factual_data": data,
                    "payload_preview": skill_groomers_payload,
                }


            except Exception as error:
                st.error(f"❌ Something went wrong with {uploaded_file.name}.")
                with st.expander("View technical error"):
                    st.exception(error)

if st.session_state.generated_files:

    st.divider()
    st.markdown("## Candidate Review")
    st.caption("Verify CV facts first, then review Skill Groomers classifications. Experience calculations remain locked to the agreed rules.")

    for filename, file_info in st.session_state.generated_files.items():
        with st.container(border=True):
            _render_review_card(filename, file_info)

            st.download_button(
                label=f"⬇️ Download Excel Report for {file_info['candidate_name']}",
                data=file_info["data"],
                file_name=f"{file_info['candidate_name']}_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{filename}",
                use_container_width=True,
            )
