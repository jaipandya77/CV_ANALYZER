import io
import json
import csv
import re
from datetime import datetime
import streamlit as st
import PyPDF2
import openpyxl
from google import genai

# === GEMINI CONFIGURATION ===
client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
MODEL = "gemini-3.5-flash-lite"

# === PDF EXTRACTOR ===
def extract_text_from_pdf(file):
    reader = PyPDF2.PdfReader(file)
    text = []
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text.append(extracted)
    return "\n".join(text)

# === RESUME SCHEMA (unchanged) ===
resume_schema = {
    "type": "object",
    "properties": {
        "candidate_name": {"type": "string"},
        "highest_qualification": {
            "type": "object",
            "properties": {
                "degree": {"type": "string"},
                "specialization": {"type": "string"},
                "institute": {"type": "string"},
                "year": {"type": "string"}
            },
            "required": ["degree", "specialization", "institute", "year"]
        },
        "work_experience": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "company": {"type": "string"},
                    "client_or_project_owner": {"type": "string"},
                    "designation": {"type": "string"},
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"},
                    "country": {"type": "string"},
                    "is_mnc": {"type": "boolean"},
                    "is_listed": {"type": "boolean"},
                    "organization_category": {
                        "type": "string",
                        "enum": ["Owner", "Contractor", "Consultant", "Freelancer", "Unknown"]
                    }
                },
                "required": [
                    "company", "client_or_project_owner", "designation",
                    "start_date", "end_date", "country",
                    "is_mnc", "is_listed", "organization_category"
                ]
            }
        },
        "certifications": {"type": "array", "items": {"type": "string"}}
    },
    "required": ["candidate_name", "highest_qualification", "work_experience", "certifications"]
}

# === GEMINI EXTRACTION ===
def extract_resume_data(raw_text):
    prompt = f"""
Extract structured HR information from this resume.

Return only information supported by the resume.
Never invent information.

Accuracy is more important than completeness.

If information is uncertain, leave the field empty instead of guessing.

Do not calculate experience, employment gaps, average tenure or job changes.

Only extract facts.

Python will perform all calculations.

============================================================
DATES
============================================================

Return dates in YYYY-MM format.

Examples:

January 2018 -> 2018-01
June 2020 -> 2020-06
14/12/2020 -> 2020-12

CRITICAL DATE RULE:
- end_date MUST ALWAYS be equal to or later than start_date.
- Never output an end_date that takes place before the start_date.
- If a date digit looks ambiguous (e.g., 2016 vs 2006), cross-reference with surrounding job history to ensure chronological consistency.

If currently employed:

"Present"

Do not calculate experience.

============================================================
HIGHEST QUALIFICATION
============================================================

Return the highest COMPLETED qualification.

Do not treat ongoing education as the highest qualification.

Extract:

degree
specialization
institute
year

Only provide specialization if it is clearly associated with
the qualification.

Only provide year if the year is clearly associated with that
specific qualification.

Do not guess a year from a nearby unrelated date.

If the institute is clearly shown in the Education section,
extract it. Do not return Unknown when the institute is visible.

============================================================
WORK EXPERIENCE
============================================================

Extract EVERY professional employment position in the resume.

IMPORTANT:

Do not stop after recent jobs.

Do not omit older jobs.

Include entry-level positions if they are actual employment,
such as:

Graduate Engineer Trainee
Junior Engineer
Site Engineer
Trainee Engineer
Engineer

Extract every distinct employer/position listed in sections such as:

Work Experience
Professional Experience
Employment History
Career History
Career Recitals
Professional Experience / Career Recitals

Do not include:

Internships
Training
Academic projects
College projects
Volunteer work

Return jobs in chronological order, oldest first.

============================================================
DIRECT EMPLOYER
============================================================

"company" means the organisation that directly employed the candidate.

"client_or_project_owner" means the client, project owner,
government authority, developer, or main organisation for whom
the project was being executed.

Do not put the client/project owner in the company field.

Example:

company:
Hafiz Construction Co. Pvt. Ltd.

client_or_project_owner:
IRCON International

============================================================
COUNTRY
============================================================

Return the actual country name.

Examples:

India
Qatar
Oman
Saudi Arabia
United Kingdom

Do not return city or state names.

============================================================
MNC
============================================================

"is_mnc" is true only when the DIRECT EMPLOYER is a multinational
or global organisation.

Otherwise false.

============================================================
LISTED
============================================================

"is_listed" is true only when the DIRECT EMPLOYER is a publicly
listed/traded company.

Otherwise false.

MNC and Listed are independent.

A company may be:

MNC only
Listed only
both MNC and Listed
neither

Do not force these into one category.

Do not classify based only on the client or project owner.

============================================================
ORGANIZATION CATEGORY
============================================================

Use these exact values:

Owner
Contractor
Consultant
Freelancer
Unknown

OWNER:

An Owner is the organisation/person that pays for and owns
the built asset.

Government departments and real-estate developers belong
to the Owner category.

CONTRACTOR:

A Contractor is an organisation that builds the asset.

Examples include civil construction, infrastructure and EPC
organisations such as:

L&T
AFCONS
Capacite

Classify the candidate as Contractor when the DIRECT EMPLOYER
is a construction/EPC/contracting organisation.

Do not classify based only on project descriptions.

CONSULTANT:

A Consultant is an organisation/person that provides advisory
services to others.

Classify as Consultant when the DIRECT EMPLOYER is an engineering,
infrastructure, management consulting, PMC, or advisory
organisation and the candidate's work is advisory/consulting/
supervisory in nature.

Do not classify as Consultant merely because:

- the project description mentions consultants
- the candidate coordinates with consultants

FREELANCER:

Return Freelancer only when the resume clearly indicates thats
the candidate works independently rather than as an employee
of an organisation.

Otherwise do not use Freelancer.

UNKNOWN:

Use Unknown when the organisation cannot confidently be classified.

============================================================
CERTIFICATIONS
============================================================

Return actual professional certifications as strings.

Do not count:

- ordinary training
- seminars
- workshops
- memberships

unless clearly presented as certifications.

If none exist:

[]

============================================================
CALCULATIONS
============================================================

Do not calculate:

Total experience
Job changes
Average tenure
Employment gaps
India experience
Outside India experience
MNC experience
Listed experience
Owner experience
Contractor experience
Consultant experience
Freelancer experience

Python will calculate all of these.

============================================================
IMPORTANT
============================================================

Never invent:

- employers
- dates
- countries
- qualifications
- company classifications
- organization classifications

The resume is the source of truth.

RESUME:

{raw_text}
"""
    response = client.models.generate_content(
        model=MODEL,
        contents=prompt,
        config={"response_mime_type": "application/json", "response_schema": resume_schema}
    )
    return json.loads(response.text)

# === DATE HELPERS (With Robust Error Handling) ===
def parse_start_date(value):
    if not value or not isinstance(value, str):
        return datetime(1900, 1, 1)
    val = value.strip()
    try:
        return datetime.strptime(val, "%Y-%m")
    except ValueError:
        # Fallback if AI provides YYYY only
        match = re.search(r"\d{4}", val)
        if match:
            return datetime(int(match.group(0)), 1, 1)
        return datetime(1900, 1, 1)

def parse_end_date(value):
    if not value or not isinstance(value, str):
        return datetime.today()
    v = value.strip().lower()
    if v in ["present", "current", "ongoing", "till date", "now"]:
        return datetime.today()
    try:
        return datetime.strptime(v, "%Y-%m")
    except ValueError:
        match = re.search(r"\d{4}", v)
        if match:
            return datetime(int(match.group(0)), 12, 31)
        return datetime.today()

def get_job_years(job):
    try:
        start = parse_start_date(job.get("start_date", ""))
        end = parse_end_date(job.get("end_date", ""))
        if end < start:
            return 0.0
        return max(0.0, (end - start).days / 365.25)
    except Exception as e:
        return 0.0

def sort_jobs(jobs): 
    return sorted(jobs, key=lambda j: parse_start_date(j.get("start_date", "")))

# === EXPERIENCE CALCULATIONS ===
def experience_by_field(jobs, field, value):
    return round(sum(get_job_years(j) for j in jobs if str(j.get(field, "")).strip().lower() == value.lower()), 1)

def experience_by_field_not_equal(jobs, field, value):
    return round(sum(get_job_years(j) for j in jobs if str(j.get(field, "")).strip().lower() != value.lower()), 1)

def experience_by_boolean_field(jobs, field):
    return round(sum(get_job_years(j) for j in jobs if j.get(field) is True), 1)

def total_experience(jobs): 
    return round(sum(get_job_years(j) for j in jobs), 1)

def employment_gap(jobs):
    sorted_j = sort_jobs(jobs)
    if len(sorted_j) < 2:
        return "No Gap"
    
    largest_gap = 0.0
    for prev, curr in zip(sorted_j, sorted_j[1:]):
        prev_end = parse_end_date(prev.get("end_date", ""))
        curr_start = parse_start_date(curr.get("start_date", ""))
        gap = (curr_start - prev_end).days / 365.25
        if gap > largest_gap:
            largest_gap = gap
            
    return round(largest_gap, 1) if largest_gap >= 1 else "No Gap"

def average_tenure(total_years, job_changes):
    if job_changes <= 0:
        return round(total_years,1)
    return round(total_years / job_changes, 1)

# === NIRF RANKING (With Partial Matching) ===
def normalize_text(value):
    if not value:
        return ""
    return " ".join(re.sub(r'[^a-zA-Z0-9\s]', '', str(value)).lower().replace("&", "and").split())

def get_nirf_ranking(institute):
    if not institute: 
        return "After 200"
    target = normalize_text(institute)
    try:
        with open("nirf_rankings.csv", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                row_inst = normalize_text(row.get("institute", ""))
                # Flexible partial matching logic
                if target in row_inst or row_inst in target:
                    cat = row.get("category", "After 200").strip()
                    if cat.lower() == "top 100": 
                        return "top 100"
                    if cat.lower() == "101-200": 
                        return "101-200"
                    return cat
    except Exception:
        pass
    return "After 200"

# === EXCEL OUTPUT ===
def write_row(sheet, cell, label, value):
    sheet[cell] = value
    st.write(f"**{label}:** {value}")

def populate_excel(data, sheet):
    # Candidate name
    write_row(sheet, "C2", "Candidate Name", data.get("candidate_name", ""))

    # Highest qualification
    qual = data.get("highest_qualification", {})
    deg = qual.get("degree", "")
    spec = qual.get("specialization", "")
    val = f"{deg} - {spec}" if (spec and spec.lower() not in deg.lower()) else deg
    write_row(sheet, "C4", "Highest Qualification", val)

    # Institute & NIRF ranking
    inst = qual.get("institute", "")
    st.write(f"**Educational Institute:** {inst}")
    write_row(sheet, "C5", "Institute Ranking", get_nirf_ranking(inst))

    # Work experience jobs
    jobs = sort_jobs(data.get("work_experience", []))
    if debug:
        st.subheader("Extracted Work Experience")
        for i, job in enumerate(jobs, 1):
            st.write(f"**Job {i}:**", job)
            st.write("Calculated Years:", round(get_job_years(job), 1))

    total = total_experience(jobs)
    write_row(sheet, "C6", "Work Experience", total)

    # Organisation categories
    org_cells = {"Contractor": "C7", "Owner": "C8", "Consultant": "C9", "Freelancer": "C10"}
    for cat, cell in org_cells.items():
        write_row(sheet, cell, f"{cat} Experience", experience_by_field(jobs, "organization_category", cat))

    # MNC and Listed
    company_cells = {"is_mnc": ("C11", "MNC Experience"), "is_listed": ("C12", "Listed Company Experience")}
    for field, (cell, label) in company_cells.items():
        write_row(sheet, cell, label, experience_by_boolean_field(jobs, field))

    # India vs outside India
    india = experience_by_field(jobs, "country", "India")
    outside_india = experience_by_field_not_equal(jobs, "country", "India")
    write_row(sheet, "C13", "India Experience", india)
    write_row(sheet, "C14", "Outside India Experience", outside_india)

    # Job changes (based on distinct company names)
    changes = 0
    prev_company = None
    for job in jobs:
        curr = normalize_text(job.get("company", ""))
        if prev_company is not None and curr != prev_company:
            changes += 1
        prev_company = curr
    write_row(sheet, "C15", "Job Changes", changes)

    write_row(sheet, "C16", "Average Tenure", average_tenure(total,changes))
    write_row(sheet, "C17", "Employment Gap", employment_gap(jobs))

    certs = data.get("certifications", [])
    write_row(sheet, "C18", "Certifications", len(certs))
    if certs:
        st.subheader("Certifications Extracted")
        for c in certs:
            st.write("-", c)

# === STREAMLIT UI ===
st.set_page_config(page_title="CV Analyzer", page_icon="📄", layout="centered")
st.markdown("""<style>
.main-title { text-align: center; font-size: 38px; font-weight: 700; margin-bottom: 5px; }
.subtitle { text-align: center; color: #666; margin-bottom: 25px; }
</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📄 CV Analyzer</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload a CV to analyse experience and generate the HR report.</div>', unsafe_allow_html=True)
st.info("Upload a PDF resume. The application extracts candidate information using Gemini AI and automatically fills the HR Excel template.")

debug = st.sidebar.checkbox("Show Debug Information")

uploaded_files = st.file_uploader("Upload CV(s)", type=["pdf"], accept_multiple_files= True, help="Only PDF files are accepted.")
if uploaded_files:
    st.success(f"Selected {len(uploaded_files)} file(s).")

if st.button("🔍 Analyse CVs", use_container_width=True):
  if not uploaded_files:
    st.warning("Please upload at least one PDF CV first.")
  else:
    # Loop through every uploaded file
    for uploaded_file in uploaded_files:
      st.write(f"--- Processing: **{uploaded_file.name}** ---")
      try:
        with st.spinner(f"Analysing {uploaded_file.name}..."):
          raw_text = extract_text_from_pdf(uploaded_file)
          data = extract_resume_data(raw_text)

          if debug:
            st.subheader(f"📋 Extracted JSON Data for {uploaded_file.name}")
            st.json(data)

          # Load template and populate sheet
          workbook = openpyxl.load_workbook("HR_Template.xlsx")
          sheet = workbook.active
          populate_excel(data, sheet)

          # Stream write directly into memory buffer
          excel_buffer = io.BytesIO()
          workbook.save(excel_buffer)
          excel_buffer.seek(0)

        st.success(f"✅ Finished analyzing {uploaded_file.name}!")

        # Provide a separate download button for each processed file
        st.download_button(
            label=f"⬇️ Download Report for {data.get('candidate_name', uploaded_file.name)}",
            data=excel_buffer,
            file_name=f"{data.get('candidate_name', 'CV')}_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=uploaded_file.name,  # Unique key is required for multiple buttons in a loop
            use_container_width=True,
        )

      except Exception as error:
        st.error(f"❌ Something went wrong with {uploaded_file.name}.")
        with st.expander("View technical error"):
          st.exception(error)
