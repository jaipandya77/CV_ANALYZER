import streamlit as st
import PyPDF2
import csv
import json
import openpyxl

from google import genai
from datetime import datetime


# ============================================================
# GEMINI
# ============================================================

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

MODEL = "gemini-3.5-flash-lite"


# ============================================================
# PDF
# ============================================================

def extract_text_from_pdf(file):

    reader = PyPDF2.PdfReader(file)

    return "\n".join(
        page.extract_text() or ""
        for page in reader.pages
    )


# ============================================================
# GEMINI
# ============================================================

resume_schema = {

    "type": "object",

    "properties": {

        "candidate_name": {
            "type": "string"
        },

        "highest_qualification": {

            "type": "object",

            "properties": {

                "degree": {
                    "type": "string"
                },

                "specialization": {
                    "type": "string"
                },

                "institute": {
                    "type": "string"
                },

                "year": {
                    "type": "string"
                }
            },

            "required": [
                "degree",
                "specialization",
                "institute",
                "year"
            ]
        },

        "work_experience": {

            "type": "array",

            "items": {

                "type": "object",

                "properties": {

                    "company": {
                        "type": "string"
                    },

                    "client_or_project_owner": {
                        "type": "string"
                    },

                    "designation": {
                        "type": "string"
                    },

                    "start_date": {
                        "type": "string"
                    },

                    "end_date": {
                        "type": "string"
                    },

                    "country": {
                        "type": "string"
                    },

                    "is_mnc": {
                        "type": "boolean"
                    },

                    "is_listed": {
                        "type": "boolean"
                    },

                    "organization_category": {

                        "type": "string",

                        "enum": [
                            "Owner",
                            "Contractor",
                            "Consultant",
                            "Freelancer",
                            "Unknown"
                        ]
                    }
                },

                "required": [

                    "company",
                    "client_or_project_owner",
                    "designation",
                    "start_date",
                    "end_date",
                    "country",
                    "is_mnc",
                    "is_listed",
                    "organization_category"
                ]
            }
        },

        "certifications": {

            "type": "array",

            "items": {
                "type": "string"
            }
        }
    },

    "required": [

        "candidate_name",
        "highest_qualification",
        "work_experience",
        "certifications"
    ]
}


def extract_resume_data(raw_text):

    # KEEP YOUR EXISTING PROMPT EXACTLY HERE
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
        config={
            "response_mime_type": "application/json",
            "response_schema": resume_schema
        }
    )

    return json.loads(response.text)


# ============================================================
# DATES
# ============================================================

def parse_start_date(value):

    return datetime.strptime(
        value.strip(),
        "%Y-%m"
    )


def parse_end_date(value):

    value = value.strip().lower()

    if value in [
        "present",
        "current",
        "ongoing",
        "till date"
    ]:
        return datetime.today()

    return datetime.strptime(
        value,
        "%Y-%m"
    )


def get_job_years(job):

    try:

        start = parse_start_date(
            job["start_date"]
        )

        end = parse_end_date(
            job["end_date"]
        )

        return max(
            0,
            (end - start).days / 365.25
        )

    except Exception:

        return 0


def sort_jobs(jobs):

    return sorted(
        jobs,
        key=lambda job: parse_start_date(
            job["start_date"]
        )
    )


# ============================================================
# CALCULATIONS
# ============================================================

def experience_by_field(jobs, field, value):

    return round(
        sum(
            get_job_years(job)
            for job in jobs
            if str(
                job.get(field, "")
            ).strip().lower()
            == value.lower()
        ),
        1
    )


def experience_by_boolean_field(jobs, field):

    return round(
        sum(
            get_job_years(job)
            for job in jobs
            if job.get(field) is True
        ),
        1
    )


def total_experience(jobs):

    return round(
        sum(
            get_job_years(job)
            for job in jobs
        ),
        1
    )


def employment_gap(jobs):

    jobs = sorted(
        jobs,
        key=lambda job: parse_start_date(job["start_date"])
    )

    largest_gap = 0

    for previous, current in zip(jobs, jobs[1:]):

        gap = (
            parse_start_date(current["start_date"]) -
            parse_end_date(previous["end_date"])
        ).days / 365.25

        if gap > largest_gap:
            largest_gap = gap

    if largest_gap > 1:
        return round(largest_gap, 1)

    return "No Gap"


def average_tenure(total_years, job_count):

    changes = job_count - 1

    return (
        round(total_years / changes, 1)
        if changes > 0
        else 0
    )


# ============================================================
# NIRF
# ============================================================

def normalize_text(value):

    return " ".join(
        str(value)
        .lower()
        .replace("&", "and")
        .split()
    )


def get_nirf_ranking(institute):

    if not institute:
        return "After 200"

    target = normalize_text(institute)

    try:

        with open(
            "nirf_rankings.csv",
            encoding="utf-8-sig"
        ) as file:

            for row in csv.DictReader(file):

                if normalize_text(
                    row.get("institute", "")
                ) == target:

                    category = row.get(
                        "category",
                        "After 200"
                    )

                    if category.lower() == "top 100":
                        return "top 100"

                    if category.lower() == "101-200":
                        return "101-200"

                    return category

    except Exception:

        pass

    return "After 200"


# ============================================================
# EXCEL
# ============================================================

def write_row(
    sheet,
    workbook,
    cell,
    label,
    value
):

    sheet[cell] = value

    st.write(
        f"{label}:",
        value
    )




def populate_excel(
    data,
    workbook,
    sheet
):

    # -------------------------------
    # C2 - Candidate
    # -------------------------------

    write_row(
        sheet,
        workbook,
        "C2",
        "Candidate Name",
        data.get(
            "candidate_name",
            ""
        )
    )

    # -------------------------------
    # C4 - Qualification
    # -------------------------------

    qualification = data.get(
        "highest_qualification",
        {}
    )
    degree = qualification.get(
        "degree",
        ""
    )

    specialization = qualification.get(
        "specialization",
        ""
    )

    value = (
        f"{degree} - {specialization}"
        if (
            specialization
            and specialization.lower()
            not in degree.lower()
        )
        else degree
    )

    write_row(
        sheet,
        workbook,
        "C4",
        "Highest Qualification",
        value
    )

    # -------------------------------
    # C5 - NIRF
    # -------------------------------

    institute = qualification.get(
        "institute",
        ""
    )

    ranking = get_nirf_ranking(
        institute
    )

    st.write(
        "Educational Institute:",
        institute
    )

    write_row(
        sheet,
        workbook,
        "C5",
        "Institute Ranking",
        ranking
    )

    # -------------------------------
    # JOBS
    # -------------------------------

    jobs = sort_jobs(
        data.get(
            "work_experience",
            []
        )
    )

    if debug:

        st.subheader(
            "Extracted Work Experience"
        )

        for i, job in enumerate(
            jobs,
            1
        ):

            st.write(
                f"Job {i}:",
                job
            )

            st.write(
                "Calculated Years:",
                round(
                    get_job_years(job),
                    1
                )
            )
    # -------------------------------
    # C6 - Total Experience
    # -------------------------------

    total = total_experience(
        jobs
    )

    write_row(
        sheet,
        workbook,
        "C6",
        "Work Experience",
        total
    )

    # -------------------------------
    # C7-C10 - HR Categories
    # -------------------------------

    organization_cells = {
        "Contractor": "C7",
        "Owner": "C8",
        "Consultant": "C9",
        "Freelancer": "C10"
    }

    for category, cell in (
        organization_cells.items()
    ):

        value = experience_by_field(
            jobs,
            "organization_category",
            category
        )

        write_row(
            sheet,
            workbook,
            cell,
            f"{category} Experience",
            value
        )

    # -------------------------------
    # C11-C12 - MNC / Listed
    # -------------------------------

    company_cells = {
        "is_mnc": (
            "C11",
            "MNC Experience"
        ),
        "is_listed": (
            "C12",
            "Listed Company Experience"
        )
    }

    for field, (cell, label) in (
        company_cells.items()
    ):

        value = experience_by_boolean_field(
            jobs,
            field
        )

        write_row(
            sheet,
            workbook,
            cell,
            label,
            value
        )

    # -------------------------------
    # C13-C14 - Country
    # -------------------------------

    india = experience_by_field(
        jobs,
        "country",
        "India"
    )

    write_row(
        sheet,
        workbook,
        "C13",
        "India Experience",
        india
    )

    write_row(
        sheet,
        workbook,
        "C14",
        "Outside India Experience",
        round(
            total - india,
            1
        )
    )

    # -------------------------------
    # C15-C18
    # -------------------------------
    changes = 0

    previous_company = None

    for job in jobs:

        company = normalize_text(
            job.get(
                "company",
                ""
            )
        )

        if previous_company is not None:

            if company != previous_company:
                changes += 1

        previous_company = company

    

    write_row(
        sheet,
        workbook,
        "C15",
        "Job Changes",
        changes
    )

    write_row(
        sheet,
        workbook,
        "C16",
        "Average Tenure",
        average_tenure(
            total,
            len(jobs)
        )
    )

    write_row(
        sheet,
        workbook,
        "C17",
        "Employment Gap",
        employment_gap(jobs)
    )

    certifications = data.get(
        "certifications",
        []
    )

    write_row(
        sheet,
        workbook,
        "C18",
        "Certifications",
        len(certifications)
    )

    if certifications:

        st.subheader(
            "Certifications Extracted"
        )

        for certification in certifications:

            st.write(
                "-",
                certification
            )


# ============================================================
# STREAMLIT
# ============================================================




st.set_page_config(
    page_title="CV Analyser",
    page_icon="📄",
    layout="centered"
)

st.markdown(
    """
    <style>
    .main-title {
        text-align: center;
        font-size: 38px;
        font-weight: 700;
        margin-bottom: 5px;
    }
    .subtitle {
        text-align: center;
        color: #666;
        margin-bottom: 25px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    '<div class="main-title">📄 CV Analyzer</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Upload a CV to analyse experience and generate the HR report.</div>',
    unsafe_allow_html=True
)

st.info(
    "Upload a PDF resume. The application extracts candidate information using Gemini AI and automatically fills the HR Excel template."
)

debug = st.sidebar.checkbox(
    "Show Debug Information"
)

uploaded_file = st.file_uploader(
    "Upload CV",
    type=["pdf"],
    help="Only PDF files are accepted."
)

if uploaded_file:

    st.success(
        f"Selected: {uploaded_file.name}"
    )

if st.button(
    "🔍 Analyse CV",
    use_container_width=True
):

    if uploaded_file is None:

        st.warning(
            "Please upload a PDF CV first."
        )

    else:

        try:

            with st.spinner(
                "Analysing CV and preparing report..."
            ):

                raw_text = extract_text_from_pdf(
                    uploaded_file
                )

                data = extract_resume_data(
                    raw_text
                )

                if debug:
                    st.subheader(
                        "📋 Extracted Information"
                    )

                    st.json(data)

                workbook = openpyxl.load_workbook(
                    "HR_Template.xlsx"
                )

                sheet = workbook.active

                populate_excel(
                    data,
                    workbook,
                    sheet
                )

            workbook.save("CV_Output.xlsx")
            st.success(
                "✅ CV analysed successfully!"
            )

            with open(
                "CV_Output.xlsx",
                "rb"
            ) as excel_file:

                st.download_button(
                    "⬇️ Download HR Report",
                    data=excel_file,
                    file_name="CV_Output.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

        except Exception as error:

            st.error(
                "❌ Something went wrong."
            )

            with st.expander(
                "View technical error"
            ):
                st.exception(error)
