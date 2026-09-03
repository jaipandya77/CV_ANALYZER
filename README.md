# 📄 CV Analyzer V8.5

An AI-assisted recruitment web application that extracts factual
information from PDF resumes, performs deterministic HR calculations,
maps candidates to **Skill Groomers** master data, generates structured
HR Excel reports, and helps HR transfer approved candidate information
into the existing Skill Groomers recruitment system.

Built using **Python**, **Streamlit**, **Google Gemini API**,
**PyPDF2**, **OpenPyXL**, and a lightweight **Chrome Extension**.

------------------------------------------------------------------------

### 🌐 **[Click Here to Try the Live Web App](https://cvanalyzer7.streamlit.app/)**

[![Streamlit
App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cvanalyzer7.streamlit.app/)

------------------------------------------------------------------------

## ✨ Features

-   📄 Upload and analyze multiple PDF resumes
-   🤖 Factual CV extraction using Google Gemini
-   📑 Direct PDF analysis for complex resume layouts
-   👤 Candidate contact and personal information extraction
-   💼 Complete employment-history extraction
-   🎓 Highest completed qualification extraction
-   🧠 Full factual CV skills and certifications
-   📊 Deterministic total-experience calculation
-   🏢 Distinct employer count
-   🔄 Job-change calculation
-   ⏳ Average-tenure calculation
-   🔗 Skill Groomers master-data mapping
-   🧩 Core Role, Functional Area, Role and Industry suggestions
-   📍 Current and native location mapping
-   🗺 Missing-state inference from supported cities where appropriate
-   🛠 Deterministic Skill Groomers Key Skills mapping
-   👥 HR review and editing before approval
-   ⚠ Mapping conflict/unresolved-field warnings
-   📋 HR Excel template generation
-   📥 Individual candidate Excel downloads
-   🌐 Skill Groomers candidate-form autofill through a Chrome extension
-   🔐 HR-controlled final Save
-   🛠 Optional Developer Debug mode

------------------------------------------------------------------------

# 🛠 Technologies Used

-   Python
-   Streamlit
-   Google Gemini API / Google GenAI Python SDK
-   PyPDF2
-   OpenPyXL
-   JSON
-   Skill Groomers Master Data
-   Chrome Extension
-   JavaScript

------------------------------------------------------------------------

# 🤖 AI Model

Google Gemini is used primarily for understanding the resume and
extracting factual structured information.

It extracts information such as candidate name, email, phone numbers,
DOB, gender, locations, employment history, designations, employment
dates, education, skills, certifications and salary when explicitly
stated.

The AI **does not calculate experience-related HR metrics**.

Python performs deterministic calculations and business rules including:

-   Total Experience
-   Number of Employers
-   Number of Job Changes
-   Average Tenure
-   Current and previous distinct employment selection
-   Skill Groomers mapping
-   Excel report generation

``` text
PDF Resume
    │
    ▼
Google Gemini
    │
    ▼
Factual Structured Data
    │
    ▼
Python Validation & Normalization
    │
    ▼
Deterministic Business Logic
    │
    ▼
Verified Candidate Data
```

------------------------------------------------------------------------

# 🗂 Skill Groomers Master Data

The application uses:

``` text
skillgroomers_master_data_v6.json
```

as its local Skill Groomers master-data source.

It contains the supported website values used by the mapping system,
including:

-   Core Roles
-   Key Skills / Keywords
-   Industries
-   Locations
-   Functional Areas
-   Sub Functional Areas
-   Roles
-   Education Qualifications
-   Courses
-   Specializations

The master data preserves Skill Groomers website labels so factual CV
information can be translated into classifications supported by the
recruitment system.

``` text
CV Information
      │
      ▼
Python Mapping Rules
      │
      ▼
Skill Groomers Master JSON
      │
      ▼
Supported SG Classification
      │
      ▼
HR Review
```

Factual CV information remains separate from the Skill Groomers
representation.

------------------------------------------------------------------------

# 🛠 Key Skills Mapping

The application maintains two separate skill representations.

### CV Key Skills

All genuine factual skills extracted from the candidate's resume can be
displayed to HR.

### Skill Groomers Key Skills

Python scans the CV for supported Skill Groomers keywords and aliases,
ranks evidenced matches deterministically, applies skill-family
deduplication, and selects a maximum of **4 Skill Groomers Key Skills**.

If fewer than four supported skills are clearly evidenced, the
application can return fewer rather than inventing another skill.

------------------------------------------------------------------------

# 📍 Location Mapping

Candidate locations are mapped against supported Skill Groomers
locations.

If an exact and unambiguous supported city is present but the state is
missing, Python can infer the state for the **Skill Groomers mapping
only**.

``` text
CV: Mumbai / State not provided
          ↓
SG: Mumbai / Maharashtra
```

The inferred state does not overwrite factual CV data. Conflicting
city/state information can be flagged for HR review.

------------------------------------------------------------------------

# 👤 Using the Application --- HR / Normal Users

This section is for HR team members and other users who only need to use
the deployed application. No Python setup or API configuration is
required.

## 🌐 1. Open the CV Analyzer

Open:

**[CV Analyzer](https://cvanalyzer7.streamlit.app/)**

## 📄 2. Upload CVs

Upload one or more PDF resumes and click:

``` text
🔍 Analyse CVs
```

## 🤖 3. Wait for Analysis

The application extracts candidate facts and employment history,
calculates HR metrics, extracts education and skills, and creates Skill
Groomers classification suggestions.

## 👤 4. Review Candidate Information

Verify:

-   Candidate and contact details
-   Current and previous employment
-   Education
-   Total Experience
-   Number of Employers
-   Number of Job Changes
-   Average Tenure
-   CV Key Skills

## 🧩 5. Review Skill Groomers Classification

Review the suggested:

-   Core Role
-   Functional Area
-   Role
-   Industry
-   Current City / State
-   Native City / State
-   Education
-   Skill Groomers Key Skills

Correct uncertain or conflicting values when required.

## ⚠️ 6. Check Review Warnings

Resolve any mapping conflicts or unresolved fields before approval. HR
should still verify important candidate information even when no
automatic conflict is detected.

## ✅ 7. Approve the Candidate

Approve the candidate after reviewing the factual information and Skill
Groomers classifications.

## 📥 8. Download the Excel Report

Use **Download Excel Report** to download the populated HR report.

## 🌐 9. Open Skill Groomers

Click:

``` text
Open Skill Groomers & Fill Candidate
```

The Skill Groomers Add Candidate page opens and the internally installed
Chrome extension fills supported fields.

## 🔎 10. Review the Skill Groomers Form

Verify the populated candidate details, classifications, locations,
experience, employment and education. Complete any intentionally manual
or unresolved fields.

## 💾 11. Save

Use Skill Groomers's own **Save** button after verification.

The Analyzer and extension do **not** automatically perform the final
Save.

### Normal User Workflow

``` text
Open CV Analyzer
      ↓
Upload PDF CV
      ↓
Analyse CV
      ↓
Review CV Facts
      ↓
Review SG Classifications
      ↓
Approve Candidate
   ↙       ↘
Excel    Open Skill Groomers
              ↓
         Form Autofilled
              ↓
          HR Reviews
              ↓
         HR Clicks Save
```

------------------------------------------------------------------------

# 👨‍💻 Developer Setup & Usage

This section is for developers who need to install, run, maintain, debug
or modify the application.

## 📂 Project Structure

``` text
CV-Analyzer/
│
├── app.py
├── resume1.xlsx
├── skillgroomers_master_data_v6.json
├── requirements.txt
├── README.md
├── .gitignore
│
├── skillgroomers_autofill_extension/
│   ├── manifest.json
│   ├── content.js
│   └── README.txt
│
└── .streamlit/
    └── secrets.toml
```

### Main Components

**`app.py`** --- Main Streamlit application.

**`resume1.xlsx`** --- Excel template populated for analyzed candidates.

**`skillgroomers_master_data_v6.json`** --- Local Skill Groomers
master-data snapshot required by the mapping system.

**`skillgroomers_autofill_extension/`** --- Browser-side extension for
filling approved candidate information into Skill Groomers.

**`.streamlit/secrets.toml`** --- Local Gemini API credentials. This
file must **NOT** be committed to GitHub.

## 📦 Requirements

Tested dependencies:

``` txt
streamlit==1.60.0
google-genai==2.14.0
PyPDF2==3.0.1
openpyxl==3.1.5
```

Install them with:

``` bash
pip install -r requirements.txt
```

## 🚀 1. Clone the Repository

``` bash
git clone https://github.com/YOUR_USERNAME/CV-Analyzer.git
cd CV-Analyzer
```

## 🔑 2. Configure Gemini API Key

Create:

``` text
.streamlit/secrets.toml
```

Add:

``` toml
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

Do not hardcode the real key in `app.py`. Keep `.streamlit/` excluded
through `.gitignore`.

For Streamlit Community Cloud, add the same secret through **App
Settings → Secrets**.

## ▶ 3. Run Locally

``` bash
streamlit run app.py
```

## 🗂 4. Master Data

Keep `skillgroomers_master_data_v6.json` with the application. When
updating it, preserve the exact Skill Groomers labels and IDs expected
by the website.

## 🧩 5. Install the Chrome Extension

Extract:

``` text
skillgroomers_autofill_easy_reliable.zip
```

Keep the extracted folder in a permanent location. The selected folder
must directly contain:

``` text
manifest.json
content.js
README.txt
```

Then:

1.  Open `chrome://extensions`.
2.  Enable **Developer mode**.
3.  Click **Load unpacked**.
4.  Select the folder containing `manifest.json`.

## 🔄 6. Update the Extension

1.  Replace changed files in the existing extension folder.
2.  Open `chrome://extensions`.
3.  Find the Skill Groomers Autofill extension.
4.  Click **Reload**.

The extension normally does not need to be uninstalled.

------------------------------------------------------------------------

# 🌐 Skill Groomers Integration

``` text
CV Analyzer
    ↓
HR Approves Candidate
    ↓
Open Skill Groomers & Fill Candidate
    ↓
Chrome Extension
    ↓
Candidate Form Autofilled
    ↓
HR Reviews
    ↓
HR Clicks Save
```

The extension fills supported fields and selects matching Skill Groomers
options. It leaves unresolved options for HR review.

It does **not**:

-   Log into Skill Groomers
-   Store usernames or passwords
-   Read or store authentication cookies
-   Bypass CAPTCHA
-   Automatically click Save
-   Automatically submit candidates

Final production submission remains under HR control.

------------------------------------------------------------------------

# 📊 Business Rules

### 💼 Total Experience

-   Employment months are counted inclusively.
-   Overlapping calendar months count once.
-   Employment gaps do not count.
-   Present employment is calculated through the current calendar month.

### 🏢 Number of Employers

The same employer counts once even when the candidate held multiple
roles there.

### 🔄 Number of Job Changes

A job change is counted when the candidate moves between distinct
employers.

### ⏳ Average Tenure

``` text
Average Tenure = Total Experience ÷ Number of Job Changes
```

### 💼 Current and Previous Employment

The most recent position is current employment. Previous employment
means the previous **distinct employer**.

### 🎓 Education

The highest **completed** qualification is used. A pursuing
qualification does not replace it.

### ❓ Missing Information

Missing factual information is left blank rather than guessed.

------------------------------------------------------------------------

# 📋 Excel Report

The application populates the HR Excel template with factual candidate
information and calculated HR metrics. Each analyzed candidate can be
downloaded as an individual Excel report.

------------------------------------------------------------------------

# 🖥 User Interface

The application provides:

-   Multiple PDF upload
-   Candidate analysis progress
-   Candidate summary cards
-   Factual candidate information
-   Full factual CV Key Skills
-   Calculated HR metrics
-   Editable Skill Groomers classifications
-   Mapping warnings
-   Candidate approval
-   Individual Excel downloads
-   Skill Groomers autofill
-   Optional Developer Debug tools

------------------------------------------------------------------------

# 🛠 Developer Debug Mode

Developer Debug mode can show:

-   CV / Excel facts
-   Raw Gemini JSON
-   Normalized factual data
-   Skill Groomers mapping summary
-   Full mapping details
-   Candidate payload preview

It is intended for development and troubleshooting rather than normal HR
use.

------------------------------------------------------------------------

# 🔐 Security

-   Gemini API keys are not hardcoded
-   `.streamlit/secrets.toml` is excluded from Git
-   Environment files should be excluded from Git
-   Skill Groomers passwords are not stored by the Analyzer
-   The extension does not perform authentication
-   Authentication cookies are not stored by the extension
-   CAPTCHA is not bypassed
-   HR reviews candidate information before submission
-   Final Skill Groomers Save remains under HR control

------------------------------------------------------------------------

# 🎯 Design Principle

``` text
AI understands.
Python calculates.
Master data classifies.
HR verifies.
Skill Groomers saves.
```

------------------------------------------------------------------------

# 👨‍💻 Author

**Jai Pandya**

CV Analyzer was developed as a Python recruitment automation project
combining AI-powered resume understanding with deterministic Python
calculations, Excel automation, Skill Groomers master-data
classification, an HR review workflow, and browser-assisted candidate
form filling.

------------------------------------------------------------------------

# 📄 License

This project is intended for internal, educational, learning, and
demonstration purposes.
