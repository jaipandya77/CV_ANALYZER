# CV Analyzer V8.5

> **AI-assisted CV analysis, deterministic HR calculations, Skill
> Groomers mapping, and browser-assisted candidate entry --- with HR
> retaining final approval.**

[![Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cvanalyzer7.streamlit.app/)

**Live Application:** https://cvanalyzer7.streamlit.app/

------------------------------------------------------------------------

## Overview

CV Analyzer V8.5 is a recruitment-support application built to reduce
repetitive CV review and candidate-entry work.

The application reads PDF resumes, extracts factual candidate
information with Google Gemini, performs HR calculations in Python, maps
supported values to Skill Groomers master data, generates an Excel
report, and helps transfer approved information into the existing Skill
Groomers candidate form.

The system deliberately separates **CV facts**, **calculated values**,
and **Skill Groomers classifications** so that mapped website values do
not overwrite factual resume information.

### Core workflow

``` text
PDF Resume
    |
    v
Gemini Factual Extraction
    |
    v
Python Validation & Calculations
    |
    +----------------------+
    |                      |
    v                      v
Excel Report       Skill Groomers Mapping
                           |
                           v
                       HR Review
                           |
                           v
                  Chrome Extension Autofill
                           |
                           v
                    HR Final Verification
                           |
                           v
                  Skill Groomers Save
```

------------------------------------------------------------------------

## Key Features

  -----------------------------------------------------------------------
  Area                                Capability
  ----------------------------------- -----------------------------------
  Resume analysis                     Multiple PDF CV uploads and direct
                                      PDF analysis

  AI extraction                       Candidate details, employment,
                                      education, skills and
                                      certifications

  Experience                          Deterministic total-experience
                                      calculation with overlap handling

  Date precision                      Year-only employment periods are
                                      treated as approximate rather than
                                      fabricated as exact months

  Employment                          Distinct employer count, job
                                      changes, current and previous
                                      employer

  Average tenure                      Deterministic calculation using the
                                      configured HR business rule

  Skills                              Full factual CV skills plus
                                      separate Skill Groomers Key Skills

  Classification                      Core Role, Functional Area, Role
                                      and Industry mapping

  Functional Area                     Designation-first mapping so career
                                      track takes priority over
                                      supporting keywords

  Location                            Skill Groomers city/state mapping
                                      with controlled state inference

  Review                              Editable classifications, warnings
                                      and unresolved-field checks

  Reporting                           Individual Excel reports using the
                                      HR template

  Integration                         Chrome extension fills supported
                                      Skill Groomers candidate fields

  Control                             HR reviews and performs the final
                                      Skill Groomers Save
  -----------------------------------------------------------------------

------------------------------------------------------------------------

# HR / Normal User Guide

This section is for HR team members using the deployed application.

**You do not need Python, an API key, VS Code, or developer tools to use
the deployed CV Analyzer.**

## 1. One-Time Chrome Extension Setup

Before using **Open Skill Groomers & Fill Candidate** on a computer for
the first time, install the internal Skill Groomers Autofill extension.

This setup is required **once per Chrome profile**.

### Install

1.  Obtain `skillgroomers_autofill_easy_reliable.zip`.
2.  Extract the ZIP to a **permanent folder** on the computer.
3.  Open Chrome.
4.  Enter `chrome://extensions` in the address bar.
5.  Turn on **Developer mode**.
6.  Click **Load unpacked**.
7.  Select the extracted folder that directly contains `manifest.json`.
8.  Confirm that **Skill Groomers Autofill** appears and is enabled.

The selected folder should look like:

``` text
skillgroomers_autofill_extension/
├── manifest.json
├── content.js
└── README.txt
```

> **Important:** Do not move or delete this folder after installation.
> Chrome needs the unpacked extension files to remain at that location.

After installation, HR does **not** need to manually operate the
extension. The CV Analyzer sends approved candidate information to it
when **Open Skill Groomers & Fill Candidate** is used.

The extension does not perform the final Save.

------------------------------------------------------------------------

## 2. Open the CV Analyzer

Open the deployed application:

**https://cvanalyzer7.streamlit.app/**

------------------------------------------------------------------------

## 3. Upload CVs

Upload one or more PDF resumes and click:

``` text
Analyse CVs
```

The application analyzes each candidate separately.

------------------------------------------------------------------------

## 4. Review Candidate Facts

Verify the information extracted from the CV, especially:

-   Candidate name and contact information
-   Current and previous employment
-   Designations and employers
-   Employment dates
-   Education
-   Total Experience
-   Number of Employers
-   Number of Job Changes
-   Average Tenure
-   CV Key Skills

Missing factual information should remain blank rather than being
guessed.

### Date-precision warning

Some CVs provide employment dates such as:

``` text
2017 - 2020
```

without months.

In these cases, an exact month-level duration cannot be proven from the
CV. The application can therefore treat experience as **approximate / a
range** instead of silently inventing January or December dates.

CVs containing proper month-level dates can continue to receive exact
month-based calculations.

------------------------------------------------------------------------

## 5. Review Skill Groomers Classification

The application separately maps the candidate to supported Skill
Groomers values.

Review:

-   Core Role
-   Functional Area / Sub Functional Area / Role
-   Industry
-   Skill Groomers Key Skills
-   Current City / State
-   Native City / State
-   Education mapping

### Functional Area mapping

Functional Area mapping prioritizes the candidate's **actual career
track**:

``` text
Current / Recent Designations
          |
          v
Designation History
          |
          v
Professional Domain
          |
          v
Supporting Skills
```

For example, a Project Manager working in Civil / Interior Fit-Out
should not become a Safety professional merely because the CV mentions
safety compliance.

------------------------------------------------------------------------

## 6. Check Warnings

If the application identifies a conflict, ambiguous mapping, missing
required value, or insufficient date precision, review it before
approval.

A clean result may show:

``` text
No mapping conflicts or unresolved fields detected.
```

This means no automatic mapping issue was found; HR should still verify
important candidate information.

------------------------------------------------------------------------

## 7. Approve Candidate

After reviewing the factual information and Skill Groomers
classifications, approve the candidate in the CV Analyzer.

------------------------------------------------------------------------

## 8. Download Excel Report

Use **Download Excel Report** when an HR Excel copy is required.

The Excel output keeps factual CV information separate from Skill
Groomers-specific mapped classifications.

------------------------------------------------------------------------

## 9. Open Skill Groomers & Autofill

Click:

``` text
Open Skill Groomers & Fill Candidate
```

The Skill Groomers Add Candidate page opens.

The installed Chrome extension then fills the supported candidate fields
using the **reviewed information** from the Analyzer.

------------------------------------------------------------------------

## 10. Final HR Review & Save

Check the populated Skill Groomers form.

Complete any intentionally manual or unresolved fields, then use Skill
Groomers's own **Save** button.

``` text
Analyzer prepares
       |
       v
Extension fills
       |
       v
HR verifies
       |
       v
Skill Groomers saves
```

**The Analyzer and Chrome extension do not automatically submit the
candidate.**

------------------------------------------------------------------------

# Developer Guide

## Technology Stack

  Component                          Purpose
  ---------------------------------- ------------------------------------------------
  Python                             Validation, calculations and application logic
  Streamlit                          Web application UI
  Google Gemini / Google GenAI SDK   Factual resume understanding
  PyPDF2                             PDF text extraction
  OpenPyXL                           Excel template population
  JSON                               Skill Groomers master-data storage
  JavaScript                         Chrome extension autofill bridge

### Tested dependencies

``` txt
streamlit==1.60.0
google-genai==2.14.0
PyPDF2==3.0.1
openpyxl==3.1.5
```

------------------------------------------------------------------------

## Project Structure

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

### Important files

**`app.py`**\
Main Streamlit application.

**`resume1.xlsx`**\
HR Excel template populated by the application.

**`skillgroomers_master_data_v6.json`**\
Local snapshot of supported Skill Groomers classifications and IDs.

**`skillgroomers_autofill_extension/`**\
Browser-side bridge used to fill reviewed candidate information into
Skill Groomers.

**`.streamlit/secrets.toml`**\
Local secret configuration. Do not commit this file.

------------------------------------------------------------------------

## Local Setup

### 1. Install dependencies

``` bash
pip install -r requirements.txt
```

### 2. Configure Gemini

Create:

``` text
.streamlit/secrets.toml
```

Add:

``` toml
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

Never hardcode the real API key in `app.py`.

### 3. Run the application

``` bash
streamlit run app.py
```

If the Python filename contains parentheses on PowerShell, quote it:

``` powershell
streamlit run "app(3).py"
```

Using `app.py` as the main filename is recommended.

------------------------------------------------------------------------

## Chrome Extension --- Developer Setup

Install the unpacked extension through:

``` text
Chrome
  -> chrome://extensions
  -> Developer mode
  -> Load unpacked
  -> folder containing manifest.json
```

### Updating the extension

When extension files change:

1.  Replace the changed files in the existing permanent extension
    folder.
2.  Open `chrome://extensions`.
3.  Find **Skill Groomers Autofill**.
4.  Click **Reload**.

Normally, the extension does not need to be uninstalled and installed
again.

------------------------------------------------------------------------

# System Design

## Factual Extraction vs Classification

The application maintains two separate representations.

### Factual CV record

Contains information supported by the candidate's resume.

Examples:

-   Employer names
-   Designations
-   Employment dates
-   Education
-   Locations
-   Skills
-   Certifications

### Skill Groomers representation

Translates supported facts into website classifications such as:

-   Core Role
-   Functional Area
-   Role
-   Industry
-   Key Skills
-   Location
-   Education

``` text
CV Facts
   |
   +-----------> Excel / HR Record
   |
   v
Deterministic Mapping
   |
   v
Skill Groomers Master Data
   |
   v
HR Review
```

**Mapping must not overwrite or contaminate the factual CV record.**

------------------------------------------------------------------------

## Experience Calculation

For employment dates with sufficient precision:

-   Employment calendar months are counted inclusively.
-   Overlapping months count once.
-   Gaps are excluded.
-   Current employment runs through the current calendar month.

This uses the **union of covered employment months**, preventing
concurrent roles from being double-counted.

### Year-only employment

A CV may state:

``` text
2008 - 2011
```

but that does not reveal whether employment began in January, June,
December, or another month.

The application should therefore preserve that uncertainty instead of
manufacturing missing months.

------------------------------------------------------------------------

## Number of Employers

The same employer counts once even when the candidate held multiple
positions within that organization.

------------------------------------------------------------------------

## Number of Job Changes

A job change is counted when the candidate moves between distinct
employers.

------------------------------------------------------------------------

## Average Tenure

The configured HR rule is:

``` text
Average Tenure = Total Experience / Number of Job Changes
```

When Total Experience is approximate because source dates lack month
precision, Average Tenure should also be treated as approximate.

------------------------------------------------------------------------

## Current & Previous Employment

-   **Current Employment:** most recent/current position.
-   **Previous Employment:** previous position at a different employer.

------------------------------------------------------------------------

## Education

The highest **completed** qualification is used.

A qualification that is still being pursued does not replace the highest
completed qualification.

------------------------------------------------------------------------

# Skill Groomers Mapping

## Key Skills

The system intentionally keeps two skill sets.

### CV Key Skills

Factual skills extracted from the resume are displayed to HR without
applying Skill Groomers's four-keyword limit.

### Skill Groomers Key Skills

Python scans the resume for supported Skill Groomers keywords and
aliases, ranks evidenced matches deterministically, removes redundant
skill-family matches, and selects up to **4** supported keywords.

If only two or three supported keywords are clearly evidenced, the
system can return fewer rather than inventing another skill.

------------------------------------------------------------------------

## Functional Area

Functional Area is **designation-first**.

Strong designation evidence takes priority over generic supporting words
in the CV.

For example:

``` text
Project Manager / Fit-Out / Civil / Site
              |
              v
Project Management / Site Engineering
```

A Safety classification should require genuine Safety/HSE/EHS
designation evidence rather than merely finding the word `safety` in
responsibilities.

The same principle applies to other specialist classifications such as
QA/QC.

------------------------------------------------------------------------

## Location

Locations are matched against supported Skill Groomers values.

When an exact, unambiguous supported city is present but the state is
missing, the system can infer the corresponding state for the **Skill
Groomers representation only**.

Example:

``` text
CV:
City  = Mumbai
State = not provided

SG mapping:
City  = Mumbai
State = Maharashtra
```

The inferred state does not rewrite the factual CV state.

Conflicting city/state information should be flagged for HR review
rather than silently overwritten.

------------------------------------------------------------------------

# Skill Groomers Integration

The integration is intentionally human-controlled.

``` text
CV Analyzer
     |
     v
HR Approves Candidate
     |
     v
Open Skill Groomers & Fill Candidate
     |
     v
Chrome Extension
     |
     v
Supported Fields Autofilled
     |
     v
HR Reviews
     |
     v
HR Clicks Save
```

The extension **does not**:

-   Log into Skill Groomers
-   Store Skill Groomers usernames or passwords
-   Read or store authentication cookies
-   Bypass CAPTCHA
-   Automatically click Save
-   Automatically submit candidates

Final production submission remains under HR control.

------------------------------------------------------------------------

# Security

-   Gemini API keys are stored outside source code.
-   `.streamlit/secrets.toml` should remain excluded from Git.
-   `.env` files should remain excluded from Git.
-   Skill Groomers credentials are not stored by the Analyzer.
-   The browser extension does not perform authentication.
-   CAPTCHA is not bypassed.
-   Candidate information is reviewed before final submission.
-   Skill Groomers retains control of the final Save.

Recommended `.gitignore` entries:

``` gitignore
# Python
__pycache__/
*.pyc

# Secrets
.env
.env.*
.streamlit/

# Virtual environments
.venv/
venv/
env/

# Temporary Excel files
~$*.xlsx

# Editor / OS
.vscode/
.idea/
.DS_Store
Thumbs.db
```

------------------------------------------------------------------------

# Troubleshooting

### `Open Skill Groomers & Fill Candidate` opens the page but does not fill

Check that:

1.  The Chrome extension is installed.
2.  It is enabled in `chrome://extensions`.
3.  The permanent extension folder has not been moved or deleted.
4.  The latest extension files are installed.
5.  After an extension update, **Reload** was clicked in
    `chrome://extensions`.

### Chrome says `Manifest file is missing or unreadable`

You probably selected the outer extracted folder.

Select the folder that **directly contains**:

``` text
manifest.json
content.js
README.txt
```

### Streamlit says the Python file has no extension

When a PowerShell filename contains parentheses, use quotes:

``` powershell
streamlit run "app(3).py"
```

Prefer renaming the main application to:

``` text
app.py
```

and run:

``` powershell
streamlit run app.py
```

------------------------------------------------------------------------

# Design Principle

> **AI understands the resume. Python calculates the metrics. Master
> data classifies supported values. HR verifies the candidate. Skill
> Groomers performs the final Save.**

------------------------------------------------------------------------

## Author

**Jai Pandya**

CV Analyzer V8.5 is a Python recruitment-automation project combining
AI-assisted resume understanding, deterministic HR calculations, Excel
automation, Skill Groomers master-data classification, human review, and
browser-assisted candidate-form filling.

------------------------------------------------------------------------

## License

This project is intended for internal, educational, learning, and
demonstration purposes.
