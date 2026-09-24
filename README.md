# CV Analyzer V9.0

> **AI-assisted CV analysis, deterministic calculations, Skill Groomers mapping, and browser-assisted candidate entry — featuring a two-step React-safe autofill and premium glassmorphism UI.**

[![Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cvanalyzer7.streamlit.app/)

**Live Application:** https://cvanalyzer7.streamlit.app/

---

## Overview

CV Analyzer V9.0 is a recruitment-support application built to reduce repetitive CV review and candidate-entry work. 

The application reads PDF resumes, extracts factual candidate information with Google Gemini, performs calculations in Python, maps supported values to Skill Groomers master data, generates an Excel report, and helps transfer approved information into the existing Skill Groomers candidate form.

The system deliberately separates **CV facts**, **calculated values**, and **Skill Groomers classifications** so that mapped website values do not overwrite factual resume information.

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
                      User Review
                           |
                           v
            2-Step Chrome Extension Autofill
                 (Main Form + DOB Calendar)
                           |
                           v
                   Final User Verification
                           |
                           v
                  Skill Groomers Save
```

---

## Key Features

| Area | Capability |
| :--- | :--- |
| **UI/UX Design** | Executive "Midnight" theme featuring glassmorphism elements, backdrop filters, and refined data cards. |
| **Resume analysis** | Multiple PDF CV uploads and direct PDF analysis. |
| **AI extraction** | Candidate details, employment, education, skills and certifications. |
| **Experience** | Deterministic total-experience calculation using absolute month-indexing to prevent off-by-one errors and handle overlaps. |
| **Date precision** | Year-only employment periods are treated as approximate rather than fabricated as exact months. |
| **Employment** | Distinct employer count, job changes, current and previous employer. |
| **Average tenure** | Deterministic calculation using the configured business rule. |
| **Skills** | Full factual CV skills plus separate Skill Groomers Key Skills. |
| **Classification** | Core Role, Functional Area, Role and Industry mapping. |
| **Functional Area** | Designation-first mapping so career track takes priority over supporting keywords. |
| **Reporting** | Individual Excel reports using the provided template. |
| **Integration** | Custom Chrome extension with Main World script injection (`pageScript.js`) to safely bypass Content Security Policy (CSP) and populate complex React calendar widgets without crashing the portal. |
| **Control** | The user reviews and performs the final Skill Groomers Save. |

---

# User Guide

This section is for users of the deployed application.

**You do not need Python, an API key, VS Code, or developer tools to use the deployed CV Analyzer.**

## 1. One-Time Chrome Extension Setup

Before using the autofill buttons on a computer for the first time, install the internal Skill Groomers Autofill extension. 

This setup is required **once per Chrome profile**.

### Install

1. Obtain `skillgroomers_autofill_easy_reliable.zip`.
2. Extract the ZIP to a **permanent folder** on the computer.
3. Open Chrome.
4. Enter `chrome://extensions` in the address bar.
5. Turn on **Developer mode**.
6. Click **Load unpacked**.
7. Select the extracted folder that directly contains `manifest.json`.
8. Confirm that **Skill Groomers Autofill** appears and is enabled.

The selected folder should look like:
``` text
skillgroomers_autofill_extension/
├── manifest.json
├── content.js
├── pageScript.js
└── README.txt
```

> **Important:** Do not move or delete this folder after installation. Chrome needs the unpacked extension files to remain at that location.

After installation, the user does **not** need to manually operate the extension. The CV Analyzer sends approved candidate information to it automatically.

---

## 2. Open the CV Analyzer

Open the deployed application:

**https://cvanalyzer7.streamlit.app/**

---

## 3. Upload CVs

Upload one or more PDF resumes and click:

``` text
Analyse Resumes
```

The application analyzes each candidate separately.

---

## 4. Review Candidate Facts

Verify the information extracted from the CV, especially:

* Candidate name and contact information
* Current and previous employment
* Designations and employers
* Employment dates
* Education
* Total Experience
* Number of Employers
* Number of Job Changes
* Average Tenure
* CV Key Skills

Missing factual information should remain blank rather than being guessed.

---

## 5. Review Skill Groomers Classification

The application separately maps the candidate to supported Skill Groomers values. Review:

* Core Role
* Functional Area / Sub Functional Area / Role
* Industry
* Skill Groomers Key Skills
* Current City / State
* Native City / State
* Education mapping

### Functional Area mapping

Functional Area mapping prioritizes the candidate's **actual career track**:

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

---

## 6. Check Warnings

If the application identifies a conflict, ambiguous mapping, or missing required value, review it and apply a manual override using the dropdown menus before approval.

---

## 7. Approve Candidate

After reviewing the factual information and Skill Groomers classifications, click the **Approve Candidate** button to unlock the export and autofill options.

---

## 8. Download Excel Report

Use **Download Excel Report** when an Excel copy is required. 

The Excel output keeps factual CV information separate from Skill Groomers-specific mapped classifications.

---

## 9. Two-Step Autofill

Because of complex React calendar widgets on the portal, candidate transfer is split into a safe two-step process to prevent portal crashes. 

1. Click **1. Autofill Add Candidate form→**. This opens the Skill Groomers Add Candidate page in a new unique tab and injects all text, dropdowns, and classifications.
2. Return to the Streamlit app and click **2. Fill DOB Only 📅**. This communicates with the *exact same tab* you just opened and safely injects the calendar dates directly into the React Fiber nodes.

---

## 10. Final Review & Save

Check the populated Skill Groomers form.

Complete any intentionally manual or unresolved fields, then use Skill Groomers's own **Save** button.

**The Analyzer and Chrome extension do not automatically submit the candidate.**

---

# Developer Guide

## Technology Stack

| Component | Purpose |
| :--- | :--- |
| Python | Validation, calculations and application logic |
| Streamlit | Web application UI |
| Google Gemini / GenAI SDK | Factual resume understanding |
| PyPDF2 | PDF text extraction |
| OpenPyXL | Excel template population |
| JSON | Skill Groomers master-data storage |
| JavaScript | Chrome extension autofill bridge with Isolated & Main World scripts |

### Tested dependencies

``` txt
streamlit
google-genai
PyPDF2
openpyxl
psutil
```

---

## Project Structure

``` text
CV-Analyzer/
│
├── appy.py
├── resume1.xlsx
├── skillgroomers_master_data_v7.json
├── requirements.txt
├── README.md
├── .gitignore
│
├── skillgroomers_autofill_extension/
│   ├── manifest.json
│   ├── content.js
│   ├── pageScript.js
│   └── README.txt
│
└── .streamlit/
    └── secrets.toml
```

### Important files

**`appy.py`**
Main Streamlit application.

**`resume1.xlsx`**
Excel template populated by the application.

**`skillgroomers_master_data_v7.json`**
Local snapshot of supported Skill Groomers classifications and IDs.

**`skillgroomers_autofill_extension/`**
Browser-side bridge used to fill reviewed candidate information into Skill Groomers. Utilizes `pageScript.js` to bypass CSP restrictions and directly manipulate Moment.js objects inside the React application's memory.

**`.streamlit/secrets.toml`**
Local secret configuration. Do not commit this file.

---

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
Never hardcode the real API key in `appy.py`.

### 3. Run the application
``` bash
streamlit run appy.py
```

---

## Chrome Extension — Developer Setup

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
1. Replace the changed files in the existing permanent extension folder.
2. Open `chrome://extensions`.
3. Find **Skill Groomers Autofill**.
4. Click **Reload**.

---

# System Design

## Experience Calculation

To bypass traditional "fencepost" (off-by-one) errors, the app avoids direct date subtraction. Instead, it relies on absolute month indexing: `(year * 12) + (month - 1)`. 

For employment dates with sufficient precision:
* Employment calendar months are counted inclusively.
* Overlapping months count once by converting ranges to integer sets.
* Gaps are excluded.
* Current employment runs through the current calendar month.

This uses the **union of covered employment months**, preventing concurrent roles from being double-counted.

---

## Security

* Gemini API keys are stored outside source code.
* `.streamlit/secrets.toml` must remain excluded from Git.
* `.env` files must remain excluded from Git.
* Skill Groomers credentials are not stored by the Analyzer.
* The browser extension does not perform authentication.
* Candidate information is reviewed before final submission.
* Skill Groomers retains control of the final Save.

---

# Design Principle

> **AI understands the resume. Python calculates the metrics. Master data classifies supported values. The user verifies the candidate. Skill Groomers performs the final Save.**

---

## Author

**Jai Pandya**

CV Analyzer V9.0 is a Python recruitment-automation project combining AI-assisted resume understanding, deterministic calculations, Excel automation, Skill Groomers master-data classification, human review, and browser-assisted candidate-form filling.
