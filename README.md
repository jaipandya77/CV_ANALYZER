# CV Analyzer

An AI-powered web application that analyzes PDF resumes and automatically fills a predefined HR Excel template.

Built using **Python**, **Streamlit**, **Google Gemini AI (Free API)**, and **OpenPyXL**, the application extracts structured candidate information from resumes and performs HR-specific calculations before generating a ready-to-use Excel report.

---

## Features

- Upload resumes in PDF format
- AI-powered resume information extraction
- Automatic HR Excel template population
- Calculates:
  - Candidate Name
  - Highest Qualification
  - Educational Institute
  - NIRF Institute Ranking
  - Total Work Experience
  - Experience as:
    - Contractor
    - Owner
    - Consultant
    - Freelancer
  - Experience in MNCs
  - Experience in Listed Companies
  - Experience in India
  - Experience Outside India
  - Number of Job Changes
  - Average Tenure
  - Employment Gaps (reported only if greater than 1 year)
  - Certifications
- Download the completed Excel report

---

# Technologies Used

- Python
- Streamlit
- Google Gemini AI
- Google Gemini Free API
- PyPDF2
- OpenPyXL

---

# AI Model

This project uses **Google Gemini's free-tier API** for resume information extraction.

Instead of relying on paid AI services or locally hosted LLMs, the application uses Google's **Gemini Flash** model through the official Google GenAI Python SDK.

### Advantages

- Free API tier for development
- Fast inference
- Structured JSON extraction
- No model training required
- Easy integration with Python

---

# Project Structure

```
CV-Analyzer/
│
├── app.py
├── HR_Template.xlsx
├── nirf_rankings.csv
├── requirements.txt
├── README.md
├── .gitignore
└── .env
```

---

# Installation

Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/CV-Analyzer.git
```

Move into the project

```bash
cd CV-Analyzer
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

# Configure API Key

Create a `.env` file in the project root.

```
GEMINI_API_KEY=YOUR_API_KEY
```

The application securely loads the API key from environment variables.

---

# Run the Application

```bash
streamlit run app.py
```

The application will automatically open in your browser.

---

# How It Works

```
          PDF Resume
               │
               ▼
      Extract Text (PyPDF2)
               │
               ▼
     Google Gemini Flash API
               │
               ▼
      Structured JSON Output
               │
               ▼
      Python Business Logic
               │
               ▼
   Experience & Gap Calculation
               │
               ▼
 Populate HR Excel Template
               │
               ▼
      Download Excel Report
```

---

# Business Rules

The application follows these HR-specific rules:

- Employment gaps are reported only if they exceed **1 year**.
- Average tenure is calculated as:

```
Total Experience ÷ Number of Job Changes
```

- Educational institute ranking is determined using the included NIRF rankings dataset.
- All calculations are performed in Python after AI extraction.

---

# Why Gemini AI?

The application uses **Gemini AI** only for extracting structured information from resumes.

All calculations—including:

- Experience calculation
- Employment gap detection
- Job changes
- Average tenure
- NIRF ranking lookup
- Excel generation

are performed entirely in Python, ensuring consistent and deterministic results.

---

# Future Improvements

- Batch resume processing
- OCR support for scanned PDFs
- Resume scoring
- Candidate ranking dashboard
- Database integration
- Support for DOCX resumes
- Export results to CSV

---

# Author

**Jai Pandya**

A Python-based automation project developed to simplify HR resume screening by combining AI-powered information extraction with automated Excel report generation.

---

# License

This project is intended for educational, demonstration, and learning purposes.
