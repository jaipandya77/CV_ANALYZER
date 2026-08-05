# 📄 CV Analyzer

An AI-powered web application that extracts information from PDF resumes and automatically generates a structured HR report in Excel format.

Built using **Python**, **Streamlit**, **Google Gemini API (Free Tier)**, and **OpenPyXL**, the application converts unstructured resumes into structured data and performs HR-specific calculations before populating a predefined Excel template.

---
---

### 🌐 **[Click Here to Try the Live Web App](https://cvanalyzer7.streamlit.app/)**

*(Or use this visual badge format)*:
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cvanalyzer7.streamlit.app/)

---

## ✨ Features

- 📄 Upload multiple resumes in PDF format simultaneously
- 🤖 AI-powered information extraction using Google Gemini
- 📋 Automatic HR Excel report generation
- 🎓 Extracts highest qualification
- 🏫 Identifies educational institute
- 📊 Finds NIRF institute ranking
- 💼 Calculates total work experience
- 🏗 Calculates experience as:
  - Contractor
  - Owner
  - Consultant
  - Freelancer
- 🏢 Calculates experience in:
  - MNCs
  - Listed Companies
- 🌍 Calculates:
  - Experience in India
  - Experience outside India (ignores jobs where the country is unknown)
- 🔄 Calculates number of job changes
- ⏳ Calculates average tenure
- ⚠ Detects employment gaps (only when greater than or equal to 1 year)
- 📜 Counts professional certifications
- 📥 Download the completed HR Excel report
- 🛠 Optional Debug Mode for viewing extracted JSON and work experience

---

# 🛠 Technologies Used

- Python
- Streamlit
- Google Gemini API (Free Tier)
- PyPDF2
- OpenPyXL
- CSV (NIRF Rankings Dataset)

---

# 🤖 AI Model

This project uses the **Google Gemini API (Free Tier)** for extracting structured information from resumes.

The AI is responsible **only for information extraction**.

It extracts:

- Candidate Name
- Highest Qualification
- Educational Institute
- Work Experience
- Company Details
- Dates
- Country
- Certifications

The AI **does not perform any calculations**.

All business logic is implemented in Python, including:

- Total experience calculation
- Employment gap detection
- Job changes
- Average tenure
- NIRF overall ranking lookup
- Organization-wise experience
- India/Outside India experience
- MNC & Listed company experience
- Excel report generation

This approach keeps calculations deterministic, accurate, and independent of AI responses.

---

# 📂 Project Structure

```
CV-Analyzer/
│
├── app.py
├── HR_Template.xlsx
├── nirf_rankings.csv
├── requirements.txt
├── README.md
├── .gitignore
└── assets/
```

---

# 🚀 Installation

Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/CV-Analyzer.git
```

Move into the project directory

```bash
cd CV-Analyzer
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

# 🔑 Configure API Key

This application uses **Streamlit Secrets** to securely store the Gemini API key.

For local development, create:

```
.streamlit/secrets.toml
```

Add:

```toml
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

When deploying to **Streamlit Community Cloud**, add the same secret through:

**App Settings → Secrets**

The application securely reads the API key using Streamlit Secrets without exposing it in the source code.

---

# ▶ Running the Application

```bash
streamlit run app.py
```

The application will automatically open in your browser.

---

# 🔄 Application Workflow

```
              One or More PDF Resume
                    │
                    ▼
         Text Extraction (PyPDF2)
                    │
                    ▼
     Google Gemini API (Free Tier)
                    │
                    ▼
       Structured JSON Extraction
                    │
                    ▼
      Python Business Logic Layer
                    │
                    ▼
    HR Calculations & Validations
                    │
                    ▼
     Populate Excel HR Template
                    │
                    ▼
    Generate Individual Excel Reports
                    │
                    ▼
              Download Reports
```

---

# 📊 Business Rules

The application follows these HR-specific business rules:

- Employment gaps are reported only when the gap is **greater than or equal to one year**.
- Average tenure is calculated as:

```
Average Tenure = Total Experience ÷ Job Changes
```

- NIRF institute rankings are retrieved using the included CSV dataset.
- Work experience is calculated from employment dates.
- Jobs with an unknown or missing country are excluded from both India and Outside India experience calculations.
- Job changes are calculated by detecting changes in employer.
- AI performs extraction only.
- Python performs every calculation and validation.

---

# 🧠 Why Google Gemini?

Instead of relying on rule-based parsing or regular expressions, this project uses Google Gemini because it can understand resumes with different layouts and formats.

Advantages include:

- Supports multiple resume formats
- Produces structured JSON output
- Fast inference
- Free API tier for learning and small projects
- Easy Python integration

---

# 🖥 User Interface

The application provides:

- Simple PDF upload interface
- Loading spinner during analysis
- Success and error notifications
- Downloadable Excel report
- Optional Debug Mode to inspect extracted JSON and intermediate calculations

---

# 👨‍💻 Author

**Jai Pandya**

CV Analyzer is my first Python project, developed to automate HR resume screening by combining AI-powered resume information extraction with Python-based business logic and Excel automation.

---

# 📄 License

This project is intended for educational, learning, and demonstration purposes.
