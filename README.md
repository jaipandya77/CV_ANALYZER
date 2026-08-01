
# CV Analyzer

An AI-powered web application that analyzes PDF resumes and automatically populates a predefined HR Excel template.

Built using **Python**, **Streamlit**, **Google Gemini API (Free Tier)**, and **OpenPyXL**, the application extracts structured candidate information from resumes and performs HR-specific calculations before generating a ready-to-download Excel report.

---

## Features

- 📄 Upload resumes in PDF format
- 🤖 AI-powered resume information extraction using Google Gemini
- 📊 Automatic HR Excel template population
- 🎓 Detects highest qualification and educational institute
- 🏫 Looks up NIRF institute rankings
- 💼 Calculates total work experience
- 📈 Calculates experience as:
  - Contractor
  - Owner
  - Consultant
  - Freelancer
- 🌍 Calculates experience:
  - In India
  - Outside India
- 🏢 Calculates experience in:
  - MNCs
  - Listed Companies
- 🔄 Calculates number of job changes
- ⏳ Calculates average tenure
- ⚠️ Detects employment gaps (only if greater than 1 year)
- 📜 Counts certifications
- 📥 Download the completed HR Excel report

---

# Technologies Used

- Python
- Streamlit
- Google Gemini API (Free Tier)
- Gemini 2.5 Flash
- PyPDF2
- OpenPyXL

---

# AI Model

This project uses **Google Gemini 2.5 Flash** through the **Google Gemini API (Free Tier)** for extracting structured information from resumes.

The AI model is used **only for extracting information** from the resume.

All business logic, including:

- Work experience calculation
- Employment gap detection
- Job change calculation
- Average tenure calculation
- NIRF institute ranking lookup
- Excel report generation

is implemented entirely in Python, ensuring accurate and deterministic results.

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

```

---

# Installation

## Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/CV-Analyzer.git
```

Move into the project directory.

```bash
cd CV-Analyzer
```

Install the required packages.

```bash
pip install -r requirements.txt
```

---

# Configure API Key

This application uses **Streamlit Secrets** to securely store the Google Gemini API key.

After deploying the application on **Streamlit Community Cloud**:

1. Open your deployed application.
2. Navigate to **Settings → Secrets**.
3. Add your Gemini API key in the following format:

```toml
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

The application securely reads the API key from Streamlit Secrets, ensuring that sensitive credentials are never stored in the source code or GitHub repository.

---

# Running the Application

Run the application locally using:

```bash
streamlit run app.py
```

The application will automatically open in your default web browser.

---

# Application Workflow

```
             PDF Resume
                  │
                  ▼
      Extract Text using PyPDF2
                  │
                  ▼
      Google Gemini 2.5 Flash API
                  │
                  ▼
       Structured JSON Response
                  │
                  ▼
      Python Business Logic Layer
                  │
                  ▼
   HR Calculations & Validations
                  │
                  ▼
 Populate HR Excel Template
                  │
                  ▼
 Download Completed Excel Report
```

---

# Business Rules

The application follows the following HR business rules:

- Employment gaps are reported **only if the gap exceeds one year**.
- Average tenure is calculated as:

```
Average Tenure = Total Experience ÷ Number of Job Changes
```

- Educational institute ranking is determined using the included NIRF rankings dataset.
- AI is responsible only for information extraction.
- All calculations and validations are performed in Python.

---

# Why Google Gemini?

Instead of building a rule-based parser using regular expressions, this project uses **Google Gemini 2.5 Flash** because it can accurately understand resumes with different layouts and formats.

Benefits include:

- Supports resumes with varying structures
- Extracts structured JSON directly
- Free API tier suitable for learning and small projects
- Fast inference
- Easy integration with Python

---

# Future Improvements

- Batch resume processing
- OCR support for scanned PDF resumes
- Resume scoring and ranking
- Candidate dashboard
- Database integration
- Support for DOCX resumes
- Export results to CSV
- Multi-user authentication

---

# Screenshots

Add screenshots of the application inside an `assets` folder.

Example:

```
assets/
├── home.png
├── upload.png
├── output.png
```

Then display them like this:

```markdown
## Home Screen

![Home](assets/home.png)

## Generated Report

![Output](assets/output.png)
```

---

# Author

**Jai Pandya**

This project was developed as my first Python application to automate HR resume screening by combining AI-powered information extraction with automated Excel report generation using Streamlit.

---

# License

This project is intended for educational, demonstration, and learning purposes.
