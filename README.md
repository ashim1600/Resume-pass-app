# Resume Builder App

Streamlit app that takes a job description and your base profile/resume, then generates a tailored resume draft.

## Features

- Tailored resume draft generation from a job description
- Keyword extraction from the job description
- Summary, skills, experience, projects, and education sections
- TXT and DOCX download output
- Optional PDF, DOCX, and TXT upload support for source material

## Installation & Setup

On macOS with managed Python environments, use a virtual environment:

```bash
cd resume-pass-app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
source .venv/bin/activate
streamlit run app.py
```

The app will be available at `http://localhost:8501`

## Screenshots

The app provides:
- Job description upload or text paste
- Resume/profile information input forms
- Tailored resume generation with keyword alignment
- Export to DOCX and TXT formats
- Quality analysis and fit assessment

## Notes

- You still need to provide your own background facts. The app will not invent experience.
- You can paste text directly if file parsing is unavailable.
- PDF and DOCX extraction depends on optional packages listed in `requirements.txt`.
- The `requirements.txt` includes: streamlit, pypdf, python-docx
