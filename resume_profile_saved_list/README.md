# Resume Profile Manager

A Flask web app for uploading, parsing, and managing candidate resume profiles. Upload PDF or DOCX resumes, auto-extract structured data (name, contact info, skills, experience, education), manage job descriptions, and compare candidates against JD requirements with a match score.

## Features

- **Resume Upload** — Upload PDF and DOCX files; raw originals are preserved
- **Auto-Parsing** — Extracts name, email, phone, LinkedIn, location, summary, skills, experience, education, certifications, and projects
- **Resume Intelligence (AI parsing)** — Optional "llm" parse mode uses a local [Ollama](https://ollama.com) model to extract identity fields and sections, instead of the regex-based Quick Parse
- **2-Column Layout Support** — Handles multi-column resume formats using word-level bounding boxes
- **Optional OCR** — Falls back to EasyOCR for image-based or scanned PDFs
- **Profile Management** — View, edit, and delete parsed profiles; public URLs via slugs
- **Skill Normalization** — Skills are tokenized into a separate `resume_skill` table for filtering
- **Role Groups** — Browse profiles grouped by job title/role
- **Job Description (JD) Management** — Add and manage JDs with skills, requirements, and keywords
- **Resume vs JD Matching** — Score a resume against a JD based on skills, experience, and keywords
- **Dashboard** — Overview stats and recently uploaded resumes
- **Bulk Compare** — Compare multiple resumes against a JD at once

## Tech Stack

| Layer | Library |
|---|---|
| Web framework | Flask |
| Database | PostgreSQL (psycopg2) |
| PDF extraction | PyPDF2, pdfplumber, PyMuPDF |
| DOCX extraction | python-docx |
| OCR (optional) | EasyOCR + PyMuPDF |
| AI parsing (optional) | Ollama (local LLM, e.g. `llama3.2:1b`) |
| Config | python-dotenv |

## Prerequisites

- Python 3.10+
- PostgreSQL 13+ running locally (or a remote instance)
- [Ollama](https://ollama.com) — optional, only needed for the "Resume Intelligence" AI parse mode

## Setup

### 1. Clone and create a virtual environment

```bash
git clone <repo-url>
cd resume_profile_saved_list
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
```

### 2. Install dependencies

```bash
pip install flask psycopg2-binary PyPDF2 pdfplumber pymupdf python-docx python-dotenv werkzeug
```

For optional OCR support (scanned / image-based PDFs):

```bash
pip install easyocr
```

The AI parser talks to Ollama over HTTP using only the standard library, so no extra
Python package is required for it.

### 3. (Optional) Set up Ollama for AI parsing

The "Resume Intelligence" parse mode runs a local LLM through [Ollama](https://ollama.com).
Quick Parse (the regex-based parser) works without it — this step is only needed if you
want the AI parse mode.

1. Install Ollama (see <https://ollama.com/download>).
2. Start the Ollama server (it usually runs automatically after install):

   ```bash
   ollama serve
   ```

3. Pull the default model:

   ```bash
   ollama pull llama3.2:1b
   ```

By default the app talks to Ollama at `http://localhost:11434` using the `llama3.2:1b`
model. Override these via the `OLLAMA_*` environment variables (see below).

### 4. Configure the database connection

Create a `.env` file in the project root:

```env
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/resume_profiles
SECRET_KEY=your-secret-key-here

# Optional — only used by the AI ("Resume Intelligence") parse mode
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_TEXT_MODEL=llama3.2:1b
OLLAMA_TEXT_TIMEOUT=60
```

The app defaults to `postgresql://postgres:postgres@localhost:5432/resume_profiles` if `DATABASE_URL` is not set.

### 5. Create the database

```bash
python create_db.py
```

This connects to the default `postgres` database and creates `resume_profiles` if it does not exist.

### 6. Create tables

Connect to `resume_profiles` and run the SQL file:

```bash
psql -U postgres -d resume_profiles -f create_tables.sql
```

Or open `create_tables.sql` in DBeaver / psql and execute it manually.

### 7. (Optional) Seed sample data

```bash
psql -U postgres -d resume_profiles -f populate.sql
```

## Running the App

```bash
python app.py
```

The app starts on `http://127.0.0.1:5000` in debug mode.

## Project Structure

```
.
├── app.py                  # Flask application — routes, parsing, DB logic
├── create_db.py            # One-time script to create the PostgreSQL database
├── create_tables.sql       # DDL — creates resume, resume_skill, schema_version tables
├── populate.sql            # Optional seed data
├── templates/              # Jinja2 HTML templates
│   ├── base.html
│   ├── dashboard.html
│   ├── upload_files.html
│   ├── profile_list.html
│   ├── profile.html
│   ├── edit.html
│   ├── groups.html
│   ├── jd_management.html
│   ├── jd_form.html
│   ├── jd_detail.html
│   ├── compare_select.html
│   └── compare_result.html
├── static/
│   └── styles.css
└── uploads/                # Uploaded resume files (gitignored in practice)
    └── raw/                # Original raw uploads + _meta.json index
```

## Key Routes

| Route | Description |
|---|---|
| `/` | Home (redirects to dashboard) |
| `/dashboard` | Overview stats and recent uploads |
| `/upload-files` | Upload one or more PDF/DOCX resumes |
| `/profiles` | List all parsed profiles |
| `/profile/<id>` | Profile detail view |
| `/edit/<id>` | Edit a profile |
| `/groups` | Browse profiles grouped by role |
| `/jd-management` | List all job descriptions |
| `/jd/add` | Add a new JD |
| `/compare/<resume_id>/<jd_id>` | Match score: resume vs JD |
| `/api/parse-resume` | JSON API — parse a resume file |
| `/api/resume` | JSON API — fetch a stored resume |

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `DATABASE_URL` | `postgresql://postgres:postgres@localhost:5432/resume_profiles` | PostgreSQL connection string |
| `SECRET_KEY` | `resume-profile-secret-key` | Flask session secret |
| `OLLAMA_BASE_URL` | `http://localhost:11434` | Ollama server URL (AI parse mode) |
| `OLLAMA_TEXT_MODEL` | `llama3.2:1b` | Ollama model used for resume parsing |
| `OLLAMA_TEXT_TIMEOUT` | `60` | Per-request Ollama timeout (seconds) |

## Database Schema

```
resume          — one row per candidate profile
resume_skill    — normalised skills (many-to-one → resume)
job_description — JDs created via the UI (created on first use)
schema_version  — tracks migration version
```
