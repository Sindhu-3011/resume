# Resume Profile Manager - Setup Guide

Complete step-by-step instructions to set up and run the Resume Profile Manager on your machine.

---

## Prerequisites

Before you start, make sure you have installed:

- **Python 3.10+** — Download from [python.org](https://www.python.org/)
- **PostgreSQL 13+** — Download from [postgresql.org](https://www.postgresql.org/)
- **Ollama** (for AI resume parsing) — Download from [ollama.com](https://ollama.com/)
- **Git** — For cloning the repository

---

## Step 1: Clone the Repository

```bash
git clone <repository-url>
cd resume_profile_saved_list
```

---

## Step 2: Set Up Python Virtual Environment

Create and activate a virtual environment to keep dependencies isolated:

```bash
# Create virtual environment
python -m venv venv

# Activate it
# On macOS/Linux:
source venv/bin/activate

# On Windows:
venv\Scripts\activate
```

You should see `(venv)` at the start of your terminal prompt when activated.

---

## Step 3: Install Python Dependencies

```bash
pip install flask psycopg2-binary PyPDF2 pdfplumber pymupdf python-docx python-dotenv werkzeug
```

**Optional** — For OCR support on scanned PDFs:
```bash
pip install easyocr
```

---

## Step 4: Set Up PostgreSQL Database

### 4a. Start PostgreSQL

- **macOS (Homebrew)**: `brew services start postgresql@15`
- **macOS (App)**: Open the PostgreSQL app
- **Windows**: PostgreSQL service should auto-start
- **Linux**: `sudo systemctl start postgresql`

Verify it's running:
```bash
pg_isready -h localhost -p 5432
```

Expected output: `localhost:5432 - accepting connections`

### 4b. Create the Database

```bash
python create_db.py
```

Output should say: `Created resume_profiles database`

### 4c. Create Database Tables

```bash
# macOS/Linux:
PGPASSWORD=postgres psql -h localhost -U postgres -d resume_profiles -f create_tables.sql

# Windows (Command Prompt):
set PGPASSWORD=postgres
psql -h localhost -U postgres -d resume_profiles -f create_tables.sql
```

Or run via Python:
```bash
python << 'EOF'
import psycopg2

conn = psycopg2.connect("postgresql://postgres:postgres@localhost:5432/resume_profiles")
cur = conn.cursor()

with open("create_tables.sql", "r") as f:
    cur.execute(f.read())

conn.commit()
conn.close()
print("Tables created successfully")
EOF
```

---

## Step 5: Set Up Ollama (for AI Resume Parsing)

### 5a. Install Ollama

Download and install from [ollama.com](https://ollama.com/)

### 5b. Start Ollama Server

```bash
ollama serve
```

This will run Ollama on `http://localhost:11434`

### 5c. Download a Model

In a **new terminal**, pull one or more models:

```bash
# Small & fast (recommended for most use cases):
ollama pull llama3.2:latest

# Larger & more accurate (if you have the disk space):
ollama pull deepseek-r1:8b
```

### 5d. Verify Models Are Ready

```bash
curl http://localhost:11434/api/tags
```

You should see JSON listing all installed models.

---

## Step 6: (Optional) Create Environment Configuration

If you want to customize settings, create a `.env` file in the project root:

```env
# Database connection (use these defaults if you followed step 4)
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/resume_profiles

# Flask secret key (for sessions)
SECRET_KEY=your-secret-key-here

# Ollama configuration (use these defaults)
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_TEXT_MODEL=llama3.2:latest
OLLAMA_TEXT_TIMEOUT=60
```

---

## Step 7: Run the Application

Make sure you're in the project directory with `(venv)` activated:

```bash
python app.py
```

You should see:
```
 * Serving Flask app 'app'
 * Debug mode: on
 * Running on http://127.0.0.1:5001
```

---

## Step 8: Open the App

Visit **http://localhost:5001** in your browser.

You should see the "Add Profile" page with a file upload form.

---

## Troubleshooting

### "Database does not exist"
- Run `python create_db.py` to create the database
- Run the SQL file to create tables

### "PostgreSQL connection refused"
- Check that PostgreSQL is running: `pg_isready -h localhost -p 5432`
- Start PostgreSQL if it's not running

### "Ollama connection refused"
- Make sure Ollama is running: `ollama serve` in another terminal
- Check that Ollama is at `http://localhost:11434`

### "Resume Intelligence option not available"
- Make sure Ollama is running and models are downloaded
- Check the browser console for errors

### Port 5001 already in use
- Find what's using it: `lsof -i :5001`
- Or edit `app.py` line 3127 to use a different port: `app.run(debug=True, use_reloader=False, port=5002)`

---

## Quick Usage Guide

1. **Upload a Resume**
   - Go to "Upload Files" or "Add Profile"
   - Select a PDF or DOCX file
   - Choose **"Resume Intelligence"** for AI parsing (requires Ollama)
   - Fields populate automatically
   - Click "Save Profile"

2. **View Profiles**
   - Go to "View Profiles"
   - Click on a profile to see details

3. **Add Job Descriptions**
   - Go to "JD Management"
   - Click "Add JD"
   - Fill in job title, role, skills, requirements

4. **Compare Resume vs JD**
   - View a profile
   - Click "Compare Against JD"
   - Select a job description
   - See match score and missing skills

---

## Directory Structure

```
resume_profile_saved_list/
├── app.py                      # Main Flask application
├── create_db.py               # Database creation script
├── create_tables.sql          # Database schema
├── venv/                      # Python virtual environment (created in step 2)
├── templates/                 # HTML templates
├── static/                    # CSS and static files
├── uploads/                   # Uploaded resume files
└── README.md                  # Original full documentation
```

---

## Need Help?

- Check the original `README.md` for more details
- Verify PostgreSQL is running
- Verify Ollama is running (for AI features)
- Check browser console for JavaScript errors
- Check terminal output for Python errors

---

## What's Next?

Once everything is working:
- Upload some resumes
- Try the Resume Intelligence AI parser
- Add job descriptions
- Compare resumes against JD requirements
- View match scores and recommendations
