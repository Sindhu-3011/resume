# Resume Profile Manager - Setup Guide

Complete step-by-step instructions to set up and run the Resume Profile Manager on your machine.

---

## Prerequisites

Before you start, make sure you have:

### Required Files
- **`resume-expert-model.tar.gz`** — The AI model file for resume parsing
  - Get this from the project owner (it's a 5.8GB file)
  - Save it somewhere on your computer (e.g., Downloads folder)
  - You'll need it in Step 5c below

### Required Software
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

With your virtual environment activated, install all dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```

This installs:
- Flask (web framework)
- PostgreSQL driver (psycopg2)
- PDF libraries (PyPDF2, pdfplumber, PyMuPDF)
- Document support (python-docx)
- PDF generation (reportlab)
- Environment config (python-dotenv)
- And other dependencies

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

### 5c. Load the Resume-Expert Model

**If you have the `resume-expert-model.tar.gz` file:**

```bash
# Extract the model to your Ollama models directory
cd ~/.ollama/models
tar -xzf /path/to/resume-expert-model.tar.gz

# Verify it's loaded
ollama list
# You should see: resume-expert:latest in the list
```

**If you don't have the model file, download one:**

In a **new terminal**, pull one of these open-source models:

```bash
# RECOMMENDED — Best balance of accuracy and speed for resume parsing:
ollama pull llama2:13b

# ALTERNATIVE — Faster, good for quick parsing (less VRAM):
ollama pull mistral

# PREMIUM — Better accuracy if you have 16GB+ VRAM available:
ollama pull llama3.2:latest
```

**Choose based on your system:**
- **4-8GB RAM** → use `mistral` (fastest)
- **8-16GB RAM** → use `llama2:13b` (recommended default)
- **16GB+ RAM** → use `llama3.2:latest` (best accuracy)

The first pull takes 5–10 minutes depending on your internet speed and model size.

### 5d. Verify Models Are Ready

```bash
curl http://localhost:11434/api/tags
```

You should see JSON listing all installed models.

---

## Step 6: (Optional) Create Environment Configuration

A `.env` file is already provided with defaults. It's configured for the `resume-expert` model:

```env
# Database connection (use these defaults if you followed step 4)
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/resume_profiles

# Flask secret key (for sessions) - change this to a random string
SECRET_KEY=your-secret-key-here

# Ollama configuration
OLLAMA_BASE_URL=http://localhost:11434
# This is pre-configured for resume-expert:latest
OLLAMA_TEXT_MODEL=resume-expert:latest
OLLAMA_TEXT_TIMEOUT=60
```

**If you used a different model from Step 5c, update `OLLAMA_TEXT_MODEL`:**
- If you extracted `resume-expert-model.tar.gz` → use `resume-expert:latest` (default - recommended)
- If you pulled `llama2:13b` → use `llama2:13b`
- If you pulled `mistral` → use `mistral`
- If you pulled `llama3.2:latest` → use `llama3.2:latest`

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
- Make sure Ollama is running: `ollama serve` in another terminal
- Check that the model from Step 5c downloaded successfully: `curl http://localhost:11434/api/tags`
- Verify the `OLLAMA_TEXT_MODEL` in `.env` matches your downloaded model
- Check the browser console for errors
- Try refreshing the page

### "Model download failed" or "Model not found"
- Check your internet connection — model files are large (5–15GB)
- If the download was interrupted, delete and retry:
  ```bash
  ollama rm llama2:13b
  ollama pull llama2:13b
  ```
- Try a smaller model if you're low on disk space: `ollama pull mistral`
- Check available disk space: `df -h` (need at least 20GB free)

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
