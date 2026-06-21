# Quick Start Guide

**Can your friend use this repo and get the same results?** YES! Here's how:

## Prerequisites Checklist

Before starting, make sure you have these installed:

- [ ] Python 3.10+ — https://python.org
- [ ] PostgreSQL 13+ — https://postgresql.org  
- [ ] Git — https://git-scm.com
- [ ] Ollama (for AI resume parsing) — https://ollama.com
- [ ] 2-4 GB disk space for Ollama models
- [ ] ~15 minutes of setup time

## Step-by-Step Setup (5 minutes)

### 1. Clone & Setup Python

```bash
git clone <repo-url>
cd resume_profile_saved_list
python -m venv venv
source venv/bin/activate    # On Windows: venv\Scripts\activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Database Setup

Make sure PostgreSQL is running, then:

```bash
python create_db.py
psql -U postgres -d resume_profiles -f create_tables.sql
```

### 4. Start Ollama (in a separate terminal)

```bash
ollama serve
```

Then in another terminal:

```bash
ollama pull resume-expert   # Recommended; ~4GB
# OR
ollama pull llama3.2:latest # Alternative; ~2GB
```

### 5. Run the App

```bash
python app.py
```

Visit **http://localhost:5001** in your browser.

---

## What You Get

✅ **Resume Upload** — PDF and DOCX support  
✅ **Smart Parsing** — Extracts name, email, phone, LinkedIn, skills, experience, education  
✅ **2-Column Layout Handling** — Works with multi-column resume formats  
✅ **Download Options** — Export as PDF or JSON  
✅ **Job Matching** — Compare resumes against job descriptions  
✅ **Profile Management** — View, edit, delete profiles  

---

## Important Notes for Your Friend

### Model Choice

The app works with multiple Ollama models:

| Model | Size | Speed | Best For |
|-------|------|-------|----------|
| `resume-expert` | 4 GB | Slow | Best accuracy for resumes |
| `llama3.2:latest` | 2 GB | Fast | Good balance |
| `llama2:13b` | 7 GB | Very slow | Maximum accuracy |

**Recommendation:** Start with `llama3.2:latest` for a good balance of speed and accuracy.

### Database

The app uses PostgreSQL. If your friend doesn't have it installed:
- **macOS**: `brew install postgresql@15`
- **Windows**: Download installer from postgresql.org
- **Linux**: `sudo apt install postgresql`

### Port Conflict

If port 5001 is already in use, edit `app.py` line ~3127:
```python
app.run(debug=True, use_reloader=False, port=5002)  # Change to different port
```

### Common Issues

**"Ollama connection refused"**
- Make sure Ollama is running: `ollama serve`
- Check it's accessible: `curl http://localhost:11434/api/tags`

**"Database does not exist"**
- Run `python create_db.py` first
- Then run `psql ... -f create_tables.sql`

**"Port 5001 already in use"**
- Change the port in app.py (see above)
- Or kill the process: `lsof -i :5001 | grep Python | awk '{print $2}' | xargs kill -9`

---

## File Structure (Important Files)

```
resume_profile_saved_list/
├── requirements.txt         ← Install dependencies from this
├── .env                     ← Database & Ollama config (pre-configured)
├── SETUP_GUIDE.md          ← Detailed setup instructions
├── app.py                  ← Main Flask application
├── create_db.py            ← Database creation script
├── create_tables.sql       ← Database schema
└── templates/              ← HTML UI
```

---

## Next Steps After Setup

1. **Upload a Resume** → Go to /upload-files
2. **Choose "Resume Intelligence"** → Uses AI parsing (Ollama)
3. **View the Profile** → See extracted data
4. **Export** → Download as PDF or JSON

---

## Support

- Read `SETUP_GUIDE.md` for detailed troubleshooting
- Read `README.md` for full documentation
- Check browser console (F12) for JavaScript errors
- Check terminal output for Python errors

---

## Environment Variables Reference

`.env` file is pre-configured with:

```env
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/resume_profiles
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_TEXT_MODEL=resume-expert:latest      # Change if using different model
OLLAMA_TEXT_TIMEOUT=60
```

Change `OLLAMA_TEXT_MODEL` if you pulled a different model:
- `llama3.2:latest`
- `llama2:13b`
- etc.

---

## TL;DR (Copy-Paste Setup)

```bash
git clone <repo-url>
cd resume_profile_saved_list
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python create_db.py
psql -U postgres -d resume_profiles -f create_tables.sql

# In another terminal:
ollama serve

# In a third terminal:
ollama pull llama3.2:latest
cd resume_profile_saved_list
source venv/bin/activate
python app.py

# Then visit http://localhost:5001
```

**That's it!** 🚀
