# Resume-Expert Model Setup

This project uses a custom **`resume-expert`** model trained specifically for resume parsing. Follow these steps to set it up.

## Quick Start

### 1. Get the Model File

You should have received `resume-expert-model.tar.gz` from the project owner. This file contains the pre-trained model optimized for resume parsing.

**File size:** ~5.8GB (needs 6-7GB free disk space)

### 2. Extract the Model

```bash
# Navigate to your Ollama models directory
cd ~/.ollama/models

# Extract the model file (replace the path with where you saved it)
tar -xzf /path/to/resume-expert-model.tar.gz
```

**Example:**
```bash
cd ~/.ollama/models
tar -xzf ~/Downloads/resume-expert-model.tar.gz
```

### 3. Verify the Model is Loaded

```bash
ollama list
```

You should see output like:
```
NAME                    ID              SIZE      MODIFIED
resume-expert:latest    7f4439f6c590    2.0 GB    just now
```

### 4. Start Ollama

Make sure Ollama is running:

```bash
ollama serve
```

Leave this terminal running.

### 5. Run the Application

In a new terminal:

```bash
# Activate your virtual environment
source venv/bin/activate

# Start the app
python app.py
```

Visit **http://localhost:5001** in your browser.

### 6. Upload a Resume

- Click "Add Profile" or "Upload Files"
- Select a resume (PDF or DOCX)
- Choose **"Resume Intelligence"** for AI parsing
- The `resume-expert` model will automatically parse the resume
- Fields populate automatically
- Click "Save Profile"

---

## Troubleshooting

### "Model not found" or "resume-expert:latest not available"
- Verify extraction was successful: `ollama list`
- If not shown, re-extract the tar.gz file to `~/.ollama/models/`
- Restart Ollama: stop the `ollama serve` process and run it again

### "Ollama connection refused"
- Make sure Ollama is running: `ollama serve` in another terminal
- Check it's running on port 11434: `curl http://localhost:11434/api/tags`

### "Resume Intelligence option not available in browser"
- Make sure Ollama is running
- Refresh the browser page (Ctrl+R or Cmd+R)
- Check the browser console for errors (F12 → Console tab)

### Out of disk space during extraction
- Check available disk: `df -h`
- Need at least 6-7GB free for the tar extraction
- Delete other large files if needed, then retry extraction

### Model extraction is slow
- The tar.gz file is 5.8GB — extraction takes 2-5 minutes depending on disk speed
- Be patient, do NOT interrupt the extraction with Ctrl+C

---

## What is Resume-Expert?

`resume-expert` is a fine-tuned language model based on Llama 3.2, optimized specifically for:
- Extracting skills, experience, and education from resumes
- Handling various resume formats (2-column, 1-column, creative layouts)
- Parsing PDFs and DOCX files accurately
- Identifying job titles, companies, and responsibilities

It provides better accuracy than general-purpose models for resume parsing tasks.

---

## Questions?

Contact the project owner or check the main `SETUP_GUIDE.md` for additional help.
