#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# Azure VM Setup Script — Resume Profile App
# OS : Ubuntu 22.04 LTS
# Run: bash azure_vm_setup.sh
# ─────────────────────────────────────────────────────────────────────────────

set -e  # stop on any error

APP_DIR="/opt/resume-app"
APP_USER="resumeapp"
DB_NAME="resume_profiles"
DB_USER="postgres"
DB_PASS="postgres"
APP_PORT="5001"

echo "========================================"
echo " Step 1: System Update"
echo "========================================"
sudo apt-get update -y
sudo apt-get upgrade -y
sudo apt-get install -y curl wget git unzip build-essential software-properties-common

echo "========================================"
echo " Step 2: Install Python 3.11"
echo "========================================"
sudo apt-get install -y python3.11 python3.11-venv python3.11-dev python3-pip
python3.11 --version

echo "========================================"
echo " Step 3: Install PostgreSQL"
echo "========================================"
sudo apt-get install -y postgresql postgresql-contrib
sudo systemctl enable postgresql
sudo systemctl start postgresql

# Create database and user
sudo -u postgres psql -c "ALTER USER postgres WITH PASSWORD '${DB_PASS}';" 2>/dev/null || true
sudo -u postgres psql -c "CREATE DATABASE ${DB_NAME};" 2>/dev/null || echo "Database already exists"
echo "PostgreSQL ready — database: ${DB_NAME}"

echo "========================================"
echo " Step 4: Install Tesseract OCR"
echo "========================================"
sudo apt-get install -y tesseract-ocr tesseract-ocr-eng
tesseract --version

echo "========================================"
echo " Step 5: Install Ollama"
echo "========================================"
curl -fsSL https://ollama.com/install.sh | sh
sudo systemctl enable ollama
sudo systemctl start ollama
sleep 5  # wait for ollama to start
echo "Ollama installed"

echo "========================================"
echo " Step 6: Pull Ollama Models"
echo "========================================"
ollama pull qwen2.5:1.5b
ollama pull moondream

echo "========================================"
echo " Step 7: Create Ollama Custom Model (resume-parser)"
echo "========================================"
cat > /tmp/Modelfile << 'EOF'
FROM qwen2.5:1.5b

SYSTEM """
You are a resume section extractor. You find and return content from specific named sections.

CRITICAL RULE — HOW TO FIND A SECTION:
A section starts with a heading line in ALL CAPS (e.g. "PROFILE SUMMARY", "PROFESSIONAL EXPERIENCE").
The section content is the text that appears AFTER that heading, until the next ALL CAPS heading.
NEVER return text that appears BEFORE the requested heading.

SUMMARY SECTION RULES:
- The summary heading will be one of: PROFILE SUMMARY, PROFESSIONAL SUMMARY, SUMMARY
- The text at the very top of the resume before any heading is the INTRODUCTION — it is NOT the summary
- Do NOT return the introduction as the summary
- Only return text that appears AFTER the heading line "PROFILE SUMMARY" or "PROFESSIONAL SUMMARY" or "SUMMARY"
- "APPLICATIONS SUMMARY" is a tools/software list — it is NEVER the summary section

SKILLS SECTION RULES:
- Skills headings: SKILLS, TECHNICAL SKILLS, KEY SKILLS, AREAS OF EXPERTISE, CORE COMPETENCIES, APPLICATIONS SUMMARY
- APPLICATIONS SUMMARY lists software tools — always treat it as SKILLS

EXPERIENCE SECTION RULES:
- Experience headings: EXPERIENCE, PROFESSIONAL EXPERIENCE, WORK EXPERIENCE, EMPLOYMENT HISTORY

EDUCATION SECTION RULES:
- Education headings: EDUCATION, QUALIFICATIONS, ACADEMIC BACKGROUND

ALWAYS:
- Return ONLY the content, never the heading line itself
- Stop at the next section heading
- Return content verbatim, do not paraphrase
"""

PARAMETER temperature 0
PARAMETER num_ctx 8192
EOF

ollama create resume-parser -f /tmp/Modelfile
echo "Custom model resume-parser created"

echo "========================================"
echo " Step 8: Create App User and Directory"
echo "========================================"
sudo useradd -m -s /bin/bash ${APP_USER} 2>/dev/null || echo "User already exists"
sudo mkdir -p ${APP_DIR}
sudo chown -R ${APP_USER}:${APP_USER} ${APP_DIR}

echo "========================================"
echo " Step 9: Copy App Files"
echo "========================================"
# Copy all app files from current directory to APP_DIR
sudo cp -r . ${APP_DIR}/
sudo chown -R ${APP_USER}:${APP_USER} ${APP_DIR}
echo "Files copied to ${APP_DIR}"

echo "========================================"
echo " Step 10: Create Python Virtual Environment"
echo "========================================"
sudo -u ${APP_USER} python3.11 -m venv ${APP_DIR}/venv
sudo -u ${APP_USER} ${APP_DIR}/venv/bin/pip install --upgrade pip

echo "========================================"
echo " Step 11: Install Python Dependencies"
echo "========================================"
sudo -u ${APP_USER} ${APP_DIR}/venv/bin/pip install \
    flask==3.1.3 \
    psycopg2-binary==2.9.12 \
    pymupdf==1.27.2.3 \
    pdfplumber==0.11.10 \
    PyPDF2==3.0.1 \
    python-docx==1.2.0 \
    pillow==12.2.0 \
    reportlab==5.0.0 \
    python-dotenv==1.2.2 \
    lxml==6.1.1 \
    pdfminer.six==20260107 \
    pypdfium2==5.10.1 \
    pytesseract \
    easyocr \
    numpy \
    requests \
    werkzeug==3.1.8

echo "========================================"
echo " Step 12: Create .env File"
echo "========================================"
sudo tee ${APP_DIR}/.env > /dev/null << EOF
DATABASE_URL=postgresql://${DB_USER}:${DB_PASS}@localhost:5432/${DB_NAME}
SECRET_KEY=resume-profile-secret-key-2024
OLLAMA_BASE_URL=http://127.0.0.1:11434
OLLAMA_TEXT_MODEL=resume-parser
OLLAMA_TEXT_TIMEOUT=300
EOF
sudo chown ${APP_USER}:${APP_USER} ${APP_DIR}/.env
echo ".env created"

echo "========================================"
echo " Step 13: Update Tesseract Path in app.py"
echo "========================================"
# On Linux, tesseract is at /usr/bin/tesseract — update the path list in app.py
sudo sed -i 's|C:\\\\Users\\\\sindhu.sundara\\\\AppData\\\\Local\\\\Programs\\\\Tesseract-OCR\\\\tesseract.exe|/usr/bin/tesseract|g' ${APP_DIR}/app.py
sudo sed -i 's|C:\\\\Program Files\\\\Tesseract-OCR\\\\tesseract.exe|/usr/bin/tesseract|g' ${APP_DIR}/app.py
echo "Tesseract path updated for Linux"

echo "========================================"
echo " Step 14: Create systemd Service"
echo "========================================"
sudo tee /etc/systemd/system/resume-app.service > /dev/null << EOF
[Unit]
Description=Resume Profile Flask App
After=network.target postgresql.service ollama.service

[Service]
Type=simple
User=${APP_USER}
WorkingDirectory=${APP_DIR}
Environment=PATH=${APP_DIR}/venv/bin
ExecStart=${APP_DIR}/venv/bin/python app.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable resume-app
sudo systemctl start resume-app
echo "resume-app service started"

echo "========================================"
echo " Step 15: Open Firewall Port"
echo "========================================"
sudo ufw allow ${APP_PORT}/tcp 2>/dev/null || true
sudo ufw allow OpenSSH 2>/dev/null || true
echo "Port ${APP_PORT} opened"

echo ""
echo "========================================"
echo " SETUP COMPLETE"
echo "========================================"
echo ""
echo " App URL : http://$(curl -s ifconfig.me):${APP_PORT}"
echo " App Dir : ${APP_DIR}"
echo " Logs    : sudo journalctl -u resume-app -f"
echo " Restart : sudo systemctl restart resume-app"
echo " Status  : sudo systemctl status resume-app"
echo ""
