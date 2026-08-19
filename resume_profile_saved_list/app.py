from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, abort, send_file, session
import psycopg2
from io import BytesIO, StringIO
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.colors import HexColor, white as RL_WHITE
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
from psycopg2.extras import RealDictCursor
from pathlib import Path
import re
import math
import hashlib
import secrets
import smtplib
from email.mime.text import MIMEText
import uuid
from collections import Counter
import zipfile
import json
import logging
import os
import urllib.request
import urllib.error
from contextlib import contextmanager
from datetime import datetime, timezone, timedelta
import threading
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from PyPDF2 import PdfReader
from docx import Document

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import easyocr
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import pymupdf as _pymupdf_check
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

# ── Tesseract OCR binary path (new code — do not modify existing code) ────────
# winget installs to AppData\Local on per-user basis; set cmd path explicitly
# so pytesseract can find the binary regardless of PATH.
try:
    import pytesseract as _pytess_check
    _TESSERACT_PATHS = [
        r"C:\Users\sindhu.sundara\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for _tpath in _TESSERACT_PATHS:
        if os.path.isfile(_tpath):
            _pytess_check.pytesseract.tesseract_cmd = _tpath
            break
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "resume-profile-secret-key")

# Auto-logout on inactivity — how long a session may sit idle (no request)
# before _require_login() below invalidates it. Configurable via env var;
# the two presets called out in the spec are 15 (default) and 30 minutes.
SESSION_TIMEOUT_MINUTES = int(os.environ.get("SESSION_TIMEOUT_MINUTES", "15"))
SESSION_WARNING_SECONDS = 60  # show the "continue session?" dialog this long before expiry

# Forgot Password / Reset Password config — see _require_permission-style
# helpers near ensure_users_table for the rest of this feature.
PASSWORD_RESET_TOKEN_MINUTES = int(os.environ.get("PASSWORD_RESET_TOKEN_MINUTES", "30"))
PASSWORD_HISTORY_COUNT = int(os.environ.get("PASSWORD_HISTORY_COUNT", "5"))
PASSWORD_MIN_LENGTH = int(os.environ.get("PASSWORD_MIN_LENGTH", "8"))
# A password older than this forces a mandatory change at next login (see
# login()'s expiry check and change_expired_password()) — 0 disables the check.
PASSWORD_MAX_AGE_DAYS = int(os.environ.get("PASSWORD_MAX_AGE_DAYS", "90"))
# No email infra existed anywhere in this app before this feature. If
# SMTP_HOST is left blank (the default — this is a local dev tool with no
# outbound mail relay configured anywhere), _send_email() logs the message
# instead of attempting a real send, so the whole flow stays testable
# without real credentials. Set SMTP_HOST (+ the rest) in .env to send real mail.
SMTP_HOST = os.environ.get("SMTP_HOST", "")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USERNAME = os.environ.get("SMTP_USERNAME", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SMTP_FROM_EMAIL = os.environ.get("SMTP_FROM_EMAIL", "no-reply@resumeprofile.local")
SMTP_USE_TLS = os.environ.get("SMTP_USE_TLS", "true").lower() == "true"

# Account lockout — locks an account after this many consecutive failed
# logins (reset to 0 by any successful login). No auto-expiry: an Admin
# must manually unlock via User Management (see user_unlock()).
MAX_FAILED_LOGIN_ATTEMPTS = int(os.environ.get("MAX_FAILED_LOGIN_ATTEMPTS", "5"))

@app.after_request
def set_cache_control(response):
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


from markupsafe import Markup, escape as _escape


@app.template_filter("verdict_tier")
def verdict_tier(value):
    """Map a verdict/tier string to its tier slug (strong|good|partial|low).

    Server-side mirror of static/verdict.js so the fit-tier COLOURS live only in
    styles.css (.verdict-* classes) instead of being hardcoded as hex values in
    each template. Keeps jd_top_matches.html in sync with the JS-driven pages.
    """
    v = str(value or "").lower()
    if "strong" in v:
        return "strong"
    if "good" in v:
        return "good"
    if "partial" in v:
        return "partial"
    return "low"


@app.template_filter("nl2br")
def nl2br(value):
    """Escape untrusted text first, THEN turn newlines into <br>.

    Resume/JD section text is auto-parsed from uploaded PDFs/DOCX and is
    therefore untrusted — the previous `| replace('\\n','<br>') | safe`
    rendered that content as raw HTML. This escapes the text (so any markup in
    the document shows as literal characters) before inserting the line breaks.
    """
    if value is None:
        return ""
    return Markup("<br>".join(_escape(line) for line in str(value).split("\n")))

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"

# PostgreSQL connection — set DATABASE_URL in .env or environment.
# Default assumes a local "resume_profiles" database owned by the current OS user.
DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/resume_profiles",
)
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "docx"}
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

SECTION_ALIASES = {
    # ── Identity / header ─────────────────────────────────────────────────────
    "summary": [
        "summary", "professional summary", "profile summary",
        "executive summary", "career summary", "career profile",
        "professional profile", "personal profile",
        "professional overview", "career overview", "overview",
        "objective", "career objective", "professional objective",
        "job objective", "employment objective", "work objective",
        "position objective", "career aim", "professional aim",
        "about", "about me", "about myself",
        "personal statement", "professional statement", "introduction",
        "professional synopsis", "career synopsis", "synopsis",
        "brief profile", "candidate profile", "career snapshot",
        "profile overview", "professional background summary",
        "qualifications summary", "summary of qualifications",
        "highlights of qualifications",
        "professional bio", "bio", "personal bio", "biography",
        "professional goals", "career goals",
        "professional mission", "mission statement",
        "at a glance", "professional snapshot",
        "career brief", "professional brief",
        "personal overview", "personal summary", "work summary",
        "candidate overview", "professional profile summary",
        "brief about me", "career statement", "profile",
        "value proposition", "professional pitch",
        "personal pitch", "candidate summary",
        "skills summary profile", "background",
        # abbreviated forms: "Prof. Summary" → "prof summary"
        "prof summary", "exec summary", "prof profile",
        "career obj", "prof objective",
    ],

    # ── Skills / competencies ─────────────────────────────────────────────────
    "skills": [
        "skills", "skill set", "skillset", "skills summary",
        "technical skills", "core skills", "key skills",
        "professional skills", "functional skills",
        "competencies", "core competencies", "key competencies",
        "professional competencies", "technical competencies",
        "skills and competencies", "competency profile",
        "technologies", "tools and technologies", "tools",
        "technical tools", "software tools", "tools and software",
        "tools and applications", "tools applications", "tools and skills",
        "tools used", "tools technologies",
        "areas of expertise", "area of expertise",
        "expertise", "domain expertise", "domain skills",
        "technical expertise", "professional expertise",
        "technical proficiencies", "proficiencies",
        "applications summary", "application summary", "applications",
        "it skills", "skills overview", "technical skill set",
        "skills and frameworks", "skills and technologies",
        "technical knowledge", "knowledge and skills",
        "skills and expertise", "key technical skills",
        "technology skills",
        "soft skills", "interpersonal skills", "personal skills",
        "transferable skills", "hard skills",
        "abilities", "capabilities", "technical abilities",
        "frameworks and tools", "programming skills",
        "languages and tools", "platforms and tools",
        "technologies and tools",
        "it proficiencies", "technical profile",
        "expertise areas", "areas of proficiency",
        "specializations", "specialties", "specialisation",
        "core technical skills", "relevant skills",
        # abbreviated / punctuated forms normalized by canonical_section_name regex
        # e.g. "I.T. Skills" → "i t skills", "Tech. Skills" → "tech skills"
        "i t skills", "tech skills", "i t proficiencies",
        "tech proficiencies", "tech expertise",
        "technical skills & tools", "technical skills and tools",
        "technical skills tools",
    ],

    # ── Work experience ───────────────────────────────────────────────────────
    "experience": [
        "EXPERIENCE",
        "Experience", "WORK EXPERIENCE", "Work Experience",
        "experience", "work experience", "professional experience",
        "employment history", "career history", "work history",
        "employment", "professional background",
        "job experience", "relevant experience", "industry experience",
        "work experience details", "experience summary",
        "employment details", "professional history",
        "career experience", "professional work experience",
        "work experience summary", "key experiences",
        "work details", "career details", "employment record",
        "professional record", "experience details",
        "professional engagements", "roles and responsibilities",
        "current experience", "past experience",
        "career overview details",
        "work profile", "working experience",
        "job history", "positions held",
        "appointments", "professional appointments",
        "corporate experience", "it experience",
        "domain experience",
        "relevant work experience", "key work experience",
        "career work history",
        "early career experience", "early career",
        # abbreviated forms: "Prof. Experience" → "prof experience"
        "prof experience", "prof work experience",
        "work exp", "professional exp",
    ],

    # ── Internships ───────────────────────────────────────────────────────────
    "internships": [
        "internship", "internships",
        "internship experience", "internship history",
        "industrial training", "industry training",
        "vocational training",
        "summer internship", "summer training",
        "academic internship", "intern experience",
        "trainee experience", "apprenticeship",
        "co op experience", "coop experience",
        "cooperative education", "practicum",
        "field experience", "work placement",
        "placement experience", "industrial placement",
        "sandwich year", "year in industry",
        "graduate training", "graduate placement",
        "part time experience", "student work experience",
    ],

    # ── Education ─────────────────────────────────────────────────────────────
    "education": [
        "EDUCATION",
        "education", "academic background", "qualifications",
        "academic qualification", "academics",
        "educational background", "academic credentials",
        "educational qualifications", "educational qualification", "academic details",
        "educational details", "academic qualifications",
        "education details", "qualification details",
        "academic history", "academic information",
        "educational information", "scholastic details",
        "scholastic background", "educational profile",
        "education and training",
        "academic record", "educational record",
        "academic profile", "studies",
        "degrees", "degrees earned", "degrees obtained",
        "formal education", "academic education",
        "educational credentials",
        "examination details", "scholastic record",
        "qualifications and education", "degree information",
        "academic qualifications and education",
        "education & certifications", "education and certifications",
        "education certifications",
    ],

    # ── Certifications ────────────────────────────────────────────────────────
    "certifications": [
        "certifications", "certificates", "certification",
        "professional certifications", "it certifications",
        "technical certifications",
        "certifications and credentials", "credentials",
        "professional credentials",
        "industry certifications",
        "certified", "certified courses",
        "certification summary",
        "licenses and certifications",
        "certificates and licenses",
        "professional development",
        "courses and certifications",
        "professional courses",
        "online courses",
        "certificates courses", "certificate courses",
        "certificates and courses", "courses and certificates",
        "certifications courses", "courses certifications",
        "certification and training",
        "certifications and training",
        "certifications and professional development",
        "certifications   professional development",
        "professional development certifications",
        "microsoft certifications", "aws certifications",
        "google certifications", "oracle certifications",
        "cisco certifications",
    ],

    # ── Training / workshops ──────────────────────────────────────────────────
    "training": [
        "training", "trainings", "training and development",
        "professional training", "technical training",
        "corporate training", "training attended",
        "training programs", "training and courses",
        "courses", "workshops", "seminars",
        "workshops and seminars", "seminars and workshops",
        "training and workshops", "training and seminars",
        "continuing education", "continuing professional development",
        "cpd",
        "professional development courses",
        "training completed", "programmes attended",
        "programs attended",
    ],

    # ── Licenses ──────────────────────────────────────────────────────────────
    "licenses": [
        "licenses", "licence", "licences", "license",
        "professional licenses", "licenses and permits",
        "state licenses", "federal licenses",
        "regulatory licenses",
        "licenses and registrations",
        "registration", "registrations",
        "professional registrations",
        "license and certification",
        "professional license",
    ],

    # ── Projects ──────────────────────────────────────────────────────────────
    "projects": [
        "project", "projects", "key projects", "project experience", "assignments",
        "notable projects", "project highlights", "academic projects",
        "personal projects", "relevant project experience",
        "project details", "project organizational details",
        "organizational details", "relevant project organizational details",
        "project and organizational details",
        "relevant project organizational details",
        "project organizational details roles and responsibilities",
        "project organizational details roles and responsibilities",
        "client details", "relevant projects",
        "project summary", "projects summary", "project overview",
        "major projects", "project work", "projects worked",
        "key projects worked", "significant projects", "project details summary",
        "live projects", "work projects", "project descriptions",
        "side projects", "open source projects",
        "portfolio", "portfolio projects",
        "technical projects", "professional projects",
        "freelance projects", "contract projects",
        "client projects", "project portfolio",
        # "Projects Completed" style headings (incl. the common "compelted" typo)
        "projects completed", "project completed", "completed projects",
        "projects compelted", "projects compeleted", "projects done",
    ],

    # ── Achievements / accomplishments ────────────────────────────────────────
    "achievements": [
        "achievements", "achievement", "key achievements",
        "professional achievements", "career achievements",
        "awards and achievements", "achievements and awards",
        "accomplishments", "key accomplishments",
        "professional accomplishments", "notable accomplishments",
        "major accomplishments", "career accomplishments",
        "honors", "honours", "academic honors", "academic honours",
        "distinction", "distinctions",
        "recognition", "accolades",
        "performance highlights", "significant achievements",
        "academic achievements",
        "notable contributions",
    ],

    # ── Awards ────────────────────────────────────────────────────────────────
    "awards": [
        "awards", "award", "awards and recognition", "awards recognition",
        "recognition and awards", "honors and recognition",
        "honours and recognition", "recognition and honors",
        "prizes", "prizes and awards",
        "scholarships", "fellowships",
        "scholarships and awards",
        "grants", "grants and awards",
        "commendations", "merits",
        "industry awards", "professional awards",
    ],

    # ── Publications ──────────────────────────────────────────────────────────
    "publications": [
        "publications", "publication", "published works",
        "published papers", "papers",
        "journal publications", "journal articles",
        "articles", "research publications",
        "peer reviewed publications", "peer reviewed articles",
        "technical publications",
        "books", "book chapters", "chapters",
        "conference papers", "conference publications",
        "technical papers",
        "white papers", "whitepapers",
        "case studies",
        "authored works", "writings", "written works",
        "papers and publications", "published research",
        "scholarly publications", "academic publications",
    ],

    # ── Research ──────────────────────────────────────────────────────────────
    "research": [
        "research", "research experience", "research work",
        "research background", "research and development",
        "r and d", "research projects",
        "research interests", "areas of research",
        "research contributions",
        "thesis", "dissertation", "research thesis",
        "academic research", "applied research",
        "research activities", "research assignments",
        "research summary", "research overview",
        "doctoral research", "postdoctoral research",
    ],

    # ── Patents ───────────────────────────────────────────────────────────────
    "patents": [
        "patents", "patent", "intellectual property",
        "patents filed", "patents granted",
        "patents and trademarks", "inventions",
        "innovations", "filed patents",
        "patent applications", "utility patents",
    ],

    # ── Conferences / presentations ───────────────────────────────────────────
    "conferences": [
        "conferences", "conference", "conference presentations",
        "presentations", "presentations and talks",
        "talks", "speaking engagements",
        "conference attendance", "conferences attended",
        "seminars attended",
        "workshops attended",
        "conferences and events", "events",
        "keynote presentations",
        "panel discussions",
        "posters", "poster presentations",
        "conference sessions", "invited talks",
        "technical talks",
    ],

    # ── Volunteer / community ─────────────────────────────────────────────────
    "volunteer": [
        "volunteer experience", "volunteering", "volunteer",
        "volunteer work", "community service",
        "community involvement", "community engagement",
        "social work", "charity work",
        "non profit experience", "nonprofit experience",
        "voluntary work", "voluntary experience",
        "pro bono work", "civic engagement",
        "community activities", "philanthropy",
        "service activities",
        "outreach", "community outreach",
        "social impact", "corporate social responsibility",
        "ngo experience",
    ],

    # ── Leadership ────────────────────────────────────────────────────────────
    "leadership": [
        "leadership", "leadership experience", "leadership roles",
        "leadership and management", "management experience",
        "team leadership", "leadership positions",
        "positions of responsibility", "leadership activities",
        "leadership and extracurricular",
        "leadership and involvement",
        "committee experience", "committee positions",
        "board experience", "board positions",
        "officer positions", "club officer",
        "leadership summary", "leadership and service",
        "management and leadership",
    ],

    # ── Extracurricular / activities / memberships ────────────────────────────
    "extracurricular": [
        "extracurricular activities", "extracurricular",
        "activities", "co curricular activities", "co curricular",
        "college activities", "campus activities",
        "student activities", "club activities",
        "student organizations", "organizations",
        "clubs and organizations", "clubs and activities",
        "university activities",
        "campus involvement", "student involvement",
        "affiliations", "professional affiliations",
        "memberships", "professional memberships",
        "association memberships", "associations",
        "clubs", "society memberships", "societies",
        "professional associations",
    ],

    # ── Languages ─────────────────────────────────────────────────────────────
    "languages": [
        "languages", "language", "language skills",
        "language proficiency", "linguistic skills",
        "languages known", "languages spoken",
        "known languages", "spoken languages",
        "communication languages",
        "foreign languages", "second languages",
        "bilingual", "multilingual",
        "natural languages", "language abilities",
        "language competencies",
    ],

    # ── Interests / hobbies ───────────────────────────────────────────────────
    "interests": [
        "interests", "areas of interest",
        "professional interests", "career interests",
        "personal interests",
        "interests and hobbies", "hobbies and interests",
        "personal interests and hobbies",
        "hobbies and activities",
        "hobbies", "hobby", "hobbies and pastimes",
        "pastimes", "activities and hobbies",
        "leisure activities", "leisure interests",
        "personal activities", "other interests",
        "recreational activities",
    ],

    # ── References ────────────────────────────────────────────────────────────
    "references": [
        "references", "reference", "professional references",
        "character references", "personal references",
        "referees", "referee", "references available",
        "references available upon request",
        "available upon request",
        "references upon request",
        "references furnished upon request",
        "reference details",
        "testimonials",
    ],

    # ── Contact information ───────────────────────────────────────────────────
    "contact": [
        "contact", "contact information", "contact details",
        "contact info", "personal details", "personal information",
        "personal data", "contact data",
        "personal contact", "address",
        "contact and personal details",
        "basic information", "basic details",
        "personal particulars", "particulars",
        "general information", "candidate information",
        "applicant information",
    ],

    # ── Social / online presence ──────────────────────────────────────────────
    "social_links": [
        "social links", "social media", "social profiles",
        "online profiles", "online presence",
        "digital profiles", "web presence",
        "linkedin", "github", "online portfolio",
        "digital portfolio", "portfolio links",
        "professional links", "professional profiles",
        "links", "websites", "website",
        "web profiles", "internet profiles",
        "social media profiles",
    ],

    # ── Career highlights ─────────────────────────────────────────────────────
    "career_highlights": [
        "career highlights", "highlights",
        "career accomplishments", "professional highlights",
        "core achievements", "top achievements",
        "key contributions", "key results",
        "significant contributions", "performance summary",
        "executive highlights",
        "key career highlights", "notable highlights",
        "career summary highlights",
    ],
}

# The 6 sections shown in the UI / stored in the DB.
_CORE_SECTIONS = ("summary", "skills", "experience", "education", "certifications", "projects")

# Maps every extended section to a core bucket (None = discard).
_SECTION_FOLD = {
    "internships":       "experience",
    "training":          "certifications",
    "licenses":          "certifications",
    "achievements":      "certifications",
    "awards":            "certifications",
    "publications":      "projects",
    "research":          "experience",
    "patents":           "certifications",
    "conferences":       "certifications",
    "volunteer":         "experience",
    "leadership":        "experience",
    "extracurricular":   None,
    "languages":         "skills",
    "interests":         None,
    "references":        None,
    "contact":           None,
    "social_links":      None,
    "career_highlights": "summary",
}

# Sections that are always discarded (fold to nothing) — e.g. "Interests"/"Hobbies",
# "References", "Contact". find_sections() treats passing through one of these as
# noise, not as evidence that an interrupted section (mid numbered-list) is done.
_THROWAWAY_SECTIONS = frozenset(k for k, v in _SECTION_FOLD.items() if v is None)

# Resume-builder-template watermarks (e.g. "webuildcv.com", "Powered by ...") that
# templates place as a bare branding line/paragraph, which then gets swept up
# verbatim into whichever section it happens to fall in. The two fragments
# sometimes appear combined on a single line (e.g. table-cell layout joins them
# with a tab), so match any whole line made up of nothing but 1-3 repetitions of
# either fragment rather than requiring each to be alone on its own line.
_WATERMARK_LINE_RE = re.compile(
    r'^\s*(?:(?:powered\s+by\b\s*)|(?:[\w-]+\.(?:com|io|co|net|ai)\b\s*)){1,3}$',
    re.I,
)

# ── DB context manager ────────────────────────────────────────────────────────

class _PgConn:
    """Thin wrapper that gives a psycopg2 connection a sqlite3-style conn.execute() API."""

    def __init__(self, raw):
        self._raw = raw

    def execute(self, sql, params=None):
        cur = self._raw.cursor(cursor_factory=RealDictCursor)
        cur.execute(sql, params if params is not None else ())
        return cur

    def commit(self):
        self._raw.commit()

    def rollback(self):
        self._raw.rollback()

    def close(self):
        self._raw.close()


@contextmanager
def db_conn():
    raw = psycopg2.connect(DATABASE_URL)
    conn = _PgConn(raw)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ── Helpers ───────────────────────────────────────────────────────────────────

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def slugify(text):
    text = (text or "profile").lower().strip()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-") or "profile"


def unique_slug(conn, base_slug, resume_id=None):
    slug = slugify(base_slug)
    candidate = slug
    counter = 2
    while True:
        if resume_id:
            existing = conn.execute(
                "SELECT id FROM resume WHERE slug = %s AND id != %s", (candidate, resume_id)
            ).fetchone()
        else:
            existing = conn.execute("SELECT id FROM resume WHERE slug = %s", (candidate,)).fetchone()
        if not existing:
            return candidate
        candidate = f"{slug}-{counter}"
        counter += 1


def sync_skills(conn, resume_id, skills_text):
    """Populate the normalised resume_skill table from the raw skills text blob."""
    conn.execute("DELETE FROM resume_skill WHERE resume_id = %s", (resume_id,))
    if not skills_text:
        return
    seen = set()
    for line in skills_text.splitlines():
        for part in re.split(r"[,|•\n]", line):
            skill = part.strip().strip("•- ")
            if skill and len(skill) > 1 and "://" not in skill and skill.lower() not in seen:
                seen.add(skill.lower())
                conn.execute(
                    "INSERT INTO resume_skill (resume_id, skill) VALUES (%s, %s)", (resume_id, skill)
                )


# ── PDF / DOCX text extraction ────────────────────────────────────────────────

# Cached EasyOCR reader — initialised once, reused on every subsequent call.
_ocr_reader = None

def _get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        logger.info("Initialising EasyOCR reader (first call only)…")
        _ocr_reader = easyocr.Reader(["en"], gpu=False)
    return _ocr_reader


def _ocr_contact_strip(path):
    """OCR only the top 25 % of page 1 to recover email/phone from icon headers.
    Returns a short string with whatever text EasyOCR finds in that strip.
    """
    try:
        import pymupdf as fitz
        import numpy as np
        doc = fitz.open(str(path))
        pg  = doc[0]
        # Clip to top quarter of the page (where contact bars live)
        clip = fitz.Rect(0, 0, pg.rect.width, pg.rect.height * 0.25)
        mat  = fitz.Matrix(150 / 72, 150 / 72)   # 150 DPI — fast enough for OCR
        pix  = pg.get_pixmap(matrix=mat, clip=clip)
        img  = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        doc.close()
        reader  = _get_ocr_reader()
        results = reader.readtext(img)
        return "\n".join(t for (_, t, c) in results if c > 0.2)
    except Exception as e:
        logger.debug(f"OCR contact strip failed: {e}")
        return ""


def _detect_two_col_split(page):
    """Return the x midpoint of the gap between two column clusters, or None.

    Uses block bounding boxes (coarser but reliable) to detect whether the page
    content clearly lives in two separate vertical strips.  Conditions:
      - ≥ 3 blocks whose right edge is left of the midpoint  (left column)
      - ≥ 3 blocks whose left  edge is right of the midpoint (right column)
      - ≤ 2 blocks that span across the midpoint (title / contact header)
      - A real gap exists: min x0 of right blocks > max x1 of left blocks
    """
    blocks = page.get_text("blocks")
    pw  = page.rect.width
    mid = pw / 2
    margin = pw * 0.08   # 8 % — tolerates slight layout asymmetry

    left     = [b for b in blocks if b[2] < mid - margin]
    right    = [b for b in blocks if b[0] > mid + margin]
    spanning = [b for b in blocks if b[0] < mid - margin and b[2] > mid + margin]

    if len(left) < 3 or len(right) < 3 or len(spanning) > 2:
        return None

    max_left_x1  = max(b[2] for b in left)
    min_right_x0 = min(b[0] for b in right)
    if min_right_x0 <= max_left_x1:
        return None   # columns overlap — not a clean 2-col layout

    return (max_left_x1 + min_right_x0) / 2


def _detect_col_gutter_words(page):
    """Find a two-column split via a vertical word gutter, or None.

    Block-based detection (`_detect_two_col_split`) misses layouts where wide
    bullet lines in the left column extend past the page midpoint — the blocks
    "span" the centre and the heuristic bails.  This word-level pass instead looks
    for a vertical strip in the middle of the page that almost no word box crosses:
    a true column gutter.  Robust because a single-column page has lines spanning
    the centre, so every candidate split is crossed by many words.
    """
    words = page.get_text("words")
    if len(words) < 40:
        return None

    pw = page.rect.width
    lo, hi = pw * 0.33, pw * 0.67          # only hunt for a gutter in the middle third
    n = len(words)

    best_x, best_cross = None, None
    x = lo
    while x <= hi:
        cross = sum(1 for w in words if w[0] < x < w[2])
        if best_cross is None or cross < best_cross:
            best_cross, best_x = cross, x
        x += 2.0

    if best_x is None:
        return None
    # The gutter must be almost empty (≤1 % of words straddle it).
    if best_cross > max(2, n * 0.01):
        return None
    # Both columns must hold a real share of the content. A narrow sidebar
    # (short skill/contact tags, or just a couple of section headings on an
    # otherwise sparse page) legitimately holds far fewer words than a wide
    # prose column even when the split is completely genuine — a flat 15%
    # share was rejecting real sidebars that clear a clean, zero-crossing
    # gutter (the strongest signal of a true column boundary) purely for
    # being word-sparse, and a flat percentage doesn't scale between a dense
    # page and a mostly-empty one anyway. A small absolute floor does the
    # same job (filters out a single stray word/artifact) without penalizing
    # sparse sidebars on either dense or sparse pages.
    left  = sum(1 for w in words if w[2] <= best_x)
    right = sum(1 for w in words if w[0] >= best_x)
    min_side = max(15, n * 0.05)
    if left < min_side or right < min_side:
        return None

    # Final guard: in a true 2-column layout each line lives mostly in one
    # column.  On a single-column page the "gutter" is just a ragged
    # right-margin gap — most lines still have words on BOTH sides.
    # If >25 % of y-level buckets have words on both sides, reject the split.
    _y_tol = 6.0
    _y_sides = {}
    for w in words:
        b = round(w[1] / _y_tol)
        prev = _y_sides.get(b, [False, False])
        _y_sides[b] = [prev[0] or w[2] <= best_x, prev[1] or w[0] >= best_x]
    _dual = sum(1 for v in _y_sides.values() if v[0] and v[1])
    if _dual > 0.25 * len(_y_sides):
        return None

    return best_x


def _words_to_text(raw_words):
    """Convert a flat list of PyMuPDF word tuples to a text string.

    Words on the same y-level (±5 pt) form a row; wide horizontal gaps (> 30 pt)
    within a row indicate sub-columns (e.g. company | role | dates).
    Section headings are always emitted on their own line so the section parser
    can pick them up reliably.
    """
    if not raw_words:
        return ""

    raw_words = sorted(raw_words, key=lambda w: w[1])   # top-to-bottom

    Y_TOL   = 5
    COL_GAP = 30

    def emit_row(row):
        if not row:
            return []
        row.sort()
        groups = [[row[0]]]
        for i in range(1, len(row)):
            if row[i][0] - row[i - 1][1] > COL_GAP:
                groups.append([])
            groups[-1].append(row[i])

        if len(groups) == 1:
            return [" ".join(wd for _, _, wd in groups[0])]

        gtexts = [" ".join(wd for _, _, wd in g) for g in groups]

        if canonical_section_name(gtexts[0]):
            return [gtexts[0]] + gtexts[1:]

        sec_idxs = [i for i in range(1, len(groups)) if canonical_section_name(gtexts[i])]
        if sec_idxs:
            sec_set = set(sec_idxs)
            non_sec = []
            for i, g in enumerate(groups):
                if i not in sec_set:
                    non_sec.extend(g)
            out = [gtexts[i] for i in sec_idxs]
            if non_sec:
                non_sec.sort()
                out.insert(0, " ".join(wd for _, _, wd in non_sec))
            return out

        all_words = [w for g in groups for w in g]
        all_words.sort()
        return [" ".join(wd for _, _, wd in all_words)]

    lines = []
    cur_row, cur_y = [], None
    for w in raw_words:
        x0, y0, x1, y1, word = w[0], w[1], w[2], w[3], w[4]
        if not word.strip():
            continue
        if cur_y is None or abs(y0 - cur_y) <= Y_TOL:
            cur_row.append((x0, x1, word))
            cur_y = cur_y if cur_y is not None else y0
        else:
            lines.extend(emit_row(cur_row))
            cur_row, cur_y = [(x0, x1, word)], y0
    lines.extend(emit_row(cur_row))
    return "\n".join(lines)


def _pdfplumber_two_col_text(path):
    """Extract text from a sidebar-style 2-column PDF using pdfplumber crops.

    For each page:
      1. Collect all word bounding boxes via extract_words().
      2. Find the X position in the 20-65% range where the fewest word boxes
         cross — this is the column gutter.
      3. Crop the page at that X, extract each side independently.
      4. Concatenate left-column text then right-column text so section
         headings (KEY SKILLS on left, WORK EXPERIENCE on right) appear in the
         correct order for find_sections().

    Returns the combined multi-page string, or None if pdfplumber is
    unavailable, the file is not a PDF, or no 2-column layout is detected.
    """
    if not HAS_PDFPLUMBER:
        return None
    try:
        import pdfplumber
        pages_text = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                # extract_words() with no extra kwargs — compatible with all versions
                try:
                    words = page.extract_words()
                except Exception:
                    pages_text.append(page.extract_text() or "")
                    continue

                if not words or len(words) < 15:
                    pages_text.append(page.extract_text() or "")
                    continue

                pw = float(page.width)
                ph = float(page.height)
                n  = len(words)

                # Scan x from 20 % to 65 % of page width for the column gutter
                lo, hi = pw * 0.20, pw * 0.65
                best_x, best_cross = None, n + 1
                x = lo
                while x <= hi:
                    cross = sum(
                        1 for w in words
                        if float(w.get('x0', 0)) < x < float(w.get('x1', 0))
                    )
                    if cross < best_cross:
                        best_cross, best_x = cross, x
                    x += 2.0

                # Allow up to 5% of words to straddle the gutter.
                # Raise tolerance to 10% when this looks like a sidebar layout
                # (right column has 2x+ more words than left): centered headings
                # in the right column can straddle a nearby gutter.
                _left_est  = sum(1 for w in words if float(w.get('x1', 0)) <= best_x)
                _right_est = sum(1 for w in words if float(w.get('x0', 0)) >= best_x)
                _cross_limit = (max(8, n * 0.10)
                                if _right_est > 2 * max(_left_est, 1)
                                else max(4, n * 0.05))
                logger.debug(
                    f"pdfplumber gutter scan: best_x={best_x:.1f} best_cross={best_cross} "
                    f"limit={_cross_limit:.1f} left_est={_left_est} right_est={_right_est}"
                )
                if best_x is None or best_cross > _cross_limit:
                    logger.debug("pdfplumber gutter rejected by cross-count — using extract_text()")
                    pages_text.append(page.extract_text() or "")
                    continue

                # Right column must be substantial (>=8%). Left can be as thin as 3%
                # because on page 2+ of sidebar PDFs the sidebar shrinks to just a few
                # personal-detail lines while the main content fills the right column.
                left_n  = sum(1 for w in words if float(w.get('x1', 0)) <= best_x)
                right_n = sum(1 for w in words if float(w.get('x0', 0)) >= best_x)
                if left_n < n * 0.03 or right_n < n * 0.08:
                    pages_text.append(page.extract_text() or "")
                    continue

                # Final guard: in a true 2-column layout each y-level has words on
                # one side only. On a single-column page most lines span both sides
                # (ragged right margin). If >40% of y-buckets are dual-side → single column.
                # Threshold is 40% (not 25%) because sidebar+main layouts with a photo
                # at the top-left have the top half populated only on the right, while
                # the bottom half has both columns — this legitimately raises dual-Y to ~30%.
                _y_tol = 6.0
                _y_sides = {}
                for w in words:
                    b = round(float(w.get('top', 0)) / _y_tol)
                    prev = _y_sides.get(b, [False, False])
                    _y_sides[b] = [prev[0] or float(w.get('x1', 0)) <= best_x,
                                   prev[1] or float(w.get('x0', 0)) >= best_x]
                _dual = sum(1 for v in _y_sides.values() if v[0] and v[1])
                if _y_sides and _dual > 0.40 * len(_y_sides):
                    pages_text.append(page.extract_text() or "")
                    continue

                # Crop each column and extract text in reading order
                try:
                    left_text  = (page.crop((0,       0, best_x, ph)).extract_text() or "").strip()
                    right_text = (page.crop((best_x,  0, pw,     ph)).extract_text() or "").strip()
                except Exception:
                    pages_text.append(page.extract_text() or "")
                    continue

                # Post-crop sanity: character-level word splits (caused by a false gutter
                # cutting through the middle of full-width lines) produce fragments like
                # "agement", "ifecycle", "anagement" that start with VOWELS (because the
                # split falls at a consonant-vowel boundary within polysyllabic words).
                # Legitimate right-column continuation lines start with full English words
                # which predominantly start with CONSONANTS ("medical", "marking", etc.).
                # If ≥2 lowercase-starting right-column lines exist and ≥60% of them start
                # with a vowel, the gutter is a false split — fall back to extract_text().
                _right_lines = [l for l in right_text.splitlines() if l.strip()]
                _lower_first_chars = [
                    l.strip()[0] for l in _right_lines[:12]
                    if l.strip() and l.strip()[0].islower()
                ]
                _vowel_frag_starts = sum(1 for c in _lower_first_chars if c in 'aeiou')
                if (len(_lower_first_chars) >= 2
                        and _vowel_frag_starts >= len(_lower_first_chars) * 0.6):
                    logger.debug(
                        "pdfplumber: right column has %d/%d vowel-starting lowercase lines "
                        "— word fragments from false 2-col split, using extract_text()",
                        _vowel_frag_starts, len(_lower_first_chars)
                    )
                    pages_text.append(page.extract_text() or "")
                    continue

                logger.info(f"pdfplumber 2-col split at x={best_x:.1f} "
                            f"(left={left_n} words, right={right_n} words)")

                # If the left column contains only sidebar-metadata headings (personal
                # details, hobbies, languages, extra-curricular) and no valuable resume
                # sections, emit ONLY the right column. This prevents continuation content
                # on the right (WE bullets, PROJECTS) from being broken by left-column
                # skip-headings that reset find_sections() state to None.
                _SIDEBAR_ONLY_RE = re.compile(
                    r'^(?:other personal details|personal information|personal details|'
                    r'hobbies?|interests?|extra\s*curricular|declaration|references|languages?)',
                    re.I
                )
                _VALUABLE_SECTION_RE = re.compile(
                    r'\b(?:skills|experience|education|projects|certifications|summary|profile)\b',
                    re.I
                )
                left_first_line = left_text.split('\n')[0].strip() if left_text else ""
                _left_is_sidebar_only = (
                    bool(_SIDEBAR_ONLY_RE.match(left_first_line))
                    and not _VALUABLE_SECTION_RE.search(left_text)
                )

                if _left_is_sidebar_only:
                    # Drop the left-column sidebar: its section headings (HOBBIES, LANGUAGES,
                    # etc.) would corrupt section state and cause right-column continuation
                    # content to be discarded by find_sections().
                    logger.info(f"Page left column is sidebar-only — emitting right column only")
                    pages_text.append(right_text)
                else:
                    # Detect label+content layout: narrow label column (SUMMARY / EXPERIENCE /
                    # SKILLS headings only) + wide content column.  If we concatenate
                    # left+right all headings appear before all content, breaking find_sections.
                    # Instead, reconstruct in true reading order (row-interleaved) so each
                    # heading immediately precedes its paragraph.
                    _left_lines  = [l.strip() for l in left_text.split('\n')  if l.strip()]
                    _right_lines = [l.strip() for l in right_text.split('\n') if l.strip()]
                    _avg_left_len  = (sum(len(l) for l in _left_lines)  / len(_left_lines))  if _left_lines  else 999
                    _short_left    = (sum(1 for l in _left_lines if len(l.split()) <= 3) / len(_left_lines)) if _left_lines else 0
                    _line_ratio    = len(_right_lines) / max(len(_left_lines), 1)
                    # word-count ratio (right_n/left_n) is more reliable than line ratio
                    # when pdfplumber compresses right-column text into few lines
                    _word_ratio    = right_n / max(left_n, 1)
                    # heading fraction: label+content layouts have ONLY section headings on the
                    # left (SKILLS / EDUCATION / etc.); content-sidebar layouts have actual data
                    # (email, phone, skill items) mixed in.  Require ≥30% of left lines to be
                    # recognised headings so we don't y-pair content-sidebar pages.
                    _left_hdr_cnt = sum(1 for l in _left_lines if canonical_section_name(l) is not None)
                    _left_heading_frac = _left_hdr_cnt / max(len(_left_lines), 1)
                    _is_label_layout = (
                        _avg_left_len < 25 and _short_left > 0.70
                        and (_line_ratio > 3 or _word_ratio > 5)
                        and _left_heading_frac >= 0.30
                    )
                    logger.debug(
                        f"label+content check: avg_left={_avg_left_len:.1f} "
                        f"short_left={_short_left:.2f} line_ratio={_line_ratio:.1f} "
                        f"word_ratio={_word_ratio:.1f} hdr_frac={_left_heading_frac:.2f} "
                        f"→ {_is_label_layout}"
                    )

                    if _is_label_layout:
                        # Label+content layout: pair each right-column content line with
                        # the heading in the left column that sits closest above it by y.
                        # Uses per-crop word coordinates (more reliable than full-page
                        # words which can have PDF-stream ordering artefacts).
                        logger.info("Label+content layout — crop-based y-pairing")
                        try:
                            _lw = page.crop((0,      0, best_x, ph)).extract_words() or []
                            _rw = page.crop((best_x, 0, pw,     ph)).extract_words() or []
                            if not _lw or not _rw:
                                raise ValueError("empty crop words")

                            # Build left-column lines: wy → text
                            _LYT = 4.0
                            _lrows = {}
                            for w in _lw:
                                wy = round(float(w.get('top', 0)) / _LYT)
                                _lrows.setdefault(wy, []).append(
                                    (float(w.get('x0', 0)), w.get('text', '')))
                            _l_lines = sorted(
                                [(wy * _LYT,
                                  ' '.join(t[1] for t in sorted(v, key=lambda t: t[0])))
                                 for wy, v in _lrows.items()],
                                key=lambda x: x[0])
                            _l_lines = [(y, t) for y, t in _l_lines if t.strip()]

                            # Build right-column lines: wy → text
                            _RYT = 3.0
                            _rrows = {}
                            for w in _rw:
                                wy = round(float(w.get('top', 0)) / _RYT)
                                _rrows.setdefault(wy, []).append(
                                    (float(w.get('x0', 0)), w.get('text', '')))
                            _r_lines = sorted(
                                [(wy * _RYT,
                                  ' '.join(t[1] for t in sorted(v, key=lambda t: t[0])))
                                 for wy, v in _rrows.items()],
                                key=lambda x: x[0])
                            _r_lines = [(y, t) for y, t in _r_lines if t.strip()]

                            # Pair each right line to the nearest left heading above it.
                            # Tolerance: heading may start up to 15pt above first content line.
                            _PTOL   = 15.0
                            _l_ys   = [y for y, _ in _l_lines]
                            _segs   = {y: [] for y, _ in _l_lines}
                            _pre    = []  # content before first heading (e.g. experience continuation)

                            for r_y, r_txt in _r_lines:
                                _bhy = None
                                for l_y in reversed(_l_ys):
                                    if l_y <= r_y + _PTOL:
                                        _bhy = l_y
                                        break
                                (_segs[_bhy] if _bhy is not None else _pre).append(r_txt)

                            _out = _pre[:]
                            for l_y, l_txt in _l_lines:
                                _out.append(l_txt)
                                _out.extend(_segs[l_y])

                            pages_text.append(
                                '\n'.join(l for l in _out if l.strip()))

                        except Exception as _exc:
                            logger.debug(f"label+content crop-pair failed: {_exc}")
                            pages_text.append(left_text + "\n" + right_text)
                    else:
                        # Strip the right-column contact/name header (all lines before the
                        # first section heading) to prevent those fragments from bleeding
                        # into adjacent left-column section verbatim during extraction.
                        _rt_lines = right_text.splitlines()
                        _rt_sec_idx = None
                        for _si, _sl in enumerate(_rt_lines):
                            if canonical_section_name(_sl.strip()):
                                _rt_sec_idx = _si
                                break
                        if _rt_sec_idx is not None and _rt_sec_idx > 0:
                            _right_hdr_block = [l.strip() for l in _rt_lines[:_rt_sec_idx] if l.strip()]
                            _right_clean = "\n".join(_rt_lines[_rt_sec_idx:])
                            logger.debug("Right col header stripped: %d lines before '%s'",
                                         len(_right_hdr_block), _rt_lines[_rt_sec_idx].strip())
                        else:
                            _right_hdr_block = []
                            _right_clean = right_text

                        # Content-sidebar layout: left column has sidebar data (personal info +
                        # skills), right column has the main content (name, title, summary…).
                        # Prepend the right column's header lines (name, title — everything
                        # before the first section heading) so that parse_resume_text sees
                        # the candidate's name and title at the very top of the document,
                        # even though the sidebar content is much longer.
                        # Only prepend the right-column header when the right column
                        # genuinely starts with a person's name (i.e. this is a true
                        # name-on-right sidebar layout, not a table split).
                        _rh_lines = []
                        _first_right = next(
                            (_rl.strip() for _rl in right_text.splitlines() if _rl.strip()), ""
                        )
                        if _looks_like_name(_first_right):
                            for _rl in _right_hdr_block:
                                if not _rl:
                                    continue
                                _rh_lines.append(_rl)
                                if len(_rh_lines) >= 3:  # name, title, optional tagline
                                    break
                        if _rh_lines:
                            pages_text.append(
                                "\n".join(_rh_lines) + "\n" + left_text + "\n" + _right_clean
                            )
                        else:
                            pages_text.append(left_text + "\n" + _right_clean)

        combined = "\n".join(pages_text).strip()
        return combined if combined else None
    except Exception as e:
        logger.debug(f"_pdfplumber_two_col_text failed: {e}")
        return None


def _chunk_lines_by_gap(lines):
    """lines: list of (y0, word_list) sorted by y0 ascending, one entry per
    rendered line within a single column. Splits into row groups wherever
    the vertical gap to the next line is notably larger than the typical
    (median) gap in this column — the signature of a real row/paragraph
    boundary (one project block ending, the next beginning) as opposed to
    ordinary line-wrap spacing within the same block. Returns a list of
    (y_lo, y_hi, word_list) tuples, one per row chunk, in top-to-bottom
    order.
    """
    if not lines:
        return []
    if len(lines) == 1:
        y0, words = lines[0]
        return [(y0, y0, list(words))]

    gaps = [lines[i + 1][0] - lines[i][0] for i in range(len(lines) - 1)]
    positive_gaps = sorted(g for g in gaps if g > 0) or [12]
    median_gap = positive_gaps[len(positive_gaps) // 2]
    threshold = max(median_gap * 1.6, median_gap + 6)

    chunk_words = list(lines[0][1])
    chunk_lo = chunk_hi = lines[0][0]
    chunks = []
    for i in range(1, len(lines)):
        y0, words = lines[i]
        if y0 - chunk_hi > threshold:
            chunks.append((chunk_lo, chunk_hi, chunk_words))
            chunk_words, chunk_lo = [], y0
        chunk_words = chunk_words + list(words)
        chunk_hi = y0
    chunks.append((chunk_lo, chunk_hi, chunk_words))
    return chunks


_LEFT_ROW_START_RE = re.compile(r"^(client|company|employer|account|customer)\s*:", re.I)


def _split_left_lines_by_row_marker(lines):
    """Split left-column lines into row chunks anchored on an explicit
    row-opener label ("Client:", "Company:", etc.) instead of vertical gap
    size. In a consulting-style "Client / Project / Description" table, the
    gap between "Client:" and "Project N:" (still the SAME row) is often
    identical to the gap between one row's last description line and the
    next row's "Client:" line — both are just one paragraph-height apart —
    so _chunk_lines_by_gap alone can't tell "still this row" from "new row"
    apart (confirmed live: it collapsed 5 distinct projects into a single
    chunk on a real resume for exactly this reason). An explicit repeating
    label has no such ambiguity. Returns None if fewer than 2 such markers
    are found, so callers can fall back to gap-based chunking for tables
    that don't use this convention.
    """
    marker_idxs = [
        i for i, (_, words) in enumerate(lines)
        if _LEFT_ROW_START_RE.match(_words_to_text(words).strip())
    ]
    if len(marker_idxs) < 2:
        return None

    chunks = []
    for k, start in enumerate(marker_idxs):
        end = marker_idxs[k + 1] if k + 1 < len(marker_idxs) else len(lines)
        row_lines = lines[start:end]
        y_lo, y_hi = row_lines[0][0], row_lines[-1][0]
        words = [w for _, ws in row_lines for w in ws]
        chunks.append((y_lo, y_hi, words))

    # Anything before the first marker (e.g. a "Relevant Project Experience:"
    # heading, or the "Project Details" column header) becomes its own
    # leading chunk rather than being dropped.
    if marker_idxs[0] > 0:
        pre_lines = lines[:marker_idxs[0]]
        y_lo, y_hi = pre_lines[0][0], pre_lines[-1][0]
        pre_words = [w for _, ws in pre_lines for w in ws]
        chunks.insert(0, (y_lo, y_hi, pre_words))

    return chunks


def _pymupdf_page_text(page):
    """Extract readable text from a page, handling 2-column layouts correctly.

    If the page has a clear 2-column structure (detected via block positions),
    the left and right columns are processed independently, ROW BY ROW, so
    each row's right-column content (e.g. "Role and Responsibilities") is
    emitted immediately after that SAME row's left-column content (e.g. the
    matching "Project Details" entry) — not as two giant blobs (whole left
    column, then whole right column) which reads as if every project's
    responsibilities were missing until the very end. Falls back to
    single-stream extraction for 1-column pages.
    """
    raw_words = page.get_text("words")
    if not raw_words:
        return ""

    # Primary: conservative block-based detection. Fallback: word-gutter detection
    # for interleaved layouts where wide left-column lines span the page centre.
    split_x = _detect_two_col_split(page) or _detect_col_gutter_words(page)
    if split_x:
        # Classify by PyMuPDF's own (block, line) grouping — the line units
        # the PDF's content stream actually rendered — not by each word's own
        # x-position. Per-word classification broke real 2-column
        # project/responsibilities tables: a left-column paragraph line that
        # happens to wrap close to the column's natural right margin has its
        # LAST word occasionally drift past split_x, so that one word alone
        # got shaved off into the right-column bucket, scrambling the
        # sentence (confirmed live: "...providing greater visibility and
        # control" lost "control" this way, repeated across every project
        # description on the page). Grouping by (block, line) keeps an
        # entire rendered line together and classifies it once, by its own
        # leftmost word — while two column headers that legitimately sit at
        # the identical y (e.g. "Project Details" / "Role and
        # Responsibilities") are already separate (block, line) pairs, so
        # they still split apart correctly.
        line_groups = {}
        for w in raw_words:
            line_groups.setdefault((w[5], w[6]), []).append(w)

        left_lines, right_lines = [], []  # each entry: (y0, word_list)
        for line_words in line_groups.values():
            y0 = min(w[1] for w in line_words)
            x0 = min(w[0] for w in line_words)
            (left_lines if x0 < split_x else right_lines).append((y0, line_words))

        left_lines.sort(key=lambda t: t[0])
        right_lines.sort(key=lambda t: t[0])

        # Row-by-row interleaving only kicks in for a confirmed repeating
        # table (an explicit "Client:"-style marker repeated ≥2 times) —
        # NOT for every 2-column page. A generic 2-column layout (e.g. main
        # content + a Skills/Contact sidebar that runs independently down
        # the whole page) isn't a row-paired table at all, and gap-based
        # chunking + y-overlap pairing would wrongly interleave sidebar
        # paragraphs into the middle of unrelated main-column content.
        # Falling back to the plain "whole left column, then whole right
        # column" concatenation is what already correctly handles that case
        # (confirmed against a real sidebar-style resume).
        left_chunks = _split_left_lines_by_row_marker(left_lines)
        if not left_chunks:
            left_words = [w for _, words in left_lines for w in words]
            right_words = [w for _, words in right_lines for w in words]
            return (_words_to_text(left_words) + "\n" + _words_to_text(right_words)).strip()

        right_chunks = _chunk_lines_by_gap(right_lines)
        if not right_chunks:
            return _words_to_text([w for _, words in left_lines for w in words])

        # Assign each right-column row to the left-column row it physically
        # sits beside — by GREATEST y-range overlap (not just the first
        # match), or, failing any overlap, the nearest left row whose range
        # starts at or before it. Index-pairing left[i]/right[i] would
        # misalign the instant either side has an extra chunk (e.g. a lone
        # column-header line forms its own tiny chunk on one side but
        # merges into the first row on the other, purely from gap-size
        # rounding) — matching by actual vertical position is what makes
        # this robust regardless of how many chunks either side ends up
        # with. Picking the FIRST overlapping chunk (rather than the best)
        # broke a page-break row: the right column's header merged with
        # that row's own first bullet into one wide chunk that technically
        # brushed the header's tiny left chunk too, and "first" wrongly
        # picked the near-zero-width header overlap over the true row.
        assigned = [[] for _ in left_chunks]
        for r_lo, r_hi, r_words in right_chunks:
            best_i, best_overlap = None, -1
            for i, (l_lo, l_hi, _) in enumerate(left_chunks):
                overlap = min(l_hi, r_hi) - max(l_lo, r_lo)
                if overlap >= 0 and overlap > best_overlap:
                    best_overlap = overlap
                    best_i = i
            if best_i is None:
                preceding = [i for i, (l_lo, _, _) in enumerate(left_chunks) if l_lo <= r_lo]
                best_i = preceding[-1] if preceding else 0
            assigned[best_i].append(r_words)

        out_parts = []
        for i, (_, _, l_words) in enumerate(left_chunks):
            out_parts.append(_words_to_text(l_words))
            for r_words in assigned[i]:
                out_parts.append(_words_to_text(r_words))
        return "\n".join(p for p in out_parts if p).strip()

    return _words_to_text(raw_words)


def _fix_wrapped_email(text):
    """Join emails that the PDF wrapped mid-TLD or mid-local-part across two lines.

    Case 1 (post-@ split):  "sindhusundaramoorthy30@gmail.c"  +  "om"
    Case 2 (pre-@ split):   "manibharathi601rav"  +  "i@gmail.com"
    """
    lines = text.split("\n")
    result = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if i + 1 < len(lines) and "@" in line:
            next_line = lines[i + 1].strip()
            # Case 1: current line ends with partial TLD (≤2 chars), next line is TLD suffix
            if re.match(r"^[a-zA-Z]{1,4}$", next_line):
                if re.search(r"@[A-Za-z0-9.\-]+\.[A-Za-z]{1,2}$", line):
                    line = line.rstrip() + next_line
                    i += 1
        elif i + 1 < len(lines) and "@" not in line:
            next_line = lines[i + 1].strip()
            # Case 2: next line starts with a short prefix (≤5 chars) then @domain.tld
            # and current line ends with 5+ alphanum chars (looks like partial local part)
            m = re.match(r"^([A-Za-z0-9]{1,5})(@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", next_line)
            if m and re.search(r"[A-Za-z0-9]{5,}$", line):
                line = line.rstrip() + next_line
                i += 1
        result.append(line)
        i += 1
    return "\n".join(result)


def extract_text_from_pdf(path):
    """Extract text from PDF.

    Order: pymupdf (fastest) → pdfplumber → PyPDF2.
    If email or phone are still missing after text extraction, a targeted OCR
    pass runs on the top 25 % of page 1 only (icon-header contact bars).
    """
    try:
        logger.info(f"Extracting text from PDF: {path}")
        text = ""

        # ── Primary: pymupdf ──────────────────────────────────────────────────
        if HAS_PYMUPDF:
            try:
                import pymupdf as fitz
                doc = fitz.open(str(path))
                parts = [_pymupdf_page_text(page) for page in doc]
                doc.close()
                text = "\n".join(parts)
                logger.info(f"pymupdf extracted {len(parts)} pages")
            except Exception as e:
                logger.debug(f"pymupdf text extract failed: {e}")

        # ── Fallback 1: pdfplumber ────────────────────────────────────────────
        if not text.strip() and HAS_PDFPLUMBER:
            try:
                import pdfplumber
                with pdfplumber.open(str(path)) as pdf:
                    parts = [p for page in pdf.pages if (p := page.extract_text())]
                    if parts:
                        text = "\n".join(parts)
                        logger.info(f"pdfplumber extracted {len(pdf.pages)} pages")
            except Exception as e:
                logger.debug(f"pdfplumber failed: {e}")

        # ── Fallback 2: PyPDF2 ────────────────────────────────────────────────
        if not text.strip():
            try:
                reader = PdfReader(str(path))
                if not reader.pages:
                    raise ValueError("PDF has no pages")
                text = "\n".join(page.extract_text() or "" for page in reader.pages)
                logger.info(f"PyPDF2 extracted {len(reader.pages)} pages")
            except Exception as e:
                logger.debug(f"PyPDF2 failed: {e}")

        if not text.strip():
            raise ValueError("Could not extract text from PDF using any method")

        # Decode CID-encoded characters. When a PDF uses a custom font without
        # a ToUnicode table, pymupdf emits glyphs as "(cid:N)". For standard
        # Latin fonts (WinAnsiEncoding / MacRomanEncoding), glyph index N
        # directly matches the ASCII/Unicode code point, so chr(N) recovers the
        # actual character. CIDs outside printable ASCII (e.g. bullet markers
        # like cid:127) are removed.
        def _cid_to_char(m):
            n = int(m.group(1))
            return chr(n) if 33 <= n <= 126 else ''
        text = re.sub(r'\(cid:(\d+)\)', _cid_to_char, text)

        # ── Targeted OCR for icon-header contact info ─────────────────────────
        # Only runs when email or phone are absent AND EasyOCR is installed.
        has_email = bool(re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text))
        has_phone_m = re.search(r"[\+]?\(?\d[\d\s\-\.\(\)]{7,20}\d", text)
        has_phone = bool(has_phone_m and len(re.sub(r"\D", "", has_phone_m.group())) >= 10)
        logger.info(f"Text layer — has_email={has_email}, has_phone={has_phone}")
        if HAS_PYMUPDF and (not has_email or not has_phone):
            if not HAS_OCR:
                logger.warning(
                    "EasyOCR is NOT installed — email/phone in icon-font headers cannot be read. "
                    "Fix: stop Flask, run  pip install easyocr  then restart."
                )
            else:
                ocr_text = _ocr_contact_strip(path)
                logger.info(f"OCR raw text: {repr(ocr_text[:300])}")
                if ocr_text:
                    # Extract only email/phone/name/location from OCR and inject as
                    # labelled lines. Never prepend raw OCR — it duplicates content
                    # that is already in the text layer.
                    contact_lines = []
                    _ocr_em = _extract_email(ocr_text) or ""
                    _ocr_ph = _extract_phone(ocr_text) or ""
                    if not has_email and _ocr_em:
                        contact_lines.append(f"Email: {_ocr_em}")
                    if not has_phone and _ocr_ph:
                        contact_lines.append(f"Phone: {_ocr_ph}")
                    # Try to extract name from OCR — the name heading is in the icon
                    # header image but absent from the PDF text layer.
                    _ocr_name = ""
                    _has_name = bool(re.search(
                        r'(?m)^\s*(?:name|full\s*name)\s*[:\-]', text, re.I
                    ))
                    if not _has_name:
                        for _ol in ocr_text.splitlines():
                            _ol = _ol.strip()
                            if not _ol:
                                continue
                            if (_ocr_em and _ocr_em in _ol):
                                continue
                            if (_ocr_ph and _ocr_ph.replace(" ", "") in _ol.replace(" ", "")):
                                continue
                            if re.search(r'@|https?://|www\.|linkedin|github|\d{6,}', _ol, re.I):
                                continue
                            if _looks_like_name(_ol):
                                _ocr_name = _ol
                                contact_lines.append(f"Name: {_ol}")
                                logger.info(f"OCR name extracted: {_ol}")
                                break
                    # Reconstruct email local part using name when OCR split it on
                    # underscore/dot or dropped a leading character.
                    # e.g. OCR "janani\najagopal@outlookcom" + name "JANANI RAJAGOPAL"
                    # → "janani_rajagopal@outlook.com"
                    if not has_email and _ocr_em and _ocr_name and '@' in _ocr_em:
                        _em_local, _em_dom = _ocr_em.split('@', 1)
                        _nwords = [w.lower() for w in _ocr_name.split() if len(w) >= 3]
                        # Fix dropped leading char: "ajagopal" is suffix of "rajagopal"
                        for _nw in _nwords:
                            if _nw.endswith(_em_local) and len(_nw) > len(_em_local):
                                _em_local = _nw
                                break
                        # Fix underscore split: OCR line before @-line is another name word
                        _ocr_ll = [_l.strip() for _l in ocr_text.splitlines()]
                        for _oi, _oline in enumerate(_ocr_ll):
                            if '@' in _oline:
                                if _oi > 0:
                                    _oprev = _ocr_ll[_oi - 1].lower()
                                    if (_oprev in _nwords and _oprev not in _em_local
                                            and re.match(r'^[a-z][a-z0-9]*$', _oprev)):
                                        _em_local = _oprev + '_' + _em_local
                                break
                        _em_new = _em_local + '@' + _em_dom
                        if _em_new != _ocr_em:
                            logger.info(f"OCR email reconstructed: {_ocr_em} → {_em_new}")
                            contact_lines = [
                                f"Email: {_em_new}" if _c.startswith("Email: ") else _c
                                for _c in contact_lines
                            ]
                    # Try to extract location from OCR header
                    _has_loc = bool(re.search(
                        r'(?m)^\s*(?:location|address|city)\s*[:\-]', text, re.I
                    ))
                    if not _has_loc:
                        _loc_pat = re.compile(
                            r'\b(?:india|uk|usa|uae|us|canada|australia|singapore|'
                            r'bangalore|bengaluru|chennai|hyderabad|pune|mumbai|delhi|'
                            r'noida|gurgaon|gurugram|coimbatore|trivandrum|kochi|'
                            r'jaipur|ahmedabad|kolkata|ludhiana|chandigarh)\b', re.I
                        )
                        for _ol in ocr_text.splitlines():
                            _ol = _ol.strip()
                            if _ol and _loc_pat.search(_ol):
                                if not re.search(r'@|\d{6,}', _ol):
                                    contact_lines.append(f"Location: {_ol}")
                                    logger.info(f"OCR location extracted: {_ol}")
                                    break
                    if contact_lines:
                        text = "\n".join(contact_lines) + "\n" + text
                        logger.info(f"OCR contact injected: {contact_lines}")

        text = _fix_wrapped_email(text)
        return text

    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}", exc_info=True)
        raise


def _extract_canvas_docx(body):
    """Extract and reconstruct text from DOCX files that use floating text boxes
    (Word Processing Canvas / Group layout).  Standard paragraph iteration yields
    nothing for these files because all content lives in anchored drawing shapes.
    """
    W_NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    MC_NS  = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    FALLBACK_TAG = f"{{{MC_NS}}}Fallback"

    def _in_fallback(el):
        for anc in el.iterancestors():
            if anc.tag == FALLBACK_TAG:
                return True
        return False

    def _box_lines(txbx):
        lines = []
        for p in txbx.findall(f"{{{W_NS}}}p"):
            t = "".join(r.text or "" for r in p.findall(f".//{{{W_NS}}}t"))
            if t.strip():
                lines.append(t.strip())
        return lines

    def _box_pos(txbx):
        """(y, x) position of the wps:wsp shape that contains this text box."""
        parent = txbx.getparent()          # wps:txbx
        if parent is None:
            return (0, 0)
        wsp = parent.getparent()           # wps:wsp
        if wsp is None or not wsp.tag.endswith("}wsp"):
            return (0, 0)
        spPr = wsp.find(f"{{{WPS_NS}}}spPr")
        if spPr is None:
            return (0, 0)
        xfrm = spPr.find(f"{{{A_NS}}}xfrm")
        if xfrm is None:
            return (0, 0)
        off = xfrm.find(f"{{{A_NS}}}off")
        if off is None:
            return (0, 0)
        return (int(off.get("y", 0)), int(off.get("x", 0)))

    # Collect non-VML-fallback text boxes
    boxes = []
    for txbx in body.findall(f".//{{{W_NS}}}txbxContent"):
        if _in_fallback(txbx):
            continue
        lines = _box_lines(txbx)
        text  = "\n".join(lines)
        if not text.strip():
            continue
        boxes.append({"pos": _box_pos(txbx), "text": text, "lines": lines})

    _JOB_WORDS = {
        "engineer", "manager", "lead", "analyst", "developer", "designer",
        "specialist", "consultant", "director", "officer", "executive",
        "programmer", "architect", "scientist", "researcher", "coordinator",
        "technician", "associate", "assistant", "head", "intern", "trainee",
    }

    # ── Classify each box ────────────────────────────────────────────────────
    for box in boxes:
        txt   = box["text"]
        first = box["lines"][0]
        pos   = box["pos"]

        # Icon / decorative single character
        if len(txt.strip()) <= 2:
            box["kind"] = "icon"
            continue

        # Shapes with explicit non-zero (y, x) positions are anchored inside the
        # experience section of the canvas — treat all such content as experience.
        if pos != (0, 0):
            box["kind"] = "experience_content"
            continue

        # Section label (matches SECTION_ALIASES, short text). Some templates put
        # the heading AND its list content in one shape (e.g. "Soft Skills" +
        # its bullet items all in a single sidebar text box) rather than in
        # separate boxes — keep any lines after the heading as that section's
        # own content instead of discarding them just because the box also
        # happens to start with a heading-shaped first line.
        sec = canonical_section_name(first)
        if sec and len(first.split()) <= 4:
            box["kind"]    = "label"
            box["section"] = sec
            if len(box["lines"]) > 1:
                box["extra_text"] = "\n".join(box["lines"][1:])
            continue

        # Contact info (phone / email present)
        if re.search(
            r"\+?\d[\d\s\-]{7,}\d|[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
            txt,
        ):
            box["kind"] = "contact"
            continue

        # Education content (degree / university keywords)
        if re.search(
            r"\b(?:university|college|school|b\.?com|b\.?sc|b\.?tech|m\.?sc|mba"
            r"|bachelor|master|degree)\b",
            txt, re.I,
        ):
            box["kind"] = "education"
            continue

        # Long single paragraph → professional summary
        if len(txt) > 200 and len(box["lines"]) <= 4:
            box["kind"] = "summary"
            continue

        # Multi-item list → skills
        if len(box["lines"]) >= 3:
            box["kind"] = "skills"
            continue

        # Short proper-case block: distinguish name vs. professional title
        words = first.split()
        if 2 <= len(words) <= 4 and all(w[0].isupper() for w in words[:3] if w.isalpha()):
            if any(w.lower() in _JOB_WORDS for w in words):
                box["kind"] = "title"   # e.g. "Validation Engineer"
            else:
                box["kind"] = "name"    # e.g. "Sindhu Sundaramoorthy"
        else:
            box["kind"] = "misc"

    # ── Reconstruct text in semantic reading order ────────────────────────────
    parts = []

    def _add(iterable):
        for b in iterable:
            parts.append(b["text"])

    # 1. Name first, then title
    _add(b for b in boxes if b.get("kind") == "name")
    _add(b for b in boxes if b.get("kind") == "title")

    # 2. Contact info — expand bare LinkedIn handle to full URL if present
    for box in boxes:
        if box.get("kind") != "contact":
            continue
        contact_lines = []
        for line in box["lines"]:
            line = line.strip()
            if re.match(r"^[a-zA-Z][a-zA-Z0-9\-]{4,}$", line):
                # Bare handle (no @, no digits only, no spaces) → LinkedIn slug
                line = f"https://www.linkedin.com/in/{line}"
            contact_lines.append(line)
        parts.append("\n".join(contact_lines))

    # 3. Professional summary
    lbl = next((b for b in boxes if b.get("kind") == "label" and b.get("section") == "summary"), None)
    if lbl:
        parts.append(lbl["lines"][0])
        if lbl.get("extra_text"):
            parts.append(lbl["extra_text"])
    _add(b for b in boxes if b.get("kind") == "summary")

    # 4. Experience — all explicitly-positioned shapes sorted by y (page order)
    exp_content = sorted(
        [b for b in boxes if b.get("kind") == "experience_content"],
        key=lambda b: b["pos"],
    )
    if exp_content:
        parts.append("EXPERIENCE")
        for b in exp_content:
            parts.append(b["text"])

    # 5. Education
    edu_lbl = next((b for b in boxes if b.get("kind") == "label" and b.get("section") == "education"), None)
    if edu_lbl:
        parts.append(edu_lbl["lines"][0])
        if edu_lbl.get("extra_text"):
            parts.append(edu_lbl["extra_text"])
    _add(b for b in boxes if b.get("kind") == "education")

    # 6. Skills — a template may have several labelled sub-groups sharing the
    # "skills" section (e.g. "Soft Skills" and "Hard Skills" as separate boxes);
    # emit every one of them, not just the first found.
    for sk_lbl in (b for b in boxes if b.get("kind") == "label" and b.get("section") == "skills"):
        parts.append(sk_lbl["lines"][0])
        if sk_lbl.get("extra_text"):
            parts.append(sk_lbl["extra_text"])
    _add(b for b in boxes if b.get("kind") == "skills")

    return "\n".join(parts)


# A short Title-Case label (no digits — excludes "Project 1:", "Client 2:" etc.)
# followed by a colon and a comma-separated list — "SIEM Tools: Splunk, QRadar,
# Microsoft Sentinel", "Firewalls: Palo Alto, Fortinet, Check Point". Matches
# the tool/skill-category lines common in a "Technical Skills" box, without
# matching narrative lines like "Role: SOC Analyst / Cyber Security Engineer"
# (no comma) or "Project 1: ..." (digit in the label).
_SKILL_CATEGORY_LINE_RE = re.compile(
    r'^[A-Z][A-Za-z/&\s]{1,30}:\s*[A-Za-z0-9].*(?:,\s*[A-Za-z0-9][^,:]*){1,}$'
)


def extract_text_from_docx(path):
    try:
        logger.info(f"Extracting text from DOCX: {path}")
        doc = Document(str(path))
        parts = []
        from docx.oxml.ns import qn
        from docx.table import Table as DocxTable
        from docx.text.paragraph import Paragraph as DocxParagraph

        # Read document section headers first — some templates put name/contact info there
        try:
            for section in doc.sections:
                for hdr in (section.header, section.first_page_header, section.even_page_header):
                    if hdr is not None:
                        for para in hdr.paragraphs:
                            t = para.text.strip()
                            if t:
                                parts.append(t)
                        for tbl in hdr.tables:
                            for row in tbl.rows:
                                # Join cells with a newline, not " | " — a space-pipe join
                                # glues the LAST line of one cell onto the FIRST line of the
                                # next cell with no line break between them (e.g. a sidebar
                                # cell's LinkedIn line fusing onto the main cell's first
                                # summary line), which hides that fused line from every
                                # downstream line-based heuristic (name/contact detection).
                                row_text = "\n".join(c.text.strip() for c in row.cells if c.text.strip())
                                if row_text:
                                    parts.append(row_text)
        except Exception:
            pass  # Headers are optional — skip silently if unavailable

        for child in doc.element.body:
            if child.tag == qn("w:p"):
                para = DocxParagraph(child, doc)
                if para.text.strip():
                    parts.append(para.text.strip())
            elif child.tag == qn("w:tbl"):
                tbl = DocxTable(child, doc)
                for row in tbl.rows:
                    # See note above — newline join prevents cross-column line-fusion,
                    # which is the main reason 2-column sidebar DOCX templates (name,
                    # title, contact block in one cell) lose their identity fields.
                    row_text = "\n".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        parts.append(row_text)

        body_text = "\n".join(parts).strip()

        # If the standard body extraction yielded nothing, the file likely uses
        # a floating text-box canvas layout — try the canvas extractor.
        if not body_text:
            logger.info("No body text found; trying canvas/text-box extraction")
            canvas_text = _extract_canvas_docx(doc.element.body)
            if canvas_text.strip():
                logger.info(f"Canvas extraction yielded {len(canvas_text)} chars")
                return canvas_text
        else:
            # Even when body text exists, supplement with canvas text boxes — some
            # templates have a sidebar (skills, contact) in floating boxes not in
            # body. A floating box's extracted text can be a large mixed blob (a
            # skills tool-list AND unrelated Project narrative bundled together),
            # so don't relocate the whole thing as one unit — that just moves the
            # misplacement instead of fixing it. Only lines that are clearly
            # shaped like a skill-category list ("SIEM Tools: Splunk, QRadar, ...")
            # get moved next to the nearest Skills heading; everything else keeps
            # the previous, safe default of landing at the end.
            try:
                canvas_text = _extract_canvas_docx(doc.element.body)
                if canvas_text.strip():
                    existing = set(l.strip().lower() for l in parts if l.strip())
                    new_lines = [
                        cl.strip() for cl in canvas_text.split('\n')
                        if cl.strip() and cl.strip().lower() not in existing
                    ]
                    skill_lines = [l for l in new_lines if _SKILL_CATEGORY_LINE_RE.match(l)]
                    other_lines = [l for l in new_lines if not _SKILL_CATEGORY_LINE_RE.match(l)]
                    if skill_lines:
                        insert_at = None
                        for _pi in range(len(parts) - 1, -1, -1):
                            if canonical_section_name(parts[_pi]) == "skills":
                                insert_at = _pi + 1
                                break
                        if insert_at is not None:
                            parts[insert_at:insert_at] = skill_lines
                        else:
                            parts.extend(skill_lines)
                    parts.extend(other_lines)
            except Exception:
                pass

        logger.info(f"Successfully extracted {len(parts)} parts from DOCX")
        return "\n".join(parts)
    except Exception as e:
        logger.error(f"Error extracting text from DOCX: {e}", exc_info=True)
        raise


def extract_resume_text(path, ext):
    if ext == "pdf":
        return extract_text_from_pdf(path)
    if ext == "docx":
        return extract_text_from_docx(path)
    raise ValueError("Unsupported file type")


_PROJECT_TABLE_HEADER_ALIASES = frozenset({
    "project details", "project detail", "project description",
})
_RESPONSIBILITY_TABLE_HEADER_ALIASES = frozenset({
    "role and responsibilities", "roles and responsibilities",
    "role and responsibility", "roles & responsibilities",
    "role & responsibilities", "role & responsibility",
    "responsibilities", "responsibility",
})


def _pdfplumber_project_table_text(path):
    """Detect a 2-column 'Project Details | Role and Responsibilities' PDF
    table and reconstruct it row by row, keeping each project paired with its
    own responsibilities.

    Plain linear text extraction reads a table cell-by-cell in content-stream
    order, which works fine for short single-line cells but scrambles a table
    once cells wrap across several lines — the two columns' lines interleave
    unpredictably, project descriptions get truncated, and unrelated nearby
    text (e.g. an Education line elsewhere on the page) can bleed in. pdfplumber
    reads each table cell as a whole, so rows stay correctly paired.

    Returns "" if pdfplumber is unavailable or no matching table is found —
    this only overrides the Projects field for this specific table shape;
    every other resume layout is completely unaffected.
    """
    if not HAS_PDFPLUMBER:
        return ""
    blocks = []
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    if not table or len(table) < 2 or len(table[0]) != 2:
                        continue
                    header = [(c or "").strip().lower() for c in table[0]]
                    if not (header[0] in _PROJECT_TABLE_HEADER_ALIASES
                            and header[1] in _RESPONSIBILITY_TABLE_HEADER_ALIASES):
                        continue
                    for row in table[1:]:
                        if len(row) != 2:
                            continue
                        details = (row[0] or "").strip()
                        resp = (row[1] or "").strip()
                        if not details:
                            continue
                        block = details
                        if resp:
                            resp_lines = [
                                "- " + rl.strip().lstrip("•●▪◦-*").strip()
                                for rl in resp.splitlines() if rl.strip()
                            ]
                            block += "\nRoles & Responsibilities:\n" + "\n".join(resp_lines)
                        blocks.append(block)
    except Exception as e:
        logger.warning(f"pdfplumber project-table extraction failed: {e}")
        return ""
    return "\n\n".join(blocks)


# ── Resume parsing ─────────────────────────────────────────────────────────────

def normalize_lines(text):
    cleaned = text.replace("•", "•")
    return [re.sub(r"\s+", " ", line).strip(" :-\t") for line in cleaned.splitlines() if line.strip()]


def canonical_section_name(line):
    cleaned = re.sub(r"[^a-zA-Z ]", " ", line).lower()
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return None
    for field, aliases in SECTION_ALIASES.items():
        if cleaned in aliases:
            return field
    return None


def _looks_like_company_name(line):
    """
    Detect if a line looks like a company name.
    E.g., "HCL TECH", "TATA Elxsi", "Lyrus Life Sciences Pvt. Ltd.", "Google India"
    """
    line = line.strip()

    # Remove trailing punctuation
    line = line.rstrip('.,;:!?')
    words = line.split()

    # Must be 2-6 words
    if not (2 <= len(words) <= 6):
        return False

    # Check if all words start with uppercase
    if not all(w and w[0].isupper() for w in words):
        return False

    # Exclude common non-company words that appear in job titles or descriptions
    exclude_words = {
        "the", "and", "or", "for", "with", "in", "at", "by", "as",
        "to", "from", "of", "on", "is", "are", "be", "been",
        "experience", "duration", "designation", "project",
        "highly", "skilled", "professional", "years", "month"
    }
    lower_words = {w.lower().rstrip('.,;:') for w in words}
    if lower_words & exclude_words:
        return False

    # Company keywords (including variations)
    company_keywords = {
        "tech", "technologies", "systems", "solutions", "services",
        "group", "corporation", "corp", "inc", "ltd", "llc", "llp",
        "company", "consulting", "consultants", "labs", "designs",
        "pvt", "private", "limited", "healthcare", "pharma", "bio",
        "sciences", "elxsi", "tcs", "infosys", "wipro", "accenture"
    }
    has_company_keyword = any(w.lower().rstrip('.,;:') in company_keywords for w in words)

    # Pattern 1: "XYZ Corp/Ltd/Inc" or "XYZ Group" etc
    if has_company_keyword:
        return True

    # Pattern 2: All words are uppercase (e.g., "HCL TECH", "TATA ELXSI")
    if all(w.isupper() for w in words):
        return True

    # Pattern 3: 2-3 capitalized words, likely a company (e.g., "Google India", "Amazon AWS")
    # But exclude if it looks like a title (contains job keywords) or a soft-skill descriptor
    # (e.g. "Good Team Player", "Optimistic Thinker" are personal-quality phrases, not companies).
    job_keywords = {
        "engineer", "manager", "lead", "analyst", "developer", "designer",
        "specialist", "consultant", "director", "officer", "executive"
    }
    soft_skill_words = {
        "team", "player", "thinker", "learner", "communicator", "listener",
        "optimistic", "motivated", "adaptable", "flexible", "reliable",
        "proactive", "dedicated", "driven", "focused", "creative",
        "innovative", "passionate", "responsible", "organized", "punctual",
        "enthusiastic", "diligent", "hardworking", "versatile", "empathetic",
        "confident", "assertive", "collaborative", "analytical", "strategic",
    }
    has_job_keyword = any(w.lower() in job_keywords for w in words)
    has_soft_skill = any(w.lower() in soft_skill_words for w in words)
    if not has_job_keyword and not has_soft_skill and len(words) <= 4:
        return True

    return False


def find_sections(lines):
    sections = {key: "" for key in SECTION_ALIASES}
    _SKIP_HEADINGS = {
        "personal", "personal information", "personal details", "other personal details",
        "references", "declaration", "interests", "languages",
        "websites portfolios and profiles", "websites and profiles",
        "websites portfolios", "and profiles",
        "additional information", "other information", "extra curricular",
        "soft skills",
        # "achievements" is NOT here — it has its own SECTION_ALIASES key and folds into certifications via _SECTION_FOLD
    }
    _SKIP_STARTSWITH = {
        "websites portfolios and profiles", "websites and profiles",
        "websites portfolios", "and profiles",
    }
    current = None
    bucket = []
    lines = list(lines)  # materialise so we can do lookahead
    _work_meta_re = re.compile(r'^\s*(duration|designation|project)\s*:', re.I)
    _numbered_item_re = re.compile(r'^\s*(\d+)\.\s')
    # When a skip-heading block (Contact/Hobbies/Languages/etc.) interrupts a
    # NUMBERED list mid-section — common on a page break where a sidebar column
    # continues between two halves of the same Work Experience/Projects list —
    # remember which section to resume into if the next real content picks the
    # numbering back up exactly where it left off. Content that ISN'T a matching
    # continuation stays unmapped, same as before this existed.
    _resume_candidate = None

    def _flush(section, bucket_lines):
        # Append rather than overwrite when this section already holds content
        # from an earlier, separate run — e.g. a resume with both "SUMMARY" and
        # a later "OBJECTIVE" heading (both canonical "summary") previously had
        # the second heading's flush silently wipe out the first heading's
        # content instead of combining them, since a plain reassignment doesn't
        # care whether anything was already there. This also covers the
        # numbered-list resume case (current was reactivated into a section that
        # already has its earlier bullets flushed) with no special-casing needed.
        if not (section and bucket_lines):
            return
        content = "\n".join(bucket_lines).strip()
        if not content:
            return
        existing = sections.get(section)
        if existing:
            # Guard against a verbatim-duplicate block re-appending under the
            # same heading — e.g. a source document whose entire body was
            # pasted in more than once, which otherwise multiplies every
            # section's content once per repeat. A genuinely new block (the
            # SUMMARY+OBJECTIVE case this append exists for) is never a
            # substring of what's already collected, so this only ever
            # suppresses true repeats.
            if content in existing or existing in content:
                if len(content) > len(existing):
                    sections[section] = content
                return
            sections[section] = existing + "\n" + content
        else:
            sections[section] = content

    for _line_idx, line in enumerate(lines):
        # A single lowercase word ending with a period is a sentence continuation
        # (e.g. "projects." wrapped from "...strategies for GxP projects."), not a heading.
        words_raw = line.split()
        if len(words_raw) == 1 and words_raw[0][-1] == "." and words_raw[0][0].islower():
            if current:
                bucket.append(line)
            continue

        cleaned_line = re.sub(r"[^a-zA-Z ]", " ", line).lower()
        cleaned_line = re.sub(r"\s+", " ", cleaned_line).strip()

        _is_skip = cleaned_line in _SKIP_HEADINGS or any(
            cleaned_line.startswith(sh) for sh in _SKIP_STARTSWITH
        )
        if _is_skip:
            if current and bucket:
                _flush(current, bucket)
                if current not in _THROWAWAY_SECTIONS:
                    _resume_candidate = current
            current = None
            bucket = []
            continue

        section = canonical_section_name(line)
        inline_remainder = None

        if not section:
            # An unlabeled company-name line starts the experience section only when the
            # next 1-2 non-empty lines look like work-experience metadata
            # ("Duration :", "Designation :", "Project :"). This prevents the person's
            # own name (e.g. "SIVARANJANI D" in a page-2 repeated header) from being
            # misidentified as a company name when current is None.
            _next_lines = lines[_line_idx + 1 : _line_idx + 3]
            _followed_by_work_meta = any(
                _work_meta_re.match(nl) for nl in _next_lines if nl.strip()
            )
            if _looks_like_company_name(line) and _followed_by_work_meta:
                section = "experience"
                inline_remainder = line  # include the company name as the first experience line
            else:
                words = line.split()
                # Lines starting with a non-letter character (bullet markers ✓, •, –,
                # numbered-list "1.", etc.) are list items, never section headings.
                # Skip the prefix-heading scan to avoid e.g. "✓ Expertise in Defect Tracking"
                # being parsed as a 2-word "✓ Expertise" → skills heading. When current is
                # None, fall through to the numbered-continuation resume check below instead
                # of dropping it outright — this is exactly the shape of a rescued line
                # ("12. ...") picking a numbered list back up after a sidebar interruption.
                _is_list_marker_line = bool(words) and not words[0][0].isalpha()
                if _is_list_marker_line and current:
                    bucket.append(line)
                    continue
                if not _is_list_marker_line:
                    for n in range(min(3, len(words)), 0, -1):
                        prefix_section = canonical_section_name(" ".join(words[:n]))
                        if prefix_section:
                            remainder = " ".join(words[n:]).strip()
                            # Single-word prefix only counts as a heading when:
                            #   • it IS the entire line (no remainder), OR
                            #   • it ends with a separator like "Skills:", OR
                            #   • it is ALL-CAPS (e.g. "EXPERIENCE VALIDATION ENGINEER..." from
                            #     row-interleaved PDF extraction — heading + inline content on same line).
                            # Title-Case words like "Experience on all ALM modules" are content.
                            _word0_bare = words[0].rstrip('.,;:!?')
                            if n == 1 and remainder and not words[0].endswith(':') and not _word0_bare.isupper():
                                break  # "Experience on all ALM modules" → content, not heading
                            section = prefix_section
                            inline_remainder = remainder if remainder else None
                            break

        if section:
            # Guard: a single Title-case word (e.g. "Qualifications") that immediately
            # follows an incomplete line (no terminal punctuation) is a wrapped continuation
            # of that line — NOT a new section heading. E.g. the bullet "…and Software\n
            # Qualifications." wraps across lines but both chunks belong to the same section.
            # ALL-CAPS words (e.g. "EDUCATION") are always genuine headings; skip the guard.
            if (len(words_raw) == 1 and current and bucket
                    and not words_raw[0].rstrip('.,;:!?').isupper()):
                # Look past any trailing empty lines to the last substantive bucket line.
                # An empty line between wrapped content (e.g. PDF page rendering gaps)
                # must not prevent the guard from seeing the prior non-empty line.
                _prev = next(
                    (bl.rstrip() for bl in reversed(bucket) if bl.strip()),
                    ""
                )
                if _prev and _prev[-1] not in '.!?;:':
                    bucket.append(line)
                    continue
            if section not in _THROWAWAY_SECTIONS:
                _resume_candidate = None  # a genuine heading resolves any skip ambiguity
            elif current and current not in _THROWAWAY_SECTIONS:
                # Entering a throwaway section (e.g. "Hobbies" → canonical "interests")
                # straight from a meaningful one is the same shape as a _SKIP_HEADINGS
                # interruption — remember it so a numbered list can resume afterward.
                # Guard on `current` being set: if we're already inside a skip/throwaway
                # run (current is None), a chained throwaway heading must NOT clobber
                # whatever _resume_candidate that run already recorded.
                _resume_candidate = current
            if section == current:
                # A bare repeat of the heading word itself (e.g. "Summary" wrapped mid-
                # summary from "Validation Summary Report.") is ignored so the section
                # doesn't restart and lose its content. But when the matched heading
                # words are just the LEAD-IN of an ordinary sentence — e.g. "Professional
                # experience in Power BI, Power Apps & Power Automate" matches the
                # "Professional Experience" heading alias — the real content living in
                # inline_remainder must still be kept, not silently dropped.
                if inline_remainder:
                    bucket.append(inline_remainder)
                continue
            _flush(current, bucket)
            # A throwaway section (e.g. "Hobbies" → canonical "interests") is discarded
            # unconditionally at the very end (its _SECTION_FOLD target is None), so its
            # content is never collected in the first place — same as current=None does
            # for a literal _SKIP_HEADINGS match. This lets a numbered list resume through
            # it exactly like it would through an explicit skip heading.
            current = None if section in _THROWAWAY_SECTIONS else section
            bucket = []
            if inline_remainder and current is not None:
                bucket.append(inline_remainder)
            continue

        if current:
            # Drop bare page-number lines (1–3 digit standalone numbers from PDF
            # page breaks, e.g. "1" between two halves of a skills list, "3" at the
            # end of the last page). Legitimate content never appears as a bare digit.
            if not re.match(r'^\s*\d{1,3}\s*$', line):
                bucket.append(line)
            continue

        # current is None: either nothing has started yet, or a skip-heading block
        # (Contact/Hobbies/Languages/etc.) just interrupted a section. In the latter
        # case, resume into that section if a line picks its numbered list back up
        # exactly where it left off — e.g. Work Experience bullets 1-11 on page 1,
        # then a whole sidebar column (contact/hobbies/languages), then "12. ..." /
        # "13. ..." continuing on page 2. Keep watching across the entire unmapped
        # stretch (it isn't going anywhere else either way) until either a match is
        # found or a genuine new heading resolves the ambiguity for good.
        if _resume_candidate is not None:
            _m = _numbered_item_re.match(line)
            if _m:
                _prev_nums = re.findall(r'(?:^|\n)\s*(\d+)\.\s', sections.get(_resume_candidate, ""))
                if _prev_nums and int(_m.group(1)) == int(_prev_nums[-1]) + 1:
                    current = _resume_candidate
                    bucket = [line]
                    _resume_candidate = None
                    continue

    _flush(current, bucket)

    # Fold extended sections into the 6 core display sections, then return only those.
    for src, dst in _SECTION_FOLD.items():
        content = sections.get(src, "")
        if content and dst:
            sep = "\n" if sections.get(dst) else ""
            sections[dst] = (sections.get(dst) or "") + sep + content

    return {k: sections.get(k, "") for k in _CORE_SECTIONS}


def parse_label_value(lines, labels):
    label_pattern = "|".join(re.escape(label) for label in labels)
    pattern = re.compile(rf"(?:^|\b)({label_pattern})\s*[:\-]\s*(.+)$", re.I)
    for line in lines[:30]:
        match = pattern.search(line)
        if match:
            return match.group(2).strip()
    return ""


def _fix_ocr_zero_o_confusion(email):
    """Correct EasyOCR's frequent digit/letter mix-up in email local parts: a
    tightly-kerned '0' (zero) in an image-based contact header is very often
    read back as the letter 'o'. Only touch an 'o'/'O' that sits directly
    between two digits, or right before '@' with a digit just before it — a
    real local part essentially never has a bare letter in exactly that
    position, but a misread '0' commonly does, so this is safe to apply
    unconditionally rather than only for OCR-sourced text.
    """
    if "@" not in email:
        return email
    local, domain = email.split("@", 1)
    local = re.sub(r"(?<=\d)[oO](?=\d)", "0", local)
    local = re.sub(r"(?<=\d)[oO]$", "0", local)
    return f"{local}@{domain}"


def _extract_email(text):
    """Return the first valid email found in text, handling common OCR artifacts."""
    # Pre-pass: join lines where an email wraps mid-TLD (e.g. "foo@gmail.c\nom").
    # Require the domain to end with exactly 1 alpha char (incomplete TLD stub) so
    # we don't accidentally join a COMPLETE email (e.g. "foo@gmail.com") with the
    # unrelated word on the next line (e.g. "Mobile").
    text = re.sub(
        r"([A-Za-z0-9._%+\-]+@(?:[A-Za-z0-9\-]+\.)+[A-Za-z]{1})\n([A-Za-z]{2,5})\b",
        r"\1\2", text,
    )

    # Pass 1: standard clean email
    m = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    if m:
        return _fix_ocr_zero_o_confusion(m.group().strip().rstrip(".,;)"))

    # Pass 2: scan lines — rebuild from the @ position handling OCR noise
    for line in text.splitlines():
        if "@" not in line:
            continue
        at_idx = line.find("@")

        # Local part: grab word(s) immediately before @
        before = line[:at_idx]
        local_match = re.search(r"([A-Za-z0-9][A-Za-z0-9.\s_%+\-]{0,50})$", before)
        if not local_match:
            continue
        local = local_match.group(1).strip()
        local = re.sub(r"\s+", ".", local)          # OCR spaces → dots
        local = re.sub(r"\.{2,}", ".", local)        # collapse duplicate dots

        # Domain part: first non-space token after @
        after_raw = line[at_idx + 1:].strip()
        domain = after_raw.split()[0] if after_raw else ""
        domain = domain.strip(".,;:()[]")
        domain = re.sub(r"\s+", "", domain)
        # OCR often drops the dot before TLD: "outlookcom" → "outlook.com"
        domain = re.sub(
            r"(?<=[a-zA-Z])(com|net|org|in|io|co|uk|edu|gov|info|biz)$",
            r".\1", domain, flags=re.I,
        )

        if local and domain and "." in domain and len(domain.split(".")[-1]) >= 2:
            return _fix_ocr_zero_o_confusion(f"{local}@{domain}")
    return ""


def _extract_phone(text):
    """Return the first phone number with 10–15 digits, fixing common OCR artifacts."""
    for line in text.splitlines():
        for m in re.finditer(r"\(?\+?\(?\d[\d\s\-\.\(\)]{7,20}\d", line):
            candidate = m.group().strip()
            digits = re.sub(r"\D", "", candidate)
            if 10 <= len(digits) <= 15:
                # OCR often reads '+' as '4' or '1' for country-code prefix
                # e.g. "491 8220 133 233" → "+91 8220 133 233"
                if re.match(r"^[41](\d{2}[\s\-])", candidate) and not candidate.startswith("+"):
                    candidate = "+" + candidate[1:]
                return candidate
    return ""


# ── Name / Title candidate helpers ────────────────────────────────────────────

_NAME_STOPWORDS = frozenset({
    "resume", "curriculum", "vitae", "cv", "profile", "candidate",
    "the", "a", "an", "and", "or", "but", "for", "of", "in", "on",
    "at", "to", "from", "with", "by", "as", "who", "that", "which",
    "highly", "skilled", "experienced", "results", "driven", "oriented",
    "seeking", "passionate", "motivated", "proactive", "dedicated",
    "dynamic", "detail", "focused", "professional", "over", "years", "year",
    "strong", "excellent", "extensive", "hands", "proven", "seasoned",
    "total", "experience", "current", "currently", "presently", "previously",
    "recently", "employer", "designation", "qualification", "available", "objective",
    "key", "skills", "competencies", "summary", "objectives", "highlights",
    "personal", "information", "education", "project", "work", "certification",
    "achievement", "language", "hobby", "extracurricular", "tools", "applications",
    "core", "section", "header", "details", "accomplishment", "award",
    "data", "analysis", "analyst", "support", "domain", "domains",
    # Job-title / designation words — these appear bold near the top of many
    # resumes (right under the name) and were being mistaken for the name itself.
    "lead", "engineer", "specialist", "manager", "consultant", "developer",
    "officer", "coordinator", "administrator", "director", "executive",
    "associate", "senior", "junior", "assistant", "validation", "regulatory",
    "affairs", "quality", "assurance", "registration", "compliance",
    "architect", "leadership", "scientist", "technician", "supervisor",
    "annexure", "appendix", "gamp", "cfr", "ich", "gxp", "fda", "qms",
    "sop", "capa", "iso",
    "enclosure", "declaration", "format",
    "biodata", "bio", "curriculum", "vitae",
})

# Degree/qualification abbreviations (dots stripped) — e.g. "M.Sc.", "B.Tech"
# were passing the name heuristic because they're short, bold, all-alpha tokens.
_DEGREE_ABBR = frozenset({
    "msc", "bsc", "mtech", "btech", "be", "me", "mba", "bba", "bca", "mca",
    "phd", "ma", "ba", "bcom", "mcom", "llb", "llm", "md", "bds", "mds",
    "bpharm", "mpharm", "bed", "med",
})


_DEGREE_ABBR_PAT = re.compile(
    r'^(?:M\.?Sc|B\.?Sc|B\.?E|M\.?E|B\.?Tech|M\.?Tech|Ph\.?D|MBA|BBA|BCA|MCA|'
    r'B\.?Com|M\.?Com|BA|MA|LLB|LLM|MBBS|MD|MS|BE|ME|BSc|MSc)\.?$',
    re.I,
)

_DEGREE_SUFFIX_PAT = re.compile(
    r',?\s*(?:M\.?Sc|B\.?Sc|B\.?E|M\.?E|B\.?Tech|M\.?Tech|Ph\.?D|MBA|BBA|BCA|MCA|'
    r'B\.?Com|M\.?Com|BA|MA|LLB|LLM|MBBS|MD|MS|BE|ME|BSc|MSc)\.?'
    r'(?:\s*[,/]\s*(?:M\.?Sc|B\.?Sc|B\.?E|M\.?E|B\.?Tech|M\.?Tech|Ph\.?D|MBA|BBA|'
    r'BCA|MCA|B\.?Com|M\.?Com|BA|MA|LLB|LLM|MBBS|MD|MS|BE|ME|BSc|MSc)\.?)*\s*$',
    re.I,
)


def _strip_degree_suffix(text):
    """Remove trailing academic degree qualifiers from a candidate name line.
    'Abhisheak Kaniyala, M.Sc.' → 'Abhisheak Kaniyala'
    'Ravi Kumar B.E., M.Tech' → 'Ravi Kumar'
    """
    return _DEGREE_SUFFIX_PAT.sub("", text).strip()


def _looks_like_name(line):
    """Return True only if *line* could plausibly be a person's full name."""
    line = line.strip()
    if not line or len(line) > 55:
        return False
    if _DEGREE_ABBR_PAT.match(line):
        return False
    if re.search(r'[,;:!?]|\.{2,}|—|–|\(|\)', line):
        return False
    if line.endswith('.') and ' ' in line:
        return False
    words = line.split()
    if not (1 <= len(words) <= 5):
        return False
    for w in words:
        clean = re.sub(r"[-'.]", "", w)
        if not clean.isalpha():
            return False
        if len(clean) > 3 and not w[0].isupper():
            return False
    lower_words = {w.lower().strip(".,;:-") for w in words}
    if lower_words & _NAME_STOPWORDS:
        return False
    alpha_words = {re.sub(r"[^a-z]", "", w.lower()) for w in words}
    if alpha_words & _DEGREE_ABBR:
        return False
    return bool(words) and words[0][0].isupper()


_TITLE_SECTION_WORDS = frozenset({
    "summary", "profile", "objective", "overview", "experience", "education",
    "skills", "certifications", "projects", "employment", "history", "background",
    "qualifications", "competencies", "expertise", "achievements", "declaration",
    "references", "hobbies", "interests", "languages",
})

_LANGUAGE_NAMES = frozenset({
    "english", "tamil", "hindi", "telugu", "kannada", "malayalam", "marathi",
    "gujarati", "bengali", "punjabi", "urdu", "odia", "assamese", "sanskrit",
    "konkani", "sindhi", "kashmiri", "manipuri", "bodo", "dogri", "maithili",
    "french", "german", "spanish", "italian", "portuguese", "russian",
    "mandarin", "chinese", "japanese", "korean", "arabic", "dutch",
})


def _looks_like_title(line):
    """Return True only if *line* could plausibly be a professional job title."""
    line = line.strip()
    if not line or len(line) > 80 or len(line) < 2:
        return False
    words = line.split()
    if len(words) > 8:
        return False
    if re.search(
        r'\bhighly\b|\bseeking\b|\blooking for\b|with \d+\s*years?|over \d+\s*years?|'
        r'\byears? of\b|\bexperienced in\b|\bdetail.oriented\b|\bresults.driven\b|'
        r'\bdedicated to\b|\bresponsible for\b|\bwho has\b',
        line, re.I,
    ):
        return False
    # Reject common resume section headings (e.g. "PROFILE SUMMARY", "WORK EXPERIENCE")
    lower_words = {w.lower().rstrip('.:') for w in words}
    if len(lower_words & _TITLE_SECTION_WORDS) >= 1 and len(words) <= 3:
        return False
    # Reject language lists (e.g. "English Tamil", "Hindi French German") — these
    # get mistaken for a short Title-Case job title when a resume has no real
    # headline and the fallback scan picks up whatever short line it can find.
    if lower_words and lower_words <= _LANGUAGE_NAMES:
        return False
    return any(w[0].isupper() for w in words if w and w[0].isalpha())


# ── Contact-header pattern: "Name    Mobile: +91...   E-Mail: ..." ──────────

_CONTACT_LABEL_PAT = re.compile(
    r'\b(?:mobile|phone|tel|email|e[\-\.\ ]?mail|contact)\s*[:\-]',
    re.I,
)


def _extract_name_from_header_line(lines):
    """
    Pull the candidate name from a line that bundles name + contact details, e.g.
    'Santhoshkumar K    Mobile: +91 8608109310    E-Mail: xxx@yyy.com'
    The name precedes the first contact label keyword.
    """
    for line in lines[:8]:
        m = _CONTACT_LABEL_PAT.search(line)
        if m:
            before = line[:m.start()].strip()
            if before and _looks_like_name(before):
                return before
    return ""


def _fuzzy_abbr_pattern(abbr):
    """Build a regex for a degree abbreviation that tolerates the dots/spaces
    resumes commonly insert between its letters (e.g. "M.Sc.", "B. Tech")."""
    return r'\.?\s*'.join(re.escape(ch) for ch in abbr)


_DEGREE_FIELD_PREFIX_RE = re.compile(
    r'\b(?:' + '|'.join(
        _fuzzy_abbr_pattern(a) for a in sorted(_DEGREE_ABBR, key=len, reverse=True)
    ) + r')\.?\s*(?:degree\s*)?in\s*$',
    re.I,
)


def _extract_title_from_para(text):
    """
    Extract a job title embedded in a summary sentence, e.g.:
    'Highly skilled ... Validation and Compliance Lead with 7 years of experience...'
    → 'Validation and Compliance Lead'
    """
    # Pattern 1: <Title Phrase> with N year(s)
    m = re.search(r'\s+with\s+(?:over\s+)?\d+\+?\s*years?\b', text[:700], re.I)
    if m:
        before = text[:m.start()].strip()
        # Grab the last run of capitalized words (with optional "and/&" connectors)
        title_m = re.search(
            r'([A-Z][a-zA-Z]+(?:\s+(?:and|&|[A-Z][a-zA-Z]+)){1,5})$',
            before,
        )
        if title_m:
            candidate = title_m.group(1).strip()
            words = candidate.split()
            # Reject "<degree> in <field>" phrasing (e.g. "M.Sc. in Analytical
            # Chemistry with 9 years of ... experience") — the captured phrase
            # is a field of study, not a job title, when it's immediately
            # preceded by a degree abbreviation + "in".
            preceding = before[:title_m.start()]
            if (2 <= len(words) <= 7
                    and not _DEGREE_FIELD_PREFIX_RE.search(preceding)):
                return candidate
    # Pattern 2: as a/an <Title>
    m2 = re.search(
        r'\bas\s+an?\s+([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+){0,5})',
        text[:700],
    )
    if m2:
        candidate = m2.group(1).strip()
        # Require >= 2 words, same as Pattern 1 above — a single word here
        # (e.g. "Lead" from "as a Lead/Design Engineer", where the capture
        # stops at the slash) is almost always a truncated fragment, not a
        # real title, and blocks the fuller top-of-resume tagline title
        # (Title priority 2) from ever being tried.
        if 2 <= len(candidate.split()) <= 7:
            return candidate
    return ""


_INTRO_SKIP_RE = re.compile(
    r'@|www\.|linkedin|github|\d{6,}|\+\d{1,3}[\s\-]?\d{3,}|(?<!\w)\d{10,}|'
    r'\b(?:mobile|phone|tel|email|e[\-\.]mail)\s*[:\-]',
    re.I,
)
_LOCATION_LINE_RE = re.compile(
    r'\b(?:india|karnataka|bangalore|bengaluru|chennai|hyderabad|pune|mumbai|delhi|'
    r'noida|gurgaon|gurugram|coimbatore|trivandrum|kochi|jaipur|ahmedabad|kolkata|'
    r'ludhiana|chandigarh|tamil\s*nadu|maharashtra|andhra|telangana|kerala|'
    r'uttar\s*pradesh|rajasthan|gujarat|west\s*bengal)\b',
    re.I,
)


def _extract_intro_paragraph(lines):
    """
    Collect the unheaded introductory block that precedes the first section heading.
    Many resumes start with a professional overview that has no label — this function
    captures it so it can be used as the Professional Summary field.
    Contact-info lines and short standalone location lines are excluded.
    """
    result = []
    for line in lines:
        if canonical_section_name(line):
            break
        stripped = line.strip()
        if not stripped or len(stripped) < 20:
            continue
        if _INTRO_SKIP_RE.search(stripped):
            continue
        # Reject short lines that are purely location data (e.g. "Coimbatore, India")
        if len(stripped) < 45 and _LOCATION_LINE_RE.search(stripped):
            continue
        result.append(stripped)
    return "\n".join(result).strip()


# Role keyword → known group role (checked in order; first match wins) ────────

_ROLE_MAPPING = [
    # CSV (checked before plain "Validation" — more specific)
    ({"csv"},                              {"lead", "senior", "head", "manager", "principal"}, "CSV Lead"),
    ({"csv"},                              set(),                                               "CSV Analyst"),
    ({"computer", "system", "validation"}, {"lead", "senior", "head", "manager"},              "CSV Lead"),
    ({"computer", "system", "validation"}, set(),                                              "CSV Analyst"),
    # CQV
    ({"cqv"},                              {"lead", "senior", "head"},                         "CQV Lead"),
    ({"cqv"},                              set(),                                               "CQV Engineer"),
    ({"commissioning", "qualification"},   {"lead", "senior"},                                 "CQV Lead"),
    ({"commissioning", "qualification"},   set(),                                              "CQV Engineer"),
    # Tosca
    ({"tosca"},                            {"lead", "senior", "head"},                         "Tosca Lead"),
    ({"tosca"},                            set(),                                               "Tosca Engineer"),
    # Automation (after Tosca so "Tosca Automation Lead" → Tosca Lead, not Automation Lead)
    ({"automation"},                       {"lead", "senior", "head", "manager"},               "Automation Lead"),
    ({"automation"},                       set(),                                               "Automation Engineer"),
    # Validation (after CSV/CQV)
    ({"validation"},                       {"lead", "compliance", "senior", "head", "manager"}, "Validation Lead"),
    ({"validation"},                       set(),                                               "Validation Engineer"),
    # Test / QA
    ({"test"},                             {"lead", "senior", "manager", "head"},               "Test Lead"),
    ({"qa"},                               {"lead", "senior", "manager", "head"},               "Test Lead"),
    ({"quality", "assurance"},             {"lead", "senior", "manager"},                       "Test Lead"),
    ({"test"},                             set(),                                               "Test Engineer"),
    ({"qa"},                               set(),                                               "QA Engineer"),
    # IT roles
    ({"full", "stack"},                    set(),                                               "Full Stack Developer"),
    ({"frontend"},                         set(),                                               "Frontend Developer"),
    ({"front", "end"},                     set(),                                               "Frontend Developer"),
    ({"backend"},                          set(),                                               "Backend Developer"),
    ({"back", "end"},                      set(),                                               "Backend Developer"),
    ({"devops"},                           set(),                                               "DevOps Engineer"),
    ({"cloud"},                            {"engineer", "architect", "developer"},              "Cloud Engineer"),
    ({"data"},                             {"engineer"},                                        "Data Engineer"),
    ({"data"},                             {"analyst"},                                         "Data Analyst"),
    ({"business"},                         {"analyst"},                                         "Business Analyst"),
    ({"scrum"},                            set(),                                               "Scrum Master"),
    ({"ui"},                               {"ux"},                                              "UI/UX Designer"),
    ({"ux"},                               set(),                                               "UI/UX Designer"),
    ({"solution"},                         {"architect"},                                       "Solution Architect"),
    ({"cybersecurity"},                    set(),                                               "Cybersecurity Engineer"),
    ({"security"},                         {"engineer", "analyst", "specialist"},              "Cybersecurity Engineer"),
    ({"database"},                         {"administrator", "admin"},                          "Database Administrator"),
    ({"dba"},                              set(),                                               "Database Administrator"),
    ({"artificial", "intelligence"},       set(),                                               "AI/ML Engineer"),
    ({"machine", "learning"},              set(),                                               "AI/ML Engineer"),
    ({"sap"},                              set(),                                               "SAP Consultant"),
    ({"salesforce"},                       set(),                                               "Salesforce Developer"),
    ({"software"},                         set(),                                               "Software Developer"),
    ({"developer"},                        set(),                                               "Software Developer"),
    ({"project"},                          {"manager"},                                         "Project Manager"),
    ({"design", "control"},               set(),                                               "Design Control Consultant"),
    ({"document", "control"},             set(),                                               "Design Control Consultant"),
]


def _map_to_group_role(raw_title):
    """Map a free-form job title to the closest known group role, or return unchanged."""
    if not raw_title:
        return raw_title
    lower_words = set(re.findall(r'\b\w+\b', raw_title.lower()))
    for required, also_has, role in _ROLE_MAPPING:
        if not (required <= lower_words):
            continue
        if also_has and not (also_has & lower_words):
            continue
        return role
    return raw_title


_PDF_CONTACT_RE = re.compile(
    r'@|\d{6,}|\+\d|\b(?:mobile|phone|tel|email|e[\-\. ]?mail|website|www\.|http)\b',
    re.I,
)


def _extract_name_from_pdf_fonts(path):
    """
    Use PyMuPDF font metadata to find the candidate name at the top of page 1.
    Looks for bold+underlined, then bold-only, then largest-font text that passes
    _looks_like_name(). Returns "" when PyMuPDF is unavailable or nothing found.
    """
    if not HAS_PYMUPDF:
        return ""
    try:
        import pymupdf as fitz
        doc = fitz.open(str(path))
        if not doc.page_count:
            doc.close()
            return ""
        page = doc[0]
        top_zone = page.rect.height * 0.30
        dict_data = page.get_text("dict")
        doc.close()
    except Exception:
        return ""

    candidates = []  # (score, font_size, text)
    max_size = 0.0

    for block in dict_data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            y0 = line.get("bbox", [0, 0, 0, 9999])[1]
            if y0 > top_zone:
                continue
            name_parts = []
            line_max_size = 0.0
            for span in line.get("spans", []):
                txt = span.get("text", "").strip()
                if not txt or _PDF_CONTACT_RE.search(txt):
                    continue
                flags = span.get("flags", 0)
                size = span.get("size", 0.0)
                is_bold = bool(flags & 16)
                is_underline = bool(flags & 4)
                line_max_size = max(line_max_size, size)
                name_parts.append((txt, size, is_bold, is_underline))
            if not name_parts:
                continue
            max_size = max(max_size, line_max_size)
            line_text = " ".join(p[0] for p in name_parts)
            bold = any(p[2] for p in name_parts)
            underline = any(p[3] for p in name_parts)
            # Try stripped version first (handles "Name, M.Sc." lines)
            stripped_line = _strip_degree_suffix(line_text)
            # Try from longest prefix down; take the first that looks like a name
            for _candidate_text in ([stripped_line, line_text] if stripped_line != line_text else [line_text]):
                words = _candidate_text.split()
                _found = False
                for n in range(min(5, len(words)), 0, -1):
                    candidate = " ".join(words[:n])
                    if _looks_like_name(candidate):
                        score = (bold and underline) * 4 + bold * 2
                        candidates.append((score, line_max_size, candidate))
                        _found = True
                        break
                if _found:
                    break

    if not candidates:
        return ""

    # Add large-font bonus now that global max_size is known
    scored = [
        (score + (1 if size >= max_size * 0.85 else 0), size, text)
        for score, size, text in candidates
    ]
    scored.sort(key=lambda x: (-x[0], -x[1]))
    return scored[0][2][:80]


_SIDEBAR_CONTACT_HEAD = re.compile(
    r'^(contact|contact details|contact information|phone|mobile|tel|telephone|'
    r'e-?mail|email|linked ?in|www|website|web|address|fax|bold profile|'
    r'social link|social)\s*:?\s*$', re.I)
_SIDEBAR_SKILLS_HEAD = re.compile(r'^(skills?|key skills|technical skills|core skills)\s*:?\s*$', re.I)
_SIDEBAR_OTHER_HEAD = re.compile(r'^(languages?|interests|hobbies|references)\s*:?\s*$', re.I)
_SIDEBAR_EMAIL = re.compile(r'^[\w.+\-]+@[\w.\-]+\.\w{2,}$')
_SIDEBAR_PHONE = re.compile(r'^[\+\(]?\d[\d\s\-\(\)]{6,}$')
_SIDEBAR_URL = re.compile(r'^(https?://|www\.)', re.I)
_SIDEBAR_BULLET = ('•', '●', '▪', '◦', '‣', '·')


def _extract_sidebar_bleed_from_experience(exp_text):
    """Split a 2-column contact/skills SIDEBAR block out of Work-Experience text.

    On sidebar-style PDFs the right/left sidebar (Contact + Skills) is read
    interleaved into the middle of the experience column, e.g.:
        • Writing test cases with use of TestNG and
        Contact
        Phone
        +91-...
        E-mail  ...  LinkedIn  ...  WWW  ...
        Skills
        Selenium WebDriver
        JAVA  ...
        • Automating all the manual test cases to ...
    The block always STARTS at a contact heading (Contact/Phone/Email/LinkedIn/
    WWW) — which never legitimately appears inside a job's bullet list — and runs
    until the real experience resumes (a bullet, a date/job line, a 'Project:'
    line, or a full prose sentence). We remove that run from experience and return
    any Skills items found inside it, so the caller can restore them to Skills.

    Returns (cleaned_experience, [skill_items]).
    """
    lines = exp_text.split('\n') if exp_text else []
    n = len(lines)
    if not any(_SIDEBAR_CONTACT_HEAD.match(l.strip()) for l in lines):
        return exp_text, []

    def _resumes_work(s):
        return (s.startswith(_SIDEBAR_BULLET)
                or bool(re.match(r'^\d{4}[\-/]', s))          # date range "2016-11"
                or 'project:' in s.lower()
                or (len(s) > 45 and ' ' in s and not _SIDEBAR_URL.match(s)))  # prose

    out, skills, i = [], [], 0
    while i < n:
        s = lines[i].strip()
        if _SIDEBAR_CONTACT_HEAD.match(s):
            in_skills = False
            j = i + 1
            while j < n:
                sj = lines[j].strip()
                is_head = (_SIDEBAR_CONTACT_HEAD.match(sj) or _SIDEBAR_SKILLS_HEAD.match(sj)
                           or _SIDEBAR_OTHER_HEAD.match(sj))
                if sj and _resumes_work(sj) and not is_head:
                    break
                if _SIDEBAR_SKILLS_HEAD.match(sj):
                    in_skills = True
                elif _SIDEBAR_OTHER_HEAD.match(sj) or _SIDEBAR_CONTACT_HEAD.match(sj):
                    in_skills = False
                elif (in_skills and sj and not _SIDEBAR_URL.match(sj)
                      and not _SIDEBAR_EMAIL.match(sj) and not _SIDEBAR_PHONE.match(sj)):
                    skills.append(sj)
                j += 1
            i = j  # drop the whole run [i, j)
            continue
        out.append(lines[i])
        i += 1
    return '\n'.join(out).strip(), skills


def _merge_recovered_skills(skills_text, recovered):
    """Prepend sidebar-recovered skill items to the skills field, de-duplicated
    case-insensitively (ignoring punctuation), preserving existing entries."""
    if not recovered:
        return skills_text
    existing = [l for l in (skills_text or "").split('\n') if l.strip()]
    seen = {re.sub(r'[^a-z0-9]', '', l.lower()) for l in existing}
    merged = []
    for item in recovered:
        key = re.sub(r'[^a-z0-9]', '', item.lower())
        if key and key not in seen:
            seen.add(key)
            merged.append(item)
    return "\n".join(merged + existing).strip()


_SKILL_SIDEBAR_BULLET_CHARS = ("-", "•", "*", "▪", "●", "○", "◦", "‣", "»", "›", "·", "", "")


def _looks_like_skill_sidebar_header(line):
    """True for a short Title-Case line that reads like a skill category header
    ("Microsoft Office Suite", "Data Management") rather than a job/company line."""
    line = line.strip()
    if not line or line.startswith(_SKILL_SIDEBAR_BULLET_CHARS):
        return False
    # Real job/company lines are the main false-positive risk here — they're
    # often ALL CAPS and/or use a comma or pipe to separate title/company/
    # location (e.g. "SENIOR EXECUTIVE, STELIS BIOPHARMA, DODDABALLAPURA",
    # "Business Analyst | Salesforce Business Analyst"). Genuine skill
    # category headers never look like that, so reject both patterns outright.
    if "," in line or "|" in line:
        return False
    letters_only = re.sub(r"[^A-Za-z]", "", line)
    if letters_only and letters_only.isupper():
        return False
    if re.search(r"\d", line):
        return False
    if len(line) > 60:
        return False
    words = line.split()
    if not (1 <= len(words) <= 6):
        return False
    _minor = {"and", "or", "of", "for", "the", "in", "&", "to"}
    for w in words:
        if w.lower() in _minor:
            continue
        if not w[0].isupper():
            return False
    return words[0][0].isupper()


def _is_skill_sidebar_bullet_or_date_line(line):
    line = line.strip()
    if not line:
        return False
    if line.startswith(_SKILL_SIDEBAR_BULLET_CHARS):
        return True
    return bool(re.search(r"\b(19|20)\d{2}\b", line))


def _rescue_skill_sidebar_from_experience(experience_text):
    """Sidebar resumes sometimes repeat their Skills column on later pages
    WITHOUT repeating the "SKILLS" heading, so that continuation has no marker
    distinguishing it from whatever section is active at that point in the
    text — usually Experience. Detect a run of 3+ consecutive "Short
    Title-Case Header" + "description" pairs (no bullets, no dates in
    between) inside Experience and pull it out.

    Returns (cleaned_experience, rescued_skills_text). rescued_skills_text is
    "" when nothing looked like a skills sidebar (the common case).
    """
    exp_lines = (experience_text or "").splitlines()
    n_lines = len(exp_lines)
    skill_blocks = []
    si = 0
    while si < n_lines:
        if _looks_like_skill_sidebar_header(exp_lines[si]):
            start = si
            pair_count = 0
            sj = si
            while sj < n_lines and _looks_like_skill_sidebar_header(exp_lines[sj]):
                sj += 1
                desc_lines = 0
                while (sj < n_lines and exp_lines[sj].strip()
                       and not _looks_like_skill_sidebar_header(exp_lines[sj])
                       and not _is_skill_sidebar_bullet_or_date_line(exp_lines[sj])):
                    desc_lines += 1
                    ends_with_comma = exp_lines[sj].rstrip().endswith(",")
                    sj += 1
                    # Only keep consuming lines while the previous one trails off
                    # with a comma (a genuine mid-sentence line wrap, e.g.
                    # "...CARR,\nFEA, QC Charts..."). Without that signal, stop
                    # after one line — otherwise an unrelated stray fragment right
                    # after a complete sentence (e.g. a duplicated/truncated
                    # leftover line from the source document) gets swallowed into
                    # this entry's description.
                    if not ends_with_comma:
                        break
                if desc_lines == 0:
                    break
                pair_count += 1
                if sj < n_lines and _is_skill_sidebar_bullet_or_date_line(exp_lines[sj]):
                    break
            end = sj
            if pair_count >= 3:
                skill_blocks.append((start, end))
            si = end if end > si else si + 1
        else:
            si += 1

    if not skill_blocks:
        return experience_text, ""

    covered = set()
    for bs, be in skill_blocks:
        covered.update(range(bs, be))
    rescued = [l for i, l in enumerate(exp_lines) if i in covered]
    kept = [l for i, l in enumerate(exp_lines) if i not in covered]
    rescued_text = "\n".join(rescued).strip()
    if not rescued_text:
        return experience_text, ""
    return "\n".join(kept).strip(), rescued_text


_HEADING_TAIL_WORDS = frozenset({
    "me", "completed", "compelted", "compeleted", "courses", "course",
    "details", "detail", "summary", "information", "info", "profile",
    "done", "worked", "history", "background", "expertise",
})


def _strip_heading_tail_prefix(text):
    """Remove a leading orphan word that is the wrapped tail of a section heading.

    Two-word headings sometimes wrap so the first word matches as the heading and
    the second word lands as the section's first content line, e.g.
    "ABOUT ME" → summary starting with "ME", or "PROJECTS COMPELTED" → projects
    starting with "COMPELTED". This drops such a leading single-word remnant.
    """
    lines = text.split('\n') if text else []
    while lines:
        first = lines[0].strip()
        key = re.sub(r'[^a-z]', '', first.lower())
        if first and ' ' not in first and key in _HEADING_TAIL_WORDS:
            lines.pop(0)
        else:
            break
    return '\n'.join(lines).strip()


def _pymupdf_summary_block(path):
    """Recover a full-width summary/objective paragraph via pymupdf block layout.

    On 2-column PDFs, a justified full-width "About Me"/"Profile Summary" paragraph
    gets its lines split into left+right halves by the linear text extractor, so
    the summary field captures only the first halves and reads truncated. pymupdf's
    block extraction keeps the paragraph intact as ONE block, so we find the summary
    heading block and return the following paragraph block. Returns "" if no summary
    heading block is found or on any error — caller only uses it when it is longer
    than the summary it already has.
    """
    try:
        import pymupdf as fitz
    except Exception:
        return ""
    head = re.compile(
        r'^(about(\s+me)?|profile\s+summary|professional\s+summary|career\s+summary|'
        r'summary|objective|career\s+objective|professional\s+objective|profile|'
        r'professional\s+profile|career\s+profile|personal\s+profile)\s*:?\s*$', re.I)
    best = ""
    try:
        doc = fitz.open(str(path))
        for page in doc:
            blocks = [b for b in page.get_text("blocks") if b[6] == 0 and b[4].strip()]
            blocks.sort(key=lambda b: (round(b[1]), b[0]))
            for idx, b in enumerate(blocks):
                blines = [l.strip() for l in b[4].splitlines() if l.strip()]
                if not blines or not head.match(blines[0]):
                    continue
                rest = "\n".join(blines[1:]).strip()
                cand = rest if len(rest) > 40 else (
                    blocks[idx + 1][4].strip() if idx + 1 < len(blocks) else "")
                if len(cand) > len(best):
                    best = cand
        doc.close()
    except Exception as exc:
        logger.debug(f"_pymupdf_summary_block failed: {exc}")
        return ""
    return best


def _stitch_wrapped_urls(text):
    """Rejoin URLs broken across lines by narrow sidebar wrapping.

    A 2-column contact sidebar wraps a long URL mid-path, e.g.
        https://www.linkedin.com/i
        n/prasanth-vb-0017b016a
    Native extraction keeps those as two lines, so URL matching grabs a garbled
    value. This concatenates a continuation line onto the preceding URL line when
    the continuation starts lowercase/digit/'/' and is pure URL-path characters —
    which excludes ALL-CAPS sidebar headings ("WWW") and spaced text, so only a
    genuine wrapped-URL tail is joined. Intended for building a scratch string
    used ONLY for URL extraction, never the main resume text.
    """
    cont = re.compile(r'^[a-z0-9/][a-z0-9\-_%/.?=&#]*$')
    url_end = re.compile(r'(?:https?://|www\.|linkedin\.com|github\.com)\S*$', re.I)
    out = []
    for ln in text.split('\n'):
        s = ln.strip()
        if out and s and cont.match(s) and url_end.search(out[-1].strip()):
            out[-1] = out[-1].rstrip() + s
        else:
            out.append(ln)
    return '\n'.join(out)


# Section headings that sometimes bleed into the Skills column on 2-column PDFs —
# everything from such a line onward is the NEXT section, not skills.
_SKILLS_BLEED_HEADINGS = frozenset({
    "languages", "language", "language known", "languages known",
    "interests", "hobbies", "hobbies and interests", "declaration",
    "personal details", "other personal details", "personal information",
    "references", "extra curricular", "extracurricular",
})

# Unlike the exact-match headings above, "Websites, Portfolios and Profiles" is
# commonly fused with its own URL on the same line by a 2-column sidebar (e.g.
# "WEBSITES, PORTFOLIOS https://www.linkedin.com/in/..."), so it can never
# equal a clean heading string outright — it must be matched as a LEAD-IN
# instead of a whole-line match.
_SKILLS_BLEED_PREFIXES = (
    "websites portfolios", "websites and profiles", "websites portfolio",
    "website portfolio", "and profiles",
)


def _truncate_skills_at_bleed(lines):
    """Cut the skills list at the first line that is actually the heading of a
    following section (e.g. a 'Languages' sidebar block read straight after the
    skills, or a 'Websites, Portfolios' sidebar heading fused with its own URL
    on the same line). Removes 'extra information apart from skills' without
    touching real skill entries above the boundary."""
    out = []
    for l in lines:
        key = re.sub(r'[^a-z ]', ' ', l.strip().lower())
        key = re.sub(r'\s+', ' ', key).strip()
        if key in _SKILLS_BLEED_HEADINGS or key.startswith(_SKILLS_BLEED_PREFIXES):
            break
        out.append(l)
    return out


_SKILL_FILLER_LEAD = re.compile(
    r'^\s*(?:proficient(?:\s+in)?|skilled(?:\s+in)?|experienced(?:\s+(?:in|with))?|'
    r'expertise(?:\s+in)?|professional\s+experience(?:\s+in)?|knowledgeable(?:\s+in)?|'
    r'hands[-\s]on\s+experience(?:\s+(?:in|with))?|good\s+knowledge\s+of|'
    r'strong\s+knowledge\s+of|working\s+knowledge\s+of|familiar\s+with|exposure\s+to|'
    r'advanced|provided|qualified)\b[\s:]*', re.I)

_SKILL_JUNK_TOKENS = frozenset({
    "editing", "review", "preparation", "document preparation", "management",
    "experience", "knowledge", "user interface", "high-performing teams",
    "compliance", "efficiency", "medical devices",
})


def _looks_like_prose_skills(skills_text):
    """True when the skills section is written as category labels + prose SENTENCES
    (e.g. 'Proficient in PTC CREO, Solid Works, ...') rather than atomic entries.
    Deliberately narrow so already-clean skill lists are left completely untouched.

    Gated on genuine filler-verb evidence ("Proficient in", "Skilled in", ...) only —
    NOT on line length. A resume that lists "Category – item1, item2, item3" or
    "Category: item1, item2" per bullet is already a clean, structured skills list
    (just naturally long because it enumerates many items); it must be displayed
    as-is, one bullet per category, not shredded into individual tokens. Line
    length alone can't tell those two formats apart, but the presence of filler
    verbs can.
    """
    lines = [l.strip() for l in (skills_text or "").split('\n') if l.strip()]
    if len(lines) < 4:
        return False
    filler = sum(1 for l in lines if _SKILL_FILLER_LEAD.match(l))
    return filler >= 2


def _split_desc_to_skills(line):
    s = re.sub(r'\([^)]*\)', ' ', line)                    # drop parentheticals
    items = []
    # Split on list separators. `and`/`&` require surrounding spaces so intra-token
    # forms like "GD&T", "R&D", "NX CAD/CAM" are never broken apart. A dash used as
    # a "Category – detail1, detail2" separator (spaces on both sides) is split too,
    # so the category label and its first item don't glue into one over-long chunk
    # that then fails the length/word-count filters below and gets dropped entirely.
    # Requiring surrounding spaces keeps tight hyphenated compounds intact
    # ("Stack-up analysis", "SAP-E Chain").
    for p in re.split(r'[;,:]|\s+and\s+|\s+&\s+|\s+[–—-]\s+', s, flags=re.I):
        p = _SKILL_FILLER_LEAD.sub('', p).strip(' .:-\t')  # strip leading filler
        # A bare leading preposition ("in Power BI" left behind when an upstream
        # heading-alias match consumed only "Professional Experience" from
        # "Professional experience in Power BI...") is never itself a skill token.
        p = re.sub(r'^(?:in|with|for|using|of)\s+', '', p, flags=re.I).strip()
        if not p or len(p) < 2 or len(p) > 40 or len(p.split()) > 4:
            continue
        if re.search(r'\b(for|to|with)\b', p, re.I):       # drop purpose clauses
            continue
        if p.lower() in _SKILL_JUNK_TOKENS or _SKILL_FILLER_LEAD.match(p):
            continue
        items.append(p)
    return items


def _extract_skill_keywords(skills_text):
    """Turn a category-label + prose skills section into atomic, one-per-line skill
    keywords. No-op unless the section is prose-formatted (see _looks_like_prose_
    skills), so resumes that already list clean skills are unaffected.

    Every line is split (both the bold category labels and their descriptions),
    because a label often carries real skill names too (e.g. 'TrackWise and
    Salesforce'); de-duplication then collapses repeats.
    """
    if not _looks_like_prose_skills(skills_text):
        return skills_text
    out = []
    for line in skills_text.split('\n'):
        out.extend(_split_desc_to_skills(line))
    seen, result = set(), []
    for s in out:
        k = re.sub(r'[^a-z0-9]', '', s.lower())
        if k and k not in seen:
            seen.add(k)
            result.append(s)
    return '\n'.join(result) if result else skills_text


def _drop_trailing_skill_fragments(lines):
    """Drop trailing single-word gutter-bleed fragments from a skills list.

    On 2-column PDFs, a skill that wraps in the sidebar (e.g. "Medical Device
    Classification") can leave its leading word ("Medical") stranded as its own
    line at the very END of the skills column, right before the next section.
    Such an orphan is spurious — it only repeats the first word of a real,
    multi-word skill already in the list. We remove a trailing single-word line
    ONLY when an earlier line actually starts with that same word followed by a
    space, so genuine standalone one-word skills (e.g. "Python", "Selenium") are
    never touched.
    """
    lines = list(lines)
    while len(lines) >= 2:
        last = lines[-1].strip()
        core = re.sub(r"[&/\-.]", "", last)
        is_single_word = last and " " not in last and core.isalpha()
        repeats_lead = any(
            o.strip().lower().startswith(last.lower() + " ") for o in lines[:-1]
        )
        if is_single_word and repeats_lead:
            lines.pop()
        else:
            break
    return lines


_SKILL_SOCIAL_FRAG_RE = re.compile(
    r'^(?:n?ked[Ii]n|[Ll]inked[Ii]n|[Ll]inked?|[Ff]acebook|[Tt]witter|'
    r'[Ii]nstagram|[Yy]outube|[Gg]ithub|[Gg]itlab)\s*$'
)
_SKILL_ROLE_ONLY_RE = re.compile(r'^[A-Z][a-z]+(?: [A-Z][a-z]+){0,3}$')  # Title-case, 1-4 words
_SKILL_ROLE_KW_RE = re.compile(
    r'\b(analyst|engineer|manager|consultant|officer|executive|director|'
    r'lead|specialist|developer|designer|coordinator|associate|intern)\b',
    re.I,
)


def _drop_name_title_and_gutter_bleed(lines, full_name, title):
    """Remove the candidate's own name/title and column-gutter-bleed fragments
    from a skills list — a 2-column layout can place the resume's own name/title
    header at the same y-position as the tail of the Skills sidebar, so the
    linear text extractor reads it as if it were another skill line.

    - Drops a line that exactly matches the already-extracted full_name/title.
    - Drops bare 1-2 char fragments (never a valid skill).
    - Drops social-media link truncations ("nkedIn", "acebook", ...).
    - Drops a bare Title-Case 1-4 word role phrase containing a job-role keyword
      (e.g. "Design Engineer") with no tool/tech markers — this catches a leaked
      title even when the `title` field itself wasn't extracted correctly.
    """
    _name = (full_name or "").strip().upper()
    _title = (title or "").strip().lower()
    out = []
    for l in lines:
        s = l.strip()
        if not s:
            continue
        if s.upper() == _name or (_title and s.lower() == _title):
            continue
        # Bare 1-2 char fragments are gutter-bleed noise ("s", "li" left over from
        # a wrapped word) — but only when purely alphabetic. A short alphanumeric
        # code ("8D", "5S") is a real, well-known quality-methodology skill.
        if len(s) <= 2 and s.isalpha():
            continue
        if _SKILL_SOCIAL_FRAG_RE.match(s):
            continue
        if (_SKILL_ROLE_ONLY_RE.match(s) and _SKILL_ROLE_KW_RE.search(s)
                and not re.search(r'[/(]|\d', s)):
            continue
        out.append(l)
    return out


def _looks_like_stray_experience_bullet(line):
    """True for a bulleted line that reads like a job-responsibility SENTENCE
    ("• Reviewed and approved V&V protocols...") rather than a short Skills
    category item ("• DHF / DMR Authorship & Review"). Used to detect Experience
    content stranded past a Skills heading — bullet punctuation alone isn't a
    reliable signal, since plenty of legitimate Skills sections bullet their own
    short items.
    """
    s = line.strip()
    if not s.startswith(("•", "●", "▪", "◦")):
        return False
    body = s.lstrip("•●▪◦").strip()
    if not body:
        return False
    if len(body) > 60:
        return True
    return body[-1:] in ".;:"


def parse_resume_text(text, name_hint=None):
    lines = normalize_lines(text)
    parsed = {
        "full_name": "", "title": "", "email": "", "phone": "", "linkedin": "", "location": "",
        "summary": "", "skills": "", "experience": "", "education": "",
        "certifications": "", "projects": "",
    }
    if not lines:
        return parsed

    # Font-extracted name from PDF metadata takes top priority — skip all other name logic.
    if name_hint and _looks_like_name(name_hint):
        parsed["full_name"] = name_hint[:80]

    # ── Email extraction ──────────────────────────────────────────────────────
    # Strategy: scan every line for a token containing @; require a dot after @.
    parsed["email"] = _extract_email(text)

    # ── Phone extraction ──────────────────────────────────────────────────────
    # Strategy: scan every line; take first candidate whose digit-only count >= 10.
    parsed["phone"] = _extract_phone(text)

    # Join LinkedIn URLs that wrap across lines. Sidebars wrap mid-path anywhere
    # (e.g. "linkedin.com/i\nn/prasanth-vb..." or "malthesh-\nkarnam-29a6b416a"),
    # so stitch general wrapped-URL tails first, then the hyphen-specific case.
    _lnk_text = _stitch_wrapped_urls(text)
    _lnk_text = re.sub(
        r'(linkedin\.com/in/[a-zA-Z0-9\-_%]*)-[ \t]*\n[ \t]*([a-zA-Z0-9])',
        lambda m: m.group(1) + '-' + m.group(2),
        _lnk_text, flags=re.I,
    )
    linkedin_match = re.search(
        r"(?:https?://)?(?:www\.)?linkedin\.com/in/([a-zA-Z0-9\-_%]+)", _lnk_text, re.I
    )
    if linkedin_match:
        parsed["linkedin"] = "https://www.linkedin.com/in/" + linkedin_match.group(1)
    else:
        # Bare handle format: "LinkedIn: john-doe" or "LinkedIn ID: john-doe-123abc"
        _bare_match = re.search(
            r'(?:linkedin|linked\s*in)(?:\s*id|\s*profile|\s*handle)?[\s:]+([a-zA-Z0-9][a-zA-Z0-9\-_%]{2,})',
            text, re.I,
        )
        if _bare_match:
            _handle = _bare_match.group(1).strip('-').strip()
            # Must look like a handle (not a common word)
            if len(_handle) >= 3 and not re.match(r'^(profile|handle|id|url|link|page|account)$', _handle, re.I):
                parsed["linkedin"] = "https://www.linkedin.com/in/" + _handle

    # Remove all LinkedIn URLs and fragments from text
    text = re.sub(r'https?://[^\s]*linkedin\.com[^\s]*', '', text, flags=re.I)
    text = re.sub(r'www\.linkedin\.com[^\s]*', '', text, flags=re.I)
    text = re.sub(r'linkedin\.com/in/[a-zA-Z0-9\-_%/@?=&#]*', '', text, flags=re.I)
    text = re.sub(r'\?skipRedirect[^\s]*|\&skipRedirect[^\s]*|#skipRedirect[^\s]*', '', text, flags=re.I)
    # Remove standalone LinkedIn profile ID fragments (e.g., "karnam-29a6b416a/")
    text = re.sub(r'^\s*[a-z]+\-[a-z0-9]+/?$', '', text, flags=re.MULTILINE | re.I)

    lines = [line.strip() for line in text.split('\n')]
    lines = [line for line in lines if line]

    if not parsed["full_name"]:
        parsed["full_name"] = parse_label_value(lines, ["name", "full name", "candidate name"])
    parsed["title"] = parse_label_value(lines, ["title", "designation", "role", "current role"])
    parsed["location"] = parse_label_value(lines, ["location", "address", "city"])

    # Pull name from a header line like "John Smith   Mobile: +91 xxx   E-Mail: yyy"
    # Must run before useful_top_lines filter, which would discard such lines.
    if not parsed["full_name"]:
        parsed["full_name"] = _extract_name_from_header_line(lines)

    _email_fragments = set()
    if parsed.get("email"):
        _em = parsed["email"]
        for _tld_len in range(1, 5):
            _frag = _em[-_tld_len:]
            if _frag.isalpha():
                _email_fragments.add(_frag.lower())

    ignored = ["resume", "curriculum vitae", "cv"]
    useful_top_lines = [
        line for line in lines[:20]
        if line.lower() not in ignored
        and line.lower().strip() not in _email_fragments
        and not (len(line.strip()) <= 4 and line.strip().isalpha() and line.strip().islower())
        and not re.search(
            r"@|www\.|linkedin|github|\d{6,}|\+\d{1,3}[\s\-]?\d{3,}|(?<!\w)\d{10,}"
            r"|date.{0,5}birth|nationality|dob\b|gender|marital|\d{1,2}[./]\d{1,2}[./]\d{2,4}"
            r"|total.{0,10}experience|current employer|current designation|qualification|designation"
            # standalone contact-sidebar labels that must never be mistaken for job titles
            r"|^email$|^mobile$|^phone$|^telephone$|^tel$|^fax$|^social\s+links?$"
            # experience-duration metadata lines like "7 Years 0 Month"
            r"|\d+\s+years?\s+\d+\s+months?",
            line, re.I,
        )
    ]
    # Name fallback: look for a line that actually looks like a person's name.
    # Prevents professional summary sentences from being assigned as the name.
    # Prioritize all-caps names (likely to be actual names in resume headers)
    #
    # Never consider a line at/after the first recognized section heading (e.g.
    # a "CORE COMPETENCIES" sub-heading like "Document Control") as a name
    # candidate — mirrors the same section-heading boundary already used for
    # title extraction just below, applied here too since a short Title-Case
    # skill-category label can otherwise pass the same 2-word name heuristic
    # as a real name once scanning runs past the resume's actual header.
    _name_scan_limit = next(
        (i for i, l in enumerate(useful_top_lines) if canonical_section_name(l)),
        len(useful_top_lines),
    )
    if not parsed["full_name"]:
        # Pass 0: the name is sometimes combined with a degree/title on one line
        # (e.g. "GURUVIGNESH Y - B.E MBA(PM) Senior Engineer (R&D Medical Device)").
        # The parentheses/hyphen make the whole line fail the name check, but the
        # segment before the first " - "/"–"/"|" separator is often just the name.
        for _nl in useful_top_lines[:min(3, _name_scan_limit)]:
            _seg = re.split(r'\s+[-–|]\s+', _nl, maxsplit=1)[0].strip()
            if _seg and _seg != _nl and _looks_like_name(_seg):
                parsed["full_name"] = _seg[:80]
                break
        # First pass: look for all-caps names in FIRST 10 lines (e.g., "SIVARANJANI D")
        if not parsed["full_name"]:
            for _nl in useful_top_lines[:min(10, _name_scan_limit)]:
                _nl_clean = _strip_degree_suffix(_nl)
                if _nl_clean.isupper() and _looks_like_name(_nl_clean):
                    parsed["full_name"] = _nl_clean[:80]
                    break
        # First pass: look for all-caps names in FIRST 10 lines (e.g., "SIVARANJANI D")
        if not parsed["full_name"]:
            for _nl in useful_top_lines[:min(10, _name_scan_limit)]:
                if _nl.isupper() and _looks_like_name(_nl):
                    parsed["full_name"] = _nl[:80]
                    break
        # Second pass: look for TWO-WORD names (common name pattern) in first 15 lines
        if not parsed["full_name"]:
            for _nl in useful_top_lines[:min(15, _name_scan_limit)]:
                _nl_clean = _strip_degree_suffix(_nl)
                _words = _nl_clean.split()
                if len(_words) == 2 and _looks_like_name(_nl_clean):
                    parsed["full_name"] = _nl_clean[:80]
                    break
        # Third pass: look for any proper-cased name in all useful lines
        if not parsed["full_name"]:
            for _nl in useful_top_lines[:_name_scan_limit]:
                _nl_clean = _strip_degree_suffix(_nl)
                if _looks_like_name(_nl_clean):
                    parsed["full_name"] = _nl_clean[:80]
                    break
        if not parsed["full_name"]:
            # Looser pass: short proper-cased line with 2-5 words and no stopwords
            for _nl in useful_top_lines[:_name_scan_limit]:
                _nwords = _nl.split()
                if (2 <= len(_nwords) <= 5 and _nl[:1].isupper() and len(_nl) <= 50
                        and not {w.lower().strip(".,;:-") for w in _nwords} & _NAME_STOPWORDS):
                    parsed["full_name"] = _nl[:80]
                    break
    # Title priority 1: extract from first paragraph sentence pattern
    # ("Validation and Compliance Lead with 7 years...") — done FIRST so sentence-embedded
    # titles beat ambiguous short lines from the top-lines scan.
    if not parsed["title"]:
        parsed["title"] = _extract_title_from_para(text[:700])
    # Title priority 2: first short proper-cased line after the name that reads
    # like a job title (not a summary sentence, not a section heading).
    # Scoped to the true preamble (lines before the first recognized section
    # heading) rather than the general-purpose useful_top_lines[:20] window —
    # once real section content starts (Skills, Personal Details, etc.), a
    # short Title-Case line found there is a sub-heading, not a job title, and
    # if the preamble genuinely has no title-like line, it's better to leave
    # the field blank than guess from unrelated section content.
    if not parsed["title"]:
        _first_heading_idx = next(
            (i for i, l in enumerate(lines[:20]) if canonical_section_name(l)),
            20,
        )
        _title_scan_lines = [l for l in useful_top_lines if l in lines[:_first_heading_idx]]
        _name_line = parsed.get("full_name", "")
        for _tl in _title_scan_lines:
            if _tl == _name_line:
                continue
            # Skip our own OCR-injected "Name:/Email:/Phone:/Location:" label
            # lines — they must never be mistaken for a job title.
            if re.match(r'^\s*(name|email|phone|location)\s*:', _tl, re.I):
                continue
            # Long taglines often bundle "<Title> | <keyword> | <keyword>..."
            # or "<Title> – <keyword list>" (e.g. "QA Specialist – Design
            # Control | Medical Device & IVD Quality") — the combined line
            # fails _looks_like_title's word-count check even though the
            # real title segment before the first separator is perfectly
            # valid on its own. Same separator pattern as the name Pass 0
            # split above.
            _tl_seg = re.split(r'\s+[-–|]\s+', _tl, maxsplit=1)[0].strip()
            if _tl_seg and _tl_seg != _tl and not canonical_section_name(_tl_seg) and _looks_like_title(_tl_seg):
                parsed["title"] = _tl_seg[:80]
                break
            if not canonical_section_name(_tl) and _looks_like_title(_tl):
                parsed["title"] = _tl[:80]
                break

    if not parsed["location"]:
        _loc_city_re = re.compile(
            r"\b(india|karnataka|bangalore|bengaluru|chennai|hyderabad|pune|mumbai|delhi|"
            r"noida|gurgaon|gurugram|coimbatore|trivandrum|kochi|jaipur|ahmedabad|kolkata|"
            r"ludhiana|chandigarh)\b", re.I
        )
        for line in lines[:20]:
            if _loc_city_re.search(line):
                # Split on pipe/bullet separators first; take only the segment
                # that contains the city keyword — avoids grabbing "Open to
                # Relocation", "LinkedIn:", years-of-experience text, etc.
                _segs = re.split(r'\s*[|•·]\s*', line)
                _city_seg = next((s for s in _segs if _loc_city_re.search(s)), line)
                # Strip email, phone, and icon characters from the city segment
                loc_line = re.sub(r'[\w.+-]+@[\w.-]+\.[a-z]{2,}', '', _city_seg, flags=re.I)
                loc_line = re.sub(r'[\+\(]?[\d\s\-\(\)]{7,}', '', loc_line)
                loc_line = re.sub(r'[^\x20-\x7E]', ' ', loc_line)
                loc_line = re.sub(r'[,;]+', ',', loc_line)
                loc_line = re.sub(r',\s*,', ',', loc_line)
                loc_line = ', '.join(p.strip() for p in loc_line.split(',') if p.strip())
                if loc_line and len(loc_line) > 2:
                    parsed["location"] = loc_line[:180]
                break

    parsed.update(find_sections(lines))

    # ── Post-process: Extract skills from 2-column layouts ──────────────────────
    # In some PDFs (e.g., Sivaranjani's), skills are interleaved with summary due to
    # column extraction. If skills section is missing, extract from summary AND certifications.
    if not parsed.get("skills"):
        skills_items = []

        # Collect lines from multiple sections that might contain skills
        sections_to_scan = []
        for _sec_key in ("summary", "certifications", "training", "achievements", "career_highlights"):
            if parsed.get(_sec_key):
                sections_to_scan.extend(parsed[_sec_key].split('\n'))

        for line in sections_to_scan:
            line_stripped = line.strip()

            if not line_stripped or len(line_stripped) < 5:
                continue

            # Must start with capital letter
            if not line_stripped[0].isupper():
                continue

            # Skip full sentences that end with periods
            if line_stripped.endswith('.'):
                continue

            # Extract skill, handling cases where skill + summary text appear on same line
            # Pattern: skill text ends where summary paragraph begins (indicated by "Experienced", "professional", etc.)
            skill_part = line_stripped
            if any(marker in line_stripped for marker in [' Experienced ', ' professional ', ' compliance ', ' and compliance']):
                # Find the skill part before summary text
                for marker in [' Experienced ', ' professional ']:
                    if marker in line_stripped:
                        skill_part = line_stripped.split(marker)[0].strip()
                        break

            # Pattern 1: Ends with & or ( (incomplete lines are skills continuing to next line)
            if skill_part.endswith('&') or skill_part.endswith('('):
                skills_items.append(skill_part)
                continue

            # Pattern 2: Contains skill indicators (&, /, parens, colons) and is reasonably short
            has_skill_indicator = any(c in skill_part for c in ['&', '/', '(', ')', ':'])
            if has_skill_indicator and len(skill_part) <= 90:
                # Accept if it has the skill pattern
                word_count = len(skill_part.split())
                if word_count <= 8:  # Allow slightly longer for multi-part skills
                    skills_items.append(skill_part)
                    continue

            # Pattern 3: Short items (1-2 words, 5-40 chars) - likely single skill names
            word_count = len(skill_part.split())
            if 1 <= word_count <= 2 and 5 <= len(skill_part) <= 40:
                skills_items.append(skill_part)
                continue

            # Pattern 4: Multi-word items with acronyms/standards (3-6 words with & or /)
            # Examples: "Medical Device Lifecycle Management & UDI", "ASTM F1980 / ISO 20417"
            if 3 <= word_count <= 6 and (has_skill_indicator or any(c.isdigit() for c in skill_part)):
                # Likely a skill with special formatting
                if len(skill_part) <= 90:
                    skills_items.append(skill_part)
                    continue

            # Pattern 5: Lines with acronyms/numbers that look like skills even without & or /
            # Examples: "MDD & EU MDR 745/2017", "ASTM F1980 / ISO 20417"
            # Accept short-medium lines (up to 5 words) that have numbers and caps
            if 2 <= word_count <= 5 and len(skill_part) <= 90:
                has_number = any(c.isdigit() for c in skill_part)
                has_caps = any(w.isupper() for w in skill_part.split())
                if has_number and has_caps and not skill_part.endswith(' that') and not skill_part.endswith(' and'):
                    skills_items.append(skill_part)

        if skills_items:
            parsed["skills"] = "\n".join(skills_items)
            # Remove extracted skills from summary to clean it up
            if parsed.get("summary"):
                summary_lines = parsed["summary"].split('\n')
                summary_cleaned = '\n'.join(l for l in summary_lines if l.strip() not in skills_items)
                parsed["summary"] = summary_cleaned.strip()

    # Capture the unheaded intro paragraph that precedes the first section heading.
    # Use it only when no explicitly labelled summary section was found; explicit
    # section headings like "PROFILE SUMMARY" take priority.
    _intro = _extract_intro_paragraph(lines)
    if _intro and not parsed.get("summary"):
        parsed["summary"] = _intro

    # Fallback: if summary is still empty, do a direct line-by-line scan for a
    # "Profile Summary" or "Professional Summary" heading and collect what follows.
    # This covers PDFs where the heading's special formatting (decorative chars,
    # 2-column extraction order, etc.) prevented the section parser from capturing it.
    if not parsed.get("summary"):
        _in_ps, _ps_lines = False, []
        for _ln in lines:
            _key = re.sub(r"[^a-zA-Z ]", " ", _ln).lower()
            _key = re.sub(r"\s+", " ", _key).strip()
            if _key in ("profile summary", "professional summary"):
                _in_ps = True
                continue
            if _in_ps:
                _sec = canonical_section_name(_ln)
                if _sec and _sec != "summary":
                    break
                if _ln.strip():
                    _ps_lines.append(_ln.strip())
        if _ps_lines:
            parsed["summary"] = "\n".join(_ps_lines)

    if parsed.get("skills"):
        _url_line_pat = re.compile(
            r"^(https?://|www\.|linkedin\.com|github\.com|gitlab\.com|portfolio|website|blog|"
            r"twitter|facebook|instagram|behance|dribbble|stackoverflow|medium\.com|kaggle)\S*$",
            re.I,
        )
        _portfolio_label_pat = re.compile(
            r"^(portfolio|website|blog|github|gitlab|twitter|facebook|instagram|behance|"
            r"dribbble|stackoverflow|medium|kaggle|profile|link|url)\s*[:\-]",
            re.I,
        )
        _non_skill_heading_pat = re.compile(
            r"^(websites?|portfolios?|websites?[,\s]+portfolios?|websites?.*profiles?|"
            r"and\s+profiles?|profiles?\s+and|"
            r"personal|references?|declaration|additional information|"
            r"other information|education|experience|certifications?|projects?)\s*$",
            re.I,
        )
        _inline_url_pat = re.compile(r"https?://\S+|www\.\S+", re.I)
        clean_skill_lines = []
        for sl in parsed["skills"].splitlines():
            stripped_sl = sl.strip()
            if _url_line_pat.match(stripped_sl):
                continue
            if _portfolio_label_pat.match(stripped_sl):
                continue
            bare = stripped_sl.strip("•-| ")
            if _url_line_pat.match(bare) or _inline_url_pat.fullmatch(bare):
                continue
            cleaned_sl = _inline_url_pat.sub("", sl).strip(" •-,|")
            if not cleaned_sl:
                continue
            if _non_skill_heading_pat.match(re.sub(r"[^a-zA-Z\s]", " ", cleaned_sl).strip()):
                continue
            clean_skill_lines.append(cleaned_sl)
        clean_skill_lines = _truncate_skills_at_bleed(clean_skill_lines)
        clean_skill_lines = _drop_trailing_skill_fragments(clean_skill_lines)
        # Remove the person's own name/title and column-gutter-bleed fragments (e.g.
        # "Design Engineer" bleeding in from the right-column header at the same
        # y-level as the tail of the Skills sidebar). Must run on whole, not-yet-
        # atomized lines — a short legitimate skill code ("8D", "5S") only looks
        # like a "bare 1-2 char fragment" AFTER the keyword-extraction step below
        # splits prose into individual tokens, so this has to run first.
        clean_skill_lines = _drop_name_title_and_gutter_bleed(
            clean_skill_lines, parsed.get("full_name", ""), parsed.get("title", "")
        )
        parsed["skills"] = "\n".join(clean_skill_lines).strip()

    if parsed.get("education"):
        _personal_pat = re.compile(
            r"\b(father|mother|parent|guardian|date.{0,5}birth|born|dob\b|gender|sex\b|"
            r"marital|nationality|citizen|passport|religion|caste|languages known|"
            r"permanent address|current address|pin\s*code|aadhar|pan\b|"
            r"place of birth|age\b|blood group|reference)\b",
            re.I,
        )
        edu_lines = [l for l in parsed["education"].splitlines() if not _personal_pat.search(l)]
        parsed["education"] = "\n".join(edu_lines).strip()

    _TITLE_REJECTS = {
        "summary", "professional summary", "profile", "career profile",
        "professional profile", "objective", "career objective", "overview",
        "about", "about me", "executive summary",
    }
    if not parsed["title"] and parsed["summary"]:
        first_line = parsed["summary"].strip().split("\n")[0].strip()
        title_match = re.match(r"^((?:[A-Z][a-zA-Z]*(?:\s+|$)){1,5})", first_line)
        if title_match:
            candidate = title_match.group(1).strip()
            if 1 <= len(candidate.split()) <= 5 and candidate.lower() not in _TITLE_REJECTS:
                parsed["title"] = candidate

    if not parsed["title"] and parsed["experience"]:
        for exp_line in parsed["experience"].splitlines()[:5]:
            exp_line = exp_line.strip()
            if exp_line and "|" not in exp_line and not re.search(r"\d", exp_line):
                title_match = re.match(r"^((?:[A-Z][a-zA-Z]*(?:\s+|$)){1,6})$", exp_line)
                if title_match:
                    candidate = title_match.group(1).strip()
                    if 1 <= len(candidate.split()) <= 6:
                        parsed["title"] = candidate
                        break

    _TITLE_REJECTS_SET = {
        "summary", "professional summary", "profile", "career profile",
        "professional profile", "objective", "career objective", "overview",
        "about", "about me", "executive summary", "skills", "experience",
        "education", "certifications", "projects",
        # Single descriptor words that are NOT job titles
        "highly", "skilled", "experienced", "dedicated", "motivated",
        "dynamic", "proactive", "seasoned", "accomplished", "passionate",
    }
    if parsed["title"]:
        title_clean = parsed["title"].strip()
        title_clean = re.split(r"[,|/\\–—]", title_clean)[0].strip()
        title_clean = re.sub(r"[\(\[].*$", "", title_clean).strip()
        title_clean = re.sub(
            r"^(summary|professional\s+summary|profile|career\s+profile|"
            r"professional\s+profile|objective|career\s+objective|overview|"
            r"about\s+me?|executive\s+summary)\s+",
            "", title_clean, flags=re.I,
        ).strip()
        title_clean = re.sub(
            r"^(experienced|skilled|dedicated|results.driven|dynamic|seasoned|"
            r"highly experienced|passionate|motivated|proactive|hands.on)\s+",
            "", title_clean, flags=re.I,
        ).strip()
        if title_clean.lower() in _TITLE_REJECTS_SET:
            title_clean = ""
        if title_clean:
            words = title_clean.split()
            if len(words) > 5:
                m = re.match(r"^((?:[A-Z][a-zA-Z]*(?:\s+|$)){1,5})", title_clean)
                title_clean = m.group(1).strip() if m else " ".join(words[:5])
        parsed["title"] = title_clean[:80]

    # Map the extracted/cleaned title to the closest known group role
    if parsed["title"]:
        parsed["title"] = _map_to_group_role(parsed["title"])

    # Recover a skills sidebar block that a 2-column layout interleaved into Work
    # Experience: an unlabeled run of "category header" + "description" pairs with
    # no distinguishing heading of its own (common when the sidebar's Skills column
    # continues onto a later page).
    if parsed.get("experience"):
        _clean_exp, _rescued_text = _rescue_skill_sidebar_from_experience(parsed["experience"])
        if _rescued_text:
            parsed["skills"] = (parsed["skills"] + "\n" + _rescued_text) if parsed.get("skills") else _rescued_text
            parsed["experience"] = _clean_exp

    # Recover a contact/skills sidebar block (Contact/Phone/Email/LinkedIn + a
    # Skills list) that a 2-column layout interleaved into Work Experience.
    if parsed.get("experience"):
        _clean_exp, _recovered_skills = _extract_sidebar_bleed_from_experience(parsed["experience"])
        if _recovered_skills or _clean_exp != parsed["experience"]:
            parsed["experience"] = _clean_exp
            if _recovered_skills:
                parsed["skills"] = _merge_recovered_skills(parsed.get("skills", ""), _recovered_skills)

    # Strip resume-builder-template watermarks (e.g. "webuildcv.com", "Powered by
    # ...") that land as a bare branding line inside whichever section they
    # happen to fall in — most commonly the tail of Skills.
    for _wm_field in ("summary", "skills", "experience", "education",
                      "certifications", "projects"):
        _wm_text = parsed.get(_wm_field)
        if _wm_text:
            _wm_lines = [l for l in _wm_text.splitlines() if not _WATERMARK_LINE_RE.match(l.strip())]
            parsed[_wm_field] = "\n".join(_wm_lines).strip()

    return parsed


def merge_resume_data(form_data, parsed_data, overwrite=False):
    merged = dict(form_data)
    for key, value in parsed_data.items():
        if value and (overwrite or not merged.get(key)):
            merged[key] = value.strip()
    return merged


# ── Ollama "Resume Intelligence" parser (text LLM) ─────────────────────────────

_OLLAMA_BASE = os.environ.get(
    "OLLAMA_BASE_URL",
    "http://localhost:11434"
)

_TEXT_MODEL = os.environ.get(
    "OLLAMA_TEXT_MODEL",
    "llama3.2:latest"
)

_OLLAMA_TEXT_TIMEOUT = int(
    os.environ.get("OLLAMA_TEXT_TIMEOUT", "60")
)

_OLLAMA_KEEP_ALIVE = os.environ.get("OLLAMA_KEEP_ALIVE", "30m")

_OLLAMA_BUDGET_SECS = int(
    os.environ.get("OLLAMA_BUDGET_SECS", "10")
)

import requests


def get_embedding(text):
    """
    Generate embedding vector using Ollama's
    nomic-embed-text model.
    """

    if not text:
        return []

    response = requests.post(
        f"{_OLLAMA_BASE}/api/embed",
        json={
            "model": "nomic-embed-text",
            "input": text[:8000]
        },
        timeout=20
    )

    response.raise_for_status()

    result = response.json()

    if not result.get("embeddings"):
        return []

    return result["embeddings"][0]
# ── Pure-Ollama extraction — NO regex field engine in this path ────────────────
#
# A 1B model cannot one-shot a whole resume (it drifts, mislabels, or returns empty
# above ~1 000 chars of input).  The reliable pattern, confirmed by testing, is:
#   • feed it small, homogeneous pieces, and
#   • use flat JSON for short identity fields, plain text for long sections.
# So we slice the (column-split-cleaned) text into section chunks by heading, then
# make one small Ollama call per piece.  Ollama does ALL the field extraction; the
# only non-LLM step is cutting the text at heading lines.

_BLANK_FIELDS = (
    "full_name", "title", "email", "phone", "linkedin", "location",
    "summary", "skills", "experience", "education", "certifications", "projects",
)

# Heading text → which field its content belongs to. Extends the canonical section
# map with a few all-caps headings small models' resumes use.
_HEADING_FIELD = {
    "technology": "skills", "technologies": "skills", "technical skills": "skills",
    "publications": "certifications", "awards": "certifications",
    "achievements": "certifications", "certificates": "certifications",
}


def _ollama_chat(prompt, *, as_json, num_predict, num_ctx=8192, timeout=None):
    """Single Ollama call. Returns the raw assistant string (or {} dict if as_json).

    `timeout` overrides the module-wide _OLLAMA_TEXT_TIMEOUT default for
    callers that are known to legitimately run longer — kept opt-in so every
    other existing caller (resume parsing, identity extraction, etc.) keeps
    its original 60s ceiling unchanged.
    """
    body = {
        "model": _TEXT_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "options": {"temperature": 0, "num_predict": num_predict, "num_ctx": num_ctx},
        "keep_alive": _OLLAMA_KEEP_ALIVE,
    }
    if as_json:
        body["format"] = "json"
    url = f"{_OLLAMA_BASE}/api/chat"
    logger.info(f"Ollama call → {url} model={body['model']}")
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout or _OLLAMA_TEXT_TIMEOUT) as r:
        content = json.loads(r.read()).get("message", {}).get("content", "").strip()
    if not as_json:
        return content
    try:
        return json.loads(content)
    except (ValueError, TypeError):
        return {}


_IDENTITY_FIELDS = ("full_name", "email", "phone", "location", "linkedin")


def _validate_identity_field(field, value):
    """Sanity-check one LLM-returned identity field; return "" if implausible."""
    value = str(value or "").strip()
    if not value:
        return ""
    if field == "email":
        return value if re.match(
            r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$', value
        ) else ""
    if field == "phone":
        digits = re.sub(r"\D", "", value)
        return value if 10 <= len(digits) <= 15 else ""
    if field == "full_name":
        if len(value) > 80 or not (1 <= len(value.split()) <= 6):
            return ""
        if re.search(r'[@\d]', value):
            return ""
        return value[:80]
    if field == "linkedin":
        m = re.search(r'linkedin\.com/in/([a-zA-Z0-9\-_%]+)', value, re.I)
        if m:
            return "https://www.linkedin.com/in/" + m.group(1)
        handle = re.sub(r'^https?://(www\.)?linkedin\.com/', '', value, flags=re.I).strip('/')
        handle = handle[3:] if handle.lower().startswith("in/") else handle
        if handle and re.match(r'^[a-zA-Z0-9][a-zA-Z0-9\-_%]{2,}$', handle):
            return "https://www.linkedin.com/in/" + handle
        return ""
    if field == "location":
        if len(value) > 100 or re.search(r'@|https?://', value):
            return ""
        return value[:180]
    return value


def _identity_value_grounded(field, value, source_text):
    """Reject values the model invented instead of read — require the value to
    actually trace back to the resume text. Small local models will confidently
    return a plausible-looking placeholder (e.g. "johndoe@example.com") when the
    real contact info isn't in the text they were given (e.g. it's trapped in an
    image-based header that OCR hasn't recovered yet); format validity alone
    can't tell a real value from a fabricated one, only presence in-text can.
    """
    text_lower = source_text.lower()
    if field == "email":
        local = value.split('@', 1)[0].lower()
        return local in text_lower
    if field == "phone":
        digits = re.sub(r"\D", "", value)
        text_digits = re.sub(r"\D", "", source_text)
        return len(digits) >= 7 and digits[-7:] in text_digits
    if field == "linkedin":
        handle = value.rstrip('/').rsplit('/', 1)[-1].lower()
        return handle in text_lower
    # full_name / location: require at least one significant word to appear —
    # exact-phrase grounding is too strict since the model may reformat casing/
    # word order, but a value with zero words present in source is fabricated.
    words = [w.lower() for w in re.findall(r"[A-Za-z]{3,}", value)]
    return any(w in text_lower for w in words) if words else False


def _ollama_identity_extract(text, missing_fields):
    """Ask the local LLM for identity fields the rule-based parser couldn't find.

    Runs on its own — independent of the content-field AI budget below — so heavy
    fields (summary/experience, which can legitimately take most of that budget)
    never starve name/contact extraction. This is the generalizing fallback for
    resume layouts that don't match any of the rule-based heuristics, and the ONLY
    fallback at all for location/linkedin, which the regex-only pass can miss
    whenever they aren't behind an explicit label.
    """
    if not missing_fields:
        return {}
    prompt = (
        "Extract the candidate's identity details from this resume header/contact "
        "area. Return ONLY this JSON, using an empty string for anything not "
        'present:\n{"full_name": "", "email": "", "phone": "", "location": "", '
        '"linkedin": ""}\n\nRESUME:\n' + text[:1800]
    )
    try:
        raw = _ollama_chat(prompt, as_json=True, num_predict=150, num_ctx=2048)
    except Exception as e:
        logger.warning(f"Identity AI extraction failed: {e}")
        return {}
    if not isinstance(raw, dict):
        return {}
    filled = {}
    for field in missing_fields:
        val = _validate_identity_field(field, raw.get(field, ""))
        if val and not _identity_value_grounded(field, val, text):
            logger.warning(f"Identity AI '{field}' = '{val}' not grounded in source text — discarding as hallucination")
            val = ""
        if val:
            filled[field] = val
    return filled


def _split_resume_chunks(text):
    """Cut clean resume text into {field: section_text} pieces at heading lines.

    Handles 2-column PDFs by scanning for all section headings first, then extracting
    content between them by position rather than linear order. This prevents columns
    from getting scrambled (e.g. "KEY SKILLS" appearing after "WORK EXPERIENCE").
    """
    lines = text.split("\n")

    # First pass: identify all section headings and their line numbers
    headings = []  # [(line_num, field_name), ...]
    for i, line in enumerate(lines):
        s = line.strip()
        if not s:
            continue
        key = s.lower().strip(":·•- ")
        field = canonical_section_name(s) or _HEADING_FIELD.get(key)
        # Treat as a heading only if it's short (real headings are 1-3 words).
        # Count only words that contain at least one letter so decorative tokens
        # like "────", "===", "•" etc. do not inflate the word count and cause
        # headings such as "── PROFILE SUMMARY ──" to be silently skipped.
        _alpha_word_count = sum(1 for w in s.split() if any(c.isalpha() for c in w))
        if field and _alpha_word_count <= 3:
            headings.append((i, field))

    # Second pass: extract content between heading positions
    chunks = {}
    for idx, (line_num, field) in enumerate(headings):
        # Start from the line after the heading
        start = line_num + 1
        # End at the next heading (or end of text)
        end = headings[idx + 1][0] if idx + 1 < len(headings) else len(lines)

        # Collect non-empty lines in this range
        content_lines = []
        for line in lines[start:end]:
            s = line.strip()
            if s:
                content_lines.append(s)

        if content_lines:
            content = "\n".join(content_lines).strip()
            # Handle duplicate section headings by appending content
            if field in chunks:
                chunks[field] = chunks[field] + "\n" + content
            else:
                chunks[field] = content

    return chunks


def _pymupdf_rawdict_scan(path):
    """Extract email/phone from page 1 using pymupdf's character-level rawdict mode.

    `get_text("rawdict")` traverses the PDF content stream character by character,
    including text in coloured/styled areas that `get_text("words")` sometimes misses
    when glyph-to-Unicode mappings are non-standard (e.g. icon fonts in contact bars).

    Returns a dict with 'email' and/or 'phone' keys — only the found ones.
    """
    if not HAS_PYMUPDF:
        return {}
    try:
        import pymupdf as fitz
        doc = fitz.open(str(path))
        page = doc[0]
        raw = page.get_text("rawdict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        doc.close()

        line_texts = []
        for block in raw.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                line_str = ""
                for span in line.get("spans", []):
                    for char_obj in span.get("chars", []):
                        ch = char_obj.get("c", "")
                        if ch:
                            line_str += ch
                if line_str.strip():
                    line_texts.append(line_str)

        full_text = "\n".join(line_texts)
        logger.debug(f"rawdict scan page-1 first 400 chars: {full_text[:400]!r}")

        found = {}
        em = _extract_email(full_text)
        if em:
            found["email"] = em
        ph = _extract_phone(full_text)
        if ph:
            found["phone"] = ph
        if found:
            logger.info(f"rawdict scan found: {found}")
        return found
    except Exception as exc:
        logger.debug(f"_pymupdf_rawdict_scan: {exc}")
        return {}


def _tesseract_contact_scan(path):
    """OCR the top portion of page 1 with pytesseract to recover email/phone.

    Designed for styled PDFs where the contact bar (dark background, white text,
    icon fonts) is NOT accessible through any PDF text-layer extraction method.
    Renders page 1 at 200 DPI, crops to the top 35 %, and runs tesseract OCR.

    Returns a dict with 'email' and/or 'phone' keys — only the found ones.
    """
    try:
        import pytesseract
        from PIL import Image
        import pymupdf as fitz
        import numpy as np

        for _tp in [
            r"C:\Users\sindhu.sundara\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        ]:
            if os.path.isfile(_tp):
                pytesseract.pytesseract.tesseract_cmd = _tp
                break

        doc = fitz.open(str(path))
        pg = doc[0]
        clip = fitz.Rect(0, 0, pg.rect.width, pg.rect.height * 0.35)
        mat  = fitz.Matrix(200 / 72, 200 / 72)   # 200 DPI — good quality for OCR
        pix  = pg.get_pixmap(matrix=mat, clip=clip, colorspace=fitz.csRGB)
        doc.close()

        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, 3
        )
        img = Image.fromarray(img_array)

        # --psm 6: assume a uniform block of text (contact bar)
        # --oem 3: LSTM engine
        ocr_text = pytesseract.image_to_string(
            img, config="--psm 6 --oem 3 -c tessedit_char_whitelist="
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@.+\\-_ "
        )
        logger.info(f"Tesseract OCR top-35%: {ocr_text[:300]!r}")

        found = {}
        em = _extract_email(ocr_text)
        if em:
            found["email"] = em
        ph = _extract_phone(ocr_text)
        if ph:
            found["phone"] = ph
        if found:
            logger.info(f"Tesseract contact scan found: {found}")
        return found
    except Exception as exc:
        logger.debug(f"_tesseract_contact_scan: {exc}")
        return {}


def _tesseract_inverted_ocr_text(path):
    """Render page 1, invert the image, and return the raw Tesseract OCR string.

    WHY INVERSION:
    Dark-background contact bars (navy/black with white text) give Tesseract very
    poor results because it expects dark ink on a light page.  Inverting the image
    first (dark background → light, white text → dark) gives full-contrast black
    text on white, which Tesseract reads with high accuracy.

    The RED channel is used for the inversion pass: a navy bar has near-zero red
    values, so inverting the red channel produces a near-white background with
    dark text — maximum contrast for the OCR engine.

    Returns the raw OCR string (may be empty on error).
    """
    try:
        import pytesseract
        from PIL import Image, ImageOps
        import pymupdf as fitz
        import numpy as np

        for _tp in [
            r"C:\Users\sindhu.sundara\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        ]:
            if os.path.isfile(_tp):
                pytesseract.pytesseract.tesseract_cmd = _tp
                break

        doc = fitz.open(str(path))
        pg  = doc[0]
        # Render top 40 % at 300 DPI — higher res improves small-font accuracy
        clip = fitz.Rect(0, 0, pg.rect.width, pg.rect.height * 0.40)
        mat  = fitz.Matrix(300 / 72, 300 / 72)
        pix  = pg.get_pixmap(matrix=mat, clip=clip, colorspace=fitz.csRGB)
        doc.close()

        img_arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, 3
        )
        img = Image.fromarray(img_arr)

        # Pass 1 — invert full colour image → dark-on-light
        img_inv = ImageOps.invert(img)

        # Pass 2 — use red channel only: navy bar → near-black → inverted = near-white
        r_ch = img.split()[0]               # red channel
        r_inv = ImageOps.invert(r_ch)       # invert: 0→255 (white bg), 255→0 (dark text)

        ocr1 = pytesseract.image_to_string(img_inv.convert("L"), config="--psm 6 --oem 3")
        ocr2 = pytesseract.image_to_string(r_inv,                 config="--psm 6 --oem 3")
        combined = ocr1 + "\n" + ocr2
        logger.info(f"Inverted OCR (pass1+pass2 first 400): {combined[:400]!r}")
        return combined
    except Exception as exc:
        logger.debug(f"_tesseract_inverted_ocr_text: {exc}")
        return ""


def _ollama_vision_contact_scan(path):
    """Send the rendered contact bar image to an Ollama vision model to read email/phone.

    The dark-background contact bars in styled resumes are NOT in the PDF text layer
    and cannot be read by any text-extraction method.  This function renders the top
    30 % of page 1 as a PNG, encodes it as base64, and sends it to the first vision-
    capable model found in Ollama (moondream, llava, bakllava, minicpm-v, etc.).

    To enable: run  `ollama pull moondream`  in a terminal (≈1.7 GB one-time download).

    Returns a dict with 'email' and/or 'phone' keys — only the validated ones.
    """
    try:
        import base64
        import io
        import json as _json
        import requests as _req
        import pymupdf as fitz
        import numpy as np
        from PIL import Image

        # Find a vision-capable model in Ollama
        _tags = _req.get(f"{_OLLAMA_BASE}/api/tags", timeout=5).json()
        _vision_kw = ("moondream", "llava", "bakllava", "minicpm-v", "vision",
                      "qwen2-vl", "gemma3", "pixtral")
        _vision_models = [
            m["name"] for m in _tags.get("models", [])
            if any(kw in m["name"].lower() for kw in _vision_kw)
        ]
        if not _vision_models:
            logger.info(
                "No vision model found in Ollama — skipping vision contact scan. "
                "Run: ollama pull moondream"
            )
            return {}

        vision_model = _vision_models[0]
        logger.info(f"Vision model selected: {vision_model}")

        # Render top 30 % of page 1 at 200 DPI
        doc  = fitz.open(str(path))
        pg   = doc[0]
        clip = fitz.Rect(0, 0, pg.rect.width, pg.rect.height * 0.30)
        mat  = fitz.Matrix(200 / 72, 200 / 72)
        pix  = pg.get_pixmap(matrix=mat, clip=clip, colorspace=fitz.csRGB)
        doc.close()

        img_arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, 3
        )
        img = Image.fromarray(img_arr)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        img_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

        prompt = (
            "This is the top of a resume. There is a dark-colored contact bar with "
            "an email address and a phone number written in white text.\n"
            "Extract ONLY the email address and phone number.\n"
            "Return ONLY valid JSON with no extra text:\n"
            '{"email": "someone@domain.com", "phone": "+91 XXXX XXX XXX"}'
        )

        payload = {
            "model": vision_model,
            "messages": [{"role": "user", "content": prompt, "images": [img_b64]}],
            "stream": False,
            "options": {"temperature": 0, "num_predict": 80},
        }
        resp = _req.post(
            f"{_OLLAMA_BASE}/api/chat", json=payload,
            timeout=_OLLAMA_TEXT_TIMEOUT
        )
        resp.raise_for_status()
        response_text = resp.json().get("message", {}).get("content", "")
        logger.info(f"Vision model raw response: {response_text[:300]!r}")

        # Extract JSON from the response
        json_m = re.search(r'\{[^}]+\}', response_text, re.DOTALL)
        if not json_m:
            return {}

        data = _json.loads(json_m.group())
        found = {}
        _em = str(data.get("email", "")).strip().rstrip(".,;)")
        _ph = str(data.get("phone", "")).strip()

        _em_re = re.compile(r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$')
        if _em and _em_re.match(_em):
            found["email"] = _em

        _ph_digits = re.sub(r"\D", "", _ph)
        if 10 <= len(_ph_digits) <= 15:
            found["phone"] = _ph

        if found:
            logger.info(f"Vision model contact found: {found}")
        return found

    except Exception as exc:
        logger.debug(f"_ollama_vision_contact_scan: {exc}")
        return {}


def _pdfplumber_whitetext_scan(path):
    """Extract email/phone by filtering pdfplumber chars to white/light-colored text.

    Dark-background contact bars (navy, black) render text in white.  pdfplumber's
    `page.chars` exposes the `non_stroking_color` (fill color) for every character
    in the PDF content stream.  Filtering to light colors isolates the contact bar
    text from the rest of the page content, then email/phone patterns are applied.

    Returns a dict with 'email' and/or 'phone' keys — only the found ones.
    """
    if not HAS_PDFPLUMBER:
        return {}
    try:
        import pdfplumber
        from collections import defaultdict
        with pdfplumber.open(str(path)) as pdf:
            if not pdf.pages:
                return {}
            chars = pdf.pages[0].chars

        if not chars:
            return {}

        def _is_light_color(color):
            """Return True for white / near-white fill colors."""
            if color is None:
                return False
            if isinstance(color, (int, float)):
                return float(color) >= 0.70
            if isinstance(color, (list, tuple)):
                vals = [float(v) for v in color]
                # RGB or CMYK (inverted for K channel)
                if len(vals) == 3:
                    return all(v >= 0.65 for v in vals)
                if len(vals) == 1:
                    return vals[0] >= 0.70
            return False

        white_chars = [c for c in chars if _is_light_color(c.get("non_stroking_color"))]
        logger.debug(f"whitetext scan: {len(white_chars)} light-colored chars out of {len(chars)}")
        if not white_chars:
            return {}

        # Group characters into lines by y-position (tolerance 4 pts)
        lines_dict = defaultdict(list)
        for c in white_chars:
            y_key = round(float(c.get("top", 0)) / 4) * 4
            lines_dict[y_key].append(c)

        line_texts = []
        for y in sorted(lines_dict.keys()):
            line_chars = sorted(lines_dict[y], key=lambda c: float(c.get("x0", 0)))
            text = "".join(c.get("text", "") for c in line_chars)
            if text.strip():
                line_texts.append(text)

        full_text = "\n".join(line_texts)
        logger.debug(f"whitetext scan text: {full_text[:300]!r}")

        found = {}
        em = _extract_email(full_text)
        if em:
            found["email"] = em
        ph = _extract_phone(full_text)
        if ph:
            found["phone"] = ph
        if found:
            logger.info(f"whitetext scan found: {found}")
        return found
    except Exception as exc:
        logger.debug(f"_pdfplumber_whitetext_scan: {exc}")
        return {}


def _pdfplumber_contact_scan(path):
    """Extract email/phone from page 1 of a styled PDF (icon-font contact bars).

    pdfplumber's extract_words() retrieves all glyph boxes including those on
    coloured backgrounds that pymupdf sometimes misses when fonts are embedded
    with a non-standard encoding.  Used as a second-chance pass before the AI
    fallback fires for email/phone.

    Returns a dict with 'email' and/or 'phone' keys — only the found ones.
    """
    if not HAS_PDFPLUMBER:
        return {}
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            if not pdf.pages:
                return {}
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=5, y_tolerance=5)
            plain = page.extract_text() or ""
            word_line = " ".join(w.get("text", "") for w in words if w.get("text", "").strip())
            combined = word_line + "\n" + plain
        found = {}
        em = _extract_email(combined)
        if em:
            found["email"] = em
        ph = _extract_phone(combined)
        if ph:
            found["phone"] = ph
        if found:
            logger.info(f"pdfplumber contact scan found: {list(found.keys())}")
        return found
    except Exception as exc:
        logger.debug(f"_pdfplumber_contact_scan: {exc}")
        return {}


_TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9\-]{3,}")


def _looks_like_proper_noun_token(tok):
    """True for tokens shaped like a proper noun / product name / acronym
    ("eLHR", "LSPath", "POMSnet", "ADC", "M-Files") rather than an ordinary
    English word ("develop", "manufacturing", "Omni"). With only a handful of
    project blocks to compare, plain words can end up statistically unique to
    one block by pure chance — this shape filter keeps that coincidence from
    being mistaken for a real, distinctive project identifier.
    """
    if any(c.isdigit() for c in tok):
        return True
    if tok.isupper() and 2 <= len(tok) <= 8:
        return True
    return any(c.isupper() for c in tok[1:])


def _reassign_project_specific_bullets(experience_text, projects_text):
    """Move Experience bullets that are clearly about ONE specific listed project
    into that project's block in Projects, instead of leaving them as generic
    Experience content.

    Purely additive / conservative by design: it only ever moves a bullet when
    it contains a proper-noun-shaped keyword (see `_looks_like_proper_noun_token`)
    that is unique to exactly one of the project blocks on THIS resume (computed
    per-upload, no hardcoded project names or domain stoplist) — product names
    (e.g. "eLHR", "LSPath") naturally pass this test, while generic domain
    vocabulary shared across every project ("validation", "CSV", "review") and
    plain capitalized words ("Omni") naturally fail it and are left alone.
    A bullet whose keywords point to more than one project, or to none, is left
    exactly where it was. Only runs when there are 2+ project blocks to disambiguate
    between — on every other resume shape this is a complete no-op.
    """
    if not experience_text or not projects_text:
        return experience_text, projects_text

    blocks = re.split(r'(?=^\s*Client:)', projects_text, flags=re.MULTILINE)
    blocks = [b for b in blocks if b.strip()]
    if len(blocks) < 2:
        return experience_text, projects_text  # nothing to disambiguate against

    # Uniqueness must be computed over ALL case variants of a word, not just the
    # shape-qualified ones — otherwise a word that's overwhelmingly common but
    # happens to appear in ALL CAPS once (e.g. a stray "ABBOTT" in one project's
    # description, vs. "Abbott" as the client name in five others) would look
    # falsely "unique" just because its other occurrences got shape-filtered out
    # before the counting happened.
    block_raw_tokens = [_TOKEN_RE.findall(b) for b in blocks]
    token_owner_count = {}
    for toks in block_raw_tokens:
        for t in {tok.lower() for tok in toks}:
            token_owner_count[t] = token_owner_count.get(t, 0) + 1
    distinctive = [
        {t.lower() for t in toks
         if token_owner_count[t.lower()] == 1 and _looks_like_proper_noun_token(t)}
        for toks in block_raw_tokens
    ]
    if not any(distinctive):
        return experience_text, projects_text

    # Split Experience into bullet segments (each starting at a "•" marker);
    # non-bullet lines (e.g. the company/title/dates table) are left untouched.
    lines = experience_text.split('\n')
    segments = []
    current = None
    for line in lines:
        if line.strip().startswith('•'):
            if current is not None:
                segments.append(current)
            current = [line]
        elif current is not None:
            current.append(line)
        else:
            segments.append([line])
    if current is not None:
        segments.append(current)

    # When a resume has its own distinct Projects section (2+ blocks, confirmed
    # above) AND Experience contains bullet points at all, those bullets did not
    # come from the Work Experience heading itself — that heading's own content is
    # the non-bullet company/title/dates table; every "•" line reaching this point
    # was appended from a "Role and Responsibilities"/"Project Details" occurrence
    # (see the heading loop above). So here ALL bullets are routed into Projects:
    # to the one specific project they name if identifiable, otherwise to a
    # general trailing section — never left mixed into the Experience table.
    kept_experience = []
    appended_to_block = [[] for _ in blocks]
    unmatched_bullets = []
    for seg in segments:
        seg_text = '\n'.join(seg)
        if not seg_text.strip().startswith('•'):
            kept_experience.append(seg_text)
            continue
        seg_tokens = {t.lower() for t in _TOKEN_RE.findall(seg_text)}
        matches = [i for i, dist in enumerate(distinctive) if seg_tokens & dist]
        if len(matches) == 1:
            appended_to_block[matches[0]].append(seg_text)
            logger.info(
                f"Moved Experience bullet to Project block {matches[0]} "
                f"(matched keyword(s): {seg_tokens & distinctive[matches[0]]})"
            )
        else:
            unmatched_bullets.append(seg_text)
            logger.info("Moved Experience bullet to general Projects section (no single distinctive project match)")

    if not any(appended_to_block) and not unmatched_bullets:
        return experience_text, projects_text  # no bullets to move — no-op

    new_blocks = []
    for block, extra in zip(blocks, appended_to_block):
        if extra:
            new_blocks.append(block.rstrip() + "\nResponsibilities:\n" + "\n".join(extra))
        else:
            new_blocks.append(block.rstrip())
    if unmatched_bullets:
        new_blocks.append("Role and Responsibilities:\n" + "\n".join(unmatched_bullets))
    return "\n".join(kept_experience).strip(), "\n".join(new_blocks).strip()


def parse_resume_with_llm_text(path):
    """Pure-Ollama resume parsing — no regex field engine.

    1. PDF → clean text (column-split aware).
    2. One Ollama call extracts the identity header into JSON (name/title/contact) —
       this is where a model genuinely beats rules, and the short header keeps it fast
       and safe from hallucination.
    3. Content sections are cut from the text at their heading lines and used verbatim
       — faithful to the resume (no invented content) and instant.
    Returns (fields_dict, "llm_text").
    """
    _path_ext = str(path).rsplit(".", 1)[-1].lower() if "." in str(path) else "pdf"
    raw_text = extract_resume_text(path, _path_ext)

    # DEBUG: write raw text to inspect extraction issues
    try:
        import os as _os
        _dbg = _os.path.join(_os.path.dirname(str(path)), "debug_raw_text.txt")
        with open(_dbg, "w", encoding="utf-8") as _f:
            _f.write(raw_text[:5000])
    except Exception:
        pass

    # For sidebar-style 2-column PDFs the pymupdf word-order extractor interleaves
    # left-column content (KEY SKILLS) with right-column content (WORK EXPERIENCE)
    # when the gutter is too close to the page edge for the block-based detector.
    # Try the pdfplumber crop-based extractor which reads each column independently.
    if str(path).lower().endswith(".pdf"):
        # Preserve any OCR-injected contact info BEFORE possibly replacing raw_text.
        # Icon-based contact bars (image text) are only visible to OCR; pdfplumber
        # reads the PDF vector layer and would miss them.
        _ocr_email = _extract_email(raw_text)
        _ocr_phone = _extract_phone(raw_text)
        _m_ocr_name = re.search(r'(?m)^\s*[Nn]ame\s*:\s*(.+)$', raw_text)
        _ocr_name_line = _m_ocr_name.group(1).strip() if _m_ocr_name else ""
        _m_ocr_loc = re.search(r'(?m)^\s*[Ll]ocation\s*:\s*(.+)$', raw_text)
        _ocr_loc_line = _m_ocr_loc.group(1).strip() if _m_ocr_loc else ""
        two_col = _pdfplumber_two_col_text(path)
        if two_col and len(two_col.strip()) >= len(raw_text.strip()) * 0.5:
            _tc = two_col
            def _cid_to_char_tc(m):
                n = int(m.group(1))
                return chr(n) if 33 <= n <= 126 else ''
            _tc = re.sub(r'\(cid:(\d+)\)', _cid_to_char_tc, _tc)
            # Reject two_col if it has significantly more word-fragment lines than
            # the pymupdf raw_text — this detects false column splits where the
            # pdfplumber gutter crop cuts through full-width single-column lines,
            # leaving partial words (e.g. "fund pa", "fee reco") at line ends.
            _FRAG_EXCL = {'and','the','of','in','to','for','at','by','or','on','an','be','as','it','its','has','was','are','may'}
            def _frag_lines(t):
                c = 0
                for _l in t.splitlines():
                    _ls = _l.strip()
                    if len(_ls) >= 15:
                        _m = re.search(r'\b([a-z]{1,4})$', _ls)
                        if _m and _m.group(1) not in _FRAG_EXCL:
                            c += 1
                return c
            # A WHOLE line that's just 1-2 letters (with optional trailing period,
            # e.g. "s." or "on" sitting alone) is a stronger corruption signal than
            # a trailing fragment — it means the column crop split a word/phrase
            # apart entirely, not just at a line's tail. `_frag_lines` alone misses
            # this since it only inspects lines >= 15 chars. Seen in practice: a
            # skills-sidebar column crop leaving "s.\non Selenium WebDriver" where
            # the source read "Skilled in Selenium WebDriver" as one line.
            _ORPHAN_FRAG_RE = re.compile(r'^[a-z]{1,2}\.?$')
            def _has_orphan_fragment_lines(t):
                return any(_ORPHAN_FRAG_RE.match(_l.strip()) for _l in t.splitlines())
            # A column crop can also misplace a fragment from the MIDDLE of a
            # sentence to the very END of the document — e.g. "...suspicious user
            # activities." loses its own tail into a stray "ser activities." line
            # appended after the resume's true last content (its Education entry).
            # A genuine re-extraction reorders content; it never invents new
            # trailing material past where the document actually ends. So when the
            # original raw_text's own true ending IS found inside two_col, anything
            # non-trivial appended after that point is fabricated, not reordered.
            _orig_tail = raw_text.strip()[-60:]
            _tail_ok = True
            if _orig_tail:
                _tail_pos = _tc.find(_orig_tail)
                if _tail_pos != -1:
                    _extra_after_tail = _tc[_tail_pos + len(_orig_tail):].strip()
                    _tail_ok = len(_extra_after_tail) <= 20
            if (_frag_lines(_tc) <= _frag_lines(raw_text) + 2
                    and not _has_orphan_fragment_lines(_tc) and _tail_ok):
                raw_text = _tc
            else:
                logger.info("two_col rejected: pdfplumber false column split detected (word fragments)")
            # Re-inject all OCR contact info if not already present in the new text.
            # Name/Location use a "labelled line already exists" check rather than a
            # raw substring search — a plain "in raw_text" check falsely matches when
            # the same words appear incidentally elsewhere (e.g. a college's city
            # matching the candidate's home city, like "...college, India, Chennai
            # 2015" in the Education section), which silently drops the re-injection.
            _post = []
            if _ocr_name_line and not re.search(r'(?m)^\s*name\s*:', raw_text, re.I):
                _post.append(f"Name: {_ocr_name_line}")
            if _ocr_loc_line and not re.search(r'(?m)^\s*(location|address|city)\s*:', raw_text, re.I):
                _post.append(f"Location: {_ocr_loc_line}")
            if _ocr_email and _ocr_email not in raw_text:
                _post.append(f"Email: {_ocr_email}")
            if _ocr_phone and _ocr_phone not in raw_text:
                _post.append(f"Phone: {_ocr_phone}")
            if _post:
                raw_text = "\n".join(_post) + "\n" + raw_text

    # Strip standalone page numbers (e.g. "\n1\n", "\n 2 \n", "\n3" at end)
    raw_text = re.sub(r'\n[ \t]*\d{1,3}[ \t]*(?=\n|$)', '', raw_text)

    # ── 1. Base parse via fast regex (instant) ───────────────────────────────────
    # Skip font-metadata name extraction when OCR already injected a "Name:" label —
    # the OCR result is more reliable than scanning body text for large-font words.
    _ocr_injected_name = bool(re.search(r'(?m)^\s*name\s*:', raw_text, re.I))
    font_name = (
        _extract_name_from_pdf_fonts(path)
        if _path_ext == "pdf" and not _ocr_injected_name
        else None
    )
    result = parse_resume_text(raw_text, name_hint=font_name)

    # Clear name/title if quick_parse grabbed a degree string instead of real values
    _degree_abbr_re = re.compile(
        r'^(M\.?\s?Sc\.?|B\.?\s?Sc\.?|M\.?\s?B\.?\s?A\.?|Ph\.?\s?D\.?'
        r'|M\.?\s?Tech\.?|B\.?\s?Tech\.?|B\.?\s?E\.?|D\.?\s?Pharm\.?|B\.?\s?Pharm\.?)$',
        re.IGNORECASE
    )
    _degree_phrase_re = re.compile(
        r'\b(master|bachelor|m\.?sc|b\.?sc|m\.?tech|b\.?tech|diploma|'
        r'certificate|analytical chemistry|pharmacy|life science)\b',
        re.IGNORECASE
    )
    _name = result.get("full_name", "").strip()
    if not _name or _degree_abbr_re.match(_name):
        logger.info(f"Name '{_name}' looks like a degree — clearing for AI fallback")
        result["full_name"] = ""
    _title = result.get("title", "").strip()
    if _degree_phrase_re.search(_title):
        logger.info(f"Title '{_title}' looks like a degree — clearing for AI fallback")
        result["title"] = ""

    # Detect name/title swap: if the "name" field contains a role keyword and the
    # "title" field looks like a person's name (2+ words, no role keyword), swap them.
    # Happens when font-size extractor picks the job-title line as the largest text.
    _role_kw_re = re.compile(
        r'\b(analyst|engineer|lead|manager|consultant|developer|specialist|'
        r'executive|officer|director|architect|scientist|associate|coordinator|'
        r'validation|quality|compliance|pharmacist|chemist|technician|programmer)\b',
        re.IGNORECASE
    )
    _cur_name  = result.get("full_name", "").strip()
    _cur_title = result.get("title", "").strip()
    if (_cur_name and _cur_title
            and _role_kw_re.search(_cur_name)
            and not _role_kw_re.search(_cur_title)
            and len(_cur_title.split()) >= 2):
        logger.info(f"Name/title swap detected — swapping '{_cur_name}' ↔ '{_cur_title}'")
        result["full_name"], result["title"] = _cur_title, _cur_name

    # ── 2. Verbatim section extraction ─────────────────────────────────────────────
    # Exact keywords only. CRITICAL: filter out word-wrap false matches by requiring
    # a blank line (double \n) before any heading — "Software \nQualifications" is a
    # wrapped sentence, not a section heading; real headings always follow a blank line.
    _EXACT_HEADINGS = [
        "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE", "EMPLOYMENT HISTORY",
        "PROFESSIONAL SUMMARY", "PROFILE SUMMARY", "APPLICATIONS SUMMARY",
        "EXECUTIVE SUMMARY", "CAREER OBJECTIVE", "PROFESSIONAL PROFILE", "PROFILE",
        "AREAS OF EXPERTISE", "TECHNICAL SKILLS", "CORE COMPETENCIES", "KEY SKILLS",
        "ACADEMIC BACKGROUND", "ACADEMIC QUALIFICATIONS",
        "EDUCATIONAL QUALIFICATION", "EDUCATIONAL QUALIFICATIONS",
        "SUMMARY", "EXPERIENCE", "EMPLOYMENT", "SKILLS", "EDUCATION",
        "QUALIFICATIONS", "CERTIFICATIONS", "PROJECTS", "ACHIEVEMENTS",
        "AWARDS", "REFERENCES",
        "TOOLS & APPLICATIONS", "TOOLS AND APPLICATIONS", "TOOLS",
        "EARLY CAREER EXPERIENCE",
        "EDUCATION & CERTIFICATIONS", "EDUCATION AND CERTIFICATIONS",
        "TECHNICAL SKILLS & TOOLS", "TECHNICAL SKILLS AND TOOLS",
        "PROJECT DETAILS", "KEY PROJECTS", "KEY PROJECTS WORKED",
        "PROJECT HIGHLIGHTS", "PROJECT SUMMARY",
        "Relevant Project/Organizational Details",
        "Project/Organizational Details",
        "PROFESSIONAL CERTIFICATIONS", "PROFESSIONAL CERTIFICATION",
        "AWARDS & RECOGNITION", "AWARDS AND RECOGNITION",
        "LICENSES AND CERTIFICATIONS", "LICENSES & CERTIFICATIONS",
        "CERTIFICATES AND LICENSES", "CERTIFICATES & LICENSES",
        "CERTIFICATION AND TRAINING", "CERTIFICATIONS AND TRAINING",
        "CERTIFICATIONS & TRAINING", "CERTIFICATION & TRAINING",
        "CERTIFICATIONS & PROFESSIONAL DEVELOPMENT",
        "CERTIFICATIONS AND PROFESSIONAL DEVELOPMENT",
        "PROFESSIONAL DEVELOPMENT", "TRAINING",
        "ACADEMIC DETAILS",
        "KEY ACHIEVEMENTS", "KEY ACCOMPLISHMENTS",
        "CAREER SUMMARY", "DECLARATION", "PERSONAL DETAILS", "OTHER PERSONAL DETAILS",
        "HOBBIES", "LANGUAGES", "OBJECTIVE",
        "Professional Summary & Skill set", "Technical & Other Proficiency",
        # Title-case headings (Arun-style resumes)
        "Professional Summary", "Work Experience", "Relevant Project Experience",
        "Key Skills", "Technical Skills", "CORE SKILLS", "Core Skills",
        "Tools & Applications", "Tools and Applications",
        # Per-role bullet lists — common in Indian IT/validation CVs that repeat a
        # "Project Details" + "Role and Responsibilities" pair once per employer.
        # Without this heading recognized, its bullets silently fall inside whatever
        # OTHER section happens to precede them in the linear text (usually Projects).
        "ROLES AND RESPONSIBILITIES", "Roles and Responsibilities",
        "ROLE AND RESPONSIBILITIES", "Role and Responsibilities",
        "KEY RESPONSIBILITIES", "Key Responsibilities",
        "RESPONSIBILITIES", "Responsibilities",
        # ── Additions to close gaps found on 2-column/sidebar resumes ──────────
        # "About Me" is a summary heading (was only in the base-parse table, so the
        # verbatim pass was blind to it and could overwrite a good summary).
        "ABOUT ME", "ABOUT", "About Me",
        # "Technical Expertise" is a skills heading (sidebar templates use it).
        "TECHNICAL EXPERTISE", "Technical Expertise", "AREA OF EXPERTISE",
        # Contact/personal sidebars must act as section BOUNDARIES (mapped to no
        # field below) so their phone/address/email don't bleed into the section
        # printed just before them (e.g. Skills).
        "CONTACT", "CONTACT DETAILS", "CONTACT INFORMATION", "CONTACT INFO",
        "PERSONAL INFORMATION", "PERSONAL INFO", "PERSONAL PROFILE", "INTERESTS",
        # "Projects Completed" / "Certificates Courses" style multi-word headings
        # (incl. the common "compelted" typo) — the trailing word previously made
        # the single-keyword match fail, so the section was unrecognised.
        "PROJECTS COMPLETED", "PROJECTS COMPELTED", "PROJECT COMPLETED",
        "COMPLETED PROJECTS", "PROJECTS DONE",
        "CERTIFICATES COURSES", "CERTIFICATE COURSES",
        "CERTIFICATES AND COURSES", "COURSES AND CERTIFICATES", "COURSES",
    ]
    _section_heading_re = re.compile(
        r'(?:^|(?<=\n))\s*('
        + '|'.join(re.escape(k) for k in _EXACT_HEADINGS)
        # Allow any trailing non-word, non-newline chars (e.g. " :-", " —", " :") after
        # the keyword so "Education :-" and "Certifications :-" are also matched.
        + r')[^\w\n]*(?=\n|$)',
        re.IGNORECASE
    )

    def _heading_field(h):
        u = h.strip().upper()
        if u in ("SUMMARY", "PROFILE SUMMARY", "PROFESSIONAL SUMMARY",
                 "CAREER SUMMARY", "EXECUTIVE SUMMARY",
                 "PROFESSIONAL SUMMARY & SKILL SET",
                 "ABOUT ME", "ABOUT", "OBJECTIVE", "CAREER OBJECTIVE",
                 "PROFESSIONAL PROFILE", "PROFILE"):
            return "summary"
        if u in ("EXPERIENCE", "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE",
                 "EMPLOYMENT HISTORY", "EMPLOYMENT",
                 "EARLY CAREER EXPERIENCE", "EARLY CAREER",
                 "ROLES AND RESPONSIBILITIES", "ROLE AND RESPONSIBILITIES",
                 "KEY RESPONSIBILITIES", "RESPONSIBILITIES"):
            return "experience"
        if u in ("SKILLS", "TECHNICAL SKILLS", "KEY SKILLS", "CORE SKILLS", "CORE COMPETENCIES",
                 "AREAS OF EXPERTISE", "AREA OF EXPERTISE",
                 "TECHNICAL SKILLS & TOOLS", "TECHNICAL SKILLS AND TOOLS",
                 "TECHNICAL & OTHER PROFICIENCY", "TECHNICAL EXPERTISE",
                 "APPLICATIONS SUMMARY"):
            # "Applications Summary" lists tools/applications (e.g. "LIMS", "JIRA",
            # "Veeva Vault") despite the word "Summary" in its name — it belongs
            # with Skills, matching the base parser's SECTION_ALIASES classification.
            # This table previously classified it as "summary" instead, so it was
            # captured TWICE: once as skills by the base parser, once appended to
            # summary here — the same content duplicated across two fields.
            return "skills"
        if u in ("EDUCATION", "QUALIFICATIONS", "ACADEMIC BACKGROUND",
                 "ACADEMIC QUALIFICATIONS", "EDUCATIONAL QUALIFICATIONS",
                 "EDUCATIONAL QUALIFICATION", "EDUCATIONAL BACKGROUND",
                 "EDUCATION & CERTIFICATIONS", "EDUCATION AND CERTIFICATIONS",
                 "ACADEMIC DETAILS", "EDUCATIONAL DETAILS"):
            return "education"
        if u in ("CERTIFICATIONS", "CERTIFICATION",
                 "PROFESSIONAL CERTIFICATIONS", "PROFESSIONAL CERTIFICATION",
                 "LICENSES AND CERTIFICATIONS", "LICENSES & CERTIFICATIONS",
                 "CERTIFICATES AND LICENSES", "CERTIFICATES & LICENSES",
                 "CERTIFICATION AND TRAINING", "CERTIFICATIONS AND TRAINING",
                 "CERTIFICATIONS & TRAINING", "CERTIFICATION & TRAINING",
                 "CERTIFICATIONS & PROFESSIONAL DEVELOPMENT",
                 "CERTIFICATIONS AND PROFESSIONAL DEVELOPMENT",
                 "PROFESSIONAL DEVELOPMENT", "TRAINING",
                 "CERTIFICATES COURSES", "CERTIFICATE COURSES",
                 "CERTIFICATES AND COURSES", "COURSES AND CERTIFICATES",
                 "COURSES", "CERTIFICATES", "CERTIFICATE"):
            return "certifications"
        if u in ("ACHIEVEMENTS", "ACHIEVEMENT", "AWARDS", "AWARD",
                 "AWARDS & RECOGNITION", "AWARDS AND RECOGNITION",
                 "KEY ACHIEVEMENTS", "KEY ACCOMPLISHMENTS",
                 "ACCOMPLISHMENTS", "ACCOMPLISHMENT"):
            return "achievements"
        if u in ("PROJECTS", "PROJECT", "RELEVANT PROJECT EXPERIENCE",
                 "PROJECT DETAILS", "KEY PROJECTS", "KEY PROJECTS WORKED",
                 "PROJECT HIGHLIGHTS", "PROJECT SUMMARY",
                 "RELEVANT PROJECT/ORGANIZATIONAL DETAILS",
                 "RELEVANT PROJECT ORGANIZATIONAL DETAILS",
                 "PROJECTS COMPLETED", "PROJECTS COMPELTED", "PROJECT COMPLETED",
                 "COMPLETED PROJECTS", "PROJECTS DONE"):
            return "projects"
        if u in ("REFERENCES",):
            return "references"
        return None

    # Validate headings: must be ALL CAPS (letter characters only).
    # This filters word-wrapped sentence continuations like "...Software\nQualifications"
    # (mixed case) while accepting real headings like "PROFILE SUMMARY", "EXPERIENCE".
    # Title-case headings are accepted as fallback if preceded by a blank line OR if the
    # preceding line ends with a sentence terminator (the regex already ensures the keyword
    # appears alone on its line, so a prior sentence-end is a reliable heading signal).
    def _preceded_by_blank_or_sentence_end(pos):
        if pos <= 1:
            return True
        line_end = pos - 1
        line_start = raw_text.rfind('\n', 0, line_end)
        prev_line = raw_text[line_start + 1:line_end] if line_start != -1 else raw_text[:line_end]
        stripped = prev_line.strip()
        if not stripped:
            return True  # blank line before heading
        # A sentence/list-item ending immediately before the heading
        return stripped[-1] in '.):–—'

    # Single-word headings that are so section-specific they're safe without
    # requiring blank-line / sentence-end context (they almost never appear
    # mid-sentence in resume text).
    _STRONG_HEADINGS = frozenset([
        'CERTIFICATIONS', 'CERTIFICATION', 'ACHIEVEMENTS', 'ACHIEVEMENT',
        'AWARDS', 'AWARD', 'REFERENCES', 'PROJECTS', 'QUALIFICATIONS',
        'EDUCATION', 'RESPONSIBILITIES', 'HOBBIES',
    ])

    # Normalise sidebar-label lines where heading is merged with content on the same line
    # (e.g. "EXPERIENCE VALIDATION ENGINEER, 11/2021" from sidebar-style PDFs).
    # Split heading onto its own line so _section_heading_re can detect it.
    _uc_single_headings = [
        h for h in _EXACT_HEADINGS
        if h == h.upper() and ' ' not in h and '&' not in h and '/' not in h
    ]
    if _uc_single_headings:
        _sidebar_norm_re = re.compile(
            r'^(' + '|'.join(re.escape(h) for h in _uc_single_headings) + r')[ \t]+(?=[A-Z\d])',
            re.MULTILINE
        )
        raw_text = _sidebar_norm_re.sub(r'\1\n', raw_text)

    candidates = [(m.start(), m.group(1).strip()) for m in _section_heading_re.finditer(raw_text)]
    headings = []
    for pos, hd in candidates:
        letters = re.sub(r'[^a-zA-Z]', '', hd)
        is_all_caps = bool(letters) and letters == letters.upper()
        # Multi-word known headings (e.g. "Key Skills", "Profile Summary") are accepted
        # without requiring all-caps or a preceding blank line — they are unambiguous
        # as section headings because the regex already anchors them to their own line.
        is_multiword_known = len(hd.split()) >= 2
        # Unambiguously section-specific single words are accepted without context
        is_strong = hd.upper() in _STRONG_HEADINGS
        if is_all_caps or is_multiword_known or is_strong or _preceded_by_blank_or_sentence_end(pos):
            headings.append((pos, hd))
        else:
            logger.debug(f"Skipped '{hd}' at {pos}: single-word mixed-case with no blank line before")

    # A summary-family heading written as "Professional Summary: <text>" — label
    # and content on the SAME line — never matches `_section_heading_re` above
    # (it requires the heading to be immediately followed by end-of-line). Left
    # unrecognized, it isn't a boundary at all, so whatever section precedes it
    # keeps reading straight through it and into the content that follows —
    # e.g. a resume whose body repeats more than once bleeds the next repeat's
    # entire summary paragraph into the previous repeat's last real section.
    # This only needs to mark WHERE such a label starts, not capture its own
    # content (that's already handled by the base regex parser/find_sections),
    # so it's tagged with a sentinel `_heading_field` never maps to a field —
    # boundary only, never overwrites `result["summary"]`.
    for m in re.finditer(
        r'(?:^|(?<=\n))\s*(?:Professional Summary|Profile Summary|Career Summary|'
        r'Executive Summary|Career Objective|Objective|About Me|About|Summary)'
        r'\s*:\s*\S',
        raw_text, re.IGNORECASE,
    ):
        headings.append((m.start(), "\x00BOUNDARY\x00"))
    headings.sort(key=lambda t: t[0])

    logger.info(f"Validated headings: {[h for _, h in headings[:12]]}")

    _qual_exp_date_re = re.compile(
        r'(?:'
        r'(?:JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|JUN(?:E)?|'
        r'JUL(?:Y)?|AUG(?:UST)?|SEP(?:TEMBER)?|OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)'
        r'\s+\d{4}\s*[-–]'          # MONTH YEAR – (e.g. JAN 2020 –)
        r'|\d{1,2}/\d{4}\s*[-–]'   # MM/YYYY – (e.g. 06/2018-)
        r')',
        re.I,
    )

    # An ALL-CAPS "job title / company" line followed by a date (e.g.
    # "PROGRAMMER ANALYST, 06/2018") — used alongside _qual_exp_date_re to confirm
    # a "Projects" section is really mislabeled work history, not genuine
    # standalone projects (which often have their own dates too, but no company
    # line like this).
    _proj_jobtitle_re = re.compile(r'(?:^|\n)[A-Z][A-Z\s/]+,\s*\d{2}/\d{4}')

    verbatim = {}
    _ambiguous_field_source = set()
    # Content sitting under headings that map to no resume field (e.g. a page-2
    # sidebar's "Other Personal Details"/"Hobbies"/"Extra Curricular"). On
    # 2-column PDFs the last Work-Experience bullets can be read AFTER one of
    # these, stranding them here where they'd otherwise be dropped. Collected so
    # the numbered-continuation rescue below can recover them.
    _unmapped_chunks = []
    # A two-column "Project Details | Roles and Responsibilities" TABLE (one
    # heading row shared across every project row, not a per-job sub-heading)
    # linearizes as two adjacent heading-like lines with nothing between them —
    # e.g. "...Relevant Project/Organizational Details\nProject/Organizational
    # Details\nRoles and Responsibilities\n<all 8 projects...>" (table header
    # split across two adjacent lines), or "...Project Details\n<5 projects'
    # descriptions>\nRole and Responsibilities\n<their role bullets>" (one
    # heading per table column, each opening its own span). Either way, "Roles
    # and Responsibilities" is always paired with a "Project Details"-family
    # heading in these validation-CV-style resumes — never a genuine standalone
    # Experience sub-heading in that context — so when it's immediately preceded
    # by one, route it to "projects" instead of the "experience" default that
    # applies when it appears as a per-job sub-heading elsewhere.
    _PROJECT_DETAILS_HEADINGS = {
        "PROJECT/ORGANIZATIONAL DETAILS", "RELEVANT PROJECT/ORGANIZATIONAL DETAILS",
        "RELEVANT PROJECT ORGANIZATIONAL DETAILS", "PROJECT DETAILS",
        "RELEVANT PROJECT EXPERIENCE",
    }
    _ROLES_RESP_HEADINGS = {
        "ROLES AND RESPONSIBILITIES", "ROLE AND RESPONSIBILITIES",
        "KEY RESPONSIBILITIES", "RESPONSIBILITIES",
    }

    for i, (pos, heading) in enumerate(headings):
        field = _heading_field(heading)
        nl_pos = raw_text.find("\n", pos)
        content_start = nl_pos + 1 if nl_pos != -1 else pos + len(heading)
        content_end = headings[i + 1][0] if i + 1 < len(headings) else len(raw_text)
        text = raw_text[content_start:content_end].strip()

        if heading.upper() in _ROLES_RESP_HEADINGS and i > 0:
            _prev_heading = headings[i - 1][1]
            if _prev_heading.upper() in _PROJECT_DETAILS_HEADINGS:
                field = "projects"
                logger.info(
                    "'%s' routed to 'projects' — immediately follows '%s' "
                    "(paired project-details/role-responsibilities table, not a "
                    "standalone Experience sub-heading)",
                    heading, _prev_heading,
                )

        # "Qualifications" headings are ambiguous — inspect content before routing.
        if heading.upper() == "QUALIFICATIONS" and text:
            if _qual_exp_date_re.search(text):
                # Has job date ranges → work experience entries mis-labelled by sidebar.
                if "experience" in verbatim:
                    verbatim["experience"] += "\n" + text
                else:
                    verbatim["experience"] = text
                logger.info(
                    "Verbatim 'experience' +%d chars re-routed from 'Qualifications' (date ranges found)",
                    len(text),
                )
                continue
            else:
                # No date ranges — check whether it's professional competency bullets
                # (common in Indian pharma/IT CVs where "Qualifications" lists skills and
                # achievements that are really a continuation of the professional summary).
                _q_edu_kw = re.search(
                    r'\b(university|college|institute|school|bachelor|master|b\.e\b|b\.tech\b|'
                    r'b\.com\b|m\.tech\b|m\.e\b|mba\b|phd\b|diploma|degree|cgpa|gpa|graduation)\b',
                    text, re.I,
                )
                _q_work_verbs = re.findall(
                    r'(?:^|\n)[ \t]*[^\w\s]?[ \t]*(?:Expertise|Managed|Performed|Prepared|Preparati|'
                    r'Worked|Working|Ensure|Contribut|Execut|Conduct|Complet|Monitor|Analys|'
                    r'Maintained|Develop|Review|Coordinat|Having|Exposure|Responsible|'
                    r'Handling|Support|Assist|Troubleshoot|Involved)',
                    text,
                )
                if not _q_edu_kw and len(_q_work_verbs) >= 2:
                    if "summary" in verbatim:
                        verbatim["summary"] += "\n" + text
                    else:
                        verbatim["summary"] = text
                    logger.info(
                        "Verbatim 'summary' +%d chars supplemented from 'Qualifications' "
                        "(professional competency bullets, no dates, no edu keywords)",
                        len(text),
                    )
                    continue
                # Otherwise: genuine academic qualifications → falls through to education

        # "Projects" / "Project Details" headings in Indian IT CVs often list work
        # assignments with company names and numeric date ranges (e.g. 06/2018-11/2021).
        # When detected, append to experience and skip setting projects. Requires
        # BOTH a date range AND an ALL-CAPS job-title/company line — a date range
        # alone isn't enough, since genuine standalone projects often have their
        # own start/end dates too and shouldn't be swept into Experience.
        if (field == "projects" and text
                and _qual_exp_date_re.search(text)
                and _proj_jobtitle_re.search(text)):
            if "experience" in verbatim:
                verbatim["experience"] += "\n" + text
            else:
                verbatim["experience"] = text
            # `result["projects"]` may already hold this same text from the earlier
            # base regex parse (parse_resume_text runs its own independent section
            # split) — clear it here so the content doesn't end up duplicated in
            # both Experience and Project Details.
            result["projects"] = ""
            logger.info(
                "Verbatim 'experience' +%d chars re-routed from '%s' (date ranges found in projects) "
                "— cleared duplicate 'projects' value from base parse",
                len(text), heading,
            )
            continue

        # "Career Summary" is ambiguous — most resumes use it for a professional
        # blurb, but some (confusingly) use it to head their job/career history.
        # Same signal as the "Qualifications" disambiguation above: job date
        # ranges mean it's really work experience, not a summary.
        if heading.upper() == "CAREER SUMMARY" and text and _qual_exp_date_re.search(text):
            if "experience" in verbatim:
                verbatim["experience"] += "\n" + text
            else:
                verbatim["experience"] = text
            logger.info(
                "Verbatim 'experience' +%d chars re-routed from 'Career Summary' (date ranges found)",
                len(text),
            )
            continue

        if not field:
            if text:
                _unmapped_chunks.append(text)
            continue
        if field in verbatim:
            # A field can legitimately recur multiple times in one resume — multiple
            # employers each with their own "Project Details"/"Role and
            # Responsibilities" block, multiple "Skills" groupings, etc. Whether a
            # later occurrence of the SAME field should REPLACE the current value or
            # be APPENDED to it depends on whether the current value was a confident
            # match or just a guess:
            if field in _ambiguous_field_source and heading.upper() != "QUALIFICATIONS":
                # The current value came from an ambiguous "Qualifications" guess;
                # a real, unambiguous heading for this field (e.g. "EDUCATION")
                # replaces that guess rather than appending to it.
                if text:
                    verbatim[field] = text
                    _ambiguous_field_source.discard(field)
                    logger.info(f"Verbatim '{field}' ({len(text)} chars) from '{heading}' (replaced ambiguous guess)")
            elif text:
                existing = verbatim[field]
                # Guard against a verbatim-duplicate block — e.g. a source
                # document whose entire body was pasted in more than once —
                # re-appending under the same field and multiplying it further.
                # A genuinely new per-job/per-client occurrence (the case this
                # append exists for) is never a substring of what's already
                # collected, so this only ever suppresses true repeats.
                if text in existing or existing in text:
                    if len(text) > len(existing):
                        verbatim[field] = text
                    logger.info(f"Verbatim '{field}' duplicate {len(text)} chars from '{heading}' skipped")
                else:
                    # Same field, another confident occurrence — append rather than
                    # overwrite (would silently lose earlier content) or drop (would
                    # silently lose this content), since resumes commonly repeat a
                    # section like this once per job/client rather than only once total.
                    verbatim[field] = existing + "\n" + text
                    logger.info(f"Verbatim '{field}' +{len(text)} chars appended from '{heading}'")
            continue
        if text:
            verbatim[field] = text
            if heading.upper() == "QUALIFICATIONS":
                _ambiguous_field_source.add(field)
            logger.info(f"Verbatim '{field}' ({len(text)} chars) from '{heading}'")

    for field, text in verbatim.items():
        result[field] = text

    # ── 2a-pre-0: don't let the base regex parser's independent "projects"
    # guess survive next to a heading-loop-recognized Experience section. The
    # base parser (parse_resume_text, run earlier as the fast first pass) uses
    # its own, separate heading-alias table and can classify some Experience
    # content as "projects" via a heading/pattern the heading-loop doesn't
    # recognize as project-specific — since the loop found real structure
    # here (it set "experience") but nothing it saw mapped to "projects",
    # whatever the base parser guessed is very often the same content already
    # sitting in Experience, just duplicated. Matches the standing rule: no
    # heading-loop-recognized Project Details heading → Projects stays empty.
    if "experience" in verbatim and "projects" not in verbatim and result.get("projects"):
        logger.info(
            "Clearing base-parser 'projects' guess (%d chars) — no Project "
            "Details heading recognized by the heading loop",
            len(result["projects"]),
        )
        result["projects"] = ""

    # ── 2a-pre: rescue Experience content stranded past a Skills heading. On
    # 2-column resumes where a Contact+Skills sidebar sits mid-page, its
    # heading's own content span runs right into whatever Experience bullets
    # happen to follow it before the next real heading (e.g. Education) —
    # nothing marks where the sidebar's own skill items end and the
    # continuation of the interrupted job entry begins. A bulleted line inside
    # Skills is only treated as stray Experience when it also READS like a full
    # job-responsibility sentence (long, or ending in sentence-terminal
    # punctuation) — plenty of legitimate Skills sections bullet their own
    # short category items (e.g. "•  DHF / DMR Authorship & Review"), and those
    # must never be mistaken for stray Experience just for having a bullet.
    if result.get("skills"):
        _sk_lines = result["skills"].split("\n")
        _bullet_idx = next(
            (i for i, l in enumerate(_sk_lines) if _looks_like_stray_experience_bullet(l)),
            None,
        )
        if _bullet_idx is not None and _bullet_idx > 0:
            _stray = "\n".join(_sk_lines[_bullet_idx:]).strip()
            result["skills"] = "\n".join(_sk_lines[:_bullet_idx]).strip()
            if _stray:
                result["experience"] = (result.get("experience", "").rstrip() + "\n" + _stray).strip()
                logger.info(
                    "Rescued %d line(s) of stray Experience content stranded past Skills",
                    len(_sk_lines) - _bullet_idx,
                )

    # ── 2a-pre-2: same idea, symmetric case — stray Skills content stranded
    # past an Education heading (e.g. a sidebar Skills list wraps onto a
    # second column/page and its tail lands after Education's own content,
    # before the next real heading). Genuine Education lines always carry a
    # year, a degree abbreviation, a CGPA/GPA marker, or an institution
    # keyword; once those stop appearing, trailing lines with none of those
    # signals are stray tool/skill names, not more education.
    if result.get("education"):
        _edu_marker_re = re.compile(
            r'\b(19|20)\d{2}\b|'
            r'\b(b\.?\s?e|b\.?\s?tech|m\.?\s?tech|mba|mca|bca|b\.?\s?sc|m\.?\s?sc|'
            r'b\.?\s?com|m\.?\s?com|phd|ph\.?d|cgpa|gpa)\b|'
            r'\b(university|college|institute|school|academy)\b|'
            r'\d+(?:\.\d+)?\s*%|'  # a bare grade/percentage line (e.g. "63%") is still education
            # A bare CGPA/GPA value (e.g. "7.5)", "9.0") is still education — but a
            # DOB-style date (e.g. "30.11.1995") must NOT match: exclude when a
            # third ".digit" group follows or precedes (a real CGPA is exactly
            # one decimal point, never three dot-separated number groups).
            r'(?<!\d\.)\b\d{1,2}\.\d{1,2}\b(?!\.\d)',
            re.IGNORECASE,
        )
        _edu_lines = result["education"].split("\n")
        _last_marker_idx = None
        for _ei, _el in enumerate(_edu_lines):
            if _edu_marker_re.search(_el):
                _last_marker_idx = _ei
        if _last_marker_idx is not None and _last_marker_idx < len(_edu_lines) - 1:
            _edu_tail = [
                l for l in _edu_lines[_last_marker_idx + 1:]
                if len(l.strip(" .")) > 1  # drop stray punctuation-only artifact lines
            ]
            # The tail past the last education marker is just as often a Declaration/
            # signature/personal-info block (e.g. "Declaration" / "I hereby declare..."
            # / "Place: ... <name>") as it is a genuine stray skills list — and that
            # content must never be rescued into Skills. Cut the tail at the first such
            # marker, mirroring the boundary the later "2a-post" Education cleanup uses,
            # so this rescue can't fire before that cleanup gets a chance to apply it.
            _edu_tail_pi_re = re.compile(
                r"^(?:personal(?:\s+(?:information|details?))?|other\s+personal|"
                r"declaration|i\s+hereby\s+declare\b.*|place\s*:.*|"
                r"father|mother|parent|guardian|date\s+of\s+birth|dob|gender|"
                r"marital\s+status|nationality|passport\s+no\.?|"
                # A 2-column sidebar can wrap "PERSONAL INFORMATION" so "PERSONAL"
                # lands on one line and "INFORMATION" lands on the next, fused with
                # whatever main-column content shares that row (e.g. "INFORMATION
                # Gender: Female") — recognized by "information" followed shortly by
                # an actual personal-detail field, not just the bare word alone
                # (which could otherwise misfire on real content like "Information
                # Technology").
                r"information\b.{0,30}(?:gender|nationality|date\s+of\s+birth|"
                r"marital\s+status))\b",
                re.I,
            )
            _pi_idx = next(
                (i for i, l in enumerate(_edu_tail) if _edu_tail_pi_re.match(l.strip())),
                None,
            )
            if _pi_idx is not None:
                _edu_tail = _edu_tail[:_pi_idx]
            _edu_stray = "\n".join(_edu_tail).strip()
            if _edu_stray:
                result["education"] = "\n".join(_edu_lines[:_last_marker_idx + 1]).strip()
                result["skills"] = (result.get("skills", "").rstrip() + "\n" + _edu_stray).strip()
                logger.info(
                    "Rescued %d line(s) of stray Skills content stranded past Education",
                    len(_edu_tail),
                )

    # ── 2b-pre: prefer the OCR-recovered About Me paragraph over a column-split,
    # fragmented native extraction. OCR reads the page in true visual order, so
    # when the text layer's own extraction splits this paragraph across a column
    # boundary (half stays near the header, the other half lands elsewhere after
    # a sidebar section), the native "summary" ends up short and grammatically
    # broken while the OCR version reconstructs the whole sentence correctly.
    _m_about_ocr = re.search(r'(?m)^\s*About\s*Me\s*:\s*(.+)$', raw_text)
    if _m_about_ocr:
        _about_ocr_text = _m_about_ocr.group(1).strip()
        if len(_about_ocr_text) > len(result.get("summary", "").replace("\n", " ").strip()):
            result["summary"] = _about_ocr_text
            logger.info(
                "Summary replaced with OCR-recovered About Me paragraph (%d chars)",
                len(_about_ocr_text),
            )

    # ── 2c-pre: rescue a numbered Work Experience tail stranded past an unmapped
    # heading. On 2-column PDFs, page 2's sidebar (Other Personal Details / Hobbies /
    # Languages / Extra Curricular) is often read BEFORE the main column's
    # continuation of Experience, so the last bullets land right after one of those
    # unrecognised headings — where they'd otherwise be dropped, since the heading
    # maps to no field. Detected by: Experience ends on a numbered bullet "N." and
    # an unmapped chunk contains the next bullet "N+1." — we then splice that
    # chunk's "N+1." onward (up to any following recognised heading, which already
    # bounds the chunk) back onto Experience. Guarded by check_parser.py so it
    # cannot silently break again (this variable was previously never built).
    if result.get("experience") and _unmapped_chunks:
        _exp_nums = re.findall(r'(?:^|\n)\s*(\d+)\.\s', result["experience"])
        if _exp_nums:
            _next_num = int(_exp_nums[-1]) + 1
            for _chunk in _unmapped_chunks:
                _m = re.search(rf'(?:^|\n)\s*{_next_num}\.\s', _chunk)
                if _m:
                    _continuation = _chunk[_m.start():].strip()
                    result["experience"] = result["experience"].rstrip() + "\n" + _continuation
                    logger.info(
                        "Rescued numbered Work Experience continuation (from #%d) "
                        "stranded past an unmapped heading",
                        _next_num,
                    )
                    break

    # ── 2a-preamble: No summary heading found — look for a professional paragraph in the
    # preamble (text before the first recognised section heading). Covers CVs that use a
    # bold specialty title (e.g. "REGULATORY AFFAIRS & QUALITY ASSURANCE") instead of an
    # explicit "PROFESSIONAL SUMMARY" heading. ──────────────────────────────────────────
    if not result.get("summary") and headings:
        _first_h_pos = headings[0][0]
        _preamble_raw = raw_text[:_first_h_pos].strip()
        if _preamble_raw:
            _contact_pat = re.compile(r'[@|\\]|\+\d{6,}|https?://', re.I)
            _allcaps_pat = re.compile(r'^[A-Z][A-Z0-9\s&+/.,():-]+$')
            _para_lines = []
            for _pl in _preamble_raw.splitlines():
                _pl = _pl.strip()
                if not _pl:
                    _para_lines = []  # blank line → paragraph boundary, reset
                    continue
                if _contact_pat.search(_pl) or (_allcaps_pat.match(_pl) and len(_pl) > 3):
                    _para_lines = []  # name/contact/all-caps-title line → reset
                    continue
                _para_lines.append(_pl)
            _para_text = " ".join(_para_lines).strip()
            if len(_para_text) >= 80 and re.search(
                r'\b(experience|expertise|professional|proven|results|background|'
                r'specialist|dedicated|accomplished|years\s+of|proficient)\b',
                _para_text, re.I,
            ):
                result["summary"] = _para_text
                logger.info("Summary extracted from preamble paragraph (%d chars)", len(_para_text))

    # ── 2a-post. Clean verbatim skills: remove website headings, URLs, split merged bullets ──
    if result.get("skills"):
        _vskill_url = re.compile(
            r"^(https?://|www\.|linkedin|github|gitlab|portfolio|website|blog)\S*$",
            re.I,
        )
        _vskill_hdr = re.compile(
            r"^(websites?(\s*,\s*|\s+)portfolios?(\s+and\s+profiles?)?|websites?|"
            r"portfolios?\s+and\s+profiles?|and\s+profiles?|profiles?\s+and|"
            r"personal\s+information|personal\s+details?|declaration|"
            r"additional\s+information|other\s+information)\s*$",
            re.I,
        )
        vsk_lines = []
        for sl in result["skills"].splitlines():
            s = sl.strip()
            if not s:
                continue
            if _vskill_url.match(s):
                continue
            bare = re.sub(r"[^a-zA-Z\s]", " ", s).strip()
            if _vskill_hdr.match(bare):
                continue
            # Split merged bullet items: "• Skill A • Skill B" → two separate items
            if s.count("•") > 1 or s.count("●") > 1:
                for p in re.split(r"\s*[•●▪◦]\s*", s):
                    p = p.strip()
                    if p:
                        vsk_lines.append(p)
            else:
                vsk_lines.append(s.lstrip("•● -").strip())
        # Remove the person's own name/title and column-gutter-bleed fragments
        # (e.g. "Design Engineer" bleeding in from the right-column header at the
        # same y-level as the tail of the Skills sidebar).
        vsk_lines = _drop_name_title_and_gutter_bleed(
            vsk_lines, result.get("full_name", ""), result.get("title", "")
        )
        # Cut a bled-in next-section heading (e.g. a "Languages" sidebar block read
        # right after skills), then drop a trailing single-word wrap fragment
        # ("Medical" left from "Medical Device Classification").
        vsk_lines = _truncate_skills_at_bleed(vsk_lines)
        vsk_lines = _drop_trailing_skill_fragments(vsk_lines)
        result["skills"] = "\n".join(l for l in vsk_lines if l.strip())

    # ── 2a-post. Clean verbatim education: stop at PERSONAL INFORMATION, remove personal lines ──
    if result.get("education"):
        # "PERSONAL INFORMATION" can be on one line OR split across two lines
        _edu_stop_pi = re.compile(
            r"\n[ \t]*(?:PERSONAL(?:[ \t]+(?:INFORMATION|DETAILS?))?|"
            r"OTHER[ \t]+PERSONAL|DECLARATION)\s*(?:\n|$)",
            re.I,
        )
        _em_pi = _edu_stop_pi.search(result["education"])
        if _em_pi:
            trimmed_pi = result["education"][:_em_pi.start()].strip()
            if trimmed_pi:
                logger.info(f"Education trimmed at PERSONAL INFORMATION boundary")
                result["education"] = trimmed_pi
        _edu_pi_pat = re.compile(
            r"\b(father|mother|parent|guardian|date.{0,5}birth|born|dob\b|gender|sex\b|"
            r"marital|nationality|citizen|passport|religion|caste|languages\s+known|"
            r"permanent\s+address|current\s+address|pin\s*code|aadhar|pan\b|"
            r"place\s+of\s+birth|age\b|blood\s+group|declaration|personal\s+information|"
            r"personal\s+details?)\b",
            re.I,
        )
        # Also drop standalone noise words like bare "PERSONAL" or "INFORMATION"
        _edu_noise = {"PERSONAL", "INFORMATION", "DETAILS", "DECLARATION", "OTHER"}
        edu_clean = [
            l for l in result["education"].splitlines()
            if not _edu_pi_pat.search(l) and l.strip().upper() not in _edu_noise
        ]
        result["education"] = "\n".join(edu_clean).strip()

    # ── 2a-edu-garbage: If education content looks like experience, clear or rescue ──
    # Symptom: label-sidebar CVs with multiple page labels (e.g. "Qualifications" on
    # page 2 sidebar) cause experience entries from later pages to land in the education
    # field.  Two sub-cases:
    #   A) Content has experience date ranges (MONTH YEAR – MONTH YEAR/PRESENT) →
    #      it's real experience that got mis-routed → merge back into experience, clear edu.
    #   B) Content has work-verb bullets but NO date ranges and NO edu keywords →
    #      it's a professional-skills/qualifications blurb → clear, let AI find real education.
    if result.get("education"):
        _edu_text = result["education"]
        _has_edu_kw = re.search(
            r'\b(university|college|institute|school|bachelor|master|b\.e\b|b\.tech\b|'
            r'b\.com\b|m\.tech\b|m\.e\b|mba\b|phd\b|diploma|degree|cgpa|gpa|'
            r'10\+2|hsc\b|sslc\b|12th|10th|graduation|affiliated)\b',
            _edu_text, re.I,
        )
        # Count experience-style lines — bullet char is optional and may be any symbol font glyph.
        _exp_verbs = re.findall(
            r'(?:^|\n)[ \t]*[^\w\s]?[ \t]*(?:Expertise|Managed|Performed|Prepared|Preparati|'
            r'Worked|Working|Ensure|Contribut|Execut|Conduct|Complet|Monitor|Analys|'
            r'Maintained|Develop|Review|Coordinat|Having|Exposure|Responsible|'
            r'Handling|Support|Assist|Troubleshoot|Involved)',
            _edu_text,
        )
        if not _has_edu_kw and len(_exp_verbs) >= 2:
            # Check whether the content contains actual experience date ranges
            # (month name + year → dashes → month/year/present).
            _exp_date_re = re.compile(
                r'(?:JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|JUN(?:E)?|'
                r'JUL(?:Y)?|AUG(?:UST)?|SEP(?:TEMBER)?|OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)'
                r'\s+\d{4}\s*[-–]',
                re.I,
            )
            if _exp_date_re.search(_edu_text):
                # Real experience entries mis-routed to education — rescue them.
                logger.info(
                    "Education field contains experience entries (date ranges found) "
                    "— merging into experience and clearing education"
                )
                if result.get("experience"):
                    result["experience"] = result["experience"] + "\n" + _edu_text
                else:
                    result["experience"] = _edu_text
            else:
                logger.info(
                    "Education field contains %d experience-style bullets (no date ranges) "
                    "— clearing for AI fallback", len(_exp_verbs)
                )
            result["education"] = ""

        elif _has_edu_kw and len(_exp_verbs) >= 2:
            # Mixed case: education has real edu content AND trailing experience bullets.
            # Cause: sidebar PDF page 2 where left column (EDUCATION + degree text) is
            # joined before right column (experience continuation bullets).
            # Fix: find the last line containing an edu keyword; everything after it
            # that looks like experience bullets gets rescued into experience.
            _edu_kw_line_re = re.compile(
                r'\b(university|college|institute|school|bachelor|master|'
                r'b\.e\b|b\.tech\b|b\.com\b|m\.tech\b|diploma|degree|'
                r'cgpa|gpa|graduation|polytechnic)\b', re.I
            )
            _edu_lines_split = _edu_text.splitlines()
            _last_edu_kw_idx = -1
            for _eli, _eln in enumerate(_edu_lines_split):
                if _edu_kw_line_re.search(_eln):
                    _last_edu_kw_idx = _eli
            if _last_edu_kw_idx >= 0 and _last_edu_kw_idx < len(_edu_lines_split) - 1:
                _trailing_exp = [l for l in _edu_lines_split[_last_edu_kw_idx + 1:] if l.strip()]
                _trailing_text = "\n".join(_trailing_exp).strip()
                if len(_trailing_text) > 30:
                    result["education"] = "\n".join(_edu_lines_split[:_last_edu_kw_idx + 1]).strip()
                    result["experience"] = (result.get("experience", "") + "\n" + _trailing_text).strip()
                    logger.info(
                        "Mixed edu+exp: moved %d trailing lines to experience", len(_trailing_exp)
                    )

    # ── 2b-proj-garbage: If projects content looks like experience (date ranges +
    # job-title/company lines), merge into experience and clear projects. ──────────
    if result.get("projects"):
        _proj_text = result["projects"]
        _proj_date_re = re.compile(
            r'(?:'
            r'(?:JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|JUN(?:E)?|'
            r'JUL(?:Y)?|AUG(?:UST)?|SEP(?:TEMBER)?|OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)'
            r'\s+\d{4}\s*[-–]'
            r'|\d{1,2}/\d{4}\s*[-–]'
            r')',
            re.I,
        )
        # Also look for ALL-CAPS job-title lines (e.g. "PROGRAMMER ANALYST, 06/2018")
        _proj_jobtitle_re = re.compile(
            r'(?:^|\n)[A-Z][A-Z\s/]+,\s*\d{2}/\d{4}',
        )
        if _proj_date_re.search(_proj_text) and _proj_jobtitle_re.search(_proj_text):
            logger.info(
                "Projects field contains work experience entries — merging into experience and clearing projects"
            )
            if result.get("experience"):
                result["experience"] = result["experience"] + "\n" + _proj_text
            else:
                result["experience"] = _proj_text
            result["projects"] = ""

    # -- 2c-skills-in-experience: sidebar resumes sometimes repeat their Skills
    # column on later pages WITHOUT repeating the "SKILLS" heading, so that
    # continuation has no marker distinguishing it from whatever section is
    # active at that point in the text -- usually Experience. Move it to Skills.
    if result.get("experience"):
        _clean_exp, _rescued_text = _rescue_skill_sidebar_from_experience(result["experience"])
        if _rescued_text:
            _sep = "\n"
            result["skills"] = (result["skills"] + _sep + _rescued_text) if result.get("skills") else _rescued_text
            result["experience"] = _clean_exp
            logger.info(
                "Rescued %d line(s) of unlabeled skills-sidebar continuation from Experience into Skills",
                len(_rescued_text.splitlines()),
            )

    # ── 2d-cert-overrun: the initial base regex parse (parse_resume_text, at the
    # very top of this function) does its own independent section split and can
    # set "certifications" from a colon-labelled line (e.g. "Certification : ...")
    # even when there's no real standalone CERTIFICATIONS heading. On resumes with
    # no blank lines between paragraphs, that independent split sometimes fails to
    # stop at the next real heading and overruns into Education/Achievements —
    # producing a field that's really just a duplicate of content already correct
    # elsewhere. Detected by: the "certifications" text fully contains both the
    # (already correct) Education and Achievements text.
    _cert_val = result.get("certifications", "").strip()
    if _cert_val:
        _edu_val = result.get("education", "").strip()
        _ach_val = result.get("achievements", "").strip()
        if _edu_val and _ach_val and _edu_val in _cert_val and _ach_val in _cert_val:
            logger.info(
                "Certifications field duplicates Education+Achievements "
                "(base-parse boundary overrun) — clearing"
            )
            result["certifications"] = ""

    # ── 2a-post. Reject title if it looks like a sentence fragment ────────────────
    _title_check = result.get("title", "").strip()
    if _title_check:
        _tw0 = _title_check.split()[0]
        _is_frag = (
            _tw0[:1].islower()            # starts lowercase → mid-sentence fragment
            or _tw0.lower().endswith("ing")  # gerund (e.g. "enhancing", "managing")
            or ("." in _title_check[:-1])    # period in the middle → sentence boundary
            or len(_title_check.split()) > 7  # too long for a title
        )
        if _is_frag:
            logger.info(f"Title '{_title_check[:60]}' looks like sentence fragment — clearing")
            result["title"] = ""

    # ── 2b. Post-process education: stop at project/experience boundary ───────
    # Handles resumes (e.g. Arun Sivakumar) where "Relevant Project Experience"
    # is directly below EDUCATION with no blank line — so heading detection missed
    # it as a boundary and the full project table was pulled into education.
    if result.get("education"):
        _edu_trim_re = re.compile(
            r'(?:\n|^)[ \t]*(?:Relevant\s+)?Project\s+Experience[:\s]*(?=\n|$)'
            r'|(?:\n|^)[ \t]*Work\s+Experience[:\s]*(?=\n|$)'
            r'|(?:\n|^)[ \t]*Employment\s+History[:\s]*(?=\n|$)',
            re.IGNORECASE,
        )
        _em = _edu_trim_re.search(result["education"])
        if _em:
            trimmed = result["education"][:_em.start()].strip()
            if trimmed:
                logger.info(
                    f"Education trimmed at project/experience boundary "
                    f"({len(result['education'])} → {len(trimmed)} chars)"
                )
                result["education"] = trimmed

    # ── 2b-fix. Restore truncated last lines (education + experience) ───────────
    # The two-column text splitter can clip the tail of a line that spans the
    # full page width, e.g. "Maharaja Engineering Colle" instead of
    # "Maharaja Engineering College, Coimbatore - 2016 to 2010".
    # Heuristic: if the last line ends with a plain letter (no punctuation,
    # digit, or closing bracket) the line was cut mid-word.
    # Fix: re-read each page with pdfplumber extract_text() (no column split)
    # and replace the truncated line with the full version found there.
    def _restore_truncated_last_line(field_text, full_page_text):
        last_line = field_text.rsplit("\n", 1)[-1].strip()
        if not last_line or not last_line[-1].isalpha():
            return field_text
        prefix = last_line[:20]
        idx = full_page_text.find(prefix)
        if idx == -1:
            return field_text
        line_end = full_page_text.find("\n", idx)
        full_line = full_page_text[
            idx : line_end if line_end != -1 else idx + 300
        ].strip()
        if len(full_line) > len(last_line):
            return field_text[: -len(last_line)] + full_line
        return field_text

    if str(path).lower().endswith(".pdf") and any(
        result.get(f) for f in ("education", "experience")
    ):
        try:
            import pdfplumber as _fix_plumber
            _fix_full_text = ""
            with _fix_plumber.open(str(path)) as _fix_pdf:
                for _fix_pg in _fix_pdf.pages:
                    _fix_full_text += (_fix_pg.extract_text() or "") + "\n"
            for _fix_field in ("education", "experience"):
                if result.get(_fix_field):
                    _before = result[_fix_field]
                    _after = _restore_truncated_last_line(_before, _fix_full_text)
                    if _after != _before:
                        result[_fix_field] = _after
                        logger.info(
                            "%s last line restored: %r → %r",
                            _fix_field,
                            _before.rsplit("\n", 1)[-1].strip(),
                            _after.rsplit("\n", 1)[-1].strip(),
                        )
        except Exception as _fix_exc:
            logger.debug("Truncated last line restore failed: %s", _fix_exc)

    # ── 2c. Styled-bar contact scan ────────────────────────────────────────────
    # For PDFs with icon-font contact bars (dark background, envelope/phone icons)
    # the main pymupdf pass sometimes fails to decode the email/phone characters.
    # Run a dedicated pdfplumber word-box scan on page 1 before the AI fallback.
    if str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        _extra_contact = _pdfplumber_contact_scan(path)
        for _cf, _cv in _extra_contact.items():
            if not result.get(_cf) and _cv:
                result[_cf] = _cv
                logger.info(f"Styled-bar contact scan filled '{_cf}': {_cv}")

    # ── 2d. Garbage detection — clear fields that contain header metadata so the
    # AI fallback can re-extract them correctly.  This handles layouts where the
    # person's all-caps name (e.g. "SIVARANJANI D") trips the company-name heuristic
    # in find_sections and causes header lines to fill the experience field.
    _header_meta_re = re.compile(
        r'total\s+experience|current\s+employer|current\s+designation'
        r'|current\s+ctc|expected\s+ctc|notice\s+period',
        re.IGNORECASE,
    )
    _exp_val = result.get("experience", "")
    if _exp_val and _header_meta_re.search(_exp_val):
        logger.info(
            f"Clearing garbage experience (header metadata detected) — AI fallback will fill it"
        )
        result["experience"] = ""

    # A single-word title that is not a recognised role descriptor is likely a
    # fragment from a split header box (e.g. "Product" from "Product Regulatory
    # Analyst" truncated at the column gutter).  Clear it so the next step can fill it.
    _title_val = result.get("title", "").strip()
    _role_fragment_ok = re.compile(
        r'\b(analyst|engineer|lead|manager|consultant|developer|specialist|'
        r'designer|director|officer|architect|executive|scientist|associate|'
        r'coordinator|technician|programmer|advisor|writer|researcher)\b',
        re.IGNORECASE,
    )
    if _title_val and len(_title_val.split()) == 1 and not _role_fragment_ok.search(_title_val):
        logger.info(
            f"Clearing single-word title fragment '{_title_val}' — regex/AI fallback will fill it"
        )
        result["title"] = ""

    # Try to recover title directly from "Current Designation : <value>" or
    # "Designation : <value>" pattern in the raw text — reliable for header-box resumes.
    if not result.get("title"):
        # Iterate all Designation matches: the header may have a truncated value
        # (e.g. "Current Designation : Product" split across columns) while the
        # work-experience section further down has the full title.
        # Use the first match that is at least 2 words or contains a role keyword.
        for _desig_m in re.finditer(
            r'(?:current\s+)?designation\s*[:\-]\s*(.+)',
            raw_text[:5000], re.IGNORECASE
        ):
            _desig_val = _desig_m.group(1).strip().split('\n')[0].strip()
            if _desig_val and (len(_desig_val.split()) >= 2 or _role_fragment_ok.search(_desig_val)):
                result["title"] = _desig_val
                logger.info(f"Title from Designation regex: '{_desig_val}'")
                break

    # Title last-resort before AI: ALL-CAPS role title at the top of experience
    # (e.g. "VALIDATION ENGINEER, 11/2021-Current") — avoids an AI call entirely.
    if not result.get("title") and result.get("experience"):
        _exp_fl = result["experience"].strip().split('\n')[0].strip()
        _exp_title_m = re.match(
            r'^([A-Z][A-Z0-9\s/&\-]{2,45}?)(?:\s*[,\-–]\s*\d|\s*$)', _exp_fl
        )
        if _exp_title_m:
            _cand = _exp_title_m.group(1).strip()
            if _role_fragment_ok.search(_cand):
                result["title"] = _cand.title()
                logger.info(f"Title from experience first line: '{result['title']}'")

    # Same idea, generalized to Title-Case entries (e.g. "Senior Lead Design
    # Engineer / Onsite Engineer | Onward Technologies Pvt Ltd for Caterpillar
    # Inc") — the ALL-CAPS regex above only matches all-uppercase job lines;
    # most resumes write the role in Title Case instead. Split on the same
    # separator used elsewhere for title/name detection and take the segment
    # before the company/date part.
    if not result.get("title") and result.get("experience"):
        _exp_fl2 = result["experience"].strip().split('\n')[0].strip()
        _exp_title_seg = re.split(r'\s+[-–|]\s+', _exp_fl2, maxsplit=1)[0].strip()
        if (_exp_title_seg and _exp_title_seg != _exp_fl2
                and 1 <= len(_exp_title_seg.split()) <= 10
                and not re.search(r'\d{4}|\d{1,2}/\d{1,2}', _exp_title_seg)
                and _role_fragment_ok.search(_exp_title_seg)):
            result["title"] = _exp_title_seg[:80]
            logger.info(f"Title from experience first line (Title-Case): '{result['title']}'")

    # ── 2c. Identity fields (name/email/phone/location/linkedin) via a dedicated
    # LLM call, run BEFORE the content-field budget below starts. Previously these
    # were mixed into the same budget-limited loop as summary/experience/skills —
    # since those heavier fields ran first and could each take several seconds on
    # the local model, the shared budget was usually exhausted before the loop
    # reached email/phone, and full_name got only one shot gated by the same clock.
    # location/linkedin had no AI fallback at all. This call is unconditional and
    # untimed against that budget, so identity extraction always gets to run.
    _missing_identity = [f for f in _IDENTITY_FIELDS if not result.get(f)]
    if _missing_identity:
        _identity_fill = _ollama_identity_extract(raw_text, _missing_identity)
        for _id_field, _id_val in _identity_fill.items():
            result[_id_field] = _id_val
            logger.info(f"Identity AI filled '{_id_field}': {_id_val}")

    # Labelled "City <x>" / "Country <y>" pair — a distinct format from the
    # hardcoded-city-name scan in parse_resume_text, and one that can legitimately
    # sit anywhere in the document (e.g. under a sidebar "OTHER PERSONAL DETAILS"
    # heading well past the header), so it's checked here against the FULL text
    # rather than the identity call's header-sized window.
    if not result.get("location"):
        _city_m = re.search(r'(?:^|\n)\s*City[ \t]+([^\n]+)', raw_text, re.I)
        _country_m = re.search(r'(?:^|\n)\s*Country[ \t]+([^\n]+)', raw_text, re.I)
        if _city_m or _country_m:
            _loc_parts = [m.group(1).strip() for m in (_city_m, _country_m) if m]
            result["location"] = ", ".join(_loc_parts)[:180]
            logger.info(f"Location from City/Country label pair: {result['location']}")

    # ── 3. AI fallback — fires only when a field is still empty ─────────────────
    _AI_FALLBACK = {
        "title": (
            "What is the current professional job title of the person in this resume? "
            "Look for lines labelled 'Current Designation:', 'Designation:', 'Title:', "
            "or a role title just below the person's name. "
            "Return ONLY the job title (e.g. 'Product Regulatory Analyst'). "
            "Do NOT return a sentence, a degree, or a summary.\n\nRESUME HEADER:\n",
            raw_text[:1200], False, 25, 512,
        ),
        "summary": (
            "Extract the professional summary section verbatim from this resume. "
            "Do not paraphrase.\n\nRESUME:\n",
            raw_text[:3000], False, 600, 2048,
        ),
        "experience": (
            "Extract the work experience from this resume. "
            "The experience section may NOT have a heading like 'Work Experience' — "
            "look for company/employer names (often in ALL CAPS or Title Case) followed by "
            "Duration, Designation/Title, and bullet-point responsibilities. "
            "Return ONLY the verbatim work experience text — company, dates, role, and bullets. "
            "Do NOT return the profile summary or certifications.\n\nRESUME:\n",
            raw_text, False, 1600, 4096,
        ),
        "skills": (
            "This resume's Skills/Technical Skills/Core Competencies section could not be "
            "found by normal parsing. Read the whole resume (summary, experience, tools "
            "mentioned) and identify the specific skills, tools, technologies, and "
            "methodologies the candidate ACTUALLY used or names — e.g. if it says "
            "'experience on all ALM modules like Requirements, Test Lab, Test Plan and "
            "Defects', extract each of those as its own item. "
            "Every item must be traceable to something explicitly written in the resume — "
            "never add a generic skill for the candidate's role/domain that isn't actually "
            "named in the text. Never repeat the same skill twice. Keep each item short "
            "(a skill/tool name, not a full sentence). If truly nothing skill-like is "
            "mentioned anywhere, return an empty array. "
            'Return ONLY JSON: {"skills":["s1","s2",...]} or {"skills":[]}.\n\nRESUME:\n',
            raw_text[:3000], True, 400, 1024,
        ),
        "education": (
            "Extract the education section verbatim from this resume.\n\nRESUME:\n",
            raw_text, False, 300, 1024,
        ),
    }

    import time as _time
    _ai_budget_start = _time.monotonic()

    for field, (prompt, ctx_text, as_json, num_predict, num_ctx) in _AI_FALLBACK.items():
        if result.get(field):
            continue  # already filled — skip AI
        if _time.monotonic() - _ai_budget_start > _OLLAMA_BUDGET_SECS:
            logger.warning(f"AI budget exhausted — skipping remaining fallback fields")
            break
        logger.info(f"AI fallback for '{field}'")
        try:
            raw = _ollama_chat(prompt + ctx_text, as_json=as_json,
                               num_predict=num_predict, num_ctx=num_ctx)
            if as_json:
                items = raw.get("skills", []) if isinstance(raw, dict) else []
                if items:
                    result[field] = "\n".join(str(s) for s in items if str(s).strip())
            else:
                if raw and raw.strip():
                    result[field] = raw.strip()
            if result.get(field):
                logger.info(f"AI filled '{field}' ({len(result[field])} chars)")
        except Exception as e:
            logger.warning(f"AI fallback '{field}' failed: {e}", exc_info=True)

    # ── 3b. Post-AI title validation ───────────────────────────────────────────
    # The AI sometimes returns a summary paragraph or "SUMMARY: ..." as the title.
    # Apply the same fragment check used at step 2a-post; if invalid, try to
    # recover the title from the first line of the experience section.
    _ai_title = result.get("title", "").strip()
    if _ai_title:
        _tw0 = _ai_title.split()[0] if _ai_title.split() else ""
        _ai_title_bad = (
            _tw0[:1].islower()
            or _tw0.lower().endswith("ing")
            or "." in _ai_title[:-1]
            or len(_ai_title.split()) > 7
            or ":" in _tw0           # "SUMMARY:" / "OBJECTIVE:" type prefix
            or _ai_title.startswith("```")  # markdown block
        )
        if _ai_title_bad:
            logger.info(f"Post-AI: title '{_ai_title[:60]}' looks like paragraph — clearing")
            result["title"] = ""

    # ── 3c. Post-AI content validation ─────────────────────────────────────────
    # When the AI can't literally find a section (no matching heading in the
    # text), a small model sometimes "helpfully" writes its own analysis or
    # summary of the resume instead of admitting nothing was found — e.g. a
    # numbered list like "5. **Technical Skills and Proficiency:**  - ..." for
    # what was supposed to be a verbatim Education extract. That's generated
    # commentary, not resume content, and must never be saved as if it were.
    _ai_meta_phrase_re = re.compile(
        r'\bbased on (?:the|this) resume\b|\bthe candidate\b|\bthis resume (?:shows|indicates|includes)\b'
        r'|\bi (?:can see|found|couldn\'?t find)\b|\bhere(?:\'s| is) (?:the|a)\b|\bno (?:education|experience|'
        r'summary|skills|certifications?) (?:section|information|details)? ?(?:found|available|listed)\b',
        re.I,
    )
    for _cf in ("summary", "experience", "education", "skills", "certifications"):
        _cv = result.get(_cf, "")
        if not _cv:
            continue
        _cv_bad = (
            "**" in _cv
            or _cv.startswith("```")
            or bool(re.search(r'(?:^|\n)\s*\d+\.\s+\*?\*?[A-Z][a-zA-Z ]+:\*?\*?', _cv))
            or bool(_ai_meta_phrase_re.search(_cv))
        )
        if _cv_bad:
            logger.warning(
                f"Post-AI validation: '{_cf}' looks like generated commentary, not verbatim "
                f"resume text — clearing ({_cv[:80]!r}...)"
            )
            result[_cf] = ""

    # ── 4. Post-AI contact validation ──────────────────────────────────────────
    # The AI model sometimes returns text fragments instead of real email/phone
    # (e.g. "sion Report." for email, or a hallucinated phone number).
    # Validate and clear any field that does not match the expected format.
    _valid_email_re = re.compile(
        r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'
    )
    _valid_phone_re = re.compile(
        r'^\+?[\d][\d\s\-\.\(\)]{6,24}$'
    )
    for _vf, _vr in [("email", _valid_email_re), ("phone", _valid_phone_re)]:
        _val = result.get(_vf, "").strip()
        if _val and not _vr.match(_val):
            logger.warning(f"Post-AI validation: clearing invalid {_vf}='{_val}'")
            result[_vf] = ""

    # ── 4b. Cross-validate AI phone: clear if its digits aren't in the raw text ─
    # The AI sometimes hallucinate a plausible-looking phone number that passes the
    # format check above but was never in the document.  Compare digit sequences:
    # if the AI phone's digits don't appear anywhere in the raw text, discard it.
    if result.get("phone"):
        _ai_ph_digits = re.sub(r'\D', '', result["phone"])
        _raw_digits   = re.sub(r'\D', '', raw_text)
        if len(_ai_ph_digits) >= 10 and _ai_ph_digits not in _raw_digits:
            logger.warning(
                f"Phone '{result['phone']}' digits not in raw text — AI hallucination, clearing"
            )
            result["phone"] = ""

    # ── 4c. Aggressive +91 / 10-digit regex on full raw text ──────────────────
    # Runs after validation + cross-check have cleared garbage values.
    # Targets Indian numbers explicitly (common in uploaded resumes).
    if not result.get("email") or not result.get("phone"):
        _email_agr_re = re.compile(
            r'\b[A-Za-z0-9][A-Za-z0-9._%+\-]*@[A-Za-z0-9.\-]+\.[A-Za-z]{2,6}\b'
        )
        _ph_91_re = re.compile(r'\+91[\s\-\.]?\d[\d\s\-\.]{7,14}\d')
        _ph_10_re = re.compile(r'(?<!\d)[6-9]\d{9}(?!\d)')

        if not result.get("email"):
            _agr_em = _email_agr_re.search(raw_text)
            if _agr_em:
                _em_val = _agr_em.group().strip().rstrip(".,;)")
                result["email"] = _em_val
                logger.info(f"Aggressive regex email: {_em_val}")

        if not result.get("phone"):
            _agr_ph = _ph_91_re.search(raw_text) or _ph_10_re.search(raw_text)
            if _agr_ph:
                result["phone"] = _agr_ph.group().strip()
                logger.info(f"Aggressive regex phone: {result['phone']}")

    # ── 4d. White-text pdfplumber scan (contact bar on dark background) ────────
    # pdfplumber chars expose non_stroking_color (fill color).  White text on a
    # dark background (navy contact bar) has color ≥ 0.7 in all channels.
    # This specifically targets the contact bar without touching the body text.
    if str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        _wt_contact = _pdfplumber_whitetext_scan(path)
        for _wtf, _wtv in _wt_contact.items():
            if not result.get(_wtf) and _wtv:
                result[_wtf] = _wtv
                logger.info(f"White-text scan filled '{_wtf}': {_wtv}")

    # ── 5. Final rawdict contact scan ──────────────────────────────────────────
    # Last resort: use pymupdf rawdict (char-level) to find email/phone that may
    # live in styled header bars with non-standard font encoding.
    # Runs only after validation has cleared garbage values.
    if str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        _rd_contact = _pymupdf_rawdict_scan(path)
        for _rdf, _rdv in _rd_contact.items():
            if not result.get(_rdf) and _rdv:
                result[_rdf] = _rdv
                logger.info(f"rawdict scan filled '{_rdf}': {_rdv}")

    # ── 5d. Inverted Tesseract OCR — highest accuracy for dark contact bars ──────
    # MUST run before step 6 (non-inverted OCR).  Step 6 struggles with white text
    # on a dark/navy background and produces OCR artefacts (e.g. "Barun.sivakumar"
    # instead of "arun.sivakumar").  Running the inverted pass first sets correct
    # values so step 6 and 6b are skipped entirely.
    if str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        _inv_ocr_early = _tesseract_inverted_ocr_text(path)
        if _inv_ocr_early.strip():
            if not result.get("email"):
                _em_early = _extract_email(_inv_ocr_early)
                if _em_early:
                    result["email"] = _em_early
                    logger.info(f"Step 5d inverted OCR email: {_em_early}")
            if not result.get("phone"):
                _ph_early = _extract_phone(_inv_ocr_early)
                if _ph_early:
                    _ph_early_digits = re.sub(r"\D", "", _ph_early)
                    if 10 <= len(_ph_early_digits) <= 15:
                        result["phone"] = _ph_early
                        logger.info(f"Step 5d inverted OCR phone: {_ph_early}")

    # ── 6. Tesseract OCR contact scan ──────────────────────────────────────────
    # pytesseract renders page 1 as an image and OCRs the top 35% — the only
    # reliable method for dark-background contact bars (white text + icon fonts)
    # that are invisible to all PDF text-layer extraction methods.
    # Runs as a final pass; fills email/phone directly from OCR output.
    if str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        _tess_contact = _tesseract_contact_scan(path)
        for _tf, _tv in _tess_contact.items():
            if not result.get(_tf) and _tv:
                result[_tf] = _tv
                logger.info(f"Tesseract OCR filled '{_tf}': {_tv}")

    # ── 6b. Inverted-image OCR + Ollama extraction ────────────────────────────
    # The standard tesseract pass (step 6) renders white-on-dark text which gives
    # Tesseract poor contrast.  This step:
    #   1. Re-renders at 300 DPI and INVERTS the image (dark→light, white→dark).
    #   2. Runs Tesseract on both the inverted RGB image and the inverted red channel
    #      (maximum contrast for navy backgrounds).
    #   3. Tries direct regex on the combined OCR output.
    #   4. If email/phone are still missing, passes the OCR text to Ollama to
    #      extract and return them as JSON — Ollama acts as a smart OCR post-processor.
    if str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        _inv_ocr = _tesseract_inverted_ocr_text(path)
        if _inv_ocr.strip():
            # Step A: direct regex on inverted OCR text
            if not result.get("email"):
                _em_inv = _extract_email(_inv_ocr)
                if _em_inv:
                    result["email"] = _em_inv
                    logger.info(f"Inverted OCR direct email: {_em_inv}")
            if not result.get("phone"):
                _ph_inv = _extract_phone(_inv_ocr)
                if _ph_inv:
                    _ph_digits = re.sub(r"\D", "", _ph_inv)
                    if 10 <= len(_ph_digits) <= 15:
                        result["phone"] = _ph_inv
                        logger.info(f"Inverted OCR direct phone: {_ph_inv}")

            # Step B: Ollama extracts from OCR text when regex still failed
            if not result.get("email") or not result.get("phone"):
                _ollama_ocr_prompt = (
                    "The text below was OCR'd from the contact bar of a resume. "
                    "Find the email address and phone number and return ONLY this JSON:\n"
                    "{\"email\": \"the_email@domain.com\", \"phone\": \"+91 XXXX XXX XXX\"}\n\n"
                    "Rules:\n"
                    "- email must contain @ and end with .com / .in / .net etc.\n"
                    "- phone must be a number starting with +91 or a 10-digit mobile\n"
                    "- if you cannot find a field, use empty string \"\"\n\n"
                    "OCR TEXT:\n" + _inv_ocr[:600]
                )
                try:
                    _ocr_json = _ollama_chat(
                        _ollama_ocr_prompt, as_json=True, num_predict=80, num_ctx=1024
                    )
                    if isinstance(_ocr_json, dict):
                        _oj_email = str(_ocr_json.get("email", "")).strip().rstrip(".,;)")
                        _oj_phone = str(_ocr_json.get("phone", "")).strip()
                        _em_re = re.compile(
                            r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'
                        )
                        if not result.get("email") and _oj_email and _em_re.match(_oj_email):
                            result["email"] = _oj_email
                            logger.info(f"Ollama-OCR email: {_oj_email}")
                        if not result.get("phone") and _oj_phone:
                            _ph_d = re.sub(r"\D", "", _oj_phone)
                            if 10 <= len(_ph_d) <= 15:
                                result["phone"] = _oj_phone
                                logger.info(f"Ollama-OCR phone: {_oj_phone}")
                except Exception as exc:
                    logger.debug(f"Ollama-OCR extraction: {exc}")

    # ── 7. Ollama with OCR context ─────────────────────────────────────────────
    # If Tesseract found partial results or nothing, build a combined context from
    # Tesseract OCR + raw_text and ask Ollama specifically for the missing fields.
    # Response is validated strictly — must match email/phone format exactly.
    _step7_ok = _time.monotonic() - _ai_budget_start <= _OLLAMA_BUDGET_SECS
    if _step7_ok and str(path).lower().endswith(".pdf") and (not result.get("email") or not result.get("phone")):
        # Build combined context: page-1 simple text + raw_text header
        _ollama_ctx_parts = []
        try:
            import pymupdf as _fz_mod
            _doc = _fz_mod.open(str(path))
            _p1_simple = _doc[0].get_text("text")
            _doc.close()
            if _p1_simple.strip():
                _ollama_ctx_parts.append(_p1_simple[:800])
        except Exception:
            pass
        _ollama_ctx_parts.append(raw_text[:800])
        _ollama_ctx = "\n---\n".join(filter(None, _ollama_ctx_parts))

        _valid_email_final = re.compile(
            r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'
        )
        _valid_phone_final = re.compile(r'^\+?[\d][\d\s\-\.\(\)]{6,24}$')

        if not result.get("email"):
            try:
                _r = _ollama_chat(
                    "The resume below belongs to a person whose email address ends with "
                    ".com or .in. Find and return ONLY the email address — nothing else.\n\n"
                    "RESUME:\n" + _ollama_ctx,
                    as_json=False, num_predict=40, num_ctx=2048,
                )
                if _r and "@" in _r:
                    _em = re.search(
                        r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', _r
                    )
                    if (_em and _valid_email_final.match(_em.group())
                            and _identity_value_grounded("email", _em.group(), _ollama_ctx)):
                        result["email"] = _em.group()
                        logger.info(f"Ollama step-7 email: {result['email']}")
            except Exception as exc:
                logger.debug(f"Ollama step-7 email: {exc}")

        if not result.get("phone"):
            try:
                _r = _ollama_chat(
                    "The resume below belongs to a person with an Indian phone number "
                    "starting with +91 or a 10-digit mobile number. "
                    "Return ONLY the phone number — nothing else.\n\n"
                    "RESUME:\n" + _ollama_ctx,
                    as_json=False, num_predict=30, num_ctx=2048,
                )
                if _r:
                    _ph = re.search(r'[\+]?\d[\d\s\-\.]{7,20}\d', _r)
                    if _ph and _valid_phone_final.match(_ph.group()):
                        _ph_digits = re.sub(r'\D', '', _ph.group())
                        if (10 <= len(_ph_digits) <= 15
                                and _identity_value_grounded("phone", _ph.group(), _ollama_ctx)):
                            result["phone"] = _ph.group().strip()
                            logger.info(f"Ollama step-7 phone: {result['phone']}")
            except Exception as exc:
                logger.debug(f"Ollama step-7 phone: {exc}")

    # ── 8. Ollama vision model contact scan ───────────────────────────────────
    # Renders page-1 top 30 % as an image and sends it to a vision-capable Ollama
    # model (moondream, llava, etc.).  This is the ONLY method that can read
    # email/phone from dark-background contact bars where the text layer is absent.
    #
    # To activate: run  `ollama pull moondream`  once in your terminal.
    # The code auto-detects any installed vision model — no config needed.
    if (_time.monotonic() - _ai_budget_start <= _OLLAMA_BUDGET_SECS
            and str(path).lower().endswith(".pdf")
            and (not result.get("email") or not result.get("phone"))):
        _vis = _ollama_vision_contact_scan(path)
        for _vf, _vv in _vis.items():
            if not result.get(_vf) and _vv:
                result[_vf] = _vv
                logger.info(f"Vision scan filled '{_vf}': {_vv}")

    # Strip resume-builder-template watermarks (e.g. "webuildcv.com",
    # "Powered by ...") from every text field. These templates place a bare
    # branding line as its own paragraph, which then gets swept up verbatim
    # into whichever section it happens to fall in. The two fragments
    # ("webuildcv.com" and "Powered by") sometimes appear combined on a
    # single line (e.g. table-cell layout joins them with a tab), so match
    # any whole line made up of nothing but 1-3 repetitions of either
    # fragment rather than requiring each to be alone on its own line.
    for _wm_field in ("summary", "skills", "experience", "education",
                      "certifications", "projects", "references"):
        _wm_text = result.get(_wm_field)
        if _wm_text:
            _wm_lines = [l for l in _wm_text.splitlines() if not _WATERMARK_LINE_RE.match(l.strip())]
            result[_wm_field] = "\n".join(_wm_lines).strip()

    result["experience"], result["projects"] = _reassign_project_specific_bullets(
        result.get("experience", ""), result.get("projects", "")
    )

    # If this PDF has a real 'Project Details | Role and Responsibilities' table,
    # the linear text extraction above scrambles it once a cell wraps across
    # several lines (see _pdfplumber_project_table_text docstring), and that
    # scrambled duplicate content can also leak into Experience via the
    # "Responsibilities" heading match and get shuffled back into Projects by
    # the reassignment step just above. Runs last so the clean, row-paired
    # reconstruction always has the final say — no-op (returns "") for every
    # other resume layout.
    if _path_ext == "pdf":
        _table_projects = _pdfplumber_project_table_text(path)
        if _table_projects:
            result["projects"] = _table_projects
            logger.info(f"Projects overridden from pdfplumber table extraction ({len(_table_projects)} chars)")

    # Remove a contact/skills sidebar block that a 2-column PDF interleaved into
    # Work Experience (Contact/Phone/Email/LinkedIn + a Skills list), and restore
    # any skills found inside it to the Skills field.
    if result.get("experience"):
        _clean_exp, _recovered_skills = _extract_sidebar_bleed_from_experience(result["experience"])
        if _recovered_skills or _clean_exp != result["experience"]:
            result["experience"] = _clean_exp
            if _recovered_skills:
                result["skills"] = _merge_recovered_skills(result.get("skills", ""), _recovered_skills)
                logger.info(
                    "Removed interleaved contact/skills sidebar from experience; "
                    "restored %d skill item(s) to Skills",
                    len(_recovered_skills),
                )
            else:
                logger.info("Removed interleaved contact sidebar block from experience")

    # Strip wrapped-heading remnants ("ME" from "ABOUT ME", "COMPELTED" from
    # "PROJECTS COMPELTED") that leaked in as a section's first line.
    for _f in ("summary", "projects", "certifications"):
        if result.get(_f):
            result[_f] = _strip_heading_tail_prefix(result[_f])

    # Recover a full-width summary paragraph that the linear extractor split across
    # a 2-column boundary (only when the block version is clearly more complete).
    if _path_ext == "pdf":
        _blk_summary = _pymupdf_summary_block(path)
        _cur_summary = str(result.get("summary") or "")
        if _blk_summary and len(_blk_summary.strip()) > len(_cur_summary.strip()) + 20:
            result["summary"] = _blk_summary.strip()
            logger.info(
                "Summary replaced with pymupdf block paragraph (%d → %d chars)",
                len(_cur_summary), len(_blk_summary),
            )

    return result, "llm_text"


def _parse_pdf_quick(path):
    """Fast rule-based PDF parse (no AI). Returns the parsed-fields dict.

    Used by bulk compare, where running the AI model on every file would be slow.
    """
    font_name = _extract_name_from_pdf_fonts(path)
    text = extract_resume_text(path, "pdf")
    return parse_resume_text(text, name_hint=font_name)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def home():
    return redirect(url_for("edit_resume"))


@app.route("/edit", methods=["GET", "POST"])
@app.route("/edit/<int:resume_id>", methods=["GET", "POST"])
def edit_resume(resume_id=None):
    perm_key = "view_add_profile" if resume_id is None else "write_profile"
    if not _require_permission(perm_key):
        abort(403)
    with db_conn() as conn:
        if request.method == "POST":
            form_data = {k: request.form.get(k, "").strip() for k in [
                "full_name", "title", "email", "phone", "linkedin", "location", "summary",
                "skills", "experience", "education", "certifications", "projects", "exp_yrs",
                "department",
            ]}

            if form_data.get("title"):
                _tc = re.split(r"[,|/\\–—]", form_data["title"])[0].strip()
                _tc = re.sub(r"[\(\[].*$", "", _tc).strip()
                _tc = re.sub(
                    r"^(experienced|skilled|dedicated|results.driven|dynamic|seasoned|"
                    r"highly experienced|passionate|motivated|proactive|hands.on)\s+",
                    "", _tc, flags=re.I,
                ).strip()
                _words = _tc.split()
                if len(_words) > 5:
                    _m = re.match(r"^((?:[A-Z][a-zA-Z]*(?:\s+|$)){1,5})", _tc)
                    _tc = _m.group(1).strip() if _m else " ".join(_words[:5])
                form_data["title"] = _tc[:80]

            resume_file = None
            if resume_id:
                existing = conn.execute(
                    "SELECT resume_file FROM resume WHERE id = %s", (resume_id,)
                ).fetchone()
                resume_file = existing["resume_file"] if existing else None

            uploaded = request.files.get("resume_file")
            parsed_data = {}

            if uploaded and uploaded.filename:
                if not allowed_file(uploaded.filename):
                    flash("Please upload a PDF or DOCX file only.", "error")
                    return redirect(
                        url_for("edit_resume", resume_id=resume_id) if resume_id
                        else url_for("edit_resume")
                    )

                name = secure_filename(uploaded.filename)
                ext = name.rsplit(".", 1)[1].lower()
                save_name = (
                    f"{slugify(form_data['full_name'] or Path(name).stem)}"
                    f"-{int(__import__('time').time())}.{ext}"
                )
                path = UPLOAD_FOLDER / save_name
                uploaded.seek(0)
                uploaded.save(path)
                resume_file = save_name

                # The Resume Intelligence (AI) parse — OCR + LLM field
                # extraction — already ran once against this exact file via
                # the /api/parse-resume AJAX preview the moment the user
                # picked it (see edit.html's file-input "change" handler),
                # and its results are already sitting in form_data below,
                # reviewed by the user before they clicked Save. Re-parsing
                # here, synchronously inside the Save request, was pure
                # duplicate work: first confirmed live (2026-08-03) running
                # the full OCR+LLM pipeline a second time (~2.5 min/Save);
                # after removing that, still re-running just the fast, non-
                # AI parse unconditionally on every file-attached Save cost
                # another ~3.5s (its own OCR contact-strip pass) even though
                # merge_resume_data(..., overwrite=False) below throws away
                # nearly all of that second parse's output whenever the
                # preview already filled the form. So only parse again when
                # the form actually looks unfilled (preview genuinely never
                # ran — JS disabled, network error) — the one case a second
                # parse is doing real, non-duplicate work.
                _preview_already_ran = bool(
                    form_data.get("full_name") and form_data.get("summary")
                    and (form_data.get("skills") or form_data.get("experience"))
                )
                if _preview_already_ran:
                    flash("Resume file saved.", "success")
                else:
                    try:
                        if ext == "pdf":
                            parsed_data = _parse_pdf_quick(path)
                        else:
                            extracted = extract_resume_text(path, ext)
                            parsed_data = parse_resume_text(extracted)
                        flash("Resume file saved.", "success")
                    except Exception:
                        flash(
                            "File uploaded, but text could not be extracted. "
                            "You can still edit the profile manually.",
                            "warning",
                        )

            merged = merge_resume_data(form_data, parsed_data, overwrite=False)
            merged["resume_file"] = resume_file

            if resume_id:
                merged["slug"] = unique_slug(
                    conn, merged.get("full_name") or f"profile-{resume_id}", resume_id
                )
                merged["id"] = resume_id

                # Captured before the UPDATE purely so the log line below can
                # show old -> new — not used for any matching decision.
                _prev_row = conn.execute(
                    "SELECT title FROM resume WHERE id = %s", (resume_id,)
                ).fetchone()
                _prev_title = _prev_row["title"] if _prev_row else None

                conn.execute(
                    """
                    UPDATE resume SET
                    full_name=%(full_name)s,
                    title=%(title)s,
                    email=%(email)s,
                    phone=%(phone)s,
                    linkedin=%(linkedin)s,
                    location=%(location)s,
                    summary=%(summary)s,
                    experience=%(experience)s,
                    education=%(education)s,
                    certifications=%(certifications)s,
                    projects=%(projects)s,
                    slug=%(slug)s,
                    resume_file=%(resume_file)s,
                    exp_yrs=%(exp_yrs)s,
                    department=%(department)s,
                    updated_at=NOW()
                    WHERE id=%(id)s
                    """,
                    merged,
                )

                sync_skills(conn, resume_id, merged.get("skills", ""))

                try:
                    jds_for_ranking = conn.execute(
                        "SELECT * FROM job_description WHERE position_status = %s",
                        (DEFAULT_POSITION_STATUS,),
                    ).fetchall()
                    # Role & Group ("title") drives _quick_estimate_jd_match's
                    # role-match bonus, so re-running this here on every Save
                    # re-evaluates candidate JD selection against whatever the
                    # title is NOW — never a stale, pre-edit selection. Logged
                    # explicitly so a Role/Group edit's effect on matching is
                    # verifiable without a debugger: which title was in effect,
                    # and exactly which JDs the re-run selected as top-3.
                    if _prev_title != merged.get("title"):
                        logger.info(
                            f"Role/Group changed for resume={resume_id}: "
                            f"'{_prev_title}' -> '{merged.get('title')}' — "
                            f"re-running matching engine against {len(jds_for_ranking)} open JD(s)"
                        )
                    top_jd_ids = _top_n_jd_ids_by_estimate(merged, jds_for_ranking, n=3)
                    logger.info(
                        f"Matching engine re-run for resume={resume_id} (title='{merged.get('title')}') "
                        f"-> new top-3 candidate JD ids={top_jd_ids}"
                    )
                    trigger_background_ai_assessment(resume_id, top_jd_ids)
                except Exception:
                    logger.exception("Failed to trigger background AI assessment")

                log_audit("Resume Profiles", "Edit", record_id=resume_id, record_label=merged.get("full_name"))
                flash("Profile updated successfully!", "success")

                return redirect(
                    url_for("profile_detail", resume_id=resume_id)
                )

            # ── Mandatory field validation (new profiles only) ───────────
            if not merged.get("exp_yrs", "").strip():
                flash("Years of Experience is required. Please select a value before saving.", "error")
                return redirect(url_for("edit_resume"))
            if not merged.get("location", "").strip():
                flash("Location is required. Please enter a location before saving.", "error")
                return redirect(url_for("edit_resume"))
            # ─────────────────────────────────────────────────────────────

            # ── Duplicate guard (new profiles only) ──────────────────────
            _dup_email = merged.get("email", "").strip()
            _dup_phone = re.sub(r"[\s\-\(\)]", "", merged.get("phone", "").strip())
            _dup_conds, _dup_params = [], []
            if _dup_email:
                _dup_conds.append("LOWER(TRIM(email)) = LOWER(TRIM(%s))")
                _dup_params.append(_dup_email)
            if _dup_phone:
                _dup_conds.append("REGEXP_REPLACE(phone, '[\\s\\-\\(\\)]', '', 'g') = %s")
                _dup_params.append(_dup_phone)
            if _dup_conds:
                _dup_row = conn.execute(
                    "SELECT id, full_name FROM resume WHERE (" + " OR ".join(_dup_conds) + ") LIMIT 1",
                    _dup_params,
                ).fetchone()
                if _dup_row:
                    flash(
                        f"A profile for \"{_dup_row['full_name']}\" already exists with the same "
                        f"email or phone number (Profile ID: {_dup_row['id']}). "
                        f"Please check before adding again.",
                        "error",
                    )
                    return redirect(url_for("edit_resume"))
            # ─────────────────────────────────────────────────────────────

            merged["slug"] = unique_slug(conn, merged.get("full_name") or "profile")
            cursor = conn.execute(
                """
                INSERT INTO resume (
                    full_name, title, email, phone, linkedin, location, summary, skills,
                    experience, education, certifications, projects, slug, resume_file,
                    exp_yrs, department, created_at, updated_at
                ) VALUES (
                    %(full_name)s, %(title)s, %(email)s, %(phone)s, %(linkedin)s, %(location)s, %(summary)s, %(skills)s,
                    %(experience)s, %(education)s, %(certifications)s, %(projects)s, %(slug)s, %(resume_file)s,
                    %(exp_yrs)s, %(department)s, NOW(), NOW()
                )
                RETURNING id
                """,
                merged,
            )

            new_id = cursor.fetchone()["id"]

            sync_skills(conn, new_id, merged.get("skills", ""))

            try:
                jds_for_ranking = conn.execute(
                    "SELECT * FROM job_description WHERE position_status = %s",
                    (DEFAULT_POSITION_STATUS,),
                ).fetchall()
                top_jd_ids = _top_n_jd_ids_by_estimate(merged, jds_for_ranking, n=3)
                trigger_background_ai_assessment(new_id, top_jd_ids)
            except Exception:
                logger.exception("Failed to trigger background AI assessment")

            log_audit("Resume Profiles", "Add", record_id=new_id, record_label=merged.get("full_name"))
            flash(
                "Profile saved successfully! Here are your top matching roles.",
                "success"
            )

            return redirect(
                url_for("profile_detail", resume_id=new_id)
            )

        # GET
        if resume_id:
            resume = conn.execute(
                "SELECT * FROM resume WHERE id = %s",
                (resume_id,)
            ).fetchone()

            if not resume:
                return "Resume not found", 404

            resume = dict(resume)

        else:
            resume = {
                k: "" for k in [
                    "full_name",
                    "title",
                    "email",
                    "phone",
                    "linkedin",
                    "location",
                    "summary",
                    "skills",
                    "experience",
                    "education",
                    "certifications",
                    "projects",
                    "slug",
                    "resume_file",
                    "department",
                ]
            }

            resume["id"] = None

    # The "Bulk Upload" tab only applies to the Add Profile (new) flow, not editing.
    raw_files, jds = ([], []) if resume_id else _get_raw_files_and_jds()
    return render_template("edit.html", resume=resume, raw_files=raw_files, jds=jds)


@app.route("/profile")
def view_current_profile():
    return redirect(url_for("profile_list"))


@app.route("/profiles")
def profile_list():
    if not _require_permission("view_profiles", write=False):
        abort(403)
    with db_conn() as conn:
        resumes = conn.execute(
            """
            SELECT id, full_name, title, email, phone, location, resume_file, created_at, updated_at
            FROM resume
            ORDER BY updated_at DESC, id DESC
            """
        ).fetchall()
    return render_template("profile_list.html", resumes=resumes)


def _quick_estimate_jd_match(resume_dict, jd_dict):
    """Fast, LLM-free estimate of resume/JD fit — keyword checklist (70%) +
    whole-document holistic similarity (20%) + a role-title exact-match bonus
    (10%). Used ONLY to pick which JDs are worth showing as the "top 3"
    candidates without running the real AI judge against every open JD (which
    could be dozens and would make the profile page take minutes to load).
    This is never the number actually displayed for a JD that has already
    been AI-assessed — see _rank_top_jd_matches, which prefers the cached
    ai_match_cache result whenever one exists.
    """
    score = calculate_match_score(resume_dict, jd_dict)
    holistic = _holistic_or_default(resume_dict, jd_dict)
    estimate_pct = round(
        (score['match_percentage'] * 0.70) +
        (holistic['fit_percentage'] * 0.20) +
        (
            10 if _jd_role_matches(
                resume_dict.get('title', ''),
                jd_dict.get('role', '')
            ) else 0
        )
    )
    return estimate_pct, score


def _rank_top_jd_matches(conn, resume_dict, jds):
    """Single source of truth for "top matching roles" ranking/labelling,
    used by both the profile page's Top 3 cards and the Export PDF.

    Candidate selection always uses the fast, LLM-free estimate above —
    running the real AI judge against every open JD just to pick 3 would make
    this page take minutes. But the percentage/tier/strengths actually
    DISPLAYED for whichever JDs land in the top 3 come from the exact same
    ai_match_cache entry the AI Holistic Assessment on the compare page reads
    (_read_cached_hybrid_match) whenever that pair has already been
    AI-assessed — never a second, independently-computed number. A JD with no
    cached AI assessment yet shows the estimate with `is_ai_judged: False`, so
    the template can flag it as provisional instead of presenting it as
    equivalent to an AI-verified score; the profile page's own script then
    upgrades it in place via the same /api/compare/.../ai-assessment endpoint,
    converging to the identical figure the compare page would show.

    Returns matches sorted best-first; caller slices to however many it needs.
    """
    candidates = []
    for jd in jds:
        jd_dict = dict(jd)
        estimate_pct, score = _quick_estimate_jd_match(resume_dict, jd_dict)
        candidates.append((estimate_pct, jd_dict, score))
    candidates.sort(key=lambda c: -c[0])

    matches = []
    uncached_jd_ids = []
    for estimate_pct, jd_dict, score in candidates[:3]:
        cached = _read_cached_hybrid_match(conn, resume_dict, jd_dict) if conn else None
        if cached is not None:
            final_score = cached['fit_percentage']
            tier_label = cached['tier_label']
            recommendation = cached['recommendation']
            strengths = cached.get('strengths') or score.get('matched_skills', [])
            is_ai_judged = True
            authenticity_status = cached.get('authenticity_status')
            authenticity_explanation = cached.get('authenticity_explanation') or ''
            jd_mirroring_label = cached.get('jd_mirroring_label')
            jd_mirroring_phrases = cached.get('jd_mirroring_phrases') or []
            logger.info(
                f"Top-3 card resume={resume_dict.get('id')} jd={jd_dict['id']}: "
                f"AI-judged pct={final_score} tier={tier_label} authenticity={authenticity_status}"
            )
        else:
            final_score = estimate_pct
            _, tier_label, recommendation = _tier_from_pct(final_score)
            strengths = score.get('matched_skills', [])
            is_ai_judged = False
            authenticity_status = None
            authenticity_explanation = ''
            # No LLM needed for this — cheap enough to compute for the
            # estimate-only display too, not just after AI verification.
            _mirroring = _jd_mirroring_risk(resume_dict, jd_dict)
            jd_mirroring_label = _mirroring["risk_label"]
            jd_mirroring_phrases = _mirroring["matched_phrases"]
            uncached_jd_ids.append(jd_dict['id'])
            logger.info(
                f"Top-3 card resume={resume_dict.get('id')} jd={jd_dict['id']}: "
                f"estimate pct={final_score} tier={tier_label} (not yet AI-assessed)"
            )

        matches.append({
            **score,
            'jd': jd_dict,
            'jd_id': jd_dict['id'],
            'jd_title': str(jd_dict.get('title') or 'Unknown'),
            'jd_role': jd_dict.get('role', ''),
            'jd_category': str(jd_dict.get('category') or ''),

            'match_percentage': score['match_percentage'],
            'final_score': final_score,

            'tier_label': tier_label,
            'verdict': tier_label,
            'recommendation': recommendation,
            'matched_skills': strengths,
            'is_ai_judged': is_ai_judged,
            'authenticity_status': authenticity_status,
            'authenticity_explanation': authenticity_explanation,
            'jd_mirroring_label': jd_mirroring_label,
            'jd_mirroring_phrases': jd_mirroring_phrases,
        })

    if uncached_jd_ids:
        # One call → one background thread → processes these sequentially,
        # not one thread per JD — concurrent local-LLM calls contend for the
        # same CPU and end up slower in aggregate than running them one at a
        # time (same reasoning as the old sequential frontend loadNext()).
        trigger_background_ai_assessment(resume_dict.get('id'), uncached_jd_ids)

    matches.sort(key=lambda x: -x['final_score'])
    return matches


@app.route("/profile/<int:resume_id>")
def profile_detail(resume_id):
    if not _require_permission("view_candidate", write=False):
        abort(403)
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not resume:
            return "Profile not found", 404

        # Calculate top 3 matching JDs
        # Resume matching only ever considers Open positions — Fulfilled/On
        # Hold/Closed JDs are excluded before scoring, not after, so they can
        # never occupy one of the top-3 slots.
        jds = conn.execute(
            "SELECT * FROM job_description WHERE position_status = %s ORDER BY created_at DESC",
            (DEFAULT_POSITION_STATUS,),
        ).fetchall()
        resume_dict = dict(resume)

        top_matches = _rank_top_jd_matches(conn, resume_dict, jds)

        return render_template("profile.html", resume=resume, top_matches=top_matches,
                               l1_comments=resume_dict.get("l1_comments") or "",
                               l2_comments=resume_dict.get("l2_comments") or "")


@app.route("/profile/<int:resume_id>/export-top-matches")
def export_top_matches(resume_id):
    """Export top 3 matching JDs with resume comparison as JSON."""
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not resume:
            return jsonify({"error": "Profile not found"}), 404

        # Calculate top 3 matching JDs
        # Resume matching only ever considers Open positions — Fulfilled/On
        # Hold/Closed JDs are excluded before scoring, not after, so they can
        # never occupy one of the top-3 slots.
        jds = conn.execute(
            "SELECT * FROM job_description WHERE position_status = %s ORDER BY created_at DESC",
            (DEFAULT_POSITION_STATUS,),
        ).fetchall()
        resume_dict = dict(resume)

        matches = []
        for jd in jds:
            jd_dict = dict(jd)
            score = calculate_match_score(resume_dict, jd_dict)
            holistic = _holistic_or_default(resume_dict, jd_dict)
            matches.append({
                'jd': jd_dict,
                'jd_id': jd['id'],
                'jd_title': jd['title'],
                'match_percentage': score['match_percentage'],
                'holistic_percentage': holistic['fit_percentage'],
                'matched_skills': score.get('matched_skills', []),
                'missing_skills': score.get('missing_skills', []),
                'matched_count': score['matched_count'],
                'total_jd_requirements': score['total_jd_requirements'],
                'tier_label': holistic['tier_label'],
                'recommendation': holistic['recommendation'],
            })

        _resume_role = resume_dict.get('title', '')
        matches.sort(key=lambda x: (not _jd_role_matches(_resume_role, x['jd'].get('role', '')), -x['holistic_percentage']))
        top_matches = matches[:3]

    # Prepare export data
    export_data = {
        'resume': {
            'full_name': resume_dict.get('full_name'),
            'title': resume_dict.get('title'),
            'email': resume_dict.get('email'),
            'phone': resume_dict.get('phone'),
            'location': resume_dict.get('location'),
            'skills': resume_dict.get('skills'),
            'experience': resume_dict.get('experience'),
            'education': resume_dict.get('education')
        },
        'top_3_matching_jds': []
    }

    for match in top_matches:
        jd_data = {
            'rank': len(export_data['top_3_matching_jds']) + 1,
            'jd_id': match['jd_id'],
            'title': match['jd_title'],
            'category': match['jd'].get('category'),
            'role': match['jd'].get('role'),
            'match_percentage': match['match_percentage'],
            'fit': match['tier_label'],
            'recommendation': match['recommendation'],
            'skills_matched': f"{match['matched_count']}/{match['total_jd_requirements']}",
            'matched_skills': match.get('matched_skills', []),
            'missing_skills': match.get('missing_skills', []),
            'details': {
                'responsibilities': match['jd'].get('responsibilities'),
                'requirements': match['jd'].get('requirements'),
                'skills': match['jd'].get('skills'),
                'keywords': match['jd'].get('keywords')
            }
        }
        export_data['top_3_matching_jds'].append(jd_data)

    filename = f"{resume_dict.get('full_name', 'resume').replace(' ', '_')}_top_3_matches.json"
    return send_file(
        BytesIO(json.dumps(export_data, indent=2).encode()),
        mimetype='application/json',
        as_attachment=True,
        download_name=filename
    )


def _pdf_hr(color=None, thickness=0.5):
    """Reusable horizontal rule for PDF reports."""
    return HRFlowable(width='100%', thickness=thickness,
                      color=color or HexColor('#e2e8f0'),
                      spaceAfter=6, spaceBefore=6)


def _pdf_skill_tags(skills, bg, fg, cols=4, col_w=1.65):
    """Render a skill list as a flat colored-cell table — no nested tables."""
    if not skills:
        return None
    tag_style = ParagraphStyle('_Tag', fontSize=8, textColor=fg,
                                fontName='Helvetica-Bold', alignment=TA_CENTER,
                                leading=11)
    cells = [Paragraph(str(s)[:38], tag_style) for s in skills]
    while len(cells) % cols != 0:
        cells.append('')
    rows = [cells[i:i + cols] for i in range(0, len(cells), cols)]
    grid = Table(rows, colWidths=[col_w * inch] * cols)
    grid.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), bg),
        ('FONTNAME',      (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0), (-1, -1), 8),
        ('TEXTCOLOR',     (0, 0), (-1, -1), fg),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 6),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 6),
        ('TOPPADDING',    (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('GRID',          (0, 0), (-1, -1), 2, RL_WHITE),
    ]))
    return grid


def _build_match_detail_story(resume_dict, jd_dict, sc, rank=None):
    """
    Build a list of ReportLab flowables for one resume-vs-JD match section.
    Used by both export_top_matches_pdf and export_compare_pdf.
    """
    C_NAVY      = HexColor('#1e293b')
    C_GREEN     = HexColor('#15803d')
    C_GREEN_BG  = HexColor('#d1fae5')
    C_RED       = HexColor('#dc2626')
    C_RED_BG    = HexColor('#fee2e2')
    C_ORANGE    = HexColor('#d97706')
    C_ORANGE_BG = HexColor('#fff7ed')
    C_PURPLE    = HexColor('#7c3aed')
    C_PURPLE_BG = HexColor('#ede9fe')
    C_BLUE      = HexColor('#2563eb')
    C_BLUE_BG   = HexColor('#dbeafe')
    C_GRAY      = HexColor('#64748b')
    C_LIGHT     = HexColor('#f8fafc')
    C_BORDER    = HexColor('#e2e8f0')

    # Verdict comes from the whole-document holistic match, not a raw percentage —
    # color keyed to the verdict's own wording so nothing here is percentage-driven.
    holistic = _holistic_or_default(resume_dict, jd_dict)
    level    = holistic['verdict']
    level_l  = level.lower()
    if 'strong' in level_l:
        lv_color, lv_bg = C_GREEN, C_GREEN_BG
    elif 'good' in level_l:
        lv_color, lv_bg = C_BLUE, C_BLUE_BG
    elif 'partial' in level_l:
        lv_color, lv_bg = C_ORANGE, C_ORANGE_BG
    else:
        lv_color, lv_bg = C_RED, C_RED_BG

    s_body  = ParagraphStyle('_Body',  fontSize=9,  textColor=C_NAVY, spaceAfter=3, leading=14)
    s_muted = ParagraphStyle('_Muted', fontSize=8.5, textColor=C_GRAY, spaceAfter=3, leading=13)

    story = []

    # ── Banner: rank + title + score ─────────────────────────────────────────
    rank_label = f'#{rank}' if rank else ''
    rank_cell = Table(
        [[Paragraph(rank_label,
                    ParagraphStyle('_Rk', fontSize=22, fontName='Helvetica-Bold',
                                   textColor=RL_WHITE, alignment=TA_CENTER))]],
        colWidths=[0.65 * inch], rowHeights=[0.65 * inch]
    )
    rank_cell.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), lv_color),
        ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',(0, 0), (-1, -1), 0),
        ('RIGHTPADDING',(0,0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING',(0,0),(-1, -1), 0),
    ]))

    title_inner = [
        Paragraph(jd_dict.get('title', ''),
                  ParagraphStyle('_JDT', fontSize=13, fontName='Helvetica-Bold',
                                  textColor=C_NAVY, spaceAfter=2)),
        Paragraph(f"{jd_dict.get('category', '')}  ·  {jd_dict.get('role', '')}",
                  ParagraphStyle('_JDC', fontSize=8.5, textColor=C_GRAY)),
    ]
    title_cell = Table([title_inner], colWidths=[4.1 * inch])
    title_cell.setStyle(TableStyle([
        ('VALIGN',      (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING',  (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING',(0,0), (-1, -1), 4),
        ('BACKGROUND',  (0, 0), (-1, -1), C_LIGHT),
    ]))

    score_cell = Table([
        [Paragraph(holistic['tier_label'],
                   ParagraphStyle('_ScLvl', fontSize=14, fontName='Helvetica-Bold',
                                   textColor=lv_color, alignment=TA_CENTER, leading=17))],
        [Paragraph(holistic['recommendation'],
                   ParagraphStyle('_ScRec', fontSize=8.5, fontName='Helvetica-Bold',
                                   textColor=lv_color, alignment=TA_CENTER, leading=11))],
    ], colWidths=[2.1 * inch])
    score_cell.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, -1), lv_bg),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING',   (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
        ('LEFTPADDING',  (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))

    banner = Table([[rank_cell, title_cell, score_cell]],
                    colWidths=[0.65 * inch, 4.1 * inch, 2.1 * inch],
                    rowHeights=[0.65 * inch])
    banner.setStyle(TableStyle([
        ('BOX',           (0, 0), (-1, -1), 1, lv_color),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(banner)
    story.append(Spacer(1, 0.14 * inch))

    # ── Match Summary table (qualitative — no percentage scores) ─────────────
    story.append(Paragraph(
        'Match Summary',
        ParagraphStyle('_SecH', fontSize=10, fontName='Helvetica-Bold',
                        textColor=C_NAVY, spaceAfter=5, spaceBefore=8)
    ))
    exp_note_short = (sc.get('experience_note') or '')[:90]
    sb_data = [
        ['Metric', 'Detail'],
        ['Overall Fit',       level],
        ['Requirements Met',  f"{sc.get('matched_count',0)} of {sc.get('total_jd_requirements',0)}"],
        ['Experience',        exp_note_short or '—'],
    ]
    sb_tbl = Table(sb_data, colWidths=[1.55 * inch, 5.75 * inch])
    sb_tbl.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, 0),  C_NAVY),
        ('TEXTCOLOR',    (0, 0), (-1, 0),  RL_WHITE),
        ('FONTNAME',     (0, 0), (-1, 0),  'Helvetica-Bold'),
        ('FONTSIZE',     (0, 0), (-1, -1), 8.5),
        ('ROWBACKGROUNDS',(0,1), (-1, -1), [C_LIGHT, RL_WHITE]),
        ('GRID',         (0, 0), (-1, -1), 0.5, C_BORDER),
        ('FONTNAME',     (1, 1), (1,  1),  'Helvetica-Bold'),
        ('TEXTCOLOR',    (1, 1), (1,  1),  lv_color),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
        ('LEFTPADDING',  (0, 0), (-1, -1), 8),
    ]))
    story.append(sb_tbl)
    story.append(Spacer(1, 0.12 * inch))

    # ── Matched / Missing skills — two-column ─────────────────────────────────
    matched  = sc.get('matched_skills', [])
    missing  = sc.get('missing_skills', [])

    def _skill_list_cell(skills, tick, color, bg):
        rows = [Paragraph(f'<b>{tick}  {s}</b>',
                           ParagraphStyle('_SI', fontSize=8.5, textColor=color,
                                           spaceAfter=2, leading=13))
                for s in skills]
        if not rows:
            rows = [Paragraph('None' if tick == '✗' else 'All matched!',
                               ParagraphStyle('_SIN', fontSize=8.5, textColor=color))]
        return rows

    left_hdr   = Paragraph(
        f'Matched Skills ({sc.get("matched_count", 0)})',
        ParagraphStyle('_MH', fontSize=10, fontName='Helvetica-Bold',
                        textColor=C_GREEN, spaceAfter=5))
    right_hdr  = Paragraph(
        f'Missing / Gap Skills ({sc.get("missing_count", 0)})',
        ParagraphStyle('_GH', fontSize=10, fontName='Helvetica-Bold',
                        textColor=C_RED, spaceAfter=5))

    left_items  = _skill_list_cell(matched, '✓', C_GREEN, C_GREEN_BG)
    right_items = _skill_list_cell(missing, '✗', C_RED,   C_RED_BG)

    skills_tbl = Table(
        [[[left_hdr]  + left_items,
          [right_hdr] + right_items]],
        colWidths=[3.35 * inch, 3.35 * inch]
    )
    skills_tbl.setStyle(TableStyle([
        ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
        ('BACKGROUND',   (0, 0), (0,  -1), HexColor('#f0fdf4')),
        ('BACKGROUND',   (1, 0), (1,  -1), HexColor('#fff5f5')),
        ('BOX',          (0, 0), (0,  -1), 0.5, C_GREEN),
        ('BOX',          (1, 0), (1,  -1), 0.5, C_RED),
        ('LEFTPADDING',  (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING',   (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 10),
    ]))
    story.append(skills_tbl)
    story.append(Spacer(1, 0.1 * inch))

    # ── Strong Areas ─────────────────────────────────────────────────────────
    strong = sc.get('strong_areas', [])
    if strong:
        story.append(Paragraph(
            'Strong Areas',
            ParagraphStyle('_SAH', fontSize=10, fontName='Helvetica-Bold',
                            textColor=C_PURPLE, spaceAfter=5, spaceBefore=6)
        ))
        tags = _pdf_skill_tags(strong, C_PURPLE_BG, C_PURPLE)
        if tags:
            story.append(tags)
        story.append(Spacer(1, 0.06 * inch))

    # ── Skill Gaps ───────────────────────────────────────────────────────────
    weak = sc.get('weak_areas', [])
    if weak:
        story.append(Paragraph(
            'Skill Gaps to Address',
            ParagraphStyle('_WAH', fontSize=10, fontName='Helvetica-Bold',
                            textColor=C_ORANGE, spaceAfter=5, spaceBefore=6)
        ))
        tags = _pdf_skill_tags(weak, C_ORANGE_BG, C_ORANGE)
        if tags:
            story.append(tags)
        story.append(Spacer(1, 0.06 * inch))

    # ── Experience Assessment ─────────────────────────────────────────────────
    exp_note  = sc.get('experience_note', '')
    jd_yrs    = sc.get('jd_years_required')
    res_yrs   = sc.get('resume_years_estimated')
    if exp_note or jd_yrs or res_yrs:
        story.append(_pdf_hr())
        story.append(Paragraph(
            'Experience Assessment',
            ParagraphStyle('_ExpH', fontSize=10, fontName='Helvetica-Bold',
                            textColor=C_NAVY, spaceAfter=5, spaceBefore=4)
        ))
        if exp_note:
            story.append(Paragraph(exp_note, s_body))
        exp_rows = []
        if jd_yrs:
            exp_rows.append(['JD Requires', f'{jd_yrs}+ years'])
        if res_yrs is not None:
            exp_rows.append(['Resume Shows', f'~{res_yrs} year(s)'])
        if exp_rows:
            exp_tbl = Table(exp_rows, colWidths=[1.5 * inch, 2.2 * inch])
            exp_tbl.setStyle(TableStyle([
                ('FONTSIZE',     (0, 0), (-1, -1), 9),
                ('FONTNAME',     (0, 0), (0,  -1), 'Helvetica-Bold'),
                ('TEXTCOLOR',    (0, 0), (0,  -1), C_NAVY),
                ('TEXTCOLOR',    (1, 0), (1,  -1), lv_color),
                ('FONTNAME',     (1, 0), (1,  -1), 'Helvetica-Bold'),
                ('TOPPADDING',   (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
                ('LEFTPADDING',  (0, 0), (-1, -1), 6),
            ]))
            story.append(exp_tbl)

    # ── Recommendations ───────────────────────────────────────────────────────
    story.append(_pdf_hr())
    story.append(Paragraph(
        'Recommendations',
        ParagraphStyle('_RecH', fontSize=10, fontName='Helvetica-Bold',
                        textColor=C_NAVY, spaceAfter=5, spaceBefore=4)
    ))
    recs = [holistic['recommendation']]
    if missing:
        top_miss = missing[:5]
        recs.append(f'Key areas to strengthen: {", ".join(top_miss)}'
                    f'{"..." if len(missing) > 5 else ""}.')
    if strong:
        recs.append(f'Candidate demonstrates strong background in: {", ".join(strong[:4])}.')
    for rec in recs:
        story.append(Paragraph(f'•  {rec}', s_body))

    return story


@app.route("/profile/<int:resume_id>/export-top-matches-pdf")
def export_top_matches_pdf(resume_id):
    """Export top 3 matching JDs as a professional candidate assessment PDF."""
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not resume:
            return "Profile not found", 404
        # Resume matching only ever considers Open positions — Fulfilled/On
        # Hold/Closed JDs are excluded before scoring, not after, so they can
        # never occupy one of the top-3 slots.
        jds = conn.execute(
            "SELECT * FROM job_description WHERE position_status = %s ORDER BY created_at DESC",
            (DEFAULT_POSITION_STATUS,),
        ).fetchall()

    resume_dict = dict(resume)

    matches = []
    for jd in jds:
        jd_dict = dict(jd)
        score = calculate_match_score(resume_dict, jd_dict)
        holistic = _holistic_or_default(resume_dict, jd_dict)
        matches.append({'jd': jd_dict, 'score': score,
                         'holistic_percentage': holistic['fit_percentage'],
                         'verdict': holistic['verdict'],
                         'tier_label': holistic['tier_label'],
                         'recommendation': holistic['recommendation']})
    _resume_role = resume_dict.get('title', '')
    matches.sort(key=lambda x: (not _jd_role_matches(_resume_role, x['jd'].get('role', '')), -x['holistic_percentage']))
    top_matches = matches[:3]

    # ── Shared palette / styles ────────────────────────────────────────────────
    C_NAVY   = HexColor('#1e293b')
    C_BLUE   = HexColor('#2563eb')
    C_GRAY   = HexColor('#64748b')
    C_LIGHT  = HexColor('#f8fafc')
    C_BORDER = HexColor('#e2e8f0')
    C_GREEN  = HexColor('#15803d')
    C_ORANGE = HexColor('#d97706')
    C_RED    = HexColor('#dc2626')
    def _verdict_col(verdict):
        v = (verdict or '').lower()
        if 'strong' in v: return C_GREEN
        if 'good' in v: return C_BLUE
        if 'partial' in v: return C_ORANGE
        return C_RED

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        topMargin=0.55 * inch, bottomMargin=0.55 * inch,
        leftMargin=0.75 * inch, rightMargin=0.75 * inch
    )
    story = []

    # ── Cover: report header strip ─────────────────────────────────────────────
    hdr_tbl = Table(
        [[Paragraph('CANDIDATE ASSESSMENT REPORT',
                    ParagraphStyle('_HDR', fontSize=9, textColor=RL_WHITE,
                                   fontName='Helvetica-Bold', alignment=TA_CENTER))]],
        colWidths=[7.0 * inch], rowHeights=[24]
    )
    hdr_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), C_NAVY),
        ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 0.2 * inch))

    # Candidate name + title + contact — explicit rowHeights guarantee no overlap
    _name_rows   = []
    _row_heights = []
    _name_rows.append([Paragraph(
        resume_dict.get('full_name', ''),
        ParagraphStyle('_CName', fontSize=24, fontName='Helvetica-Bold',
                        textColor=C_NAVY, alignment=TA_CENTER, leading=30))])
    _row_heights.append(0.50 * inch)   # 36 pt — fits 24 pt font centred with room
    if resume_dict.get('title'):
        _name_rows.append([Paragraph(
            resume_dict['title'],
            ParagraphStyle('_CRole', fontSize=11, textColor=C_GRAY,
                            alignment=TA_CENTER, leading=16))])
        _row_heights.append(0.30 * inch)   # 21.6 pt — fits 11 pt font
    # Contact: split each field on | to de-duplicate combined parser output
    _seen_cp = set()
    _contact_parts = []
    for _f in [resume_dict.get('email'), resume_dict.get('phone'),
                resume_dict.get('location')]:
        for _tok in (_f or '').replace('|', '\n').split('\n'):
            _tok = _tok.strip()
            if _tok and _tok not in _seen_cp:
                _seen_cp.add(_tok)
                _contact_parts.append(_tok)
    _contact_parts = _contact_parts[:4]
    if _contact_parts:
        _name_rows.append([Paragraph(
            '  |  '.join(_contact_parts),
            ParagraphStyle('_CCont', fontSize=9, textColor=C_GRAY,
                            alignment=TA_CENTER, leading=12))])
        _row_heights.append(0.26 * inch)   # 18.7 pt — fits 9 pt font
    _name_tbl = Table(_name_rows, colWidths=[7.0 * inch], rowHeights=_row_heights)
    _name_tbl.setStyle(TableStyle([
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
    ]))
    story.append(_name_tbl)

    story.append(Spacer(1, 0.2 * inch))
    story.append(_pdf_hr(C_BLUE, 1.2))
    story.append(Spacer(1, 0.08 * inch))

    # Overview summary table
    story.append(Paragraph(
        'TOP 3 MATCHING ROLES — OVERVIEW',
        ParagraphStyle('_OvH', fontSize=10, fontName='Helvetica-Bold',
                        textColor=C_BLUE, spaceAfter=7)
    ))
    ov_data = [['Rank', 'Job Role', 'Category', 'Overall Fit', 'Recommendation', 'Matched', 'Missing']]
    for i, m in enumerate(top_matches, 1):
        sc = m['score']
        ov_data.append([
            f'#{i}',
            m['jd'].get('title', ''),
            m['jd'].get('category', ''),
            m['tier_label'],
            m['recommendation'],
            str(sc['matched_count']),
            str(sc['missing_count']),
        ])
    ov_tbl = Table(ov_data,
                    colWidths=[0.4*inch, 1.75*inch, 1.0*inch, 0.95*inch, 1.5*inch, 0.65*inch, 0.7*inch])
    ov_style = TableStyle([
        ('BACKGROUND',    (0, 0), (-1,  0), C_NAVY),
        ('TEXTCOLOR',     (0, 0), (-1,  0), RL_WHITE),
        ('FONTNAME',      (0, 0), (-1,  0), 'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0), (-1, -1), 8.5),
        ('ROWBACKGROUNDS',(0, 1), (-1, -1), [C_LIGHT, RL_WHITE]),
        ('GRID',          (0, 0), (-1, -1), 0.5, C_BORDER),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN',         (1, 0), (2,  -1), 'LEFT'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING',    (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING',   (0, 0), (-1, -1), 6),
    ])
    for i, m in enumerate(top_matches, 1):
        col = _verdict_col(m['verdict'])
        ov_style.add('TEXTCOLOR',  (3, i), (4, i), col)
        ov_style.add('FONTNAME',   (3, i), (4, i), 'Helvetica-Bold')
    ov_tbl.setStyle(ov_style)
    story.append(ov_tbl)

    # Candidate skills snapshot
    if resume_dict.get('skills'):
        story.append(Spacer(1, 0.12 * inch))
        story.append(_pdf_hr())
        story.append(Paragraph(
            'Candidate Key Skills',
            ParagraphStyle('_CSH', fontSize=10, fontName='Helvetica-Bold',
                            textColor=C_NAVY, leading=14, spaceAfter=0)
        ))
        story.append(Spacer(1, 0.1 * inch))
        raw_skills = [s.strip()
                       for s in resume_dict['skills'].replace(',', '\n').split('\n')
                       if s.strip()][:20]
        tags = _pdf_skill_tags(raw_skills, HexColor('#dbeafe'), C_BLUE, cols=4, col_w=1.65)
        if tags:
            story.append(tags)

    story.append(PageBreak())

    # ── Per-match detail pages ─────────────────────────────────────────────────
    for idx, match in enumerate(top_matches, 1):
        detail = _build_match_detail_story(resume_dict, match['jd'], match['score'], rank=idx)
        story.extend(detail)
        if idx < len(top_matches):
            story.append(PageBreak())

    # Footer
    story.append(Spacer(1, 0.25 * inch))
    story.append(_pdf_hr())
    story.append(Paragraph(
        f'Generated by Resume Profile Manager  ·  Candidate: {resume_dict.get("full_name", "")}',
        ParagraphStyle('_Ftr', fontSize=7.5, textColor=C_GRAY, alignment=TA_CENTER, leading=11)
    ))

    doc.build(story)
    buffer.seek(0)
    fname = f"{resume_dict.get('full_name', 'candidate').replace(' ', '_')}_assessment_report.pdf"
    return send_file(buffer, mimetype='application/pdf', as_attachment=True, download_name=fname)


@app.route("/compare/<int:resume_id>/<int:jd_id>/export-pdf")
def export_compare_pdf(resume_id, jd_id):
    """Export a single resume vs JD comparison as a professional assessment PDF."""
    with db_conn() as conn:
        ensure_jd_table(conn)
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        jd     = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
    if not resume or not jd:
        return "Resume or JD not found", 404

    resume_dict = dict(resume)
    jd_dict     = dict(jd)
    sc          = calculate_match_score(resume_dict, jd_dict)

    C_NAVY  = HexColor('#1e293b')
    C_BLUE  = HexColor('#2563eb')
    C_GRAY  = HexColor('#64748b')

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        topMargin=0.55 * inch, bottomMargin=0.55 * inch,
        leftMargin=0.75 * inch, rightMargin=0.75 * inch
    )
    story = []

    # Report header strip
    hdr_tbl = Table(
        [[Paragraph('CANDIDATE MATCH REPORT',
                    ParagraphStyle('_HR2', fontSize=9, textColor=RL_WHITE,
                                   fontName='Helvetica-Bold', alignment=TA_CENTER))]],
        colWidths=[7.0 * inch], rowHeights=[24]
    )
    hdr_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), C_NAVY),
        ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 0.18 * inch))

    # Candidate name + contact — explicit rowHeights guarantee no overlap
    _cn2_rows   = []
    _cn2_rh     = []
    _cn2_rows.append([Paragraph(
        resume_dict.get('full_name', ''),
        ParagraphStyle('_CN2', fontSize=22, fontName='Helvetica-Bold',
                        textColor=C_NAVY, alignment=TA_CENTER, leading=28))])
    _cn2_rh.append(0.48 * inch)   # 34.6 pt — fits 22 pt font centred
    if resume_dict.get('title'):
        _cn2_rows.append([Paragraph(
            resume_dict['title'],
            ParagraphStyle('_CR2', fontSize=10, textColor=C_GRAY,
                            alignment=TA_CENTER, leading=15))])
        _cn2_rh.append(0.28 * inch)
    _seen_cp2 = set()
    _contact2  = []
    for _f in [resume_dict.get('email'), resume_dict.get('phone'),
                resume_dict.get('location')]:
        for _tok in (_f or '').replace('|', '\n').split('\n'):
            _tok = _tok.strip()
            if _tok and _tok not in _seen_cp2:
                _seen_cp2.add(_tok)
                _contact2.append(_tok)
    _contact2 = _contact2[:4]
    if _contact2:
        _cn2_rows.append([Paragraph(
            '  |  '.join(_contact2),
            ParagraphStyle('_CC2', fontSize=8.5, textColor=C_GRAY,
                            alignment=TA_CENTER, leading=12))])
        _cn2_rh.append(0.24 * inch)
    _cn2_tbl = Table(_cn2_rows, colWidths=[7.0 * inch], rowHeights=_cn2_rh)
    _cn2_tbl.setStyle(TableStyle([
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
    ]))
    story.append(_cn2_tbl)
    story.append(Spacer(1, 0.15 * inch))
    story.append(_pdf_hr(C_BLUE, 1.2))
    story.append(Spacer(1, 0.08 * inch))

    # Compared-against label
    story.append(Paragraph(
        f'Compared Against:  <b>{jd_dict.get("title", "")}</b>'
        f'  ·  {jd_dict.get("category", "")}',
        ParagraphStyle('_CAL', fontSize=9.5, textColor=C_GRAY, spaceAfter=10)
    ))

    # Detailed match section
    story.extend(_build_match_detail_story(resume_dict, jd_dict, sc, rank=None))

    # Footer
    story.append(Spacer(1, 0.2 * inch))
    story.append(_pdf_hr())
    story.append(Paragraph(
        f'Generated by Resume Profile Manager  ·  '
        f'{resume_dict.get("full_name", "")} vs {jd_dict.get("title", "")}',
        ParagraphStyle('_Ft2', fontSize=7.5, textColor=C_GRAY, alignment=TA_CENTER)
    ))

    doc.build(story)
    buffer.seek(0)
    cname = resume_dict.get('full_name', 'candidate').replace(' ', '_')
    jname = jd_dict.get('title', 'role').replace(' ', '_')
    log_audit("Export", "Export PDF", record_id=f"{resume_id}:{jd_id}",
              record_label=f"{resume_dict.get('full_name')} vs {jd_dict.get('title')}")
    return send_file(
        buffer, mimetype='application/pdf', as_attachment=True,
        download_name=f'{cname}_vs_{jname}_report.pdf'
    )


@app.route("/profile/<int:resume_id>/export-rich-pdf")
def export_rich_profile_pdf(resume_id):
    """Export a professional assessment PDF: overview table + per-JD breakdown."""
    from reportlab.platypus import Table, TableStyle, HRFlowable, PageBreak
    from reportlab.lib import colors as rl_colors

    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not resume:
            return "Profile not found", 404
        # Resume matching only ever considers Open positions — Fulfilled/On
        # Hold/Closed JDs are excluded before scoring, not after, so they can
        # never occupy one of the top-3 slots.
        jds = conn.execute(
            "SELECT * FROM job_description WHERE position_status = %s ORDER BY created_at DESC",
            (DEFAULT_POSITION_STATUS,),
        ).fetchall()

        resume_dict = dict(resume)
        top_matches = _rank_top_jd_matches(conn, resume_dict, jds)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            topMargin=0.5*inch, bottomMargin=0.5*inch,
                            leftMargin=0.75*inch, rightMargin=0.75*inch)

    C_DARK    = rl_colors.HexColor('#1e293b')
    C_INDIGO  = rl_colors.HexColor('#4f46e5')
    C_GREEN   = rl_colors.HexColor('#10b981')
    C_ORANGE  = rl_colors.HexColor('#f59e0b')
    C_DORANG  = rl_colors.HexColor('#d97706')
    C_RED     = rl_colors.HexColor('#ef4444')
    C_MUTED   = rl_colors.HexColor('#64748b')
    C_BGBLUE  = rl_colors.HexColor('#eef2ff')
    C_BGGRN   = rl_colors.HexColor('#d1fae5')
    C_BGRED   = rl_colors.HexColor('#fee2e2')
    C_BGORG   = rl_colors.HexColor('#fef3c7')
    C_BGINDIG = rl_colors.HexColor('#ede9fe')
    C_WHITE   = rl_colors.white
    C_HDR     = rl_colors.HexColor('#1e293b')
    C_LINE    = rl_colors.HexColor('#3b82f6')

    def _scol(verdict):
        v = (verdict or '').lower()
        if 'strong' in v: return C_GREEN
        if 'good' in v: return C_INDIGO
        if 'partial' in v: return C_ORANGE
        return C_RED

    def _ps(nm, **kw):
        return ParagraphStyle(nm, **kw)

    W = A4[0] - 1.5 * inch
    story = []

    # ── PAGE 1: HEADER BAR ────────────────────────────────────────────────────
    hdr_bar = Table(
        [[Paragraph('CANDIDATE ASSESSMENT REPORT',
                    _ps('RP_HdrBar', fontName='Helvetica-Bold', fontSize=13,
                        textColor=C_WHITE, alignment=TA_CENTER))]],
        colWidths=[W]
    )
    hdr_bar.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), C_HDR),
        ('TOPPADDING',    (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('LEFTPADDING',   (0, 0), (-1, -1), 10),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 10),
    ]))
    story.append(hdr_bar)
    story.append(Spacer(1, 0.18 * inch))

    cname  = str(resume_dict.get('full_name') or 'Candidate')
    ctitle = str(resume_dict.get('title') or '')
    cemail = str(resume_dict.get('email') or '')
    cphone = str(resume_dict.get('phone') or '')
    cloc   = str(resume_dict.get('location') or '')
    contact_parts = [x for x in [cemail, cphone, cloc] if x]

    story.append(Paragraph(cname,
                            _ps('RP_Name', fontName='Helvetica-Bold', fontSize=26,
                                textColor=C_DARK, leading=32, alignment=TA_CENTER)))
    if ctitle:
        story.append(Paragraph(ctitle,
                                _ps('RP_Title', fontName='Helvetica', fontSize=13,
                                    textColor=C_MUTED, leading=18, alignment=TA_CENTER,
                                    spaceAfter=2)))
    if contact_parts:
        story.append(Paragraph(' | '.join(contact_parts),
                                _ps('RP_Contact', fontName='Helvetica', fontSize=9,
                                    textColor=C_MUTED, leading=14, alignment=TA_CENTER,
                                    spaceAfter=6)))
    story.append(Spacer(1, 0.08 * inch))
    story.append(HRFlowable(width=W, thickness=2, color=C_LINE, spaceAfter=14))

    # ── OVERVIEW TABLE ────────────────────────────────────────────────────────
    story.append(Paragraph('TOP 3 MATCHING ROLES — OVERVIEW',
                            _ps('RP_OvSec', fontName='Helvetica-Bold', fontSize=12,
                                textColor=C_INDIGO, spaceAfter=8)))

    ov_hc = _ps('RP_OvHC', fontName='Helvetica-Bold', fontSize=9,
                textColor=C_WHITE, alignment=TA_CENTER)
    ov_hl = _ps('RP_OvHL', fontName='Helvetica-Bold', fontSize=9, textColor=C_WHITE)
    ov_data = [[
        Paragraph('Rank',           ov_hc),
        Paragraph('Job Role',       ov_hl),
        Paragraph('Category',       ov_hl),
        Paragraph('Overall Fit',    ov_hc),
        Paragraph('Recommendation', ov_hc),
        Paragraph('Matched',        ov_hc),
        Paragraph('Missing',        ov_hc),
    ]]
    for ri, m in enumerate(top_matches, 1):
        verdict = m['verdict']
        sc     = _scol(verdict)
        mc     = int(m['matched_count'])
        ms_cnt = int(m['missing_count'])
        ov_data.append([
            Paragraph(f'#{ri}',
                      _ps(f'RP_OvRk{ri}', fontName='Helvetica-Bold',
                          fontSize=10, textColor=C_DARK, alignment=TA_CENTER)),
            Paragraph(str(m['jd_title']),
                      _ps(f'RP_OvT{ri}', fontName='Helvetica', fontSize=9,
                          textColor=C_DARK, leading=12)),
            Paragraph(str(m['jd_category']),
                      _ps(f'RP_OvC{ri}', fontName='Helvetica', fontSize=9,
                          textColor=C_MUTED, leading=12)),
            Paragraph(m['tier_label'],
                      _ps(f'RP_OvSc{ri}', fontName='Helvetica-Bold',
                          fontSize=9, textColor=sc, alignment=TA_CENTER)),
            Paragraph(m['recommendation'],
                      _ps(f'RP_OvRec{ri}', fontName='Helvetica-Bold',
                          fontSize=8, textColor=sc, alignment=TA_CENTER)),
            Paragraph(str(mc),
                      _ps(f'RP_OvMc{ri}', fontName='Helvetica-Bold',
                          fontSize=9, textColor=C_GREEN, alignment=TA_CENTER)),
            Paragraph(str(ms_cnt),
                      _ps(f'RP_OvMs{ri}', fontName='Helvetica-Bold',
                          fontSize=9, textColor=C_RED, alignment=TA_CENTER)),
        ])

    ov_cols  = [0.06*W, 0.22*W, 0.15*W, 0.14*W, 0.24*W, 0.09*W, 0.10*W]
    ov_style = [
        ('BACKGROUND',    (0, 0), (-1, 0),  C_DARK),
        ('TOPPADDING',    (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
        ('LEFTPADDING',   (0, 0), (-1, -1), 7),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 7),
        ('GRID',          (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#e2e8f0')),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
    ]
    for ri in range(1, len(ov_data)):
        ov_style.append(('BACKGROUND', (0, ri), (-1, ri),
                         C_WHITE if ri % 2 == 1 else rl_colors.HexColor('#f8fafc')))
    ov_tbl = Table(ov_data, colWidths=ov_cols)
    ov_tbl.setStyle(TableStyle(ov_style))
    story.append(ov_tbl)
    story.append(Spacer(1, 0.2 * inch))

    # ── CANDIDATE KEY SKILLS ──────────────────────────────────────────────────
    skills_raw = str(resume_dict.get('skills') or '')
    skill_list = [s.strip() for s in skills_raw.split('\n') if s.strip()]

    if skill_list:
        story.append(Paragraph('Candidate Key Skills',
                                _ps('RP_SkSec', fontName='Helvetica-Bold', fontSize=12,
                                    textColor=C_DARK, spaceAfter=8)))
        tag_sty   = _ps('RP_SkTag', fontName='Helvetica', fontSize=8.5,
                        textColor=C_DARK, leading=13, alignment=TA_CENTER)
        empty_tag = Paragraph('', tag_sty)
        tag_rows  = []
        row       = []
        for i, sk in enumerate(skill_list):
            row.append(Paragraph(str(sk), tag_sty))
            if len(row) == 4 or i == len(skill_list) - 1:
                while len(row) < 4:
                    row.append(empty_tag)
                tag_rows.append(row)
                row = []
        if tag_rows:
            sk_tbl = Table(tag_rows, colWidths=[W / 4] * 4)
            sk_tbl.setStyle(TableStyle([
                ('BACKGROUND',    (0, 0), (-1, -1), C_BGBLUE),
                ('TOPPADDING',    (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LEFTPADDING',   (0, 0), (-1, -1), 6),
                ('RIGHTPADDING',  (0, 0), (-1, -1), 6),
                ('GRID',          (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#c7d2fe')),
                ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(sk_tbl)

    # ── HELPER: tag grid (4-column) ───────────────────────────────────────────
    def _tag_grid(items, tag_bg, text_col, border_col, pfx):
        safe = list(items) if items else []
        if not safe:
            return None
        t_sty  = _ps(f'RP_TG{pfx}', fontName='Helvetica', fontSize=8.5,
                     textColor=text_col, leading=13, alignment=TA_CENTER)
        empty  = Paragraph('', t_sty)
        rows   = []
        row    = []
        for i, sk in enumerate(safe):
            row.append(Paragraph(str(sk), t_sty))
            if len(row) == 4 or i == len(safe) - 1:
                while len(row) < 4:
                    row.append(empty)
                rows.append(row)
                row = []
        tbl = Table(rows, colWidths=[W / 4] * 4)
        tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), tag_bg),
            ('TOPPADDING',    (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING',   (0, 0), (-1, -1), 6),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 6),
            ('GRID',          (0, 0), (-1, -1), 0.5, border_col),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        return tbl

    # ── HELPER: skills list (nested single-column Table) ─────────────────────
    def _skill_inner(items, label, label_col, text_col, marker, pfx):
        safe = list(items) if items else []
        rows = [[Paragraph(label,
                           _ps(f'RP_SL{pfx}', fontName='Helvetica-Bold', fontSize=9,
                               textColor=label_col, leading=13, spaceAfter=2))]]
        if safe:
            item_sty = _ps(f'RP_SI{pfx}', fontName='Helvetica', fontSize=8.5,
                           textColor=text_col, leading=13, leftIndent=4)
            for sk in safe[:20]:
                rows.append([Paragraph(f'{marker} {str(sk)}', item_sty)])
            if len(safe) > 20:
                rows.append([Paragraph(
                    f'+ {len(safe) - 20} more',
                    _ps(f'RP_SM{pfx}', fontName='Helvetica', fontSize=8,
                        textColor=C_MUTED, leading=11, leftIndent=4)
                )])
        else:
            rows.append([Paragraph('None',
                                   _ps(f'RP_SN{pfx}', fontName='Helvetica', fontSize=8.5,
                                       textColor=C_MUTED, leading=12, leftIndent=4))])
        inner = Table(rows, colWidths=[W * 0.5 - 28])
        inner.setStyle(TableStyle([
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING',   (0, 0), (-1, -1), 0),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ]))
        return inner

    # ── PER-JD DETAILED SECTIONS ──────────────────────────────────────────────
    RANKS = ['#1', '#2', '#3']

    for idx, m in enumerate(top_matches):
        verdict     = m['verdict']
        sc          = _scol(verdict)
        jdt         = str(m['jd_title'])
        exp_note    = str(m.get('experience_note') or '')
        res_yrs     = m.get('resume_years_estimated')
        strong      = list(m.get('strong_areas') or [])
        weak        = list(m.get('weak_areas') or [])
        matched     = list(m.get('matched_skills') or [])
        missing     = list(m.get('missing_skills') or [])
        match_count = int(m['matched_count'])
        total       = int(m['total_jd_requirements'])
        miss_count  = int(m['missing_count'])

        story.append(PageBreak())

        # JD banner: colored rank badge | title | tier + recommendation
        banner = Table([[
            Paragraph(RANKS[idx],
                      _ps(f'RP_Rk{idx}', fontName='Helvetica-Bold', fontSize=18,
                          textColor=C_WHITE, alignment=TA_CENTER, leading=22)),
            Paragraph(jdt,
                      _ps(f'RP_JdT{idx}', fontName='Helvetica-Bold', fontSize=13,
                          textColor=C_DARK, leading=18)),
            Paragraph(f"<b>{m['tier_label']}</b><br/><font size=\"9\">{m['recommendation']}</font>",
                      _ps(f'RP_JdP{idx}', fontName='Helvetica-Bold', fontSize=12,
                          textColor=sc, alignment=TA_RIGHT, leading=15)),
        ]], colWidths=[0.12 * W, 0.48 * W, 0.40 * W])
        banner.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (0, 0),   sc),
            ('BACKGROUND',    (1, 0), (-1, -1), C_WHITE),
            ('TOPPADDING',    (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('LEFTPADDING',   (0, 0), (0, 0),   6),
            ('LEFTPADDING',   (1, 0), (1, 0),   14),
            ('LEFTPADDING',   (2, 0), (2, 0),   8),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 10),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX',           (0, 0), (-1, -1), 1.5, sc),
        ]))
        story.append(banner)
        story.append(Spacer(1, 0.15 * inch))

        # Match Summary table (qualitative — no percentage scores)
        story.append(Paragraph('Match Summary',
                                _ps(f'RP_MBH{idx}', fontName='Helvetica-Bold', fontSize=11,
                                    textColor=C_DARK, spaceAfter=6)))
        mb_wh  = _ps(f'RP_MBW{idx}',  fontName='Helvetica-Bold', fontSize=9, textColor=C_WHITE)
        mb_met = _ps(f'RP_MBM{idx}', fontName='Helvetica', fontSize=9, textColor=C_DARK)
        mb_det = _ps(f'RP_MBD{idx}', fontName='Helvetica', fontSize=9, textColor=C_MUTED)
        exp_det = exp_note[:90] if exp_note else 'N/A'

        mb_data = [
            [Paragraph('Metric', mb_wh), Paragraph('Detail', mb_wh)],
            [Paragraph('Overall Fit', mb_met),
             Paragraph(verdict, _ps(f'RP_MBOv{idx}', fontName='Helvetica-Bold', fontSize=9, textColor=sc))],
            [Paragraph('Recommendation', mb_met),
             Paragraph(m['recommendation'], _ps(f'RP_MBRec{idx}', fontName='Helvetica-Bold', fontSize=9, textColor=sc))],
            [Paragraph('Requirements Met', mb_met),
             Paragraph(f'{match_count} of {total}', mb_det)],
            [Paragraph('Experience', mb_met),
             Paragraph(exp_det, mb_det)],
        ]
        mb_style = [
            ('BACKGROUND',    (0, 0), (-1, 0),  C_DARK),
            ('TOPPADDING',    (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING',   (0, 0), (-1, -1), 8),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 8),
            ('GRID',          (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#e2e8f0')),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ]
        for ri in range(1, len(mb_data)):
            mb_style.append(('BACKGROUND', (0, ri), (-1, ri),
                             C_WHITE if ri % 2 == 1 else rl_colors.HexColor('#f8fafc')))
        mb_tbl = Table(mb_data, colWidths=[0.28 * W, 0.72 * W])
        mb_tbl.setStyle(TableStyle(mb_style))
        story.append(mb_tbl)
        story.append(Spacer(1, 0.15 * inch))

        # Matched / Missing skills side by side
        ms = Table([[
            _skill_inner(matched, f'Matched Skills ({match_count})',
                         C_GREEN, C_GREEN, '•', f'{idx}a'),
            _skill_inner(missing, f'Missing / Gap Skills ({miss_count})',
                         C_RED, C_RED, '×', f'{idx}b'),
        ]], colWidths=[W * 0.5, W * 0.5])
        ms.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (0, 0), C_BGGRN),
            ('BACKGROUND',    (1, 0), (1, 0), C_BGRED),
            ('TOPPADDING',    (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LEFTPADDING',   (0, 0), (-1, -1), 14),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 14),
            ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
            ('LINEAFTER',     (0, 0), (0, 0),   1, rl_colors.HexColor('#d1d5db')),
        ]))
        story.append(ms)
        story.append(Spacer(1, 0.14 * inch))

        # Strong Areas
        if strong:
            story.append(Paragraph('Strong Areas',
                                    _ps(f'RP_StrH{idx}', fontName='Helvetica-Bold', fontSize=11,
                                        textColor=C_INDIGO, spaceAfter=6)))
            tbl = _tag_grid(strong, C_BGINDIG, C_INDIGO,
                            rl_colors.HexColor('#ddd6fe'), f'{idx}str')
            if tbl:
                story.append(tbl)
            story.append(Spacer(1, 0.12 * inch))

        # Skill Gaps
        if weak:
            story.append(Paragraph('Skill Gaps to Address',
                                    _ps(f'RP_GapH{idx}', fontName='Helvetica-Bold', fontSize=11,
                                        textColor=C_ORANGE, spaceAfter=6)))
            tbl = _tag_grid(weak, C_BGORG, C_DORANG,
                            rl_colors.HexColor('#fcd34d'), f'{idx}gap')
            if tbl:
                story.append(tbl)
            story.append(Spacer(1, 0.12 * inch))

        # Experience Assessment
        story.append(Paragraph('Experience Assessment',
                                _ps(f'RP_ExpH{idx}', fontName='Helvetica-Bold', fontSize=11,
                                    textColor=C_DARK, spaceAfter=6)))
        if exp_note:
            story.append(Paragraph(exp_note,
                                    _ps(f'RP_ExpN{idx}', fontName='Helvetica', fontSize=9,
                                        textColor=C_MUTED, leading=14, spaceAfter=4)))
        if res_yrs is not None:
            yr_row = Table([[
                Paragraph('Resume Shows',
                           _ps(f'RP_YrL{idx}', fontName='Helvetica-Bold',
                               fontSize=9, textColor=C_DARK)),
                Paragraph(f'~{res_yrs} year(s)',
                           _ps(f'RP_YrV{idx}', fontName='Helvetica-Bold',
                               fontSize=9, textColor=C_ORANGE)),
            ]], colWidths=[0.25 * W, 0.35 * W])
            yr_row.setStyle(TableStyle([
                ('TOPPADDING',    (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('LEFTPADDING',   (0, 0), (-1, -1), 0),
                ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
                ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(yr_row)
        story.append(Spacer(1, 0.12 * inch))

        # Recommendations
        story.append(Paragraph('Recommendations',
                                _ps(f'RP_RecH{idx}', fontName='Helvetica-Bold', fontSize=11,
                                    textColor=C_DARK, spaceAfter=6)))
        story.append(Paragraph(f"• {m['recommendation']}",
                                _ps(f'RP_Rec1{idx}', fontName='Helvetica', fontSize=9,
                                    textColor=C_DARK, leading=14, spaceAfter=4)))
        if weak:
            top_gaps = ', '.join(str(g) for g in weak[:4])
            if len(weak) > 4:
                top_gaps += '...'
            story.append(Paragraph(f'• Key areas to strengthen: {top_gaps}',
                                    _ps(f'RP_Rec2{idx}', fontName='Helvetica', fontSize=9,
                                        textColor=C_DARK, leading=14)))
        story.append(Spacer(1, 0.2 * inch))

    doc.build(story)
    buffer.seek(0)
    safe = (resume_dict.get('full_name') or 'candidate').replace(' ', '_')
    log_audit("Export", "Export PDF", record_id=resume_id, record_label=resume_dict.get('full_name'))
    return send_file(buffer, mimetype='application/pdf',
                     as_attachment=True, download_name=f'{safe}_assessment_report.pdf')


@app.route("/profile/slug/<slug>")
def public_profile(slug):
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE slug = %s", (slug,)).fetchone()
    if not resume:
        return "Profile not found", 404
    return render_template("profile.html", resume=resume)


def generate_resume_pdf(resume):
    """Generate a PDF from resume data using reportlab."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                          topMargin=0.5*inch, bottomMargin=0.5*inch,
                          leftMargin=0.75*inch, rightMargin=0.75*inch)
    story = []
    styles = getSampleStyleSheet()

    # Header with name and title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor='#1e293b',
        spaceAfter=2,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    story.append(Paragraph(resume.get('full_name', 'Resume'), title_style))

    if resume.get('title'):
        subtitle_style = ParagraphStyle(
            'Subtitle',
            parent=styles['Normal'],
            fontSize=12,
            textColor='#64748b',
            spaceAfter=12,
            alignment=TA_CENTER
        )
        story.append(Paragraph(resume.get('title'), subtitle_style))

    # Contact Info
    contact_parts = []
    if resume.get('email'):
        contact_parts.append(resume['email'])
    if resume.get('phone'):
        contact_parts.append(resume['phone'])
    if resume.get('location'):
        contact_parts.append(resume['location'])

    if contact_parts:
        contact_style = ParagraphStyle(
            'Contact',
            parent=styles['Normal'],
            fontSize=9,
            textColor='#475569',
            spaceAfter=16,
            alignment=TA_CENTER
        )
        story.append(Paragraph(' | '.join(contact_parts), contact_style))

    # Section styling
    section_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontSize=12,
        textColor='#1e293b',
        spaceAfter=8,
        spaceBefore=8,
        fontName='Helvetica-Bold',
        borderColor='#e2e8f0',
        borderWidth=0,
        borderPadding=0,
    )

    # Summary
    if resume.get('summary'):
        story.append(Paragraph('PROFESSIONAL SUMMARY', section_style))
        body_style = ParagraphStyle(
            'Body',
            parent=styles['Normal'],
            fontSize=9,
            alignment=TA_JUSTIFY,
            spaceAfter=12
        )
        story.append(Paragraph(resume['summary'], body_style))

    # Skills
    if resume.get('skills'):
        story.append(Paragraph('TECHNICAL SKILLS', section_style))
        skills_list = [s.strip() for s in resume['skills'].split('\n') if s.strip()]
        skills_text = ' • '.join(skills_list[:20])  # Limit to 20 skills
        story.append(Paragraph(skills_text, ParagraphStyle(
            'Skills', parent=styles['Normal'], fontSize=9, spaceAfter=12
        )))

    # Experience
    if resume.get('experience'):
        story.append(Paragraph('WORK EXPERIENCE', section_style))
        exp_text = resume['experience'].replace('\n', '<br/>')
        story.append(Paragraph(exp_text, ParagraphStyle(
            'Experience', parent=styles['Normal'], fontSize=9, spaceAfter=12
        )))

    # Education
    if resume.get('education'):
        story.append(Paragraph('EDUCATION', section_style))
        edu_text = resume['education'].replace('\n', '<br/>')
        story.append(Paragraph(edu_text, ParagraphStyle(
            'Education', parent=styles['Normal'], fontSize=9, spaceAfter=12
        )))

    # Projects
    if resume.get('projects'):
        story.append(Paragraph('PROJECTS', section_style))
        proj_text = resume['projects'].replace('\n', '<br/>')
        story.append(Paragraph(proj_text, ParagraphStyle(
            'Projects', parent=styles['Normal'], fontSize=9, spaceAfter=12
        )))

    # Build PDF
    doc.build(story)
    buffer.seek(0)
    return buffer


@app.route("/profile/<int:resume_id>/download-pdf")
def download_resume_pdf(resume_id):
    """Download resume as PDF."""
    if not _require_permission("download_resume", write=False):
        abort(403)
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()

    if not resume:
        return "Profile not found", 404

    resume_dict = dict(resume)
    pdf_buffer = generate_resume_pdf(resume_dict)

    filename = f"{resume_dict.get('full_name', 'Resume').replace(' ', '_')}.pdf"
    log_audit("Downloads", "Download", record_id=resume_id, record_label=resume_dict.get('full_name'))
    return send_file(
        pdf_buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )


@app.route("/profile/<int:resume_id>/extract")
def extract_resume_data(resume_id):
    """Extract resume data as JSON."""
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()

    if not resume:
        return jsonify({"error": "Profile not found"}), 404

    resume_dict = dict(resume)
    # Remove sensitive/internal fields
    resume_dict.pop('id', None)
    resume_dict.pop('created_at', None)
    resume_dict.pop('updated_at', None)

    filename = f"{resume_dict.get('full_name', 'resume').replace(' ', '_')}_extracted.json"

    return send_file(
        BytesIO(json.dumps(resume_dict, indent=2).encode()),
        mimetype='application/json',
        as_attachment=True,
        download_name=filename
    )


@app.route("/profile/<int:resume_id>/delete", methods=["POST"])
def delete_resume(resume_id):
    if not _require_permission("delete_profile"):
        abort(403)
    with db_conn() as conn:
        row = conn.execute("SELECT resume_file, full_name FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not row:
            return "Profile not found", 404
        if row["resume_file"]:
            (UPLOAD_FOLDER / row["resume_file"]).unlink(missing_ok=True)
        conn.execute("DELETE FROM resume WHERE id = %s", (resume_id,))
    log_audit("Resume Profiles", "Delete", record_id=resume_id, record_label=row["full_name"])
    flash("Profile deleted.", "success")
    return redirect(url_for("profile_list"))


@app.route("/uploads/<path:filename>")
def uploaded_file(filename):
    # Only serve allowed file types
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ALLOWED_EXTENSIONS:
        abort(403)
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


# ── New: Auto-detect "Years of Experience" from summary/experience text ────────
# Purely additive — feeds the exp_yrs dropdown a best-guess value; the user can
# still review/override it before saving, same as every other auto-filled field.

_EXP_YRS_BUCKETS = [
    (1, "0-1 yrs"), (2, "1-2 yrs"), (3, "2-3 yrs"), (4, "3-4 yrs"),
    (5, "4-5 yrs"), (6, "5-6 yrs"), (7, "6-7 yrs"), (8, "7-8 yrs"),
    (10, "8-10 yrs"), (12, "10-12 yrs"), (15, "12-15 yrs"),
]

_EXP_YRS_PATTERNS = [
    # "X+ years of [professional/relevant/total/work/industry] experience"
    # — allow up to 2 filler words between "of" and "experience" since resumes
    # phrase this many different ways ("of professional experience", "of
    # relevant industry experience", etc.).
    re.compile(r'(\d+(?:\.\d+)?)\s*\+\s*years?\s+(?:of\s+)?(?:\w+\s+){0,2}experience', re.I),
    re.compile(r'(\d+(?:\.\d+)?)\s*\+\s*yrs?\b', re.I),
    re.compile(r'(\d+(?:\.\d+)?)\s*\+\s*years?\b', re.I),
    re.compile(r'over\s+(\d+(?:\.\d+)?)\s*years?', re.I),
    re.compile(r'(\d+(?:\.\d+)?)\s*years?\s+(?:of\s+)?(?:\w+\s+){0,2}experience', re.I),
]


def _scan_for_exp_years(text):
    """Return the first regex match's numeric years value, or None."""
    for pattern in _EXP_YRS_PATTERNS:
        m = pattern.search(text)
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                continue
    return None


def _years_to_bucket(years):
    for cap, bucket in _EXP_YRS_BUCKETS:
        if years < cap:
            return bucket
    return "15+ yrs"


_EXP_DATE_TOKEN_RE = re.compile(
    r'\b(?:JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|JUN(?:E)?|'
    r'JUL(?:Y)?|AUG(?:UST)?|SEP(?:TEMBER)?|OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)'
    r'\.?\s+(\d{4})\b'
    r'|\b\d{1,2}/(\d{4})\b',
    re.I,
)
_EXP_PRESENT_RE = re.compile(r'\b(?:present|current|till\s*date|to\s*date|ongoing|now)\b', re.I)


def _estimate_years_from_dates(text):
    """Fallback for resumes with no explicit "X years of experience" phrase:
    estimate total experience from the earliest job-start year through the
    latest end year (or the current year, if any entry says "Present" /
    "Current" / etc.) found in the employment date ranges themselves —
    e.g. "2/2018 — Present", "6/2013 – 4/2017". Returns None when there
    aren't at least two distinct dated years to span.
    """
    years_found = [
        int(m.group(1) or m.group(2))
        for m in _EXP_DATE_TOKEN_RE.finditer(text)
    ]
    if len(years_found) < 2:
        return None
    earliest, latest = min(years_found), max(years_found)
    if _EXP_PRESENT_RE.search(text):
        from datetime import datetime as _dt
        latest = max(latest, _dt.now().year)
    span = latest - earliest
    return float(span) if 0 < span <= 50 else None


def _detect_years_of_experience(parsed, path=None, ext=None):
    """Scan the parsed summary/experience text for phrases like "7+ years" or
    "4+ years exp" and map the number to the closest exp_yrs dropdown bucket.

    Many resumes state years-of-experience as a tagline right under the name/
    title, BEFORE the first section heading — that text is intentionally
    excluded from the "summary" field elsewhere in this codebase (treated as
    an "introduction", not the summary), so it's invisible to the fast path
    below. When the fast path finds nothing and a file path is supplied, fall
    back to scanning the full raw extracted text (which does include that
    intro tagline).

    Returns "" when no confident match is found.
    """
    text = f"{parsed.get('summary', '')}\n{parsed.get('experience', '')}"[:2000]
    years = _scan_for_exp_years(text) if text.strip() else None
    if years is None and path is not None and ext is not None:
        try:
            full_text = extract_resume_text(path, ext)[:2000]
            years = _scan_for_exp_years(full_text)
        except Exception:
            pass
    if years is None:
        # No explicit "X years of experience" phrase anywhere — estimate the
        # span from the employment date ranges in the experience text itself,
        # rather than leaving exp_yrs blank when dated work history exists.
        years = _estimate_years_from_dates((parsed.get('experience', '') or '')[:4000])
    return _years_to_bucket(years) if years is not None else ""


@app.route("/api/parse-resume", methods=["POST"])
def parse_resume_api():
    uploaded = request.files.get("resume_file")
    if not uploaded or not uploaded.filename:
        return jsonify({"success": False, "message": "Please choose a resume file."}), 400
    if not allowed_file(uploaded.filename):
        return jsonify({"success": False, "message": "Please upload a PDF or DOCX file only."}), 400

    uploaded.seek(0, 2)
    file_size = uploaded.tell()
    uploaded.seek(0)
    if file_size > 25 * 1024 * 1024:
        return jsonify({
            "success": False,
            "message": f"File size ({file_size / 1024 / 1024:.1f} MB) exceeds the 25 MB limit.",
        }), 400

    try:
        name = secure_filename(uploaded.filename)
        ext = name.rsplit(".", 1)[1].lower()
        # Unique per request — this used to be a fixed "tmp-parse.<ext>" path
        # shared by every call to this endpoint. The dev server runs
        # threaded (concurrent requests), so two people (or the same person,
        # two tabs) previewing a resume around the same time could overwrite
        # or delete each other's temp file mid-parse: one request reads back
        # a different file than it just uploaded, or its file vanishes
        # (unlink in the other request's `finally`) while still being read.
        # This is the most likely explanation for "parsing worked, then
        # intermittently didn't, for the same file" — a single-process,
        # single-request test can never reproduce it, only concurrent real
        # usage on the shared server can.
        temp_path = UPLOAD_FOLDER / f"tmp-parse-{uuid.uuid4().hex}.{ext}"
        uploaded.save(str(temp_path))
        # parse_mode: "llm" = Resume Intelligence (AI), anything else = Quick Parse
        parse_mode  = request.form.get("parse_mode", "llm")
        parser_used = "standard"
        logger.info(f"API parse-resume: parse_mode={parse_mode}, ext={ext}")
        try:
            if ext in ("pdf", "docx") and parse_mode == "llm":
                try:
                    parsed, parser_used = parse_resume_with_llm_text(temp_path)
                except Exception as e:
                    logger.warning("Resume Intelligence failed (%s); using Quick Parse.", e, exc_info=True)
                    if ext == "pdf":
                        parsed = _parse_pdf_quick(temp_path)
                    else:
                        parsed = parse_resume_text(extract_resume_text(temp_path, ext))
                    parser_used = "text (fallback)"
            elif ext == "pdf":
                parsed = _parse_pdf_quick(temp_path)
                parser_used = "text"
            else:
                text = extract_resume_text(temp_path, ext)
                parsed = parse_resume_text(text)
                parser_used = "text"
            # New: auto-detect "X+ years of experience" from the parsed text and
            # suggest a matching exp_yrs dropdown bucket (user can still override
            # it). Runs before the temp file is deleted below, since the raw-text
            # fallback needs to re-read it.
            parsed["exp_yrs"] = _detect_years_of_experience(parsed, temp_path, ext)
        finally:
            temp_path.unlink(missing_ok=True)
        return jsonify({"success": True, "message": "Fields extracted from resume.", "data": parsed, "parser_used": parser_used})
    except Exception as e:
        logger.error(f"Error parsing resume: {e}", exc_info=True)
        return jsonify({"success": False, "message": f"Could not extract text: {e}"}), 422


@app.route("/api/resume")
def resume_api():
    with db_conn() as conn:
        resumes = conn.execute(
            "SELECT * FROM resume ORDER BY updated_at DESC, id DESC"
        ).fetchall()
    return jsonify([dict(row) for row in resumes])


# ── New: Role Definitions ─────────────────────────────────────────────────────

REGULATORY_ROLES = [
    "Design Control Consultant",
    "IFU Technical Writer",
    "IFU Team Lead",
    "Product Registration Specialist",
    "Labeling Specialist",
]

VALIDATION_ROLES = [
    "CSV Analyst", "CSV Lead", "Validation Engineer", "Validation Lead",
    "CQV Engineer", "CQV Lead", "Automation Engineer", "Automation Lead",
    "Tosca Engineer", "Tosca Lead", "Test Engineer", "Test Lead",
]

IT_ROLES = [
    "Software Developer", "Full Stack Developer", "Frontend Developer",
    "Backend Developer", "DevOps Engineer", "Cloud Engineer", "Data Analyst",
    "Data Engineer", "Business Analyst", "QA Engineer", "Automation Tester",
    "Project Manager", "Scrum Master", "UI/UX Designer", "Solution Architect",
    "Cybersecurity Engineer", "Database Administrator", "AI/ML Engineer",
    "SAP Consultant", "Salesforce Developer",
]

ALL_JD_ROLES = {
    "Regulatory Affairs": REGULATORY_ROLES,
    "Validation Roles": VALIDATION_ROLES,
    "IT Roles": IT_ROLES,
}

# ── New: Raw Upload Folder ────────────────────────────────────────────────────

RAW_UPLOAD_FOLDER = UPLOAD_FOLDER / "raw"
RAW_UPLOAD_FOLDER.mkdir(exist_ok=True)

_RAW_META_PATH = RAW_UPLOAD_FOLDER / "_meta.json"
_RAW_UPLOAD_ALLOWED = {"pdf", "docx", "doc"}  # allowed inside ZIP / direct


def _load_raw_meta():
    if _RAW_META_PATH.exists():
        try:
            return json.loads(_RAW_META_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_raw_meta(meta):
    try:
        _RAW_META_PATH.write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


# ── New: Dashboard Route ──────────────────────────────────────────────────────

@app.route("/dashboard")
def dashboard():
    if not _require_permission("view_dashboard", write=False):
        abort(403)
    from datetime import datetime as _dt
    with db_conn() as conn:
        total_resumes = conn.execute("SELECT COUNT(*) AS count FROM resume").fetchone()["count"]

        role_counts = conn.execute(
            """
            SELECT COALESCE(NULLIF(title, ''), 'No Title') AS role, COUNT(*) AS count
            FROM resume
            GROUP BY COALESCE(NULLIF(title, ''), 'No Title')
            ORDER BY count DESC
            LIMIT 20
            """
        ).fetchall()

        recent_activity = conn.execute(
            "SELECT id, full_name, title, created_at, updated_at"
            " FROM resume ORDER BY updated_at DESC LIMIT 10"
        ).fetchall()

        # ── NEW: Candidate details ─────────────────────────────────────────
        candidate_rows = conn.execute(
            "SELECT id, full_name, title, location, experience, summary, created_at,"
            " COALESCE(status, 'New') AS status,"
            " COALESCE(interview_status, 'To Be Interviewed') AS interview_status,"
            " COALESCE(exp_yrs, '') AS exp_yrs"
            " FROM resume ORDER BY created_at DESC LIMIT 200"
        ).fetchall()

        # ── NEW: Upload trend (last 30 days) ───────────────────────────────
        upload_trend = conn.execute(
            """
            SELECT TO_CHAR(DATE(created_at), 'DD Mon') AS day_label,
                   COUNT(*) AS cnt
            FROM resume
            WHERE created_at >= NOW() - INTERVAL '30 days'
            GROUP BY DATE(created_at)
            ORDER BY DATE(created_at)
            """
        ).fetchall()

        # ── NEW: HR-style dashboard metrics ─────────────────────────────────
        ensure_requirement_table(conn)
        ensure_interview_schedule_table(conn)

        open_jobs_count = conn.execute(
            "SELECT COUNT(*) AS count FROM requirement WHERE status = 'Open'"
        ).fetchone()["count"] or 0

        interviews_count = conn.execute(
            "SELECT COUNT(*) AS count FROM interview_schedule"
        ).fetchone()["count"] or 0

        # "Offers" = candidates whose offer was actually released (regardless of
        # whether they've since joined). Deliberately excludes plain "Selected by
        # Client" — being selected doesn't yet mean an offer letter went out.
        offers_count = conn.execute(
            "SELECT COUNT(*) AS count FROM resume WHERE interview_status IN"
            " ('Client Offer Released', 'Joined Vaisesika')"
        ).fetchone()["count"] or 0

        hired_count = conn.execute(
            "SELECT COUNT(*) AS count FROM resume WHERE interview_status = 'Joined Vaisesika'"
        ).fetchone()["count"] or 0

        # "Profile Shared" — reached "Profile Shared With Client" or any later stage
        # (including candidates the client later rejected, since they did get shared).
        profile_shared_count = conn.execute(
            "SELECT COUNT(*) AS count FROM resume WHERE interview_status IN"
            " ('Profile Shared With Client', 'Client Offer Released', 'Selected by Client',"
            "  'Joined Vaisesika', 'Rejected By Client')"
        ).fetchone()["count"] or 0

        # "Client Selected" — reached "Selected by Client" or later (a narrower
        # subset of Profile Shared, and a superset of Offer Released below).
        client_selected_count = conn.execute(
            "SELECT COUNT(*) AS count FROM resume WHERE interview_status IN"
            " ('Client Offer Released', 'Selected by Client', 'Joined Vaisesika')"
        ).fetchone()["count"] or 0

        # Funnel labels mirror the exact wording used in Interview Status /
        # Candidate Details elsewhere on this dashboard, so the two agree.
        funnel_stages = [
            {"label": "Applied", "count": total_resumes},
            {"label": "Profile Shared", "count": profile_shared_count},
            {"label": "Client Selected", "count": client_selected_count},
            {"label": "Offer Released", "count": offers_count},
            {"label": "Joined", "count": hired_count},
        ]

        department_counts = conn.execute(
            """
            SELECT COALESCE(NULLIF(department, ''), 'Unassigned') AS dept, COUNT(*) AS count
            FROM resume
            GROUP BY COALESCE(NULLIF(department, ''), 'Unassigned')
            ORDER BY count DESC
            """
        ).fetchall()

        recent_interviews = conn.execute(
            """
            SELECT i.id, i.position, i.interviewer, i.interview_date, i.status,
                   r.id AS resume_id, r.full_name
            FROM interview_schedule i
            JOIN resume r ON r.id = i.resume_id
            ORDER BY i.created_at DESC
            LIMIT 8
            """
        ).fetchall()

        top_requirements = conn.execute(
            """
            SELECT id, requirement_code, requirement_name, client, division,
                   num_requirement, status, profiles_shared, interviewed, offered, onsite_offshore
            FROM requirement
            WHERE status = 'Open'
            ORDER BY created_at DESC
            LIMIT 6
            """
        ).fetchall()

        all_candidates_min = conn.execute(
            "SELECT id, full_name FROM resume ORDER BY full_name"
        ).fetchall()


    # ── NEW: Helpers to extract exp years and current org ──────────────────
    def _cd_exp_years(summary_txt, exp_txt):
        txt = (summary_txt or "")[:600] + " " + (exp_txt or "")[:400]
        m = re.search(r'(\d+)\s*\+?\s*years?\s+(?:of\s+)?(?:experience|exp\b)', txt, re.I)
        if m:
            return m.group(1) + "+"
        m = re.search(r'(\d+)\s*\+\s*years?', txt, re.I)
        if m:
            return m.group(1) + "+"
        return "—"

    def _initials(name):
        parts = (name or "").split()
        if not parts:
            return "?"
        return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()

    candidate_details = [
        {
            "id": r["id"],
            "full_name": r["full_name"] or "—",
            "title": r["title"] or "—",
            "location": r["location"] or "—",
            "exp_years": r["exp_yrs"] if r["exp_yrs"] else _cd_exp_years(r["summary"], r["experience"]),
            "status": r["status"] or "New",
            "interview_status": r["interview_status"] or "To Be Interviewed",
            "initials": _initials(r["full_name"] or ""),
            "summary_snippet": (r["summary"] or "").strip()[:130],
            "created_at": r["created_at"].strftime("%Y-%m-%d") if r["created_at"] else "—",
        }
        for r in candidate_rows
    ]

    status_counts = {"New": 0, "Reviewed": 0, "Shortlisted": 0}
    for _c in candidate_details:
        status_counts[_c["status"]] = status_counts.get(_c["status"], 0) + 1

    interview_status_counts = {}
    for _c in candidate_details:
        _ist = _c["interview_status"]
        interview_status_counts[_ist] = interview_status_counts.get(_ist, 0) + 1

    return render_template(
        "dashboard.html",
        total_resumes=total_resumes,
        role_counts=list(role_counts),
        recent_activity=list(recent_activity),
        candidate_details=candidate_details,
        upload_trend=list(upload_trend),
        status_counts=status_counts,
        interview_status_counts=interview_status_counts,
        interview_statuses=INTERVIEW_STATUSES,
        interview_status_colors=INTERVIEW_STATUS_COLORS,
        interview_status_fallback_color=INTERVIEW_STATUS_FALLBACK_COLOR,
        open_jobs_count=open_jobs_count,
        interviews_count=interviews_count,
        offers_count=offers_count,
        hired_count=hired_count,
        funnel_stages=funnel_stages,
        department_counts=list(department_counts),
        recent_interviews=list(recent_interviews),
        top_requirements=list(top_requirements),
        all_candidates=list(all_candidates_min),
        interview_schedule_statuses=INTERVIEW_SCHEDULE_STATUSES,
    )


# ── New: Upload Files Route ───────────────────────────────────────────────────

def _get_raw_files_and_jds():
    """Shared context for the bulk-upload UI — used standalone at /upload-files
    and embedded as the "Bulk Upload" tab on the Add Profile page."""
    from datetime import datetime as _dt
    meta = _load_raw_meta()
    raw_files = []
    if RAW_UPLOAD_FOLDER.exists():
        for p in sorted(RAW_UPLOAD_FOLDER.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True):
            if p.is_file() and p.suffix.lower().lstrip(".") in ALLOWED_EXTENSIONS:
                st = p.stat()
                fm = meta.get(p.name, {})
                raw_files.append({
                    "name": p.name,
                    "size": st.st_size,
                    "uploaded_at": fm.get("uploaded_at") or _dt.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
                    "ext": p.suffix.lower().lstrip("."),
                    "source": fm.get("source", "direct"),
                    "original": fm.get("original", p.name),
                    "zip_name": fm.get("zip_name"),
                })

    # Pass JD list for the bulk-compare panel
    jds = []
    try:
        with db_conn() as conn:
            ensure_jd_table(conn)
            seed_jds(conn)
            jds = conn.execute(
                "SELECT id, title, category FROM job_description ORDER BY category, title"
            ).fetchall()
    except Exception:
        pass

    return raw_files, list(jds)


@app.route("/upload-files", methods=["GET", "POST"])
def upload_files():
    if not _require_permission("view_add_profile"):
        abort(403)
    import time as _time
    from datetime import datetime as _dt
    if request.method == "POST":
        files = request.files.getlist("files")
        results = []
        meta = _load_raw_meta()

        for f in files:
            if not f or not f.filename:
                continue
            name = secure_filename(f.filename)
            ext = name.rsplit(".", 1)[1].lower() if "." in name else ""

            # ── ZIP: extract each resume inside it ───────────────────────────
            if ext == "zip":
                f.seek(0)
                tmp_zip = RAW_UPLOAD_FOLDER / f"_tmp_{int(_time.time())}.zip"
                f.save(str(tmp_zip))
                try:
                    with zipfile.ZipFile(tmp_zip, "r") as zf:
                        for member in zf.namelist():
                            member_name = Path(member).name
                            # Skip hidden dotfiles and Office lock files (e.g.
                            # "~$report.docx") — these are OS/app artifacts left
                            # in the ZIP, not real resumes, and fail to parse.
                            if not member_name or member_name.startswith(".") or member_name.startswith("~$"):
                                continue
                            m_ext = member_name.rsplit(".", 1)[-1].lower() if "." in member_name else ""
                            if m_ext not in _RAW_UPLOAD_ALLOWED:
                                continue
                            m_stem = Path(member_name).stem
                            save_name = f"{slugify(m_stem)}-{int(_time.time())}.{m_ext}"
                            dest = RAW_UPLOAD_FOLDER / save_name
                            with zf.open(member) as src, open(str(dest), "wb") as dst:
                                dst.write(src.read())
                            now_str = _dt.now().strftime("%Y-%m-%d %H:%M")
                            meta[save_name] = {
                                "original": member_name, "source": "zip",
                                "zip_name": f.filename, "uploaded_at": now_str,
                            }
                            results.append({
                                "name": save_name, "original": member_name,
                                "size": dest.stat().st_size, "status": "success",
                                "source": "zip", "zip_name": f.filename,
                            })
                except Exception as e:
                    results.append({"name": f.filename, "status": "error",
                                    "message": f"ZIP error: {e}"})
                finally:
                    tmp_zip.unlink(missing_ok=True)
                continue

            # ── Existing PDF / DOCX logic (unchanged) ────────────────────────
            if not allowed_file(f.filename):
                results.append({"name": f.filename, "status": "error",
                                 "message": "Unsupported format (PDF/DOCX/ZIP only)"})
                continue
            stem = Path(name).stem
            save_name = f"{slugify(stem)}-{int(_time.time())}.{ext}"
            dest = RAW_UPLOAD_FOLDER / save_name
            f.seek(0)
            f.save(str(dest))
            now_str = _dt.now().strftime("%Y-%m-%d %H:%M")
            meta[save_name] = {
                "original": f.filename, "source": "direct", "uploaded_at": now_str,
            }
            results.append({
                "name": save_name, "original": f.filename,
                "size": dest.stat().st_size, "status": "success", "source": "direct",
            })

        _save_raw_meta(meta)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return jsonify({"results": results})
        ok = sum(1 for r in results if r["status"] == "success")
        if ok:
            flash(f"Uploaded {ok} file(s) successfully.", "success")
        return redirect(url_for("upload_files"))

    # GET
    raw_files, jds = _get_raw_files_and_jds()
    return render_template("upload_files.html", raw_files=raw_files, jds=jds)


@app.route("/uploads/raw/<path:filename>")
def raw_uploaded_file(filename):
    safe = secure_filename(filename)
    ext = safe.rsplit(".", 1)[-1].lower() if "." in safe else ""
    if ext not in ALLOWED_EXTENSIONS:
        abort(403)
    return send_from_directory(str(RAW_UPLOAD_FOLDER), safe)


@app.route("/api/delete-upload/<path:filename>", methods=["POST"])
def delete_raw_upload(filename):
    safe = secure_filename(filename)
    p = RAW_UPLOAD_FOLDER / safe
    if p.exists() and p.is_file():
        p.unlink()
        meta = _load_raw_meta()
        meta.pop(safe, None)
        _save_raw_meta(meta)
        return jsonify({"success": True})
    return jsonify({"success": False, "message": "File not found"}), 404


@app.route("/api/delete-all-uploads", methods=["POST"])
def delete_all_uploads():
    """Delete every raw uploaded file and clear metadata."""
    deleted = 0
    if RAW_UPLOAD_FOLDER.exists():
        for p in list(RAW_UPLOAD_FOLDER.iterdir()):
            if p.is_file() and p.name != '_meta.json':
                p.unlink()
                deleted += 1
    _save_raw_meta({})
    return jsonify({"success": True, "deleted": deleted})


@app.route("/api/process-raw-files", methods=["POST"])
def process_raw_files():
    """Parse all unprocessed files in RAW_UPLOAD_FOLDER and insert into resume DB."""
    _req_data = request.get_json(silent=True) or {}
    bulk_department = (_req_data.get("department") or "").strip()
    meta = _load_raw_meta()
    processed = []
    errors = []

    if not RAW_UPLOAD_FOLDER.exists():
        return jsonify({"processed": 0, "errors": [], "skipped": 0})

    # A file's "parsed" flag can go stale if its resume record was later deleted
    # (or the dev DB was reset) — the meta tracking has no way to notice that on
    # its own, so every subsequent bulk-upload run silently skips it forever,
    # never re-parsing it even though there's no matching profile anymore. Check
    # which referenced resume_ids are still real in one query, and treat a
    # "parsed" flag pointing at a deleted resume as if the file were unparsed.
    _referenced_ids = {
        fm.get("resume_id") for fm in meta.values()
        if fm.get("parsed") and fm.get("resume_id") is not None
    }
    _live_ids = set()
    if _referenced_ids:
        with db_conn() as conn:
            _rows = conn.execute(
                "SELECT id FROM resume WHERE id = ANY(%s)", (list(_referenced_ids),)
            ).fetchall()
            _live_ids = {r["id"] for r in _rows}

    files_to_process = []
    for p in RAW_UPLOAD_FOLDER.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith('_') or p.suffix.lower() not in ('.pdf', '.docx', '.doc'):
            continue
        fm = meta.get(p.name, {})
        _rid = fm.get('resume_id')
        if fm.get('parsed') and (_rid is None or _rid in _live_ids):
            continue
        files_to_process.append(p)

    skipped = []
    for p in files_to_process:
        ext = p.suffix.lower().lstrip('.')
        try:
            # Bulk uploads use the fast regex/heading parser only — no Ollama
            # calls at all. A bulk batch runs every file sequentially inside
            # one request, so per-file AI latency (identity extraction is
            # unconditional, plus _AI_FALLBACK) multiplies across the whole
            # batch; skipping AI here restores the original bulk-upload speed.
            # Single-profile uploads (Add Profile) and the Test Resume sandbox
            # are unaffected and still get full AI-assisted parsing.
            if ext == 'pdf':
                parsed_data = _parse_pdf_quick(p)
            else:
                parsed_data = parse_resume_text(extract_resume_text(p, ext))

            form_data = {k: '' for k in [
                'full_name', 'title', 'email', 'phone', 'linkedin', 'location', 'summary',
                'skills', 'experience', 'education', 'certifications', 'projects',
            ]}
            merged = merge_resume_data(form_data, parsed_data, overwrite=False)
            merged['resume_file'] = p.name
            # New: auto-detect "X+ years of experience" for bulk uploads too,
            # same as the single-profile Add Profile flow.
            merged['exp_yrs'] = _detect_years_of_experience(merged, p, ext)
            # New: apply the Department picked once for this bulk batch,
            # same column the single-profile Add Profile flow already uses.
            merged['department'] = bulk_department

            if merged.get('title'):
                _tc = re.split(r'[,|/\\–—]', merged['title'])[0].strip()
                _tc = re.sub(r'[\(\[].*$', '', _tc).strip()
                _tc = re.sub(
                    r'^(experienced|skilled|dedicated|results.driven|dynamic|seasoned|'
                    r'highly experienced|passionate|motivated|proactive|hands.on)\s+',
                    '', _tc, flags=re.I,
                ).strip()
                _words = _tc.split()
                if len(_words) > 5:
                    _m = re.match(r'^((?:[A-Z][a-zA-Z]*(?:\s+|$)){1,5})', _tc)
                    _tc = _m.group(1).strip() if _m else ' '.join(_words[:5])
                merged['title'] = _tc[:80]

            with db_conn() as conn:
                # ── Duplicate guard (same as single-profile save) ────────
                _dup_email = (merged.get('email') or '').strip()
                _dup_phone = re.sub(r"[\s\-\(\)]", "", (merged.get('phone') or '').strip())
                _dup_conds, _dup_params = [], []
                if _dup_email:
                    _dup_conds.append("LOWER(TRIM(email)) = LOWER(TRIM(%s))")
                    _dup_params.append(_dup_email)
                if _dup_phone:
                    _dup_conds.append("REGEXP_REPLACE(phone, '[\\s\\-\\(\\)]', '', 'g') = %s")
                    _dup_params.append(_dup_phone)
                if _dup_conds:
                    _dup_row = conn.execute(
                        "SELECT id, full_name FROM resume WHERE (" + " OR ".join(_dup_conds) + ") LIMIT 1",
                        _dup_params,
                    ).fetchone()
                    if _dup_row:
                        meta[p.name] = dict(meta.get(p.name, {}), parsed=True, duplicate_of=_dup_row['id'])
                        skipped.append({
                            'file': p.name,
                            'name': merged.get('full_name') or p.stem,
                            'existing_id': _dup_row['id'],
                            'existing_name': _dup_row['full_name'],
                        })
                        continue
                # ─────────────────────────────────────────────────────────

                merged['slug'] = unique_slug(conn, merged.get('full_name') or 'profile')
                cursor = conn.execute(
                    """
                    INSERT INTO resume (
                        full_name, title, email, phone, linkedin, location, summary, skills,
                        experience, education, certifications, projects, slug, resume_file,
                        exp_yrs, department, created_at, updated_at
                    ) VALUES (
                        %(full_name)s, %(title)s, %(email)s, %(phone)s, %(linkedin)s, %(location)s, %(summary)s, %(skills)s,
                        %(experience)s, %(education)s, %(certifications)s, %(projects)s, %(slug)s, %(resume_file)s,
                        %(exp_yrs)s, %(department)s, NOW(), NOW()
                    )
                    RETURNING id
                    """,
                    merged,
                )
                new_id = cursor.fetchone()['id']
                sync_skills(conn, new_id, merged.get('skills', ''))

            meta[p.name] = dict(meta.get(p.name, {}), parsed=True, resume_id=new_id)
            processed.append({
                'file': p.name,
                'name': merged.get('full_name') or p.stem,
                'id': new_id,
            })
        except Exception as e:
            errors.append({'file': p.name, 'error': str(e)})

    _save_raw_meta(meta)
    return jsonify({
        'processed': len(processed),
        'skipped': len(skipped),
        'duplicates': skipped,
        'errors': errors,
        'candidates': processed,
    })


@app.route("/api/export-compare-pdf", methods=["POST"])
def export_bulk_compare_pdf():
    """Generate a PDF from bulk-compare results for download."""
    from reportlab.platypus import Table, TableStyle, HRFlowable
    from reportlab.lib import colors as rl_colors

    data = request.get_json(silent=True) or {}
    results = data.get('results', [])
    jd_title = data.get('jd_title', 'Job Description')

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            topMargin=0.5*inch, bottomMargin=0.5*inch,
                            leftMargin=0.65*inch, rightMargin=0.65*inch)

    COL_DARK   = rl_colors.HexColor('#1e293b')
    COL_ACCENT = rl_colors.HexColor('#6366f1')
    COL_GREEN  = rl_colors.HexColor('#10b981')
    COL_RED    = rl_colors.HexColor('#ef4444')
    COL_ORANGE = rl_colors.HexColor('#f59e0b')
    COL_MUTED  = rl_colors.HexColor('#64748b')
    COL_WHITE  = rl_colors.white
    COL_HDR    = rl_colors.HexColor('#312e81')

    def sc(verdict):
        v = (verdict or '').lower()
        if 'strong' in v: return COL_GREEN
        if 'good' in v: return COL_ACCENT
        if 'partial' in v: return COL_ORANGE
        return COL_RED

    story = []
    W = A4[0] - 1.3*inch

    # Header
    hdr_tbl = Table([[
        Paragraph(f'Compare Report', ParagraphStyle('H', fontName='Helvetica-Bold',
                  fontSize=20, textColor=COL_WHITE, leading=24)),
        Paragraph(f'vs {jd_title}', ParagraphStyle('H2', fontName='Helvetica',
                  fontSize=10, textColor=rl_colors.HexColor('#c7d2fe'), leading=14)),
    ]], colWidths=[W*0.55, W*0.45])
    hdr_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), COL_HDR),
        ('TOPPADDING', (0,0), (-1,-1), 18),
        ('BOTTOMPADDING', (0,0), (-1,-1), 18),
        ('LEFTPADDING', (0,0), (-1,-1), 20),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 0.2*inch))

    if not results:
        story.append(Paragraph('No results to display.', ParagraphStyle('N', fontName='Helvetica', fontSize=11)))
    else:
        # Summary table
        hdr_row = [
            Paragraph('#', ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8.5,
                      textColor=COL_WHITE, alignment=TA_CENTER)),
            Paragraph('Candidate', ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8.5,
                      textColor=COL_WHITE)),
            Paragraph('Title', ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8.5,
                      textColor=COL_WHITE)),
            Paragraph('Fit', ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8.5,
                      textColor=COL_WHITE, alignment=TA_CENTER)),
            Paragraph('Recommendation', ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8.5,
                      textColor=COL_WHITE, alignment=TA_CENTER)),
            Paragraph('Matched', ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8.5,
                      textColor=COL_WHITE, alignment=TA_CENTER)),
        ]
        tbl_data = [hdr_row]
        row_colors = []
        for i, r in enumerate(results):
            verdict = r.get('verdict', 'Not enough data to assess')
            tier_label = r.get('tier_label', verdict)
            recommendation = r.get('recommendation', '')
            color = sc(verdict)
            bg = rl_colors.HexColor('#f8fafc') if i % 2 == 0 else COL_WHITE
            row_colors.append(bg)
            matched_str = ', '.join(r.get('matched_skills', [])[:5])
            if len(r.get('matched_skills', [])) > 5:
                matched_str += f" +{len(r['matched_skills'])-5}"
            tbl_data.append([
                Paragraph(str(i+1), ParagraphStyle('TD', fontName='Helvetica-Bold', fontSize=9,
                          textColor=color, alignment=TA_CENTER)),
                Paragraph(r.get('candidate_name', r.get('file', '—')),
                          ParagraphStyle('TD2', fontName='Helvetica-Bold', fontSize=8.5,
                                         textColor=COL_DARK)),
                Paragraph(r.get('title', '—'),
                          ParagraphStyle('TD3', fontName='Helvetica', fontSize=8,
                                         textColor=COL_MUTED)),
                Paragraph(tier_label, ParagraphStyle('TD4', fontName='Helvetica-Bold', fontSize=8.5,
                          textColor=color, alignment=TA_CENTER)),
                Paragraph(recommendation, ParagraphStyle('TD4b', fontName='Helvetica-Bold', fontSize=7.5,
                          textColor=color, alignment=TA_CENTER)),
                Paragraph(f"{r.get('matched_count',0)}/{r.get('total_jd_requirements',0)}",
                          ParagraphStyle('TD5', fontName='Helvetica', fontSize=8.5,
                                         textColor=COL_MUTED, alignment=TA_CENTER)),
            ])

        col_w = [0.05*W, 0.22*W, 0.17*W, 0.15*W, 0.24*W, 0.17*W]
        main_tbl = Table(tbl_data, colWidths=col_w, repeatRows=1)
        ts = [
            ('BACKGROUND',    (0,0), (-1,0), COL_DARK),
            ('TOPPADDING',    (0,0), (-1,-1), 7),
            ('BOTTOMPADDING', (0,0), (-1,-1), 7),
            ('LEFTPADDING',   (0,0), (-1,-1), 8),
            ('RIGHTPADDING',  (0,0), (-1,-1), 8),
            ('GRID',          (0,0), (-1,-1), 0.4, rl_colors.HexColor('#e2e8f0')),
            ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ]
        for i, bg in enumerate(row_colors):
            ts.append(('BACKGROUND', (0, i+1), (-1, i+1), bg))
        main_tbl.setStyle(TableStyle(ts))
        story.append(main_tbl)
        story.append(Spacer(1, 0.3*inch))

        # Per-candidate matched/missing detail
        detail_sec = ParagraphStyle('DS', fontName='Helvetica-Bold', fontSize=11,
                                    textColor=COL_DARK, spaceBefore=8, spaceAfter=6)
        story.append(Paragraph('Skills Detail Per Candidate', detail_sec))
        story.append(HRFlowable(width=W, thickness=0.8,
                                color=rl_colors.HexColor('#e2e8f0'), spaceAfter=8))

        for r in results:
            if r.get('error'):
                continue
            verdict = r.get('verdict', 'Not enough data to assess')
            recommendation = r.get('recommendation', '')
            color = sc(verdict)
            name = r.get('candidate_name', r.get('file', '?'))
            matched = r.get('matched_skills', [])
            missing = r.get('missing_skills', [])

            cand_hdr = Table([[
                Paragraph(name, ParagraphStyle('CN', fontName='Helvetica-Bold',
                          fontSize=10, textColor=COL_WHITE, leading=14)),
                Paragraph(f'{verdict}<br/><font size="8">{recommendation}</font>',
                          ParagraphStyle('CPct', fontName='Helvetica-Bold', fontSize=10,
                                         textColor=COL_WHITE, alignment=TA_CENTER, leading=13)),
            ]], colWidths=[W*0.6, W*0.4])
            cand_hdr.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), color),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('LEFTPADDING', (0,0), (-1,-1), 12),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            story.append(cand_hdr)

            m_text = ', '.join(matched) if matched else 'None'
            g_text = ', '.join(missing) if missing else 'None'
            detail_row = Table([[
                [Paragraph('Matched Skills', ParagraphStyle('ML', fontName='Helvetica-Bold',
                            fontSize=8, textColor=COL_GREEN, leading=12, spaceAfter=3)),
                 Paragraph(m_text, ParagraphStyle('MV', fontName='Helvetica', fontSize=7.5,
                            textColor=COL_DARK, leading=11))],
                [Paragraph('Gap / Missing Skills', ParagraphStyle('GL', fontName='Helvetica-Bold',
                            fontSize=8, textColor=COL_RED, leading=12, spaceAfter=3)),
                 Paragraph(g_text, ParagraphStyle('GV', fontName='Helvetica', fontSize=7.5,
                            textColor=COL_DARK, leading=11))],
            ]], colWidths=[W*0.5, W*0.5])
            detail_row.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (0,0), rl_colors.HexColor('#d1fae5')),
                ('BACKGROUND', (1,0), (1,0), rl_colors.HexColor('#fee2e2')),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('LEFTPADDING', (0,0), (-1,-1), 10),
                ('RIGHTPADDING', (0,0), (-1,-1), 10),
                ('LINEAFTER', (0,0), (0,0), 0.5, rl_colors.HexColor('#d1d5db')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ]))
            story.append(detail_row)
            story.append(Spacer(1, 0.12*inch))

    doc.build(story)
    buffer.seek(0)
    return send_file(buffer, mimetype='application/pdf',
                     as_attachment=True, download_name='compare_report.pdf')


@app.route("/api/bulk-compare", methods=["POST"])
def bulk_compare():
    data = request.get_json(silent=True) or {}
    filenames = data.get("files", [])
    jd_id = data.get("jd_id")

    if not filenames:
        return jsonify({"success": False, "message": "No files selected."}), 400
    if not jd_id:
        return jsonify({"success": False, "message": "No JD selected."}), 400

    try:
        with db_conn() as conn:
            ensure_jd_table(conn)
            jd = conn.execute(
                "SELECT * FROM job_description WHERE id = %s", (int(jd_id),)
            ).fetchone()
        if not jd:
            return jsonify({"success": False, "message": "JD not found."}), 404
        jd_dict = dict(jd)
    except Exception as e:
        return jsonify({"success": False, "message": f"DB error: {e}"}), 500

    results = []
    for filename in filenames:
        safe = secure_filename(filename)
        path = RAW_UPLOAD_FOLDER / safe
        if not path.exists():
            results.append({"file": filename, "candidate_name": safe,
                             "error": "File not found", "fit_percentage": 0, "verdict": "Error"})
            continue
        ext = safe.rsplit(".", 1)[-1].lower() if "." in safe else ""
        if ext not in ALLOWED_EXTENSIONS:
            results.append({"file": filename, "candidate_name": safe,
                             "error": "Unsupported format", "fit_percentage": 0, "verdict": "Error"})
            continue
        try:
            if ext == "pdf":
                parsed = _parse_pdf_quick(path)
            else:
                text = extract_resume_text(path, ext)
                parsed = parse_resume_text(text)
            score = calculate_match_score(parsed, jd_dict)
            holistic = _holistic_or_default(parsed, jd_dict)
            results.append({
                "file": filename,
                "candidate_name": parsed.get("full_name") or Path(filename).stem,
                "title": parsed.get("title", ""),
                "email": parsed.get("email", ""),
                "fit_percentage": holistic["fit_percentage"],
                "verdict": holistic["verdict"],
                "tier_label": holistic["tier_label"],
                "recommendation": holistic["recommendation"],
                "matched_count": score["matched_count"],
                "missing_count": score["missing_count"],
                "total_jd_requirements": score["total_jd_requirements"],
                "matched_skills": score["matched_skills"][:10],
                "missing_skills": score["missing_skills"][:10],
                "strong_areas": score["strong_areas"][:6],
                "weak_areas": score["weak_areas"][:6],
                "experience_note": score["experience_note"],
            })
        except Exception as e:
            results.append({
                "file": filename,
                "candidate_name": Path(filename).stem,
                "error": str(e),
                "fit_percentage": 0,
                "verdict": "Error",
                "matched_count": 0,
                "missing_count": 0,
                "total_jd_requirements": 0,
                "matched_skills": [],
                "missing_skills": [],
            })

    results.sort(key=lambda x: x.get("fit_percentage", 0), reverse=True)
    return jsonify({
        "success": True,
        "jd_title": jd_dict["title"],
        "jd_category": jd_dict.get("category", ""),
        "total": len(results),
        "results": results,
    })


# ── New: Groups Route ─────────────────────────────────────────────────────────

@app.route("/groups")
@app.route("/groups/<path:role>")
def groups(role=None):
    profiles = []
    if role:
        with db_conn() as conn:
            profiles = conn.execute(
                "SELECT id, full_name, title, email, phone, location, updated_at"
                " FROM resume WHERE title ILIKE %s ORDER BY updated_at DESC",
                (f"%{role}%",),
            ).fetchall()
    return render_template(
        "groups.html",
        all_roles=ALL_JD_ROLES,
        selected_role=role,
        profiles=list(profiles),
    )


# ── New: JD Management & Resume Matching ─────────────────────────────────────

JD_UPLOAD_FOLDER = UPLOAD_FOLDER / "jd"
JD_UPLOAD_FOLDER.mkdir(exist_ok=True)


def ensure_ai_match_cache_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ai_match_cache (
            resume_id          INTEGER NOT NULL,
            jd_id              INTEGER NOT NULL,
            fit_percentage     INTEGER NOT NULL,
            verdict            VARCHAR(200) DEFAULT '',
            rationale          TEXT DEFAULT '',
            extra              JSONB DEFAULT '{}'::jsonb,
            resume_updated_at  TIMESTAMPTZ,
            jd_updated_at      TIMESTAMPTZ,
            computed_at        TIMESTAMP DEFAULT NOW(),
            PRIMARY KEY (resume_id, jd_id)
        )
    """)

    conn.execute("""
        ALTER TABLE ai_match_cache
        ADD COLUMN IF NOT EXISTS extra JSONB
        DEFAULT '{}'::jsonb
    """)


def ensure_role_match_cache_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS role_match_cache (
            resume_id      INTEGER NOT NULL,
            jd_id          INTEGER NOT NULL,
            fit_percentage INTEGER NOT NULL,
            similarity     NUMERIC(10,6),
            updated_at     TIMESTAMP DEFAULT NOW(),

            PRIMARY KEY (resume_id, jd_id)
        )
    """)
    # Defensive migration — this table already existed in some dev environments
    # before the "extra" column (strengths/concerns/suggested_roles) was added.
    conn.execute("ALTER TABLE ai_match_cache ADD COLUMN IF NOT EXISTS extra JSONB DEFAULT '{}'::jsonb")


def ensure_jd_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS job_description (
            id               SERIAL PRIMARY KEY,
            title            VARCHAR(200) NOT NULL,
            role             VARCHAR(200) DEFAULT '',
            category         VARCHAR(100) DEFAULT 'General',
            responsibilities TEXT DEFAULT '',
            requirements     TEXT DEFAULT '',
            skills           TEXT DEFAULT '',
            keywords         TEXT DEFAULT '',
            jd_file          VARCHAR(500),
            position_status  VARCHAR(20) DEFAULT 'Open',
            created_at       TIMESTAMP DEFAULT NOW(),
            updated_at       TIMESTAMP DEFAULT NOW()
        )
    """)


# Single source of truth for the position-status enum — shared by JD Management
# and Requirement Management so the two never drift apart (see REQUIREMENT_STATUSES
# below, which reuses this list instead of defining its own).
POSITION_STATUSES = ["Open", "Fulfilled", "On Hold", "Closed"]
DEFAULT_POSITION_STATUS = "Open"


def _ensure_jd_position_status_col():
    """Add position_status to job_description for DBs created before this
    column existed, and backfill any NULLs to the default — the ADD COLUMN
    DEFAULT already back-fills existing rows, but a NULL can still slip in via
    an explicit NULL insert, so this is a belt-and-braces pass."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE job_description ADD COLUMN IF NOT EXISTS "
                "position_status VARCHAR(20) DEFAULT 'Open'"
            )
            conn.execute(
                "UPDATE job_description SET position_status = 'Open' "
                "WHERE position_status IS NULL"
            )
    except Exception:
        pass


_ensure_jd_position_status_col()


# Document-workflow status — separate from position_status (a staffing
# concept: Open/Fulfilled/On Hold/Closed). This tracks the JD *document's*
# review lifecycle: a JD is drafted, submitted for approval, approved or
# rejected (or sent back for changes), then explicitly published.
JD_WORKFLOW_STATUSES = ["Draft", "Pending Approval", "Approved", "Rejected", "Changes Requested", "Published"]
DEFAULT_JD_WORKFLOW_STATUS = "Draft"


def _ensure_jd_workflow_cols():
    """Adds the approval-workflow columns to job_description for DBs created
    before this feature existed. New rows default to 'Published' at the SQL
    level deliberately: every JD that already exists is already live and in
    active use by matching/compare/requirements, so defaulting existing rows
    to 'Draft' would silently pull them out of normal use. jd_add() overrides
    this default at INSERT time to start new JDs at 'Draft'."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE job_description ADD COLUMN IF NOT EXISTS "
                "workflow_status VARCHAR(30) NOT NULL DEFAULT 'Published'"
            )
            conn.execute(
                "ALTER TABLE job_description ADD COLUMN IF NOT EXISTS "
                "created_by INTEGER REFERENCES app_user(id)"
            )
            conn.execute("ALTER TABLE job_description ADD COLUMN IF NOT EXISTS submitted_at TIMESTAMPTZ")
            conn.execute(
                "ALTER TABLE job_description ADD COLUMN IF NOT EXISTS "
                "decided_by INTEGER REFERENCES app_user(id)"
            )
            conn.execute("ALTER TABLE job_description ADD COLUMN IF NOT EXISTS decided_at TIMESTAMPTZ")
            conn.execute("ALTER TABLE job_description ADD COLUMN IF NOT EXISTS decision_comment TEXT")
            conn.execute("ALTER TABLE job_description ADD COLUMN IF NOT EXISTS published_at TIMESTAMPTZ")
    except Exception:
        pass


_ensure_jd_workflow_cols()


def _snapshot_jd_version(conn, jd_row, edited_by):
    """Insert one jd_version row capturing jd_row's current field values,
    numbered one past whatever's already stored for this jd_id. Called with
    the PRE-edit row before an edit is applied to a Published JD, and with
    the PRE-rollback row before a rollback overwrites the live row — either
    way the snapshot always captures what's about to be overwritten, so
    history never loses a state."""
    ensure_jd_version_table(conn)
    next_num = (conn.execute(
        "SELECT COALESCE(MAX(version_number), 0) + 1 AS n FROM jd_version WHERE jd_id = %s",
        (jd_row["id"],),
    ).fetchone()["n"])
    conn.execute(
        """
        INSERT INTO jd_version
            (jd_id, version_number, title, role, category, responsibilities,
             requirements, skills, keywords, position_status, edited_by)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """,
        (jd_row["id"], next_num, jd_row["title"], jd_row["role"], jd_row["category"],
         jd_row["responsibilities"], jd_row["requirements"], jd_row["skills"],
         jd_row["keywords"], jd_row["position_status"], edited_by),
    )


def ensure_jd_version_table(conn):
    """One snapshot row per edit made to a JD that was already Published at
    edit time (see jd_edit) — never updated or deleted, so history only ever
    grows, including the extra snapshot rollback itself adds."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS jd_version (
            id               SERIAL PRIMARY KEY,
            jd_id            INTEGER NOT NULL REFERENCES job_description(id) ON DELETE CASCADE,
            version_number   INTEGER NOT NULL,
            title            TEXT, role TEXT, category TEXT, responsibilities TEXT,
            requirements     TEXT, skills TEXT, keywords TEXT, position_status TEXT,
            edited_by        INTEGER REFERENCES app_user(id),
            edited_at        TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )


PREDEFINED_JDS = [
    # ── From IFU JD PDF ──────────────────────────────────────────────────────
    {
        "title": "IFU Technical Writer",
        "role": "IFU Technical Writer",
        "category": "Regulatory Affairs",
        "responsibilities": (
            "Create, revise, and publish high-quality Instructions for Use (IFU)\n"
            "Ensure documentation complies with QSR, ISO, and internal QMS standards\n"
            "Coordinate translation activities with external agencies and internal stakeholders\n"
            "Maintain adherence to company style guides, templates, and QMS\n"
            "Collaborate with Product Development, Regulatory Affairs, Quality Assurance\n"
            "Support CAPA and complaint investigations by providing documentation expertise\n"
            "Contribute to continuous improvement of documentation standards and templates"
        ),
        "requirements": (
            "Bachelor's or Master's degree in Technical Communication, English, or Life Sciences\n"
            "2-3 years of technical writing experience within life sciences, medical device, or IVD\n"
            "Excellent English speaking, writing, and editing skills\n"
            "Strong understanding of regulated content development including labeling requirements\n"
            "Hands-on experience with CMS/CCMS, ideally DITA/XML structured authoring\n"
            "Experience working with translations and simplified English\n"
            "Experience with EU IVDR documentation requirements"
        ),
        "skills": (
            "Technical Writing\nIFU\nQSR\nISO\nQMS\nCMS\nCCMS\nDITA\nXML\n"
            "EU IVDR\nEU MDR\nTranslation\nRegulatory Compliance\nLabeling\nDocumentation\n"
            "Style Guide\nCAPA\nRisk Communication\nMedical Device\nIVD\n"
            "Simplified English\nStructured Authoring\nLife Sciences"
        ),
        "keywords": "IFU, technical writer, DITA, XML, IVDR, medical device, regulatory, QMS, labeling",
    },
    # ── From CSV JD PDF – Lead ────────────────────────────────────────────────
    {
        "title": "CSV Lead",
        "role": "CSV Lead",
        "category": "Validation Roles",
        "responsibilities": (
            "Lead and perform end-to-end validation activities for complex projects\n"
            "Author end-to-end validation deliverables\n"
            "Support automation, cloud qualification, and new IT projects\n"
            "Provide compliance consulting for GxP IT systems\n"
            "Perform CAPA and RCA using 5 x Why approach\n"
            "Manage test execution and reporting\n"
            "High level of stakeholder management across multiple concurrent tasks\n"
            "Participate in audits and inspections"
        ),
        "requirements": (
            "5+ years of experience in CSV or Computer System Validation\n"
            "Excellence in CSV and Equipment Qualification\n"
            "Hands-on experience with Agile, Waterfall, and hybrid models\n"
            "Expert in HP ALM, Azure, EDMS\n"
            "Strong IT compliance and GxP knowledge\n"
            "21 CFR Part 11 compliance experience"
        ),
        "skills": (
            "CSV\nComputer System Validation\nEquipment Qualification\nHP ALM\nAzure\nEDMS\n"
            "Agile\nWaterfall\nCAPA\nRCA\nGxP\nGMP\n21 CFR Part 11\nEU Annex 11\n"
            "GAMP 5\nRegulatory Compliance\nIT Compliance\nRisk Assessment\nStakeholder Management\n"
            "Cloud Qualification\nAutomation\nTOSCA\nValgenesis\nIQ\nOQ\nPQ\nURS"
        ),
        "keywords": "CSV Lead, validation, GxP, compliance, 21CFR, HP ALM, Azure, EDMS, GAMP5",
    },
    # ── From CSV JD PDF – Analyst ─────────────────────────────────────────────
    {
        "title": "CSV Analyst",
        "role": "CSV Analyst",
        "category": "Validation Roles",
        "responsibilities": (
            "Perform validation activities for systems/applications from various business areas\n"
            "Develop plans, execute, and deliver services with acceptable service evaluations\n"
            "Perform CAPA and RCA using 5 x Why approach\n"
            "Prepare validation documents including URS, IRA, FRA, IQ, OQ, PQ protocols\n"
            "Responsible for Software Validation, test script preparation, execution and review\n"
            "Ensure Risk Management Plan aligned to ICH Q9\n"
            "Perform 21CFR Part 11 Gap assessment for Computerized systems"
        ),
        "requirements": (
            "M.Sc. or B.Sc. in Analytical Chemistry, Life Sciences, or related field\n"
            "3-5 years of experience in CSV or Computer System Validation\n"
            "Good knowledge in GxP, 21CFR PART 11, EU Annex 11, GAMP 5\n"
            "Experience in V-model, Agile, Waterfall methodologies\n"
            "Knowledge of EDMS systems such as Valgenesis\n"
            "Experience with LIMS software\n"
            "Understanding of equipment qualification (IQ, OQ, PQ)"
        ),
        "skills": (
            "CSV\nComputer System Validation\nGAMP 5\n21 CFR Part 11\nEU Annex 11\n"
            "GxP\nGMP\nGDP\nIQ\nOQ\nPQ\nURS\nDQ\nFRA\nIRA\nCAPA\nRisk Assessment\n"
            "EDMS\nValgenesis\nLIMS\nSAP\nHP ALM\nAgile\nWaterfall\nV-model\nICH Q9\n"
            "Equipment Qualification\nSoftware Validation"
        ),
        "keywords": "CSV, GAMP5, validation, pharmaceutical, compliance, regulatory, 21CFR, GxP, LIMS",
    },
    # ── From CSV JD PDF – Compliance ──────────────────────────────────────────
    {
        "title": "IT Compliance Consultant (CSV)",
        "role": "Validation Engineer",
        "category": "Validation Roles",
        "responsibilities": (
            "Provide compliance consulting for IT projects and deliver CSV compliance deliverables\n"
            "Perform stakeholder management involving multiple stakeholders\n"
            "Manage multiple concurrent tasks with flexibility\n"
            "Perform IT risk assessments\n"
            "Execute risk-based validation strategy and deliverables\n"
            "Participate in audits and inspections related to the projects"
        ),
        "requirements": (
            "Experience in IT Compliance, CSV, and Application Support\n"
            "CSV, Regulatory Compliance, 21 CFR Part 11 expertise\n"
            "Strong IT compliance and GxP knowledge\n"
            "Written and Verbal communication skills"
        ),
        "skills": (
            "CSV\nIT Compliance\nRegulatory Compliance\n21 CFR Part 11\nGxP\n"
            "Application Support\nStakeholder Management\nRisk Assessment\nCAPA\nRCA\n"
            "Audit\nValidation\nComputer System Validation\nGMP"
        ),
        "keywords": "CSV, IT compliance, 21CFR, GxP, regulatory, validation, audit",
    },
    # ── From CSV JD PDF – Testing / TOSCA ────────────────────────────────────
    {
        "title": "CSV Testing Engineer (TOSCA/LeapWork)",
        "role": "Tosca Engineer",
        "category": "Validation Roles",
        "responsibilities": (
            "Creation and execution of manual and automated scripts using LeapWork/TOSCA\n"
            "End-to-end testing including requirement analysis, test data, test case creation, execution, defect management\n"
            "Organizing and monitoring defect management process and regression tests\n"
            "Define and implement test automation strategy including roadmap and tools\n"
            "Act as SME around Testing across business units\n"
            "Ensure quality of deliverables in alignment with stakeholder expectations"
        ),
        "requirements": (
            "2+ years of experience in CSV and Testing\n"
            "TOSCA or LeapWork experience required\n"
            "CSV, STLC Principles\n"
            "Regulatory Compliance, 21 CFR Part 11\n"
            "Azure DevOps, TIMS experience\n"
            "Experience in Agile Set-up\n"
            "ITIL understanding"
        ),
        "skills": (
            "TOSCA\nLeapWork\nLeap Work\nTest Automation\nSTLC\nCSV\n"
            "21 CFR Part 11\nAzure DevOps\nTIMS\nITIL\nAgile\n"
            "Regression Testing\nDefect Management\nTest Case Creation\nTest Scripts\n"
            "Regulatory Compliance\nComputer System Validation\nManual Testing"
        ),
        "keywords": "TOSCA, LeapWork, test automation, CSV, testing, 21CFR, Azure DevOps, STLC",
    },
    # ── From Product Registration Specialist PDF ──────────────────────────────
    {
        "title": "Product Registration Specialist",
        "role": "Product Registration Specialist",
        "category": "Regulatory Affairs",
        "responsibilities": (
            "Lead global regulatory strategies for product registrations, renewals, and change controls\n"
            "Prepare, review, and submit regulatory dossiers under MDR, IVDR, and regional frameworks\n"
            "Manage end-to-end submission lifecycle including gap analysis and response coordination\n"
            "Partner with R&D, Quality, and Manufacturing teams\n"
            "Coordinate with Notified Bodies and health authorities\n"
            "Maintain regulatory documentation and databases\n"
            "Support labeling reviews and post-market compliance activities"
        ),
        "requirements": (
            "Minimum 12 years of experience in global Regulatory Affairs for IVDs and Medical Devices\n"
            "Proven track record in hands-on submission execution\n"
            "Strong understanding of MDR, IVDR, FDA 510(k)/PMA, and regional product registration pathways\n"
            "Degree in Life Sciences, Regulatory Affairs, or Biomedical Engineering"
        ),
        "skills": (
            "MDR\nIVDR\nFDA 510k\nPMA\nRegulatory Affairs\nIVD\nMedical Device\n"
            "Notified Bodies\nRegulatory Dossiers\nGap Analysis\nChange Control\n"
            "Post-market Compliance\nLabeling\nDocument Control\nRegulatory Strategy\n"
            "Submission Management\nLife Sciences\nBiomedical Engineering\nAPAC\nLATAM\nEMEA"
        ),
        "keywords": "regulatory affairs, MDR, IVDR, FDA, medical device, IVD, submission, registration",
    },
    # ── Generic Validation Roles ──────────────────────────────────────────────
    {
        "title": "Validation Engineer",
        "role": "Validation Engineer",
        "category": "Validation Roles",
        "responsibilities": (
            "Execute validation protocols and prepare validation deliverables\n"
            "Prepare and review IQ, OQ, PQ protocols\n"
            "Author validation plans, reports, and traceability matrices\n"
            "Support CAPA and deviation investigations\n"
            "Participate in risk assessments and risk management activities\n"
            "Maintain compliance with GxP guidelines"
        ),
        "requirements": (
            "B.Sc. or M.Sc. in Life Sciences, Engineering, or related field\n"
            "2-4 years of experience in validation\n"
            "Knowledge of 21 CFR Part 11 and GAMP 5\n"
            "Experience with equipment and software validation"
        ),
        "skills": (
            "Validation\nIQ\nOQ\nPQ\nURS\nFRA\nGAMP 5\n21 CFR Part 11\n"
            "GxP\nGMP\nCAPA\nRisk Management\nEquipment Qualification\n"
            "Software Validation\nEDMS\nCompliance"
        ),
        "keywords": "validation engineer, GxP, IQ OQ PQ, GAMP5, compliance, pharmaceutical",
    },
    {
        "title": "Automation Engineer (Validation)",
        "role": "Automation Engineer",
        "category": "Validation Roles",
        "responsibilities": (
            "Design, develop, and maintain test automation frameworks\n"
            "Create and execute automated test scripts\n"
            "Integrate automation into CI/CD pipelines\n"
            "Perform root cause analysis and defect management\n"
            "Support regulatory compliance activities"
        ),
        "requirements": (
            "3+ years of test automation experience\n"
            "Proficiency in TOSCA, Selenium, or similar tools\n"
            "Knowledge of CSV and regulatory compliance\n"
            "Good scripting skills"
        ),
        "skills": (
            "Test Automation\nTOSCA\nSelenium\nCI/CD\nSTLC\nCSV\n"
            "Agile\nDefect Management\nRegression Testing\nScripting\n"
            "Python\nJava\nJavaScript\nJira\nAzure DevOps"
        ),
        "keywords": "automation engineer, TOSCA, Selenium, test automation, CI/CD, scripting",
    },
    # ── IT Roles ──────────────────────────────────────────────────────────────
    {
        "title": "Software Developer",
        "role": "Software Developer",
        "category": "IT Technology",
        "responsibilities": (
            "Design, develop, test, and maintain software applications\n"
            "Write clean, efficient, and well-documented code\n"
            "Collaborate with cross-functional teams to define and implement features\n"
            "Participate in code reviews and provide constructive feedback\n"
            "Troubleshoot, debug, and upgrade existing systems"
        ),
        "requirements": (
            "Bachelor's degree in Computer Science or related field\n"
            "3+ years of software development experience\n"
            "Proficiency in one or more programming languages\n"
            "Experience with databases and SQL\n"
            "Knowledge of software development lifecycle (SDLC)"
        ),
        "skills": (
            "Java\nPython\nJavaScript\nC#\nSQL\nGit\nREST API\n"
            "Microservices\nAgile\nScrum\nSDLC\nOOP\nUnit Testing\nCI/CD\nDocker"
        ),
        "keywords": "software developer, programming, Java, Python, JavaScript, REST API, Agile, SDLC",
    },
    {
        "title": "Full Stack Developer",
        "role": "Full Stack Developer",
        "category": "IT Technology",
        "responsibilities": (
            "Develop both front-end and back-end components of web applications\n"
            "Design and implement RESTful APIs and microservices\n"
            "Create responsive and user-friendly UI components\n"
            "Manage databases and ensure data integrity"
        ),
        "requirements": (
            "3+ years of full stack development experience\n"
            "Proficiency in frontend frameworks such as React, Angular, or Vue\n"
            "Backend experience with Node.js, Python, Java, or similar\n"
            "Experience with relational and NoSQL databases"
        ),
        "skills": (
            "React\nAngular\nVue.js\nNode.js\nPython\nJava\nJavaScript\nTypeScript\n"
            "HTML\nCSS\nREST API\nGraphQL\nSQL\nMongoDB\nPostgreSQL\nDocker\nKubernetes\nGit\nCI/CD"
        ),
        "keywords": "full stack, React, Node.js, REST API, JavaScript, frontend, backend, TypeScript",
    },
    {
        "title": "DevOps Engineer",
        "role": "DevOps Engineer",
        "category": "IT Technology",
        "responsibilities": (
            "Design and implement CI/CD pipelines\n"
            "Manage cloud infrastructure and containerized environments\n"
            "Monitor system performance and ensure availability\n"
            "Automate infrastructure provisioning using IaC tools"
        ),
        "requirements": (
            "3+ years of DevOps or SRE experience\n"
            "Experience with cloud platforms such as AWS, Azure, or GCP\n"
            "Proficiency in Docker and Kubernetes\n"
            "Infrastructure as Code experience with Terraform or Ansible"
        ),
        "skills": (
            "Docker\nKubernetes\nAWS\nAzure\nGCP\nCI/CD\nJenkins\nGitLab CI\nGitHub Actions\n"
            "Terraform\nAnsible\nPython\nBash\nLinux\nMonitoring\nPrometheus\nGrafana"
        ),
        "keywords": "DevOps, CI/CD, Docker, Kubernetes, AWS, Azure, Terraform, Jenkins, automation",
    },
    {
        "title": "QA Engineer",
        "role": "QA Engineer",
        "category": "IT Technology",
        "responsibilities": (
            "Design, develop, and execute test cases and test plans\n"
            "Perform manual and automated testing of software applications\n"
            "Identify, document, and track defects\n"
            "Implement and maintain test automation frameworks"
        ),
        "requirements": (
            "2+ years of QA testing experience\n"
            "Experience with test automation frameworks such as Selenium, Cypress, or TOSCA\n"
            "Knowledge of STLC and software testing methodologies\n"
            "Proficiency in defect tracking tools such as Jira or Azure DevOps"
        ),
        "skills": (
            "Selenium\nCypress\nTOSCA\nJira\nAzure DevOps\nSTLC\n"
            "Test Automation\nManual Testing\nAPI Testing\nPostman\n"
            "Regression Testing\nPerformance Testing\nDefect Management\nAgile\nSQL"
        ),
        "keywords": "QA engineer, testing, Selenium, automation, Jira, STLC, defect management, Agile",
    },
    {
        "title": "Data Analyst",
        "role": "Data Analyst",
        "category": "IT Technology",
        "responsibilities": (
            "Collect, process, and analyze large datasets to extract meaningful insights\n"
            "Create dashboards and visualizations using BI tools\n"
            "Prepare reports and presentations for business stakeholders\n"
            "Ensure data quality and accuracy"
        ),
        "requirements": (
            "Bachelor's degree in Statistics, Mathematics, Computer Science, or related field\n"
            "2+ years of data analysis experience\n"
            "Proficiency in SQL and data querying\n"
            "Experience with BI tools such as Power BI or Tableau\n"
            "Python or R for data analysis"
        ),
        "skills": (
            "SQL\nPython\nR\nPower BI\nTableau\nExcel\nData Analysis\n"
            "Statistics\nVisualization\nETL\nPandas\nNumPy\nBusiness Intelligence\nKPI\nDashboard"
        ),
        "keywords": "data analyst, SQL, Power BI, Tableau, Python, statistics, visualization, BI",
    },
    {
        "title": "AI/ML Engineer",
        "role": "AI/ML Engineer",
        "category": "IT Technology",
        "responsibilities": (
            "Design, develop, and deploy machine learning models\n"
            "Work with large datasets to train and evaluate ML models\n"
            "Implement MLOps pipelines for model deployment\n"
            "Research and apply state-of-the-art ML techniques"
        ),
        "requirements": (
            "Bachelor's or Master's degree in Computer Science, Mathematics, or related field\n"
            "3+ years of ML/AI development experience\n"
            "Proficiency in Python and ML frameworks\n"
            "Strong knowledge of statistics and mathematics"
        ),
        "skills": (
            "Python\nTensorFlow\nPyTorch\nMachine Learning\nDeep Learning\nNLP\nScikit-learn\n"
            "Pandas\nNumPy\nMLOps\nDocker\nKubernetes\nAWS\nStatistics\nSpark\nSQL"
        ),
        "keywords": "AI, ML, machine learning, deep learning, Python, TensorFlow, PyTorch, NLP, MLOps",
    },
    {
        "title": "SAP Consultant",
        "role": "SAP Consultant",
        "category": "IT Technology",
        "responsibilities": (
            "Implement, configure, and support SAP modules\n"
            "Analyze business requirements and translate them into SAP solutions\n"
            "Perform system integration testing and user acceptance testing\n"
            "Provide end-user training and documentation"
        ),
        "requirements": (
            "3+ years of SAP consulting experience\n"
            "Expertise in one or more SAP modules such as S/4HANA, MM, SD, FICO\n"
            "Good understanding of business processes\n"
            "Strong analytical and problem-solving skills"
        ),
        "skills": (
            "SAP\nS/4HANA\nSAP MM\nSAP SD\nSAP FICO\nSAP HR\nSAP ERP\n"
            "ABAP\nFIORI\nSAP BTP\nBusiness Analysis\nIntegration Testing\nUAT\n"
            "Data Migration\nSAP Activate\nAgile"
        ),
        "keywords": "SAP, S/4HANA, SAP consultant, ERP, ABAP, MM, SD, FICO, implementation",
    },
    {
        "title": "Design Control Consultant (IVD)",
        "role": "Design Control Consultant",
        "category": "Regulatory Affairs",
        "responsibilities": (
            "Review and strengthen Design Control processes in accordance with FDA 21 CFR Part 820, ISO 13485, and EU IVDR (2017/746)\n"
            "Ensure compliance across all Design and Development lifecycle stages (planning, inputs, outputs, verification, validation, transfer, DHF)\n"
            "Conduct detailed gap assessments against regulatory requirements and internal procedures\n"
            "Identify compliance risks and deficiencies in Design History Files (DHF) and related documentation\n"
            "Develop and execute remediation strategies and action plans\n"
            "Harmonize design control processes across multiple product lines (especially IVDs) and global/regional regulatory frameworks\n"
            "Standardize SOPs, templates, and workflows to ensure consistency and scalability\n"
            "Deliver gap assessment report with actionable recommendations\n"
            "Produce audit-ready documentation including remediated and compliant Design Control processes"
        ),
        "requirements": (
            "Bachelor's or Master's degree in Engineering, Life Sciences, or related field\n"
            "5 to 10 years of experience in Medical Devices with strong exposure to IVD products\n"
            "Proven experience in Design Control implementation and remediation\n"
            "Proven experience conducting gap assessments and process harmonization across global standards\n"
            "Strong knowledge of FDA 21 CFR Part 820 / QSR, ISO 13485, and EU IVDR requirements\n"
            "Experience with Design History Files (DHF) and lifecycle management"
        ),
        "skills": (
            "Design Control\nIVD\nIn Vitro Diagnostics\nFDA 21 CFR Part 820\nISO 13485\n"
            "EU IVDR\nQSR\nDHF\nDesign History File\nGap Assessment\nProcess Harmonization\n"
            "SOPs\nMedical Devices\nRegulatory Compliance\nRisk Assessment\nIQ OQ PQ\n"
            "Validation\nDocumentation\nAudit Readiness\nSDLC\nGxP"
        ),
        "keywords": (
            "Design Control, IVD, In Vitro Diagnostics, FDA 21 CFR Part 820, ISO 13485, "
            "EU IVDR, DHF, gap assessment, design history file, SOPs, medical devices, QSR, "
            "regulatory compliance, process harmonization, validation, audit ready"
        ),
    },
    # ── From IFU Team Lead JD PDF ─────────────────────────────────────────────
    {
        "title": "IFU Team Lead",
        "role": "IFU Team Lead",
        "category": "Regulatory Affairs",
        "responsibilities": (
            "Lead and mentor a team of IFU writers, manuals writers, and label creators\n"
            "Represent the team in cross-functional projects and management meetings\n"
            "Drive team activities and deliverables within defined scope, ensuring timelines and quality\n"
            "Create, revise, and publish high-quality Instructions for Use (IFU) and User Manuals\n"
            "Ensure compliance with QSR, ISO, IVDD/IVDR, FDA 21 CFR Part 11, ISO 13485\n"
            "Manage and coordinate translations to ensure accuracy and consistency\n"
            "Collaborate with global teams: Product Development, Regulatory Affairs, QA, Marketing\n"
            "Continuously improve documentation standards, templates, and processes\n"
            "Support quality investigations including deviations, CAPAs, and complaints"
        ),
        "requirements": (
            "Bachelor's or Master's degree in Technical Communication, English, Journalism, or Life Sciences\n"
            "5-8 years of technical writing experience within life sciences, medical device, or IVD\n"
            "Proven expertise in developing IFUs, DFUs, product labels, and user manuals\n"
            "Strong understanding of regulated content development, usability, and risk communication\n"
            "Experience working with translations and simplified English\n"
            "Familiarity with visual and multimedia tools (e.g., Adobe Illustrator)\n"
            "Prior training in QSR; experience with EU IVDR / MDR is a plus\n"
            "Excellent English speaking, writing, and editing skills"
        ),
        "skills": (
            "IFU\nUser Manuals\nDFU\nTechnical Writing\nTeam Lead\nMentoring\n"
            "QSR\nISO 13485\nFDA 21 CFR Part 11\nEU IVDR\nEU MDR\nIVDD\n"
            "Translation Management\nSimplified English\nCMS\nCCMS\nDITA\nXML\n"
            "Adobe Illustrator\nLabeling\nDocumentation\nCAPA\nDeviation\n"
            "Medical Device\nIVD\nRegulatory Compliance\nLife Sciences\nStyle Guide"
        ),
        "keywords": "IFU Team Lead, IFU writer, technical writing, IVDR, MDR, medical device, regulatory, QMS, labeling, team lead",
    },
    # ── From Labeling JD PDF ──────────────────────────────────────────────────
    {
        "title": "Labeling Specialist",
        "role": "Labeling Specialist",
        "category": "Regulatory Affairs",
        "responsibilities": (
            "Create and update labels and box prints in alignment with procedures, regulatory expectations, and market needs\n"
            "Ensure documentation complies with QSR, ISO, and internal quality and regulatory standards\n"
            "Maintain adherence to company style guides, templates, and quality management systems (QMS)\n"
            "Collaborate with Product Development, Regulatory Affairs, Quality Assurance, Marketing, and Global Operations\n"
            "Drive and manage activities within the assigned area of responsibility\n"
            "Contribute to continuous improvement of documentation standards and templates\n"
            "Support Deviation, CAPA, and complaint investigations by providing documentation expertise"
        ),
        "requirements": (
            "University degree in the life science field and/or relevant work experience\n"
            "Excellent English speaking, writing, and editing skills\n"
            "2-3 years of label creation experience in life sciences, medical devices, or IVD industry\n"
            "Familiarity with visual and multimedia tools (e.g., Adobe InDesign, Illustrator)\n"
            "Proven ability to manage multiple documentation projects in a global, regulated environment\n"
            "Experience with EU IVDR documentation requirements"
        ),
        "skills": (
            "Labeling\nLabel Design\nBox Print\nAdobe InDesign\nAdobe Illustrator\n"
            "QSR\nISO\nQMS\nEU IVDR\nEU MDR\nRegulatory Compliance\nDocumentation\n"
            "Style Guide\nCAPA\nDeviation\nMedical Device\nIVD\nLife Sciences\n"
            "Technical Writing\nMultilingual\nTranslation\nCross-functional Collaboration"
        ),
        "keywords": "labeling, label design, IVD, medical device, IVDR, regulatory, QMS, Adobe InDesign, documentation",
    },
]


def seed_jds(conn):
    """Populate the starter JD set on a genuinely empty table only. This runs
    on every /jd-management page load, so re-checking title-by-title (as
    before) meant deleting a predefined JD got silently undone on the very
    next page load — the delete worked, but the next GET re-seeded it right
    back in. A one-time "table is empty" check preserves the first-run
    starter content without fighting the user's own deletions afterward.
    """
    (count_row,) = conn.execute("SELECT COUNT(*) AS n FROM job_description").fetchall()
    if count_row["n"] > 0:
        return
    for jd in PREDEFINED_JDS:
        conn.execute(
            """
            INSERT INTO job_description
                (title, role, category, responsibilities, requirements, skills, keywords)
            VALUES
                (%(title)s, %(role)s, %(category)s, %(responsibilities)s,
                 %(requirements)s, %(skills)s, %(keywords)s)
            """,
            jd,
        )


# ── Matching Algorithm ────────────────────────────────────────────────────────
def cosine_similarity(v1, v2):

    if not v1 or not v2:
        return 0

    dot = sum(a * b for a, b in zip(v1, v2))

    mag1 = math.sqrt(sum(a * a for a in v1))
    mag2 = math.sqrt(sum(a * a for a in v2))

    if mag1 == 0 or mag2 == 0:
        return 0

    return dot / (mag1 * mag2)


def resume_match_text(resume):

    return "\n".join([
        str(resume.get("title", "")),
        str(resume.get("summary", "")),
        str(resume.get("skills", "")),
        str(resume.get("experience", "")),
        str(resume.get("projects", "")),
        str(resume.get("certifications", ""))
    ])


def jd_match_text(jd):

    return "\n".join([
        str(jd.get("title", "")),
        str(jd.get("skills", "")),
        str(jd.get("requirements", "")),
        str(jd.get("responsibilities", "")),
        str(jd.get("keywords", ""))
    ])


def rebuild_role_match_cache(resume_id):
    print("CACHE BUILD STARTED", resume_id)
    with db_conn() as conn:
        resume = conn.execute(
            "SELECT * FROM resume WHERE id = %s", (resume_id,)
        ).fetchone()
        if not resume:
            return
        resume = dict(resume)    
        resume_embedding = get_embedding(    
            resume_match_text(resume)  
      )
        if not resume_embedding:      
            return
        jds = conn.execute(
            """
            SELECT *
            FROM job_description         
            WHERE position_status = 'Open'         
            """
        ).fetchall()
        print("JD count:", len(jds))

        for jd in jds:
            jd_dict = dict(jd)
            jd_embedding = get_embedding(              
                jd_match_text(jd_dict)          
                )

            if not jd_embedding:
                continue

            similarity = cosine_similarity(
                resume_embedding,          
                jd_embedding  
            )

            fit_percentage = round(similarity * 100)

            conn.execute(
                """
                INSERT INTO role_match_cache            
                (
                   resume_id,           
                   jd_id,       
                   fit_percentage,                
                   similarity,              
                   updated_at
                )
                VALUES
                (              
                %s,                
                %s,           
                %s,            
                %s,                    
                NOW()             
                )

                ON CONFLICT (resume_id, jd_id)

                DO UPDATE SET                   
                fit_percentage = EXCLUDED.fit_percentage,                    
                similarity = EXCLUDED.similarity,               
                updated_at = NOW()
                """,
                (                 
                    resume_id,                  
                    jd_dict["id"],                    
                    fit_percentage,                    
                    similarity
                )            
            )

            
def _normalize(text):
    return re.sub(r'\s+', ' ', (text or '').lower()).strip()


def _jd_role_matches(resume_title, jd_role):
    """True if the candidate's own declared role/title matches a JD's role.

    Used to prioritize a JD for the candidate's own specialty ahead of pure
    skill-keyword-overlap scoring in the Top Matches ranking — otherwise a
    JD can score lower than an unrelated one purely because its specific
    tool/skill list has less textual overlap with the resume, even though
    the role itself is an exact match (e.g. a Cybersecurity Engineer
    candidate not seeing the Cybersecurity Engineer JD in their own top 3).
    Both sides are drawn from the same role vocabulary (the Groups dropdown
    used at upload / the JD role dropdown), so an exact match once
    whitespace differences are ignored ("Cyber Security" vs "Cybersecurity")
    is the right bar — no fuzzy/partial matching, to avoid conflating
    distinct roles like "QA Engineer" and "QA Lead".
    """
    a = re.sub(r'\s+', '', (resume_title or '').lower())
    b = re.sub(r'\s+', '', (jd_role or '').lower())
    return bool(a) and bool(b) and a == b


# JD-posting logistics lines (not actual requirements) that show up mixed into
# free-text requirements fields on any JD, not just one specific posting.
_JD_METADATA_LINE_RE = re.compile(
    r'^(location|type|salary|duration|department|employment\s*type|work\s*mode|shift|reports?\s*to)\s*:',
    re.I,
)


def _parse_jd_items(skills_text, requirements_text, keywords_text):
    combined = "\n".join(filter(None, [
        str(skills_text or ''), str(requirements_text or ''), str(keywords_text or '')
    ]))
    seen, items = set(), []
    for line in combined.splitlines():
        line = re.sub(r'^[•\-–*►◆▸▪\d\.\)\s]+', '', line).strip()
        if not line or _JD_METADATA_LINE_RE.match(line):
            continue
        for part in re.split(r'[,;]', line):
            part = part.strip()
            key = _normalize(part)
            if 2 <= len(part) <= 80 and key not in seen:
                seen.add(key)
                items.append(part)
    return items


# Common abbreviation/full-form pairs so e.g. a JD's "JavaScript" matches a resume
# that only ever wrote "JS", without needing an LLM call to recognize the synonym.
# Grouped as sets (not a flat dict) so any member in a group pulls in every other
# member as a variant to check.
_SKILL_SYNONYM_GROUPS = [
    {"js", "javascript"}, {"ts", "typescript"}, {"py", "python"},
    {"ml", "machine learning"}, {"ai", "artificial intelligence"},
    {"k8s", "kubernetes"}, {"aws", "amazon web services"},
    {"gcp", "google cloud platform"}, {"nlp", "natural language processing"},
    {"oop", "object oriented programming"}, {"qa", "quality assurance"},
    {"sql", "structured query language"}, {"iac", "infrastructure as code"},
    {"sre", "site reliability engineering"}, {"crm", "customer relationship management"},
    {"erp", "enterprise resource planning"}, {"bi", "business intelligence"},
    {"etl", "extract transform load"}, {"llm", "large language model"},
    {"ux", "user experience"}, {"ui", "user interface"},
    {"devops", "development operations"}, {"iot", "internet of things"},
    {"ci/cd", "cicd", "ci cd", "continuous integration continuous deployment"},
]


def _synonym_variants(skill_n):
    """All known synonym phrases for an already-normalized skill, plus itself."""
    variants = {skill_n}
    for group in _SKILL_SYNONYM_GROUPS:
        if skill_n in group:
            variants |= group
    return variants


def _skill_matches(skill, corpus):
    skill_n = _normalize(skill)
    corpus_n = _normalize(corpus)
    # Short skills (IQ, FRA, CAPA, ...) need a whole-word check — as plain substrings
    # they'd false-positive inside unrelated words (e.g. "iq" inside "techniques").
    if len(skill_n) <= 4:
        if re.search(r'\b' + re.escape(skill_n) + r'\b', corpus_n):
            return True
    elif skill_n in corpus_n:
        return True
    for variant in _synonym_variants(skill_n) - {skill_n}:
        if len(variant) <= 4:
            if re.search(r'\b' + re.escape(variant) + r'\b', corpus_n):
                return True
        elif variant in corpus_n:
            return True
    words = [w for w in skill_n.split() if len(w) > 3]
    if len(words) >= 2 and all(w in corpus_n for w in words[:2]):
        return True
    return False


def _extract_years_required(text):
    for pat in [
        r'minimum\s+(\d+)\s*years?',
        r'at\s+least\s+(\d+)\s*years?',
        r'(\d+)\+\s*years?',
        r'(\d+)\s*[-–]\s*\d+\s*years?',
        r'(\d+)\s*years?\s+of\s+experience',
    ]:
        m = re.search(pat, text, re.I)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass
    return None



# Generic-language words filtered out of the whole-document comparison so the
# similarity reflects shared domain/role content, not shared filler words —
# not specific to any one JD or resume.
_DOC_STOPWORDS = set("""
a an the and or but if then else for of to in on at by with without within into from as is are was
were be been being this that these those it its youre your our their his her he she they them i we
do does did have has had will would can could should shall may might must not no nor so such than too
very s t don now years year experience strong excellent good working skills location type permanent
full time hybrid remote responsible ability including etc using well also both across each per any
role team teams work works related including
""".split())


def _doc_word_freq(text):
    words = re.findall(r"[a-z]{3,}", (text or '').lower())
    return Counter(w for w in words if w not in _DOC_STOPWORDS)


_JD_MIRROR_NGRAM_SIZE = 6
_JD_MIRROR_MIN_JD_NGRAMS = 5     # too little distinctive JD text to judge meaningfully below this
_JD_MIRROR_MIN_MATCHES = 3       # fewer matches than this is noise, not a real signal
_JD_MIRROR_SOME_THRESHOLD = 8    # % overlap
_JD_MIRROR_HIGH_THRESHOLD = 30   # % overlap

# Deterministic backstop for the "achievements_measurable" authenticity rule
# (see _evaluate_authenticity / _hybrid_match) — matches a measurable,
# verifiable claim: a percentage, a team size, a before/after reduction
# ("45min to 4min"), a scale ("2M transactions/day"), a named improvement
# with a number attached, etc.
# Found live (2026-08-07): a fully fabricated resume — real company names,
# real-looking date ranges, generic project titles for every claimed skill,
# grammatically smooth — scored "Strong Evidence" (70-80) from the LLM
# judge alone under the old evidence-depth system. It has ZERO numbers
# anywhere. The LLM's own judgment proved too easy to satisfy with "looks
# structurally complete" rather than "is actually specific" (confirmed:
# rewording the prompt shifted this same resume's score all over the place,
# 35 to 65, depending on phrasing — small-model judgment on this specific
# question is not reliable enough to trust alone). This regex is the
# reliable, un-gameable half of the signal, and is now the one HARD GATE in
# _AUTHENTICITY_RULES: a resume with literally no quantified claims cannot
# be labeled "Likely Genuine", no matter what the other 7 checks say.
_QUANT_SPECIFICITY_RE = re.compile(
    r"\d+\s*%"
    r"|\d+\s*[kmb]\+?\s*(?:transactions?|requests?|users?|customers?|records?|documents?|incidents?|"
    r"tickets?|systems?|sites?|reports?|deployments?|releases?|rows?|queries)"
    r"|team\s+of\s+\d+|\d+[\s-]*(?:member|engineer|person|people)"
    r"|\$\s*\d|\d+\s*(?:crore|lakh|million|billion)"
    r"|\d+\s*(?:min|minutes|hours?|days?|weeks?|months?)\s*(?:to|->|→)\s*\d+"
    r"|from\s+\d+.{0,15}to\s+\d+"
    r"|(?:reduced|increased|improved|cut|grew|boosted|decreased|saved)\b[^.]{0,40}?\d+",
    re.I,
)
_QUANT_SPECIFICITY_MIN_TEXT_LEN = 150  # skip the hard gate for a stub-length resume — nothing to judge yet


def _quantified_specificity_count(text):
    return len(_QUANT_SPECIFICITY_RE.findall(text or ""))


# Resume-authenticity signal: 8 consistency checks, one deterministic (the
# proven regex above) plus 7 LLM-judged plain booleans (much more stable
# for a small model than a 0-100 scale — see _llm_judge_match), combined
# via a uniform failure-count threshold across all 8. Modeling this as a
# registry (rather than one hardcoded if/else) is what makes it genuinely
# modular: adding a new check later is one entry here, not a rewrite of
# _evaluate_authenticity's decision logic.
#
# `achievements_measurable` was originally an unconditional hard gate (any
# resume with zero quantified claims was force-flagged, regardless of the
# other 7). Live data (2026-08-07) showed this over-triggers: many resumes
# in genuinely non-metric-driven domains (compliance/QA/audit work) have
# zero quantified claims *and* are completely genuine — the hard gate
# flagged most of them, defeating the point of a differentiating signal
# (the same failure mode that killed the old tri-bucket Evidence Depth
# system). The theoretical worst case the hard gate guarded against (a
# fabricated resume that fools every one of the other 7 checks) never
# actually showed up in live testing — the real fabricated case
# ("Priya Narayanan", see below) always tripped at least one other check
# too. So `achievements_measurable` is now one equal vote among 8, and
# `_AUTHENTICITY_FAIL_THRESHOLD = 2` (the user's own spec: "multiple
# inconsistencies") is what decides — a resume can lack quantified
# achievements alone and still read as genuine, but that plus even one more
# red flag is enough to ask a recruiter to take a closer look.
_AUTHENTICITY_FAIL_THRESHOLD = 2
_AUTHENTICITY_MIN_CHECKS_PRESENT = 4  # below this, treat as insufficient signal, not "clean"

_AUTHENTICITY_RULES = [
    {"key": "achievements_measurable", "label": "Achievements contain measurable evidence",
     "source": "deterministic"},
    {"key": "skills_match_projects", "label": "Skills are supported by project experience",
     "source": "llm"},
    {"key": "experience_matches_responsibilities", "label": "Responsibilities align with experience",
     "source": "llm"},
    {"key": "timeline_consistent", "label": "Timeline is consistent",
     "source": "llm"},
    {"key": "designation_matches_experience", "label": "Designation matches years of experience",
     "source": "llm"},
    {"key": "technologies_demonstrated", "label": "Technologies are demonstrated in experience",
     "source": "llm"},
    {"key": "projects_realistic", "label": "Projects appear technically realistic",
     "source": "llm"},
    {"key": "no_major_contradictions", "label": "No major contradictions found",
     "source": "llm"},
]
_AUTHENTICITY_SOFT_RULE_KEYS = [r["key"] for r in _AUTHENTICITY_RULES if r["key"] != "achievements_measurable"]


def _evaluate_authenticity(resume_dict, consistency_checks):
    """Binary resume-authenticity decision — "Likely Genuine" or "Needs
    Manual Verification" — built from _AUTHENTICITY_RULES. Replaces the old
    tri-bucket Evidence Depth system, which mostly collapsed everything into
    "Shallow" once candidates started using AI writing tools, making it
    useless for differentiating candidates.

    `consistency_checks` is the raw dict the LLM returned (or None if the
    LLM judge didn't run at all — Ollama down). Each of the 7 LLM-judged
    keys may individually be missing/malformed without invalidating the
    rest, same graceful-degradation spirit as the old evidence_depth field.

    Returns None (not a fallback bucket) when consistency_checks itself is
    None, so callers can distinguish "known result" from "not assessed".
    """
    if consistency_checks is None:
        return None

    quant_text = f"{resume_dict.get('experience') or ''} {resume_dict.get('projects') or ''}"
    achievements_assessed = len(quant_text.strip()) >= _QUANT_SPECIFICITY_MIN_TEXT_LEN
    achievements_pass = _quantified_specificity_count(quant_text) > 0 if achievements_assessed else None

    rules = {}
    for rule in _AUTHENTICITY_RULES:
        key = rule["key"]
        if key == "achievements_measurable":
            if achievements_pass is None:
                rules[key] = {"pass": None, "reason": "Resume too short to assess.", "source": rule["source"]}
            else:
                reason = ("Contains measurable outcomes, metrics, or scale." if achievements_pass
                          else "No measurable outcomes, metrics, or specific quantified details "
                               "found anywhere in the experience/projects text.")
                rules[key] = {"pass": achievements_pass, "reason": reason, "source": rule["source"]}
            continue

        raw = consistency_checks.get(key) if isinstance(consistency_checks, dict) else None
        if isinstance(raw, dict) and isinstance(raw.get("pass"), bool):
            rules[key] = {"pass": raw["pass"], "reason": str(raw.get("reason") or "").strip()[:120],
                          "source": rule["source"]}
        else:
            rules[key] = {"pass": None, "reason": "", "source": rule["source"]}

    present = [r["key"] for r in _AUTHENTICITY_RULES if rules[r["key"]]["pass"] is not None]
    failed_count = sum(1 for k in present if rules[k]["pass"] is False)

    insufficient_signal = len(present) < _AUTHENTICITY_MIN_CHECKS_PRESENT
    if insufficient_signal:
        status = "Needs Manual Verification"
    elif failed_count >= _AUTHENTICITY_FAIL_THRESHOLD:
        status = "Needs Manual Verification"
    else:
        status = "Likely Genuine"

    # Positive reasons reuse the rule's success-framed label directly (e.g.
    # "Skills are supported by project experience") — negative reasons use
    # the specific concern text instead (the LLM's own reason, or the
    # deterministic explanation), since reusing the same label with a
    # warning icon would misleadingly read as a positive claim.
    reasons_positive = [r["label"] for r in _AUTHENTICITY_RULES if rules[r["key"]]["pass"] is True][:4]
    reasons_negative = [
        rules[r["key"]]["reason"] or f"{r['label']} — not confirmed"
        for r in _AUTHENTICITY_RULES if rules[r["key"]]["pass"] is False
    ][:4]
    if insufficient_signal and not reasons_negative:
        reasons_negative = ["Not enough signal from the AI reviewer to confirm consistency"][:4]

    if status == "Likely Genuine":
        explanation = (
            f"This resume passed {len(reasons_positive)} of {len(_AUTHENTICITY_RULES)} consistency "
            "checks with no significant contradictions. " + (
                (reasons_positive[0] + ", and " + reasons_positive[1] + ".") if len(reasons_positive) >= 2
                else (reasons_positive[0] + "." if reasons_positive else "")
            )
        ).strip()
    else:
        lead = ("Insufficient signal to confirm consistency." if insufficient_signal
                else f"This resume shows {len(reasons_negative)} consistency concern(s) worth a closer look.")
        explanation = (lead + " " + ("; ".join(reasons_negative) + "." if reasons_negative else "")).strip()

    return {
        "status": status,
        "reasons_positive": reasons_positive,
        "reasons_negative": reasons_negative,
        "explanation": explanation,
        "rules": rules,
    }


def _extract_meaningful_ngrams(text, n=_JD_MIRROR_NGRAM_SIZE):
    """Word n-grams ("shingles") from text, keeping only ones with enough
    non-stopword content to be distinctive phrasing rather than generic
    connective filler ("...and the ability to work with..."). Used to
    detect verbatim phrase reuse — see _jd_mirroring_risk.
    """
    words = re.findall(r"[a-z0-9]+", (text or "").lower())
    grams = set()
    for i in range(len(words) - n + 1):
        gram = words[i:i + n]
        meaningful = sum(1 for w in gram if w not in _DOC_STOPWORDS and len(w) > 2)
        if meaningful >= 3:
            grams.add(" ".join(gram))
    return grams


def _jd_mirroring_risk(resume_dict, jd_dict):
    """Cheap, deterministic, no-LLM check: how much of the JD's own
    distinctive phrasing (6-word shingles) appears verbatim in the resume.
    High overlap is a strong, independent tell that a resume was derived by
    closely following — or literally pasting from — the JD's own text,
    rather than describing genuine, independently-written experience: the
    AI-generated/JD-copied resume pattern this feature targets.

    Deliberately separate from the keyword/skill match score: a resume
    sharing required SKILL NAMES with a JD is a good match, not a red flag.
    This instead looks at whether whole PHRASES (not just terms) recur,
    which a real candidate describing their own work essentially never does
    by coincidence — confirmed empirically: an independently-written resume
    and a paraphrased-in-own-words resume both scored 0% against a real JD,
    a resume with one lifted sentence scored ~11%, and a near-verbatim copy
    scored ~69%.

    Computed fresh every time (not cached) — it's pure string comparison
    over already-fetched text, cheap enough to redo on every read, so it
    always reflects the current resume/JD content with no staleness concern
    the way an LLM-computed score needs cache invalidation for.
    """
    jd_text = " ".join(str(jd_dict.get(k) or "") for k in
                        ("skills", "requirements", "responsibilities", "keywords"))
    resume_text = " ".join(str(resume_dict.get(k) or "") for k in
                            ("summary", "skills", "experience", "projects"))

    jd_grams = _extract_meaningful_ngrams(jd_text)
    if len(jd_grams) < _JD_MIRROR_MIN_JD_NGRAMS:
        return {"overlap_pct": 0, "risk_label": None, "matched_phrases": []}

    resume_grams = _extract_meaningful_ngrams(resume_text)
    matched = jd_grams & resume_grams
    overlap_pct = round(len(matched) / len(jd_grams) * 100)

    if len(matched) < _JD_MIRROR_MIN_MATCHES:
        risk_label = None
    elif overlap_pct >= _JD_MIRROR_HIGH_THRESHOLD:
        risk_label = "High JD-Text Overlap"
    elif overlap_pct >= _JD_MIRROR_SOME_THRESHOLD:
        risk_label = "Some JD-Text Overlap"
    else:
        risk_label = None

    return {
        "overlap_pct": overlap_pct,
        "risk_label": risk_label,
        "matched_phrases": sorted(matched, key=len, reverse=True)[:5],
    }


def _whole_doc_similarity(jd_dict, resume_dict):
    """Cosine similarity between the JD's full text and the resume's full text,
    treated as whole documents rather than a checklist of discrete skills.
    A shared-vocabulary comparison across the whole JD and whole resume gives
    credit for genuine overall domain/role overlap (shared responsibilities,
    tools, terminology) instead of scoring a category to zero the moment one
    specific line item isn't a literal hit — closer to how a recruiter skims
    a resume against a JD as a whole rather than checking off a line-by-line
    list. Returns (similarity 0..1, shared significant terms sorted by
    combined frequency).
    """
    jd_text = " ".join(str(jd_dict.get(k) or '') for k in
                        ('title', 'skills', 'requirements', 'responsibilities', 'keywords'))
    resume_text = " ".join(str(resume_dict.get(k) or '') for k in
                            ('title', 'summary', 'skills', 'experience', 'projects', 'certifications', 'education'))
    jd_freq = _doc_word_freq(jd_text)
    res_freq = _doc_word_freq(resume_text)
    if not jd_freq or not res_freq:
        return 0.0, []

    common = set(jd_freq) & set(res_freq)
    dot = sum(jd_freq[w] * res_freq[w] for w in common)
    jd_norm = math.sqrt(sum(v * v for v in jd_freq.values()))
    res_norm = math.sqrt(sum(v * v for v in res_freq.values()))
    similarity = dot / (jd_norm * res_norm) if jd_norm and res_norm else 0.0

    shared = sorted(common, key=lambda w: jd_freq[w] * res_freq[w], reverse=True)
    return similarity, shared


def _holistic_match(resume_dict, jd_dict):
    """Whole-document fit assessment: cosine similarity between the JD's full
    text and the resume's full text, rather than a line-by-line skills
    checklist — see _whole_doc_similarity(). A checklist scores a candidate
    down for every individual missing acronym even when the two documents are
    clearly the same domain overall; comparing the documents as a whole avoids
    that, at the cost of being a cruder, less specific signal.
    """
    similarity, shared_terms = _whole_doc_similarity(jd_dict, resume_dict)
    if not shared_terms and similarity == 0.0:
        return None
    # Raw cosine similarity between two natural-language documents (unweighted
    # term frequency, no IDF corpus available) sits low in absolute terms even
    # for a strong match, so it's rescaled onto a friendlier 0-100 range rather
    # than reported as-is; only relative ordering/tiering matters since the UI
    # doesn't surface the raw number.
    pct = max(0, min(100, round(similarity * 220)))

    # Tier/label/recommendation come from _tier_from_pct() — the one place in
    # the app that maps a percentage to a tier — so this never drifts out of
    # sync with the AI hybrid assessment or the Top 3 Matching Roles cards
    # (they used to each bucket the same percentage differently).
    verdict, tier_label, recommendation = _tier_from_pct(pct)

    rationale = f"Compared the full job description against the full resume as whole documents; {verdict.lower()} based on shared domain content."

    return {
        "fit_percentage": pct,
        "verdict": verdict,
        "tier_label": tier_label,
        "recommendation": recommendation,
        "rationale": rationale[:500],
        "strengths": shared_terms[:6],
        "concerns": "",
        "suggested_roles": [],
    }


def _holistic_or_default(resume_dict, jd_dict):
    """_holistic_match(), normalized to always return a full dict (never None)
    so every caller gets the same keys (including tier_label/recommendation)
    without repeating its own "if holistic else ..." fallback.
    """
    holistic = _holistic_match(resume_dict, jd_dict)
    if holistic:
        return holistic
    return {
        "fit_percentage": 0,
        "verdict": "Not enough data to assess",
        "tier_label": "Not enough data",
        "recommendation": "— No assessment available",
        "rationale": "",
        "strengths": [],
        "concerns": "",
        "suggested_roles": [],
    }


def _estimate_exp_years(experience_text):
    from datetime import datetime as _dt2
    if not experience_text:
        return None
    current_year = _dt2.now().year
    ranges = re.findall(
        r'(\d{4})\s*[-–—to]+\s*(present|current|till\s*date|\d{4})',
        experience_text, re.I
    )
    total = 0
    for start, end in ranges:
        try:
            s = int(start)
            e = current_year if re.match(r'present|current|till', end, re.I) else int(end)
            if 1990 <= s <= current_year and e >= s:
                total += e - s
        except Exception:
            pass
    return total if total > 0 else None


def calculate_match_score(resume_d, jd_d, include_title=True):
    """include_title=False is used by the AI hybrid judgment path
    (_hybrid_match) so that editing pure organizational metadata (the
    Role & Group field, which sets `title`) can never change the keyword
    grounding facts fed to the LLM — only the fast pre-AI estimate
    (_quick_estimate_jd_match) and its own candidate-selection role-bonus
    still consider title, unchanged from before."""
    corpus_fields = ['skills', 'experience', 'summary', 'projects', 'certifications', 'education']
    if include_title:
        corpus_fields.append('title')
    corpus = " ".join(str(resume_d.get(k) or '') for k in corpus_fields)
    items = _parse_jd_items(jd_d.get('skills'), jd_d.get('requirements'), jd_d.get('keywords'))
    if not items:
        return {
            'match_percentage': 0, 'skills_match_percentage': 0, 'experience_match_percentage': 0,
            'matched_skills': [], 'missing_skills': [], 'total_jd_requirements': 0,
            'matched_count': 0, 'missing_count': 0, 'match_level': 'No Data',
            'match_level_color': 'gray', 'strong_areas': [], 'weak_areas': [],
            'experience_note': 'No skills data found in this JD.',
        }
    matched, missing = [], []
    for item in items:
        (matched if _skill_matches(item, corpus) else missing).append(item)
    skills_pct = len(matched) / len(items) * 100
    jd_exp_text = " ".join(str(jd_d.get(k) or '') for k in ['requirements', 'responsibilities'])
    jd_yrs = _extract_years_required(jd_exp_text)
    res_yrs = _estimate_exp_years(str(resume_d.get('experience') or ''))
    if jd_yrs and res_yrs is not None:
        exp_pct = min(100.0, res_yrs / jd_yrs * 100)
        exp_note = f"Estimated {res_yrs} yr(s) in resume; JD requires {jd_yrs}+ yr(s)"
    elif jd_yrs:
        exp_pct = 60.0
        exp_note = f"JD requires {jd_yrs}+ yr(s); experience timeline not detected in resume"
    else:
        exp_pct = 80.0
        exp_note = "No specific experience years requirement stated in JD"
    final = min(100.0, skills_pct * 0.80 + exp_pct * 0.20)
    if final >= 80:
        level, color = 'Strong Match', 'green'
    elif final >= 60:
        level, color = 'Good Match', 'blue'
    elif final >= 40:
        level, color = 'Partial Match', 'orange'
    else:
        level, color = 'Low Match', 'red'
    return {
        'match_percentage': round(final),
        'skills_match_percentage': round(skills_pct),
        'experience_match_percentage': round(exp_pct),
        'matched_skills': matched,
        'missing_skills': missing,
        'total_jd_requirements': len(items),
        'matched_count': len(matched),
        'missing_count': len(missing),
        'match_level': level,
        'match_level_color': color,
        'strong_areas': matched[:12],
        'weak_areas': missing[:12],
        'experience_note': exp_note,
        'jd_years_required': jd_yrs,
        'resume_years_estimated': res_yrs,
    }


# ── LLM-as-judge hybrid assessment ────────────────────────────────────────────
# calculate_match_score() is a literal keyword/skill checklist — a resume that
# says "CI/CD pipeline automation" gets no credit against a JD line item of
# "Jenkins" even though a recruiter would call that a plausible match. The
# functions below add a real LLM judgment on top of (not instead of) that
# keyword scan: the keyword findings are fed into the prompt as grounding
# facts, and the LLM is asked to reason about what the literal scan misses
# (adjacent tools, transferable skills, seniority) and explain its verdict in
# plain language, with its own confidence.

# ── Weighted category scoring model ───────────────────────────────────────────
# The AI Holistic Assessment score is a deterministic weighted sum of 9 named
# categories, each judged by the LLM using contextual/semantic understanding
# (not literal keyword matching — that's what calculate_match_score() already
# does, and its findings are fed into the prompt below as grounding facts for
# the Skills/Tools/Experience categories specifically). The weighting itself
# happens in Python, never left to the LLM to compute, so it's auditable and
# never varies. Weights sum to 1.0 — see _CATEGORY_WEIGHTS below.
_CATEGORY_WEIGHTS = {
    "skills":           0.22,  # Required + preferred skills
    "responsibilities": 0.13,  # Roles & responsibilities alignment
    "experience":       0.15,  # Years of experience + relevance of that experience
    "domain":           0.10,  # Domain / industry experience
    "certifications":   0.08,
    "education":        0.07,
    "regulatory":       0.08,  # Regulatory / standards knowledge (ISO, FDA, GDPR, ...)
    "tools":            0.10,  # Tools & technologies
    "soft_skills":      0.07,
}
assert abs(sum(_CATEGORY_WEIGHTS.values()) - 1.0) < 1e-9, "category weights must sum to 100%"

# The exact example values a prior prompt version put in front of the model
# (see _llm_judge_match) — phi4-mini was found copying these back verbatim
# regardless of resume/JD content, which is why every ai_match_cache row
# converged on ~69%. Kept as a tripwire even after removing that example
# from the prompt, in case a future prompt edit reintroduces the pattern.
_KNOWN_POISONED_EXAMPLE_SCORES = {
    "skills": 75, "responsibilities": 60, "experience": 80, "domain": 50,
    "certifications": 40, "education": 90, "regulatory": 100, "tools": 65,
    "soft_skills": 50,
}

_CATEGORY_LABELS = {
    "skills": "Skills", "responsibilities": "Responsibilities", "experience": "Experience",
    "domain": "Domain Experience", "certifications": "Certifications", "education": "Education",
    "regulatory": "Regulatory/Standards", "tools": "Tools & Technologies", "soft_skills": "Soft Skills",
}

_CONFIDENCE_AGREEMENT_WEIGHT = 0.45
_CONFIDENCE_SUFFICIENCY_WEIGHT = 0.20
_CONFIDENCE_SELF_WEIGHT = 0.35
# Bumped whenever the scoring formula/prompt shape changes — every cached
# ai_match_cache row is keyed to this version (see _read_cached_hybrid_match),
# so a version bump makes every existing cache entry a miss and forces a
# fresh, correctly-computed re-assessment instead of serving an old model's
# number under the new UI.
_HYBRID_ALGO_VERSION = "hybrid-v8-authenticity-calibration"


def _resume_match_fingerprint(resume_dict):
    """Hash of only the resume fields that actually feed AI/keyword matching
    (see calculate_match_score's include_title=False and _llm_judge_match,
    which both deliberately exclude Title/Role & Group from the AI judgment).
    _read_cached_hybrid_match compares this instead of the raw
    resume.updated_at timestamp, so editing purely organizational metadata
    (Role & Group, location, contact info) can never invalidate an
    otherwise-identical cached AI assessment — confirmed live 2026-08-04 that
    a metadata-only edit was forcing a full, wasteful recompute that (before
    Title was excluded above) even produced a visibly different percentage.
    """
    fields = ("summary", "skills", "experience", "projects", "certifications", "education", "exp_yrs")
    raw = "\x1f".join(str(resume_dict.get(k) or "") for k in fields)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _tier_from_pct(pct):
    """Single source of truth for percentage → (verdict, tier_label,
    recommendation). Every scorer in the app — the keyword+LLM hybrid
    assessment, the whole-document holistic fallback, and the Top 3 Matching
    Roles ranking — buckets through this one function, so the same
    percentage always reads the same tier everywhere. These used to be three
    separate copies of this if/elif chain at different cutoffs (70/45/25 in
    two places, 80/60/40 in the third), so the exact same score could show
    "Good Fit" in one component and "Partial Fit" in another for the same
    resume/JD pair — the score-inconsistency bug this fixes.
    """
    if pct >= 80:
        return ("Strong overall fit", "Strong Fit",
                "🌟 Great fit for this role — advance to next round")
    elif pct >= 60:
        return ("Good overall fit", "Good Fit",
                "👍 Good fit for this role — worth interviewing")
    elif pct >= 40:
        return ("Partial overall fit", "Partial Fit",
                "🤔 Partial fit — worth a closer look")
    else:
        return ("Limited overall fit", "Weak Fit",
                "⏭ Not a strong fit for this role")


def _clamp_pct(value, default=50):
    try:
        return max(0, min(100, int(round(float(value)))))
    except (TypeError, ValueError):
        return default


def _unwrap_category_score(value):
    """Defensive unwrap for a small-model quirk seen in practice: a category
    score occasionally comes back wrapped as {"description text": 60}
    instead of plain 60, when the model echoes the category's descriptive
    hint back as a nested key. If given a dict, pull out its first value;
    otherwise pass the value through unchanged for _clamp_pct to handle.
    """
    if isinstance(value, dict) and value:
        return next(iter(value.values()))
    return value


def _coerce_str_list(value, limit):
    if not isinstance(value, list):
        return []
    out = []
    for item in value:
        s = str(item).strip()
        if s:
            out.append(s[:80])
        if len(out) >= limit:
            break
    return out


def _coerce_interview_questions(raw, category):
    """Validate/coerce one category's raw Ollama response into a list of
    plain question dicts, dropping any individual malformed entry rather
    than failing the whole category — same graceful-degradation spirit as
    the rest of this codebase's LLM-output handling (see _coerce_str_list).
    `category` is trusted (comes from our own per-category prompt loop, not
    the model), so it isn't re-validated here.
    """
    if not isinstance(raw, dict):
        return []
    items = raw.get("questions")
    if not isinstance(items, list):
        return []

    out = []
    for item in items:
        if not isinstance(item, dict):
            continue
        question = str(item.get("question") or "").strip()[:400]
        if not question:
            continue
        difficulty = str(item.get("difficulty") or "").strip().title()
        if difficulty not in INTERVIEW_QUESTION_DIFFICULTIES:
            difficulty = "Medium"
        out.append({
            "category": category,
            "question": question,
            "expected_answer": str(item.get("expected_answer") or "").strip()[:600],
            "key_points": "\n".join(_coerce_str_list(item.get("key_points"), 6)),
            "difficulty": difficulty,
        })
    return out


def _llm_judge_match(resume_dict, jd_dict, keyword_result):
    """Asks the local Ollama model to judge JD/resume fit as a human recruiter
    would — but instead of one opaque "fit_percentage", it scores each of the
    9 named categories in _CATEGORY_WEIGHTS independently (0-100), using
    contextual/semantic judgment rather than literal keyword matching (that's
    what calculate_match_score() already does; its matched/missing findings
    are fed in below as grounding facts for the model's skills/tools/
    experience judgment, not as something it re-derives from scratch). The
    weighted sum that turns these into one final percentage happens entirely
    in Python (_hybrid_match), never left to the LLM, so the weighting is
    auditable and never drifts.

    Returns None (never raises) if the model is unreachable, times out, or
    returns something we can't make sense of — callers fall back to the
    pure-heuristic score in that case, same as every other Ollama caller here.
    """
    jd_text = "\n".join(f"{label}: {str(jd_dict.get(key) or '')[:1500]}" for label, key in [
        ("Title", "title"), ("Role", "role"), ("Skills", "skills"),
        ("Requirements", "requirements"), ("Responsibilities", "responsibilities"),
        ("Keywords", "keywords"),
    ])
    # Title deliberately excluded — it's the Role & Group field, pure
    # organizational metadata the UI lets a user change independent of the
    # actual resume content. Including it here made editing only Role/Group
    # change the AI-judged percentage for an unrelated JD (confirmed live
    # 2026-08-04: same skills/experience, only Title edited, same JD's score
    # moved from 58% to 50%) — the deterministic LLM call was correctly
    # reflecting a real prompt-input change, but from the user's perspective
    # metadata-only edits shouldn't move a score at all. calculate_match_score
    # (below, include_title=False) is likewise kept title-free so the
    # grounding facts handed to the LLM can't leak title influence either.
    resume_text = "\n".join(f"{label}: {str(resume_dict.get(key) or '')[:1500]}" for label, key in [
        ("Summary", "summary"), ("Skills", "skills"),
        ("Experience", "experience"), ("Projects", "projects"),
        ("Certifications", "certifications"), ("Education", "education"),
    ])
    matched = keyword_result.get("matched_skills") or []
    missing = keyword_result.get("missing_skills") or []
    jd_yrs = keyword_result.get("jd_years_required")
    res_yrs = keyword_result.get("resume_years_estimated")
    exp_fact = (
        f"JD requires {jd_yrs}+ year(s); resume shows ~{res_yrs} year(s)." if jd_yrs and res_yrs is not None
        else f"JD requires {jd_yrs}+ year(s); resume experience timeline unclear." if jd_yrs
        else "No specific years-of-experience requirement stated in the JD."
    )

    # The category MEANINGS are given as a plain bulleted list, kept fully
    # separate from the JSON template below. The template uses bare "<int>"
    # placeholders (no concrete numbers, no descriptive text inline with a
    # key) rather than either of the two formats tried previously:
    #   - abstract "<0-100 int, description>" hints inline with the key —
    #     phi4-mini would echo the description text back as a nested object,
    #     e.g. {"skills": {"required+preferred skills coverage": 60}}.
    #   - a concrete, fully-filled-in EXAMPLE with plain numbers — confirmed
    #     live (2026-08-03) that phi4-mini copies those exact numbers back
    #     verbatim regardless of resume/JD content: a strong Python-backend
    #     match and a completely unrelated graphic-designer resume against
    #     the same JD both returned the identical example scores, which is
    #     why every single assessment in ai_match_cache converged on ~69%.
    # A bare "<int>" placeholder gives the model the JSON shape without
    # anything it can plausibly parrot, and was verified live to produce
    # properly differentiated scores across strong/weak/mid-match resumes
    # (including empty and garbage-text edge cases, which still fail closed
    # to plain integers rather than nested objects).
    prompt = (
        "You are a technical recruiter judging how well a candidate's resume fits a job description. "
        "Score these 9 categories independently, each 0-100, using your own contextual/semantic "
        "judgment based ONLY on the specific JOB DESCRIPTION and RESUME given below — credit adjacent "
        "tools, transferable skills, and equivalent experience the exact wording wouldn't literally "
        "match, not just verbatim overlap:\n"
        "- skills: required + preferred skills coverage\n"
        "- responsibilities: how well past roles map to this JD's responsibilities\n"
        "- experience: years AND relevance of that experience\n"
        "- domain: domain/industry experience overlap\n"
        "- certifications: relevant certifications held vs wanted\n"
        "- education: education requirement fit\n"
        "- regulatory: regulatory/standards knowledge (e.g. ISO/FDA/GDPR); use 100 if none is required\n"
        "- tools: tools & technologies overlap\n"
        "- soft_skills: communication/leadership/collaboration signals; use 50 if no signal either way\n\n"
        "Additionally run 7 CONSISTENCY CHECKS on the resume — internal consistency only, NOT "
        "whether the resume was written or polished with AI assistance. CRITICAL: these checks are "
        "about whether the resume is coherent with ITSELF, NOT about whether the candidate meets "
        "THIS JD's requirements — that comparison is already scored separately above. A candidate "
        "having fewer years than the JD wants, working in a different domain than this JD, or "
        "lacking a skill this JD asks for, is a FIT gap, not an inconsistency — never fail a check "
        "for that reason. Only fail a check if the resume contradicts itself or claims something "
        "implausible on its own terms, regardless of this specific JD:\n"
        "- skills_match_projects: are the skills claimed actually demonstrated somewhere in the "
        "candidate's OWN project/experience descriptions, not just listed by name?\n"
        "- experience_matches_responsibilities: do the responsibilities described for a past role "
        "plausibly fit the seniority/duration the candidate claims for THAT role — judged against "
        "the candidate's own claims, not this JD's bar?\n"
        "- timeline_consistent: do the candidate's own date ranges make sense (no overlaps, gaps, "
        "or impossible chronology)?\n"
        "- designation_matches_experience: does the seniority of the candidate's job title fit the "
        "candidate's own total years of experience?\n"
        "- technologies_demonstrated: do the tools/technologies the candidate lists under skills "
        "show up anywhere in their own work experience or projects?\n"
        "- projects_realistic: do the candidate's project descriptions read as technically "
        "plausible work, not generic filler that could apply to any candidate — judge this on its "
        "own merits, not against this JD's domain?\n"
        "- no_major_contradictions: is the resume free of direct self-contradictions (e.g. "
        "conflicting titles, durations, or claims elsewhere in the same document)?\n"
        "For each, answer pass=true only if the resume genuinely supports it, pass=false if it "
        "doesn't, with one short reason each. Do not penalize a check just because the resume is "
        "well-written, or because the candidate is a partial fit for this JD — judge only whether "
        "the resume's own specifics line up with each other.\n\n"
        f"KEYWORD SCAN (grounding facts, not the final answer): {len(matched)}/{len(matched) + len(missing)} "
        f"JD skill/requirement items matched verbatim ({keyword_result.get('match_percentage', 0)}%).\n"
        f"Matched: {', '.join(matched[:20]) or 'none'}\n"
        f"Missing: {', '.join(missing[:20]) or 'none'}\n"
        f"EXPERIENCE FACT: {exp_fact}\n\n"
        f"JOB DESCRIPTION:\n{jd_text}\n\n"
        f"RESUME:\n{resume_text}\n\n"
        "Respond with ONLY a single JSON object, no other text before or after it, in EXACTLY this "
        "shape — every category_scores value must be a plain integer from 0 to 100 that YOU compute "
        "specifically for the resume and JD above, never an object, a string, or a placeholder:\n"
        '{"category_scores": {"skills": <int>, "responsibilities": <int>, "experience": <int>, '
        '"domain": <int>, "certifications": <int>, "education": <int>, "regulatory": <int>, '
        '"tools": <int>, "soft_skills": <int>}, '
        '"verdict": "<Strong Match|Good Match|Partial Match|Low Match>", '
        '"rationale": "<2-4 sentences grounded in specific resume/JD content explaining the verdict>", '
        '"strengths": ["<short phrase>", ...up to 6], '
        '"concerns": "<1-2 sentences on the biggest gaps, or empty string if none>", '
        '"suggested_roles": ["<role name>", ...up to 3, or empty list], '
        '"confidence": <int 0-100, how confident you are in this judgment given the information available>, '
        '"confidence_reason": "<one short sentence>", '
        '"consistency_checks": {'
        '"skills_match_projects": {"pass": <bool>, "reason": "<short phrase>"}, '
        '"experience_matches_responsibilities": {"pass": <bool>, "reason": "<short phrase>"}, '
        '"timeline_consistent": {"pass": <bool>, "reason": "<short phrase>"}, '
        '"designation_matches_experience": {"pass": <bool>, "reason": "<short phrase>"}, '
        '"technologies_demonstrated": {"pass": <bool>, "reason": "<short phrase>"}, '
        '"projects_realistic": {"pass": <bool>, "reason": "<short phrase>"}, '
        '"no_major_contradictions": {"pass": <bool>, "reason": "<short phrase>"}}}'
    )

    try:
        # On CPU-only Ollama, generation time scales with num_predict — 850
        # output tokens covers 9 category scores, 7 consistency-check
        # objects, a short rationale, and a handful of short phrases with
        # headroom, without the runaway cost of an unbounded cap. This now
        # runs at most once per resume/JD pair ever (cached + advisory-
        # locked, see _get_or_compute_hybrid_match), not repeatedly per page
        # view, so this cap is about output quality, not working around
        # being called repeatedly.
        raw = _ollama_chat(prompt, as_json=True, num_predict=850, num_ctx=8192)
    except Exception as e:
        logger.warning(f"LLM judge call failed: {e}", exc_info=True)
        return None

    if not isinstance(raw, dict) or not isinstance(raw.get("category_scores"), dict):
        logger.warning(f"LLM judge returned unusable response: {str(raw)[:200]}")
        return None

    raw_categories = raw["category_scores"]
    category_scores = {
        cat: _clamp_pct(_unwrap_category_score(raw_categories.get(cat)), default=None)
        for cat in _CATEGORY_WEIGHTS
    }
    missing_categories = [cat for cat, val in category_scores.items() if val is None]
    if missing_categories:
        logger.warning(f"LLM judge omitted categories {missing_categories}, response: {str(raw)[:200]}")
        return None

    if category_scores == _KNOWN_POISONED_EXAMPLE_SCORES:
        # Defense-in-depth against the exact regression fixed on 2026-08-03:
        # the model parroting a literal prompt example back verbatim instead
        # of judging the actual resume/JD. Odds of a genuine, independently
        # -computed judgment landing on this exact 9-value combination are
        # negligible, so treat it as a copy and fall back to the heuristic
        # rather than caching a fabricated score.
        logger.warning(f"LLM judge returned the known-poisoned example verbatim — treating as invalid, response: {str(raw)[:200]}")
        return None

    # consistency_checks is deliberately NOT required the way the 9 category
    # scores are — it's a supplementary signal (see _evaluate_authenticity),
    # not part of the auditable weighted score. If the model omits it, or
    # malforms individual keys, the real match assessment must still stand;
    # _evaluate_authenticity degrades each check to "not assessed"
    # individually rather than invalidating the whole judgment the way a
    # missing category score does.
    raw_checks = raw.get("consistency_checks")
    consistency_checks = {}
    if isinstance(raw_checks, dict):
        for key in _AUTHENTICITY_SOFT_RULE_KEYS:
            entry = raw_checks.get(key)
            if isinstance(entry, dict) and isinstance(entry.get("pass"), bool):
                consistency_checks[key] = {
                    "pass": entry["pass"],
                    "reason": str(entry.get("reason") or "").strip()[:120],
                }

    return {
        "category_scores": category_scores,
        "verdict": str(raw.get("verdict") or "").strip()[:100],
        "rationale": str(raw.get("rationale") or "").strip()[:600],
        "strengths": _coerce_str_list(raw.get("strengths"), 6),
        "concerns": str(raw.get("concerns") or "").strip()[:400],
        "suggested_roles": _coerce_str_list(raw.get("suggested_roles"), 3),
        "confidence": _clamp_pct(raw.get("confidence")),
        "confidence_reason": str(raw.get("confidence_reason") or "").strip()[:200],
        "consistency_checks": consistency_checks,
    }


def _hybrid_match(resume_dict, jd_dict):
    """The primary "AI Assessment": keyword checklist + LLM-as-judge, blended.
    Never raises and never returns None — falls back to the pure-heuristic
    _holistic_or_default() (with a Low confidence badge) if the LLM judge is
    unavailable, so a down/slow Ollama server degrades the UX rather than
    breaking it.

    The returned dict always carries `is_ai_judged`: True only when the real
    LLM judge actually ran, False for the heuristic fallback. Callers that
    cache this result (see _store_hybrid_match) must only persist it when
    is_ai_judged is True — caching a fallback would otherwise permanently
    freeze that resume/JD pair at a low-confidence keyword-only estimate even
    after Ollama comes back, since a cache hit is never re-checked against a
    live LLM again.
    """
    # include_title=False: this keyword scan feeds the LLM's grounding facts
    # (matched/missing lists), so it must exclude Title too — otherwise a
    # Role/Group-only edit could still shift the AI's judgment indirectly
    # through those grounding facts even with Title removed from resume_text
    # above.
    keyword_result = calculate_match_score(resume_dict, jd_dict, include_title=False)
    llm_result = _llm_judge_match(resume_dict, jd_dict, keyword_result)
    # Pure string comparison, no LLM involved — computed regardless of
    # whether the AI judge succeeded, so the fallback path (Ollama down)
    # still surfaces this flag instead of losing it entirely.
    jd_mirroring = _jd_mirroring_risk(resume_dict, jd_dict)

    if llm_result is None:
        fallback = _holistic_or_default(resume_dict, jd_dict)
        fallback["keyword_match_percentage"] = keyword_result.get("match_percentage", 0)
        fallback["category_scores"] = None
        fallback["confidence"] = 25
        fallback["confidence_label"] = "Low"
        fallback["confidence_reason"] = "AI judge unavailable — showing keyword + heuristic estimate only."
        fallback["algo_version"] = _HYBRID_ALGO_VERSION
        fallback["is_ai_judged"] = False
        # No LLM ran, so there's no basis for an authenticity judgment —
        # left as "not assessed" rather than guessed at.
        fallback["authenticity_status"] = None
        fallback["authenticity_reasons_positive"] = []
        fallback["authenticity_reasons_negative"] = []
        fallback["authenticity_explanation"] = ""
        fallback["authenticity_rules"] = {}
        fallback["jd_mirroring_pct"] = jd_mirroring["overlap_pct"]
        fallback["jd_mirroring_label"] = jd_mirroring["risk_label"]
        fallback["jd_mirroring_phrases"] = jd_mirroring["matched_phrases"]
        return fallback

    keyword_pct = keyword_result.get("match_percentage", 0)
    category_scores = llm_result["category_scores"]
    # The ONLY place the final percentage is computed — a deterministic
    # weighted sum of the 9 category scores, weights defined once in
    # _CATEGORY_WEIGHTS. The LLM never computes or reports this number
    # itself, so it's fully auditable and reproducible from the stored
    # category_scores alone.
    final_pct = round(sum(category_scores[cat] * weight for cat, weight in _CATEGORY_WEIGHTS.items()))
    verdict, tier_label, recommendation = _tier_from_pct(final_pct)

    total_items = keyword_result.get("total_jd_requirements", 0)
    resume_word_count = len(re.findall(r"\w+", " ".join(
        str(resume_dict.get(k) or '') for k in ("skills", "experience", "summary"))))
    if total_items >= 5 and resume_word_count >= 80:
        data_sufficiency = 100
    elif total_items >= 2 and resume_word_count >= 30:
        data_sufficiency = 60
    else:
        data_sufficiency = 30

    # Agreement between the literal keyword scan and the semantic weighted
    # score — a big gap between "what's written verbatim" and "what the
    # category judgment concluded" is itself a signal the assessment
    # deserves a closer human look, not just a raw percentage.
    agreement = 100 - abs(keyword_pct - final_pct)
    confidence_pct = round(
        agreement * _CONFIDENCE_AGREEMENT_WEIGHT
        + data_sufficiency * _CONFIDENCE_SUFFICIENCY_WEIGHT
        + llm_result["confidence"] * _CONFIDENCE_SELF_WEIGHT
    )
    confidence_label = "High" if confidence_pct >= 75 else "Medium" if confidence_pct >= 50 else "Low"

    # Binary resume-authenticity signal — see _evaluate_authenticity /
    # _AUTHENTICITY_RULES. consistency_checks is always a dict here (never
    # None) since llm_result exists on this path; _evaluate_authenticity
    # only returns None when the LLM didn't run at all (the fallback path
    # above).
    authenticity = _evaluate_authenticity(resume_dict, llm_result.get("consistency_checks", {}))
    logger.info(
        f"Authenticity status={authenticity['status']} "
        f"failed_rules={[k for k, v in authenticity['rules'].items() if v['pass'] is False]}"
    )

    return {
        "fit_percentage": final_pct,
        "verdict": verdict,
        "tier_label": tier_label,
        "recommendation": recommendation,
        "rationale": llm_result["rationale"],
        "strengths": llm_result["strengths"],
        "concerns": llm_result["concerns"],
        "suggested_roles": llm_result["suggested_roles"],
        "keyword_match_percentage": keyword_pct,
        "category_scores": category_scores,
        "confidence": confidence_pct,
        "confidence_label": confidence_label,
        "confidence_reason": llm_result["confidence_reason"] or (
            "Keyword scan and category judgment agree closely." if agreement >= 80
            else "Keyword scan and category judgment differ somewhat."
        ),
        "algo_version": _HYBRID_ALGO_VERSION,
        "is_ai_judged": True,
        # Binary resume-authenticity signal — see _evaluate_authenticity.
        # Deliberately kept out of fit_percentage/category_scores: a
        # flagged resume still gets an honest skill-match score, this is a
        # separate signal for a recruiter to verify in interview, not a
        # score penalty.
        "authenticity_status": authenticity["status"],
        "authenticity_reasons_positive": authenticity["reasons_positive"],
        "authenticity_reasons_negative": authenticity["reasons_negative"],
        "authenticity_explanation": authenticity["explanation"],
        "authenticity_rules": authenticity["rules"],
        # Second, independent authenticity signal — see _jd_mirroring_risk.
        # Not LLM-derived, so it's available even when authenticity isn't.
        "jd_mirroring_pct": jd_mirroring["overlap_pct"],
        "jd_mirroring_label": jd_mirroring["risk_label"],
        "jd_mirroring_phrases": jd_mirroring["matched_phrases"],
    }


def _read_cached_hybrid_match(conn, resume, jd):
    """Read-only lookup into ai_match_cache. Returns the cached hybrid
    assessment dict if a genuine AI-judged result exists and is still fresh
    (neither side edited since it was computed, same algo version) — else
    None. Never computes or blocks on Ollama; callers that need a result even
    when nothing is cached should use _get_or_compute_hybrid_match() instead.

    This is the single place that decides whether a cached row counts —
    every reader (the compare page, the Top 3 Matching Roles cards) goes
    through it so they can never disagree about whether a pair has already
    been AI-assessed.
    """
    ensure_ai_match_cache_table(conn)
    cached = conn.execute(
        "SELECT * FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
        (resume["id"], jd["id"]),
    ).fetchone()
    if not cached:
        return None

    # resume.updated_at is TIMESTAMPTZ (tz-aware) but job_description.updated_at
    # is a naive TIMESTAMP — normalize both sides to naive before comparing so
    # this doesn't depend on which of the two inconsistent column types a given
    # value happened to come from.
    _naive = lambda dt: dt.replace(tzinfo=None) if dt and dt.tzinfo else dt
    extra = cached["extra"] or {}
    if not (
        # Content fingerprint, not the raw resume.updated_at timestamp — a
        # Save that only touches Role & Group (or location/contact info)
        # still bumps updated_at but must NOT invalidate an otherwise
        # identical cached AI assessment (see _resume_match_fingerprint).
        extra.get("resume_match_fingerprint") == _resume_match_fingerprint(resume)
        and _naive(cached["jd_updated_at"]) == _naive(jd["updated_at"])
        and extra.get("algo_version") == _HYBRID_ALGO_VERSION
        # Cache rows written before is_ai_judged existed (or a fallback that
        # slipped through pre-fix) don't carry `is_ai_judged: true` — treat
        # them as a miss so they self-heal to a real AI judgment on next view
        # instead of serving a stale keyword-only estimate forever.
        and extra.get("is_ai_judged")
    ):
        return None

    jd_mirroring = _jd_mirroring_risk(resume, jd)

    return {
        "fit_percentage": cached["fit_percentage"],
        "verdict": cached["verdict"],
        "tier_label": extra.get("tier_label", cached["verdict"]),
        "recommendation": extra.get("recommendation", ""),
        "rationale": cached["rationale"],
        "strengths": extra.get("strengths", []),
        "concerns": extra.get("concerns", ""),
        "suggested_roles": extra.get("suggested_roles", []),
        "keyword_match_percentage": extra.get("keyword_match_percentage"),
        "category_scores": extra.get("category_scores"),
        "category_weights": _CATEGORY_WEIGHTS,
        "confidence": extra.get("confidence"),
        "confidence_label": extra.get("confidence_label", ""),
        "confidence_reason": extra.get("confidence_reason", ""),
        "algo_version": _HYBRID_ALGO_VERSION,
        "is_ai_judged": True,
        "authenticity_status": extra.get("authenticity_status"),
        "authenticity_reasons_positive": extra.get("authenticity_reasons_positive", []),
        "authenticity_reasons_negative": extra.get("authenticity_reasons_negative", []),
        "authenticity_explanation": extra.get("authenticity_explanation", ""),
        "authenticity_rules": extra.get("authenticity_rules", {}),
        # Computed fresh on every read, not persisted in `extra` — pure
        # string comparison over the current resume/JD text, so it never
        # needs cache-invalidation logic and always reflects live content.
        "jd_mirroring_pct": jd_mirroring["overlap_pct"],
        "jd_mirroring_label": jd_mirroring["risk_label"],
        "jd_mirroring_phrases": jd_mirroring["matched_phrases"],
    }


def _store_hybrid_match(conn, resume, jd, assessment):
    """Persist a genuinely AI-judged hybrid assessment so every other view of
    this resume/JD pair reuses it instead of recomputing. Only ever called
    with assessment['is_ai_judged'] True — a fallback (LLM unavailable or
    timed out) is intentionally never written here, see _hybrid_match.
    """
    extra_json = json.dumps({
        "tier_label": assessment.get("tier_label", assessment["verdict"]),
        "recommendation": assessment.get("recommendation", ""),
        "strengths": assessment.get("strengths", []),
        "concerns": assessment.get("concerns", ""),
        "suggested_roles": assessment.get("suggested_roles", []),
        "keyword_match_percentage": assessment.get("keyword_match_percentage"),
        "category_scores": assessment.get("category_scores"),
        "confidence": assessment.get("confidence"),
        "confidence_label": assessment.get("confidence_label", ""),
        "confidence_reason": assessment.get("confidence_reason", ""),
        "algo_version": assessment.get("algo_version", _HYBRID_ALGO_VERSION),
        "is_ai_judged": True,
        "resume_match_fingerprint": _resume_match_fingerprint(resume),
        "authenticity_status": assessment.get("authenticity_status"),
        "authenticity_reasons_positive": assessment.get("authenticity_reasons_positive", []),
        "authenticity_reasons_negative": assessment.get("authenticity_reasons_negative", []),
        "authenticity_explanation": assessment.get("authenticity_explanation", ""),
        "authenticity_rules": assessment.get("authenticity_rules", {}),
    })
    conn.execute(
        """
        INSERT INTO ai_match_cache
            (resume_id, jd_id, fit_percentage, verdict, rationale, extra,
             resume_updated_at, jd_updated_at, computed_at)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
        ON CONFLICT (resume_id, jd_id) DO UPDATE SET
            fit_percentage = EXCLUDED.fit_percentage,
            verdict = EXCLUDED.verdict,
            rationale = EXCLUDED.rationale,
            extra = EXCLUDED.extra,
            resume_updated_at = EXCLUDED.resume_updated_at,
            jd_updated_at = EXCLUDED.jd_updated_at,
            computed_at = NOW()
        """,
        (resume["id"], jd["id"], assessment["fit_percentage"], assessment["verdict"],
         assessment["rationale"], extra_json, resume["updated_at"], jd["updated_at"]),
    )


def _get_or_compute_hybrid_match(conn, resume, jd):
    """Single entry point for "the AI Holistic Assessment" of one resume/JD
    pair. Every caller — the compare page's async fetch, the Top 3 Matching
    Roles cards' background upgrade — goes through this so they can never
    independently compute or display a different number for the same pair:
    reuse the cached genuine AI judgment when one exists and nothing has
    changed since; otherwise run the real LLM judge now and cache the
    result, unless the LLM was unreachable (see _hybrid_match /
    _store_hybrid_match).

    Serialized per (resume_id, jd_id) with a Postgres advisory lock — not an
    in-process Python lock — because this app can end up with more than one
    server process bound to the same port (e.g. a stale process left running
    from a previous launch alongside a freshly-restarted one; Windows will
    happily let both bind). Two such processes racing to compute the same
    uncached pair would each call the LLM independently and whichever
    finished last would silently overwrite the other's cache row — which is
    exactly what made the displayed percentage flip between refreshes. A
    Postgres advisory lock is enforced by the database itself, so it
    serializes correctly across processes, not just threads within one.
    pg_advisory_lock is session-scoped: it releases automatically if the
    connection dies before the `finally` runs, so a crashed request can
    never leave the pair permanently locked.
    """
    cached = _read_cached_hybrid_match(conn, resume, jd)
    if cached is not None:
        logger.info(f"AI holistic assessment cache hit → resume={resume['id']} jd={jd['id']} pct={cached['fit_percentage']}")
        return cached

    resume_id, jd_id = resume["id"], jd["id"]
    conn.execute("SELECT pg_advisory_lock(%s, %s)", (resume_id, jd_id))
    try:
        # Re-check: another process/thread may have just finished computing
        # this exact pair while we were waiting for the lock.
        cached = _read_cached_hybrid_match(conn, resume, jd)
        if cached is not None:
            logger.info(f"AI holistic assessment cache hit after lock wait → resume={resume_id} jd={jd_id} pct={cached['fit_percentage']}")
            return cached

        logger.info(f"AI holistic assessment cache miss → resume={resume_id} jd={jd_id}, computing")
        assessment = _hybrid_match(dict(resume), dict(jd))
        if assessment.get("is_ai_judged"):
            _store_hybrid_match(conn, resume, jd, assessment)
            logger.info(f"AI holistic assessment computed → resume={resume_id} jd={jd_id} pct={assessment['fit_percentage']}")
        else:
            logger.warning(f"AI judge unavailable, not caching fallback → resume={resume_id} jd={jd_id}")
        return assessment
    finally:
        conn.execute("SELECT pg_advisory_unlock(%s, %s)", (resume_id, jd_id))


# ── Background AI analysis ────────────────────────────────────────────────────
# Every place that used to make the user's own HTTP request block on a real
# LLM call (up to a couple of minutes on local hardware) now instead kicks off
# a detached background thread and returns immediately — the frontend polls
# a cheap, read-only status endpoint (see compare_ai_status()) until the
# result lands in ai_match_cache, instead of holding one HTTP request open
# and giving up with "unavailable" the moment a client-side timeout fires
# while the computation is still healthy and still running server-side.
_ai_assessment_in_flight = set()
_ai_assessment_in_flight_lock = threading.Lock()


def _background_ai_assessment_worker(resume_id, jd_ids):
    for jd_id in jd_ids:
        try:
            with db_conn() as conn:
                resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
                jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
                if not resume or not jd:
                    continue
                _get_or_compute_hybrid_match(conn, resume, jd)
        except Exception:
            logger.exception(f"Background AI assessment failed → resume={resume_id} jd={jd_id}")
        finally:
            with _ai_assessment_in_flight_lock:
                _ai_assessment_in_flight.discard((resume_id, jd_id))


def trigger_background_ai_assessment(resume_id, jd_ids):
    """Fire-and-forget: start computing the real AI Holistic Assessment for
    the given (resume_id, jd_id) pairs in a background thread, so analysis
    is already under way — or done — by the time anyone's browser asks for
    it, instead of only starting when a page view's fetch request arrives.

    Safe to call repeatedly for the same pair: an in-process guard skips
    spawning a duplicate thread for a pair that's already being computed
    (avoids piling up redundant background threads each idly blocked on the
    same Postgres advisory lock), and _get_or_compute_hybrid_match's
    advisory lock protects correctness even across separate server
    processes regardless.
    """
    to_launch = []
    with _ai_assessment_in_flight_lock:
        for jd_id in jd_ids:
            key = (resume_id, jd_id)
            if key not in _ai_assessment_in_flight:
                _ai_assessment_in_flight.add(key)
                to_launch.append(jd_id)
    if not to_launch:
        return
    logger.info(f"Triggering background AI assessment → resume={resume_id} jds={to_launch}")
    threading.Thread(
        target=_background_ai_assessment_worker, args=(resume_id, to_launch), daemon=True
    ).start()


def _top_n_jd_ids_by_estimate(resume_dict, jds, n=3):
    """Fast, LLM-free ranking (see _quick_estimate_jd_match) used only to
    decide which JDs are worth kicking off a background AI assessment for
    right after Save — analysing every Open JD in the background would
    waste CPU on roles nowhere near a realistic top-3.
    """
    scored = []
    for jd in jds:
        jd_dict = dict(jd)
        estimate_pct, _ = _quick_estimate_jd_match(resume_dict, jd_dict)
        scored.append((estimate_pct, jd_dict["id"]))
    scored.sort(key=lambda t: -t[0])
    return [jd_id for _, jd_id in scored[:n]]


# ── JD Routes ─────────────────────────────────────────────────────────────────

@app.route("/jd-management")
def jd_management():
    if not _require_permission("view_jd", write=False):
        abort(403)
    status_filter = (request.args.get("status") or "").strip()
    search_query = (request.args.get("q") or "").strip()
    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_requirement_table(conn)
        seed_jds(conn)
        conditions, params = [], []
        if status_filter in POSITION_STATUSES:
            conditions.append("position_status = %s")
            params.append(status_filter)
        if search_query:
            # "Keyword" covers the keywords/skills text fields too, since
            # those aren't visible table columns the client-side category/
            # status tabs could ever search — this has to be a real query.
            like = f"%{search_query}%"
            conditions.append("(title ILIKE %s OR keywords ILIKE %s OR skills ILIKE %s OR CAST(id AS TEXT) = %s)")
            params += [like, like, like, search_query]
        where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        jds = conn.execute(
            "SELECT id, title, role, category, position_status, workflow_status, created_at"
            f" FROM job_description {where_sql} ORDER BY category, title",
            params,
        ).fetchall()
        active_counts = conn.execute(
            "SELECT jd_id, COUNT(*) AS cnt FROM requirement "
            "WHERE jd_id IS NOT NULL AND status != 'Closed' GROUP BY jd_id"
        ).fetchall()
    active_count_by_jd = {r["jd_id"]: r["cnt"] for r in active_counts}
    jds = [dict(jd, active_requirement_count=active_count_by_jd.get(jd["id"], 0)) for jd in jds]
    return render_template("jd_management.html", jds=jds,
                           statuses=POSITION_STATUSES, status_filter=status_filter,
                           search_query=search_query)


# Every field except the optional file upload is mandatory (2026-08-13:
# explicit follow-up ask — the form previously only required Job Title,
# letting a JD save with everything else blank).
_JD_REQUIRED_FIELDS = [
    ("title", "Job Title"), ("role", "Role"), ("category", "Category"),
    ("responsibilities", "Responsibilities"), ("requirements", "Requirements / Qualifications"),
    ("skills", "Required Skills"), ("keywords", "Keywords"),
]


def _validate_jd_required_fields(data):
    missing = [label for key, label in _JD_REQUIRED_FIELDS if not data.get(key)]
    if missing:
        return f"Please fill in all required fields: {', '.join(missing)}."
    return None


@app.route("/jd/add", methods=["GET", "POST"])
def jd_add():
    if not _require_permission("write_jd"):
        abort(403)
    if request.method == "POST":
        import time as _t
        data = {k: request.form.get(k, "").strip() for k in
                ["title", "role", "category", "responsibilities", "requirements", "skills", "keywords"]}
        data["position_status"] = request.form.get("position_status", "").strip() or DEFAULT_POSITION_STATUS
        if data["position_status"] not in POSITION_STATUSES:
            data["position_status"] = DEFAULT_POSITION_STATUS
        data["jd_file"] = None
        uploaded = request.files.get("jd_file")
        if uploaded and uploaded.filename and allowed_file(uploaded.filename):
            name = secure_filename(uploaded.filename)
            ext = name.rsplit(".", 1)[1].lower()
            save_name = f"{slugify(Path(name).stem)}-{int(_t.time())}.{ext}"
            dest = JD_UPLOAD_FOLDER / save_name
            uploaded.seek(0)
            uploaded.save(str(dest))
            data["jd_file"] = save_name
            if not data["responsibilities"]:
                try:
                    txt = extract_resume_text(dest, ext)
                    data["responsibilities"] = (txt or "")[:4000]
                except Exception:
                    pass
        error = _validate_jd_required_fields(data)
        if error:
            flash(error, "error")
            return render_template("jd_form.html", jd=None, all_roles=ALL_JD_ROLES,
                                   statuses=POSITION_STATUSES, default_status=DEFAULT_POSITION_STATUS), 400

        # Duplicate-JD warning — this app's JD model has no "Department" field,
        # so Category (e.g. "IT Roles"/"Validation Roles") is the closest
        # existing equivalent. Not a hard block: "Add Anyway" (confirm_duplicate=1)
        # bypasses this exact check on the resubmit.
        if request.form.get("confirm_duplicate") != "1":
            with db_conn() as conn:
                ensure_jd_table(conn)
                dup = conn.execute(
                    "SELECT id, title FROM job_description WHERE LOWER(title) = LOWER(%s) AND category = %s LIMIT 1",
                    (data["title"], data["category"]),
                ).fetchone()
            if dup:
                flash(f"A JD titled \"{dup['title']}\" already exists in the {data['category']} category.",
                      "error")
                return render_template("jd_form.html", jd=data, all_roles=ALL_JD_ROLES,
                                       statuses=POSITION_STATUSES, default_status=DEFAULT_POSITION_STATUS,
                                       duplicate_warning=True)

        data["created_by"] = session.get("user_id")
        with db_conn() as conn:
            ensure_jd_table(conn)
            cursor = conn.execute(
                """
                INSERT INTO job_description
                    (title, role, category, responsibilities, requirements, skills, keywords, jd_file,
                     position_status, workflow_status, created_by)
                VALUES
                    (%(title)s, %(role)s, %(category)s, %(responsibilities)s,
                     %(requirements)s, %(skills)s, %(keywords)s, %(jd_file)s, %(position_status)s,
                     'Draft', %(created_by)s)
                RETURNING id
                """,
                data,
            )
            new_jd_id = cursor.fetchone()["id"]
        log_audit("JD Management", "Add", record_id=new_jd_id, record_label=data["title"])
        flash(f"Job Description '{data['title']}' added as a Draft. Submit it for approval when ready.", "success")
        return redirect(url_for("jd_detail", jd_id=new_jd_id))
    return render_template("jd_form.html", jd=None, all_roles=ALL_JD_ROLES,
                           statuses=POSITION_STATUSES, default_status=DEFAULT_POSITION_STATUS)


@app.route("/jd/<int:jd_id>")
def jd_detail(jd_id):
    if not _require_permission("view_jd", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_requirement_table(conn)
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if not jd:
            return "Job Description not found", 404
        resumes = conn.execute(
            "SELECT id, full_name, title FROM resume ORDER BY updated_at DESC LIMIT 30"
        ).fetchall()
        creator = conn.execute(
            "SELECT full_name FROM app_user WHERE id = %s", (jd["created_by"],)
        ).fetchone() if jd["created_by"] else None
        decider = conn.execute(
            "SELECT full_name FROM app_user WHERE id = %s", (jd["decided_by"],)
        ).fetchone() if jd["decided_by"] else None
        linked_requirements = conn.execute(
            "SELECT id, requirement_code, requirement_name, status FROM requirement "
            "WHERE jd_id = %s ORDER BY created_at DESC",
            (jd_id,),
        ).fetchall()
    linked_requirements = list(linked_requirements)
    active_linked = [r for r in linked_requirements if r["status"] != "Closed"]
    if active_linked:
        codes = ", ".join(r["requirement_code"] for r in active_linked)
        delete_confirm_message = (
            f"This JD has {len(active_linked)} active linked requirement(s) ({codes}). "
            "They will be unlinked, not deleted. Continue deleting this JD?"
        )
    else:
        delete_confirm_message = f'Are you sure you want to delete "{jd["title"]}"? This cannot be undone.'
    return render_template(
        "jd_detail.html", jd=dict(jd), resumes=list(resumes),
        creator_name=creator["full_name"] if creator else None,
        decider_name=decider["full_name"] if decider else None,
        linked_requirements=linked_requirements,
        active_linked_requirements=active_linked,
        delete_confirm_message=delete_confirm_message,
    )


@app.route("/jd/<int:jd_id>/download-pdf")
def download_jd_pdf(jd_id):
    """Download JD as PDF."""
    with db_conn() as conn:
        ensure_jd_table(conn)
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()

    if not jd:
        return "Job Description not found", 404

    jd_dict = dict(jd)

    # A Draft/Rejected JD is not yet finalized content — this is the one
    # place a JD's content can actually leave the authenticated app (e.g.
    # emailed to a candidate), so it's blocked here rather than requiring
    # every future export surface to remember the same check.
    if jd_dict.get("workflow_status") in ("Draft", "Rejected"):
        flash("This JD is in Draft/Rejected status and cannot be exported until it's approved.", "error")
        return redirect(url_for("jd_detail", jd_id=jd_id))

    # Generate PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                          topMargin=0.5*inch, bottomMargin=0.5*inch,
                          leftMargin=0.75*inch, rightMargin=0.75*inch)
    story = []
    styles = getSampleStyleSheet()

    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor='#1e293b',
        spaceAfter=4,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    story.append(Paragraph(jd_dict.get('title', 'Job Description'), title_style))

    # Category & Role
    if jd_dict.get('category') or jd_dict.get('role') or jd_dict.get('position_status'):
        meta_parts = []
        if jd_dict.get('category'):
            meta_parts.append(f"Category: {jd_dict['category']}")
        if jd_dict.get('role'):
            meta_parts.append(f"Role: {jd_dict['role']}")
        if jd_dict.get('position_status'):
            meta_parts.append(f"Position Status: {jd_dict['position_status']}")
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=10,
            textColor='#64748b',
            spaceAfter=16,
            alignment=TA_CENTER
        )
        story.append(Paragraph(' | '.join(meta_parts), meta_style))

    # Section styling
    section_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontSize=12,
        textColor='#1e293b',
        spaceAfter=8,
        spaceBefore=8,
        fontName='Helvetica-Bold'
    )

    body_style = ParagraphStyle(
        'Body',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_JUSTIFY,
        spaceAfter=12
    )

    # Responsibilities
    if jd_dict.get('responsibilities'):
        story.append(Paragraph('RESPONSIBILITIES', section_style))
        resp_text = jd_dict['responsibilities'].replace('\n', '<br/>')
        story.append(Paragraph(resp_text, body_style))

    # Requirements
    if jd_dict.get('requirements'):
        story.append(Paragraph('REQUIREMENTS / QUALIFICATIONS', section_style))
        req_text = jd_dict['requirements'].replace('\n', '<br/>')
        story.append(Paragraph(req_text, body_style))

    # Skills
    if jd_dict.get('skills'):
        story.append(Paragraph('REQUIRED SKILLS', section_style))
        skills_list = [s.strip() for s in jd_dict['skills'].split('\n') if s.strip()]
        skills_text = ' • '.join(skills_list)
        story.append(Paragraph(skills_text, body_style))

    # Keywords
    if jd_dict.get('keywords'):
        story.append(Paragraph('KEYWORDS', section_style))
        story.append(Paragraph(jd_dict['keywords'], body_style))

    # Build PDF
    doc.build(story)
    buffer.seek(0)

    filename = f"{jd_dict.get('title', 'JD').replace(' ', '_')}.pdf"
    log_audit("Downloads", "Download", record_id=jd_id, record_label=jd_dict.get('title'))
    return send_file(
        buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )


@app.route("/jd/<int:jd_id>/extract")
def extract_jd_data(jd_id):
    """Extract JD data as JSON."""
    with db_conn() as conn:
        ensure_jd_table(conn)
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()

    if not jd:
        return jsonify({"error": "Job Description not found"}), 404

    jd_dict = dict(jd)
    # Remove internal fields
    jd_dict.pop('id', None)
    jd_dict.pop('created_at', None)
    jd_dict.pop('updated_at', None)
    jd_dict.pop('jd_file', None)

    filename = f"{jd_dict.get('title', 'jd').replace(' ', '_')}_details.json"

    return send_file(
        BytesIO(json.dumps(jd_dict, indent=2).encode()),
        mimetype='application/json',
        as_attachment=True,
        download_name=filename
    )


@app.route("/jd/<int:jd_id>/edit", methods=["GET", "POST"])
def jd_edit(jd_id):
    if not _require_permission("write_jd"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        if request.method == "POST":
            data = {k: request.form.get(k, "").strip() for k in
                    ["title", "role", "category", "responsibilities", "requirements", "skills", "keywords"]}
            data["position_status"] = request.form.get("position_status", "").strip() or DEFAULT_POSITION_STATUS
            if data["position_status"] not in POSITION_STATUSES:
                data["position_status"] = DEFAULT_POSITION_STATUS
            data["id"] = jd_id
            error = _validate_jd_required_fields(data)
            if error:
                flash(error, "error")
                return render_template("jd_form.html", jd=data, all_roles=ALL_JD_ROLES,
                                       statuses=POSITION_STATUSES, default_status=DEFAULT_POSITION_STATUS), 400
            current = conn.execute(
                "SELECT * FROM job_description WHERE id = %s", (jd_id,)
            ).fetchone()
            if current and current["workflow_status"] == "Published":
                _snapshot_jd_version(conn, current, session.get("user_id"))
            conn.execute(
                """
                UPDATE job_description SET
                    title=%(title)s, role=%(role)s, category=%(category)s,
                    responsibilities=%(responsibilities)s, requirements=%(requirements)s,
                    skills=%(skills)s, keywords=%(keywords)s, position_status=%(position_status)s,
                    updated_at=NOW()
                WHERE id=%(id)s
                """,
                data,
            )
            log_audit("JD Management", "Edit", record_id=jd_id, record_label=data["title"])
            flash("Job Description updated.", "success")
            return redirect(url_for("jd_detail", jd_id=jd_id))
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if not jd:
            return "Job Description not found", 404
    return render_template("jd_form.html", jd=dict(jd), all_roles=ALL_JD_ROLES,
                           statuses=POSITION_STATUSES, default_status=DEFAULT_POSITION_STATUS)


@app.route("/jd/<int:jd_id>/delete", methods=["POST"])
def jd_delete(jd_id):
    if not _require_permission("delete_jd"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        row = conn.execute(
            "SELECT jd_file, title FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if row and row["jd_file"]:
            (JD_UPLOAD_FOLDER / (row["jd_file"] or "")).unlink(missing_ok=True)
        conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))
    log_audit("JD Management", "Delete", record_id=jd_id, record_label=row["title"] if row else None)
    flash("Job Description deleted.", "success")
    return redirect(url_for("jd_management"))


def _jd_link(jd_id):
    """Absolute URL to the JD's detail page, used in notification emails so
    the recipient can click straight through to review/act on it (login is
    still required — this is a deep link, not a bypass-auth magic link)."""
    try:
        return url_for("jd_detail", jd_id=jd_id, _external=True)
    except RuntimeError:
        return None  # no active request context (e.g. called from a test/script)


def _notify_jd_approvers(jd_id, jd_title, submitted_by_name):
    """Emails every active Hiring Manager/Admin that a JD needs their
    review — role-based routing, not a per-JD approver assignment."""
    with db_conn() as conn:
        approvers = conn.execute(
            "SELECT email, full_name FROM app_user WHERE role IN ('admin', 'hiring_manager') AND is_active"
        ).fetchall()
    link = _jd_link(jd_id)
    link_line = f"\n\nReview it here: {link}" if link else ""
    for a in approvers:
        if not a["email"]:
            continue
        _send_email(
            a["email"],
            f"JD pending your approval: {jd_title}",
            f"Hello {a['full_name'] or ''},\n\n"
            f"\"{jd_title}\" was submitted by {submitted_by_name or 'a user'} and is now "
            f"pending your approval in JD Management.{link_line}\n\n"
            "Open the link above (log in if prompted), then use Approve, Reject, or "
            "Request Changes on the JD's Approval card.",
        )


def _notify_jd_creator_of_decision(jd_row, decision, comment):
    """Emails the JD's creator when a decision is made. Silently skipped for
    legacy JDs with no created_by (created before this feature existed)."""
    if not jd_row.get("created_by"):
        return
    with db_conn() as conn:
        creator = conn.execute(
            "SELECT email, full_name FROM app_user WHERE id = %s", (jd_row["created_by"],)
        ).fetchone()
    if not creator or not creator["email"]:
        return
    comment_line = f"\n\nComment: {comment}" if comment else ""
    link = _jd_link(jd_row["id"])
    link_line = f"\n\nView it here: {link}" if link else ""
    _send_email(
        creator["email"],
        f"Your JD \"{jd_row['title']}\" was {decision}",
        f"Hello {creator['full_name'] or ''},\n\n"
        f"\"{jd_row['title']}\" was {decision.lower()}.{comment_line}{link_line}",
    )


@app.route("/jd/<int:jd_id>/submit-for-approval", methods=["POST"])
def jd_submit_for_approval(jd_id):
    if not _require_permission("write_jd"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        row = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not row:
            return "Job Description not found", 404
        if row["workflow_status"] not in ("Draft", "Changes Requested"):
            flash("Only a Draft or Changes-Requested JD can be submitted for approval.", "error")
            return redirect(url_for("jd_detail", jd_id=jd_id))
        conn.execute(
            "UPDATE job_description SET workflow_status = 'Pending Approval', submitted_at = NOW(), "
            "created_by = COALESCE(created_by, %s) WHERE id = %s",
            (session.get("user_id"), jd_id),
        )
    log_audit("JD Management", "Submit for Approval", record_id=jd_id, record_label=row["title"])
    _notify_jd_approvers(jd_id, row["title"], session.get("full_name"))
    flash("Submitted for approval.", "success")
    return redirect(url_for("jd_detail", jd_id=jd_id))


def _jd_decision(jd_id, new_status, decision_label, require_comment):
    """Shared body for approve/reject/request-changes — same guards
    (must be Pending Approval, creator can't act on their own submission),
    same field updates, same notify+audit steps; only the target status,
    label, and comment-required-ness differ per caller."""
    if not _require_permission("approve_jd"):
        abort(403)
    comment = request.form.get("comment", "").strip()
    if require_comment and not comment:
        return jsonify({"error": "A comment is required for this decision."}), 400
    with db_conn() as conn:
        ensure_jd_table(conn)
        row = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not row:
            return "Job Description not found", 404
        if row["workflow_status"] != "Pending Approval":
            flash("This JD is not currently pending approval.", "error")
            return redirect(url_for("jd_detail", jd_id=jd_id))
        if row["created_by"] and session.get("user_id") == row["created_by"]:
            flash("You cannot act on a JD you submitted yourself — another Hiring Manager "
                  "or Admin must review it.", "error")
            return redirect(url_for("jd_detail", jd_id=jd_id))
        conn.execute(
            "UPDATE job_description SET workflow_status = %s, decided_by = %s, decided_at = NOW(), "
            "decision_comment = %s WHERE id = %s",
            (new_status, session.get("user_id"), comment, jd_id),
        )
    log_audit("JD Management", decision_label, record_id=jd_id, record_label=row["title"],
              details=comment or None)
    _notify_jd_creator_of_decision(row, decision_label, comment)
    flash(f"JD {decision_label.lower()}.", "success")
    return redirect(url_for("jd_detail", jd_id=jd_id))


@app.route("/jd/<int:jd_id>/approve", methods=["POST"])
def jd_approve(jd_id):
    return _jd_decision(jd_id, "Approved", "Approved", require_comment=False)


@app.route("/jd/<int:jd_id>/reject", methods=["POST"])
def jd_reject(jd_id):
    return _jd_decision(jd_id, "Rejected", "Rejected", require_comment=True)


@app.route("/jd/<int:jd_id>/request-changes", methods=["POST"])
def jd_request_changes(jd_id):
    return _jd_decision(jd_id, "Changes Requested", "Changes Requested", require_comment=True)


@app.route("/jd/<int:jd_id>/publish", methods=["POST"])
def jd_publish(jd_id):
    if not _require_permission("approve_jd"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        row = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not row:
            return "Job Description not found", 404
        # Route-level guard, not just a hidden button: a Rejected (or any
        # non-Approved) JD can never be published, no matter how this route
        # is invoked.
        if row["workflow_status"] != "Approved":
            flash("Only an Approved JD can be published.", "error")
            return redirect(url_for("jd_detail", jd_id=jd_id))
        conn.execute(
            "UPDATE job_description SET workflow_status = 'Published', published_at = NOW() WHERE id = %s",
            (jd_id,),
        )
    log_audit("JD Management", "Publish", record_id=jd_id, record_label=row["title"])
    flash("Job Description published.", "success")
    return redirect(url_for("jd_detail", jd_id=jd_id))


def _jd_field_diff(a, b):
    """Field names that differ between two JD-shaped dicts (version snapshot
    or the live row) — drives both the versions list's "changed" summary and
    the compare page's highlighting."""
    fields = ["title", "role", "category", "responsibilities", "requirements", "skills",
              "keywords", "position_status"]
    return [f for f in fields if (a.get(f) or "") != (b.get(f) or "")]


@app.route("/jd/<int:jd_id>/versions")
def jd_versions(jd_id):
    if not _require_permission("view_jd", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_jd_version_table(conn)
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not jd:
            return "Job Description not found", 404
        versions = conn.execute(
            """
            SELECT v.*, u.full_name AS editor_name
            FROM jd_version v LEFT JOIN app_user u ON u.id = v.edited_by
            WHERE v.jd_id = %s ORDER BY v.version_number DESC
            """,
            (jd_id,),
        ).fetchall()
    versions = [dict(v) for v in versions]
    # Each version snapshot captured the state right BEFORE that edit — so
    # "what changed" in edit N is the diff between version N and whatever
    # came right after it (version N+1, or the live row for the newest one).
    for i, v in enumerate(versions):
        newer = versions[i - 1] if i > 0 else dict(jd)
        v["changed_fields"] = _jd_field_diff(v, newer)
    return render_template("jd_versions.html", jd=dict(jd), versions=versions)


def _load_jd_snapshot(conn, jd_id, ref):
    """ref is either the literal 'current' (the live job_description row) or
    a jd_version id — used by the compare view so either side of a
    comparison can be "what it looks like right now"."""
    if ref == "current":
        row = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        return dict(row) if row else None, "Current"
    row = conn.execute(
        "SELECT v.*, u.full_name AS editor_name FROM jd_version v "
        "LEFT JOIN app_user u ON u.id = v.edited_by WHERE v.id = %s AND v.jd_id = %s",
        (ref, jd_id),
    ).fetchone()
    return (dict(row), f"Version {row['version_number']}") if row else (None, None)


@app.route("/jd/<int:jd_id>/versions/compare")
def jd_version_compare(jd_id):
    if not _require_permission("view_jd", write=False):
        abort(403)
    a_ref = request.args.get("a", "current")
    b_ref = request.args.get("b", "current")
    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_jd_version_table(conn)
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not jd:
            return "Job Description not found", 404
        a_snapshot, a_label = _load_jd_snapshot(conn, jd_id, a_ref)
        b_snapshot, b_label = _load_jd_snapshot(conn, jd_id, b_ref)
    if not a_snapshot or not b_snapshot:
        return "Version not found", 404
    fields = ["title", "role", "category", "responsibilities", "requirements", "skills",
              "keywords", "position_status"]
    rows = [{"field": f, "a": a_snapshot.get(f) or "", "b": b_snapshot.get(f) or "",
             "differs": (a_snapshot.get(f) or "") != (b_snapshot.get(f) or "")} for f in fields]
    return render_template("jd_version_compare.html", jd=dict(jd), rows=rows,
                           a_label=a_label, b_label=b_label)


@app.route("/jd/<int:jd_id>/versions/<int:version_id>/rollback", methods=["POST"])
def jd_version_rollback(jd_id, version_id):
    # Admin/approver-only, per spec — reuses the same gate as the approval
    # decisions rather than write_jd, since rollback is a higher-privilege
    # action than a normal edit.
    if not _require_permission("approve_jd"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_jd_version_table(conn)
        current = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not current:
            return "Job Description not found", 404
        target = conn.execute(
            "SELECT * FROM jd_version WHERE id = %s AND jd_id = %s", (version_id, jd_id)
        ).fetchone()
        if not target:
            return "Version not found", 404
        # Preserve the pre-rollback state as its own version first — history
        # only ever grows, a rollback is never destructive to prior history.
        _snapshot_jd_version(conn, current, session.get("user_id"))
        conn.execute(
            """
            UPDATE job_description SET
                title=%s, role=%s, category=%s, responsibilities=%s,
                requirements=%s, skills=%s, keywords=%s, position_status=%s, updated_at=NOW()
            WHERE id=%s
            """,
            (target["title"], target["role"], target["category"], target["responsibilities"],
             target["requirements"], target["skills"], target["keywords"],
             target["position_status"], jd_id),
        )
    log_audit("JD Management", "Rollback", record_id=jd_id, record_label=current["title"],
              details=f"Rolled back to version {target['version_number']}")
    flash(f"Rolled back to version {target['version_number']}.", "success")
    return redirect(url_for("jd_detail", jd_id=jd_id))


@app.route("/uploads/jd/<path:filename>")
def jd_uploaded_file(filename):
    safe = secure_filename(filename)
    ext = safe.rsplit(".", 1)[-1].lower() if "." in safe else ""
    if ext not in ALLOWED_EXTENSIONS:
        abort(403)
    return send_from_directory(str(JD_UPLOAD_FOLDER), safe)


@app.route("/compare/<int:resume_id>")
def compare_select_jd(resume_id):
    if not _require_permission("compare_resume"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        seed_jds(conn)
        resume = conn.execute(
            "SELECT id, full_name, title FROM resume WHERE id = %s", (resume_id,)
        ).fetchone()
        if not resume:
            return "Resume not found", 404
        jds = conn.execute(
            "SELECT id, title, role, category FROM job_description ORDER BY category, title"
        ).fetchall()
    return render_template("compare_select.html", resume=dict(resume), jds=list(jds))


@app.route("/compare/<int:resume_id>/<int:jd_id>")
def compare_result(resume_id, jd_id):
    if not _require_permission("compare_resume"):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not resume or not jd:
            return "Resume or JD not found", 404
        all_jds = conn.execute(
            "SELECT id, title, category FROM job_description ORDER BY category, title"
        ).fetchall()
    result = calculate_match_score(dict(resume), dict(jd))
    log_audit("Compare Resume", "Compare", record_id=f"{resume_id}:{jd_id}",
              record_label=f"{resume['full_name']} vs {jd['title']}")
    trigger_background_ai_assessment(resume_id, [jd_id])
    return render_template(
        "compare_result.html",
        resume=dict(resume),
        jd=dict(jd),
        result=result,
        all_jds=list(all_jds),
    )


@app.route("/api/compare/<int:resume_id>/<int:jd_id>/ai-assessment")
def compare_ai_assessment(resume_id, jd_id):
    """Fetched asynchronously by compare_result.html / profile.html after the
    page loads. Delegates entirely to _get_or_compute_hybrid_match() — the
    one function that reads/writes ai_match_cache — so this endpoint and the
    Top 3 Matching Roles cards can never end up showing two different
    percentages for the same resume/JD pair.
    """
    logger.info(f"AI assessment requested → resume={resume_id} jd={jd_id}")
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not resume or not jd:
            return jsonify({"error": "Resume or JD not found"}), 404

        assessment = _get_or_compute_hybrid_match(conn, resume, jd)
    return jsonify(assessment)


@app.route("/api/compare/<int:resume_id>/<int:jd_id>/ai-status")
def compare_ai_status(resume_id, jd_id):
    """Fast, read-only cache check — never computes and never blocks on the
    advisory lock (contrast with compare_ai_assessment above). This is what
    the frontend polls every few seconds while a background AI assessment
    is running (see trigger_background_ai_assessment), so no browser
    request ever sits open for the full one-to-several minutes a real LLM
    judgment can take on local hardware.
    """
    with db_conn() as conn:
        # Full rows, not just id/updated_at — _read_cached_hybrid_match now
        # compares a content fingerprint over summary/skills/experience/etc.
        # (see _resume_match_fingerprint) and computes jd_mirroring_risk
        # fresh on every read (see _jd_mirroring_risk), both of which need
        # the actual skills/requirements/responsibilities columns present,
        # not just id/updated_at.
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not resume or not jd:
            return jsonify({"ready": False, "error": "Resume or JD not found"}), 404
        cached = _read_cached_hybrid_match(conn, resume, jd)
    if cached is None:
        return jsonify({"ready": False})
    return jsonify({"ready": True, **cached})


# ── NEW FEATURE: Top-matching resumes for a JD (reverse of Compare Resumes) ──
# Given a JD, scores every stored resume against it with the same
# calculate_match_score() used by the one-resume-at-a-time compare flow, and
# returns the top 3. Purely additive — new route, new template, no existing
# route/behavior touched.
@app.route("/jd/<int:jd_id>/top-matches")
def jd_top_matches(jd_id):
    with db_conn() as conn:
        ensure_jd_table(conn)
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if not jd:
            return "Job Description not found", 404
        resumes = conn.execute("SELECT * FROM resume").fetchall()

    jd_dict = dict(jd)
    scored = []
    for r in resumes:
        r_dict = dict(r)
        score = calculate_match_score(r_dict, jd_dict)
        holistic = _holistic_or_default(r_dict, jd_dict)
        scored.append({"resume": r_dict, "score": score, "holistic": holistic})

    scored.sort(key=lambda x: x["holistic"]["fit_percentage"], reverse=True)
    top_matches = scored[:3]

    return render_template(
        "jd_top_matches.html",
        jd=jd_dict,
        top_matches=top_matches,
        total_scored=len(scored),
    )


# ── Interview Questions & Answers (per-JD, under JD Management) ─────────────
# Fully additive: its own table (interview_questions), its own routes, its
# own template. Never touches resume parsing, JD matching, ai_match_cache, or
# any existing Profile/Resume screen.

# Per-category Ollama calls needed to reach the user's 25-question spec
# (10 Technical / 5 Functional / 5 Scenario / 5 Behavioral). Each call asks
# for at most 5 questions — found live (2026-08-12) that a single 10-question
# call reliably TIMED OUT at 180s on CPU-only Ollama while every 5-question
# call succeeded comfortably, so Technical is split into two 5-question
# calls rather than raised to a much longer timeout. This also follows this
# codebase's established lesson that a small local model is far less
# reliable one-shotting a large, richly-detailed JSON response than being
# fed small, homogeneous pieces (see parse_resume_with_llm_text and
# _llm_judge_match's prompt-design comments).
_INTERVIEW_QUESTION_BATCHES = [
    ("Technical", 5), ("Technical", 5),
    ("Functional", 5),
    ("Scenario Based", 5),
    ("Behavioral", 5),
]

# One-line style hint per category — found live (2026-08-12) that without
# this, Functional/Scenario/Behavioral all converged on near-identical
# "describe a time when..." phrasing (genuine content, correctly JD-specific,
# but not actually distinguishing the 3 question styles from each other).
_INTERVIEW_QUESTION_STYLE_HINTS = {
    "Technical": "test the candidate's domain/technical knowledge directly — not a personal story.",
    "Functional": "ask how the candidate would carry out a specific day-to-day responsibility of this role.",
    "Scenario Based": "pose a hypothetical situation (\"Imagine...\"/\"Suppose...\") and ask how they'd handle it.",
    "Behavioral": "ask the candidate to describe a specific past experience (\"Tell me about a time...\").",
}


def _build_jd_text_for_questions(jd_dict):
    return (
        f"Title: {jd_dict.get('title') or ''}\n"
        f"Role: {jd_dict.get('role') or ''}\n"
        f"Responsibilities: {jd_dict.get('responsibilities') or ''}\n"
        f"Requirements: {jd_dict.get('requirements') or ''}\n"
        f"Skills: {jd_dict.get('skills') or ''}\n"
        f"Keywords: {jd_dict.get('keywords') or ''}"
    )


def _generate_interview_questions_for_category(jd_text, category, count):
    """One Ollama call for one category. Returns a list of coerced question
    dicts (possibly empty if the model fails or returns nothing usable) —
    never raises for bad/malformed JSON (mirrors _ollama_chat's own
    contract), only for a genuine call failure (Ollama down/timeout), which
    the caller catches.
    """
    style_hint = _INTERVIEW_QUESTION_STYLE_HINTS.get(category, "")
    prompt = (
        "You are an experienced technical interviewer. Analyze the following Job "
        f"Description and create {count} {category} interview questions specifically "
        f"for this role. {style_hint}\n"
        "Questions must be directly relevant to the responsibilities, skills, "
        "technologies, qualifications, and experience mentioned in the JD. Do not "
        "invent technologies or requirements that are not relevant to it.\n\n"
        f"JOB DESCRIPTION:\n{jd_text}\n\n"
        "Respond with ONLY a single JSON object, no other text before or after it, in "
        "EXACTLY this shape:\n"
        '{"questions": [{"question": <string>, "expected_answer": <string>, '
        '"key_points": [<string>, ...up to 6], "difficulty": <"Easy"|"Medium"|"High">}, '
        f"... {count} items total]}}"
    )
    # Each call is capped at 5 questions (see _INTERVIEW_QUESTION_BATCHES) —
    # 1000 covers question + a several-sentence answer + up to 6 short
    # bullets, x5, with headroom; 180s timeout comfortably fits a 5-question
    # batch on CPU-only Ollama per live testing.
    raw = _ollama_chat(prompt, as_json=True, num_predict=1000, num_ctx=8192, timeout=180)
    return _coerce_interview_questions(raw, category)


@app.route("/jd/<int:jd_id>/interview-questions")
def jd_interview_questions(jd_id):
    if not _require_permission("view_jd", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_interview_questions_table(conn)
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not jd:
            return "Job Description not found", 404
        rows = conn.execute(
            "SELECT * FROM interview_questions WHERE jd_id = %s ORDER BY category, id",
            (jd_id,),
        ).fetchall()

    questions_by_category = {cat: [] for cat in INTERVIEW_QUESTION_CATEGORIES}
    for row in rows:
        questions_by_category.setdefault(row["category"], []).append(dict(row))

    return render_template(
        "interview_questions.html",
        jd=dict(jd),
        categories=INTERVIEW_QUESTION_CATEGORIES,
        difficulties=INTERVIEW_QUESTION_DIFFICULTIES,
        questions_by_category=questions_by_category,
        has_questions=len(rows) > 0,
    )


@app.route("/jd/<int:jd_id>/interview-questions/generate", methods=["POST"])
def generate_interview_questions(jd_id):
    """Generates (or regenerates) the full ~25-question set for one JD. Runs
    5 sequential Ollama calls (2 Technical + 1 each of Functional/Scenario/
    Behavioral) — see _INTERVIEW_QUESTION_BATCHES and
    _generate_interview_questions_for_category for why. A batch that comes
    back empty (model failure/timeout for just that call) is simply omitted
    rather than failing the whole request; only a total Ollama outage
    (every batch empty) is reported as an error.

    Gated on view_jd (not write_jd): every role that can view a JD at all
    can trigger its own one-time auto-generation (see interview_questions.html's
    inline script) — this must not require write access, or a view-only role
    opening a brand new JD first would get stuck on a permanent "unable to
    generate" error instead of ever seeing its auto-populated questions.
    """
    if not _require_permission("view_jd", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_jd_table(conn)
        jd = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not jd:
            return jsonify({"success": False, "message": "Job Description not found"}), 404
        jd_dict = dict(jd)

    jd_text = _build_jd_text_for_questions(jd_dict)
    all_questions = []
    for category, count in _INTERVIEW_QUESTION_BATCHES:
        try:
            all_questions.extend(_generate_interview_questions_for_category(jd_text, category, count))
        except Exception as e:
            # Caught per-batch, not around the whole loop — one flaky/
            # timed-out call (a real risk with 5 sequential CPU-bound local
            # calls) shouldn't lose the batches that already succeeded.
            logger.warning(f"Interview question generation failed for jd={jd_id} category={category}: {e}",
                            exc_info=True)

    if not all_questions:
        return jsonify({"success": False,
                         "message": "Unable to generate interview questions. Please try again."}), 502

    with db_conn() as conn:
        ensure_interview_questions_table(conn)
        conn.execute("DELETE FROM interview_questions WHERE jd_id = %s", (jd_id,))
        for q in all_questions:
            conn.execute(
                """
                INSERT INTO interview_questions
                    (jd_id, category, question, expected_answer, key_points, difficulty)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (jd_id, q["category"], q["question"], q["expected_answer"], q["key_points"], q["difficulty"]),
            )

    log_audit("JD Management", "Generate Interview Questions", record_id=jd_id,
              record_label=jd_dict.get("title"), details=f"{len(all_questions)} questions")

    questions_by_category = {cat: [] for cat in INTERVIEW_QUESTION_CATEGORIES}
    for q in all_questions:
        questions_by_category[q["category"]].append(q)
    return jsonify({"success": True, "questions": questions_by_category})


@app.route("/jd/<int:jd_id>/interview-questions/add", methods=["POST"])
def add_interview_question(jd_id):
    if not _require_permission("write_jd"):
        abort(403)
    data = request.get_json(silent=True) or {}
    category = str(data.get("category") or "").strip().title()
    question = str(data.get("question") or "").strip()
    if category not in INTERVIEW_QUESTION_CATEGORIES:
        return jsonify({"ok": False, "error": "Invalid category"}), 400
    if not question:
        return jsonify({"ok": False, "error": "Question is required"}), 400
    difficulty = str(data.get("difficulty") or "Medium").strip().title()
    if difficulty not in INTERVIEW_QUESTION_DIFFICULTIES:
        difficulty = "Medium"
    expected_answer = str(data.get("expected_answer") or "").strip()[:600]
    key_points = data.get("key_points")
    if isinstance(key_points, list):
        key_points = "\n".join(_coerce_str_list(key_points, 6))
    else:
        key_points = str(key_points or "").strip()

    with db_conn() as conn:
        ensure_jd_table(conn)
        ensure_interview_questions_table(conn)
        jd = conn.execute("SELECT title FROM job_description WHERE id = %s", (jd_id,)).fetchone()
        if not jd:
            return jsonify({"ok": False, "error": "Job Description not found"}), 404
        new_id = conn.execute(
            """
            INSERT INTO interview_questions
                (jd_id, category, question, expected_answer, key_points, difficulty)
            VALUES (%s, %s, %s, %s, %s, %s)
            RETURNING id
            """,
            (jd_id, category, question, expected_answer, key_points, difficulty),
        ).fetchone()["id"]

    log_audit("JD Management", "Add Interview Question", record_id=new_id, record_label=jd["title"])
    return jsonify({"ok": True, "question": {
        "id": new_id, "category": category, "question": question,
        "expected_answer": expected_answer, "key_points": key_points, "difficulty": difficulty,
    }})


@app.route("/api/interview-question/<int:q_id>/update", methods=["POST"])
def update_interview_question(q_id):
    if not _require_permission("write_jd"):
        abort(403)
    data = request.get_json(silent=True) or {}
    question = str(data.get("question") or "").strip()
    if not question:
        return jsonify({"ok": False, "error": "Question is required"}), 400
    difficulty = str(data.get("difficulty") or "Medium").strip().title()
    if difficulty not in INTERVIEW_QUESTION_DIFFICULTIES:
        difficulty = "Medium"
    expected_answer = str(data.get("expected_answer") or "").strip()[:600]
    key_points = data.get("key_points")
    if isinstance(key_points, list):
        key_points = "\n".join(_coerce_str_list(key_points, 6))
    else:
        key_points = str(key_points or "").strip()

    with db_conn() as conn:
        ensure_interview_questions_table(conn)
        row = conn.execute("SELECT id FROM interview_questions WHERE id = %s", (q_id,)).fetchone()
        if not row:
            return jsonify({"ok": False, "error": "Question not found"}), 404
        conn.execute(
            """
            UPDATE interview_questions
            SET question = %s, expected_answer = %s, key_points = %s, difficulty = %s,
                updated_at = NOW()
            WHERE id = %s
            """,
            (question, expected_answer, key_points, difficulty, q_id),
        )
    log_audit("JD Management", "Edit Interview Question", record_id=q_id)
    return jsonify({"ok": True})


@app.route("/api/interview-question/<int:q_id>/delete", methods=["POST"])
def delete_interview_question(q_id):
    if not _require_permission("write_jd"):
        abort(403)
    with db_conn() as conn:
        ensure_interview_questions_table(conn)
        row = conn.execute("SELECT id FROM interview_questions WHERE id = %s", (q_id,)).fetchone()
        if not row:
            return jsonify({"ok": False, "error": "Question not found"}), 404
        conn.execute("DELETE FROM interview_questions WHERE id = %s", (q_id,))
    log_audit("JD Management", "Delete Interview Question", record_id=q_id)
    return jsonify({"ok": True})


# ── NEW: Candidate list API (used by dashboard refresh) ──────────────────────
@app.route("/api/candidates")
def api_candidates():
    if not _require_permission("view_candidate", write=False):
        abort(403)
    with db_conn() as conn:
        rows = conn.execute(
            "SELECT id, full_name, title, location, experience, summary, created_at,"
            " COALESCE(status, 'New') AS status,"
            " COALESCE(interview_status, 'To Be Interviewed') AS interview_status,"
            " COALESCE(exp_yrs, '') AS exp_yrs"
            " FROM resume ORDER BY created_at DESC LIMIT 200"
        ).fetchall()

    def _ini(name):
        parts = (name or "").split()
        if not parts: return "?"
        return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()

    def _exp(summary_txt, exp_txt):
        txt = (summary_txt or "")[:600] + " " + (exp_txt or "")[:400]
        m = re.search(r'(\d+)\s*\+?\s*years?\s+(?:of\s+)?(?:experience|exp\b)', txt, re.I)
        if m: return m.group(1) + "+"
        m = re.search(r'(\d+)\s*\+\s*years?', txt, re.I)
        if m: return m.group(1) + "+"
        return "—"

    candidates = [
        {
            "id": r["id"],
            "full_name": r["full_name"] or "—",
            "title": r["title"] or "—",
            "location": r["location"] or "—",
            "exp_years": r["exp_yrs"] if r["exp_yrs"] else _exp(r["summary"], r["experience"]),
            "status": r["status"] or "New",
            "interview_status": r["interview_status"] or "To Be Interviewed",
            "initials": _ini(r["full_name"] or ""),
            "summary_snippet": (r["summary"] or "").strip()[:130],
            "created_at": r["created_at"].strftime("%Y-%m-%d") if r["created_at"] else "—",
        }
        for r in rows
    ]
    status_counts = {"New": 0, "Reviewed": 0, "Shortlisted": 0}
    for c in candidates:
        status_counts[c["status"]] = status_counts.get(c["status"], 0) + 1
    interview_status_counts = {}
    for c in candidates:
        _ist = c["interview_status"]
        interview_status_counts[_ist] = interview_status_counts.get(_ist, 0) + 1
    return jsonify({"candidates": candidates, "status_counts": status_counts, "interview_status_counts": interview_status_counts})


# ── NEW: Candidate status update route ───────────────────────────────────────
@app.route("/candidate-status", methods=["POST"])
def update_candidate_status():
    data = request.get_json(silent=True) or {}
    resume_id = data.get("id")
    status = data.get("status")
    if not resume_id or status not in ("New", "Reviewed", "Shortlisted"):
        return jsonify({"ok": False, "error": "invalid"}), 400
    with db_conn() as conn:
        conn.execute(
            "UPDATE resume SET status = %s WHERE id = %s",
            (status, resume_id),
        )
    return jsonify({"ok": True})


# ── NEW: Ensure status column exists (runs once at startup) ──────────────────
def _ensure_status_col():
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE resume ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'New'"
            )
    except Exception:
        pass

_ensure_status_col()


# ── NEW: Interview Status column + route ─────────────────────────────────────

INTERVIEW_STATUSES = [
    "To Be Interviewed",
    "Profile Shared With Client",
    "Can be considered for Future Roles",
    "Client Offer Released",
    "Joined Vaisesika",
    "Not Available Currently",
    "Not Shortlisted",
    "Position Closed",
    "Rejected By Client",
    "Rejected Internally",
    "Selected by Client",
]

# Single source of truth for interview-status colours. The dashboard donut and
# the "Interview Status" bar list both read this, so the same status can never
# render in two different colours across the two panels.
INTERVIEW_STATUS_COLORS = {
    "To Be Interviewed":                  "#6366f1",
    "Profile Shared With Client":         "#0284c7",
    "Can be considered for Future Roles": "#8b5cf6",
    "Client Offer Released":              "#059669",
    "Joined Vaisesika":                   "#10b981",
    "Not Available Currently":            "#64748b",
    "Not Shortlisted":                    "#f59e0b",
    "Position Closed":                    "#dc2626",
    "Rejected By Client":                 "#ef4444",
    "Rejected Internally":                "#991b1b",
    "Selected by Client":                 "#16a34a",
}
INTERVIEW_STATUS_FALLBACK_COLOR = "#6366f1"


def _ensure_interview_status_col():
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE resume ADD COLUMN IF NOT EXISTS interview_status TEXT DEFAULT 'To Be Interviewed'"
            )
    except Exception:
        pass


_ensure_interview_status_col()


def _ensure_comment_cols():
    try:
        with db_conn() as conn:
            conn.execute("ALTER TABLE resume ADD COLUMN IF NOT EXISTS l1_comments TEXT DEFAULT ''")
            conn.execute("ALTER TABLE resume ADD COLUMN IF NOT EXISTS l2_comments TEXT DEFAULT ''")
    except Exception:
        pass


_ensure_comment_cols()


def _ensure_exp_yrs_col():
    try:
        with db_conn() as conn:
            conn.execute("ALTER TABLE resume ADD COLUMN IF NOT EXISTS exp_yrs TEXT DEFAULT ''")
    except Exception:
        pass


_ensure_exp_yrs_col()


def _ensure_department_col():
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE resume ADD COLUMN IF NOT EXISTS department TEXT DEFAULT ''"
            )
    except Exception:
        pass


try:
    with db_conn() as conn:
        ensure_role_match_cache_table(conn)
except Exception as e:
    logger.error(f"Role cache table creation failed: {e}")


def _migrate_jd_category_it_roles():
    try:
        with db_conn() as conn:
            conn.execute(
                "UPDATE job_description SET category = 'IT Technology' WHERE category = 'IT Roles'"
            )
    except Exception:
        pass


_ensure_department_col()
_migrate_jd_category_it_roles()


DEPARTMENTS = ["IT", "QA", "Finance", "HR", "Sales", "Marketing", "Operations", "Other"]

# Requirement Management reuses the same Position Status enum as JD Management
# (POSITION_STATUSES, defined alongside job_description) rather than keeping a
# second, drift-prone copy of the same four values.
REQUIREMENT_STATUSES = POSITION_STATUSES


def _migrate_requirement_status_filled_to_fulfilled():
    """One-time rename: the requirement.status enum used to say 'Filled'
    where it now says 'Fulfilled' (to match JD Management's Position Status
    wording) — existing rows must be updated or they'd show a value that no
    longer appears in the dropdown."""
    try:
        with db_conn() as conn:
            conn.execute(
                "UPDATE requirement SET status = 'Fulfilled' WHERE status = 'Filled'"
            )
    except Exception:
        pass


def ensure_requirement_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS requirement (
            id                SERIAL PRIMARY KEY,
            requirement_code  TEXT NOT NULL,
            requirement_name  TEXT NOT NULL,
            client            TEXT DEFAULT '',
            division          TEXT DEFAULT '',
            num_requirement   INTEGER DEFAULT 1,
            status            TEXT DEFAULT 'Open',
            profiles_shared   INTEGER DEFAULT 0,
            interviewed       INTEGER DEFAULT 0,
            offered           INTEGER DEFAULT 0,
            onsite_offshore   TEXT DEFAULT '',
            created_at        TIMESTAMPTZ DEFAULT NOW(),
            updated_at        TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )


ONSITE_OFFSHORE_OPTIONS = ["Onsite", "Offshore", "Hybrid"]
PRIORITY_LEVELS = ["Low", "Medium", "High", "Critical"]


def _ensure_onsite_offshore_col():
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE requirement ADD COLUMN IF NOT EXISTS onsite_offshore TEXT DEFAULT ''"
            )
    except Exception:
        pass


def _ensure_requirement_recruiter_priority_cols():
    """Assigned Recruiter + Priority — used for the Requirement Management
    filter set. Nullable/defaulted so existing rows are unaffected."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE requirement ADD COLUMN IF NOT EXISTS "
                "assigned_recruiter_id INTEGER REFERENCES app_user(id)"
            )
            conn.execute(
                "ALTER TABLE requirement ADD COLUMN IF NOT EXISTS priority TEXT DEFAULT 'Medium'"
            )
    except Exception:
        pass


def _ensure_requirement_jd_col():
    """Links a requirement to the (single) JD it was raised against. A JD can
    have many requirements linked to it, so the FK lives on requirement, not
    a join table. ON DELETE SET NULL: deleting a JD must never delete the
    requisition records raised against it — it only drops the link."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE requirement ADD COLUMN IF NOT EXISTS "
                "jd_id INTEGER REFERENCES job_description(id) ON DELETE SET NULL"
            )
    except Exception:
        pass


_migrate_requirement_status_filled_to_fulfilled()
_ensure_onsite_offshore_col()
_ensure_requirement_jd_col()
_ensure_requirement_recruiter_priority_cols()


def ensure_interview_schedule_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS interview_schedule (
            id              SERIAL PRIMARY KEY,
            resume_id       INTEGER NOT NULL REFERENCES resume(id) ON DELETE CASCADE,
            position        TEXT DEFAULT '',
            interviewer     TEXT DEFAULT '',
            interview_date  DATE,
            status          TEXT DEFAULT 'Scheduled',
            created_at      TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )


# Child table of job_description — one row per generated/hand-added interview
# question. Deliberately real columns, not a JSONB blob (contrast
# ai_match_cache.extra): questions need individual edit/delete, which a
# one-row-per-question table supports directly, matching the
# interview_schedule precedent above rather than the cache-blob convention.
INTERVIEW_QUESTION_CATEGORIES = ["Technical", "Functional", "Scenario Based", "Behavioral"]
INTERVIEW_QUESTION_DIFFICULTIES = ["Easy", "Medium", "High"]


def ensure_interview_questions_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS interview_questions (
            id              SERIAL PRIMARY KEY,
            jd_id           INTEGER NOT NULL REFERENCES job_description(id) ON DELETE CASCADE,
            category        VARCHAR(20) NOT NULL,
            question        TEXT NOT NULL,
            expected_answer TEXT DEFAULT '',
            key_points      TEXT DEFAULT '',
            difficulty      VARCHAR(10) DEFAULT 'Medium',
            created_at      TIMESTAMPTZ DEFAULT NOW(),
            updated_at      TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )


def _migrate_interview_question_labels():
    """One-time rename to match the category/difficulty labels now used
    throughout the UI: 'Scenario' -> 'Scenario Based', 'Hard' -> 'High'.
    Existing questions/answers are relabeled in place, never deleted —
    every old row keeps its question+answer, just under the new label."""
    try:
        with db_conn() as conn:
            ensure_interview_questions_table(conn)
            conn.execute("UPDATE interview_questions SET category = 'Scenario Based' WHERE category = 'Scenario'")
            conn.execute("UPDATE interview_questions SET difficulty = 'High' WHERE difficulty = 'Hard'")
    except Exception:
        pass


_migrate_interview_question_labels()


# ── Requirement Management Routes ────────────────────────────────────────────

@app.route("/requirement-management")
def requirement_management():
    if not _require_permission("view_requirement", write=False):
        abort(403)
    search_query = (request.args.get("q") or "").strip()
    with db_conn() as conn:
        ensure_requirement_table(conn)
        ensure_jd_table(conn)
        conditions, params = [], []
        if search_query:
            like = f"%{search_query}%"
            conditions.append(
                "(r.requirement_code ILIKE %s OR r.requirement_name ILIKE %s OR r.client ILIKE %s "
                "OR r.division ILIKE %s OR CAST(r.id AS TEXT) = %s)"
            )
            params += [like, like, like, like, search_query]
        where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        requirements = conn.execute(
            "SELECT r.*, jd.title AS jd_title, u.full_name AS recruiter_name FROM requirement r "
            "LEFT JOIN job_description jd ON jd.id = r.jd_id "
            f"LEFT JOIN app_user u ON u.id = r.assigned_recruiter_id {where_sql} "
            "ORDER BY r.created_at DESC",
            params,
        ).fetchall()
    return render_template("requirement_management.html", requirements=list(requirements),
                           search_query=search_query, priority_levels=PRIORITY_LEVELS,
                           statuses=REQUIREMENT_STATUSES)


def _jd_options_for_requirement(conn, current_jd_id=None):
    """Published JDs only (the spec's "select an existing published JD"),
    plus the requirement's currently-linked JD even if its status has since
    moved on — so editing a requirement never silently drops an existing
    link from the dropdown."""
    rows = [dict(r) for r in conn.execute(
        "SELECT id, title, workflow_status FROM job_description "
        "WHERE workflow_status = 'Published' ORDER BY title"
    ).fetchall()]
    if current_jd_id and not any(r["id"] == current_jd_id for r in rows):
        extra = conn.execute(
            "SELECT id, title, workflow_status FROM job_description WHERE id = %s", (current_jd_id,)
        ).fetchone()
        if extra:
            rows.append(dict(extra))
    return rows


def _recruiter_options(conn):
    """Users eligible to be a requirement's Assigned Recruiter."""
    return [dict(r) for r in conn.execute(
        "SELECT id, full_name, username FROM app_user "
        "WHERE role IN ('recruiter', 'hiring_manager') AND is_active ORDER BY full_name"
    ).fetchall()]


def _validate_jd_link(conn, jd_id_raw):
    """Empty selection is valid (means "no link"). Returns (jd_id_or_None,
    error_message_or_None) — server-side re-check beyond the dropdown
    filtering, since a Draft/Rejected JD must never be linkable even if the
    form is tampered with directly."""
    jd_id_raw = (jd_id_raw or "").strip()
    if not jd_id_raw:
        return None, None
    try:
        jd_id = int(jd_id_raw)
    except ValueError:
        return None, "Invalid Job Description selection."
    row = conn.execute("SELECT workflow_status FROM job_description WHERE id = %s", (jd_id,)).fetchone()
    if not row:
        return None, "Selected Job Description no longer exists."
    if row["workflow_status"] in ("Draft", "Rejected"):
        return None, "A Draft or Rejected JD cannot be linked to a requirement."
    return jd_id, None


@app.route("/requirement/add", methods=["GET", "POST"])
def requirement_add():
    if not _require_permission("write_requirement"):
        abort(403)
    if request.method == "POST":
        data = {k: request.form.get(k, "").strip() for k in
                ["requirement_name", "client", "division", "status", "onsite_offshore", "priority"]}
        if data["priority"] not in PRIORITY_LEVELS:
            data["priority"] = "Medium"
        for k in ["num_requirement", "profiles_shared", "interviewed", "offered"]:
            try:
                data[k] = int(request.form.get(k, "0") or "0")
            except ValueError:
                data[k] = 0
        recruiter_raw = request.form.get("assigned_recruiter_id", "").strip()
        data["assigned_recruiter_id"] = int(recruiter_raw) if recruiter_raw.isdigit() else None
        with db_conn() as conn:
            ensure_requirement_table(conn)
            ensure_jd_table(conn)
            jd_id, jd_error = _validate_jd_link(conn, request.form.get("jd_id"))
            if jd_error:
                flash(jd_error, "error")
                return render_template(
                    "requirement_form.html", requirement=None, statuses=REQUIREMENT_STATUSES,
                    onsite_offshore_options=ONSITE_OFFSHORE_OPTIONS, priority_levels=PRIORITY_LEVELS,
                    jd_options=_jd_options_for_requirement(conn), recruiter_options=_recruiter_options(conn),
                ), 400
            data["jd_id"] = jd_id
            # requirement_code is system-generated, not user-entered — insert
            # with a placeholder, then derive the real code from the row's
            # own new id (unique by construction, no collision check needed).
            cursor = conn.execute(
                """
                INSERT INTO requirement
                    (requirement_code, requirement_name, client, division, num_requirement,
                     status, profiles_shared, interviewed, offered, onsite_offshore, jd_id,
                     assigned_recruiter_id, priority)
                VALUES
                    ('', %(requirement_name)s, %(client)s, %(division)s,
                     %(num_requirement)s, %(status)s, %(profiles_shared)s, %(interviewed)s, %(offered)s,
                     %(onsite_offshore)s, %(jd_id)s, %(assigned_recruiter_id)s, %(priority)s)
                RETURNING id
                """,
                data,
            )
            new_req_id = cursor.fetchone()["id"]
            new_code = f"REQ-{new_req_id:05d}"
            conn.execute("UPDATE requirement SET requirement_code = %s WHERE id = %s", (new_code, new_req_id))
        log_audit("Requirements", "Add", record_id=new_req_id, record_label=data["requirement_name"])
        flash(f"Requirement '{data['requirement_name']}' added as {new_code}.", "success")
        return redirect(url_for("requirement_management"))
    with db_conn() as conn:
        ensure_jd_table(conn)
        jd_options = _jd_options_for_requirement(conn)
        recruiter_options = _recruiter_options(conn)
    return render_template("requirement_form.html", requirement=None,
                            statuses=REQUIREMENT_STATUSES, onsite_offshore_options=ONSITE_OFFSHORE_OPTIONS,
                            jd_options=jd_options, recruiter_options=recruiter_options,
                            priority_levels=PRIORITY_LEVELS)


@app.route("/requirement/<int:req_id>")
def requirement_detail(req_id):
    if not _require_permission("view_requirement", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_requirement_table(conn)
        ensure_jd_table(conn)
        requirement = conn.execute(
            "SELECT r.*, jd.title AS jd_title, u.full_name AS recruiter_name FROM requirement r "
            "LEFT JOIN job_description jd ON jd.id = r.jd_id "
            "LEFT JOIN app_user u ON u.id = r.assigned_recruiter_id WHERE r.id = %s",
            (req_id,),
        ).fetchone()
        if not requirement:
            return "Requirement not found", 404
    return render_template("requirement_detail.html", requirement=dict(requirement))


@app.route("/requirement/<int:req_id>/edit", methods=["GET", "POST"])
def requirement_edit(req_id):
    if not _require_permission("write_requirement"):
        abort(403)
    with db_conn() as conn:
        ensure_requirement_table(conn)
        ensure_jd_table(conn)
        if request.method == "POST":
            data = {k: request.form.get(k, "").strip() for k in
                    ["requirement_name", "client", "division", "status", "onsite_offshore", "priority"]}
            if data["priority"] not in PRIORITY_LEVELS:
                data["priority"] = "Medium"
            for k in ["num_requirement", "profiles_shared", "interviewed", "offered"]:
                try:
                    data[k] = int(request.form.get(k, "0") or "0")
                except ValueError:
                    data[k] = 0
            recruiter_raw = request.form.get("assigned_recruiter_id", "").strip()
            data["assigned_recruiter_id"] = int(recruiter_raw) if recruiter_raw.isdigit() else None
            jd_id, jd_error = _validate_jd_link(conn, request.form.get("jd_id"))
            if jd_error:
                flash(jd_error, "error")
                existing = conn.execute("SELECT * FROM requirement WHERE id = %s", (req_id,)).fetchone()
                return render_template(
                    "requirement_form.html", requirement=dict(existing) if existing else None,
                    statuses=REQUIREMENT_STATUSES, onsite_offshore_options=ONSITE_OFFSHORE_OPTIONS,
                    jd_options=_jd_options_for_requirement(conn, existing["jd_id"] if existing else None),
                    recruiter_options=_recruiter_options(conn), priority_levels=PRIORITY_LEVELS,
                ), 400
            # requirement_code is intentionally left out of this UPDATE — it's
            # system-generated once at creation and immutable afterward.
            data["jd_id"] = jd_id
            data["id"] = req_id
            conn.execute(
                """
                UPDATE requirement SET
                    requirement_name=%(requirement_name)s,
                    client=%(client)s, division=%(division)s, num_requirement=%(num_requirement)s,
                    status=%(status)s, profiles_shared=%(profiles_shared)s,
                    interviewed=%(interviewed)s, offered=%(offered)s, onsite_offshore=%(onsite_offshore)s,
                    jd_id=%(jd_id)s, assigned_recruiter_id=%(assigned_recruiter_id)s, priority=%(priority)s,
                    updated_at=NOW()
                WHERE id=%(id)s
                """,
                data,
            )
            log_audit("Requirements", "Edit", record_id=req_id, record_label=data["requirement_name"])
            flash("Requirement updated.", "success")
            return redirect(url_for("requirement_detail", req_id=req_id))
        requirement = conn.execute(
            "SELECT * FROM requirement WHERE id = %s", (req_id,)
        ).fetchone()
        if not requirement:
            return "Requirement not found", 404
        jd_options = _jd_options_for_requirement(conn, requirement["jd_id"])
        recruiter_options = _recruiter_options(conn)
    return render_template("requirement_form.html", requirement=dict(requirement),
                            statuses=REQUIREMENT_STATUSES, onsite_offshore_options=ONSITE_OFFSHORE_OPTIONS,
                            jd_options=jd_options, recruiter_options=recruiter_options,
                            priority_levels=PRIORITY_LEVELS)


@app.route("/requirement/<int:req_id>/delete", methods=["POST"])
def requirement_delete(req_id):
    if not _require_permission("delete_requirement"):
        abort(403)
    with db_conn() as conn:
        ensure_requirement_table(conn)
        row = conn.execute("SELECT requirement_name FROM requirement WHERE id = %s", (req_id,)).fetchone()
        conn.execute("DELETE FROM requirement WHERE id = %s", (req_id,))
    log_audit("Requirements", "Delete", record_id=req_id, record_label=row["requirement_name"] if row else None)
    flash("Requirement deleted.", "success")
    return redirect(url_for("requirement_management"))


@app.route("/api/schedule-interview", methods=["POST"])
def schedule_interview():
    if not _require_permission("schedule_interview"):
        abort(403)
    data = request.get_json(silent=True) or {}
    resume_id = data.get("resume_id")
    interviewer = (data.get("interviewer") or "").strip()
    interview_date = (data.get("interview_date") or "").strip() or None
    status = (data.get("status") or "Scheduled").strip()
    if not resume_id or not interviewer or not interview_date:
        return jsonify({"ok": False, "error": "resume_id, interviewer, and interview_date are required"}), 400
    with db_conn() as conn:
        ensure_interview_schedule_table(conn)
        resume_row = conn.execute(
            "SELECT title, full_name FROM resume WHERE id = %s", (resume_id,)
        ).fetchone()
        if not resume_row:
            return jsonify({"ok": False, "error": "Candidate not found"}), 404
        new_interview_id = conn.execute(
            """
            INSERT INTO interview_schedule (resume_id, position, interviewer, interview_date, status)
            VALUES (%s, %s, %s, %s, %s)
            RETURNING id
            """,
            (resume_id, resume_row["title"] or "", interviewer, interview_date, status),
        ).fetchone()["id"]
    log_audit(
        "Interviews", "Schedule", record_id=new_interview_id,
        record_label=f"{resume_row['full_name']} with {interviewer} ({status})",
    )
    return jsonify({"ok": True})


# Interview EVENT status — the 4-value Scheduled/Completed/Pending/Cancelled
# lifecycle of one scheduled interview, shown in the Recent Interviews table.
# Distinct from INTERVIEW_STATUSES above, which tracks a candidate's much
# longer overall pipeline stage (Profile Shared, Selected by Client, etc.)
# shown in Candidate Details — the two have never been the same concept and
# aren't meant to be merged.
INTERVIEW_SCHEDULE_STATUSES = ["Scheduled", "Completed", "Pending", "Cancelled"]


@app.route("/api/interview/<int:interview_id>/update", methods=["POST"])
def update_interview_schedule(interview_id):
    """Edit an existing scheduled interview's status/date/interviewer.

    The candidate (resume_id) is intentionally never editable here — once an
    interview is created it stays tied to the same candidate; rescheduling
    against a different candidate means creating a new interview instead.
    """
    data = request.get_json(silent=True) or {}
    interviewer = (data.get("interviewer") or "").strip()
    interview_date = (data.get("interview_date") or "").strip() or None
    status = (data.get("status") or "").strip()

    if status not in INTERVIEW_SCHEDULE_STATUSES:
        return jsonify({"ok": False, "error": "Invalid status"}), 400
    if not interviewer or not interview_date:
        return jsonify({"ok": False, "error": "interviewer and interview_date are required"}), 400

    with db_conn() as conn:
        ensure_interview_schedule_table(conn)
        row = conn.execute(
            """
            SELECT i.id, i.interviewer AS current_interviewer, r.full_name
            FROM interview_schedule i
            JOIN resume r ON r.id = i.resume_id
            WHERE i.id = %s
            """,
            (interview_id,),
        ).fetchone()
        if not row:
            return jsonify({"ok": False, "error": "Interview not found"}), 404

        # "Assigned Only" per the RBAC matrix: an Interviewer may update only
        # interviews currently assigned to them (best-effort name match, see
        # _is_my_assigned_interview) — every other role needs the standard
        # "full" permission level.
        if not _require_permission("update_interview_status"):
            if session.get("role") != "interviewer" or not _is_my_assigned_interview(row["current_interviewer"]):
                abort(403)

        conn.execute(
            """
            UPDATE interview_schedule
            SET interviewer = %s, interview_date = %s, status = %s
            WHERE id = %s
            """,
            (interviewer, interview_date, status, interview_id),
        )
    log_audit(
        "Interviews", "Edit", record_id=interview_id,
        record_label=f"{row['full_name']} with {interviewer} ({status})",
    )
    return jsonify({"ok": True})


@app.route("/profile/<int:resume_id>/save-comments", methods=["POST"])
def save_profile_comments(resume_id):
    data = request.get_json(silent=True) or {}
    l1 = data.get("l1_comments", "")
    l2 = data.get("l2_comments", "")
    with db_conn() as conn:
        conn.execute(
            "UPDATE resume SET l1_comments = %s, l2_comments = %s WHERE id = %s",
            (l1, l2, resume_id),
        )
    return jsonify({"ok": True})


@app.route("/api/check-duplicate", methods=["POST"])
def check_duplicate_profile():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip()
    phone = (data.get("phone") or "").strip()
    exclude_id = data.get("exclude_id")

    if not email and not phone:
        return jsonify({"duplicate": False})

    with db_conn() as conn:
        conditions, params = [], []
        if email:
            conditions.append("LOWER(TRIM(email)) = LOWER(TRIM(%s))")
            params.append(email)
        if phone:
            clean = re.sub(r"[\s\-\(\)]", "", phone)
            conditions.append("REGEXP_REPLACE(phone, '[\\s\\-\\(\\)]', '', 'g') = %s")
            params.append(clean)
        query = (
            "SELECT id, full_name, email, phone FROM resume WHERE ("
            + " OR ".join(conditions) + ")"
        )
        if exclude_id:
            query += " AND id != %s"
            params.append(exclude_id)
        query += " LIMIT 1"
        row = conn.execute(query, params).fetchone()

    if row:
        return jsonify({
            "duplicate": True,
            "name": row["full_name"] or "Unknown",
            "email": row["email"] or "",
            "phone": row["phone"] or "",
            "id": row["id"],
        })
    return jsonify({"duplicate": False})


@app.route("/candidate-interview-status", methods=["POST"])
def update_candidate_interview_status():
    data = request.get_json(silent=True) or {}
    resume_id = data.get("id")
    status = data.get("status")
    if not resume_id or status not in INTERVIEW_STATUSES:
        return jsonify({"ok": False, "error": "invalid"}), 400
    with db_conn() as conn:
        conn.execute(
            "UPDATE resume SET interview_status = %s WHERE id = %s", (status, resume_id)
        )
    return jsonify({"ok": True})


# ── Audit Trail ───────────────────────────────────────────────────────────────
# Single reusable logging service (log_audit) backing one table (audit_log),
# used by every module instead of each one rolling its own logging. Two
# coverage layers:
#   1. Explicit log_audit(...) calls at meaningful business actions (Login,
#      Logout, Add/Edit/Delete, Compare, Export, Download) — these carry
#      precise "record affected" context (e.g. which resume/JD) that no
#      generic hook could infer from the URL alone.
#   2. A generic @app.after_request hook (_audit_generic_view, registered
#      further down) that logs a "View" action for every successful GET to
#      an authenticated HTML page not already covered by an explicit call —
#      so a brand-new route gets baseline audit coverage automatically, with
#      no extra code, satisfying "any future module automatically".

def ensure_audit_log_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_log (
            id            SERIAL PRIMARY KEY,
            user_id       INTEGER,
            username      VARCHAR(150),
            role          VARCHAR(50),
            created_at    TIMESTAMPTZ DEFAULT NOW(),
            module        VARCHAR(100) NOT NULL,
            action        VARCHAR(100) NOT NULL,
            record_id     VARCHAR(100),
            record_label  VARCHAR(300),
            ip_address    VARCHAR(64),
            user_agent    TEXT,
            status        VARCHAR(20) DEFAULT 'Success',
            details       TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_created_at ON audit_log (created_at DESC)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_user_id ON audit_log (user_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_module ON audit_log (module)")


def _client_ip():
    # Trust X-Forwarded-For's first hop only if this ever sits behind a
    # reverse proxy; falls back to the direct peer address for the common
    # case (local/dev, no proxy) where request.remote_addr is authoritative.
    forwarded = request.headers.get("X-Forwarded-For", "")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.remote_addr or ""


def log_audit(module, action, record_id=None, record_label=None, status="Success", details=None):
    """The one function every route calls to write an audit trail entry —
    never raises (a logging failure must never break the user-facing
    action it's describing), and always opens its own short-lived
    connection so it works regardless of whether the caller has one open.
    Pulls the acting user from the Flask session and request metadata
    (IP/user-agent) automatically, so callers only ever need to supply the
    business-specific bits: which module, what action, and which record.
    """
    try:
        with db_conn() as conn:
            ensure_audit_log_table(conn)
            conn.execute(
                """
                INSERT INTO audit_log
                    (user_id, username, role, module, action, record_id, record_label,
                     ip_address, user_agent, status, details)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    session.get("user_id"), session.get("username"), session.get("role"),
                    str(module)[:100], str(action)[:100],
                    str(record_id)[:100] if record_id is not None else None,
                    str(record_label)[:300] if record_label is not None else None,
                    _client_ip()[:64], (request.headers.get("User-Agent") or "")[:500],
                    str(status)[:20], (str(details)[:2000] if details else None),
                ),
            )
    except Exception:
        logger.exception(f"Audit log write failed → module={module} action={action}")


# Endpoint-name prefix → friendly module label, used by the generic
# after_request hook to auto-label "View" events for routes that don't have
# an explicit log_audit() call of their own. Falls back to the endpoint's own
# name (title-cased) for anything not listed — see _audit_generic_view.
_AUDIT_MODULE_BY_ENDPOINT_PREFIX = [
    ("dashboard", "Dashboard"), ("home", "Dashboard"),
    ("user_", "User Management"),
    ("profile_", "Resume Profiles"), ("edit_resume", "Resume Profiles"),
    ("public_profile", "Resume Profiles"), ("upload_files", "Resume Profiles"),
    ("groups", "Resume Profiles"),
    ("jd_", "JD Management"),
    ("requirement_", "Requirements"),
    ("compare_", "Compare Resume"), ("bulk_compare", "Compare Resume"),
    ("export_", "Export"), ("download_", "Downloads"),
    ("candidate_", "Dashboard"), ("api_candidates", "Dashboard"),
]

# Endpoints the generic hook must never log a "View" for: auth (already
# explicitly logged), the keepalive ping (fires every few minutes per user,
# pure noise), and anything serving raw bytes rather than a page a person
# consciously navigated to.
_AUDIT_VIEW_EXEMPT_ENDPOINTS = {
    "login", "logout", "static", "session_keepalive",
    "uploaded_file", "raw_uploaded_file", "jd_uploaded_file",
    "compare_result",  # explicitly logged as a "Compare" action instead, see compare_result()
}


def _audit_module_for_endpoint(endpoint):
    for prefix, module in _AUDIT_MODULE_BY_ENDPOINT_PREFIX:
        if endpoint.startswith(prefix):
            return module
    return endpoint.replace("_", " ").title()


@app.after_request
def _audit_generic_view(response):
    """Baseline "View" coverage for every authenticated GET page-load that
    isn't already explicitly logged (see the module docstring above) — this
    is what makes a brand-new route show up in the audit trail automatically
    with zero extra code, not just the routes someone remembered to
    instrument. Explicit log_audit() calls for Add/Edit/Delete/etc already
    cover their own request, so this only fires for GET to avoid double
    logging the same POST twice.
    """
    try:
        if (request.method == "GET" and response.status_code == 200
                and request.endpoint and request.endpoint not in _AUDIT_VIEW_EXEMPT_ENDPOINTS
                and session.get("user_id")
                and response.mimetype == "text/html"):
            record_id = next(iter(request.view_args.values()), None) if request.view_args else None
            log_audit(_audit_module_for_endpoint(request.endpoint), "View", record_id=record_id)
    except Exception:
        logger.exception("Generic audit view-logging hook failed")
    return response


_AUDIT_SORT_COLUMNS = {"created_at", "username", "module", "action", "status"}
_AUDIT_PER_PAGE = 25


def _audit_trail_query(args):
    """Shared filter/sort parsing + query building for the audit trail page
    and its CSV export, so the two can never drift out of sync on what
    "the currently filtered view" means.
    """
    search = (args.get("q") or "").strip()
    module_filter = (args.get("module") or "").strip()
    action_filter = (args.get("action") or "").strip()
    user_filter = (args.get("user") or "").strip()
    date_from = (args.get("date_from") or "").strip()
    date_to = (args.get("date_to") or "").strip()
    sort = args.get("sort", "created_at")
    if sort not in _AUDIT_SORT_COLUMNS:
        sort = "created_at"
    direction = "asc" if args.get("dir") == "asc" else "desc"

    conditions, params = [], []
    if search:
        like = f"%{search}%"
        conditions.append(
            "(username ILIKE %s OR record_label ILIKE %s OR module ILIKE %s "
            "OR action ILIKE %s OR COALESCE(details, '') ILIKE %s)"
        )
        params += [like, like, like, like, like]
    if module_filter:
        conditions.append("module = %s")
        params.append(module_filter)
    if action_filter:
        conditions.append("action = %s")
        params.append(action_filter)
    if user_filter:
        conditions.append("username = %s")
        params.append(user_filter)
    if date_from:
        conditions.append("created_at >= %s")
        params.append(date_from)
    if date_to:
        conditions.append("created_at < (%s::date + INTERVAL '1 day')")
        params.append(date_to)

    where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    filters = {
        "search": search, "module_filter": module_filter, "action_filter": action_filter,
        "user_filter": user_filter, "date_from": date_from, "date_to": date_to,
        "sort": sort, "direction": direction,
    }
    return where_sql, params, filters


@app.route("/audit-trail")
def audit_trail():
    if not _require_permission("view_audit_trail", write=False):
        abort(403)

    page = max(1, int(request.args.get("page", 1) or 1))
    where_sql, params, filters = _audit_trail_query(request.args)

    with db_conn() as conn:
        ensure_audit_log_table(conn)
        total = conn.execute(f"SELECT COUNT(*) AS c FROM audit_log {where_sql}", params).fetchone()["c"]
        rows = conn.execute(
            f"""
            SELECT * FROM audit_log {where_sql}
            ORDER BY {filters['sort']} {filters['direction']}, id {filters['direction']}
            LIMIT %s OFFSET %s
            """,
            params + [_AUDIT_PER_PAGE, (page - 1) * _AUDIT_PER_PAGE],
        ).fetchall()
        modules = [r["module"] for r in conn.execute(
            "SELECT DISTINCT module FROM audit_log ORDER BY module").fetchall()]
        actions = [r["action"] for r in conn.execute(
            "SELECT DISTINCT action FROM audit_log ORDER BY action").fetchall()]
        users = [r["username"] for r in conn.execute(
            "SELECT DISTINCT username FROM audit_log WHERE username IS NOT NULL ORDER BY username").fetchall()]

    total_pages = max(1, (total + _AUDIT_PER_PAGE - 1) // _AUDIT_PER_PAGE)
    return render_template(
        "audit_trail.html", rows=list(rows), total=total, page=page, total_pages=total_pages,
        per_page=_AUDIT_PER_PAGE, modules=modules, actions=actions, users=users, **filters,
    )


@app.route("/audit-trail/export.csv")
def audit_trail_export_csv():
    if not _require_permission("export_audit_trail"):
        abort(403)
    import csv as _csv

    where_sql, params, filters = _audit_trail_query(request.args)
    with db_conn() as conn:
        ensure_audit_log_table(conn)
        rows = conn.execute(
            f"SELECT * FROM audit_log {where_sql} ORDER BY created_at DESC",
            params,
        ).fetchall()

    buffer = StringIO()
    writer = _csv.writer(buffer)
    writer.writerow(["Date & Time", "User ID", "Username", "Role", "Module", "Action",
                      "Record ID", "Record/Entity", "Status", "IP Address", "User Agent", "Details"])
    for r in rows:
        writer.writerow([
            r["created_at"].strftime("%Y-%m-%d %H:%M:%S") if r["created_at"] else "",
            r["user_id"] or "", r["username"] or "", r["role"] or "", r["module"], r["action"],
            r["record_id"] or "", r["record_label"] or "", r["status"], r["ip_address"] or "",
            r["user_agent"] or "", r["details"] or "",
        ])
    log_audit("Audit Trail", "Export", status="Success", details=f"{len(rows)} rows")
    csv_bytes = buffer.getvalue().encode("utf-8-sig")  # BOM so Excel opens UTF-8 cleanly
    return send_file(
        BytesIO(csv_bytes), mimetype="text/csv", as_attachment=True,
        download_name=f"audit_trail_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.csv",
    )


# ── User Management & Authentication (new feature — additive only) ──────────

# 5-role RBAC per the client-supplied permission matrix (2026-08-13). Existing
# accounts are deliberately NOT auto-migrated onto these — a legacy `role`
# value (e.g. the old 'user') simply isn't a key in ROLE_PERMISSIONS below, so
# _permission_level() safely falls back to "none" for every gated feature
# until an admin explicitly reassigns that account via User Management. See
# user_management.html's "Needs Role Assignment" flag.
USER_ROLES = ["admin", "recruiter", "hiring_manager", "interviewer", "viewer_auditor"]
ROLE_LABELS = {
    "admin": "Admin",
    "recruiter": "Recruiter",
    "hiring_manager": "Hiring Manager",
    "interviewer": "Interviewer",
    "viewer_auditor": "Viewer / Auditor",
}
_DEFAULT_ADMIN_PASSWORD = "Admin@123"

# Permission matrix — one dict per feature, each mapping role -> "full" (view
# + write) | "view" (read-only, write controls hidden/blocked) | "none" (no
# access, route 403s). Built directly from the client's RBAC spreadsheet.
# Interviewer's "Assigned Only" case (update_interview_status) isn't a static
# level — it depends on the specific record — so it's handled separately, see
# _is_my_assigned_interview.
ROLE_PERMISSIONS = {
    "view_dashboard": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "full", "viewer_auditor": "full",
    },
    "view_candidate": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "full", "viewer_auditor": "view",
    },
    "schedule_interview": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "update_interview_status": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",  # interviewer's real access is "assigned" — see _is_my_assigned_interview
    },
    "view_jd": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "full", "viewer_auditor": "view",
    },
    "write_jd": {  # Add + Edit JD
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "delete_jd": {
        "admin": "full", "recruiter": "none", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    # Not a row in the client's spreadsheet — added for the JD approval
    # workflow (2026-08-13): Approve/Reject/Request Changes/Publish/Rollback
    # a JD. Single-level approval, role-based (no per-JD approver
    # assignment) — any Hiring Manager or Admin can act on any JD pending
    # approval, mirroring delete_jd's existing admin+hiring_manager pairing.
    "approve_jd": {
        "admin": "full", "recruiter": "none", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "view_requirement": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "full", "viewer_auditor": "view",
    },
    "write_requirement": {  # Add + Edit Requirement
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "delete_requirement": {
        "admin": "full", "recruiter": "none", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "view_users": {
        "admin": "full", "recruiter": "none", "hiring_manager": "none",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "write_users": {  # Add + Edit + Delete User
        "admin": "full", "recruiter": "none", "hiring_manager": "none",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "view_audit_trail": {
        "admin": "full", "recruiter": "view", "hiring_manager": "view",
        "interviewer": "view", "viewer_auditor": "view",
    },
    "export_audit_trail": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "view_profiles": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "full", "viewer_auditor": "view",
    },
    "write_profile": {  # upload/edit/save/bulk-upload
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "delete_profile": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    "download_resume": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "full", "viewer_auditor": "view",
    },
    "view_add_profile": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
    # Not a row in the client's spreadsheet — added per explicit follow-up
    # request (2026-08-14): Interviewer and Viewer/Auditor should not see or
    # use Compare Resume. Same access pattern as write_profile/schedule_interview.
    "compare_resume": {
        "admin": "full", "recruiter": "full", "hiring_manager": "full",
        "interviewer": "none", "viewer_auditor": "none",
    },
}


def _permission_level(key):
    return ROLE_PERMISSIONS.get(key, {}).get(session.get("role"), "none")


def _require_permission(key, write=True):
    """One inline `if not _require_permission(...): abort(403)` call at the
    top of a view function, same convention every route in this app uses.
    write=True requires "full" (a write/mutating action); write=False
    accepts "full" or "view" (read-only access to the page)."""
    level = _permission_level(key)
    if write:
        return level == "full"
    return level in ("full", "view")


def _is_my_assigned_interview(interviewer_field):
    """Interviewer's "Assigned Only" access to update_interview_status:
    interview_schedule.interviewer is free text (no FK to app_user exists),
    so this is a best-effort name match against the logged-in user's
    full_name — not a guaranteed identity link, since the field was never
    designed as one.

    Found live (2026-08-14): an interviewer's own account full_name ("Kumar")
    didn't exact-match the fuller name typed when the interview was
    scheduled ("Satheesh Kumar") — same person, different name format,
    correctly assigned to them, but incorrectly blocked. Fixed with a
    word-subset match: one side's words must be fully contained in the
    other's (not just any shared word), so "Kumar" ⊆ {"Satheesh","Kumar"}
    matches, but "Kumar Patel" vs "Kumar Singh" — a real, different-person
    collision — correctly does not (neither word set is a subset of the
    other).
    """
    me = (session.get("full_name") or "").strip().lower()
    them = (interviewer_field or "").strip().lower()
    if not me or not them:
        return False
    if me == them:
        return True
    me_words, them_words = set(me.split()), set(them.split())
    return me_words <= them_words or them_words <= me_words


def ensure_users_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_user (
            id            SERIAL PRIMARY KEY,
            username      TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name     TEXT DEFAULT '',
            email         TEXT DEFAULT '',
            role          TEXT DEFAULT 'user',
            is_active     BOOLEAN DEFAULT TRUE,
            created_at    TIMESTAMPTZ DEFAULT NOW(),
            updated_at    TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )
    count = conn.execute("SELECT COUNT(*) AS count FROM app_user").fetchone()["count"]
    if not count:
        conn.execute(
            """
            INSERT INTO app_user (username, password_hash, full_name, role)
            VALUES (%s, %s, %s, 'admin')
            """,
            ("admin", generate_password_hash(_DEFAULT_ADMIN_PASSWORD), "Administrator"),
        )


def _ensure_password_changed_at_col():
    """Bumped only on an actual password change (self-service reset or an
    admin editing a user's password) — never by unrelated profile edits.
    Snapshotted into the session at login and compared on every request by
    _require_login(); a mismatch means the password changed since this
    session's cookie was issued, so the old session is invalidated. This is
    the only way to achieve "log out other sessions" without a server-side
    session store — this app's sessions are Flask's default stateless
    signed cookies, there is no session table to delete rows from."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE app_user ADD COLUMN IF NOT EXISTS "
                "password_changed_at TIMESTAMPTZ DEFAULT NOW()"
            )
    except Exception:
        pass


_ensure_password_changed_at_col()


def _ensure_last_login_col():
    """NULL = never logged in. Stamped by login() and by
    change_expired_password() (which also ends in an authenticated session,
    just via the forced-password-change detour)."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE app_user ADD COLUMN IF NOT EXISTS last_login_at TIMESTAMPTZ"
            )
    except Exception:
        pass


_ensure_last_login_col()


def ensure_user_role_history_table(conn):
    """One row per role assignment/change for a user — the initial role at
    creation (old_role NULL) plus every subsequent change via user_edit().
    Never updated or deleted, so this is a append-only history."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS user_role_history (
            id          SERIAL PRIMARY KEY,
            user_id     INTEGER NOT NULL REFERENCES app_user(id) ON DELETE CASCADE,
            old_role    TEXT,
            new_role    TEXT NOT NULL,
            changed_by  INTEGER REFERENCES app_user(id),
            changed_at  TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )


def _ensure_account_lockout_cols():
    """failed_login_attempts: consecutive count, reset to 0 by any
    successful login. locked_at: NULL = not locked; a timestamp = locked
    (and when) — no auto-expiry, only user_unlock() (Admin-only) clears it."""
    try:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE app_user ADD COLUMN IF NOT EXISTS "
                "failed_login_attempts INTEGER DEFAULT 0"
            )
            conn.execute("ALTER TABLE app_user ADD COLUMN IF NOT EXISTS locked_at TIMESTAMPTZ")
    except Exception:
        pass


_ensure_account_lockout_cols()


def ensure_password_reset_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS password_reset_token (
            id          SERIAL PRIMARY KEY,
            user_id     INTEGER NOT NULL REFERENCES app_user(id) ON DELETE CASCADE,
            token_hash  TEXT NOT NULL UNIQUE,
            created_at  TIMESTAMPTZ DEFAULT NOW(),
            expires_at  TIMESTAMPTZ NOT NULL,
            used_at     TIMESTAMPTZ
        )
        """
    )


def ensure_password_history_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS password_history (
            id            SERIAL PRIMARY KEY,
            user_id       INTEGER NOT NULL REFERENCES app_user(id) ON DELETE CASCADE,
            password_hash TEXT NOT NULL,
            created_at    TIMESTAMPTZ DEFAULT NOW()
        )
        """
    )


def _generate_reset_token():
    """Returns (raw_token, token_hash). Only token_hash is ever persisted —
    the raw token exists solely in the emailed link and the user's browser,
    mirroring how a password itself is never stored in plain form."""
    raw = secrets.token_urlsafe(32)
    return raw, hashlib.sha256(raw.encode()).hexdigest()


def _validate_password_complexity(password):
    if len(password) < PASSWORD_MIN_LENGTH:
        return False, f"Password must be at least {PASSWORD_MIN_LENGTH} characters long."
    if not re.search(r"[a-z]", password):
        return False, "Password must contain at least one lowercase letter."
    if not re.search(r"[A-Z]", password):
        return False, "Password must contain at least one uppercase letter."
    if not re.search(r"\d", password):
        return False, "Password must contain at least one number."
    if not re.search(r"[^A-Za-z0-9]", password):
        return False, "Password must contain at least one special character."
    return True, ""


# Deliberately a simple "looks like an email" shape check (local@domain.tld),
# not a full RFC 5322 validator — good enough to catch typos/garbage without
# rejecting real-world addresses a stricter regex might choke on.
_EMAIL_FORMAT_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _validate_email_format(email):
    return bool(_EMAIL_FORMAT_RE.match(email))


def _password_reused(conn, user_id, new_password, current_hash):
    """True if new_password matches the current password or any of the last
    PASSWORD_HISTORY_COUNT historical ones for this user."""
    if current_hash and check_password_hash(current_hash, new_password):
        return True
    ensure_password_history_table(conn)
    rows = conn.execute(
        "SELECT password_hash FROM password_history WHERE user_id = %s "
        "ORDER BY created_at DESC LIMIT %s",
        (user_id, PASSWORD_HISTORY_COUNT),
    ).fetchall()
    return any(check_password_hash(row["password_hash"], new_password) for row in rows)


def _record_password_change(conn, user_id, old_hash):
    """Call this BEFORE overwriting app_user.password_hash with the new one
    — archives the outgoing hash for future reuse checks, prunes to the
    newest PASSWORD_HISTORY_COUNT, and bumps password_changed_at so every
    other outstanding session for this user is invalidated on its next
    request (see _ensure_password_changed_at_col)."""
    ensure_password_history_table(conn)
    if old_hash:
        conn.execute(
            "INSERT INTO password_history (user_id, password_hash) VALUES (%s, %s)",
            (user_id, old_hash),
        )
    conn.execute(
        """
        DELETE FROM password_history WHERE id IN (
            SELECT id FROM password_history WHERE user_id = %s
            ORDER BY created_at DESC OFFSET %s
        )
        """,
        (user_id, PASSWORD_HISTORY_COUNT),
    )
    conn.execute("UPDATE app_user SET password_changed_at = NOW() WHERE id = %s", (user_id,))


def _send_email(to_addr, subject, body):
    """Stdlib-only (no new dependency). If SMTP_HOST is unset (the default
    — this app has no mail relay configured anywhere), logs the message
    instead of attempting a real send, so this feature is fully testable
    without real SMTP credentials. Never raises — a failed/unsent email
    must not break the password-reset flow itself; the reset link/new
    password already took effect regardless of whether the notification
    email made it out.
    """
    if not to_addr:
        logger.warning(f"_send_email: no recipient address, skipping. subject={subject!r}")
        return
    if not SMTP_HOST:
        logger.info(f"[DEV MODE — no SMTP_HOST configured] Email to {to_addr}: {subject}\n{body}")
        return
    try:
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = SMTP_FROM_EMAIL
        msg["To"] = to_addr
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10) as server:
            if SMTP_USE_TLS:
                server.starttls()
            if SMTP_USERNAME:
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(SMTP_FROM_EMAIL, [to_addr], msg.as_string())
        logger.info(f"Email sent via {SMTP_HOST} to {to_addr}: {subject!r}")
    except Exception as e:
        logger.warning(f"_send_email failed to {to_addr}: {e}", exc_info=True)


@app.context_processor
def _inject_auth_context():
    return {
        "current_username": session.get("username"),
        "current_full_name": session.get("full_name"),
        "current_role": session.get("role"),
        "current_user_id": session.get("user_id"),
        "role_labels": ROLE_LABELS,
        # Single source of truth shared with the backend's own gating
        # (_require_permission) — templates hide a button with the exact
        # same permission check the route itself enforces, so there's no
        # separate/divergent notion of "can this role do X" in the UI layer.
        "can_view": lambda key: _permission_level(key) in ("full", "view"),
        "can_write": lambda key: _permission_level(key) == "full",
        # Same word-subset name match _require_permission's assigned-only
        # branch uses server-side — exposed here so templates never
        # reimplement (and potentially diverge from) this matching logic.
        "is_my_assigned_interview": _is_my_assigned_interview,
        # Consumed by base.html's inline bootstrap for static/session-timeout.js
        # — kept in one place (this processor) so every page agrees on the
        # configured duration without each template hardcoding it.
        "session_timeout_ms": SESSION_TIMEOUT_MINUTES * 60 * 1000,
        "session_warning_ms": SESSION_WARNING_SECONDS * 1000,
    }


_LOGIN_EXEMPT_ENDPOINTS = {"login", "static", "forgot_password", "reset_password", "change_expired_password"}


@app.before_request
def _require_login():
    """Single global auth gate for the whole app — also enforces the
    inactivity timeout here so every module gets it automatically, with no
    per-route code. `session['last_active']` is a sliding window: any
    request from an authenticated user pushes it forward, and if more than
    SESSION_TIMEOUT_MINUTES has elapsed since the last one, the session is
    invalidated and the user is bounced to login with an explanatory
    message — this is the server-side backstop that holds even if the
    browser tab was frozen/closed and the client-side warning JS
    (static/session-timeout.js) never got to run.
    """
    if request.endpoint is None or request.endpoint in _LOGIN_EXEMPT_ENDPOINTS:
        return None
    if not session.get("user_id"):
        return redirect(url_for("login", next=request.path))

    last_active_raw = session.get("last_active")
    now = datetime.now(timezone.utc)
    if last_active_raw:
        try:
            last_active = datetime.fromisoformat(last_active_raw)
        except ValueError:
            last_active = now
        idle_seconds = (now - last_active).total_seconds()
        if idle_seconds > SESSION_TIMEOUT_MINUTES * 60:
            log_audit("Auth", "Session Timeout", status="Success",
                      details=f"idle for {round(idle_seconds)}s (limit {SESSION_TIMEOUT_MINUTES}m)")
            session.clear()
            flash("Your session has expired due to inactivity. Please log in again.", "error")
            return redirect(url_for("login", next=request.path))

    # Password-change session invalidation: this app's sessions are
    # stateless signed cookies (no server-side session store/table), so
    # "log out all other sessions after a password reset" is implemented by
    # snapshotting app_user.password_changed_at into the session at login
    # and comparing it here on every request. A mismatch means the password
    # changed (self-service reset or an admin edit) since this cookie was
    # issued, so it's invalidated exactly like the inactivity-timeout branch
    # above. One extra indexed lookup per authenticated request — this
    # before_request hook made zero DB calls previously.
    stored_pca = session.get("password_changed_at")
    if stored_pca:
        with db_conn() as conn:
            row = conn.execute(
                "SELECT password_changed_at FROM app_user WHERE id = %s", (session["user_id"],)
            ).fetchone()
        current_pca = row["password_changed_at"].isoformat() if row and row["password_changed_at"] else None
        if current_pca and current_pca != stored_pca:
            log_audit("Auth", "Session Invalidated", status="Success",
                      details="password changed since this session was issued")
            session.clear()
            flash("Your password was changed. Please log in again.", "error")
            return redirect(url_for("login", next=request.path))

    session["last_active"] = now.isoformat()
    return None


@app.route("/api/session/keepalive")
def session_keepalive():
    """Pinged by static/session-timeout.js when the user clicks "Continue
    Session" on the inactivity warning dialog. No-op beyond a plain 200 —
    _require_login() above already refreshed session['last_active'] for
    this request before this view even runs, since it's just another
    authenticated request.
    """
    return jsonify({"ok": True, "timeout_ms": SESSION_TIMEOUT_MINUTES * 60 * 1000})


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        with db_conn() as conn:
            ensure_users_table(conn)
            row = conn.execute(
                "SELECT * FROM app_user WHERE LOWER(username) = LOWER(%s)", (username,)
            ).fetchone()
        # A locked account is rejected unconditionally — checked BEFORE the
        # password comparison, so even the correct password never succeeds
        # while locked_at is set. No auto-expiry; only user_unlock() (Admin
        # -only, User Management) clears this. Admin accounts are exempt
        # (see the failure-tracking block below for why) — this check also
        # covers the edge case of an admin account somehow already carrying
        # a locked_at value (e.g. from before this exemption existed).
        if row and row["locked_at"] and row["role"] != "admin":
            log_audit("Auth", "Login", record_label=row["username"], status="Failure",
                      details="Account locked")
            flash("Your account has been locked due to too many failed login attempts. "
                  "Please contact your administrator to unlock it.", "error")
            return render_template("login.html", next=request.form.get("next", ""))

        if row and row["is_active"] and check_password_hash(row["password_hash"], password):
            with db_conn() as conn:
                conn.execute(
                    "UPDATE app_user SET failed_login_attempts = 0, last_login_at = NOW() WHERE id = %s",
                    (row["id"],),
                )
            # Credentials are correct, but a password older than
            # PASSWORD_MAX_AGE_DAYS blocks normal login — only a temporary
            # marker is set (not the real session), so _require_login() still
            # treats this user as logged out everywhere except the mandatory
            # change-password screen (see change_expired_password() and its
            # entry in _LOGIN_EXEMPT_ENDPOINTS).
            if PASSWORD_MAX_AGE_DAYS and row["password_changed_at"]:
                age_days = (datetime.now(timezone.utc) - row["password_changed_at"]).days
                if age_days >= PASSWORD_MAX_AGE_DAYS:
                    session["pending_password_change_user_id"] = row["id"]
                    log_audit("Auth", "Login", record_label=row["username"], status="Failure",
                              details=f"Password expired ({age_days} days old) — must be changed")
                    flash("Your password has expired and must be changed before you can continue.", "error")
                    return redirect(url_for("change_expired_password"))
            session["user_id"] = row["id"]
            session["username"] = row["username"]
            session["full_name"] = row["full_name"]
            session["role"] = row["role"]
            session["last_active"] = datetime.now(timezone.utc).isoformat()
            session["password_changed_at"] = (
                row["password_changed_at"].isoformat() if row["password_changed_at"] else None
            )
            log_audit("Auth", "Login", record_label=row["username"], status="Success")
            return redirect(request.form.get("next") or url_for("dashboard"))

        if row:
            # A real account, wrong password (or disabled) — track the
            # consecutive-failure count and lock once it hits the
            # configured threshold. A nonexistent username (the `else`
            # implied by `not row`) never reaches here, so it never gets an
            # attempts-remaining counter — nothing to attach a count to,
            # and it avoids adding a *second* signal beyond the generic
            # message for a username that doesn't exist.
            #
            # Admin accounts are exempt from ever being LOCKED (found live
            # 2026-08-14: the only admin account got itself locked out,
            # which meant no one was left who could unlock it — a
            # self-inflicted total-lockout risk this app can't recover
            # from without direct DB access). The failure count is still
            # tracked for audit visibility, it just never flips locked_at.
            is_admin_account = row["role"] == "admin"
            with db_conn() as conn:
                new_count = conn.execute(
                    "UPDATE app_user SET failed_login_attempts = failed_login_attempts + 1 "
                    "WHERE id = %s RETURNING failed_login_attempts",
                    (row["id"],),
                ).fetchone()["failed_login_attempts"]
                if new_count >= MAX_FAILED_LOGIN_ATTEMPTS and not is_admin_account:
                    conn.execute("UPDATE app_user SET locked_at = NOW() WHERE id = %s", (row["id"],))
            if new_count >= MAX_FAILED_LOGIN_ATTEMPTS and not is_admin_account:
                _send_email(
                    row["email"],
                    "Your account has been locked — Resume Profile",
                    f"Hello {row['full_name'] or row['username']},\n\n"
                    f"Your account was locked after {new_count} consecutive failed login attempts. "
                    "Please contact your administrator to unlock it.\n\n"
                    "If this wasn't you, your administrator should also verify no one else is "
                    "trying to access your account.",
                )
                log_audit("Auth", "Login", record_label=row["username"], status="Failure",
                          details=f"Account locked after {new_count} failed attempts")
                flash("Your account has been locked due to too many failed login attempts. "
                      "Please contact your administrator to unlock it.", "error")
                return render_template("login.html", next=request.form.get("next", ""))
            if is_admin_account:
                # No attempts-remaining/lock-threat messaging for admin — it
                # would be misleading since this account can never actually
                # lock. Still logged for audit so unusually high failure
                # counts on an admin account remain visible/investigable.
                log_audit("Auth", "Login", record_label=row["username"], status="Failure",
                          details=f"Failed attempt {new_count} (admin — exempt from lockout)")
                flash("Invalid username or password.", "error")
                return render_template("login.html", next=request.form.get("next", ""))
            remaining = MAX_FAILED_LOGIN_ATTEMPTS - new_count
            log_audit("Auth", "Login", record_label=row["username"], status="Failure",
                      details=f"Failed attempt {new_count}/{MAX_FAILED_LOGIN_ATTEMPTS}")
            flash(f"Invalid username or password. {remaining} attempt(s) remaining "
                  "before your account is locked.", "error")
            return render_template("login.html", next=request.form.get("next", ""))

        log_audit("Auth", "Login", record_label=username, status="Failure",
                  details="Invalid username or password")
        flash("Invalid username or password.", "error")
        return render_template("login.html", next=request.form.get("next", ""))
    with db_conn() as conn:
        ensure_users_table(conn)
    return render_template("login.html", next=request.args.get("next", ""))


@app.route("/logout")
def logout():
    username = session.get("username")
    session.clear()
    if request.args.get("reason") == "inactivity":
        log_audit("Auth", "Session Timeout", record_label=username, status="Success",
                  details="client-side inactivity timer")
        flash("Your session has expired due to inactivity. Please log in again.", "error")
    else:
        log_audit("Auth", "Logout", record_label=username, status="Success")
    return redirect(url_for("login"))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        # Same confirmation message regardless of whether the email is
        # registered — avoids leaking which addresses have accounts. This
        # still satisfies "an unregistered email cannot be used to reset a
        # password": nothing is created/sent unless a real, active account
        # matches, only the user-facing wording doesn't reveal which case
        # occurred.
        generic_message = (
            "If that email is registered, we've sent a password reset link to it. "
            "The link is valid for %d minutes." % PASSWORD_RESET_TOKEN_MINUTES
        )
        if email:
            with db_conn() as conn:
                ensure_users_table(conn)
                ensure_password_reset_table(conn)
                user = conn.execute(
                    "SELECT * FROM app_user WHERE LOWER(email) = LOWER(%s) AND is_active = TRUE",
                    (email,),
                ).fetchone()
                if user:
                    # Only the latest requested link should ever work —
                    # invalidate any earlier unused ones for this user first.
                    conn.execute(
                        "UPDATE password_reset_token SET used_at = NOW() "
                        "WHERE user_id = %s AND used_at IS NULL",
                        (user["id"],),
                    )
                    raw_token, token_hash = _generate_reset_token()
                    expires_at = datetime.now(timezone.utc) + timedelta(minutes=PASSWORD_RESET_TOKEN_MINUTES)
                    conn.execute(
                        "INSERT INTO password_reset_token (user_id, token_hash, expires_at) "
                        "VALUES (%s, %s, %s)",
                        (user["id"], token_hash, expires_at),
                    )
                    reset_link = url_for("reset_password", token=raw_token, _external=True)
                    _send_email(
                        user["email"],
                        "Password Reset Request — Resume Profile",
                        f"Hello {user['full_name'] or user['username']},\n\n"
                        f"A password reset was requested for your account. Click the link below "
                        f"to set a new password. This link is valid for {PASSWORD_RESET_TOKEN_MINUTES} "
                        f"minutes and can only be used once.\n\n{reset_link}\n\n"
                        "If you didn't request this, you can safely ignore this email.",
                    )
                    log_audit("Auth", "Forgot Password Requested", record_label=user["username"],
                              status="Success")
                else:
                    log_audit("Auth", "Forgot Password Requested", record_label=email,
                              status="Failure", details="No matching active account")
        flash(generic_message, "success")
        return redirect(url_for("login"))
    return render_template("forgot_password.html")


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    token_hash = hashlib.sha256(token.encode()).hexdigest()
    with db_conn() as conn:
        ensure_password_reset_table(conn)
        token_row = conn.execute(
            "SELECT * FROM password_reset_token WHERE token_hash = %s", (token_hash,)
        ).fetchone()

    error = None
    if not token_row:
        error = "This password reset link is invalid."
    elif token_row["used_at"]:
        error = "This password reset link has already been used. Please request a new one."
    elif token_row["expires_at"] < datetime.now(timezone.utc):
        error = "This password reset link has expired. Please request a new one."

    if error:
        return render_template("reset_password.html", error=error)

    if request.method == "POST":
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if new_password != confirm_password:
            flash("New password and confirmation do not match.", "error")
            return render_template("reset_password.html", error=None)

        ok, complexity_error = _validate_password_complexity(new_password)
        if not ok:
            flash(complexity_error, "error")
            return render_template("reset_password.html", error=None)

        with db_conn() as conn:
            ensure_password_reset_table(conn)
            # Re-fetch and re-validate — defends against a race between the
            # GET that rendered this form and this POST (e.g. the same link
            # opened twice, or it expiring/getting used in between).
            token_row = conn.execute(
                "SELECT * FROM password_reset_token WHERE token_hash = %s", (token_hash,)
            ).fetchone()
            if not token_row or token_row["used_at"] or token_row["expires_at"] < datetime.now(timezone.utc):
                flash("This password reset link is no longer valid. Please request a new one.", "error")
                return redirect(url_for("forgot_password"))

            user = conn.execute(
                "SELECT * FROM app_user WHERE id = %s", (token_row["user_id"],)
            ).fetchone()
            if not user:
                flash("Account not found.", "error")
                return redirect(url_for("forgot_password"))

            if _password_reused(conn, user["id"], new_password, user["password_hash"]):
                flash(f"You cannot reuse any of your last {PASSWORD_HISTORY_COUNT} passwords.", "error")
                return render_template("reset_password.html", error=None)

            _record_password_change(conn, user["id"], user["password_hash"])
            conn.execute(
                "UPDATE app_user SET password_hash = %s, updated_at = NOW() WHERE id = %s",
                (generate_password_hash(new_password), user["id"]),
            )
            conn.execute(
                "UPDATE password_reset_token SET used_at = NOW() WHERE id = %s", (token_row["id"],)
            )

        _send_email(
            user["email"],
            "Your password was changed — Resume Profile",
            f"Hello {user['full_name'] or user['username']},\n\n"
            "This is a confirmation that your password was just changed. If you did not make "
            "this change, contact your administrator immediately.\n\n"
            "You have been logged out of all other sessions and must log in again with your new password.",
        )
        log_audit("Auth", "Password Reset", record_label=user["username"], status="Success")
        flash("Your password has been reset. Please log in with your new password.", "success")
        return redirect(url_for("login"))

    return render_template("reset_password.html", error=None)


@app.route("/change-expired-password", methods=["GET", "POST"])
def change_expired_password():
    """Reached only via login()'s expiry redirect, which sets
    session['pending_password_change_user_id'] instead of establishing a
    real session — so _require_login() still blocks every other page for
    this user until they set a new password here. Reuses the exact same
    complexity/reuse/history calls reset_password() and user_edit() already
    use, just with a different entry point (a temporary marker instead of a
    one-time emailed token, since the user already proved their identity
    with their current password moments ago)."""
    user_id = session.get("pending_password_change_user_id")
    if not user_id:
        return redirect(url_for("login"))
    with db_conn() as conn:
        user = conn.execute("SELECT * FROM app_user WHERE id = %s", (user_id,)).fetchone()
    if not user:
        session.pop("pending_password_change_user_id", None)
        return redirect(url_for("login"))

    if request.method == "POST":
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if new_password != confirm_password:
            flash("New password and confirmation do not match.", "error")
            return render_template("change_expired_password.html")

        ok, complexity_error = _validate_password_complexity(new_password)
        if not ok:
            flash(complexity_error, "error")
            return render_template("change_expired_password.html")

        with db_conn() as conn:
            if _password_reused(conn, user["id"], new_password, user["password_hash"]):
                flash(f"You cannot reuse any of your last {PASSWORD_HISTORY_COUNT} passwords.", "error")
                return render_template("change_expired_password.html")
            _record_password_change(conn, user["id"], user["password_hash"])
            conn.execute(
                "UPDATE app_user SET password_hash = %s, updated_at = NOW(), last_login_at = NOW() WHERE id = %s",
                (generate_password_hash(new_password), user["id"]),
            )
            # Re-fetch rather than stamping a Python-side timestamp — this
            # must match the DB's stored value exactly, since _require_login()
            # compares this snapshot against a fresh DB read on every request.
            updated_user = conn.execute(
                "SELECT password_changed_at FROM app_user WHERE id = %s", (user["id"],)
            ).fetchone()

        session.pop("pending_password_change_user_id", None)
        session["user_id"] = user["id"]
        session["username"] = user["username"]
        session["full_name"] = user["full_name"]
        session["role"] = user["role"]
        session["last_active"] = datetime.now(timezone.utc).isoformat()
        session["password_changed_at"] = (
            updated_user["password_changed_at"].isoformat() if updated_user["password_changed_at"] else None
        )
        log_audit("Auth", "Password Changed (Expired)", record_label=user["username"], status="Success")
        flash("Your password has been changed.", "success")
        return redirect(url_for("dashboard"))

    return render_template("change_expired_password.html")


# ── User Management routes (admin only) ──────────────────────────────────────

@app.route("/users")
def user_management():
    if not _require_permission("view_users", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_users_table(conn)
        users = conn.execute("SELECT * FROM app_user ORDER BY created_at ASC").fetchall()
    return render_template("user_management.html", users=list(users))

@app.route("/users/add", methods=["GET", "POST"])
def user_add():
    if not _require_permission("write_users"):
        abort(403)
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip()
        role = request.form.get("role", "viewer_auditor")
        if role not in USER_ROLES:
            role = "viewer_auditor"
        if not username or not password:
            flash("Username and password are required.", "error")
            return render_template("user_form.html", user=None, roles=USER_ROLES)
        # Email itself stays optional (unchanged) — but if one is entered, it
        # must at least look like a real address.
        if email and not _validate_email_format(email):
            flash("Please enter a valid email address.", "error")
            return render_template("user_form.html", user=None, roles=USER_ROLES)
        # Same complexity policy the self-service reset flow enforces (see
        # _validate_password_complexity) — introduced alongside password
        # reset since no policy existed anywhere before; applying it only to
        # self-service reset would let an admin-created password be weaker
        # than a self-reset one.
        ok, complexity_error = _validate_password_complexity(password)
        if not ok:
            flash(complexity_error, "error")
            return render_template("user_form.html", user=None, roles=USER_ROLES)
        with db_conn() as conn:
            ensure_users_table(conn)
            existing = conn.execute(
                "SELECT id FROM app_user WHERE LOWER(username) = LOWER(%s)", (username,)
            ).fetchone()
            if existing:
                flash(f"Username '{username}' is already taken.", "error")
                return render_template("user_form.html", user=None, roles=USER_ROLES)
            if email:
                existing_email = conn.execute(
                    "SELECT id FROM app_user WHERE LOWER(email) = LOWER(%s)", (email,)
                ).fetchone()
                if existing_email:
                    flash(f"Email '{email}' is already in use by another account.", "error")
                    return render_template("user_form.html", user=None, roles=USER_ROLES)
            cursor = conn.execute(
                """
                INSERT INTO app_user (username, password_hash, full_name, email, role)
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id
                """,
                (username, generate_password_hash(password), full_name, email, role),
            )
            new_user_id = cursor.fetchone()["id"]
            # Initial role assignment — old_role NULL marks this as account
            # creation rather than a later change, so history is complete
            # from the start, not just from the first edit onward.
            ensure_user_role_history_table(conn)
            conn.execute(
                "INSERT INTO user_role_history (user_id, old_role, new_role, changed_by) "
                "VALUES (%s, NULL, %s, %s)",
                (new_user_id, role, session.get("user_id")),
            )
        log_audit("User Management", "Add", record_id=new_user_id, record_label=username)
        flash(f"User '{username}' added.", "success")
        return redirect(url_for("user_management"))
    return render_template("user_form.html", user=None, roles=USER_ROLES)


@app.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
def user_edit(user_id):
    if not _require_permission("write_users"):
        abort(403)
    with db_conn() as conn:
        ensure_users_table(conn)
        if request.method == "POST":
            full_name = request.form.get("full_name", "").strip()
            email = request.form.get("email", "").strip()
            role = request.form.get("role", "viewer_auditor")
            if role not in USER_ROLES:
                role = "viewer_auditor"
            # Captured before the UPDATE below overwrites it — compared after
            # the save succeeds to decide whether a role-history row is needed.
            existing_role_row = conn.execute("SELECT role FROM app_user WHERE id = %s", (user_id,)).fetchone()
            old_role = existing_role_row["role"] if existing_role_row else None
            # Email itself stays optional (unchanged) — but if one is entered,
            # it must at least look like a real address, and must not already
            # belong to a different account.
            if email and not _validate_email_format(email):
                flash("Please enter a valid email address.", "error")
                user_row = conn.execute("SELECT * FROM app_user WHERE id = %s", (user_id,)).fetchone()
                return render_template("user_form.html", user=dict(user_row), roles=USER_ROLES)
            if email:
                existing_email = conn.execute(
                    "SELECT id FROM app_user WHERE LOWER(email) = LOWER(%s) AND id != %s", (email, user_id)
                ).fetchone()
                if existing_email:
                    flash(f"Email '{email}' is already in use by another account.", "error")
                    user_row = conn.execute("SELECT * FROM app_user WHERE id = %s", (user_id,)).fetchone()
                    return render_template("user_form.html", user=dict(user_row), roles=USER_ROLES)
            is_active = bool(request.form.get("is_active"))
            new_password = request.form.get("password", "")
            if new_password:
                ok, complexity_error = _validate_password_complexity(new_password)
                if not ok:
                    flash(complexity_error, "error")
                    user_row = conn.execute("SELECT * FROM app_user WHERE id = %s", (user_id,)).fetchone()
                    return render_template("user_form.html", user=dict(user_row), roles=USER_ROLES)
                current = conn.execute(
                    "SELECT password_hash FROM app_user WHERE id = %s", (user_id,)
                ).fetchone()
                if _password_reused(conn, user_id, new_password, current["password_hash"] if current else None):
                    flash(f"That user cannot reuse any of their last {PASSWORD_HISTORY_COUNT} passwords.", "error")
                    user_row = conn.execute("SELECT * FROM app_user WHERE id = %s", (user_id,)).fetchone()
                    return render_template("user_form.html", user=dict(user_row), roles=USER_ROLES)
                _record_password_change(conn, user_id, current["password_hash"] if current else None)
                conn.execute(
                    """
                    UPDATE app_user SET full_name=%s, email=%s, role=%s, is_active=%s,
                        password_hash=%s, updated_at=NOW()
                    WHERE id=%s
                    """,
                    (full_name, email, role, is_active, generate_password_hash(new_password), user_id),
                )
            else:
                conn.execute(
                    """
                    UPDATE app_user SET full_name=%s, email=%s, role=%s, is_active=%s, updated_at=NOW()
                    WHERE id=%s
                    """,
                    (full_name, email, role, is_active, user_id),
                )
            if old_role != role:
                ensure_user_role_history_table(conn)
                conn.execute(
                    "INSERT INTO user_role_history (user_id, old_role, new_role, changed_by) "
                    "VALUES (%s, %s, %s, %s)",
                    (user_id, old_role, role, session.get("user_id")),
                )
            if session.get("user_id") == user_id:
                session["full_name"] = full_name
                session["role"] = role
            log_audit("User Management", "Edit", record_id=user_id, record_label=full_name)
            flash("User updated.", "success")
            return redirect(url_for("user_management"))
        user = conn.execute("SELECT * FROM app_user WHERE id = %s", (user_id,)).fetchone()
        if not user:
            return "User not found", 404
    return render_template("user_form.html", user=dict(user), roles=USER_ROLES)


@app.route("/users/<int:user_id>/role-history")
def user_role_history(user_id):
    if not _require_permission("view_users", write=False):
        abort(403)
    with db_conn() as conn:
        ensure_users_table(conn)
        ensure_user_role_history_table(conn)
        user = conn.execute("SELECT id, username, full_name FROM app_user WHERE id = %s", (user_id,)).fetchone()
        if not user:
            return "User not found", 404
        history = conn.execute(
            "SELECT h.*, u.full_name AS changed_by_name FROM user_role_history h "
            "LEFT JOIN app_user u ON u.id = h.changed_by "
            "WHERE h.user_id = %s ORDER BY h.changed_at DESC",
            (user_id,),
        ).fetchall()
    return render_template("user_role_history.html", user=dict(user), history=list(history),
                           role_labels=ROLE_LABELS)


@app.route("/users/<int:user_id>/delete", methods=["POST"])
def user_delete(user_id):
    if not _require_permission("write_users"):
        abort(403)
    if session.get("user_id") == user_id:
        flash("You cannot delete your own account while logged in.", "error")
        return redirect(url_for("user_management"))
    with db_conn() as conn:
        row = conn.execute("SELECT username FROM app_user WHERE id = %s", (user_id,)).fetchone()
        conn.execute("DELETE FROM app_user WHERE id = %s", (user_id,))
    log_audit("User Management", "Delete", record_id=user_id, record_label=row["username"] if row else None)
    flash("User deleted.", "success")
    return redirect(url_for("user_management"))


@app.route("/users/<int:user_id>/unlock", methods=["POST"])
def user_unlock(user_id):
    if not _require_permission("write_users"):
        abort(403)
    with db_conn() as conn:
        conn.execute(
            "UPDATE app_user SET failed_login_attempts = 0, locked_at = NULL WHERE id = %s",
            (user_id,),
        )
        row = conn.execute("SELECT username FROM app_user WHERE id = %s", (user_id,)).fetchone()
    log_audit("User Management", "Unlock", record_id=user_id, record_label=row["username"] if row else None)
    flash("Account unlocked.", "success")
    return redirect(url_for("user_management"))


try:
    with db_conn() as conn:
        ensure_audit_log_table(conn)
except Exception as e:
    logger.error(f"Audit log table creation failed: {e}")


if __name__ == "__main__":
    app.run(
        debug=True,
        use_reloader=False,
        # Without this, Werkzeug's dev server handles one request at a time —
        # a single slow AI-assessment call (real local-LLM inference, up to a
        # minute) blocks every other request behind it, including unrelated
        # page loads and static assets, which is what made the app look like
        # it "keeps loading forever" under any concurrent use. Safe to enable
        # here: every request opens its own DB connection (see db_conn()),
        # so there's no shared connection for concurrent requests to race on.
        threaded=True,
        port=5001,
        host="0.0.0.0"
    )