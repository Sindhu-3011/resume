from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, abort, send_file
import psycopg2
from io import BytesIO
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
from psycopg2.extras import RealDictCursor
from pathlib import Path
import re
import zipfile
import json
import logging
import os
import urllib.request
import urllib.error
from contextlib import contextmanager
from werkzeug.utils import secure_filename
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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "resume-profile-secret-key")

@app.after_request
def set_cache_control(response):
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

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
    "summary": [
        "summary", "professional summary", "profile summary", "profile",
        "objective", "career objective",
        "about", "about me", "overview", "executive summary", "career profile",
        "professional profile", "personal profile", "career summary",
        "professional overview", "career overview",
    ],
    "skills": [
        "skills", "technical skills", "core competencies", "key skills", "skill set",
        "competencies", "technologies", "tools and technologies", "technical expertise",
        "areas of expertise", "technical competencies", "key competencies",
        "skills and competencies", "skills and technologies", "tools technologies",
        "applications summary", "application summary", "applications", "tools used",
        "tools and applications", "technical tools", "software tools", "tools and software",
    ],
    "experience": [
        "experience", "work experience", "professional experience", "employment history",
        "career history", "work history", "employment", "professional background",
        "job experience", "relevant experience", "industry experience",
        "work experience details", "experience summary", "employment details",
        "professional history", "career experience", "professional work experience",
        "work experience summary", "key experiences",
    ],
    "education": [
        "education", "academic background", "qualifications", "academic qualification",
        "academics", "educational background", "academic credentials",
        "educational qualifications", "academic details",
    ],
    "certifications": [
        "certifications", "certificates", "certification", "training", "trainings",
        "courses", "licenses", "professional development", "professional certifications",
        "awards and certifications", "achievements", "awards", "awards certifications",
    ],
    "projects": [
        "projects", "key projects", "project experience", "assignments",
        "notable projects", "project highlights", "academic projects",
        "personal projects", "relevant project experience",
        "project details", "project organizational details",
        "organizational details", "relevant project organizational details",
        "project and organizational details", "project organizational details roles and responsibilities",
        "client details", "relevant projects",
    ],
}

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
    # Both columns must hold a real share of the content.
    left  = sum(1 for w in words if w[2] <= best_x)
    right = sum(1 for w in words if w[0] >= best_x)
    if left < n * 0.15 or right < n * 0.15:
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


def _pymupdf_page_text(page):
    """Extract readable text from a page, handling 2-column layouts correctly.

    If the page has a clear 2-column structure (detected via block positions),
    the left and right columns are processed independently and concatenated —
    preventing content from the two columns from being mixed on the same line.
    Falls back to single-stream extraction for 1-column pages.
    """
    raw_words = page.get_text("words")
    if not raw_words:
        return ""

    # Primary: conservative block-based detection. Fallback: word-gutter detection
    # for interleaved layouts where wide left-column lines span the page centre.
    split_x = _detect_two_col_split(page) or _detect_col_gutter_words(page)
    if split_x:
        left_words  = [w for w in raw_words if w[0] <  split_x]
        right_words = [w for w in raw_words if w[0] >= split_x]
        return (_words_to_text(left_words) + "\n" + _words_to_text(right_words)).strip()

    return _words_to_text(raw_words)


def _fix_wrapped_email(text):
    """Join emails that the PDF wrapped mid-TLD across two lines.

    e.g.  "sindhusundaramoorthy30@gmail.c"  +  "om"  →  "…@gmail.com"
    """
    lines = text.split("\n")
    result = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if i + 1 < len(lines) and "@" in line:
            next_line = lines[i + 1].strip()
            # Next line is 1-4 plain letters — likely a truncated TLD suffix
            if re.match(r"^[a-zA-Z]{1,4}$", next_line):
                # Current line ends with a partial domain: @host.xy  (TLD ≤ 2 chars)
                if re.search(r"@[A-Za-z0-9.\-]+\.[A-Za-z]{1,2}$", line):
                    line = line.rstrip() + next_line
                    i += 1   # consume the continuation line
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
                if ocr_text:
                    # Extract only email/phone from OCR and inject as labelled lines.
                    # Never prepend raw OCR — it duplicates content that is already in
                    # the text layer, which corrupts summary/section parsing.
                    contact_lines = []
                    if not has_email:
                        em = _extract_email(ocr_text)
                        if em:
                            contact_lines.append(f"Email: {em}")
                    if not has_phone:
                        ph = _extract_phone(ocr_text)
                        if ph:
                            contact_lines.append(f"Phone: {ph}")
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

        # Section label (matches SECTION_ALIASES, short text)
        sec = canonical_section_name(first)
        if sec and len(first.split()) <= 4:
            box["kind"]    = "label"
            box["section"] = sec
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
    _add(b for b in boxes if b.get("kind") == "education")

    # 6. Skills
    sk_lbl = next((b for b in boxes if b.get("kind") == "label" and b.get("section") == "skills"), None)
    if sk_lbl:
        parts.append(sk_lbl["lines"][0])
    _add(b for b in boxes if b.get("kind") == "skills")

    return "\n".join(parts)


def extract_text_from_docx(path):
    try:
        logger.info(f"Extracting text from DOCX: {path}")
        doc = Document(str(path))
        parts = []
        from docx.oxml.ns import qn
        from docx.table import Table as DocxTable
        from docx.text.paragraph import Paragraph as DocxParagraph
        for child in doc.element.body:
            if child.tag == qn("w:p"):
                para = DocxParagraph(child, doc)
                if para.text.strip():
                    parts.append(para.text.strip())
            elif child.tag == qn("w:tbl"):
                tbl = DocxTable(child, doc)
                for row in tbl.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        parts.append(row_text)

        # If the standard body extraction yielded nothing, the file likely uses
        # a floating text-box canvas layout — try the canvas extractor.
        if not "".join(parts).strip():
            logger.info("No body text found; trying canvas/text-box extraction")
            canvas_text = _extract_canvas_docx(doc.element.body)
            if canvas_text.strip():
                logger.info(f"Canvas extraction yielded {len(canvas_text)} chars")
                return canvas_text

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
    # But exclude if it looks like a title (contains job keywords)
    job_keywords = {
        "engineer", "manager", "lead", "analyst", "developer", "designer",
        "specialist", "consultant", "director", "officer", "executive"
    }
    has_job_keyword = any(w.lower() in job_keywords for w in words)
    if not has_job_keyword and len(words) <= 4:
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
        "achievements", "soft skills",
    }
    _SKIP_STARTSWITH = {
        "websites portfolios and profiles", "websites and profiles",
        "websites portfolios", "and profiles",
    }
    current = None
    bucket = []
    for line in lines:
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
                sections[current] = "\n".join(bucket).strip()
            current = None
            bucket = []
            continue

        section = canonical_section_name(line)
        inline_remainder = None

        if not section:
            # An unlabeled company-name line starts the experience section ONLY when no
            # section is active yet (a resume that opens with jobs and no "Experience"
            # heading). Once a section is active, a company-looking line is treated as
            # that section's content — otherwise capitalised lines in projects/education/
            # awards would hijack the flow and shred the real experience section.
            if current is None and _looks_like_company_name(line):
                section = "experience"
            else:
                words = line.split()
                for n in range(min(3, len(words)), 0, -1):
                    prefix_section = canonical_section_name(" ".join(words[:n]))
                    if prefix_section:
                        remainder = " ".join(words[n:]).strip()
                        # Single-word prefix (e.g. "Experience") only counts as a heading when:
                        #   • it IS the entire line (no remainder), OR
                        #   • it ends with a separator like "Skills:" — not a content sentence.
                        if n == 1 and remainder and not words[0].endswith(':'):
                            break  # "Experience on all ALM modules" → content, not heading
                        section = prefix_section
                        inline_remainder = remainder if remainder else None
                        break

        if section:
            if section == current:
                # "Summary" appearing mid-summary (e.g. wrapped "Validation Summary Report.")
                # — ignore it so the current section doesn't restart and lose its content.
                continue
            if current and bucket:
                sections[current] = "\n".join(bucket).strip()
            current = section
            bucket = []
            if inline_remainder:
                bucket.append(inline_remainder)
            continue

        if current:
            bucket.append(line)

    if current and bucket:
        sections[current] = "\n".join(bucket).strip()
    return sections


def parse_label_value(lines, labels):
    label_pattern = "|".join(re.escape(label) for label in labels)
    pattern = re.compile(rf"(?:^|\b)({label_pattern})\s*[:\-]\s*(.+)$", re.I)
    for line in lines[:30]:
        match = pattern.search(line)
        if match:
            return match.group(2).strip()
    return ""


def _extract_email(text):
    """Return the first valid email found in text, handling common OCR artifacts."""
    # Pass 1: standard clean email
    m = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    if m:
        return m.group().strip().rstrip(".,;)")

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
            return f"{local}@{domain}"
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
    "total", "experience", "current", "employer", "designation", "qualification",
    "key", "skills", "competencies", "summary", "objectives", "highlights",
    "personal", "information", "education", "project", "work", "certification",
    "achievement", "language", "hobby", "extracurricular", "tools", "applications",
    "core", "section", "header", "details", "accomplishment", "award",
})


def _looks_like_name(line):
    """Return True only if *line* could plausibly be a person's full name."""
    line = line.strip()
    if not line or len(line) > 55:
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
    return bool(words) and words[0][0].isupper()


_TITLE_SECTION_WORDS = frozenset({
    "summary", "profile", "objective", "overview", "experience", "education",
    "skills", "certifications", "projects", "employment", "history", "background",
    "qualifications", "competencies", "expertise", "achievements", "declaration",
    "references", "hobbies", "interests", "languages",
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
            if 2 <= len(words) <= 7:
                return candidate
    # Pattern 2: as a/an <Title>
    m2 = re.search(
        r'\bas\s+an?\s+([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+){0,5})',
        text[:700],
    )
    if m2:
        candidate = m2.group(1).strip()
        if 1 <= len(candidate.split()) <= 7:
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
            # Try from longest prefix down; take the first that looks like a name
            words = line_text.split()
            for n in range(min(5, len(words)), 0, -1):
                candidate = " ".join(words[:n])
                if _looks_like_name(candidate):
                    score = (bold and underline) * 4 + bold * 2
                    candidates.append((score, line_max_size, candidate))
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

    # Join LinkedIn URLs that wrap across lines (e.g. "malthesh-\nkarnam-29a6b416a")
    _lnk_text = re.sub(
        r'(linkedin\.com/in/[a-zA-Z0-9\-_%]*)-[ \t]*\n[ \t]*([a-zA-Z0-9])',
        lambda m: m.group(1) + '-' + m.group(2),
        text, flags=re.I,
    )
    linkedin_match = re.search(
        r"(?:https?://)?(?:www\.)?linkedin\.com/in/([a-zA-Z0-9\-_%]+)", _lnk_text, re.I
    )
    if linkedin_match:
        parsed["linkedin"] = "https://www.linkedin.com/in/" + linkedin_match.group(1)

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
            r"|total experience|current employer|current designation|qualification|designation",
            line, re.I,
        )
    ]
    # Name fallback: look for a line that actually looks like a person's name.
    # Prevents professional summary sentences from being assigned as the name.
    # Prioritize all-caps names (likely to be actual names in resume headers)
    if not parsed["full_name"]:
        # First pass: look for all-caps names in FIRST 10 lines (e.g., "SIVARANJANI D")
        for _nl in useful_top_lines[:10]:
            if _nl.isupper() and _looks_like_name(_nl):
                parsed["full_name"] = _nl[:80]
                break
        # Second pass: look for TWO-WORD names (common name pattern) in first 15 lines
        if not parsed["full_name"]:
            for _nl in useful_top_lines[:15]:
                _words = _nl.split()
                if len(_words) == 2 and _looks_like_name(_nl):
                    parsed["full_name"] = _nl[:80]
                    break
        # Third pass: look for any proper-cased name in all useful lines
        if not parsed["full_name"]:
            for _nl in useful_top_lines:
                if _looks_like_name(_nl):
                    parsed["full_name"] = _nl[:80]
                    break
        if not parsed["full_name"]:
            # Looser pass: short proper-cased line with 2-5 words and no stopwords
            for _nl in useful_top_lines:
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
    if not parsed["title"]:
        _name_line = parsed.get("full_name", "")
        for _tl in useful_top_lines:
            if _tl == _name_line:
                continue
            if not canonical_section_name(_tl) and _looks_like_title(_tl):
                parsed["title"] = _tl[:80]
                break

    if not parsed["location"]:
        for line in lines[:20]:
            if re.search(
                r"\b(india|karnataka|bangalore|bengaluru|chennai|hyderabad|pune|mumbai|delhi|"
                r"noida|gurgaon|gurugram|coimbatore|trivandrum|kochi|jaipur|ahmedabad|kolkata|"
                r"ludhiana|chandigarh)\b",
                line, re.I,
            ):
                parsed["location"] = line[:180]
                break

    parsed.update(find_sections(lines))

    # ── Post-process: Extract skills from 2-column layouts ──────────────────────
    # In some PDFs (e.g., Sivaranjani's), skills are interleaved with summary due to
    # column extraction. If skills section is missing, extract from summary AND certifications.
    if not parsed.get("skills"):
        skills_items = []

        # Collect lines from multiple sections that might contain skills
        sections_to_scan = []
        if parsed.get("summary"):
            sections_to_scan.extend(parsed["summary"].split('\n'))
        if parsed.get("certifications"):
            sections_to_scan.extend(parsed["certifications"].split('\n'))

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

    return parsed


def merge_resume_data(form_data, parsed_data, overwrite=False):
    merged = dict(form_data)
    for key, value in parsed_data.items():
        if value and (overwrite or not merged.get(key)):
            merged[key] = value.strip()
    return merged


# ── Ollama "Resume Intelligence" parser (text LLM) ─────────────────────────────

_OLLAMA_BASE         = os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434")
_TEXT_MODEL          = os.environ.get("OLLAMA_TEXT_MODEL", "llama3.2:1b")
_OLLAMA_TEXT_TIMEOUT = int(os.environ.get("OLLAMA_TEXT_TIMEOUT", "60"))


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


def _ollama_chat(prompt, *, as_json, num_predict):
    """Single Ollama call. Returns the raw assistant string (or {} dict if as_json)."""
    body = {
        "model": _TEXT_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "options": {"temperature": 0, "num_predict": num_predict, "num_ctx": 4096},
    }
    if as_json:
        body["format"] = "json"
    req = urllib.request.Request(
        f"{_OLLAMA_BASE}/api/chat",
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=_OLLAMA_TEXT_TIMEOUT) as r:
        content = json.loads(r.read()).get("message", {}).get("content", "").strip()
    if not as_json:
        return content
    try:
        return json.loads(content)
    except (ValueError, TypeError):
        return {}


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
        if field and len(s.split()) <= 3:
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
    raw_text = extract_resume_text(path, "pdf")
    result = {k: "" for k in _BLANK_FIELDS}

    # ── 1. Name / title / location via Ollama ────────────────────────────────────
    # These are the fields a model genuinely beats rules at. Explicit field=description
    # anchoring stops a 1B model echoing an empty template; 3 fields + a short header
    # window is the sweet spot for reliability (more fields make it blank-out).
    ident = _ollama_chat(
        "Extract from this resume header. full_name=the person name, "
        "title=their professional job role (not a company), location=their city.\n"
        'Return ONLY JSON: {"full_name":"","title":"","location":""}\n\n'
        "RESUME:\n" + raw_text[:700],
        as_json=True, num_predict=120,
    )
    for k in ("full_name", "title", "location"):
        v = ident.get(k, "")
        if isinstance(v, list):
            v = " ".join(str(x) for x in v)
        result[k] = str(v).strip() if v else ""

    # Map the Ollama-extracted title to the closest predefined role group. (A 1B model
    # can't reliably pick from the 33-role list — tested — but the title it extracts is
    # accurate, and keyword-mapping that title to a group is deterministic and correct.)
    if result["title"]:
        result["title"] = _map_to_group_role(result["title"])

    # ── 2. Contact fields — deterministic, 100% reliable (not the heuristic engine) ─
    result["email"] = _extract_email(raw_text)
    result["phone"] = _extract_phone(raw_text)

    # Extract LinkedIn URL, handling wrapped URLs (newlines in the middle)
    # First, normalize whitespace in URLs (convert newlines to nothing within URLs)
    normalized_text = re.sub(
        r'(linkedin\.com/in/[a-z0-9\-_%]*)\s*\n\s*([a-z0-9])',
        r'\1\2',
        raw_text,
        flags=re.I
    )
    m = re.search(
        r"(?:https?://)?(?:www\.)?linkedin\.com/in/([A-Za-z0-9\-_%/?=&#]+)",
        normalized_text,
        re.I
    )
    if m:
        url_part = m.group(1)
        # Remove trailing query params that might have been captured incorrectly
        url_part = re.sub(r'[?&#].*$', lambda x: x.group(0) if 'skipRedirect' in x.group(0) else '', url_part)
        result["linkedin"] = "https://www.linkedin.com/in/" + url_part

    # ── 3. Content sections — use robust regex-based section finder for better handling ──
    # The simple line-by-line chunker (_split_resume_chunks) fails on 2-column PDFs where
    # sections can be interleaved. The find_sections() function has more robust logic for
    # detecting section boundaries using the full SECTION_ALIASES and handles edge cases.
    lines = raw_text.split('\n')
    chunks = find_sections(lines)

    # ── 4. Post-process: 2-column PDFs interleave skills with experience bullets
    # Use Ollama to extract just the skill list from the messy section.
    if chunks.get("skills"):
        skills_text = chunks["skills"]

        # Use Ollama to extract the actual skill list (it can pick them out of the mess)
        try:
            skill_list_json = _ollama_chat(
                "Extract ONLY the technical skill names from this messy resume section. "
                "Return a JSON list with the extracted skills. Ignore numbered bullets (experience items). "
                'Return ONLY JSON: {"skills":["skill1","skill2",...]}.\n\n'
                "SECTION:\n" + skills_text,
                as_json=True, num_predict=200,
            )
            logger.info(f"Ollama skill extraction result: {skill_list_json}")

            if skill_list_json and skill_list_json.get("skills"):
                chunks["skills"] = '\n'.join(skill_list_json["skills"])
                logger.info(f"Updated skills section from {len(skills_text)} chars to {len(chunks['skills'])} chars")
            else:
                logger.warning(f"Ollama skill extraction returned empty or invalid JSON: {skill_list_json}")
        except Exception as e:
            logger.warning(f"Ollama skill extraction failed: {e}. Keeping original skills section.")

        # Extract numbered bullets (1., 2., etc.) as experience if not already there
        experience_bullets = []
        for line in skills_text.split('\n'):
            line_stripped = line.strip()
            if re.match(r'^\d{1,2}\.\s+', line_stripped):
                experience_bullets.append(line_stripped)

        if experience_bullets:
            experience_text = '\n'.join(experience_bullets)
            if chunks.get("experience"):
                chunks["experience"] = chunks["experience"] + '\n' + experience_text
            else:
                chunks["experience"] = experience_text

    # ── 5. Clean up experience section — remove LinkedIn URLs and contact info
    if chunks.get("experience"):
        exp_text = chunks["experience"]
        # Remove complete LinkedIn URLs
        exp_text = re.sub(
            r'https?://(?:www\.)?linkedin\.com/in/[^\s\n]*(?:\s*\n\s*[a-zA-Z0-9\-_%/@?=&#]*)?',
            '',
            exp_text,
            flags=re.I
        )
        # Remove orphaned LinkedIn URL fragments (e.g., "karnam-29a6b416a/?skipRedirect=true")
        # These are fragments from wrapped URLs in PDFs
        exp_text = re.sub(
            r'^[a-z]+\-[a-z0-9]+/?[^\n]*(?:skipRedirect|linkedin)[^\n]*\n?',
            '',
            exp_text,
            flags=re.MULTILINE | re.I
        )
        # Remove any remaining orphaned protocol prefixes
        exp_text = re.sub(r'^\s*https?://\s*\n', '', exp_text, flags=re.MULTILINE)
        chunks["experience"] = exp_text.strip()

    # ── 6. Clean up education section — use Ollama to extract just education details
    # 2-column PDFs have personal info on left, education on right. Interleaved extraction is messy.
    # Use Ollama to extract just the degrees and institutions from the mixed content.
    if chunks.get("education"):
        edu_text = chunks["education"]

        # Use Ollama to extract just education information
        try:
            edu_json = _ollama_chat(
                "Extract ONLY the educational qualifications (degrees and institutions) from this mixed text. "
                "Ignore personal info fields like Email, Mobile, Phone, Total work experience, Social Link, etc. "
                'Return ONLY JSON: {"education":"degree1 from institute1, degree2 from institute2, ..."}.\n\n'
                "TEXT:\n" + edu_text,
                as_json=True, num_predict=150,
            )

            if edu_json and edu_json.get("education"):
                chunks["education"] = edu_json["education"]
                logger.info(f"Updated education section via Ollama")
            else:
                # Fallback to regex cleanup if Ollama extraction fails
                logger.warning("Ollama education extraction returned empty, using regex cleanup")
                edu_text = re.sub(r'Email\s+', '', edu_text, flags=re.I)
                edu_text = re.sub(r'Mobile\s+', '', edu_text, flags=re.I)
                edu_text = re.sub(r'\+?\(?\d{1,4}\)?[\s\-.]?\d{3,}[\s\-.]?\d{3,}[\s\-.]?\d{0,4}', '', edu_text)
                edu_text = re.sub(r'^(?:Total work experience|Social Link|City|Country|Languages?|Hobbies?)[^\n]*\n?', '', edu_text, flags=re.MULTILINE | re.I)
                edu_text = re.sub(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '', edu_text)
                edu_text = re.sub(r'\n\s*\n+', '\n', edu_text)
                chunks["education"] = edu_text.strip()
        except Exception as e:
            logger.warning(f"Education cleanup via Ollama failed: {e}")

    for field in ("summary", "skills", "experience", "education", "certifications", "projects"):
        if chunks.get(field):
            result[field] = chunks[field]

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
    with db_conn() as conn:
        if request.method == "POST":
            form_data = {k: request.form.get(k, "").strip() for k in [
                "full_name", "title", "email", "phone", "linkedin", "location", "summary",
                "skills", "experience", "education", "certifications", "projects",
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

                # parse_mode: "llm" = Resume Intelligence (AI), anything else = Quick Parse
                parse_mode = request.form.get("parse_mode", "0")
                try:
                    if ext == "pdf" and parse_mode == "llm":
                        try:
                            parsed_data, _ = parse_resume_with_llm_text(path)
                        except Exception as e:
                            logger.warning("Resume Intelligence failed (%s); using Quick Parse.", e)
                            parsed_data = _parse_pdf_quick(path)
                    elif ext == "pdf":
                        parsed_data = _parse_pdf_quick(path)
                    else:
                        extracted = extract_resume_text(path, ext)
                        parsed_data = parse_resume_text(extracted)
                    flash("Resume uploaded and data extracted successfully.", "success")
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
                conn.execute(
                    """
                    UPDATE resume SET
                        full_name=%(full_name)s, title=%(title)s, email=%(email)s, phone=%(phone)s,
                        linkedin=%(linkedin)s, location=%(location)s, summary=%(summary)s, skills=%(skills)s,
                        experience=%(experience)s, education=%(education)s, certifications=%(certifications)s,
                        projects=%(projects)s, slug=%(slug)s, resume_file=%(resume_file)s,
                        updated_at=NOW()
                    WHERE id=%(id)s
                    """,
                    merged,
                )
                sync_skills(conn, resume_id, merged.get("skills", ""))
                return redirect(url_for("profile_detail", resume_id=resume_id))

            merged["slug"] = unique_slug(conn, merged.get("full_name") or "profile")
            cursor = conn.execute(
                """
                INSERT INTO resume (
                    full_name, title, email, phone, linkedin, location, summary, skills,
                    experience, education, certifications, projects, slug, resume_file,
                    created_at, updated_at
                ) VALUES (
                    %(full_name)s, %(title)s, %(email)s, %(phone)s, %(linkedin)s, %(location)s, %(summary)s, %(skills)s,
                    %(experience)s, %(education)s, %(certifications)s, %(projects)s, %(slug)s, %(resume_file)s,
                    NOW(), NOW()
                )
                RETURNING id
                """,
                merged,
            )
            new_id = cursor.fetchone()["id"]
            sync_skills(conn, new_id, merged.get("skills", ""))
            flash("Profile saved successfully! Here are your top matching roles.", "success")
            return redirect(url_for("profile_detail", resume_id=new_id))

        # GET
        if resume_id:
            resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
            if not resume:
                return "Resume not found", 404
            resume = dict(resume)
        else:
            resume = {k: "" for k in [
                "full_name", "title", "email", "phone", "linkedin", "location",
                "summary", "skills", "experience", "education", "certifications",
                "projects", "slug", "resume_file",
            ]}
            resume["id"] = None

    return render_template("edit.html", resume=resume)


@app.route("/profile")
def view_current_profile():
    return redirect(url_for("profile_list"))


@app.route("/profiles")
def profile_list():
    with db_conn() as conn:
        resumes = conn.execute(
            """
            SELECT id, full_name, title, email, phone, location, resume_file, created_at, updated_at
            FROM resume
            ORDER BY updated_at DESC, id DESC
            """
        ).fetchall()
    return render_template("profile_list.html", resumes=resumes)


@app.route("/profile/<int:resume_id>")
def profile_detail(resume_id):
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not resume:
            return "Profile not found", 404

        # Calculate top 3 matching JDs
        jds = conn.execute("SELECT * FROM job_description ORDER BY created_at DESC").fetchall()
        resume_dict = dict(resume)

        matches = []
        for jd in jds:
            jd_dict = dict(jd)
            score = calculate_match_score(resume_dict, jd_dict)
            matches.append({
                'jd_id': jd['id'],
                'jd_title': jd['title'],
                'match_percentage': score['match_percentage'],
                'matched_count': score['matched_count'],
                'total_jd_requirements': score['total_jd_requirements']
            })

        matches.sort(key=lambda x: x['match_percentage'], reverse=True)
        top_matches = matches[:3]
        is_weak = top_matches and top_matches[0]['match_percentage'] < 50

    return render_template("profile.html", resume=resume, top_matches=top_matches, is_weak=is_weak)


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
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()

    if not resume:
        return "Profile not found", 404

    resume_dict = dict(resume)
    pdf_buffer = generate_resume_pdf(resume_dict)

    filename = f"{resume_dict.get('full_name', 'Resume').replace(' ', '_')}.pdf"
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
    with db_conn() as conn:
        row = conn.execute("SELECT resume_file FROM resume WHERE id = %s", (resume_id,)).fetchone()
        if not row:
            return "Profile not found", 404
        if row["resume_file"]:
            (UPLOAD_FOLDER / row["resume_file"]).unlink(missing_ok=True)
        conn.execute("DELETE FROM resume WHERE id = %s", (resume_id,))
    flash("Profile deleted.", "success")
    return redirect(url_for("profile_list"))


@app.route("/uploads/<path:filename>")
def uploaded_file(filename):
    # Only serve allowed file types
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ALLOWED_EXTENSIONS:
        abort(403)
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


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
        temp_path = UPLOAD_FOLDER / f"tmp-parse.{ext}"
        uploaded.save(str(temp_path))
        # parse_mode: "llm" = Resume Intelligence (AI), anything else = Quick Parse
        parse_mode  = request.form.get("parse_mode", "0")
        parser_used = "standard"
        try:
            if ext == "pdf" and parse_mode == "llm":
                try:
                    parsed, parser_used = parse_resume_with_llm_text(temp_path)
                except Exception as e:
                    logger.warning("Resume Intelligence failed (%s); using Quick Parse.", e)
                    parsed = _parse_pdf_quick(temp_path)
                    parser_used = "text (fallback)"
            elif ext == "pdf":
                parsed = _parse_pdf_quick(temp_path)
                parser_used = "text"
            else:
                text = extract_resume_text(temp_path, ext)
                parsed = parse_resume_text(text)
                parser_used = "text"
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

VALIDATION_ROLES = [
    "CSV Analyst", "CSV Lead", "Validation Engineer", "Validation Lead",
    "CQV Engineer", "CQV Lead", "Automation Engineer", "Automation Lead",
    "Tosca Engineer", "Tosca Lead", "Test Engineer", "Test Lead",
    "Design Control Consultant",
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
    from datetime import datetime as _dt
    with db_conn() as conn:
        total_resumes = conn.execute("SELECT COUNT(*) AS count FROM resume").fetchone()["count"]

        today_resumes = conn.execute(
            "SELECT id, full_name, title, created_at FROM resume"
            " WHERE created_at::date = CURRENT_DATE ORDER BY created_at DESC"
        ).fetchall()

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

        total_roles = conn.execute(
            "SELECT COUNT(DISTINCT NULLIF(title, '')) AS count FROM resume"
        ).fetchone()["count"] or 0

    return render_template(
        "dashboard.html",
        total_resumes=total_resumes,
        today_resumes=list(today_resumes),
        role_counts=list(role_counts),
        recent_activity=list(recent_activity),
        total_roles=total_roles,
        all_jd_roles=ALL_JD_ROLES,
    )


# ── New: Upload Files Route ───────────────────────────────────────────────────

@app.route("/upload-files", methods=["GET", "POST"])
def upload_files():
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
                            if not member_name or member_name.startswith("."):
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

    return render_template("upload_files.html", raw_files=raw_files, jds=list(jds))


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
                             "error": "File not found", "match_percentage": 0,
                             "match_level": "Error", "match_level_color": "gray"})
            continue
        ext = safe.rsplit(".", 1)[-1].lower() if "." in safe else ""
        if ext not in ALLOWED_EXTENSIONS:
            results.append({"file": filename, "candidate_name": safe,
                             "error": "Unsupported format", "match_percentage": 0,
                             "match_level": "Error", "match_level_color": "gray"})
            continue
        try:
            if ext == "pdf":
                parsed = _parse_pdf_quick(path)
            else:
                text = extract_resume_text(path, ext)
                parsed = parse_resume_text(text)
            score = calculate_match_score(parsed, jd_dict)
            results.append({
                "file": filename,
                "candidate_name": parsed.get("full_name") or Path(filename).stem,
                "title": parsed.get("title", ""),
                "email": parsed.get("email", ""),
                "match_percentage": score["match_percentage"],
                "match_level": score["match_level"],
                "match_level_color": score["match_level_color"],
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
                "match_percentage": 0,
                "match_level": "Error",
                "match_level_color": "gray",
                "matched_count": 0,
                "missing_count": 0,
                "total_jd_requirements": 0,
                "matched_skills": [],
                "missing_skills": [],
            })

    results.sort(key=lambda x: x.get("match_percentage", 0), reverse=True)
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
            created_at       TIMESTAMP DEFAULT NOW(),
            updated_at       TIMESTAMP DEFAULT NOW()
        )
    """)


PREDEFINED_JDS = [
    # ── From IFU JD PDF ──────────────────────────────────────────────────────
    {
        "title": "IFU Technical Writer",
        "role": "Validation Engineer",
        "category": "Validation Roles",
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
        "role": "Validation Lead",
        "category": "Validation Roles",
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
        "category": "IT Roles",
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
        "category": "IT Roles",
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
        "category": "IT Roles",
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
        "category": "IT Roles",
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
        "category": "IT Roles",
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
        "category": "IT Roles",
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
        "category": "IT Roles",
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
        "category": "Validation Roles",
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
]


def seed_jds(conn):
    existing_titles = {
        row["title"]
        for row in conn.execute("SELECT title FROM job_description").fetchall()
    }
    for jd in PREDEFINED_JDS:
        if jd["title"] not in existing_titles:
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

def _normalize(text):
    return re.sub(r'\s+', ' ', (text or '').lower()).strip()


def _parse_jd_items(skills_text, requirements_text, keywords_text):
    combined = "\n".join(filter(None, [
        str(skills_text or ''), str(requirements_text or ''), str(keywords_text or '')
    ]))
    seen, items = set(), []
    for line in combined.splitlines():
        line = re.sub(r'^[•\-–*►◆▸▪\d\.\)\s]+', '', line).strip()
        if not line:
            continue
        for part in re.split(r'[,;]', line):
            part = part.strip()
            key = _normalize(part)
            if 2 <= len(part) <= 80 and key not in seen:
                seen.add(key)
                items.append(part)
    return items


def _skill_matches(skill, corpus):
    skill_n = _normalize(skill)
    corpus_n = _normalize(corpus)
    if skill_n in corpus_n:
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


def calculate_match_score(resume_d, jd_d):
    corpus = " ".join(str(resume_d.get(k) or '') for k in
                      ['skills', 'experience', 'summary', 'projects', 'certifications', 'education', 'title'])
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


# ── JD Routes ─────────────────────────────────────────────────────────────────

@app.route("/jd-management")
def jd_management():
    with db_conn() as conn:
        ensure_jd_table(conn)
        seed_jds(conn)
        jds = conn.execute(
            "SELECT id, title, role, category, created_at"
            " FROM job_description ORDER BY category, title"
        ).fetchall()
    return render_template("jd_management.html", jds=list(jds))


@app.route("/jd/add", methods=["GET", "POST"])
def jd_add():
    if request.method == "POST":
        import time as _t
        data = {k: request.form.get(k, "").strip() for k in
                ["title", "role", "category", "responsibilities", "requirements", "skills", "keywords"]}
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
        with db_conn() as conn:
            ensure_jd_table(conn)
            conn.execute(
                """
                INSERT INTO job_description
                    (title, role, category, responsibilities, requirements, skills, keywords, jd_file)
                VALUES
                    (%(title)s, %(role)s, %(category)s, %(responsibilities)s,
                     %(requirements)s, %(skills)s, %(keywords)s, %(jd_file)s)
                """,
                data,
            )
        flash(f"Job Description '{data['title']}' added.", "success")
        return redirect(url_for("jd_management"))
    return render_template("jd_form.html", jd=None, all_roles=ALL_JD_ROLES)


@app.route("/jd/<int:jd_id>")
def jd_detail(jd_id):
    with db_conn() as conn:
        ensure_jd_table(conn)
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if not jd:
            return "Job Description not found", 404
        resumes = conn.execute(
            "SELECT id, full_name, title FROM resume ORDER BY updated_at DESC LIMIT 30"
        ).fetchall()
    return render_template("jd_detail.html", jd=dict(jd), resumes=list(resumes))


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
    if jd_dict.get('category') or jd_dict.get('role'):
        meta_parts = []
        if jd_dict.get('category'):
            meta_parts.append(f"Category: {jd_dict['category']}")
        if jd_dict.get('role'):
            meta_parts.append(f"Role: {jd_dict['role']}")
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
    with db_conn() as conn:
        ensure_jd_table(conn)
        if request.method == "POST":
            data = {k: request.form.get(k, "").strip() for k in
                    ["title", "role", "category", "responsibilities", "requirements", "skills", "keywords"]}
            data["id"] = jd_id
            conn.execute(
                """
                UPDATE job_description SET
                    title=%(title)s, role=%(role)s, category=%(category)s,
                    responsibilities=%(responsibilities)s, requirements=%(requirements)s,
                    skills=%(skills)s, keywords=%(keywords)s, updated_at=NOW()
                WHERE id=%(id)s
                """,
                data,
            )
            flash("Job Description updated.", "success")
            return redirect(url_for("jd_detail", jd_id=jd_id))
        jd = conn.execute(
            "SELECT * FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if not jd:
            return "Job Description not found", 404
    return render_template("jd_form.html", jd=dict(jd), all_roles=ALL_JD_ROLES)


@app.route("/jd/<int:jd_id>/delete", methods=["POST"])
def jd_delete(jd_id):
    with db_conn() as conn:
        ensure_jd_table(conn)
        row = conn.execute(
            "SELECT jd_file FROM job_description WHERE id = %s", (jd_id,)
        ).fetchone()
        if row and row["jd_file"]:
            (JD_UPLOAD_FOLDER / (row["jd_file"] or "")).unlink(missing_ok=True)
        conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))
    flash("Job Description deleted.", "success")
    return redirect(url_for("jd_management"))


@app.route("/uploads/jd/<path:filename>")
def jd_uploaded_file(filename):
    safe = secure_filename(filename)
    ext = safe.rsplit(".", 1)[-1].lower() if "." in safe else ""
    if ext not in ALLOWED_EXTENSIONS:
        abort(403)
    return send_from_directory(str(JD_UPLOAD_FOLDER), safe)


@app.route("/compare/<int:resume_id>")
def compare_select_jd(resume_id):
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
    return render_template(
        "compare_result.html",
        resume=dict(resume),
        jd=dict(jd),
        result=result,
        all_jds=list(all_jds),
    )


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False, port=5001)
