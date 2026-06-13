from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, abort
import psycopg2
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
from psycopg2.extras import RealDictCursor
from pathlib import Path
import re
import logging
import os
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
        "summary", "professional summary", "profile", "objective", "career objective",
        "about", "about me", "overview", "executive summary", "career profile",
        "professional profile", "personal profile", "career summary",
        "professional overview", "career overview",
    ],
    "skills": [
        "skills", "technical skills", "core competencies", "key skills", "skill set",
        "competencies", "technologies", "tools and technologies", "technical expertise",
        "areas of expertise", "technical competencies", "key competencies",
        "skills and competencies", "skills and technologies", "tools technologies",
    ],
    "experience": [
        "experience", "work experience", "professional experience", "employment history",
        "career history", "work history", "employment", "professional background",
        "job experience", "relevant experience", "industry experience",
        "work experience details",
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


def _pymupdf_page_text(page):
    """
    Extract page text using word-level bounding boxes.

    Words at the same vertical level (±5 pt) form one row.  Within a row,
    a gap larger than 30 pt between the right edge of one word and the left
    edge of the next indicates separate columns.

    Two-column resume layout (section label left, content right):
      If the leftmost column group is a known section alias (e.g. "SUMMARY"),
      it is emitted as its own line and the right column as further lines.
      → "SUMMARY" then "Validation Engineer..." → section detected ✓

    Table rows (company | role | dates all at same y):
      If the leftmost group is NOT a section alias, all groups are joined.
      → "Vaisesika Consulting Pvt Ltd CSV Lead June 2021 Till Present" ✓
    """
    raw_words = page.get_text("words")   # (x0, y0, x1, y1, word, blk, ln, wn)
    if not raw_words:
        return ""

    raw_words = sorted(raw_words, key=lambda w: w[1])   # top-to-bottom by y0

    Y_TOL   = 5    # pt — words within 5 pt vertically share a line
    COL_GAP = 30   # pt — gap between right edge of prev word and left of next

    def emit_row(row):
        if not row:
            return []
        row.sort()   # by x0

        # Split into column groups wherever consecutive words have a large x-gap
        groups = [[row[0]]]
        for i in range(1, len(row)):
            x1_prev = row[i - 1][1]   # right edge of previous word
            x0_curr = row[i][0]       # left edge of current word
            if x0_curr - x1_prev > COL_GAP:
                groups.append([])
            groups[-1].append(row[i])

        if len(groups) == 1:
            return [" ".join(wd for _, _, wd in groups[0])]

        # Two-column check: is the leftmost group a section heading?
        first_text = " ".join(wd for _, _, wd in groups[0])
        if canonical_section_name(first_text):
            # Emit label as its own line; each right-column group as a content line
            out = [first_text]
            for g in groups[1:]:
                out.append(" ".join(wd for _, _, wd in g))
            return out
        else:
            # Table row or normal multi-column text: join everything left-to-right
            all_words = []
            for g in groups:
                all_words.extend(g)
            all_words.sort()
            return [" ".join(wd for _, _, wd in all_words)]

    lines = []
    cur_row = []   # list of (x0, x1, word)
    cur_y   = None

    for w in raw_words:
        x0, y0, x1, y1, word = w[0], w[1], w[2], w[3], w[4]
        if not word.strip():
            continue
        if cur_y is None or abs(y0 - cur_y) <= Y_TOL:
            cur_row.append((x0, x1, word))
            if cur_y is None:
                cur_y = y0
        else:
            lines.extend(emit_row(cur_row))
            cur_row = [(x0, x1, word)]
            cur_y = y0

    lines.extend(emit_row(cur_row))
    return "\n".join(lines)


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


def find_sections(lines):
    sections = {key: "" for key in SECTION_ALIASES}
    _SKIP_HEADINGS = {
        "personal", "personal information", "personal details",
        "references", "declaration", "hobbies", "interests", "languages",
        "websites portfolios and profiles", "websites and profiles",
        "websites portfolios", "and profiles",
        "additional information", "other information",
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
        for m in re.finditer(r"[\+]?\(?\d[\d\s\-\.\(\)]{7,20}\d", line):
            candidate = m.group().strip()
            digits = re.sub(r"\D", "", candidate)
            if 10 <= len(digits) <= 15:
                # OCR often reads '+' as '4' or '1' for country-code prefix
                # e.g. "491 8220 133 233" → "+91 8220 133 233"
                if re.match(r"^[41](\d{2}[\s\-])", candidate) and not candidate.startswith("+"):
                    candidate = "+" + candidate[1:]
                return candidate
    return ""


def parse_resume_text(text):
    lines = normalize_lines(text)
    parsed = {
        "full_name": "", "title": "", "email": "", "phone": "", "linkedin": "", "location": "",
        "summary": "", "skills": "", "experience": "", "education": "",
        "certifications": "", "projects": "",
    }
    if not lines:
        return parsed

    # ── Email extraction ──────────────────────────────────────────────────────
    # Strategy: scan every line for a token containing @; require a dot after @.
    parsed["email"] = _extract_email(text)

    # ── Phone extraction ──────────────────────────────────────────────────────
    # Strategy: scan every line; take first candidate whose digit-only count >= 10.
    parsed["phone"] = _extract_phone(text)

    linkedin_match = re.search(
        r"(?:https?://)?(?:www\.)?linkedin\.com/in/([a-zA-Z0-9\-_%]+)", text, re.I
    )
    if linkedin_match:
        parsed["linkedin"] = "https://www.linkedin.com/in/" + linkedin_match.group(1)

    parsed["full_name"] = parse_label_value(lines, ["name", "full name", "candidate name"])
    parsed["title"] = parse_label_value(lines, ["title", "designation", "role", "current role"])
    parsed["location"] = parse_label_value(lines, ["location", "address", "city"])

    _email_fragments = set()
    if parsed.get("email"):
        _em = parsed["email"]
        for _tld_len in range(1, 5):
            _frag = _em[-_tld_len:]
            if _frag.isalpha():
                _email_fragments.add(_frag.lower())

    ignored = ["resume", "curriculum vitae", "cv"]
    useful_top_lines = [
        line for line in lines[:15]
        if line.lower() not in ignored
        and line.lower().strip() not in _email_fragments
        and not (len(line.strip()) <= 4 and line.strip().isalpha() and line.strip().islower())
        and not re.search(
            r"@|www\.|linkedin|github|\d{6,}|\+\d{1,3}[\s\-]?\d{3,}|(?<!\w)\d{10,}"
            r"|date.{0,5}birth|nationality|dob\b|gender|marital|\d{1,2}[./]\d{1,2}[./]\d{2,4}",
            line, re.I,
        )
    ]
    if not parsed["full_name"] and useful_top_lines:
        parsed["full_name"] = useful_top_lines[0][:120]
    if not parsed["title"] and len(useful_top_lines) > 1 and not canonical_section_name(useful_top_lines[1]):
        parsed["title"] = useful_top_lines[1][:150]

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

    return parsed


def merge_resume_data(form_data, parsed_data, overwrite=False):
    merged = dict(form_data)
    for key, value in parsed_data.items():
        if value and (overwrite or not merged.get(key)):
            merged[key] = value.strip()
    return merged


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

                try:
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
            return redirect(url_for("profile_list"))

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
    return render_template("profile.html", resume=resume)


@app.route("/profile/slug/<slug>")
def public_profile(slug):
    with db_conn() as conn:
        resume = conn.execute("SELECT * FROM resume WHERE slug = %s", (slug,)).fetchone()
    if not resume:
        return "Profile not found", 404
    return render_template("profile.html", resume=resume)


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
        try:
            text = extract_resume_text(temp_path, ext)
            parsed = parse_resume_text(text)
        finally:
            temp_path.unlink(missing_ok=True)
        return jsonify({"success": True, "message": "Fields extracted from resume.", "data": parsed})
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


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
