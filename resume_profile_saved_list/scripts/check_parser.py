#!/usr/bin/env python
"""
Resume-parser regression guard.

Run this before committing any change to the parsing pipeline (app.py). It exists
because the parser has repeatedly broken on previously-working resumes whenever new
functionality was added — most recently a block that referenced an undefined
`_unmapped_chunks`, which crashed the ENTIRE AI parser and silently downgraded every
upload to a low-quality "Quick Parse". That bug was invisible because there were no
tests and the crash was swallowed by a try/except.

This guard has two layers:

  1. STATIC  — runs pyflakes over app.py and fails on any "undefined name". This is
               the exact class of the _unmapped_chunks bug: a new/renamed feature
               referencing a variable that doesn't exist. Caught without executing.

  2. RUNTIME — generates small sample resume PDFs (single-column and two-column) and
               runs the real parse_resume_with_llm_text() over them, asserting it does
               not crash and that core sections are populated. AI/LLM calls are stubbed
               out so the check is fast and works offline (no Postgres, no Ollama).

Exit code is non-zero on any failure, so a git pre-commit hook can block the commit.

Usage:
    python scripts/check_parser.py
"""
import os
import sys
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
APP = os.path.join(ROOT, "app.py")

FAILURES = []


def _fail(msg):
    FAILURES.append(msg)
    print("  FAIL:", msg)


# ─────────────────────────────────────────────────────────────────────────────
# Layer 1 — static: no undefined names anywhere in app.py
# ─────────────────────────────────────────────────────────────────────────────
def check_undefined_names():
    print("[1/2] static check: undefined names (pyflakes)")
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pyflakes", APP],
            capture_output=True, text=True,
        )
    except Exception as exc:  # pyflakes not installed
        print("  WARN: pyflakes unavailable (%s) — install with `pip install pyflakes`" % exc)
        print("        static undefined-name check SKIPPED (runtime check still runs).")
        return
    lines = (proc.stdout + proc.stderr).splitlines()
    undefined = [ln for ln in lines if "undefined name" in ln]
    if undefined:
        for ln in undefined:
            _fail("undefined name — " + ln.strip())
    else:
        print("  OK: no undefined names")


# ─────────────────────────────────────────────────────────────────────────────
# Layer 2 — runtime: the parser runs without crashing and finds core sections
# ─────────────────────────────────────────────────────────────────────────────
def _make_pdf(path, story_flowables):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate
    doc = SimpleDocTemplate(path, pagesize=A4,
                            topMargin=36, bottomMargin=36, leftMargin=36, rightMargin=36)
    doc.build(story_flowables)


def _single_column_pdf(path):
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, Spacer
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("JAMIE RIVERA", ss["Title"]),
        Paragraph("Senior Software Engineer", n),
        Spacer(1, 8),
        Paragraph("SUMMARY", h),
        Paragraph("Senior engineer with 9 years building web services in Python and Go.", n),
        Paragraph("SKILLS", h),
        Paragraph("Python, Go, PostgreSQL, Docker, Kubernetes, AWS, REST APIs", n),
        Paragraph("WORK EXPERIENCE", h),
        Paragraph("Lead Engineer, Acme Corp, 2020-01 - Current", n),
        Paragraph("Built and scaled the billing platform to 2M users.", n),
        Paragraph("Engineer, Globex, 2016-06 - 2019-12", n),
        Paragraph("Owned the search indexing pipeline.", n),
        Paragraph("EDUCATION", h),
        Paragraph("B.Tech Computer Science, State University, 2012 - 2016", n),
        Paragraph("PROJECTS", h),
        Paragraph("OpenTrace: distributed tracing library used by 30+ teams.", n),
    ]
    _make_pdf(path, story)


def _tricky_headings_pdf(path):
    """Single-column PDF (reliable extraction order) that reproduces the heading
    problems seen on real sidebar resumes: an 'ABOUT ME' summary, a 'CONTACT'
    block whose phone/address/email must NOT bleed into Skills, and the typo'd
    'PROJECTS COMPELTED' / 'CERTIFICATES COURSES' section headings."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, Spacer
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("SAM PATEL", ss["Title"]),
        Paragraph("Mechanical Engineer", n),
        Paragraph("ABOUT ME", h),
        Paragraph("Mechanical engineer focused on planning and process optimization.", n),
        Paragraph("EDUCATION", h),
        Paragraph("Crescent University, B.E Mechanical, 2016 - 2020", n),
        Paragraph("SKILLS", h),
        Paragraph("AutoCAD", n), Paragraph("CREO", n), Paragraph("MS Office", n),
        Paragraph("CONTACT", h),
        Paragraph("+91-9789475003", n),
        Paragraph("No:08, NH3, Silapathikaram Street, Maraimalai Nagar", n),
        Paragraph("contactperson@example.com", n),
        Paragraph("WORK EXPERIENCE", h),
        Paragraph("Best Enterprises, Mechanical Engineer, May 2021 - Jul 2024", n),
        Paragraph("Managed inventory and production planning end to end.", n),
        Paragraph("PROJECTS COMPELTED", h),
        Paragraph("Investigation of material hardness of composites.", n),
        Paragraph("CERTIFICATES COURSES", h),
        Paragraph("Lean Six Sigma Greenbelt (manufacturing).", n),
    ]
    _make_pdf(path, story)


def _assert_tricky(parsed):
    """Custom checks for the tricky-headings fixture. Returns a list of problems."""
    problems = []
    skills = str(parsed.get("skills") or "")
    # CONTACT must act as a boundary — its phone/address/email must not leak into Skills.
    for leak in ("9789475003", "Silapathikaram", "@example.com"):
        if leak in skills:
            problems.append("CONTACT content leaked into skills: %r" % leak)
    # Typo/multi-word section headings must be recognized, not swallowed.
    if "material hardness" not in str(parsed.get("projects") or "").lower():
        problems.append("'PROJECTS COMPELTED' not recognized — projects missing the project text")
    if "six sigma" not in str(parsed.get("certifications") or "").lower():
        problems.append("'CERTIFICATES COURSES' not recognized — certifications missing the course")
    return problems


def _stranded_experience_pdf(path):
    """Single-column PDF where the last Work-Experience bullets are stranded AFTER
    an unmapped sidebar heading (Hobbies), mimicking a 2-column resume whose page-2
    sidebar is read before the main column's continuation. Bullets 4-5 must still
    be rescued into experience (this is the '_unmapped_chunks' rescue that was
    once broken by an undefined variable)."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("ALEX MORGAN", ss["Title"]),
        Paragraph("Design Engineer", n),
        Paragraph("WORK EXPERIENCE", h),
        Paragraph("HCL Technologies, Design Engineer, Jun 2019 - Present", n),
        Paragraph("1. Analyzed medical device classifications against IFUs.", n),
        Paragraph("2. Conducted gap assessments for EU MDR compliance.", n),
        Paragraph("3. Performed requirements traceability remediation.", n),
        Paragraph("HOBBIES", h),
        Paragraph("Cricket, Singing", n),
        Paragraph("4. Performed DHF gap assessments and remediation strategies.", n),
        Paragraph("5. Worked on DHF binder creation in the PLM system.", n),
        Paragraph("EDUCATION", h),
        Paragraph("M.Tech, Ballari Institute of Technology, 2017", n),
    ]
    _make_pdf(path, story)


def _assert_stranded(parsed):
    """The stranded continuation bullets (4 and 5) must end up in experience."""
    problems = []
    exp = str(parsed.get("experience") or "").lower()
    if "dhf gap assessments" not in exp:
        problems.append("bullet #4 (stranded past 'HOBBIES') was dropped from experience")
    if "dhf binder" not in exp:
        problems.append("bullet #5 (stranded past 'HOBBIES') was dropped from experience")
    return problems


def _two_column_pdf(path):
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle
    ss = getSampleStyleSheet()
    h, n = ss["Heading3"], ss["BodyText"]
    left = [
        Paragraph("SKILLS", h),
        Paragraph("AutoCAD", n), Paragraph("CREO", n), Paragraph("MS Office", n),
        Paragraph("EDUCATION", h),
        Paragraph("State University", n), Paragraph("B.E Mechanical", n), Paragraph("2016 - 2020", n),
        Paragraph("CONTACT", h),
        Paragraph("test@example.com", n), Paragraph("+91-9000000000", n),
    ]
    right = [
        Paragraph("SAM TAYLOR", ss["Title"]),
        Paragraph("Mechanical Engineer", n),
        Paragraph("WORK EXPERIENCE", h),
        Paragraph("Best Enterprises, Mechanical Engineer, May 21 - Jul 24", n),
        Paragraph("Managed inventory and production planning.", n),
        Paragraph("Reduced material wastage by 23% through better scheduling.", n),
        Paragraph("PROJECTS", h),
        Paragraph("Investigation of material hardness of composites.", n),
    ]
    tbl = Table([[left, right]], colWidths=[170, 340])
    tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    _make_pdf(path, [tbl])


def _summary_alias_pdf(path):
    """Resume using 'EXECUTIVE SUMMARY' (not the far more common 'SUMMARY') plus a
    separate later 'OBJECTIVE' heading. Exercises two things end-to-end through the
    real AI-path pipeline: (1) 'EXECUTIVE SUMMARY' must be recognized as a summary
    heading by the verbatim/_EXACT_HEADINGS table, not just the base SECTION_ALIASES
    table; (2) its content must not be silently overwritten when the later OBJECTIVE
    heading (same canonical section) is reached — both must be present."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("JORDAN LEE", ss["Title"]),
        Paragraph("Quality Engineer", n),
        Paragraph("EXECUTIVE SUMMARY", h),
        Paragraph("Quality engineer with 9 years in medical device design control.", n),
        Paragraph("SKILLS", h),
        Paragraph("ISO 13485", n), Paragraph("CAPA", n), Paragraph("Risk Management", n),
        Paragraph("OBJECTIVE", h),
        Paragraph("Seeking a senior design control role in a growing organization.", n),
        Paragraph("WORK EXPERIENCE", h),
        Paragraph("Acme Devices, Quality Engineer, Jan 2018 - Present", n),
        Paragraph("Led DHF remediation across the product portfolio.", n),
    ]
    _make_pdf(path, story)


def _assert_summary_alias(parsed):
    problems = []
    summary = str(parsed.get("summary") or "")
    if "9 years in medical device" not in summary:
        problems.append("'EXECUTIVE SUMMARY' content missing — heading not recognized")
    if "senior design control role" not in summary:
        problems.append("later 'OBJECTIVE' content was not merged (overwrote EXECUTIVE SUMMARY instead)")
    return problems


def _edu_cgpa_tail_pdf(path):
    """Education is the LAST section, and its final line is a bare CGPA value
    with no '%' sign (e.g. "7.5)") — the "2a-pre-2" rescue (stray Skills content
    stranded past Education) must recognize this as still-education content, not
    an unmarked stray line to sweep into Skills. Real-world case: a resume whose
    last Education line reads "...SRM Institute of Science & Technology (CGPA:
    7.5)" split across two lines by the PDF layout, where "7.5)" alone doesn't
    obviously look like education without a percent sign or degree keyword."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("TAYLOR MORGAN", ss["Title"]),
        Paragraph("Senior Fund Accountant", n),
        Paragraph("TECHNICAL SKILLS", h),
        Paragraph("NAV Calculation & Validation", n),
        Paragraph("Capital Calls & Investor Distributions", n),
        Paragraph("SAP ERP & Tally", n),
        Paragraph("EDUCATION", h),
        Paragraph("Master of Business Administration - SRM Institute of Science & Technology", n),
        Paragraph("(CGPA: 9.0) Bachelor of Commerce (General) - SRM Institute (CGPA:", n),
        Paragraph("7.5)", n),
    ]
    _make_pdf(path, story)


def _assert_edu_cgpa_tail(parsed):
    problems = []
    skills = str(parsed.get("skills") or "")
    education = str(parsed.get("education") or "")
    if "7.5" in skills:
        problems.append("bare CGPA tail '7.5)' was wrongly swept from Education into Skills: %r" % skills)
    if "7.5" not in education:
        problems.append("bare CGPA tail '7.5)' was lost from Education entirely: %r" % education)
    return problems


def _roles_resp_table_pdf(path):
    """A 'Project Details | Roles and Responsibilities' table (common in Indian
    IT/validation CVs) linearizes as two adjacent heading-like lines with an
    empty span between them, then everything after — real-world case: Satheesh
    Kumar's resume. "Roles and Responsibilities" maps to "experience" by
    default (it's also a legitimate per-job Experience sub-heading elsewhere),
    but here it must route to "projects" since it immediately follows a
    "Project Details"-family heading — otherwise the entire project table gets
    swallowed into Experience and Projects stays empty."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("MORGAN LEE", ss["Title"]),
        Paragraph("Validation Analyst", n),
        Paragraph("Work Experience", h),
        Paragraph("Acme Labs, CSV Analyst, Jan 2021 - Present", n),
        Paragraph("Relevant Project/Organizational Details", h),
        Paragraph("Project/Organizational Details", h),
        Paragraph("Roles and Responsibilities", h),
        Paragraph("Client: Acme Pharma", n),
        Paragraph("Project 1: Batch Record System (Jan 2022 to Dec 2022)", n),
        Paragraph("Project Description: Electronic batch record system for manufacturing.", n),
        Paragraph("Validation tester and handled the project as single member.", n),
        Paragraph("Supported validation plan.", n),
    ]
    _make_pdf(path, story)


def _assert_roles_resp_table(parsed):
    problems = []
    experience = str(parsed.get("experience") or "")
    projects = str(parsed.get("projects") or "")
    if "Batch Record System" not in projects:
        problems.append("project table content was not routed to 'projects': %r" % projects)
    if "Validation tester and handled the project" not in projects:
        problems.append("'Roles and Responsibilities' content was not routed to 'projects': %r" % projects)
    if "Batch Record System" in experience or "Validation tester and handled" in experience:
        problems.append("project table content leaked into 'experience' instead of (or in addition to) 'projects': %r" % experience)
    if "Acme Labs" not in experience:
        problems.append("genuine Work Experience content was lost: %r" % experience)
    return problems


def _sidebar_personal_bleed_pdf(path):
    """A 2-column sidebar layout where 'WEBSITES, PORTFOLIOS AND PROFILES' is
    fused with its own URL on one line (so it can't be an exact heading match),
    and 'PERSONAL INFORMATION' wraps across two lines each fused with
    unrelated main-column content at the same Y-position ('PERSONAL' + a DOB
    line, 'INFORMATION' + 'Gender: Female') — real-world case: Sindhu
    Sundaramoorthy's resume. Two separate bugs let this leak into Skills:
    (1) a CGPA-decimal marker regex falsely matched the '30.11' inside the DOB
    date, shifting the Education 'last marker' index past the DOB line and
    misclassifying the bare 'INFORMATION Gender: Female' continuation as an
    unmarked stray line to rescue into Skills; (2) the Skills-bleed truncation
    only matched a whole-line heading, never the fused 'WEBSITES, PORTFOLIOS
    <url>' line."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        Paragraph("SINDHU SUNDARAMOORTHY", ss["Title"]),
        Paragraph("Quality Analyst", n),
        Paragraph("SKILLS", h),
        Paragraph("Requirements analysis, Validation testing, Project management", n),
        Paragraph("WEBSITES, PORTFOLIOS https://www.linkedin.com/in/sindhu-example/", n),
        Paragraph("EDUCATION", h),
        Paragraph("SRM University, B.Com (CGPA: 7.5)", n),
        Paragraph("PERSONAL Date of birth: 30.11.1995 Nationality: Indian", n),
        Paragraph("INFORMATION Gender: Female", n),
    ]
    _make_pdf(path, story)


def _assert_sidebar_personal_bleed(parsed):
    problems = []
    skills = str(parsed.get("skills") or "")
    education = str(parsed.get("education") or "")
    up_skills = skills.upper()
    for sign in ("WEBSITES", "GENDER", "NATIONALITY", "DATE OF BIRTH", "INFORMATION"):
        if sign in up_skills:
            problems.append("personal-info/website bleed leaked into Skills (%r found): %r" % (sign, skills))
    if "Requirements analysis" not in skills:
        problems.append("genuine Skills content was lost: %r" % skills)
    if "7.5" not in education and "CGPA" not in education.upper():
        problems.append("genuine Education content (CGPA) was lost: %r" % education)
    return problems


def _unlabeled_projects_pdf(path):
    """A bare 'Projects' heading (not the 'Project Details' table-header
    variant) listing several clients, where only the FIRST client spells out
    its own 'RESPONSIBILITIES' sub-heading and the rest are bare bullets with
    no heading at all — real-world case: Saravana Kumar Sathiamoorthy's resume.
    'RESPONSIBILITIES' defaults to 'experience' (it's also a legitimate per-job
    Experience sub-heading elsewhere), and previously only flipped to
    'projects' when immediately preceded by a 'Project Details'-family
    heading — a plain 'Projects' heading didn't count, so the sub-heading (and
    everything after it, up to the next real heading) got stolen into
    Experience while Projects was left truncated to just the first client's
    name/duration line. Also covers per-skill role-based summary headings
    (common in QA/testing resumes) that must fold into Experience as a
    trailing "Work summary" block instead of bleeding into Certifications, and
    Achievements merging into Certifications alongside real certificates."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    ss = getSampleStyleSheet()
    h, n = ss["Heading2"], ss["BodyText"]
    story = [
        # Name wraps across two lines in a narrow header column — both lines
        # score identically (same bold Title style), so the font-based name
        # extractor must merge them instead of only keeping the first.
        Paragraph("Jordan", ss["Title"]),
        Paragraph("Casey", ss["Title"]),
        Paragraph("Automation Test Lead", n),
        # Email + phone in the text layer so this stays on the fast PyMuPDF
        # font-based name path, matching the real resume — without them, missing
        # contact info would trigger the (slow) EasyOCR fallback path instead.
        Paragraph("jordan.casey@example.com (+1) 5551234567", n),
        Paragraph("Work Experience", h),
        Paragraph("Aug 2010 - Present", n),
        Paragraph("Automation Test Lead, Atos", n),
        Paragraph("Led testing efforts for 14 years, driving efficiency gains.", n),
        Paragraph("Projects", h),
        Paragraph("45 Months", n),
        Paragraph("Stellantis", n),
        Paragraph("RESPONSIBILITIES", n),
        Paragraph("Building QE Assist, a GenAI-powered testing tool.", n),
        Paragraph("58 Months", n),
        Paragraph("Fedex", n),
        Paragraph("Plan with Scrum teams on sprint planning.", n),
        Paragraph("Certification", h),
        Paragraph("ISTQB Foundation Level Tester", n),
        Paragraph("Certified Scrum Master", n),
        Paragraph("AUTOMATION TESTING SUMMARY", h),
        Paragraph("Design, develop and maintain automated test scripts.", n),
        Paragraph("MANUAL TESTING SUMMARY", h),
        Paragraph("Well versed with all stages of SDLC and STLC.", n),
        Paragraph("SCRUM MASTER SUMMARY", h),
        Paragraph("Scrum Master for multiple teams, present KPIs.", n),
        Paragraph("ACHIEVEMENTS", h),
        Paragraph("Received Individual Value Award for Q1.", n),
    ]
    _make_pdf(path, story)


def _assert_unlabeled_projects(parsed):
    problems = []
    full_name = str(parsed.get("full_name") or "")
    if full_name != "Jordan Casey":
        problems.append("two-line wrapped name was not merged into 'full_name': %r" % full_name)
    experience = str(parsed.get("experience") or "")
    projects = str(parsed.get("projects") or "")
    certifications = str(parsed.get("certifications") or "")
    if "Fedex" not in projects or "Plan with Scrum teams" not in projects:
        problems.append("unlabeled second client block leaked out of 'projects': %r" % projects)
    if "Stellantis" not in projects or "Building QE Assist" not in projects:
        problems.append("first client's RESPONSIBILITIES content was not routed to 'projects': %r" % projects)
    if "Fedex" in experience or "Building QE Assist" in experience:
        problems.append("project content leaked into 'experience': %r" % experience)
    if "Certification" in projects or "ISTQB" in projects:
        problems.append("a bare 'Certification' heading was not recognized as a boundary — "
                         "Certification content leaked into 'projects': %r" % projects)
    if "Atos" not in experience:
        problems.append("genuine Work Experience content was lost: %r" % experience)
    if "Work summary" not in experience or "Scrum Master for multiple teams" not in experience:
        problems.append("role-based summary headings were not folded into Experience as 'Work summary': %r" % experience)
    if "AUTOMATION TESTING SUMMARY" in certifications or "SDLC" in certifications:
        problems.append("role-based summary content leaked into 'certifications': %r" % certifications)
    if "ISTQB" not in certifications:
        problems.append("genuine Certification content was lost: %r" % certifications)
    if "Individual Value Award" not in certifications:
        problems.append("Achievements were not merged into 'certifications': %r" % certifications)
    return problems


def _stub_out_ai(app):
    """Neutralize every LLM call so the check is fast and offline. We test the
    deterministic pipeline + crash-safety, not the model."""
    def _noop_chat(prompt, *, as_json=False, num_predict=0, num_ctx=8192, timeout=None):
        return {} if as_json else ""
    app._ollama_chat = _noop_chat
    app._OLLAMA_BASE = "http://127.0.0.1:9"   # unreachable → any stray direct call fails fast
    app._OLLAMA_TEXT_TIMEOUT = 1


def check_parsing_runtime():
    print("[2/2] runtime check: parser runs without crashing and finds sections")
    # Import must not require a live DB — app.py's schema-ensure calls swallow errors.
    os.environ.setdefault("DATABASE_URL", "postgresql://invalid:invalid@127.0.0.1:1/none")
    sys.path.insert(0, ROOT)
    import logging
    logging.disable(logging.CRITICAL)
    try:
        import app
    except Exception as exc:
        _fail("could not import app.py: %r" % exc)
        return
    _stub_out_ai(app)

    tmp = tempfile.mkdtemp(prefix="parser_guard_")
    # Each case: (label, pdf_path, builder, required_non_empty_fields, custom_assert_or_None)
    cases = [
        ("single-column", os.path.join(tmp, "single.pdf"), _single_column_pdf,
         # single column is the well-supported case → assert real quality
         ["full_name", "experience", "skills", "education"], None),
        ("tricky-headings", os.path.join(tmp, "tricky.pdf"), _tricky_headings_pdf,
         # ABOUT ME / CONTACT-boundary / typo'd PROJECTS+CERTIFICATES headings
         ["experience", "skills", "projects", "certifications"], _assert_tricky),
        ("stranded-experience", os.path.join(tmp, "stranded.pdf"), _stranded_experience_pdf,
         # numbered Work-Experience bullets stranded past an unmapped sidebar heading
         ["experience"], _assert_stranded),
        ("two-column", os.path.join(tmp, "two.pdf"), _two_column_pdf,
         # two-column quality is a known limitation → assert only crash-safety + experience
         ["experience"], None),
        ("summary-alias", os.path.join(tmp, "summary_alias.pdf"), _summary_alias_pdf,
         # "EXECUTIVE SUMMARY" heading recognized + merged with a later "OBJECTIVE"
         ["summary", "skills", "experience"], _assert_summary_alias),
        ("edu-cgpa-tail", os.path.join(tmp, "edu_cgpa_tail.pdf"), _edu_cgpa_tail_pdf,
         # bare CGPA value ("7.5)") at the end of Education must stay in Education
         ["skills", "education"], _assert_edu_cgpa_tail),
        ("roles-resp-table", os.path.join(tmp, "roles_resp_table.pdf"), _roles_resp_table_pdf,
         # "Project Details | Roles and Responsibilities" table must route to
         # 'projects', not swallow itself into 'experience'
         ["experience", "projects"], _assert_roles_resp_table),
        ("sidebar-personal-bleed", os.path.join(tmp, "sidebar_personal_bleed.pdf"), _sidebar_personal_bleed_pdf,
         # fused "WEBSITES, PORTFOLIOS <url>" and wrapped "PERSONAL"/"INFORMATION"
         # sidebar lines must never leak into Skills
         ["skills", "education"], _assert_sidebar_personal_bleed),
        ("unlabeled-projects", os.path.join(tmp, "unlabeled_projects.pdf"), _unlabeled_projects_pdf,
         # a bare "Projects" heading with only its first client labeled
         # "RESPONSIBILITIES" must not leak later unlabeled clients into
         # Experience; a two-line wrapped name must be merged; role-based
         # summary headings fold into Experience as "Work summary";
         # Achievements merge into Certifications
         ["full_name", "experience", "projects", "certifications"], _assert_unlabeled_projects),
    ]
    for label, path, builder, required, custom in cases:
        try:
            builder(path)
        except Exception as exc:
            _fail("could not build %s fixture: %r" % (label, exc))
            continue
        try:
            result = app.parse_resume_with_llm_text(path)
        except Exception as exc:
            import traceback
            _fail("parse_resume_with_llm_text CRASHED on %s fixture: %r" % (label, exc))
            traceback.print_exc()
            continue
        parsed = result[0] if isinstance(result, tuple) else result
        if not isinstance(parsed, dict):
            parsed = getattr(parsed, "__dict__", {})
        missing = [f for f in required if not str(parsed.get(f) or "").strip()]
        problems = list(custom(parsed)) if custom else []
        if missing:
            problems.append("expected non-empty %s (got empty)" % missing)
        if problems:
            for p in problems:
                _fail("%s fixture: %s" % (label, p))
        else:
            print("  OK: %s fixture parsed and passed all checks" % label)

    # The stranded-experience rescue must also work through the Quick-Parse path
    # (parse_resume_text/find_sections) directly — this is the path Bulk Upload
    # always uses (no AI), so a rescue that only works via parse_resume_with_llm_text
    # silently fails to help the far more common bulk-uploaded multi-page resume.
    stranded_path = os.path.join(tmp, "stranded.pdf")
    try:
        raw_text = app.extract_resume_text(stranded_path, "pdf")
        quick_parsed = app.parse_resume_text(raw_text)
    except Exception as exc:
        import traceback
        _fail("parse_resume_text CRASHED on stranded-experience fixture: %r" % exc)
        traceback.print_exc()
    else:
        problems = list(_assert_stranded(quick_parsed))
        if problems:
            for p in problems:
                _fail("stranded-experience (Quick-Parse path): %s" % p)
        else:
            print("  OK: stranded-experience fixture also rescued via Quick-Parse path")


def check_skill_fragment_helper():
    """Lock the trailing skill-fragment cleanup: a wrapped one-word orphan
    ("Medical" from "Medical Device Classification") must be dropped, while a
    genuine standalone one-word skill ("Python") must be kept."""
    print("[3/3] logic check: trailing skill-fragment cleanup")
    os.environ.setdefault("DATABASE_URL", "postgresql://invalid:invalid@127.0.0.1:1/none")
    sys.path.insert(0, ROOT)
    import logging
    logging.disable(logging.CRITICAL)
    try:
        import app
    except Exception as exc:
        _fail("could not import app.py for skill-fragment check: %r" % exc)
        return
    fn = getattr(app, "_drop_trailing_skill_fragments", None)
    if fn is None:
        _fail("_drop_trailing_skill_fragments helper is missing")
        return
    got = fn(["Medical Device Classification", "Good Team Player", "Medical"])
    if "Medical" in got:
        _fail("trailing wrap-fragment 'Medical' was NOT dropped: %r" % got)
    kept = fn(["Java", "Selenium", "Python"])
    if kept != ["Java", "Selenium", "Python"]:
        _fail("standalone one-word skills were wrongly dropped: %r" % kept)

    # Skills bleed-truncation: a bled-in "Languages" section must be cut off.
    trunc = getattr(app, "_truncate_skills_at_bleed", None)
    if trunc is None:
        _fail("_truncate_skills_at_bleed helper is missing")
    else:
        got = trunc(["Java", "Python", "Languages", "Tamil", "English"])
        if got != ["Java", "Python"]:
            _fail("skills bleed-truncation failed to cut 'Languages' section: %r" % got)

    # Wrapped-URL stitching: a linkedin URL split across lines must rejoin, and an
    # ALL-CAPS sidebar heading ("WWW") on the next line must NOT be joined.
    stitch = getattr(app, "_stitch_wrapped_urls", None)
    if stitch is None:
        _fail("_stitch_wrapped_urls helper is missing")
    else:
        out = stitch("https://www.linkedin.com/i\nn/prasanth-vb-0017b016a\nWWW\nBold Profile")
        if "linkedin.com/in/prasanth-vb-0017b016a" not in out:
            _fail("wrapped LinkedIn URL was not stitched: %r" % out)
        if "0017b016aWWW" in out or "016aWWW" in out:
            _fail("stitching wrongly joined an ALL-CAPS heading onto the URL: %r" % out)

    # Interleaved contact/skills sidebar must be stripped from experience, its
    # skills recovered, and CLEAN experience must be left untouched.
    bleed = getattr(app, "_extract_sidebar_bleed_from_experience", None)
    if bleed is None:
        _fail("_extract_sidebar_bleed_from_experience helper is missing")
    else:
        sample = ("• Writing test cases with TestNG and\nContact\nPhone\n"
                  "+91-9444965882\nE-mail\nx@y.com\nLinkedIn\nSkills\n"
                  "Selenium WebDriver\nJAVA\n• Automating all the manual test cases")
        clean, sk = bleed(sample)
        if any(w in clean for w in ("Contact", "Phone", "+91", "@y.com", "LinkedIn")):
            _fail("sidebar contact block not removed from experience: %r" % clean)
        if "Selenium WebDriver" not in sk or "JAVA" not in sk:
            _fail("sidebar skills not recovered from experience: %r" % sk)
        clean_exp = "• Did X\n• Did Y\nManaged the team and delivered projects on schedule"
        clean2, sk2 = bleed(clean_exp)
        if clean2 != clean_exp or sk2:
            _fail("sidebar-bleed helper wrongly altered clean experience: %r / %r" % (clean2, sk2))

    # find_sections must not drop content whose lead-in words merely resemble a
    # section-heading alias (e.g. "Professional experience in Power BI..." starts
    # with "Professional Experience", a recognized heading synonym for the section
    # we're already inside) — only a BARE repeat of the heading word itself (no
    # real remainder) should be swallowed as a wrapped-heading artifact.
    fs = getattr(app, "find_sections", None)
    if fs is None:
        _fail("find_sections helper is missing")
    else:
        lines = [
            "WORK EXPERIENCE",
            "Acme Corp, Design Engineer, Jan 2020 - Present",
            "Professional experience in Power BI, Power Apps & Power Automate",
            "Skilled in managing projects to ensure timely completion",
        ]
        exp = fs(lines).get("experience", "")
        if "Power BI" not in exp:
            _fail("find_sections dropped real content after a heading-alias lead-in: %r" % exp)

        # Two DIFFERENT headings mapping to the same canonical section (e.g. a
        # resume with both "SUMMARY" and, further down, a separate "OBJECTIVE")
        # must MERGE, not have the second one silently overwrite the first. This
        # was a real bug: Suresh's actual profile-summary paragraph under "SUMMARY"
        # was being wiped out entirely by a later "OBJECTIVE" heading's content.
        summary_lines = [
            "SUMMARY",
            "Highly accomplished engineer with 7+ years of experience.",
            "SKILLS",
            "AutoCAD", "SolidWorks",
            "OBJECTIVE",
            "To secure a challenging position in product design.",
        ]
        merged = fs(summary_lines).get("summary", "")
        if "Highly accomplished engineer" not in merged:
            _fail("SUMMARY content was lost when a later OBJECTIVE heading was reached: %r" % merged)
        if "challenging position" not in merged:
            _fail("OBJECTIVE content was not merged into summary: %r" % merged)

        # New summary-alias headings (Executive Summary, Professional Profile,
        # Profile, Career Objective) must resolve to canonical "summary" and
        # extract their content — these were previously missing from the
        # verbatim/_EXACT_HEADINGS table used by the AI parsing path, even
        # though the base SECTION_ALIASES table already recognized them.
        canon = getattr(app, "canonical_section_name", None)
        if canon is not None:
            for heading, expect in (
                ("EXECUTIVE SUMMARY", "summary"), ("Professional Profile", "summary"),
                ("Profile", "summary"), ("Career Objective", "summary"),
                ("About Me", "summary"), ("SUMMARY", "summary"),
            ):
                if canon(heading) != expect:
                    _fail("canonical_section_name(%r) = %r, expected %r" % (heading, canon(heading), expect))

    # Wrapped-heading remnants ("ME" from "ABOUT ME", "COMPELTED" from "PROJECTS
    # COMPELTED") must be stripped from a section's first line; normal content kept.
    striptail = getattr(app, "_strip_heading_tail_prefix", None)
    if striptail is None:
        _fail("_strip_heading_tail_prefix helper is missing")
    else:
        if striptail("ME\nTo secure a challenging position\ngrowth") != "To secure a challenging position\ngrowth":
            _fail("did not strip 'ME' remnant from summary")
        if striptail("COMPELTED\nInvestigation of material hardness") != "Investigation of material hardness":
            _fail("did not strip 'COMPELTED' remnant from projects")
        keepnormal = striptail("Led the BREXIT project\nUpdated 728 drawings")
        if keepnormal != "Led the BREXIT project\nUpdated 728 drawings":
            _fail("wrongly stripped a normal first line: %r" % keepnormal)

    # Prose/category skills → atomic keywords; already-clean lists left untouched.
    kw = getattr(app, "_extract_skill_keywords", None)
    if kw is None:
        _fail("_extract_skill_keywords helper is missing")
    else:
        prose = ("CAD Design Tools\nProficient in PTC CREO, Solid Works, AutoCAD "
                 "(3D modeling), experienced with Catia-V5 and NX CAD/CAM\n"
                 "TrackWise and Salesforce\nProficient in TrackWise and Salesforce\n"
                 "Design and Quality Analysis\nQualified Defect Preventive Analyst (DPA), "
                 "Design Configuration Controller (CC), GD&T knowledge")
        got = kw(prose)
        items = [l.strip() for l in got.split('\n') if l.strip()]
        for need in ("PTC CREO", "Solid Works", "AutoCAD", "NX CAD/CAM", "TrackWise", "Salesforce"):
            if need not in items:
                _fail("keyword extraction dropped a real skill %r: %r" % (need, items))
        if "GD&T knowledge" not in items or "GD" in items:
            _fail("keyword extraction split GD&T incorrectly: %r" % items)
        if any(w in got for w in ("Proficient in", "experienced with")):
            _fail("keyword extraction left filler words in: %r" % got)
        # Already-clean atomic skills must be returned unchanged (gate is a no-op).
        clean = "Selenium WebDriver\nJAVA\nCypress\nPlaywright\nPuppeteer"
        if kw(clean) != clean:
            _fail("keyword extraction wrongly altered an already-clean skills list: %r" % kw(clean))
        # A resume that lists "Category – item1, item2, item3" per bullet (no filler
        # verbs — just naturally long lines enumerating many items) is ALREADY a
        # clean, structured skills list and must be left completely untouched, one
        # bullet per category, not shredded into individual tokens. Line length
        # alone must never be the trigger for atomization — only genuine prose
        # filler-verb evidence should be.
        structured = (
            "Design Control & Design History File (DHF)\n"
            "Verification & Validation (V&V) – Protocols, Reports, Aging studies, Functional & Packaging tests\n"
            "Regulatory Compliance – MDR, 21 CFR 820, ISO 13485, ISO 14971, ISO 10993, ISO 15223\n"
            "Risk Management – DFMEA, PFMEA, UFMEA, RMP"
        )
        if kw(structured) != structured:
            _fail("keyword extraction wrongly shredded a clean 'Category - item, item' skills list: %r" % kw(structured))

    # "Category – detail1, detail2" on a SINGLE line (em-dash separator, not the
    # two-line "Header\nProficient in ..." format) must split so the category
    # label doesn't glue onto the first detail into one over-long chunk that
    # then fails the length filter and drops the ENTIRE bullet, header included.
    split_fn = getattr(app, "_split_desc_to_skills", None)
    if split_fn is None:
        _fail("_split_desc_to_skills helper is missing")
    else:
        items = split_fn("Verification & Validation (V&V) – Protocols, Reports, Aging studies")
        for need in ("Verification", "Validation", "Protocols", "Reports", "Aging studies"):
            if need not in items:
                _fail("dash-separated category/detail line lost %r: %r" % (need, items))
        # A tight hyphenated compound (no surrounding spaces) must NOT be split.
        tight = split_fn("Stack-up analysis")
        if tight != ["Stack-up analysis"]:
            _fail("dash-splitting wrongly broke a hyphenated compound word: %r" % tight)

    # The candidate's own name/title bleeding into Skills from a column boundary
    # must be dropped, but short legitimate skill CODES ("8D", "5S") must survive —
    # this exact regression happened once: the filter was wired in AFTER keyword
    # extraction already atomized skills, so a real 2-char code looked like a
    # "bare gutter-bleed fragment" and got dropped along with the real leak.
    nametitle = getattr(app, "_drop_name_title_and_gutter_bleed", None)
    if nametitle is None:
        _fail("_drop_name_title_and_gutter_bleed helper is missing")
    else:
        bled = ["ISO 13485", "Solid Works", "Malthesh Karnam", "Design Engineer", "GD&T knowledge"]
        got = nametitle(bled, "Malthesh Karnam", "")
        if "Malthesh Karnam" in got:
            _fail("name/title filter did not drop the candidate's own name from skills: %r" % got)
        if "Design Engineer" in got:
            _fail("name/title filter did not drop a bare role-title leak from skills: %r" % got)
        for keep in ("ISO 13485", "Solid Works", "GD&T knowledge"):
            if keep not in got:
                _fail("name/title filter wrongly dropped a real skill %r: %r" % (keep, got))
        atomized = ["CAPA", "8D", "PPAP", "5S", "Six Sigma"]
        got2 = nametitle(atomized, "Malthesh Karnam", "")
        if got2 != atomized:
            _fail("name/title filter dropped a short legitimate skill code: %r -> %r" % (atomized, got2))

    # Unlabeled skills-sidebar continuation bleeding into Experience (a run of
    # 3+ "Category Header" + "description" pairs, no bullets/dates between) must
    # be pulled out into Skills, and clean bulleted experience left untouched.
    rescue = getattr(app, "_rescue_skill_sidebar_from_experience", None)
    if rescue is None:
        _fail("_rescue_skill_sidebar_from_experience helper is missing")
    else:
        header_ok = getattr(app, "_looks_like_skill_sidebar_header", None)
        if header_ok is not None and not header_ok("Microsoft Office Suite"):
            _fail("_looks_like_skill_sidebar_header wrongly rejected a real category header")
        bled_exp = (
            "- Provided task support and guided trainees\n"
            "- Experienced in TrackWise and Salesforce\n"
            "Microsoft Office Suite\n"
            "Advanced Excel skills: Macros, VBA User Forms, pivot tables, and charts\n"
            "Data Management\n"
            "Professional experience in Power BI, Power Apps & Power Automate\n"
            "TrackWise and Salesforce\n"
            "Proficient in TrackWise for IVD products and Salesforce with hands-on experience\n"
            "New Product Lead Design (NPD) | Aruna Alloys and Steels Pvt Ltd\n"
            "- Prepared 3D modeling (CREO, SolidWorks) and 2D drafting (AutoCAD)"
        )
        clean_exp, rescued = rescue(bled_exp)
        if "Microsoft Office Suite" not in rescued or "Data Management" not in rescued:
            _fail("skill-sidebar rescue failed to recover category headers: %r" % rescued)
        if "Power BI" not in rescued:
            _fail("skill-sidebar rescue failed to recover a description line: %r" % rescued)
        if "Microsoft Office Suite" in clean_exp or "Data Management" in clean_exp:
            _fail("skill-sidebar rescue left the sidebar block behind in experience: %r" % clean_exp)
        if "Prepared 3D modeling" not in clean_exp or "TrackWise and Salesforce" not in clean_exp.split("\n")[1]:
            _fail("skill-sidebar rescue removed genuine experience content: %r" % clean_exp)
        clean_only = "- Did X\n- Did Y\nManaged the team and delivered projects on schedule"
        clean2, rescued2 = rescue(clean_only)
        if clean2 != clean_only or rescued2:
            _fail("skill-sidebar rescue wrongly altered clean experience: %r / %r" % (clean2, rescued2))

    # Font-based name extraction must merge a genuinely wrapped two-line name
    # (different Y position, same style) but must NOT merge same-line text-run
    # fragments some resume-builder PDF exports emit as separate PyMuPDF "line"
    # entries at the same Y (e.g. a job-title line split into "Automation" +
    # "Test" runs) — that previously produced a fake merged "name" like
    # "Automation Test" pulled straight out of the job title, not the person's
    # actual name.
    name_fn = getattr(app, "_extract_name_from_pdf_fonts", None)
    if name_fn is None:
        _fail("_extract_name_from_pdf_fonts helper is missing")
    else:
        from reportlab.pdfgen import canvas as _canvas
        same_line_path = os.path.join(tempfile.mkdtemp(prefix="parser_guard_"), "same_line.pdf")
        c = _canvas.Canvas(same_line_path)
        c.setFont("Helvetica-Bold", 18)
        c.drawString(72, 760, "Riley")
        c.drawString(72, 732, "Morgan")
        # Two separate drawString calls at the IDENTICAL y — simulates a PDF
        # exporter emitting one text run per word on a single visual line.
        c.setFont("Helvetica", 10)
        c.drawString(72, 704, "Automation")
        c.drawString(140, 704, "Test")
        c.drawString(72, 690, "Lead")
        c.save()
        got_name = name_fn(same_line_path)
        if got_name != "Riley Morgan":
            _fail("font-based name extraction gave %r, expected the wrapped two-line "
                  "name 'Riley Morgan' merged correctly" % got_name)

        # Real-world case: Saravana Kumar Sathiamoorthy's resume styled the job
        # title bold at 10.8pt while the actual name sat above it, unstyled, at
        # 24pt. Being the largest text on the page must win over a smaller
        # bold line — otherwise the job title's name-shaped prefix ("Automation
        # Test", since "...Lead" fails the name check) gets returned instead.
        bold_title_path = os.path.join(tempfile.mkdtemp(prefix="parser_guard_"), "bold_title.pdf")
        c2 = _canvas.Canvas(bold_title_path)
        c2.setFont("Helvetica", 24)
        c2.drawString(72, 760, "Jamie Rivera")
        c2.setFont("Helvetica-Bold", 11)
        c2.drawString(72, 720, "Automation Test Lead")
        c2.save()
        got_name2 = name_fn(bold_title_path)
        if got_name2 != "Jamie Rivera":
            _fail("font-based name extraction gave %r — a smaller bold job-title "
                  "line outscored the actual (larger, unstyled) name" % got_name2)

    if not FAILURES:
        print("  OK: skills + experience + heading-tail + keyword helpers all correct")


def main():
    print("=" * 64)
    print("Resume-parser regression guard")
    print("=" * 64)
    check_undefined_names()
    check_parsing_runtime()
    check_skill_fragment_helper()
    print("-" * 64)
    if FAILURES:
        print("GUARD FAILED with %d issue(s):" % len(FAILURES))
        for f in FAILURES:
            print("  - " + f)
        print("\nDo not commit until these are resolved.")
        return 1
    print("GUARD PASSED — parser imports cleanly and extracts core sections.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
