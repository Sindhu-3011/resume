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
