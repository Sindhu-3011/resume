"""Regression test for the shared-temp-file race condition found live on
2026-08-05 while investigating a report that resume parsing "was working
fine" and then intermittently broke again for the same file (Arun
Sivakumar's and Satheesh's resumes).

Plain-Python, no pytest dependency — run directly:

    python tests/test_upload_concurrency.py

Root cause: /api/parse-resume wrote every uploaded file to a FIXED path —
UPLOAD_FOLDER / f"tmp-parse.{ext}" — regardless of who was using it. The
dev server runs threaded (concurrent requests), so two people (or the same
person, two browser tabs) previewing a resume around the same time could
overwrite or delete each other's temp file mid-parse: one request reads
back a different file than it just uploaded, or its file vanishes (unlink
in the other request's `finally`) while still being read. Confirmed live
against the real server with two real, different resumes fired
concurrently — each request correctly got back only its own parsed
content, no cross-contamination.

Fix: the temp path is now unique per request (uuid4-suffixed). This test
proves the actual guarantee that eliminates the race — every call gets a
never-before-seen path, structurally, not just "usually" under favorable
timing — rather than trying to catch the race in the act under real
thread scheduling, which is inherently flaky to assert on in an automated
suite.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def test_parse_preview_never_reuses_the_same_temp_path():
    """Every call to /api/parse-resume must get its own, never-before-seen
    temp file path. The old code used a fixed
    UPLOAD_FOLDER/"tmp-parse.<ext>" path derived only from the file
    extension, so ANY two calls (same filename or not, concurrent or even
    just close in time) collided on the identical path every single time —
    that's what let two people (or two tabs) previewing a resume around
    the same time overwrite or delete each other's temp file mid-parse.
    """
    from io import BytesIO
    from reportlab.pdfgen import canvas

    def make_pdf(marker):
        buf = BytesIO()
        c = canvas.Canvas(buf)
        c.drawString(72, 720, f"Full Name: {marker} Testperson")
        c.save()
        return buf.getvalue()

    real_parse_llm = appmod.parse_resume_with_llm_text
    seen_paths = []

    def fake_parse(path):
        seen_paths.append(str(path))
        return ({"full_name": "x"}, "llm_text")

    appmod.parse_resume_with_llm_text = fake_parse
    try:
        client = appmod.app.test_client()
        client.post("/login", data={"username": "admin", "password": "Admin@123"})
        for marker in ("Alpha", "Bravo", "Charlie"):
            # Same original filename every time ("resume.pdf") — the old
            # fixed-path code derived the temp name only from the
            # extension, so an identical original filename each call is
            # exactly the scenario that always collided before.
            r = client.post(
                "/api/parse-resume",
                data={"resume_file": (BytesIO(make_pdf(marker)), "resume.pdf"), "parse_mode": "llm"},
                content_type="multipart/form-data",
            )
            check(f"{marker} preview call succeeded", r.status_code == 200 and r.get_json().get("success"))
    finally:
        appmod.parse_resume_with_llm_text = real_parse_llm

    check("3 temp paths captured", len(seen_paths) == 3)
    check("every temp path used was unique — no two calls ever share a path",
          len(set(seen_paths)) == len(seen_paths))
    check("the old fixed 'tmp-parse.<ext>' path is no longer used",
          all(os.path.basename(p) != "tmp-parse.pdf" for p in seen_paths))


if __name__ == "__main__":
    tests = [test_parse_preview_never_reuses_the_same_temp_path]
    for t in tests:
        print(f"\n-- {t.__name__} --")
        try:
            t()
        except Exception:
            import traceback
            print(f"[ERROR] {t.__name__} raised an exception:")
            traceback.print_exc()
            failures.append(t.__name__)

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        print("Failures:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    else:
        print("All upload-concurrency checks passed.")
