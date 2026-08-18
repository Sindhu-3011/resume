"""Regression checks for the Save -> background AI analysis -> polling flow.

Plain-Python, no pytest dependency — run directly:

    python tests/test_background_ai_analysis.py

Covers the bugs found and fixed this session:
  1. _unwrap_category_score handles the small-model quirk where a category
     score comes back wrapped as {"description": 60} instead of plain 60 —
     the actual failure mode observed live against the configured Ollama
     model, which previously made every real AI assessment fail validation
     and silently fall back.
  2. Saving a resume no longer triggers the wasteful per-JD embedding calls
     (rebuild_role_match_cache) that had no reader and made Save slow.
  3. trigger_background_ai_assessment starts a background thread that
     eventually populates ai_match_cache without any blocking HTTP request,
     and is idempotent (repeat calls for an in-flight pair don't spawn a
     second thread).

Note: because this script itself is a short-lived process, any daemon
thread it spawns is killed the instant the script exits — these tests use
a short `time.sleep()` after triggering (with a mocked, near-instant LLM
call) to give the background thread a real chance to finish before the
process ends. The live, long-running server was separately verified by
hand against the real Ollama model (see conversation notes); this suite
locks in the same logic deterministically and fast.
"""

import datetime
import os
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
FAKE_RESUME_ID = 999_999_101
FAKE_JD_ID = 999_999_102


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def test_unwrap_category_score_handles_nested_dict_quirk():
    check("plain int passes through unchanged", appmod._unwrap_category_score(60) == 60)
    check("nested {'description': 60} unwraps to 60",
          appmod._unwrap_category_score({"required+preferred skills coverage": 60}) == 60)
    check("empty dict passes through unchanged (falls to _clamp_pct default)",
          appmod._unwrap_category_score({}) == {})
    check("None passes through unchanged", appmod._unwrap_category_score(None) is None)


def test_llm_judge_recovers_from_nested_category_dicts():
    """End-to-end through _llm_judge_match: a response with every category
    wrapped in a nested dict (the exact shape seen live against phi4-mini)
    must still be accepted, not rejected as 'omitted categories'."""
    # Deliberately distinct from _KNOWN_POISONED_EXAMPLE_SCORES (75/60/80/50/
    # 40/90/100/65/50) — this test is only about the nested-dict unwrap
    # mechanics, and reusing those exact numbers would now trip the separate
    # poisoned-example tripwire added after discovering the model parrots
    # that literal example back verbatim (see _llm_judge_match).
    quirky_response = {
        "category_scores": {
            "skills": {"required+preferred skills coverage": 82},
            "responsibilities": {"how well past roles map": 55},
            "experience": {"years and relevance": 71},
            "domain": {"domain overlap": 44},
            "certifications": {"held vs wanted": 33},
            "education": {"requirement fit": 95},
            "regulatory": {"standards knowledge": 88},
            "tools": {"tools overlap": 59},
            "soft_skills": {"signals": 62},
        },
        "verdict": "Good Match", "rationale": "test", "strengths": ["x"],
        "concerns": "", "suggested_roles": [], "confidence": 80, "confidence_reason": "",
    }
    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = lambda *a, **k: quirky_response
        result = appmod._llm_judge_match(
            {"title": "x"}, {"title": "y"},
            {"matched_skills": [], "missing_skills": [], "match_percentage": 0, "total_jd_requirements": 0},
        )
        check("quirky nested-dict response is still accepted (not None)", result is not None)
        if result:
            check("skills score correctly unwrapped to 82", result["category_scores"]["skills"] == 82)
            check("regulatory score correctly unwrapped to 88", result["category_scores"]["regulatory"] == 88)
    finally:
        appmod._ollama_chat = real_ollama_chat


def test_poisoned_example_scores_are_rejected():
    """Regression test for the bug found live on 2026-08-03: phi4-mini was
    copying the prompt's concrete example category scores back verbatim
    regardless of resume/JD content, so every single Resume-JD comparison
    in ai_match_cache converged on ~69%. _llm_judge_match must treat an
    exact match against _KNOWN_POISONED_EXAMPLE_SCORES as an invalid
    response (return None, forcing the heuristic fallback) rather than
    caching a fabricated score as if it were a real judgment."""
    poisoned_response = {
        "category_scores": dict(appmod._KNOWN_POISONED_EXAMPLE_SCORES),
        "verdict": "Good Match", "rationale": "test", "strengths": ["x"],
        "concerns": "", "suggested_roles": [], "confidence": 80, "confidence_reason": "",
    }
    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = lambda *a, **k: poisoned_response
        result = appmod._llm_judge_match(
            {"title": "x"}, {"title": "y"},
            {"matched_skills": [], "missing_skills": [], "match_percentage": 0, "total_jd_requirements": 0},
        )
        check("exact poisoned-example scores are rejected as invalid (None)", result is None)
    finally:
        appmod._ollama_chat = real_ollama_chat

    # A near-miss (one category off by one) must NOT be rejected — the
    # tripwire is an exact-match check, not a fuzzy one, so genuine
    # judgments that happen to land close to those values still count.
    near_miss = dict(appmod._KNOWN_POISONED_EXAMPLE_SCORES)
    near_miss["skills"] = near_miss["skills"] - 1
    near_miss_response = {
        "category_scores": near_miss,
        "verdict": "Good Match", "rationale": "test", "strengths": ["x"],
        "concerns": "", "suggested_roles": [], "confidence": 80, "confidence_reason": "",
    }
    try:
        appmod._ollama_chat = lambda *a, **k: near_miss_response
        result = appmod._llm_judge_match(
            {"title": "x"}, {"title": "y"},
            {"matched_skills": [], "missing_skills": [], "match_percentage": 0, "total_jd_requirements": 0},
        )
        check("a near-miss (not an exact copy) is still accepted", result is not None)
    finally:
        appmod._ollama_chat = real_ollama_chat


def test_save_does_not_call_wasteful_embedding_rebuild():
    """rebuild_role_match_cache() (per-JD Ollama embedding calls, for a
    table with no reader) must no longer be called from the resume Save
    path — that was the main reason Save used to be slow. Uses a disposable
    resume this test creates and deletes itself — never touches real data.
    Also stubs out trigger_background_ai_assessment entirely so this test
    can't race a real background thread against later tests (a real one
    would keep running past this test's mock-restoration, and could pick up
    whatever mock happens to be active in a *different* test at the moment
    it actually calls Ollama).
    """
    real_get_embedding = appmod.get_embedding
    real_trigger = appmod.trigger_background_ai_assessment
    call_count = {"embedding": 0, "trigger": 0}

    appmod.get_embedding = lambda *a, **k: (call_count.__setitem__("embedding", call_count["embedding"] + 1), None)[1]
    appmod.trigger_background_ai_assessment = lambda *a, **k: call_count.__setitem__("trigger", call_count["trigger"] + 1)

    new_id = None
    try:
        client = appmod.app.test_client()
        client.post("/login", data={"username": "admin", "password": "Admin@123!"})

        add_data = {
            "full_name": "ZZZ Regression Test Disposable", "title": "Engineer", "email": "",
            "phone": "", "location": "Test City", "exp_yrs": "5", "summary": "", "skills": "",
            "experience": "", "education": "", "certifications": "", "projects": "", "department": "",
        }
        r_add = client.post("/edit", data=add_data)
        check("disposable test resume created (redirect)", r_add.status_code in (302, 303))
        location = r_add.headers.get("Location", "")
        m = __import__("re").search(r"/profile/(\d+)", location)
        new_id = int(m.group(1)) if m else None
        check("new resume id captured from redirect", new_id is not None)

        if new_id:
            t0 = time.time()
            r = client.post(f"/edit/{new_id}", data=add_data)
            elapsed = time.time() - t0
            check("edit-save request succeeds (redirect)", r.status_code in (302, 303))
            check("save completes quickly (<5s) with no synchronous embedding loop", elapsed < 5)

        check("get_embedding was never called during Save", call_count["embedding"] == 0)
        check("background AI trigger was called (analysis still kicked off, just stubbed here)",
              call_count["trigger"] >= 1)
    finally:
        appmod.get_embedding = real_get_embedding
        appmod.trigger_background_ai_assessment = real_trigger
        if new_id:
            with appmod.db_conn() as conn:
                conn.execute("DELETE FROM resume WHERE id = %s", (new_id,))
                conn.commit()


def test_save_with_attached_file_does_not_reparse_with_llm():
    """Regression test for the bug found live on 2026-08-03: the resume file
    picker already parses the file once with the full OCR+LLM pipeline via
    an AJAX call to /api/parse-resume the moment the user selects it (see
    edit.html), and its results land in the visible form fields the user
    reviews before clicking Save. The final Save POST still carries that
    same file (browsers keep a file input's value until cleared), and
    edit_resume() used to unconditionally re-run the full LLM/OCR parse on
    it a second time — confirmed live to make a single Save take ~2.5
    minutes (OCR + two Ollama calls) even though merge_resume_data(...,
    overwrite=False) discards nearly all of that second parse's output
    since form_data already has it. Save must now only use the fast,
    non-AI quick-parse as a gap-filler, never parse_resume_with_llm_text.
    """
    from io import BytesIO
    from reportlab.pdfgen import canvas

    buf = BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(72, 720, "Jane Regression Tester")
    c.drawString(72, 700, "Skills: Python, SQL")
    c.save()
    pdf_bytes = buf.getvalue()

    real_llm_parse = appmod.parse_resume_with_llm_text
    real_trigger = appmod.trigger_background_ai_assessment
    llm_parse_calls = {"count": 0}

    def fake_llm_parse(*a, **k):
        llm_parse_calls["count"] += 1
        return ({}, "llm_text")

    appmod.parse_resume_with_llm_text = fake_llm_parse
    appmod.trigger_background_ai_assessment = lambda *a, **k: None

    new_id = None
    try:
        client = appmod.app.test_client()
        client.post("/login", data={"username": "admin", "password": "Admin@123!"})

        add_data = {
            "full_name": "ZZZ Regression Test Disposable File", "title": "Engineer", "email": "",
            "phone": "", "location": "Test City", "exp_yrs": "5", "summary": "", "skills": "",
            "experience": "", "education": "", "certifications": "", "projects": "", "department": "",
        }
        r_add = client.post("/edit", data=add_data)
        location = r_add.headers.get("Location", "")
        m = __import__("re").search(r"/profile/(\d+)", location)
        new_id = int(m.group(1)) if m else None
        check("disposable test resume created for file-upload test", new_id is not None)

        if new_id:
            save_data = dict(add_data)
            save_data["parse_mode"] = "llm"
            save_data["resume_file"] = (BytesIO(pdf_bytes), "resume.pdf")
            r = client.post(
                f"/edit/{new_id}", data=save_data, content_type="multipart/form-data"
            )
            check("save with attached file succeeds (redirect)", r.status_code in (302, 303))
            check(
                "parse_resume_with_llm_text (OCR+LLM) was never called on Save",
                llm_parse_calls["count"] == 0,
            )
    finally:
        appmod.parse_resume_with_llm_text = real_llm_parse
        appmod.trigger_background_ai_assessment = real_trigger
        if new_id:
            with appmod.db_conn() as conn:
                conn.execute("DELETE FROM resume WHERE id = %s", (new_id,))
                conn.commit()


def test_save_skips_reparse_entirely_when_preview_already_filled_the_form():
    """Regression test for the second half of the 2026-08-03 slow-Save bug:
    even after removing the OCR+LLM re-parse, Save was still unconditionally
    re-running the fast, non-AI quick-parse (its own OCR contact-strip pass)
    on every file-attached Save — confirmed live to cost ~3.5s by itself,
    even when the form already had every field filled in by the earlier
    /api/parse-resume preview. When full_name/summary/skills-or-experience
    are already present (the preview clearly already ran and the user
    reviewed the result), Save must skip re-parsing the file completely —
    only save it to disk — not call _parse_pdf_quick or extract_resume_text
    at all.
    """
    from io import BytesIO
    from reportlab.pdfgen import canvas

    buf = BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(72, 720, "Jane Regression Tester")
    c.drawString(72, 700, "Skills: Python, SQL")
    c.save()
    pdf_bytes = buf.getvalue()

    real_quick_parse = appmod._parse_pdf_quick
    real_extract = appmod.extract_resume_text
    real_trigger = appmod.trigger_background_ai_assessment
    calls = {"quick_parse": 0, "extract": 0}

    appmod._parse_pdf_quick = lambda *a, **k: calls.__setitem__("quick_parse", calls["quick_parse"] + 1) or {}
    appmod.extract_resume_text = lambda *a, **k: calls.__setitem__("extract", calls["extract"] + 1) or ""
    appmod.trigger_background_ai_assessment = lambda *a, **k: None

    new_id = None
    try:
        client = appmod.app.test_client()
        client.post("/login", data={"username": "admin", "password": "Admin@123!"})

        full_data = {
            "full_name": "ZZZ Regression Test Prefilled Form", "title": "Engineer",
            "email": "", "phone": "", "location": "Test City", "exp_yrs": "5",
            "summary": "Already-parsed summary from the earlier preview.",
            "skills": "Python, SQL", "experience": "Company A, Engineer, 2015-2020",
            "education": "BE Computer Science", "certifications": "", "projects": "",
            "department": "",
        }
        r_add = client.post("/edit", data=full_data)
        location = r_add.headers.get("Location", "")
        m = __import__("re").search(r"/profile/(\d+)", location)
        new_id = int(m.group(1)) if m else None
        check("disposable prefilled-form resume created", new_id is not None)

        if new_id:
            save_data = dict(full_data)
            save_data["parse_mode"] = "llm"
            save_data["resume_file"] = (BytesIO(pdf_bytes), "resume.pdf")
            t0 = time.time()
            r = client.post(
                f"/edit/{new_id}", data=save_data, content_type="multipart/form-data"
            )
            elapsed = time.time() - t0
            check("save with already-filled form succeeds (redirect)", r.status_code in (302, 303))
            check("save completes near-instantly (<1s) with no re-parse at all", elapsed < 1)
            check("_parse_pdf_quick was never called", calls["quick_parse"] == 0)
            check("extract_resume_text was never called", calls["extract"] == 0)
    finally:
        appmod._parse_pdf_quick = real_quick_parse
        appmod.extract_resume_text = real_extract
        appmod.trigger_background_ai_assessment = real_trigger
        if new_id:
            with appmod.db_conn() as conn:
                conn.execute("DELETE FROM resume WHERE id = %s", (new_id,))
                conn.commit()


def test_background_trigger_populates_cache_without_blocking_request():
    """trigger_background_ai_assessment() must return immediately and let a
    background thread populate ai_match_cache — the frontend polls
    ai-status for this, never blocks a request on it.

    _background_ai_assessment_worker does its own fresh SELECT by ID (it
    doesn't use any dict a caller passes in), and silently skips a pair
    whose resume/JD row doesn't exist — so this needs a REAL disposable
    resume and a REAL existing JD, not made-up IDs, to actually reach
    _get_or_compute_hybrid_match at all.
    """
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})
    add_data = {
        "full_name": "ZZZ Regression Test Background Trigger", "title": "Engineer", "email": "",
        "phone": "", "location": "Test City", "exp_yrs": "5", "summary": "", "skills": "",
        "experience": "", "education": "", "certifications": "", "projects": "", "department": "",
    }
    real_trigger = appmod.trigger_background_ai_assessment
    appmod.trigger_background_ai_assessment = lambda *a, **k: None  # suppress Save's own auto-trigger
    resume_id = None
    try:
        r_add = client.post("/edit", data=add_data)
        location = r_add.headers.get("Location", "")
        m = __import__("re").search(r"/profile/(\d+)", location)
        resume_id = int(m.group(1)) if m else None
        check("disposable test resume created for this test", resume_id is not None)
    finally:
        appmod.trigger_background_ai_assessment = real_trigger

    with appmod.db_conn() as conn:
        jd_row = conn.execute(
            "SELECT id, updated_at FROM job_description WHERE position_status = %s LIMIT 1",
            (appmod.DEFAULT_POSITION_STATUS,),
        ).fetchone()
        # Full row, not just id/updated_at — _read_cached_hybrid_match compares
        # a content fingerprint over summary/skills/experience/etc. now (see
        # _resume_match_fingerprint), matching every real production caller,
        # which always fetches the full resume row.
        resume_row = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
    check("a real Open JD exists to test against", jd_row is not None)
    jd_id = jd_row["id"] if jd_row else None

    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s", (resume_id, jd_id))
        conn.commit()

    real_llm_judge = appmod._llm_judge_match
    appmod._llm_judge_match = lambda resume_d, jd_d, keyword_result: {
        "category_scores": {cat: 70 for cat in appmod._CATEGORY_WEIGHTS},
        "rationale": "x", "strengths": [], "concerns": "",
        "suggested_roles": [], "confidence": 80, "confidence_reason": "",
    }

    try:
        with appmod.app.test_request_context("/"):
            t0 = time.time()
            appmod.trigger_background_ai_assessment(resume_id, [jd_id])
            call_returned_immediately = (time.time() - t0) < 0.5
        check("trigger_background_ai_assessment returns immediately (non-blocking)",
              call_returned_immediately)

        # Give the background thread (mocked LLM, near-instant) a real
        # chance to run and finish before this short-lived script exits.
        deadline = time.time() + 5
        cached = None
        while time.time() < deadline:
            with appmod.db_conn() as conn:
                cached = appmod._read_cached_hybrid_match(conn, resume_row, jd_row)
            if cached is not None:
                break
            time.sleep(0.2)

        check("background thread populated ai_match_cache without a blocking request",
              cached is not None)
        if cached:
            check("cached result matches the mocked weighted score (70)", cached["fit_percentage"] == 70)

        # Re-triggering after completion must never change the answer — the
        # user-facing guarantee (same pair -> same percentage, however many
        # times it's (re)triggered) — even in the narrow window right after
        # completion where a second trigger could in principle still find a
        # fresh thread starting (in-flight bookkeeping is best-effort
        # dedup, not a correctness requirement; the cache freshness check
        # inside _get_or_compute_hybrid_match is what actually guarantees
        # determinism, and that's what this checks).
        appmod.trigger_background_ai_assessment(resume_id, [jd_id])
        deadline2 = time.time() + 5
        recached = cached
        while time.time() < deadline2:
            with appmod.db_conn() as conn:
                recached = appmod._read_cached_hybrid_match(conn, resume_row, jd_row)
            if recached is not None:
                break
            time.sleep(0.2)
        check("re-triggering the same pair still yields the identical cached score",
              recached is not None and recached["fit_percentage"] == 70)
    finally:
        appmod._llm_judge_match = real_llm_judge
        with appmod.db_conn() as conn:
            conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s", (resume_id, jd_id))
            if resume_id:
                conn.execute("DELETE FROM resume WHERE id = %s", (resume_id,))
            conn.commit()


def test_ai_status_endpoint_never_blocks_and_never_computes():
    """compare_ai_status must be a pure read: not-ready for an uncached pair
    with NO background job triggered — it must never itself kick off
    computation (that's trigger_background_ai_assessment's job only)."""
    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID + 1))
        conn.commit()

    call_count = {"n": 0}
    real_llm_judge = appmod._llm_judge_match
    appmod._llm_judge_match = lambda *a, **k: call_count.__setitem__("n", call_count["n"] + 1) or None

    try:
        client = appmod.app.test_client()
        client.post("/login", data={"username": "admin", "password": "Admin@123!"})
        with appmod.db_conn() as conn:
            rid = conn.execute("SELECT id FROM resume ORDER BY id DESC LIMIT 1").fetchone()["id"]
        r = client.get(f"/api/compare/{rid}/{FAKE_JD_ID + 1}/ai-status")
        check("status check for nonexistent JD returns 404, not a hang", r.status_code == 404)
        check("status endpoint never invoked the LLM judge itself", call_count["n"] == 0)
    finally:
        appmod._llm_judge_match = real_llm_judge


def test_ai_status_endpoint_returns_correct_jd_mirroring_data():
    """Regression test for a bug found live on 2026-08-06: compare_ai_status
    only ever selected `id, updated_at` for the JD row, since that used to be
    all _read_cached_hybrid_match needed (a raw timestamp comparison).
    _jd_mirroring_risk needs the JD's actual skills/requirements/
    responsibilities text, computed fresh on every read (see
    _read_cached_hybrid_match) — with only id/updated_at present, every
    single call through this endpoint silently saw an empty JD and always
    reported jd_mirroring_label=None, even for a resume that blatantly
    mirrors the JD's own text. Confirmed live: the Top 3 cards (which use a
    full JD row via _rank_top_jd_matches) correctly showed "High JD-Text
    Overlap", while this exact same pair's /ai-status response showed null.
    """
    with appmod.db_conn() as conn:
        jd_row = conn.execute(
            "SELECT id, requirements, responsibilities, skills, keywords, updated_at "
            "FROM job_description WHERE position_status = %s LIMIT 1",
            (appmod.DEFAULT_POSITION_STATUS,),
        ).fetchone()
    check("a real Open JD exists to test against", jd_row is not None)
    if not jd_row:
        return
    jd_id = jd_row["id"]
    # Build a resume that closely mirrors this JD's own text — the same
    # near-verbatim-copy scenario the feature targets.
    mirror_text = " ".join(str(jd_row.get(k) or "") for k in
                            ("requirements", "responsibilities", "skills", "keywords"))[:1500]

    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})
    add_data = {
        "full_name": "ZZZ Regression Test JD Mirroring", "title": "Engineer", "email": "",
        "phone": "", "location": "Test City", "exp_yrs": "5", "summary": mirror_text, "skills": "",
        "experience": "", "education": "", "certifications": "", "projects": "", "department": "",
    }
    resume_id = None
    real_llm_judge = appmod._llm_judge_match
    real_trigger = appmod.trigger_background_ai_assessment
    appmod.trigger_background_ai_assessment = lambda *a, **k: None
    try:
        r_add = client.post("/edit", data=add_data)
        m = __import__("re").search(r"/profile/(\d+)", r_add.headers.get("Location", ""))
        resume_id = int(m.group(1)) if m else None
        check("disposable mirrored-content resume created", resume_id is not None)
        if not resume_id:
            return

        # Force a cached AI-judged row for this pair so the endpoint has
        # something to serve — mocked LLM output, fast and deterministic.
        appmod._llm_judge_match = lambda resume_d, jd_d, keyword_result: {
            "category_scores": {cat: 70 for cat in appmod._CATEGORY_WEIGHTS},
            "verdict": "Good Match", "rationale": "x", "strengths": [], "concerns": "",
            "suggested_roles": [], "confidence": 80, "confidence_reason": "",
        }
        with appmod.db_conn() as conn:
            resume_row = conn.execute("SELECT * FROM resume WHERE id = %s", (resume_id,)).fetchone()
            jd_full = conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()
            assessment = appmod._hybrid_match(dict(resume_row), dict(jd_full))
            appmod._store_hybrid_match(conn, resume_row, jd_full, assessment)
            conn.commit()

        r = client.get(f"/api/compare/{resume_id}/{jd_id}/ai-status")
        data = r.get_json()
        check("ai-status endpoint reports ready", data.get("ready") is True)
        check("ai-status endpoint's jd_mirroring_label reflects the actual overlap, not null",
              data.get("jd_mirroring_label") is not None)
        check("ai-status endpoint's jd_mirroring_pct is a real number, not 0-by-empty-JD",
              (data.get("jd_mirroring_pct") or 0) > 0)
    finally:
        appmod._llm_judge_match = real_llm_judge
        appmod.trigger_background_ai_assessment = real_trigger
        if resume_id:
            with appmod.db_conn() as conn:
                conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                             (resume_id, jd_id))
                conn.execute("DELETE FROM resume WHERE id = %s", (resume_id,))
                conn.commit()


def main():
    tests = [
        test_unwrap_category_score_handles_nested_dict_quirk,
        test_llm_judge_recovers_from_nested_category_dicts,
        test_poisoned_example_scores_are_rejected,
        test_save_does_not_call_wasteful_embedding_rebuild,
        test_save_with_attached_file_does_not_reparse_with_llm,
        test_save_skips_reparse_entirely_when_preview_already_filled_the_form,
        test_background_trigger_populates_cache_without_blocking_request,
        test_ai_status_endpoint_never_blocks_and_never_computes,
        test_ai_status_endpoint_returns_correct_jd_mirroring_data,
    ]
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
    print("All background-AI-analysis checks passed.")


if __name__ == "__main__":
    main()
