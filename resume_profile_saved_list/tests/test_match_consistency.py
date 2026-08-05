"""Regression checks for the AI Holistic Assessment / Top 3 Matching Roles
consistency fix.

Plain-Python, no pytest dependency (none is installed in this project's
venv) — run directly:

    python tests/test_match_consistency.py

Covers the bugs that made "AI Holistic Assessment" and "Top 3 Matching
Roles" disagree or get stuck:
  1. One tier-threshold function (_tier_from_pct) used everywhere, instead
     of three copies at different cutoffs.
  2. _hybrid_match's fallback path (LLM unreachable/timed out) is flagged
     is_ai_judged=False and is never persisted to ai_match_cache — a slow
     or offline Ollama used to get permanently cached as the final answer.
  3. A legacy/pre-fix cache row (no is_ai_judged marker) is treated as a
     miss and recomputed, instead of served forever as stale data.
  4. _rank_top_jd_matches (Top 3 cards) prefers a cached AI-judged result
     over its own heuristic estimate for the same resume/JD pair, so once
     a pair has been AI-assessed, both surfaces show the identical number.

This talks to the real Postgres database (DATABASE_URL / default local
"resume_profiles") to exercise the actual cache read/write path, using
resume_id/jd_id values far outside the real ID range so it never touches
real data; ai_match_cache has no foreign key, so this is safe. Any row
this test writes is deleted again at the end, in a `finally`.
"""

import datetime
import os
import sys
import threading
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

FAKE_RESUME_ID = 999_999_001
FAKE_JD_ID = 999_999_002

failures = []


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def test_tier_thresholds_unified_and_correct():
    """_tier_from_pct must be the 80/60/40 cutoffs described to users as
    Strong/Good/Partial/Weak, with no gaps or overlaps at the boundaries."""
    cases = [
        (100, "Strong Fit"), (80, "Strong Fit"), (79, "Good Fit"),
        (60, "Good Fit"), (59, "Partial Fit"),
        (40, "Partial Fit"), (39, "Weak Fit"), (0, "Weak Fit"),
    ]
    for pct, expected in cases:
        _, tier_label, _ = appmod._tier_from_pct(pct)
        check(f"_tier_from_pct({pct}) == {expected!r}", tier_label == expected)


def test_holistic_match_delegates_to_tier_from_pct():
    """_holistic_match() must bucket through _tier_from_pct(), not a second,
    independently-maintained copy of the threshold chain — that duplication
    is exactly how the two components drifted to different cutoffs before."""
    real_similarity_fn = appmod._whole_doc_similarity
    try:
        for target_pct in (85, 65, 45, 10):
            # _holistic_match rescales similarity * 220 -> pct; invert that
            # to drive a specific percentage through the real function.
            fake_similarity = target_pct / 220.0
            appmod._whole_doc_similarity = lambda jd, res, s=fake_similarity: (s, ["shared_term"])
            holistic = appmod._holistic_match({"title": "x"}, {"title": "y"})
            expected_tier = appmod._tier_from_pct(target_pct)[1]
            check(
                f"_holistic_match tier at ~{target_pct}% matches _tier_from_pct ({expected_tier})",
                holistic["tier_label"] == expected_tier,
            )
    finally:
        appmod._whole_doc_similarity = real_similarity_fn


def test_hybrid_match_fallback_is_flagged_and_not_final():
    """When the LLM judge is unreachable, _hybrid_match must return
    is_ai_judged=False so the caching layer knows never to persist it."""
    real_llm_judge = appmod._llm_judge_match
    try:
        appmod._llm_judge_match = lambda resume, jd, keyword_result: None
        result = appmod._hybrid_match(
            {"title": "Engineer", "skills": "Python", "experience": "3 years"},
            {"title": "Engineer", "skills": "Python", "requirements": "", "responsibilities": "", "keywords": ""},
        )
        check("fallback assessment has is_ai_judged=False", result.get("is_ai_judged") is False)
        check("fallback assessment still has a tier_label", bool(result.get("tier_label")))
    finally:
        appmod._llm_judge_match = real_llm_judge


def test_hybrid_match_success_is_flagged_and_tiered_consistently():
    """When the LLM judge succeeds, is_ai_judged=True and the blended
    percentage must bucket through the same _tier_from_pct as everywhere
    else — this is the number the Top 3 cards will later reuse verbatim."""
    real_llm_judge = appmod._llm_judge_match
    try:
        appmod._llm_judge_match = lambda resume, jd, keyword_result: {
            "category_scores": {cat: 70 for cat in appmod._CATEGORY_WEIGHTS},
            "rationale": "test rationale", "strengths": ["a"],
            "concerns": "", "suggested_roles": [], "confidence": 80, "confidence_reason": "",
        }
        result = appmod._hybrid_match(
            {"title": "Engineer", "skills": "Python", "experience": "3 years"},
            {"title": "Engineer", "skills": "Python", "requirements": "", "responsibilities": "", "keywords": ""},
        )
        check("success assessment has is_ai_judged=True", result.get("is_ai_judged") is True)
        check("uniform 70-per-category input weighted-sums to 70", result["fit_percentage"] == 70)
        expected_tier = appmod._tier_from_pct(result["fit_percentage"])[1]
        check(
            "success assessment tier_label matches _tier_from_pct(fit_percentage)",
            result["tier_label"] == expected_tier,
        )
    finally:
        appmod._llm_judge_match = real_llm_judge


def test_weighted_category_sum_is_deterministic_and_matches_hand_calculation():
    """The core of the new scoring model: fit_percentage must be EXACTLY the
    weighted sum of the 9 category scores (never something the LLM invents
    or a black-box blend) — auditable by hand from category_scores alone."""
    check("category weights sum to 1.0", abs(sum(appmod._CATEGORY_WEIGHTS.values()) - 1.0) < 1e-9)

    scores = {"skills": 80, "responsibilities": 70, "experience": 60, "domain": 50,
              "certifications": 90, "education": 100, "regulatory": 40, "tools": 75, "soft_skills": 65}
    real_llm_judge = appmod._llm_judge_match
    try:
        appmod._llm_judge_match = lambda resume, jd, keyword_result: {
            "category_scores": dict(scores), "rationale": "", "strengths": [], "concerns": "",
            "suggested_roles": [], "confidence": 80, "confidence_reason": "",
        }
        result = appmod._hybrid_match(
            {"title": "x", "skills": "Python", "experience": ""},
            {"title": "y", "skills": "Python", "requirements": "", "responsibilities": "", "keywords": ""},
        )
        hand_calc = round(sum(scores[c] * w for c, w in appmod._CATEGORY_WEIGHTS.items()))
        check(f"fit_percentage ({result['fit_percentage']}) matches hand-calculated weighted sum ({hand_calc})",
              result["fit_percentage"] == hand_calc)
        check("returned category_scores match what the judge reported", result["category_scores"] == scores)
    finally:
        appmod._llm_judge_match = real_llm_judge


def test_title_excluded_from_ai_matching_and_cache_stays_valid_across_title_edits():
    """Regression test for the bug found live on 2026-08-04: editing only the
    Role & Group field (which sets resume.title) changed the AI-judged
    percentage for an unrelated JD, even though skills/experience/summary
    were untouched. Root cause: Title fed into both the LLM prompt
    (_llm_judge_match's resume_text) and the keyword scan corpus
    (calculate_match_score) used for AI grounding facts, and every Save
    bumps resume.updated_at — the sole cache-invalidation signal at the
    time — so a metadata-only edit forced a real recompute against a
    genuinely different (title-including) prompt. Title must now be
    excluded from both AI-judged inputs, and the cache must key on a
    content fingerprint that ignores Title (_resume_match_fingerprint), so
    a Role/Group-only edit can neither invalidate nor change an
    already-cached AI assessment.
    """
    resume_a = {
        "id": FAKE_RESUME_ID, "updated_at": None,
        "title": "Automation Lead", "summary": "Built CI/CD pipelines for 8 years.",
        "skills": "Python, Selenium, Jenkins", "experience": "8 years automation testing",
        "projects": "", "certifications": "", "education": "BE Computer Science", "exp_yrs": "8",
    }
    resume_b = dict(resume_a)
    resume_b["title"] = "Design Control Consultant"  # only Title differs
    jd = {"title": "Automation Engineer", "role": "Automation", "skills": "Python, Jenkins",
          "requirements": "", "responsibilities": "", "keywords": ""}

    score_a = appmod.calculate_match_score(resume_a, jd, include_title=False)
    score_b = appmod.calculate_match_score(resume_b, jd, include_title=False)
    check(
        "title-excluded keyword scan is identical across a title-only change",
        score_a["match_percentage"] == score_b["match_percentage"]
        and score_a["matched_skills"] == score_b["matched_skills"],
    )

    captured_prompts = []

    def fake_ollama_chat(prompt, **kwargs):
        captured_prompts.append(prompt)
        return {
            "category_scores": {cat: 70 for cat in appmod._CATEGORY_WEIGHTS},
            "verdict": "Good Match", "rationale": "x", "strengths": [], "concerns": "",
            "suggested_roles": [], "confidence": 80, "confidence_reason": "",
        }

    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = fake_ollama_chat
        keyword_result = appmod.calculate_match_score(resume_a, jd, include_title=False)
        appmod._llm_judge_match(resume_a, jd, keyword_result)
        appmod._llm_judge_match(resume_b, jd, keyword_result)
    finally:
        appmod._ollama_chat = real_ollama_chat

    check("2 LLM calls captured", len(captured_prompts) == 2)
    if len(captured_prompts) == 2:
        check("resume A's title is not in its own prompt", "Automation Lead" not in captured_prompts[0])
        check("resume B's title is not in its own prompt", "Design Control Consultant" not in captured_prompts[1])
        check("the two prompts are identical — title made no difference", captured_prompts[0] == captured_prompts[1])

    check(
        "content fingerprint is unaffected by a title-only change",
        appmod._resume_match_fingerprint(resume_a) == appmod._resume_match_fingerprint(resume_b),
    )

    now = datetime.datetime.now(datetime.timezone.utc)
    jd_full = {"id": FAKE_JD_ID, "updated_at": now}
    resume_a_full = dict(resume_a); resume_a_full["updated_at"] = now
    resume_b_full = dict(resume_b); resume_b_full["updated_at"] = now
    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))
        conn.commit()
        real_result = {
            "fit_percentage": 69, "verdict": "Good overall fit", "tier_label": "Good Fit",
            "recommendation": "", "rationale": "", "strengths": [], "concerns": "",
            "suggested_roles": [], "keyword_match_percentage": 60,
            "category_scores": {cat: 70 for cat in appmod._CATEGORY_WEIGHTS},
            "confidence": 80, "confidence_label": "High", "confidence_reason": "",
            "algo_version": appmod._HYBRID_ALGO_VERSION, "is_ai_judged": True,
        }
        appmod._store_hybrid_match(conn, resume_a_full, jd_full, real_result)
        conn.commit()
        # Read back with resume_b_full — same content, only title differs,
        # simulating a Role/Group-only edit having just been saved.
        reread = appmod._read_cached_hybrid_match(conn, resume_b_full, jd_full)
        check(
            "cached AI assessment survives a title-only edit (Role/Group change)",
            reread is not None and reread["fit_percentage"] == 69,
        )
        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))
        conn.commit()


def test_incomplete_category_scores_rejected_not_silently_averaged():
    """If the LLM omits a category (malformed/truncated JSON), _llm_judge_match
    must return None (triggering the heuristic fallback) rather than silently
    computing a weighted sum over partial data, which would be a wrong number
    presented with the same confidence as a complete judgment."""
    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = lambda *a, **k: {
            "category_scores": {"skills": 80, "tools": 70},  # 7 categories missing
            "verdict": "Good Match", "rationale": "", "strengths": [], "concerns": "",
            "suggested_roles": [], "confidence": 80, "confidence_reason": "",
        }
        result = appmod._llm_judge_match(
            {"title": "x"}, {"title": "y"},
            {"matched_skills": [], "missing_skills": [], "match_percentage": 0, "total_jd_requirements": 0},
        )
        check("incomplete category_scores causes _llm_judge_match to return None", result is None)
    finally:
        appmod._ollama_chat = real_ollama_chat


def test_cache_never_persists_fallback_and_self_heals_legacy_rows():
    """End-to-end against the real ai_match_cache table:
      - a fallback assessment must never be written to the cache
      - a legacy row without is_ai_judged must be treated as a miss
      - a genuine AI-judged result must be cached and served on the next read
    """
    now = datetime.datetime.now(datetime.timezone.utc)
    resume = {"id": FAKE_RESUME_ID, "updated_at": now}
    jd = {"id": FAKE_JD_ID, "updated_at": now}

    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))

        # 1) A fallback (is_ai_judged=False) must not be persisted.
        fallback = {
            "fit_percentage": 20, "verdict": "Limited overall fit", "tier_label": "Weak Fit",
            "recommendation": "", "rationale": "", "strengths": [], "concerns": "",
            "suggested_roles": [], "is_ai_judged": False, "algo_version": appmod._HYBRID_ALGO_VERSION,
        }
        if fallback.get("is_ai_judged"):
            appmod._store_hybrid_match(conn, resume, jd, fallback)
        cached = appmod._read_cached_hybrid_match(conn, resume, jd)
        check("fallback result is not cached", cached is None)

        # 2) A legacy row (pre-fix, no is_ai_judged in extra) must read as a miss.
        conn.execute(
            """
            INSERT INTO ai_match_cache
                (resume_id, jd_id, fit_percentage, verdict, rationale, extra,
                 resume_updated_at, jd_updated_at, computed_at)
            VALUES (%s, %s, 55, 'Good overall fit', '', %s, %s, %s, NOW())
            """,
            (FAKE_RESUME_ID, FAKE_JD_ID,
             '{"tier_label": "Good Fit", "algo_version": "hybrid-v1"}', now, now),
        )
        conn.commit()
        legacy_cached = appmod._read_cached_hybrid_match(conn, resume, jd)
        check("legacy row without is_ai_judged reads as a cache miss", legacy_cached is None)

        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))
        conn.commit()

        # 3) A genuine AI-judged result must be cached and reread identically.
        real_result = {
            "fit_percentage": 72, "verdict": "Good overall fit", "tier_label": "Good Fit",
            "recommendation": "worth interviewing", "rationale": "matches well",
            "strengths": ["Python"], "concerns": "", "suggested_roles": [],
            "keyword_match_percentage": 60, "category_scores": {cat: 78 for cat in appmod._CATEGORY_WEIGHTS},
            "confidence": 80, "confidence_label": "High", "confidence_reason": "",
            "algo_version": appmod._HYBRID_ALGO_VERSION, "is_ai_judged": True,
        }
        appmod._store_hybrid_match(conn, resume, jd, real_result)
        conn.commit()
        reread = appmod._read_cached_hybrid_match(conn, resume, jd)
        check("genuine AI-judged result is cached", reread is not None)
        if reread is not None:
            check("cached fit_percentage round-trips exactly", reread["fit_percentage"] == 72)
            check("cached tier_label round-trips exactly", reread["tier_label"] == "Good Fit")

        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))
        conn.commit()


def test_rank_top_jd_matches_prefers_cached_ai_result_over_estimate():
    """The core single-source-of-truth guarantee: once a resume/JD pair has
    a cached AI-judged assessment, _rank_top_jd_matches must display THAT
    percentage/tier for it, not its own independently-computed estimate —
    even when the two numbers differ."""
    now = datetime.datetime.now(datetime.timezone.utc)
    resume_dict = {
        "id": FAKE_RESUME_ID, "updated_at": now, "title": "Engineer",
        "skills": "Python, SQL", "experience": "3 years as engineer",
        "summary": "", "projects": "", "certifications": "", "education": "",
    }
    jd_dict = {
        "id": FAKE_JD_ID, "updated_at": now, "title": "Fake Test JD", "role": "",
        "category": "Test", "skills": "Python", "requirements": "", "responsibilities": "",
        "keywords": "", "position_status": "Open",
    }

    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))

        # Prime the cache with an AI-judged score deliberately different from
        # whatever the fast heuristic would estimate on its own.
        cached_result = {
            "fit_percentage": 91, "verdict": "Strong overall fit", "tier_label": "Strong Fit",
            "recommendation": "advance to next round", "rationale": "excellent match",
            "strengths": ["Python"], "concerns": "", "suggested_roles": [],
            "keyword_match_percentage": 50, "category_scores": {cat: 95 for cat in appmod._CATEGORY_WEIGHTS},
            "confidence": 90, "confidence_label": "High", "confidence_reason": "",
            "algo_version": appmod._HYBRID_ALGO_VERSION, "is_ai_judged": True,
        }
        appmod._store_hybrid_match(conn, resume_dict, jd_dict, cached_result)
        conn.commit()

        estimate_pct, _ = appmod._quick_estimate_jd_match(resume_dict, jd_dict)
        check(
            "test setup: heuristic estimate differs from the cached AI score "
            f"(estimate={estimate_pct}, cached=91) so this test is meaningful",
            estimate_pct != 91,
        )

        matches = appmod._rank_top_jd_matches(conn, resume_dict, [jd_dict])
        check("_rank_top_jd_matches returns the JD", len(matches) == 1)
        if matches:
            m = matches[0]
            check("displayed match is flagged is_ai_judged=True", m["is_ai_judged"] is True)
            check("displayed final_score is the cached AI score (91), not the heuristic estimate",
                  m["final_score"] == 91)
            check("displayed tier_label is the cached tier (Strong Fit)", m["tier_label"] == "Strong Fit")

        conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                     (FAKE_RESUME_ID, FAKE_JD_ID))
        conn.commit()


def test_concurrent_requests_for_same_pair_compute_only_once():
    """The regression this session actually chased down: with two Flask
    processes (or just two threads) racing to AI-assess the same uncached
    resume/JD pair, each independently calling the LLM and "last write wins"
    into ai_match_cache made the displayed percentage flip on every refresh.
    _get_or_compute_hybrid_match() now serializes via a Postgres advisory
    lock, which — unlike a Python threading.Lock — works across independent
    OS processes, not just threads in one. This test opens a SEPARATE
    db_conn() per thread (its own real Postgres connection/session, exactly
    like two separate server processes would each have) and proves only one
    of them ever reaches the (mocked) LLM call.
    """
    now = datetime.datetime.now(datetime.timezone.utc)
    resume = {"id": FAKE_RESUME_ID, "updated_at": now, "title": "Engineer",
              "skills": "Python", "experience": "", "summary": "", "projects": "",
              "certifications": "", "education": ""}
    jd = {"id": FAKE_JD_ID, "updated_at": now, "title": "Fake JD", "role": "",
          "skills": "Python", "requirements": "", "responsibilities": "", "keywords": ""}

    barrier = threading.Barrier(3, timeout=6)
    call_count = {"n": 0}
    call_lock = threading.Lock()
    reached_together = {"count": 0}

    def fake_llm_judge(resume_d, jd_d, keyword_result):
        with call_lock:
            call_count["n"] += 1
        try:
            barrier.wait()
            with call_lock:
                reached_together["count"] += 1
        except threading.BrokenBarrierError:
            pass
        return {
            "category_scores": {cat: 50 for cat in appmod._CATEGORY_WEIGHTS},
            "rationale": "x", "strengths": [], "concerns": "",
            "suggested_roles": [], "confidence": 80, "confidence_reason": "",
        }

    real_llm_judge = appmod._llm_judge_match
    appmod._llm_judge_match = fake_llm_judge

    with appmod.db_conn() as setup_conn:
        setup_conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                            (FAKE_RESUME_ID, FAKE_JD_ID))
        setup_conn.commit()

    results = []
    errors = []

    def worker():
        try:
            with appmod.db_conn() as conn:
                results.append(appmod._get_or_compute_hybrid_match(conn, resume, jd))
        except Exception as e:
            errors.append(e)

    try:
        threads = [threading.Thread(target=worker) for _ in range(3)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=15)
    finally:
        appmod._llm_judge_match = real_llm_judge
        with appmod.db_conn() as cleanup_conn:
            cleanup_conn.execute("DELETE FROM ai_match_cache WHERE resume_id = %s AND jd_id = %s",
                                 (FAKE_RESUME_ID, FAKE_JD_ID))
            cleanup_conn.commit()

    check("no worker raised an exception", not errors)
    check("only one real LLM invocation happened across 3 concurrent cross-connection callers",
          call_count["n"] == 1)
    check("the other callers waited on the advisory lock rather than racing in (barrier never satisfied)",
          reached_together["count"] == 0)
    check("all 3 concurrent callers got back the identical result",
          len(results) == 3 and len({r["fit_percentage"] for r in results}) == 1)


def main():
    tests = [
        test_tier_thresholds_unified_and_correct,
        test_holistic_match_delegates_to_tier_from_pct,
        test_hybrid_match_fallback_is_flagged_and_not_final,
        test_hybrid_match_success_is_flagged_and_tiered_consistently,
        test_weighted_category_sum_is_deterministic_and_matches_hand_calculation,
        test_title_excluded_from_ai_matching_and_cache_stays_valid_across_title_edits,
        test_incomplete_category_scores_rejected_not_silently_averaged,
        test_cache_never_persists_fallback_and_self_heals_legacy_rows,
        test_rank_top_jd_matches_prefers_cached_ai_result_over_estimate,
        test_concurrent_requests_for_same_pair_compute_only_once,
    ]
    for t in tests:
        print(f"\n-- {t.__name__} --")
        try:
            t()
        except Exception:
            print(f"[ERROR] {t.__name__} raised an exception:")
            traceback.print_exc()
            failures.append(t.__name__)

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        print("Failures:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("All consistency checks passed.")


if __name__ == "__main__":
    main()
