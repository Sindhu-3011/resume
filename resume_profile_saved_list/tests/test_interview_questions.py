"""Regression checks for the Interview Questions & Answers feature (JD
Management -> per-JD AI-generated interview questions).

Plain-Python, no pytest dependency — run directly:

    python tests/test_interview_questions.py

This feature is fully additive (its own table, its own routes, its own
template) — these tests only exercise the new surface area, never resume
parsing, JD matching, or ai_match_cache.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
FAKE_JD_ID = 999_999_201
FAKE_JD_ID_2 = 999_999_202


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def _login(client):
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})


def _insert_fake_jd(conn, jd_id, title, requirements="", skills=""):
    conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))
    conn.execute(
        """
        INSERT INTO job_description (id, title, role, category, responsibilities, requirements, skills, keywords)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """,
        (jd_id, title, "Test Role", "IT Roles", "", requirements, skills, ""),
    )


def _cleanup(conn, *jd_ids):
    for jd_id in jd_ids:
        conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))  # cascades interview_questions


def _fixed_ollama_response(*a, **k):
    return {"questions": [
        {"question": "Q1?", "expected_answer": "A1", "key_points": ["p1", "p2"], "difficulty": "Easy"},
        {"question": "Q2?", "expected_answer": "A2", "key_points": ["p1"], "difficulty": "High"},
        {"question": "Q3?", "expected_answer": "A3", "key_points": [], "difficulty": "bogus"},
    ]}


def test_coerce_interview_questions_drops_malformed_and_defaults_difficulty():
    raw = {"questions": [
        {"question": "Real question?", "expected_answer": "ans", "key_points": ["a", "b"], "difficulty": "High"},
        {"question": "", "expected_answer": "should be dropped, no question text"},
        {"question": "No difficulty given?"},
        "not even a dict",
    ]}
    out = appmod._coerce_interview_questions(raw, "Technical")
    check("empty-question entry is dropped", len(out) == 2)
    check("valid question keeps its own difficulty", out[0]["difficulty"] == "High")
    check("missing difficulty defaults to Medium", out[1]["difficulty"] == "Medium")
    check("category is tagged from the trusted caller-supplied value", all(q["category"] == "Technical" for q in out))
    check("key_points coerced to newline-joined string", out[0]["key_points"] == "a\nb")

    check("non-dict raw input returns empty list", appmod._coerce_interview_questions("not a dict", "Technical") == [])
    check("missing 'questions' key returns empty list", appmod._coerce_interview_questions({}, "Technical") == [])


def test_generate_creates_questions_grouped_by_category():
    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = _fixed_ollama_response
        with appmod.db_conn() as conn:
            _insert_fake_jd(conn, FAKE_JD_ID, "Design Control Consultant",
                             requirements="FDA 21 CFR Part 820, ISO 13485, EU IVDR, DHF, V&V, Risk Management")
            conn.commit()

        client = appmod.app.test_client()
        _login(client)
        r = client.post(f"/jd/{FAKE_JD_ID}/interview-questions/generate")
        data = r.get_json()
        check("generate endpoint returns success", data.get("success") is True)
        check("generate returns questions grouped by all 4 categories",
              set(data.get("questions", {}).keys()) == set(appmod.INTERVIEW_QUESTION_CATEGORIES))

        with appmod.db_conn() as conn:
            rows = conn.execute("SELECT * FROM interview_questions WHERE jd_id = %s", (FAKE_JD_ID,)).fetchall()
        # 3 coerced questions (the malformed-difficulty one still keeps its
        # question text, only its difficulty gets defaulted) x 5 batches
        # (Technical is split into 2 batches, see _INTERVIEW_QUESTION_BATCHES).
        check("15 rows stored (3 per batch x 5 batches)", len(rows) == 15)
        categories_present = {row["category"] for row in rows}
        check("all 4 categories represented in storage", categories_present == set(appmod.INTERVIEW_QUESTION_CATEGORIES))
        check("bogus difficulty was defaulted to Medium in storage",
              any(row["difficulty"] == "Medium" for row in rows))
    finally:
        appmod._ollama_chat = real_ollama_chat
        with appmod.db_conn() as conn:
            _cleanup(conn, FAKE_JD_ID)
            conn.commit()


def test_regenerate_replaces_rather_than_appends():
    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = _fixed_ollama_response
        with appmod.db_conn() as conn:
            _insert_fake_jd(conn, FAKE_JD_ID, "Validation Engineer")
            conn.commit()

        client = appmod.app.test_client()
        _login(client)
        client.post(f"/jd/{FAKE_JD_ID}/interview-questions/generate")
        client.post(f"/jd/{FAKE_JD_ID}/interview-questions/generate")

        with appmod.db_conn() as conn:
            rows = conn.execute("SELECT * FROM interview_questions WHERE jd_id = %s", (FAKE_JD_ID,)).fetchall()
        check("regenerating does not duplicate the question set (still 15, not 30)", len(rows) == 15)
    finally:
        appmod._ollama_chat = real_ollama_chat
        with appmod.db_conn() as conn:
            _cleanup(conn, FAKE_JD_ID)
            conn.commit()


def test_questions_are_scoped_per_jd_and_never_cross_over():
    """The user's own stated test case: questions generated for one JD must
    never appear when viewing a different JD."""
    real_ollama_chat = appmod._ollama_chat
    try:
        appmod._ollama_chat = _fixed_ollama_response
        with appmod.db_conn() as conn:
            _insert_fake_jd(conn, FAKE_JD_ID, "Design Control Consultant")
            _insert_fake_jd(conn, FAKE_JD_ID_2, "Validation Engineer")
            conn.commit()

        client = appmod.app.test_client()
        _login(client)
        client.post(f"/jd/{FAKE_JD_ID}/interview-questions/generate")
        client.post(f"/jd/{FAKE_JD_ID_2}/interview-questions/generate")

        with appmod.db_conn() as conn:
            jd1_ids = {row["id"] for row in conn.execute(
                "SELECT id FROM interview_questions WHERE jd_id = %s", (FAKE_JD_ID,)).fetchall()}
            jd2_rows = conn.execute(
                "SELECT id FROM interview_questions WHERE jd_id = %s", (FAKE_JD_ID_2,)).fetchall()
        check("JD 1 has its own questions", len(jd1_ids) == 15)
        check("no row from JD 1 appears under JD 2", all(row["id"] not in jd1_ids for row in jd2_rows))
    finally:
        appmod._ollama_chat = real_ollama_chat
        with appmod.db_conn() as conn:
            _cleanup(conn, FAKE_JD_ID, FAKE_JD_ID_2)
            conn.commit()


def test_generate_fails_safely_when_ollama_is_completely_down():
    real_ollama_chat = appmod._ollama_chat
    try:
        def _boom(*a, **k):
            raise RuntimeError("Ollama unavailable")
        appmod._ollama_chat = _boom

        with appmod.db_conn() as conn:
            _insert_fake_jd(conn, FAKE_JD_ID, "Design Control Consultant")
            conn.commit()

        client = appmod.app.test_client()
        _login(client)
        r = client.post(f"/jd/{FAKE_JD_ID}/interview-questions/generate")
        data = r.get_json()
        check("generate endpoint reports failure, not a 500 crash", data.get("success") is False)
        check("failure message matches the spec's required user-facing text",
              data.get("message") == "Unable to generate interview questions. Please try again.")
        check("response status is a non-2xx error code", r.status_code >= 400)

        with appmod.db_conn() as conn:
            rows = conn.execute("SELECT * FROM interview_questions WHERE jd_id = %s", (FAKE_JD_ID,)).fetchall()
        check("no rows were inserted on total failure", len(rows) == 0)

        # The JD page itself must still render normally, not error out.
        r2 = client.get(f"/jd/{FAKE_JD_ID}/interview-questions")
        check("interview-questions page still renders when Ollama is down", r2.status_code == 200)
    finally:
        appmod._ollama_chat = real_ollama_chat
        with appmod.db_conn() as conn:
            _cleanup(conn, FAKE_JD_ID)
            conn.commit()


def test_add_edit_delete_single_question_round_trip():
    with appmod.db_conn() as conn:
        _insert_fake_jd(conn, FAKE_JD_ID, "Design Control Consultant")
        conn.commit()

    client = appmod.app.test_client()
    _login(client)
    try:
        r = client.post(f"/jd/{FAKE_JD_ID}/interview-questions/add", json={
            "category": "technical", "question": "Manually added question?",
            "expected_answer": "Manual answer.", "key_points": ["k1", "k2"], "difficulty": "easy",
        })
        data = r.get_json()
        check("add returns ok", data.get("ok") is True)
        q_id = data["question"]["id"]
        check("category is normalized to title case", data["question"]["category"] == "Technical")
        check("difficulty is normalized to title case", data["question"]["difficulty"] == "Easy")

        r = client.post(f"/api/interview-question/{q_id}/update", json={
            "question": "Updated question?", "expected_answer": "Updated answer.",
            "key_points": "single point", "difficulty": "High",
        })
        check("update returns ok", r.get_json().get("ok") is True)
        with appmod.db_conn() as conn:
            row = conn.execute("SELECT * FROM interview_questions WHERE id = %s", (q_id,)).fetchone()
        check("question text updated in DB", row["question"] == "Updated question?")
        check("difficulty updated in DB", row["difficulty"] == "High")

        r = client.post(f"/api/interview-question/{q_id}/delete")
        check("delete returns ok", r.get_json().get("ok") is True)
        with appmod.db_conn() as conn:
            row = conn.execute("SELECT * FROM interview_questions WHERE id = %s", (q_id,)).fetchone()
        check("question is gone from DB after delete", row is None)
    finally:
        with appmod.db_conn() as conn:
            _cleanup(conn, FAKE_JD_ID)
            conn.commit()


def test_add_question_rejects_invalid_category_and_empty_question():
    with appmod.db_conn() as conn:
        _insert_fake_jd(conn, FAKE_JD_ID, "Design Control Consultant")
        conn.commit()

    client = appmod.app.test_client()
    _login(client)
    try:
        r = client.post(f"/jd/{FAKE_JD_ID}/interview-questions/add", json={
            "category": "NotARealCategory", "question": "Some question?",
        })
        check("invalid category is rejected with 400", r.status_code == 400)

        r = client.post(f"/jd/{FAKE_JD_ID}/interview-questions/add", json={
            "category": "Technical", "question": "",
        })
        check("empty question is rejected with 400", r.status_code == 400)
    finally:
        with appmod.db_conn() as conn:
            _cleanup(conn, FAKE_JD_ID)
            conn.commit()


def main():
    tests = [
        test_coerce_interview_questions_drops_malformed_and_defaults_difficulty,
        test_generate_creates_questions_grouped_by_category,
        test_regenerate_replaces_rather_than_appends,
        test_questions_are_scoped_per_jd_and_never_cross_over,
        test_generate_fails_safely_when_ollama_is_completely_down,
        test_add_edit_delete_single_question_round_trip,
        test_add_question_rejects_invalid_category_and_empty_question,
    ]
    for t in tests:
        print(f"\n-- {t.__name__} --")
        try:
            t()
        except Exception as e:
            print(f"[ERROR] {t.__name__} raised {e!r}")
            failures.append(f"{t.__name__} raised {e!r}")

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All interview-questions checks passed.")


if __name__ == "__main__":
    main()
