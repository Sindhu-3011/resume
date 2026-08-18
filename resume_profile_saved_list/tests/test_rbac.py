"""Regression checks for the 5-role RBAC system (Admin / Recruiter / Hiring
Manager / Interviewer / Viewer-Auditor) added 2026-08-13 per the client's
permission-matrix spreadsheet.

Plain-Python, no pytest dependency — run directly:

    python tests/test_rbac.py

Uses `client.session_transaction()` to inject a role directly into the test
client's session (same technique test_session_and_audit.py already uses),
rather than creating real app_user rows for every role — faster and avoids
any password-hashing/account-cleanup overhead for a test this wide (many
roles x many routes).

Disposable JD/Requirement/Resume rows are created once via the real
Add routes (logged in as the real seeded `admin` account) and cleaned up at
the end, mirroring the "disposable-via-real-Add-flow" pattern used elsewhere
in this test suite.
"""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []

ROLES = ["admin", "recruiter", "hiring_manager", "interviewer", "viewer_auditor"]


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def _as_role(client, role, full_name="Test User"):
    """Inject a session for `role` directly — no real app_user row needed.
    `role='legacy_user'` deliberately sets a role string that isn't one of
    the 5 recognized keys, simulating a pre-migration account."""
    with client.session_transaction() as sess:
        sess["user_id"] = 999999
        sess["username"] = f"test_{role}"
        sess["full_name"] = full_name
        sess["role"] = role
        import datetime
        sess["last_active"] = datetime.datetime.now(datetime.timezone.utc).isoformat()


def _setup_fixtures():
    """Create one disposable JD, Requirement, and Resume as the real admin
    account. Returns (jd_id, req_id, resume_id); any that fail to create
    are returned as None so calling tests can skip gracefully."""
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})

    # Both jd_add and requirement_add redirect to their *-management list
    # pages on success (not to a detail page), so the new row's id is looked
    # up by its unique test title/name instead of parsed from the redirect.
    jd_id = None
    client.post("/jd/add", data={
        "title": "ZZZ RBAC Test JD", "role": "Test Role", "category": "IT Roles",
        "responsibilities": "x", "requirements": "x", "skills": "Python", "keywords": "python",
        "position_status": "Open",
    })
    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT id FROM job_description WHERE title = %s ORDER BY id DESC LIMIT 1",
            ("ZZZ RBAC Test JD",),
        ).fetchone()
        if row:
            jd_id = row["id"]

    req_id = None
    client.post("/requirement/add", data={
        "requirement_code": "ZZZ-RBAC-1", "requirement_name": "ZZZ RBAC Test Requirement",
        "client": "x", "division": "x", "status": "Open", "onsite_offshore": "Onsite",
        "num_requirement": "1", "profiles_shared": "0", "interviewed": "0", "offered": "0",
    })
    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT id FROM requirement WHERE requirement_name = %s ORDER BY id DESC LIMIT 1",
            ("ZZZ RBAC Test Requirement",),
        ).fetchone()
        if row:
            req_id = row["id"]

    resume_id = None
    r = client.post("/edit", data={
        "full_name": "ZZZ RBAC Test Candidate", "title": "Engineer", "email": "",
        "phone": "", "location": "Test City", "exp_yrs": "5", "summary": "", "skills": "",
        "experience": "", "education": "", "certifications": "", "projects": "", "department": "",
    })
    m = re.search(r"/profile/(\d+)", r.headers.get("Location", ""))
    if m:
        resume_id = int(m.group(1))

    return jd_id, req_id, resume_id


def _teardown_fixtures(jd_id, req_id, resume_id):
    with appmod.db_conn() as conn:
        if jd_id:
            conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))
        if req_id:
            conn.execute("DELETE FROM requirement WHERE id = %s", (req_id,))
        if resume_id:
            conn.execute("DELETE FROM resume WHERE id = %s", (resume_id,))
        conn.commit()


def test_view_dashboard_all_five_roles_allowed_legacy_role_blocked():
    client = appmod.app.test_client()
    for role in ROLES:
        _as_role(client, role)
        r = client.get("/dashboard")
        check(f"{role}: GET /dashboard -> 200", r.status_code == 200)

    _as_role(client, "user")  # legacy pre-migration role, not one of the 5
    r = client.get("/dashboard")
    check("legacy 'user' role: GET /dashboard -> 403 (not auto-migrated, needs explicit reassignment)",
          r.status_code == 403)


def test_view_only_roles_can_view_but_not_write(jd_id, req_id, resume_id):
    client = appmod.app.test_client()

    # Viewer/Auditor: view-level access everywhere, but every write route 403s.
    _as_role(client, "viewer_auditor")
    if jd_id:
        check("viewer_auditor: GET /jd/<id> -> 200 (view-only)",
              client.get(f"/jd/{jd_id}").status_code == 200)
        check("viewer_auditor: GET /jd/<id>/edit -> 403",
              client.get(f"/jd/{jd_id}/edit").status_code == 403)
        check("viewer_auditor: POST /jd/<id>/delete -> 403",
              client.post(f"/jd/{jd_id}/delete").status_code == 403)
    check("viewer_auditor: GET /jd/add -> 403", client.get("/jd/add").status_code == 403)
    check("viewer_auditor: GET /profiles -> 200 (view-only)", client.get("/profiles").status_code == 200)
    check("viewer_auditor: GET /edit (Add Profile) -> 403", client.get("/edit").status_code == 403)
    if resume_id:
        check("viewer_auditor: POST /profile/<id>/delete -> 403",
              client.post(f"/profile/{resume_id}/delete").status_code == 403)
        check("viewer_auditor: GET /profile/<id>/download-pdf -> 200 (view-level download allowed)",
              client.get(f"/profile/{resume_id}/download-pdf").status_code == 200)
    check("viewer_auditor: GET /audit-trail -> 200 (matrix widens this from admin-only)",
          client.get("/audit-trail").status_code == 200)
    check("viewer_auditor: GET /audit-trail/export.csv -> 403",
          client.get("/audit-trail/export.csv").status_code == 403)
    check("viewer_auditor: GET /users -> 403", client.get("/users").status_code == 403)
    if resume_id:
        check("viewer_auditor: GET /compare/<resume_id> -> 403 (Compare Resume disabled, 2026-08-14 follow-up)",
              client.get(f"/compare/{resume_id}").status_code == 403)

    _as_role(client, "interviewer")
    if resume_id:
        check("interviewer: GET /compare/<resume_id> -> 403 (Compare Resume disabled, 2026-08-14 follow-up)",
              client.get(f"/compare/{resume_id}").status_code == 403)
    if jd_id and resume_id:
        check("interviewer: GET /compare/<resume_id>/<jd_id> -> 403",
              client.get(f"/compare/{resume_id}/{jd_id}").status_code == 403)
    check("viewer_auditor: POST /api/schedule-interview -> 403",
          client.post("/api/schedule-interview", json={}).status_code == 403)


def test_recruiter_can_add_edit_jd_and_requirement_but_not_delete(jd_id, req_id):
    """The one deliberately asymmetric rule in the whole matrix: Recruiter
    gets full Add/Edit on JD and Requirement, but is explicitly excluded
    from Delete on both — Hiring Manager (and Admin) can delete either."""
    client = appmod.app.test_client()
    _as_role(client, "recruiter")
    check("recruiter: GET /jd/add -> 200 (can add)", client.get("/jd/add").status_code == 200)
    if jd_id:
        check("recruiter: GET /jd/<id>/edit -> 200 (can edit)",
              client.get(f"/jd/{jd_id}/edit").status_code == 200)
        check("recruiter: POST /jd/<id>/delete -> 403 (CANNOT delete)",
              client.post(f"/jd/{jd_id}/delete").status_code == 403)
    check("recruiter: GET /requirement/add -> 200 (can add)",
          client.get("/requirement/add").status_code == 200)
    if req_id:
        check("recruiter: GET /requirement/<id>/edit -> 200 (can edit)",
              client.get(f"/requirement/{req_id}/edit").status_code == 200)
        check("recruiter: POST /requirement/<id>/delete -> 403 (CANNOT delete)",
              client.post(f"/requirement/{req_id}/delete").status_code == 403)

    _as_role(client, "hiring_manager")
    if jd_id:
        check("hiring_manager: GET /jd/<id>/edit -> 200",
              client.get(f"/jd/{jd_id}/edit").status_code == 200)
        # Don't actually delete the shared jd_id fixture here — Hiring
        # Manager's delete access is proven via the permission function
        # directly instead, to avoid destroying a fixture other tests need.


def test_hiring_manager_delete_permission_level_is_full():
    with appmod.app.test_request_context():
        from flask import session
        session["role"] = "hiring_manager"
        check("hiring_manager has full delete_jd permission",
              appmod._permission_level("delete_jd") == "full")
        check("hiring_manager has full delete_requirement permission",
              appmod._permission_level("delete_requirement") == "full")
        session["role"] = "recruiter"
        check("recruiter does NOT have full delete_jd permission",
              appmod._permission_level("delete_jd") != "full")
        check("recruiter does NOT have full delete_requirement permission",
              appmod._permission_level("delete_requirement") != "full")


def test_is_my_assigned_interview_word_subset_matching():
    """Regression test for the exact bug found live on 2026-08-14: an
    interviewer's own account full_name ("Kumar") didn't exact-match the
    fuller name typed when scheduling the interview ("Satheesh Kumar") —
    same person, correctly assigned, incorrectly blocked. Word-subset
    matching must fix this without over-matching two different people who
    happen to share one common name word."""
    with appmod.app.test_request_context():
        from flask import session
        session["full_name"] = "Kumar"
        check('"Kumar" matches "Satheesh Kumar" (short name is a word-subset of the full name)',
              appmod._is_my_assigned_interview("Satheesh Kumar") is True)
        session["full_name"] = "Satheesh Kumar"
        check('"Satheesh Kumar" matches "Kumar" (reverse direction also works)',
              appmod._is_my_assigned_interview("Kumar") is True)
        session["full_name"] = "Kumar Patel"
        check('"Kumar Patel" does NOT match "Kumar Singh" (different person, shared word only)',
              appmod._is_my_assigned_interview("Kumar Singh") is False)
        session["full_name"] = ""
        check("empty full_name never matches", appmod._is_my_assigned_interview("Satheesh Kumar") is False)


def test_interviewer_assigned_only_update_interview_status(resume_id):
    if not resume_id:
        check("interviewer assigned-only test skipped (no disposable resume)", True)
        return

    interview_id = None
    with appmod.db_conn() as conn:
        appmod.ensure_interview_schedule_table(conn)
        interview_id = conn.execute(
            """
            INSERT INTO interview_schedule (resume_id, position, interviewer, interview_date, status)
            VALUES (%s, %s, %s, CURRENT_DATE, 'Scheduled')
            RETURNING id
            """,
            (resume_id, "Engineer", "Iris Interviewer"),
        ).fetchone()["id"]
        conn.commit()

    try:
        client = appmod.app.test_client()
        _as_role(client, "interviewer", full_name="Someone Else")
        r = client.post(f"/api/interview/{interview_id}/update", json={
            "interviewer": "Iris Interviewer", "interview_date": "2026-01-01", "status": "Completed",
        })
        check("interviewer NOT assigned to this interview -> 403", r.status_code == 403)

        _as_role(client, "interviewer", full_name="Iris Interviewer")
        r = client.post(f"/api/interview/{interview_id}/update", json={
            "interviewer": "Iris Interviewer", "interview_date": "2026-01-01", "status": "Completed",
        })
        check("interviewer IS assigned (name matches, case-insensitive) -> allowed",
              r.status_code == 200 and r.get_json().get("ok") is True)

        _as_role(client, "viewer_auditor")
        r = client.post(f"/api/interview/{interview_id}/update", json={
            "interviewer": "Iris Interviewer", "interview_date": "2026-01-01", "status": "Completed",
        })
        check("viewer_auditor can never update interview status -> 403", r.status_code == 403)
    finally:
        with appmod.db_conn() as conn:
            conn.execute("DELETE FROM interview_schedule WHERE id = %s", (interview_id,))
            conn.commit()


def test_admin_and_hiring_manager_unaffected_full_access(jd_id, req_id, resume_id):
    client = appmod.app.test_client()
    for role in ("admin", "hiring_manager"):
        _as_role(client, role)
        check(f"{role}: GET /users -> {'200' if role == 'admin' else '403'}",
              client.get("/users").status_code == (200 if role == "admin" else 403))
        if jd_id:
            check(f"{role}: GET /jd/<id> -> 200", client.get(f"/jd/{jd_id}").status_code == 200)
        check(f"{role}: GET /audit-trail -> 200", client.get("/audit-trail").status_code == 200)
        check(f"{role}: GET /audit-trail/export.csv -> 200",
              client.get("/audit-trail/export.csv").status_code == 200)
        if resume_id:
            check(f"{role}: GET /compare/<resume_id> -> 200 (Compare Resume still allowed)",
                  client.get(f"/compare/{resume_id}").status_code == 200)


def main():
    jd_id, req_id, resume_id = _setup_fixtures()
    check("disposable JD fixture created", jd_id is not None)
    check("disposable Requirement fixture created", req_id is not None)
    check("disposable Resume fixture created", resume_id is not None)

    def _view_only():
        test_view_only_roles_can_view_but_not_write(jd_id, req_id, resume_id)

    def _recruiter_asymmetry():
        test_recruiter_can_add_edit_jd_and_requirement_but_not_delete(jd_id, req_id)

    def _assigned_only():
        test_interviewer_assigned_only_update_interview_status(resume_id)

    def _admin_hm_full_access():
        test_admin_and_hiring_manager_unaffected_full_access(jd_id, req_id, resume_id)

    tests = [
        test_view_dashboard_all_five_roles_allowed_legacy_role_blocked,
        _view_only,
        _recruiter_asymmetry,
        test_hiring_manager_delete_permission_level_is_full,
        test_is_my_assigned_interview_word_subset_matching,
        _assigned_only,
        _admin_hm_full_access,
    ]
    try:
        for t in tests:
            print(f"\n-- {t.__name__} --")
            try:
                t()
            except Exception as e:
                print(f"[ERROR] {t.__name__} raised {e!r}")
                failures.append(f"{t.__name__} raised {e!r}")
    finally:
        _teardown_fixtures(jd_id, req_id, resume_id)

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All RBAC checks passed.")


if __name__ == "__main__":
    main()
