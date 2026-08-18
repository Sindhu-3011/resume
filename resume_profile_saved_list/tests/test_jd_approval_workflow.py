"""Regression checks for the JD approval workflow, version history, and
JD<->Requirement linking feature (JD Management / Requirement Management).

Plain-Python, no pytest dependency — run directly:

    python tests/test_jd_approval_workflow.py

Session state (user id/role) is injected directly via session_transaction()
rather than a real username/password login — the same technique
test_account_lockout.py's admin-unlock test uses — so this suite never
depends on the real seeded admin account's actual current password, which
has repeatedly drifted independently of the test suite in this app. Real
disposable app_user rows are still created for any user whose id needs to
satisfy a foreign key (job_description.created_by/decided_by).

Uses disposable job_description/app_user/requirement rows (ZZZ-prefixed /
999_999_4xx ids), cleaned up in a finally block. No real SMTP server exists
in this environment — every scenario monkeypatches _send_email to capture
what would have been sent rather than actually sending it.
"""

import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
FAKE_JD_ID = 999_999_401
SUBMITTER_ID = 999_999_402      # role: recruiter — can create/submit, cannot approve
APPROVER_ID = 999_999_403       # role: hiring_manager — can approve
SECOND_APPROVER_ID = 999_999_404  # role: hiring_manager — a second, independent approver
FAKE_REQ_ID = 999_999_405
_extra_jd_ids = []  # ids created via the real jd_add route, swept up on cleanup


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def _session_as(client, user_id, username, full_name, role):
    with client.session_transaction() as sess:
        sess["user_id"] = user_id
        sess["username"] = username
        sess["full_name"] = full_name
        sess["role"] = role
        sess["last_active"] = datetime.datetime.now(datetime.timezone.utc).isoformat()


def _create_test_users():
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE id IN (%s, %s, %s)",
                     (SUBMITTER_ID, APPROVER_ID, SECOND_APPROVER_ID))
        conn.commit()
        for uid, uname, full_name, role in [
            (SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter"),
            (APPROVER_ID, "zzz_jdwf_approver", "ZZZ JD Approver", "hiring_manager"),
            (SECOND_APPROVER_ID, "zzz_jdwf_approver2", "ZZZ JD Approver Two", "hiring_manager"),
        ]:
            conn.execute(
                "INSERT INTO app_user (id, username, password_hash, full_name, email, role, is_active) "
                "VALUES (%s, %s, %s, %s, %s, %s, TRUE)",
                (uid, uname, appmod.generate_password_hash("Whatever1!"), full_name,
                 f"{uname}@example.com", role),
            )
        conn.commit()


def _insert_fake_jd(jd_id=FAKE_JD_ID, workflow_status="Draft", created_by=None, title=None):
    fields = {
        "id": jd_id, "title": title or f"ZZZ Test JD {jd_id}", "role": "Test Role",
        "category": "IT Roles", "responsibilities": "Do the thing.", "requirements": "",
        "skills": "", "keywords": "", "position_status": "Open",
        "workflow_status": workflow_status, "created_by": created_by,
    }
    with appmod.db_conn() as conn:
        appmod.ensure_jd_table(conn)
        conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))
        conn.commit()
        conn.execute(
            """
            INSERT INTO job_description
                (id, title, role, category, responsibilities, requirements, skills, keywords,
                 position_status, workflow_status, created_by)
            VALUES (%(id)s, %(title)s, %(role)s, %(category)s, %(responsibilities)s, %(requirements)s,
                    %(skills)s, %(keywords)s, %(position_status)s, %(workflow_status)s, %(created_by)s)
            """,
            fields,
        )
        conn.commit()


def _fetch_jd(jd_id=FAKE_JD_ID):
    with appmod.db_conn() as conn:
        return conn.execute("SELECT * FROM job_description WHERE id = %s", (jd_id,)).fetchone()


def _capture_sent_emails():
    sent = []
    real_send = appmod._send_email

    def _fake_send(to_addr, subject, body):
        sent.append({"to": to_addr, "subject": subject, "body": body})

    appmod._send_email = _fake_send

    def _restore():
        appmod._send_email = real_send

    return sent, _restore


def test_jd_add_creates_draft_owned_by_submitter():
    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")
    r = client.post("/jd/add", data={
        "title": "ZZZ JD Workflow Add Test", "role": "Test Role", "category": "IT Roles",
        "responsibilities": "Do the thing.", "requirements": "3+ years.", "skills": "Python",
        "keywords": "python, testing", "position_status": "Open",
    })
    check("jd_add redirects on success", r.status_code == 302)
    m = re.search(r"/jd/(\d+)", r.headers.get("Location", ""))
    check("redirect location contains the new JD id", m is not None)
    if m:
        new_id = int(m.group(1))
        _extra_jd_ids.append(new_id)
        row = _fetch_jd(new_id)
        check("new JD defaults to Draft", row["workflow_status"] == "Draft")
        check("new JD is owned by the submitting user", row["created_by"] == SUBMITTER_ID)


def test_jd_add_warns_on_duplicate_title_and_category_then_allows_override():
    _insert_fake_jd(workflow_status="Published", title="ZZZ Duplicate Check JD")
    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")

    base_payload = {
        "title": "ZZZ Duplicate Check JD", "role": "Test Role", "category": "IT Roles",
        "responsibilities": "Do the thing.", "requirements": "3+ years.", "skills": "Python",
        "keywords": "python", "position_status": "Open",
    }

    r = client.post("/jd/add", data=base_payload)
    check("same title+category is blocked with a warning, not created", r.status_code == 200)
    check("warning banner is shown", b"Possible duplicate JD" in r.data)
    with appmod.db_conn() as conn:
        count = conn.execute(
            "SELECT COUNT(*) AS c FROM job_description WHERE title = %s", ("ZZZ Duplicate Check JD",)
        ).fetchone()["c"]
    check("no second row was inserted for the blocked attempt", count == 1)

    different_category_payload = dict(base_payload, category="Validation Roles")
    r = client.post("/jd/add", data=different_category_payload)
    check("same title but a different category is NOT flagged as a duplicate", r.status_code == 302)
    m = re.search(r"/jd/(\d+)", r.headers.get("Location", ""))
    if m:
        _extra_jd_ids.append(int(m.group(1)))

    override_payload = dict(base_payload, confirm_duplicate="1")
    r = client.post("/jd/add", data=override_payload)
    check("'Add Anyway' (confirm_duplicate=1) bypasses the warning and creates it", r.status_code == 302)
    m = re.search(r"/jd/(\d+)", r.headers.get("Location", ""))
    check("redirect location contains the new JD id", m is not None)
    if m:
        _extra_jd_ids.append(int(m.group(1)))
    with appmod.db_conn() as conn:
        count = conn.execute(
            "SELECT COUNT(*) AS c FROM job_description WHERE title = %s", ("ZZZ Duplicate Check JD",)
        ).fetchone()["c"]
    check("now 3 rows share this title (seed + different-category + confirmed-duplicate)", count == 3)


def test_pdf_export_blocked_for_draft_and_rejected():
    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")

    for status in ("Draft", "Rejected"):
        _insert_fake_jd(workflow_status=status)
        r = client.get(f"/jd/{FAKE_JD_ID}/download-pdf")
        check(f"{status} JD's PDF download is blocked (redirected, not a PDF)",
              r.status_code == 302 and "/jd/" in r.headers.get("Location", ""))

    for status in ("Approved", "Published"):
        _insert_fake_jd(workflow_status=status)
        r = client.get(f"/jd/{FAKE_JD_ID}/download-pdf")
        check(f"{status} JD's PDF download still works", r.status_code == 200
              and r.mimetype == "application/pdf")


def test_jd_management_search_finds_by_keyword_and_id():
    _insert_fake_jd(workflow_status="Published", title="ZZZ Searchable Widget Engineer")
    with appmod.db_conn() as conn:
        conn.execute(
            "UPDATE job_description SET keywords = %s, skills = %s WHERE id = %s",
            ("zzzuniquekeywordxyz", "Widgetology", FAKE_JD_ID),
        )
        conn.commit()

    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")

    r = client.get("/jd-management?q=zzzuniquekeywordxyz")
    check("search by a keyword only present in the keywords field finds the JD",
          b"ZZZ Searchable Widget Engineer" in r.data)

    r = client.get(f"/jd-management?q={FAKE_JD_ID}")
    check("search by numeric JD id finds the JD", b"ZZZ Searchable Widget Engineer" in r.data)

    r = client.get("/jd-management?q=zzz_definitely_not_a_real_keyword")
    check("a non-matching search excludes the JD", b"ZZZ Searchable Widget Engineer" not in r.data)
    check("a non-matching search shows the no-results message", b"No Job Descriptions match your search" in r.data)


def test_submit_for_approval_transitions_and_notifies_approvers():
    _insert_fake_jd(workflow_status="Draft", created_by=SUBMITTER_ID)
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")
        r = client.post(f"/jd/{FAKE_JD_ID}/submit-for-approval")
        check("submit-for-approval redirects", r.status_code == 302)
        row = _fetch_jd()
        check("status moves to Pending Approval", row["workflow_status"] == "Pending Approval")
        check("submitted_at is stamped", row["submitted_at"] is not None)
        recipients = {e["to"] for e in sent}
        check("both hiring-manager approvers were emailed",
              "zzz_jdwf_approver@example.com" in recipients and
              "zzz_jdwf_approver2@example.com" in recipients)
    finally:
        restore()


def test_submit_for_approval_rejected_from_wrong_state():
    _insert_fake_jd(workflow_status="Published", created_by=SUBMITTER_ID)
    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")
    client.post(f"/jd/{FAKE_JD_ID}/submit-for-approval")
    row = _fetch_jd()
    check("a Published JD cannot be re-submitted for approval", row["workflow_status"] == "Published")


def test_self_approval_blocked_but_other_approver_succeeds():
    # APPROVER_ID both created this JD and is a valid approver-role user —
    # the block must trigger on "you created this", not on lacking the role.
    _insert_fake_jd(workflow_status="Pending Approval", created_by=APPROVER_ID)
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        _session_as(client, APPROVER_ID, "zzz_jdwf_approver", "ZZZ JD Approver", "hiring_manager")
        r = client.post(f"/jd/{FAKE_JD_ID}/approve")
        check("self-approval attempt does not error, just redirects with a flash", r.status_code == 302)
        row = _fetch_jd()
        check("status is unchanged after a blocked self-approval", row["workflow_status"] == "Pending Approval")

        client2 = appmod.app.test_client()
        _session_as(client2, SECOND_APPROVER_ID, "zzz_jdwf_approver2", "ZZZ JD Approver Two", "hiring_manager")
        r2 = client2.post(f"/jd/{FAKE_JD_ID}/approve", data={"comment": "Looks good"})
        check("a different approver can approve the same JD", r2.status_code == 302)
        row = _fetch_jd()
        check("status moves to Approved", row["workflow_status"] == "Approved")
        check("decided_by is the second approver", row["decided_by"] == SECOND_APPROVER_ID)
        check("the original creator (also the blocked approver) was emailed the decision",
              any(e["to"] == "zzz_jdwf_approver@example.com" for e in sent))
    finally:
        restore()


def test_reject_requires_comment_and_blocks_publish():
    _insert_fake_jd(workflow_status="Pending Approval", created_by=SUBMITTER_ID)
    client = appmod.app.test_client()
    _session_as(client, SECOND_APPROVER_ID, "zzz_jdwf_approver2", "ZZZ JD Approver Two", "hiring_manager")
    r = client.post(f"/jd/{FAKE_JD_ID}/reject")
    check("reject without a comment is rejected with 400", r.status_code == 400)
    row = _fetch_jd()
    check("status unchanged after the rejected (no-comment) attempt", row["workflow_status"] == "Pending Approval")

    r = client.post(f"/jd/{FAKE_JD_ID}/reject", data={"comment": "Missing required skills section"})
    check("reject with a comment succeeds", r.status_code == 302)
    row = _fetch_jd()
    check("status moves to Rejected", row["workflow_status"] == "Rejected")

    r = client.post(f"/jd/{FAKE_JD_ID}/publish")
    check("publish redirects rather than erroring on a Rejected JD", r.status_code == 302)
    row = _fetch_jd()
    check("a Rejected JD can never be published", row["workflow_status"] == "Rejected")


def test_request_changes_then_resubmit():
    _insert_fake_jd(workflow_status="Pending Approval", created_by=SUBMITTER_ID)
    approver_client = appmod.app.test_client()
    _session_as(approver_client, SECOND_APPROVER_ID, "zzz_jdwf_approver2", "ZZZ JD Approver Two", "hiring_manager")
    approver_client.post(f"/jd/{FAKE_JD_ID}/request-changes", data={"comment": "Please add more detail"})
    row = _fetch_jd()
    check("status moves to Changes Requested", row["workflow_status"] == "Changes Requested")

    submitter_client = appmod.app.test_client()
    _session_as(submitter_client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")
    submitter_client.post(f"/jd/{FAKE_JD_ID}/submit-for-approval")
    row = _fetch_jd()
    check("a Changes-Requested JD can be resubmitted", row["workflow_status"] == "Pending Approval")


def test_editing_published_jd_versions_but_draft_edit_does_not():
    with appmod.db_conn() as conn:
        appmod.ensure_jd_version_table(conn)
        conn.execute("DELETE FROM jd_version WHERE jd_id = %s", (FAKE_JD_ID,))
        conn.commit()

    _insert_fake_jd(workflow_status="Published", title="Original Title")
    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")
    client.post(f"/jd/{FAKE_JD_ID}/edit", data={
        "title": "Edited Title", "role": "Test Role", "category": "IT Roles",
        "responsibilities": "Do the thing.", "requirements": "3+ years.", "skills": "Python",
        "keywords": "python", "position_status": "Open",
    })
    with appmod.db_conn() as conn:
        versions = conn.execute(
            "SELECT * FROM jd_version WHERE jd_id = %s ORDER BY version_number", (FAKE_JD_ID,)
        ).fetchall()
    check("editing a Published JD creates exactly one version row", len(versions) == 1)
    if versions:
        check("the version snapshot holds the PRE-edit title", versions[0]["title"] == "Original Title")

    _insert_fake_jd(workflow_status="Draft", title="Draft Title")
    client.post(f"/jd/{FAKE_JD_ID}/edit", data={
        "title": "Edited Draft Title", "role": "Test Role", "category": "IT Roles",
        "responsibilities": "Do the thing.", "requirements": "3+ years.", "skills": "Python",
        "keywords": "python", "position_status": "Open",
    })
    with appmod.db_conn() as conn:
        versions = conn.execute("SELECT * FROM jd_version WHERE jd_id = %s", (FAKE_JD_ID,)).fetchall()
    check("editing a Draft JD creates no version row", len(versions) == 0)


def test_version_compare_and_rollback():
    with appmod.db_conn() as conn:
        appmod.ensure_jd_version_table(conn)
        conn.execute("DELETE FROM jd_version WHERE jd_id = %s", (FAKE_JD_ID,))
        conn.commit()

    _insert_fake_jd(workflow_status="Published", title="Version One Title")
    editor_client = appmod.app.test_client()
    _session_as(editor_client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")
    editor_client.post(f"/jd/{FAKE_JD_ID}/edit", data={
        "title": "Version Two Title", "role": "Test Role", "category": "IT Roles",
        "responsibilities": "Do the thing.", "requirements": "3+ years.", "skills": "Python",
        "keywords": "python", "position_status": "Open",
    })
    with appmod.db_conn() as conn:
        v1 = conn.execute(
            "SELECT * FROM jd_version WHERE jd_id = %s AND version_number = 1", (FAKE_JD_ID,)
        ).fetchone()
    check("version 1 was recorded", v1 is not None)

    r = editor_client.get(f"/jd/{FAKE_JD_ID}/versions/compare?a={v1['id']}&b=current")
    check("compare page renders", r.status_code == 200)
    check("compare page shows the older title", b"Version One Title" in r.data)
    check("compare page shows the current title", b"Version Two Title" in r.data)

    approver_client = appmod.app.test_client()
    _session_as(approver_client, SECOND_APPROVER_ID, "zzz_jdwf_approver2", "ZZZ JD Approver Two", "hiring_manager")
    r = approver_client.post(f"/jd/{FAKE_JD_ID}/versions/{v1['id']}/rollback")
    check("rollback redirects", r.status_code == 302)
    row = _fetch_jd()
    check("rollback restores the older title onto the live row", row["title"] == "Version One Title")
    with appmod.db_conn() as conn:
        versions = conn.execute("SELECT * FROM jd_version WHERE jd_id = %s", (FAKE_JD_ID,)).fetchall()
    check("rollback adds a new version rather than shrinking history", len(versions) == 2)


def test_requirement_linking_validates_workflow_status_and_delete_unlinks():
    with appmod.db_conn() as conn:
        appmod.ensure_requirement_table(conn)
        conn.execute("DELETE FROM requirement WHERE id = %s", (FAKE_REQ_ID,))
        conn.commit()
        conn.execute(
            "INSERT INTO requirement (id, requirement_code, requirement_name, status) "
            "VALUES (%s, 'ZZZ-REQ-1', 'ZZZ Test Requirement', 'Open')",
            (FAKE_REQ_ID,),
        )
        conn.commit()

    client = appmod.app.test_client()
    _session_as(client, SUBMITTER_ID, "zzz_jdwf_submitter", "ZZZ JD Submitter", "recruiter")

    _insert_fake_jd(workflow_status="Draft")
    r = client.post(f"/requirement/{FAKE_REQ_ID}/edit", data={
        "requirement_code": "ZZZ-REQ-1", "requirement_name": "ZZZ Test Requirement",
        "client": "", "division": "", "status": "Open", "onsite_offshore": "",
        "num_requirement": "1", "profiles_shared": "0", "interviewed": "0", "offered": "0",
        "jd_id": str(FAKE_JD_ID),
    })
    check("linking a Draft JD to a requirement is rejected", r.status_code == 400)
    with appmod.db_conn() as conn:
        req = conn.execute("SELECT jd_id FROM requirement WHERE id = %s", (FAKE_REQ_ID,)).fetchone()
    check("requirement's jd_id is still unset after the rejected link attempt", req["jd_id"] is None)

    _insert_fake_jd(workflow_status="Published")
    r = client.post(f"/requirement/{FAKE_REQ_ID}/edit", data={
        "requirement_code": "ZZZ-REQ-1", "requirement_name": "ZZZ Test Requirement",
        "client": "", "division": "", "status": "Open", "onsite_offshore": "",
        "num_requirement": "1", "profiles_shared": "0", "interviewed": "0", "offered": "0",
        "jd_id": str(FAKE_JD_ID),
    })
    check("linking a Published JD to a requirement succeeds", r.status_code == 302)
    with appmod.db_conn() as conn:
        req = conn.execute("SELECT jd_id FROM requirement WHERE id = %s", (FAKE_REQ_ID,)).fetchone()
    check("requirement is now linked to the JD", req["jd_id"] == FAKE_JD_ID)

    r = client.get(f"/jd/{FAKE_JD_ID}")
    check("the JD detail page lists the linked requirement", b"ZZZ-REQ-1" in r.data)

    # recruiter (SUBMITTER_ID) has no delete_jd permission — deleting requires
    # admin/hiring_manager, matching ROLE_PERMISSIONS["delete_jd"].
    approver_client = appmod.app.test_client()
    _session_as(approver_client, SECOND_APPROVER_ID, "zzz_jdwf_approver2", "ZZZ JD Approver Two", "hiring_manager")
    r = approver_client.post(f"/jd/{FAKE_JD_ID}/delete")
    check("hiring_manager can delete the JD", r.status_code == 302)
    with appmod.db_conn() as conn:
        req = conn.execute("SELECT jd_id FROM requirement WHERE id = %s", (FAKE_REQ_ID,)).fetchone()
    check("deleting the JD unlinks the requirement rather than deleting it", req is not None and req["jd_id"] is None)


def main():
    _create_test_users()
    tests = [
        test_jd_add_creates_draft_owned_by_submitter,
        test_jd_add_warns_on_duplicate_title_and_category_then_allows_override,
        test_pdf_export_blocked_for_draft_and_rejected,
        test_jd_management_search_finds_by_keyword_and_id,
        test_submit_for_approval_transitions_and_notifies_approvers,
        test_submit_for_approval_rejected_from_wrong_state,
        test_self_approval_blocked_but_other_approver_succeeds,
        test_reject_requires_comment_and_blocks_publish,
        test_request_changes_then_resubmit,
        test_editing_published_jd_versions_but_draft_edit_does_not,
        test_version_compare_and_rollback,
        test_requirement_linking_validates_workflow_status_and_delete_unlinks,
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
        with appmod.db_conn() as conn:
            for jd_id in [FAKE_JD_ID] + _extra_jd_ids:
                conn.execute("DELETE FROM job_description WHERE id = %s", (jd_id,))
            conn.execute("DELETE FROM requirement WHERE id = %s", (FAKE_REQ_ID,))
            conn.execute("DELETE FROM app_user WHERE id IN (%s, %s, %s)",
                         (SUBMITTER_ID, APPROVER_ID, SECOND_APPROVER_ID))
            conn.commit()

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All JD approval-workflow checks passed.")


if __name__ == "__main__":
    main()
