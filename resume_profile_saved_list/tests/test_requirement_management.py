"""Regression checks for Requirement Management's system-generated ID,
Recruiter/Priority fields, and search.

Plain-Python, no pytest dependency — run directly:

    python tests/test_requirement_management.py

Session state (user id/role) is injected directly via session_transaction()
— same technique test_jd_approval_workflow.py uses — so this suite never
depends on the real seeded admin account's actual current password.
"""

import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
RECRUITER_ID = 999_999_501   # role: recruiter — can create/edit requirements
FAKE_REQ_ID = 999_999_502    # a directly-inserted disposable row (legacy-style, manual code)
_extra_req_ids = []          # ids created via the real requirement_add route, swept up on cleanup


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


def _create_test_user():
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE id = %s", (RECRUITER_ID,))
        conn.commit()
        conn.execute(
            "INSERT INTO app_user (id, username, password_hash, full_name, email, role, is_active) "
            "VALUES (%s, %s, %s, %s, %s, 'recruiter', TRUE)",
            (RECRUITER_ID, "zzz_reqmgmt_recruiter", appmod.generate_password_hash("Whatever1!"),
             "ZZZ Req Mgmt Recruiter", "zzz.reqmgmt.recruiter@example.com"),
        )
        conn.commit()


def _insert_legacy_requirement():
    """A pre-existing-style row with a manually-typed code, inserted directly
    (bypassing the route) — proves the new system-generation logic never
    touches already-existing rows."""
    with appmod.db_conn() as conn:
        appmod.ensure_requirement_table(conn)
        conn.execute("DELETE FROM requirement WHERE id = %s", (FAKE_REQ_ID,))
        conn.commit()
        conn.execute(
            "INSERT INTO requirement (id, requirement_code, requirement_name, client, division, status) "
            "VALUES (%s, 'REQ-1001', 'ZZZ Legacy Requirement', 'Acme Corp', 'IT', 'Open')",
            (FAKE_REQ_ID,),
        )
        conn.commit()


def test_new_requirement_gets_system_generated_id():
    client = appmod.app.test_client()
    _session_as(client, RECRUITER_ID, "zzz_reqmgmt_recruiter", "ZZZ Req Mgmt Recruiter", "recruiter")

    r = client.post("/requirement/add", data={
        "requirement_name": "ZZZ Auto ID Test", "client": "Acme", "division": "IT",
        "status": "Open", "onsite_offshore": "Onsite", "priority": "High",
        "num_requirement": "2", "profiles_shared": "0", "interviewed": "0", "offered": "0",
    })
    check("requirement_add redirects on success", r.status_code == 302)

    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT * FROM requirement WHERE requirement_name = %s", ("ZZZ Auto ID Test",)
        ).fetchone()
    check("a row was created", row is not None)
    if row:
        _extra_req_ids.append(row["id"])
        check("requirement_code is not blank", bool(row["requirement_code"]))
        check("requirement_code matches the system-generated REQ-##### format based on its own id",
              row["requirement_code"] == f"REQ-{row['id']:05d}")
        check("priority was persisted", row["priority"] == "High")
        check("assigned_recruiter_id is unset (none was chosen)", row["assigned_recruiter_id"] is None)


def test_add_form_has_no_editable_requirement_code_field():
    client = appmod.app.test_client()
    _session_as(client, RECRUITER_ID, "zzz_reqmgmt_recruiter", "ZZZ Req Mgmt Recruiter", "recruiter")
    r = client.get("/requirement/add")
    check("Add form has no editable requirement_code input",
          b'name="requirement_code"' not in r.data)


def test_recruiter_and_priority_persist_through_add_and_edit():
    client = appmod.app.test_client()
    _session_as(client, RECRUITER_ID, "zzz_reqmgmt_recruiter", "ZZZ Req Mgmt Recruiter", "recruiter")

    r = client.post("/requirement/add", data={
        "requirement_name": "ZZZ Recruiter Priority Test", "client": "Acme", "division": "Finance",
        "status": "Open", "onsite_offshore": "Offshore", "priority": "Critical",
        "assigned_recruiter_id": str(RECRUITER_ID),
        "num_requirement": "1", "profiles_shared": "0", "interviewed": "0", "offered": "0",
    })
    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT * FROM requirement WHERE requirement_name = %s", ("ZZZ Recruiter Priority Test",)
        ).fetchone()
    check("recruiter and priority were saved on create",
          row["assigned_recruiter_id"] == RECRUITER_ID and row["priority"] == "Critical")
    _extra_req_ids.append(row["id"])
    original_code = row["requirement_code"]

    r = client.post(f"/requirement/{row['id']}/edit", data={
        "requirement_name": "ZZZ Recruiter Priority Test", "client": "Acme", "division": "Finance",
        "status": "Open", "onsite_offshore": "Offshore", "priority": "Low",
        "assigned_recruiter_id": "",
        "num_requirement": "1", "profiles_shared": "0", "interviewed": "0", "offered": "0",
    })
    check("edit redirects on success", r.status_code == 302)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT * FROM requirement WHERE id = %s", (row["id"],)).fetchone()
    check("priority updated to the new value", row["priority"] == "Low")
    check("recruiter cleared back to unassigned", row["assigned_recruiter_id"] is None)
    check("requirement_code is unchanged by the edit (immutable once assigned)",
          row["requirement_code"] == original_code)


def test_search_finds_by_code_name_client_and_id():
    _insert_legacy_requirement()
    client = appmod.app.test_client()
    _session_as(client, RECRUITER_ID, "zzz_reqmgmt_recruiter", "ZZZ Req Mgmt Recruiter", "recruiter")

    r = client.get("/requirement-management?q=REQ-1001")
    check("search by requirement code finds it", b"ZZZ Legacy Requirement" in r.data)

    r = client.get("/requirement-management?q=Acme Corp")
    check("search by client name finds it", b"ZZZ Legacy Requirement" in r.data)

    r = client.get(f"/requirement-management?q={FAKE_REQ_ID}")
    check("search by numeric id finds it", b"ZZZ Legacy Requirement" in r.data)

    r = client.get("/requirement-management?q=zzz_definitely_not_a_real_match")
    check("a non-matching search excludes it", b"ZZZ Legacy Requirement" not in r.data)
    check("a non-matching search shows the no-results message",
          b"No Requirements match your search" in r.data)


def test_migration_never_touches_existing_manually_coded_rows():
    _insert_legacy_requirement()
    # Re-running the additive migrations (as happens on every app start) must
    # never alter a pre-existing row's already-set requirement_code.
    appmod._ensure_requirement_recruiter_priority_cols()
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT requirement_code, priority FROM requirement WHERE id = %s", (FAKE_REQ_ID,)).fetchone()
    check("legacy manually-typed code is untouched", row["requirement_code"] == "REQ-1001")
    check("priority backfilled to the column default for a pre-existing row", row["priority"] == "Medium")


def main():
    _create_test_user()
    tests = [
        test_new_requirement_gets_system_generated_id,
        test_add_form_has_no_editable_requirement_code_field,
        test_recruiter_and_priority_persist_through_add_and_edit,
        test_search_finds_by_code_name_client_and_id,
        test_migration_never_touches_existing_manually_coded_rows,
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
            for req_id in [FAKE_REQ_ID] + _extra_req_ids:
                conn.execute("DELETE FROM requirement WHERE id = %s", (req_id,))
            conn.execute("DELETE FROM app_user WHERE id = %s", (RECRUITER_ID,))
            conn.commit()

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All requirement-management checks passed.")


if __name__ == "__main__":
    main()
