"""Regression checks for email validation on user account creation/editing:
duplicate-email prevention and invalid-format rejection.

Plain-Python, no pytest dependency — run directly:

    python tests/test_user_email_validation.py

Session state (user id/role) is injected directly via session_transaction()
— same technique other suites in this app use. A real disposable admin
app_user row is created since these actions require write_users.
"""

import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
ADMIN_ID = 999_999_701
EXISTING_USERNAME = "zzz_email_existing"
EXISTING_EMAIL = "zzz.email.taken@example.com"
_extra_user_ids = []


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


def _create_admin_test_user():
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE id = %s", (ADMIN_ID,))
        conn.commit()
        conn.execute(
            "INSERT INTO app_user (id, username, password_hash, full_name, email, role, is_active) "
            "VALUES (%s, %s, %s, %s, %s, 'admin', TRUE)",
            (ADMIN_ID, "zzz_emailval_admin", appmod.generate_password_hash("Whatever1!"),
             "ZZZ Email Val Admin", "zzz.emailval.admin@example.com"),
        )
        conn.commit()


def _create_existing_user():
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE username = %s", (EXISTING_USERNAME,))
        conn.commit()
        user_id = conn.execute(
            "INSERT INTO app_user (username, password_hash, full_name, email, role, is_active) "
            "VALUES (%s, %s, %s, %s, 'recruiter', TRUE) RETURNING id",
            (EXISTING_USERNAME, appmod.generate_password_hash("Whatever1!"),
             "ZZZ Email Existing", EXISTING_EMAIL),
        ).fetchone()["id"]
        conn.commit()
    return user_id


def _admin_client():
    client = appmod.app.test_client()
    _session_as(client, ADMIN_ID, "zzz_emailval_admin", "ZZZ Email Val Admin", "admin")
    return client


def test_add_rejects_invalid_email_format():
    client = _admin_client()
    r = client.post("/users/add", data={
        "username": "zzz_email_bad_format", "password": "Whatever1!",
        "full_name": "ZZZ Bad Format", "email": "not-an-email", "role": "viewer_auditor",
    })
    check("invalid email format is rejected (re-renders the form)", r.status_code == 200)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT id FROM app_user WHERE username = %s", ("zzz_email_bad_format",)).fetchone()
    check("no user was created with the invalid email", row is None)


def test_add_rejects_duplicate_email_case_insensitively():
    client = _admin_client()
    r = client.post("/users/add", data={
        "username": "zzz_email_dup_attempt", "password": "Whatever1!",
        "full_name": "ZZZ Dup Attempt", "email": EXISTING_EMAIL.upper(), "role": "viewer_auditor",
    })
    check("duplicate email (different case) is rejected", r.status_code == 200)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT id FROM app_user WHERE username = %s", ("zzz_email_dup_attempt",)).fetchone()
    check("no user was created with the duplicate email", row is None)


def test_add_accepts_valid_unique_email_and_blank_email():
    client = _admin_client()
    r = client.post("/users/add", data={
        "username": "zzz_email_good", "password": "Whatever1!",
        "full_name": "ZZZ Good Email", "email": "zzz.good.email@example.com", "role": "viewer_auditor",
    })
    check("a valid, unique email is accepted", r.status_code == 302)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT id FROM app_user WHERE username = %s", ("zzz_email_good",)).fetchone()
    check("the user was created", row is not None)
    if row:
        _extra_user_ids.append(row["id"])

    r = client.post("/users/add", data={
        "username": "zzz_email_blank", "password": "Whatever1!",
        "full_name": "ZZZ Blank Email", "email": "", "role": "viewer_auditor",
    })
    check("email remains optional — a blank email is still accepted", r.status_code == 302)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT id FROM app_user WHERE username = %s", ("zzz_email_blank",)).fetchone()
    check("the user with no email was created", row is not None)
    if row:
        _extra_user_ids.append(row["id"])


def test_edit_rejects_invalid_format_and_duplicate_but_allows_keeping_own_email(existing_user_id):
    client = _admin_client()

    r = client.post(f"/users/{existing_user_id}/edit", data={
        "full_name": "ZZZ Email Existing", "email": "still-not-an-email", "role": "recruiter", "password": "",
    })
    check("edit rejects an invalid email format", r.status_code == 200)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT email FROM app_user WHERE id = %s", (existing_user_id,)).fetchone()
    check("the email was NOT overwritten with the invalid value", row["email"] == EXISTING_EMAIL)

    other_id = _extra_user_ids[0] if _extra_user_ids else None
    if other_id:
        with appmod.db_conn() as conn:
            other_email = conn.execute("SELECT email FROM app_user WHERE id = %s", (other_id,)).fetchone()["email"]
        r = client.post(f"/users/{existing_user_id}/edit", data={
            "full_name": "ZZZ Email Existing", "email": other_email, "role": "recruiter", "password": "",
        })
        check("edit rejects an email already used by a different account", r.status_code == 200)
        with appmod.db_conn() as conn:
            row = conn.execute("SELECT email FROM app_user WHERE id = %s", (existing_user_id,)).fetchone()
        check("the email was NOT overwritten with the duplicate value", row["email"] == EXISTING_EMAIL)

    # Resubmitting the account's OWN unchanged email must not be flagged as a
    # duplicate-of-itself.
    r = client.post(f"/users/{existing_user_id}/edit", data={
        "full_name": "ZZZ Email Existing Renamed", "email": EXISTING_EMAIL, "role": "recruiter", "password": "",
    })
    check("keeping one's own existing email on save succeeds", r.status_code == 302)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT full_name, email FROM app_user WHERE id = %s", (existing_user_id,)).fetchone()
    check("the rest of the edit (full_name) still applied", row["full_name"] == "ZZZ Email Existing Renamed")
    check("the email is unchanged", row["email"] == EXISTING_EMAIL)


def main():
    _create_admin_test_user()
    existing_user_id = _create_existing_user()

    def t4():
        test_edit_rejects_invalid_format_and_duplicate_but_allows_keeping_own_email(existing_user_id)

    tests = [
        test_add_rejects_invalid_email_format,
        test_add_rejects_duplicate_email_case_insensitively,
        test_add_accepts_valid_unique_email_and_blank_email,
        t4,
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
            conn.execute("DELETE FROM app_user WHERE username = %s", (EXISTING_USERNAME,))
            for uid in _extra_user_ids:
                conn.execute("DELETE FROM app_user WHERE id = %s", (uid,))
            conn.execute("DELETE FROM app_user WHERE id = %s", (ADMIN_ID,))
            conn.commit()

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All user-email-validation checks passed.")


if __name__ == "__main__":
    main()
