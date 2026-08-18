"""Regression checks for last-login timestamps and role-change history on
user accounts.

Plain-Python, no pytest dependency — run directly:

    python tests/test_user_account_audit.py

Session state (user id/role) is injected directly via session_transaction()
— same technique other suites in this app use — so this never depends on
the real seeded admin account's actual current password. A real disposable
admin app_user row is still created, since user_role_history.changed_by is
a foreign key into app_user.
"""

import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
ADMIN_ID = 999_999_601          # role: admin — performs the add/edit actions below
TARGET_USERNAME = "zzz_audit_target"
TARGET_EMAIL = "zzz.audit.target@example.com"
TARGET_PASSWORD = "CorrectPass1!"
_extra_user_ids = []  # ids created via the real user_add route, swept up on cleanup


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
            (ADMIN_ID, "zzz_audit_admin", appmod.generate_password_hash("Whatever1!"),
             "ZZZ Audit Admin", "zzz.audit.admin@example.com"),
        )
        conn.commit()


def _create_target_user(role="recruiter"):
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE username = %s", (TARGET_USERNAME,))
        conn.commit()
        user_id = conn.execute(
            "INSERT INTO app_user (username, password_hash, full_name, email, role, is_active) "
            "VALUES (%s, %s, %s, %s, %s, TRUE) RETURNING id",
            (TARGET_USERNAME, appmod.generate_password_hash(TARGET_PASSWORD),
             "ZZZ Audit Target", TARGET_EMAIL, role),
        ).fetchone()["id"]
        conn.commit()
    return user_id


def _role_history_rows(user_id):
    with appmod.db_conn() as conn:
        appmod.ensure_user_role_history_table(conn)
        return conn.execute(
            "SELECT * FROM user_role_history WHERE user_id = %s ORDER BY id", (user_id,)
        ).fetchall()


def test_last_login_stamped_on_successful_login():
    user_id = _create_target_user()
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT last_login_at FROM app_user WHERE id = %s", (user_id,)).fetchone()
    check("a brand-new user has never logged in", row["last_login_at"] is None)

    client = appmod.app.test_client()
    r = client.post("/login", data={"username": TARGET_USERNAME, "password": TARGET_PASSWORD})
    check("login succeeds", r.status_code == 302)

    with appmod.db_conn() as conn:
        row = conn.execute("SELECT last_login_at FROM app_user WHERE id = %s", (user_id,)).fetchone()
    check("last_login_at is now stamped", row["last_login_at"] is not None)
    age_seconds = (datetime.datetime.now(datetime.timezone.utc) - row["last_login_at"]).total_seconds()
    check("the stamp is recent (within the last minute)", age_seconds < 60)


def test_user_add_logs_initial_role_history():
    client = appmod.app.test_client()
    _session_as(client, ADMIN_ID, "zzz_audit_admin", "ZZZ Audit Admin", "admin")

    r = client.post("/users/add", data={
        "username": "zzz_audit_new_user", "password": "BrandNew1!",
        "full_name": "ZZZ New User", "email": "zzz.new@example.com", "role": "hiring_manager",
    })
    check("user_add redirects on success", r.status_code == 302)

    with appmod.db_conn() as conn:
        row = conn.execute("SELECT id FROM app_user WHERE username = %s", ("zzz_audit_new_user",)).fetchone()
    check("the new user was created", row is not None)
    if row:
        _extra_user_ids.append(row["id"])
        history = _role_history_rows(row["id"])
        check("exactly one history row exists (the initial assignment)", len(history) == 1)
        if history:
            check("old_role is NULL for the initial assignment", history[0]["old_role"] is None)
            check("new_role matches what was submitted", history[0]["new_role"] == "hiring_manager")
            check("changed_by is the acting admin", history[0]["changed_by"] == ADMIN_ID)


def test_user_edit_logs_role_change_but_not_a_no_op_save():
    user_id = _create_target_user(role="recruiter")
    client = appmod.app.test_client()
    _session_as(client, ADMIN_ID, "zzz_audit_admin", "ZZZ Audit Admin", "admin")

    # A save that does NOT change the role must not add a history row.
    r = client.post(f"/users/{user_id}/edit", data={
        "full_name": "ZZZ Audit Target Renamed", "email": TARGET_EMAIL, "role": "recruiter", "password": "",
    })
    check("no-op-role edit redirects", r.status_code == 302)
    check("no history row added when the role didn't change", len(_role_history_rows(user_id)) == 0)

    # A save that DOES change the role must add exactly one history row.
    r = client.post(f"/users/{user_id}/edit", data={
        "full_name": "ZZZ Audit Target Renamed", "email": TARGET_EMAIL, "role": "hiring_manager", "password": "",
    })
    check("role-change edit redirects", r.status_code == 302)
    history = _role_history_rows(user_id)
    check("exactly one history row added for the role change", len(history) == 1)
    if history:
        check("old_role is the previous role", history[0]["old_role"] == "recruiter")
        check("new_role is the newly submitted role", history[0]["new_role"] == "hiring_manager")
        check("changed_by is the acting admin", history[0]["changed_by"] == ADMIN_ID)


def test_role_history_page_renders_and_is_admin_gated():
    user_id = _create_target_user(role="recruiter")
    with appmod.db_conn() as conn:
        appmod.ensure_user_role_history_table(conn)
        conn.execute(
            "INSERT INTO user_role_history (user_id, old_role, new_role, changed_by) VALUES (%s, NULL, %s, %s)",
            (user_id, "recruiter", ADMIN_ID),
        )
        conn.commit()

    admin_client = appmod.app.test_client()
    _session_as(admin_client, ADMIN_ID, "zzz_audit_admin", "ZZZ Audit Admin", "admin")
    r = admin_client.get(f"/users/{user_id}/role-history")
    check("admin can view the role-history page", r.status_code == 200)
    check("the page shows the recorded role", b"Recruiter" in r.data)

    non_admin_client = appmod.app.test_client()
    _session_as(non_admin_client, 999_999_602, "zzz_audit_recruiter", "ZZZ Audit Recruiter", "recruiter")
    r = non_admin_client.get(f"/users/{user_id}/role-history")
    check("a non-admin role is blocked from the role-history page", r.status_code == 403)


def main():
    _create_admin_test_user()
    tests = [
        test_last_login_stamped_on_successful_login,
        test_user_add_logs_initial_role_history,
        test_user_edit_logs_role_change_but_not_a_no_op_save,
        test_role_history_page_renders_and_is_admin_gated,
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
            conn.execute("DELETE FROM app_user WHERE username = %s", (TARGET_USERNAME,))
            for uid in _extra_user_ids:
                conn.execute("DELETE FROM app_user WHERE id = %s", (uid,))
            conn.execute("DELETE FROM app_user WHERE id = %s", (ADMIN_ID,))
            conn.commit()

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All user-account-audit checks passed.")


if __name__ == "__main__":
    main()
