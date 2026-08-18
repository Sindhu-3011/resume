"""Regression checks for account lockout on repeated failed logins.

Plain-Python, no pytest dependency — run directly:

    python tests/test_account_lockout.py

Uses a disposable app_user row created directly via SQL, cleaned up at the
end. No real SMTP server exists in this environment — every scenario
monkeypatches _send_email to capture what would have been sent rather than
actually sending it.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
TEST_USERNAME = "zzz_lockout_test"
TEST_EMAIL = "zzz.lockout.test@example.com"
TEST_PASSWORD = "CorrectPass1!"
TEST_ADMIN_USERNAME = "zzz_lockout_admin_test"
TEST_ADMIN_PASSWORD = "AdminCorrect1!"


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def _create_test_user():
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE username = %s", (TEST_USERNAME,))
        conn.commit()
        user_id = conn.execute(
            """
            INSERT INTO app_user (username, password_hash, full_name, email, role, is_active)
            VALUES (%s, %s, %s, %s, 'recruiter', TRUE)
            RETURNING id
            """,
            (TEST_USERNAME, appmod.generate_password_hash(TEST_PASSWORD),
             "ZZZ Lockout Test", TEST_EMAIL),
        ).fetchone()["id"]
        conn.commit()
    return user_id


def _create_admin_test_user():
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE username = %s", (TEST_ADMIN_USERNAME,))
        conn.commit()
        user_id = conn.execute(
            """
            INSERT INTO app_user (username, password_hash, full_name, email, role, is_active)
            VALUES (%s, %s, %s, %s, 'admin', TRUE)
            RETURNING id
            """,
            (TEST_ADMIN_USERNAME, appmod.generate_password_hash(TEST_ADMIN_PASSWORD),
             "ZZZ Lockout Admin Test", "zzz.lockout.admin@example.com"),
        ).fetchone()["id"]
        conn.commit()
    return user_id


def _reset_lockout_state(user_id):
    with appmod.db_conn() as conn:
        conn.execute(
            "UPDATE app_user SET failed_login_attempts = 0, locked_at = NULL WHERE id = %s",
            (user_id,),
        )
        conn.commit()


def _cleanup(user_id):
    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM app_user WHERE id = %s", (user_id,))
        conn.commit()


def _capture_sent_emails():
    sent = []
    real_send = appmod._send_email

    def _fake_send(to_addr, subject, body):
        sent.append({"to": to_addr, "subject": subject, "body": body})

    appmod._send_email = _fake_send

    def _restore():
        appmod._send_email = real_send

    return sent, _restore


def test_failed_attempts_show_decrementing_remaining_count(user_id):
    _reset_lockout_state(user_id)
    client = appmod.app.test_client()
    max_attempts = appmod.MAX_FAILED_LOGIN_ATTEMPTS
    for attempt_number in range(1, max_attempts):  # every attempt short of locking
        r = client.post("/login", data={"username": TEST_USERNAME, "password": "WrongPass!"})
        expected_remaining = max_attempts - attempt_number
        check(f"attempt {attempt_number}: shows {expected_remaining} remaining",
              f"{expected_remaining} attempt(s) remaining".encode() in r.data)
    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT failed_login_attempts, locked_at FROM app_user WHERE id = %s", (user_id,)
        ).fetchone()
    check(f"failed_login_attempts is {max_attempts - 1} after {max_attempts - 1} failures",
          row["failed_login_attempts"] == max_attempts - 1)
    check("account is not yet locked", row["locked_at"] is None)


def test_nth_failure_locks_account_and_sends_email(user_id):
    _reset_lockout_state(user_id)
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        max_attempts = appmod.MAX_FAILED_LOGIN_ATTEMPTS
        r = None
        for _ in range(max_attempts):
            r = client.post("/login", data={"username": TEST_USERNAME, "password": "WrongPass!"})
        check(f"the {max_attempts}th failure shows the locked message",
              b"account has been locked" in r.data)
        check("exactly one lockout email was sent", len(sent) == 1)
        if sent:
            check("lockout email went to the account's registered address", sent[0]["to"] == TEST_EMAIL)
            check("lockout email subject mentions locking", "locked" in sent[0]["subject"].lower())

        with appmod.db_conn() as conn:
            row = conn.execute(
                "SELECT failed_login_attempts, locked_at FROM app_user WHERE id = %s", (user_id,)
            ).fetchone()
        check("locked_at is now set", row["locked_at"] is not None)
        check(f"failed_login_attempts reached {max_attempts}", row["failed_login_attempts"] == max_attempts)
    finally:
        restore()


def test_locked_account_rejected_even_with_correct_password(user_id):
    # Reuses the lock from the previous test's state (or locks it fresh if
    # run standalone) — the point is the CORRECT password must still fail.
    with appmod.db_conn() as conn:
        conn.execute("UPDATE app_user SET locked_at = NOW() WHERE id = %s", (user_id,))
        conn.commit()
    client = appmod.app.test_client()
    r = client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})
    check("correct password on a locked account is still rejected", r.status_code == 200)
    check("locked-account message is shown even with the right password",
          b"account has been locked" in r.data)
    r2 = client.get("/dashboard")
    check("no session was established for the locked account",
          r2.status_code == 302 and "/login" in r2.headers.get("Location", ""))


def test_successful_login_resets_failed_attempt_counter(user_id):
    _reset_lockout_state(user_id)
    client = appmod.app.test_client()
    max_attempts = appmod.MAX_FAILED_LOGIN_ATTEMPTS

    client.post("/login", data={"username": TEST_USERNAME, "password": "WrongPass!"})
    client.post("/login", data={"username": TEST_USERNAME, "password": "WrongPass!"})
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT failed_login_attempts FROM app_user WHERE id = %s", (user_id,)).fetchone()
    check("2 failures recorded before the successful login", row["failed_login_attempts"] == 2)

    r = client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})
    check("correct password succeeds (redirect)", r.status_code == 302)

    with appmod.db_conn() as conn:
        row = conn.execute("SELECT failed_login_attempts FROM app_user WHERE id = %s", (user_id,)).fetchone()
    check("failed_login_attempts reset to 0 after a successful login", row["failed_login_attempts"] == 0)

    client2 = appmod.app.test_client()
    r = client2.post("/login", data={"username": TEST_USERNAME, "password": "WrongPass!"})
    check(f"next failure after a reset shows {max_attempts - 1} remaining, not {max_attempts - 3}",
          f"{max_attempts - 1} attempt(s) remaining".encode() in r.data)


def test_admin_unlock_restores_login(user_id):
    with appmod.db_conn() as conn:
        conn.execute(
            "UPDATE app_user SET locked_at = NOW(), failed_login_attempts = %s WHERE id = %s",
            (appmod.MAX_FAILED_LOGIN_ATTEMPTS, user_id),
        )
        conn.commit()

    # Session-injected admin role (same technique test_rbac.py uses) rather
    # than a real username/password login — avoids ever depending on the
    # real seeded admin account's actual current password, which can drift
    # (e.g. changed via User Management in the live app) independently of
    # this test suite.
    import datetime
    admin_client = appmod.app.test_client()
    with admin_client.session_transaction() as sess:
        sess["user_id"] = 999998
        sess["username"] = "zzz_test_admin"
        sess["full_name"] = "ZZZ Test Admin"
        sess["role"] = "admin"
        sess["last_active"] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    r = admin_client.post(f"/users/{user_id}/unlock")
    check("admin unlock request redirects to user management", r.status_code == 302)

    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT locked_at, failed_login_attempts FROM app_user WHERE id = %s", (user_id,)
        ).fetchone()
    check("locked_at cleared by unlock", row["locked_at"] is None)
    check("failed_login_attempts reset to 0 by unlock", row["failed_login_attempts"] == 0)

    client = appmod.app.test_client()
    r = client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})
    check("account can log in again after being unlocked", r.status_code == 302)


def test_nonexistent_username_never_shows_attempt_counter():
    client = appmod.app.test_client()
    r = client.post("/login", data={"username": "zzz_definitely_not_a_real_user", "password": "whatever"})
    check("nonexistent username gets the generic message", b"Invalid username or password" in r.data)
    check("nonexistent username never shows an attempts-remaining counter",
          b"attempt(s) remaining" not in r.data)


def test_admin_account_is_never_locked(admin_user_id):
    """Regression test for the exact incident found live on 2026-08-14: the
    only admin account locked itself out, leaving no one able to unlock it.
    Admin accounts must keep tracking failed attempts (for audit) but never
    actually flip locked_at, and must never show lock-threat messaging that
    would be misleading since it can't happen."""
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        max_attempts = appmod.MAX_FAILED_LOGIN_ATTEMPTS
        r = None
        for _ in range(max_attempts + 3):  # well past the normal threshold
            r = client.post("/login", data={"username": TEST_ADMIN_USERNAME, "password": "WrongPass!"})
            check("admin failure never shows attempts-remaining/lock-threat messaging",
                  b"attempt(s) remaining" not in r.data and b"account has been locked" not in r.data)
        check("admin failure always shows the plain generic message", b"Invalid username or password" in r.data)
        check("no lockout email was ever sent for the admin account", len(sent) == 0)

        with appmod.db_conn() as conn:
            row = conn.execute(
                "SELECT failed_login_attempts, locked_at FROM app_user WHERE id = %s", (admin_user_id,)
            ).fetchone()
        check(f"failed_login_attempts still tracked ({max_attempts + 3}) for audit visibility",
              row["failed_login_attempts"] == max_attempts + 3)
        check("locked_at was never set, despite far exceeding the normal threshold",
              row["locked_at"] is None)

        r = client.post("/login", data={"username": TEST_ADMIN_USERNAME, "password": TEST_ADMIN_PASSWORD})
        check("admin can still log in with the correct password after many failures", r.status_code == 302)
    finally:
        restore()


def main():
    user_id = _create_test_user()
    admin_user_id = _create_admin_test_user()

    def t1(): test_failed_attempts_show_decrementing_remaining_count(user_id)
    def t2(): test_nth_failure_locks_account_and_sends_email(user_id)
    def t3(): test_locked_account_rejected_even_with_correct_password(user_id)
    def t4(): test_successful_login_resets_failed_attempt_counter(user_id)
    def t5(): test_admin_unlock_restores_login(user_id)
    def t6(): test_admin_account_is_never_locked(admin_user_id)

    tests = [
        t1,
        t2,
        t3,
        t4,
        t5,
        test_nonexistent_username_never_shows_attempt_counter,
        t6,
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
        _cleanup(user_id)
        _cleanup(admin_user_id)

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All account-lockout checks passed.")


if __name__ == "__main__":
    main()
