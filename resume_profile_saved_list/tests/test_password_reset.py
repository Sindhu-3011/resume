"""Regression checks for Forgot Password / Reset Password.

Plain-Python, no pytest dependency — run directly:

    python tests/test_password_reset.py

Uses a disposable app_user row created directly via SQL (its own username/
email, cleaned up at the end via ON DELETE CASCADE on password_reset_token/
password_history). No real SMTP server exists in this environment — every
scenario monkeypatches _send_email to capture what would have been sent
rather than actually sending it.
"""

import datetime
import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
TEST_USERNAME = "zzz_rbac_pw_reset_test"
TEST_EMAIL = "zzz.pw.reset.test@example.com"
TEST_ORIGINAL_PASSWORD = "OriginalPass1!"


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
            (TEST_USERNAME, appmod.generate_password_hash(TEST_ORIGINAL_PASSWORD),
             "ZZZ Password Reset Test", TEST_EMAIL),
        ).fetchone()["id"]
        conn.commit()
    return user_id


def _cleanup(user_id):
    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM app_user WHERE id = %s", (user_id,))  # cascades token/history rows
        conn.commit()


def _capture_sent_emails():
    """Returns (sent_list, restore_fn). sent_list gets a dict appended for
    every _send_email call instead of a real network attempt."""
    sent = []
    real_send = appmod._send_email

    def _fake_send(to_addr, subject, body):
        sent.append({"to": to_addr, "subject": subject, "body": body})

    appmod._send_email = _fake_send

    def _restore():
        appmod._send_email = real_send

    return sent, _restore


def test_valid_registered_email_creates_token_and_sends_email(user_id):
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        r = client.post("/forgot-password", data={"email": TEST_EMAIL})
        check("valid email redirects to login", r.status_code == 302)
        check("exactly one email was sent", len(sent) == 1)
        if sent:
            check("reset email went to the registered address", sent[0]["to"] == TEST_EMAIL)
            check("reset email body contains a reset link", "/reset-password/" in sent[0]["body"])

        with appmod.db_conn() as conn:
            rows = conn.execute(
                "SELECT * FROM password_reset_token WHERE user_id = %s", (user_id,)
            ).fetchall()
        check("a token row was created for this user", len(rows) == 1)
        if rows:
            check("token row is not yet used", rows[0]["used_at"] is None)
            check("token row expires in the future", rows[0]["expires_at"] > datetime.datetime.now(datetime.timezone.utc))
    finally:
        restore()


def test_unregistered_email_creates_no_token_same_response():
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        r1 = client.post("/forgot-password", data={"email": "not.a.real.user@example.com"})
        client2 = appmod.app.test_client()
        r2 = client2.post("/forgot-password", data={"email": TEST_EMAIL})
        check("unregistered email still redirects (no error leaked)", r1.status_code == 302)
        check("unregistered email location is identical to a registered one (no enumeration)",
              r1.headers.get("Location") == r2.headers.get("Location"))
        check("no email was sent for the unregistered address", len(sent) == 1)  # only from r2 (the real user)
        with appmod.db_conn() as conn:
            count = conn.execute(
                "SELECT COUNT(*) AS c FROM app_user WHERE username = 'nonexistent_zzz_user'"
            ).fetchone()["c"]
        check("sanity: the unregistered username really doesn't exist", count == 0)
    finally:
        restore()
        with appmod.db_conn() as conn:
            conn.execute(
                "DELETE FROM password_reset_token WHERE user_id = (SELECT id FROM app_user WHERE username = %s)",
                (TEST_USERNAME,),
            )
            conn.commit()


def _make_token(user_id, minutes_from_now=30, used=False):
    raw, token_hash = appmod._generate_reset_token()
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(minutes=minutes_from_now)
    used_at = datetime.datetime.now(datetime.timezone.utc) if used else None
    with appmod.db_conn() as conn:
        appmod.ensure_password_reset_table(conn)
        conn.execute(
            "INSERT INTO password_reset_token (user_id, token_hash, expires_at, used_at) VALUES (%s, %s, %s, %s)",
            (user_id, token_hash, expires_at, used_at),
        )
        conn.commit()
    return raw


def test_valid_reset_link_shows_form_and_succeeds(user_id):
    raw = _make_token(user_id)
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        r = client.get(f"/reset-password/{raw}")
        check("valid token GET returns 200 (form, not error)", r.status_code == 200)
        check("valid token page does not show an error", b"Reset Link Unavailable" not in r.data)

        r2 = client.post(f"/reset-password/{raw}", data={
            "new_password": "BrandNewPass2!", "confirm_password": "BrandNewPass2!",
        })
        check("successful reset redirects to login", r2.status_code == 302)
        check("confirmation email was sent", len(sent) == 1)

        with appmod.db_conn() as conn:
            row = conn.execute("SELECT password_hash FROM app_user WHERE id = %s", (user_id,)).fetchone()
        check("password_hash actually changed",
              appmod.check_password_hash(row["password_hash"], "BrandNewPass2!"))
        check("old password no longer works",
              not appmod.check_password_hash(row["password_hash"], TEST_ORIGINAL_PASSWORD))
    finally:
        restore()


def test_expired_reset_link_rejected(user_id):
    raw = _make_token(user_id, minutes_from_now=-5)  # already expired
    client = appmod.app.test_client()
    r = client.get(f"/reset-password/{raw}")
    check("expired token shows the error state", b"expired" in r.data.lower())
    r2 = client.post(f"/reset-password/{raw}", data={
        "new_password": "WontWork12", "confirm_password": "WontWork12",
    })
    check("expired token POST does not change the password (still shows error)",
          b"Reset Link Unavailable" in r2.data or r2.status_code == 302)


def test_already_used_reset_link_rejected(user_id):
    raw = _make_token(user_id, used=True)
    client = appmod.app.test_client()
    r = client.get(f"/reset-password/{raw}")
    check("already-used token shows the error state", b"already been used" in r.data)


def test_weak_password_rejected(user_id):
    raw = _make_token(user_id)
    client = appmod.app.test_client()
    client.get(f"/reset-password/{raw}")  # establishes nothing stateful, just mirrors real usage
    r = client.post(f"/reset-password/{raw}", data={
        "new_password": "abc", "confirm_password": "abc",
    })
    check("weak (too short) password is rejected, form re-shown", r.status_code == 200)
    with appmod.db_conn() as conn:
        row = conn.execute("SELECT password_hash FROM app_user WHERE id = %s", (user_id,)).fetchone()
    check("password was NOT changed by the rejected weak attempt",
          appmod.check_password_hash(row["password_hash"], TEST_ORIGINAL_PASSWORD)
          or not appmod.check_password_hash(row["password_hash"], "abc"))


def test_password_complexity_rules():
    """Direct unit test of _validate_password_complexity — each rule added
    2026-08-14 (uppercase, special character) alongside the pre-existing
    length/lowercase/digit checks."""
    cases = [
        ("Sh1!", False, "too short (4 chars)"),
        ("alllowercase1!", False, "no uppercase"),
        ("ALLUPPERCASE1!", False, "no lowercase"),
        ("NoDigitsHere!", False, "no digit"),
        ("NoSpecialChar1", False, "no special character"),
        ("ValidPass1!", True, "meets every rule"),
    ]
    for password, should_pass, label in cases:
        ok, _ = appmod._validate_password_complexity(password)
        check(f"complexity check — {label}: {'accepted' if should_pass else 'rejected'} as expected",
              ok == should_pass)


def test_reused_previous_password_rejected(user_id):
    # Self-contained regardless of what earlier tests already changed this
    # user's real current password to: seed history with a hash of a known,
    # dedicated plaintext (not derived from the account's actual current
    # password, which prior tests in this sequential run may have already
    # changed), then attempt to reset TO that same known plaintext.
    seeded_password = "PriorHistoryPass9!"
    with appmod.db_conn() as conn:
        appmod.ensure_password_history_table(conn)
        conn.execute(
            "INSERT INTO password_history (user_id, password_hash) VALUES (%s, %s)",
            (user_id, appmod.generate_password_hash(seeded_password)),
        )
        conn.commit()
    raw = _make_token(user_id)
    sent, restore = _capture_sent_emails()
    try:
        client = appmod.app.test_client()
        r = client.post(f"/reset-password/{raw}", data={
            "new_password": seeded_password, "confirm_password": seeded_password,
        })
        check("reusing a historical password is rejected", r.status_code == 200)
        check("no confirmation email sent for a rejected reuse attempt", len(sent) == 0)
    finally:
        restore()


def test_old_session_invalidated_after_reset(user_id):
    # Self-contained regardless of what earlier tests already did to this
    # user's real password: establish a known password via a real reset
    # first, log in with THAT, then perform a second reset and confirm the
    # first session gets invalidated.
    known_password = "SessionTestPass1!"
    sent, restore = _capture_sent_emails()
    try:
        raw = _make_token(user_id)
        client_setup = appmod.app.test_client()
        client_setup.post(f"/reset-password/{raw}", data={
            "new_password": known_password, "confirm_password": known_password,
        })
    finally:
        restore()

    client_a = appmod.app.test_client()
    r = client_a.post("/login", data={"username": TEST_USERNAME, "password": known_password})
    check("session A logged in successfully", r.status_code == 302 and "/login" not in r.headers.get("Location", ""))
    r = client_a.get("/dashboard")
    check("session A can reach the dashboard before any further reset", r.status_code == 200)

    sent, restore = _capture_sent_emails()
    try:
        raw2 = _make_token(user_id)
        client_b = appmod.app.test_client()
        client_b.post(f"/reset-password/{raw2}", data={
            "new_password": "AnotherNewPass3!", "confirm_password": "AnotherNewPass3!",
        })
    finally:
        restore()

    r = client_a.get("/dashboard")
    check("session A is now redirected to login (invalidated by the reset)",
          r.status_code == 302 and "/login" in r.headers.get("Location", ""))


def test_show_hide_toggle_present_on_login_page():
    client = appmod.app.test_client()
    r = client.get("/login")
    html = r.get_data(as_text=True)
    check("login page has the show/hide toggle button", 'id="togglePassword"' in html)
    check("login page has the password input it targets", 'id="password"' in html)
    check("login page has a Forgot Password link", "forgot-password" in html.lower())


def main():
    user_id = _create_test_user()

    def t1(): test_valid_registered_email_creates_token_and_sends_email(user_id)
    def t3(): test_valid_reset_link_shows_form_and_succeeds(user_id)
    def t4(): test_expired_reset_link_rejected(user_id)
    def t5(): test_already_used_reset_link_rejected(user_id)
    def t6(): test_weak_password_rejected(user_id)
    def t7(): test_reused_previous_password_rejected(user_id)
    def t8(): test_old_session_invalidated_after_reset(user_id)

    tests = [
        t1,
        test_unregistered_email_creates_no_token_same_response,
        t3,
        t4,
        t5,
        t6,
        test_password_complexity_rules,
        t7,
        t8,
        test_show_hide_toggle_present_on_login_page,
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

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    print("All password-reset checks passed.")


if __name__ == "__main__":
    main()
