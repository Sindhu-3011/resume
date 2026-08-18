"""Regression checks for the password-expiry feature (PASSWORD_MAX_AGE_DAYS).

Plain-Python, no pytest dependency — run directly:

    python tests/test_password_expiry.py

Uses a disposable app_user row, cleaned up in a finally block.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []
TEST_USERNAME = "zzz_pwexpiry_test"
TEST_EMAIL = "zzz.pwexpiry.test@example.com"
TEST_PASSWORD = "CorrectPass1!"
TEST_USER_ID = None


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def _create_test_user(password_changed_days_ago):
    with appmod.db_conn() as conn:
        appmod.ensure_users_table(conn)
        conn.execute("DELETE FROM app_user WHERE username = %s", (TEST_USERNAME,))
        conn.commit()
        user_id = conn.execute(
            """
            INSERT INTO app_user (username, password_hash, full_name, email, role, is_active,
                                   password_changed_at)
            VALUES (%s, %s, %s, %s, 'recruiter', TRUE, NOW() - (%s || ' days')::interval)
            RETURNING id
            """,
            (TEST_USERNAME, appmod.generate_password_hash(TEST_PASSWORD),
             "ZZZ Password Expiry Test", TEST_EMAIL, password_changed_days_ago),
        ).fetchone()["id"]
        conn.commit()
    return user_id


def _cleanup(user_id):
    with appmod.db_conn() as conn:
        conn.execute("DELETE FROM app_user WHERE id = %s", (user_id,))
        conn.commit()


def test_expired_password_blocks_login_until_changed(user_id):
    max_age = appmod.PASSWORD_MAX_AGE_DAYS
    client = appmod.app.test_client()
    r = client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})
    check("expired-password login redirects to the change-password route",
          r.status_code == 302 and "/change-expired-password" in r.headers.get("Location", ""))

    with client.session_transaction() as sess:
        check("no real session was established (no user_id)", "user_id" not in sess)
        check("a pending-change marker was set instead", sess.get("pending_password_change_user_id") == user_id)

    r = client.get("/dashboard")
    check("dashboard is still inaccessible (redirected to login)",
          r.status_code == 302 and "/login" in r.headers.get("Location", ""))


def test_change_expired_password_rejects_weak_and_reused(user_id):
    client = appmod.app.test_client()
    client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})

    r = client.post("/change-expired-password", data={"new_password": "weak", "confirm_password": "weak"})
    check("weak new password is rejected", r.status_code == 200)
    with client.session_transaction() as sess:
        check("still pending after a weak-password attempt",
              sess.get("pending_password_change_user_id") == user_id)

    r = client.post("/change-expired-password",
                     data={"new_password": TEST_PASSWORD, "confirm_password": TEST_PASSWORD})
    check("reusing the current password is rejected", r.status_code == 200)
    with client.session_transaction() as sess:
        check("still pending after a reused-password attempt",
              sess.get("pending_password_change_user_id") == user_id)


def test_change_expired_password_succeeds_and_establishes_real_session(user_id):
    client = appmod.app.test_client()
    client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})

    new_password = "BrandNewPass9!"
    r = client.post("/change-expired-password",
                     data={"new_password": new_password, "confirm_password": new_password})
    check("valid new password redirects to the dashboard", r.status_code == 302
          and "/dashboard" in r.headers.get("Location", ""))

    with client.session_transaction() as sess:
        check("pending marker cleared", "pending_password_change_user_id" not in sess)
        check("a real session was established", sess.get("user_id") == user_id)

    r = client.get("/dashboard")
    check("dashboard is now reachable with the real session", r.status_code == 200)

    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT password_changed_at FROM app_user WHERE id = %s", (user_id,)
        ).fetchone()
    check("password_changed_at was bumped to now (not still 91+ days old)",
          (appmod.datetime.now(appmod.timezone.utc) - row["password_changed_at"]).days < 1)

    client2 = appmod.app.test_client()
    r = client2.post("/login", data={"username": TEST_USERNAME, "password": new_password})
    check("can log in normally with the new password afterward", r.status_code == 302
          and "/dashboard" in r.headers.get("Location", ""))


def test_recent_password_logs_in_normally():
    user_id = _create_test_user(5)  # 5 days old — well within the default 90-day window
    try:
        client = appmod.app.test_client()
        r = client.post("/login", data={"username": TEST_USERNAME, "password": TEST_PASSWORD})
        check("a recent password logs straight in (no expiry detour)",
              r.status_code == 302 and "/dashboard" in r.headers.get("Location", ""))
        with client.session_transaction() as sess:
            check("a real session was established immediately", sess.get("user_id") == user_id)
    finally:
        _cleanup(user_id)


def main():
    max_age = appmod.PASSWORD_MAX_AGE_DAYS
    user_id = _create_test_user(max_age + 1)

    def t1(): test_expired_password_blocks_login_until_changed(user_id)
    def t2(): test_change_expired_password_rejects_weak_and_reused(user_id)
    def t3(): test_change_expired_password_succeeds_and_establishes_real_session(user_id)

    tests = [t1, t2, t3, test_recent_password_logs_in_normally]
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
    print("All password-expiry checks passed.")


if __name__ == "__main__":
    main()
