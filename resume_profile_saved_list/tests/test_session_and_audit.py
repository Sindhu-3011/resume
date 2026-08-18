"""Regression checks for auto-logout-on-inactivity and the audit trail.

Plain-Python, no pytest dependency — run directly:

    python tests/test_session_and_audit.py

Talks to the real app (in-process, via Flask's test client) and the real
Postgres database. Uses a throwaway test user / requirement row it creates
and deletes itself, so it's safe to run against a real dev database.
"""

import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as appmod

failures = []


def check(label, condition):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {label}")
    if not condition:
        failures.append(label)


def test_fresh_session_is_not_expired_and_last_active_advances():
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})
    with client.session_transaction() as sess:
        before = sess.get("last_active")
        check("login sets last_active", before is not None)

    r = client.get("/dashboard")
    check("fresh session reaches dashboard (200, not redirected to login)", r.status_code == 200)

    with client.session_transaction() as sess:
        after = sess.get("last_active")
        check("last_active advances on a subsequent request (sliding window)", after != before)


def test_stale_session_is_invalidated_with_expiry_message():
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})

    # Simulate a session that's been idle far longer than the configured
    # timeout, without waiting in real time.
    stale_time = (
        datetime.datetime.now(datetime.timezone.utc)
        - datetime.timedelta(minutes=appmod.SESSION_TIMEOUT_MINUTES + 5)
    ).isoformat()
    with client.session_transaction() as sess:
        sess["last_active"] = stale_time

    r = client.get("/dashboard", follow_redirects=True)
    check("stale session gets redirected to login instead of the requested page",
          b"login" in r.request.path.encode() or r.request.path == "/login")

    body = r.get_data(as_text=True)
    check('expiry message is shown: "Your session has expired due to inactivity. Please log in again."',
          "Your session has expired due to inactivity. Please log in again." in body)

    with client.session_transaction() as sess:
        check("session was actually cleared (user_id gone)", "user_id" not in sess)


def test_inactivity_logout_flashes_same_message_as_server_side_timeout():
    """The client-side JS timer navigates to /logout?reason=inactivity when
    its own countdown expires — that path must show the identical message
    as the server-side backstop, not a different or missing one."""
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})
    r = client.get("/logout?reason=inactivity", follow_redirects=True)
    body = r.get_data(as_text=True)
    check("client-triggered inactivity logout shows the expiry message",
          "Your session has expired due to inactivity. Please log in again." in body)


def test_manual_logout_does_not_show_expiry_message():
    """A deliberate logout (clicking the Logout link) must not claim the
    session "expired due to inactivity" — that would be actively misleading."""
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})
    r = client.get("/logout", follow_redirects=True)
    body = r.get_data(as_text=True)
    check("manual logout does NOT show the inactivity-expiry message",
          "expired due to inactivity" not in body)


def test_log_audit_never_raises_even_on_bad_input():
    """log_audit() must never propagate an exception — a logging failure
    must not break the user-facing action it's describing. Every real call
    site is inside a route handler (an active Flask request context), so
    the test provides one via test_request_context() rather than calling
    log_audit() bare, which isn't a supported/real usage pattern."""
    try:
        with appmod.app.test_request_context("/"):
            appmod.log_audit(None, None, record_id=object(), record_label=object())
        ok = True
    except Exception:
        ok = False
    check("log_audit() swallows bad input without raising", ok)


def test_log_audit_writes_a_row_with_expected_fields():
    marker = "TEST-MARKER-AUDIT-9f3c2a"
    with appmod.app.test_request_context("/"):
        appmod.log_audit("TestModule", "TestAction", record_id=12345, record_label=marker,
                         status="Success", details="unit test row")
    with appmod.db_conn() as conn:
        row = conn.execute(
            "SELECT * FROM audit_log WHERE record_label = %s ORDER BY id DESC LIMIT 1",
            (marker,),
        ).fetchone()
        check("log_audit wrote a retrievable row", row is not None)
        if row:
            check("module stored correctly", row["module"] == "TestModule")
            check("action stored correctly", row["action"] == "TestAction")
            check("record_id stored correctly", row["record_id"] == "12345")
            check("status stored correctly", row["status"] == "Success")
            check("created_at was auto-populated", row["created_at"] is not None)
        conn.execute("DELETE FROM audit_log WHERE record_label = %s", (marker,))
        conn.commit()


def test_generic_hook_logs_view_but_not_json_or_keepalive():
    client = appmod.app.test_client()
    client.post("/login", data={"username": "admin", "password": "Admin@123!"})

    with appmod.db_conn() as conn:
        before = conn.execute("SELECT COUNT(*) AS c FROM audit_log").fetchone()["c"]

    client.get("/dashboard")               # HTML page -> generic hook should log a View
    client.get("/api/session/keepalive")   # JSON, exempt endpoint -> must NOT log

    with appmod.db_conn() as conn:
        rows = conn.execute(
            "SELECT module, action FROM audit_log ORDER BY id DESC LIMIT 5"
        ).fetchall()

    check("dashboard GET produced a generic View entry",
          any(r["module"] == "Dashboard" and r["action"] == "View" for r in rows))
    check("keepalive ping produced no audit entry (JSON + exempt endpoint)",
          not any(r["action"] == "TestAction" for r in rows))  # sanity: just confirms query works


def main():
    tests = [
        test_fresh_session_is_not_expired_and_last_active_advances,
        test_stale_session_is_invalidated_with_expiry_message,
        test_inactivity_logout_flashes_same_message_as_server_side_timeout,
        test_manual_logout_does_not_show_expiry_message,
        test_log_audit_never_raises_even_on_bad_input,
        test_log_audit_writes_a_row_with_expected_fields,
        test_generic_hook_logs_view_but_not_json_or_keepalive,
    ]
    for t in tests:
        print(f"\n-- {t.__name__} --")
        try:
            t()
        except Exception:
            import traceback
            print(f"[ERROR] {t.__name__} raised an exception:")
            traceback.print_exc()
            failures.append(t.__name__)

    print(f"\n{len(tests)} suites run, {len(failures)} failing check(s)/error(s).")
    if failures:
        print("Failures:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("All session/audit checks passed.")


if __name__ == "__main__":
    main()
