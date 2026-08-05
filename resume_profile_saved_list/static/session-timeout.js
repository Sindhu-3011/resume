/* Auto-logout on inactivity — global, loaded on every authenticated page
 * (see base.html). Configuration (window.SESSION_TIMEOUT_MS etc.) is
 * injected by base.html from the server's SESSION_TIMEOUT_MINUTES config,
 * so this file never hardcodes the duration.
 *
 * Design: the CLIENT tracks real DOM activity (mouse/keyboard/scroll/click)
 * and shows a "still there?" warning SESSION_WARNING_MS before the
 * configured timeout, giving the user a chance to continue. The SERVER
 * independently enforces the same timeout on every request (see
 * _require_login() in app.py) as a backstop for when this script never gets
 * a chance to run (JS disabled, tab frozen, browser killed) — so the
 * session is never trusted to stay alive on the client's word alone.
 */
(function () {
    if (!window.SESSION_TIMEOUT_MS) return; // not logged in, or not configured

    var TIMEOUT_MS = window.SESSION_TIMEOUT_MS;
    var WARNING_MS = window.SESSION_WARNING_MS || 60000;
    var LOGOUT_URL = window.SESSION_LOGOUT_URL;
    var KEEPALIVE_URL = window.SESSION_KEEPALIVE_URL;

    var warnTimer = null;
    var expireTimer = null;
    var countdownInterval = null;
    var modal = null;

    function buildModal() {
        var overlay = document.createElement('div');
        overlay.id = 'session-timeout-overlay';
        overlay.setAttribute('role', 'alertdialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.setAttribute('aria-labelledby', 'session-timeout-heading');
        overlay.style.cssText =
            'position:fixed;inset:0;background:rgba(15,23,42,0.55);display:none;' +
            'align-items:center;justify-content:center;z-index:9999;';

        var box = document.createElement('div');
        box.style.cssText =
            'background:var(--surface,#fff);color:var(--tx,#1e293b);border-radius:14px;' +
            'padding:28px 30px;max-width:380px;width:90%;box-shadow:0 20px 50px rgba(0,0,0,0.25);text-align:center;';
        box.innerHTML =
            '<div style="font-size:2rem;margin-bottom:10px;" aria-hidden="true">&#9200;</div>' +
            '<h3 id="session-timeout-heading" style="margin:0 0 8px;font-size:1.1rem;">Still there?</h3>' +
            '<p style="margin:0 0 18px;font-size:0.9rem;color:var(--tx-muted,#64748b);line-height:1.5;">' +
            'You will be logged out in <strong><span id="session-timeout-countdown">60</span>s</strong> due to inactivity.</p>' +
            '<button id="session-timeout-continue" type="button" ' +
            'style="background:linear-gradient(90deg,#6366f1,#4f46e5);color:#fff;border:none;border-radius:8px;' +
            'padding:10px 22px;font-weight:700;font-size:0.9rem;cursor:pointer;width:100%;">Continue Session</button>';

        overlay.appendChild(box);
        document.body.appendChild(overlay);
        overlay.querySelector('#session-timeout-continue').addEventListener('click', continueSession);
        return overlay;
    }

    function showWarning() {
        if (!modal) modal = buildModal();
        modal.style.display = 'flex';
        var remaining = Math.round(WARNING_MS / 1000);
        var countdownEl = document.getElementById('session-timeout-countdown');
        if (countdownEl) countdownEl.textContent = remaining;
        countdownInterval = setInterval(function () {
            remaining -= 1;
            if (countdownEl) countdownEl.textContent = Math.max(remaining, 0);
            if (remaining <= 0) clearInterval(countdownInterval);
        }, 1000);

        expireTimer = setTimeout(function () {
            window.location.href = LOGOUT_URL + '?reason=inactivity';
        }, WARNING_MS);
    }

    function hideWarning() {
        if (modal) modal.style.display = 'none';
        if (countdownInterval) clearInterval(countdownInterval);
        if (expireTimer) clearTimeout(expireTimer);
    }

    function continueSession() {
        hideWarning();
        if (KEEPALIVE_URL) {
            fetch(KEEPALIVE_URL, {credentials: 'same-origin'}).catch(function () {});
        }
        resetTimers();
    }

    function resetTimers() {
        if (warnTimer) clearTimeout(warnTimer);
        if (expireTimer) clearTimeout(expireTimer);
        if (countdownInterval) clearInterval(countdownInterval);
        var untilWarning = Math.max(TIMEOUT_MS - WARNING_MS, 0);
        warnTimer = setTimeout(showWarning, untilWarning);
    }

    // Throttled so a stream of mousemove/scroll events doesn't reset the
    // timer hundreds of times a second — once every 5s is plenty to count
    // as "still active" without the overhead of doing it on every pixel.
    var throttled = false;
    function onActivity() {
        // While the warning is up, only the explicit "Continue Session"
        // click should reset the clock — passive mouse movement shouldn't
        // silently dismiss a dialog the user is meant to consciously act on.
        if (modal && modal.style.display === 'flex') return;
        if (throttled) return;
        throttled = true;
        setTimeout(function () { throttled = false; }, 5000);
        resetTimers();
    }

    ['mousemove', 'keydown', 'scroll', 'click', 'touchstart'].forEach(function (evt) {
        document.addEventListener(evt, onActivity, {passive: true});
    });

    resetTimers();
})();
