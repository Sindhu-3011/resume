# DEVELOPER_GUIDE.md — Conventions, Testing, Debugging, Change Impact

Read this file for: coding standards, folder structure, how to add a feature, debugging tips,
what breaks what. For architecture/routes/DB, see **PROJECT_CONTEXT.md**. For resume parsing / AI
matching internals, see **AI_PIPELINE.md**. (`README.md`/`QUICK_START.md`/`SETUP_GUIDE.md` already
cover install/run instructions and a user-facing feature list — not duplicated here.)

## 1. Folder structure

```
app.py                  Single Flask monolith (~12,100 lines) — routes, DB, parsing, AI, everything
create_tables.sql       Base schema (resume, resume_skill) — run manually, not by the app
populate.sql            Seed data
create_db.py            One-off DB bootstrap helper
requirements.txt        Python deps
.env                    Config (DATABASE_URL, SECRET_KEY, OLLAMA_*, SESSION_TIMEOUT_MINUTES)
templates/              Jinja2 templates (one per page — see PROJECT_CONTEXT.md §6)
static/                 styles.css, session-timeout.js, verdict.js — no build step
tests/                  5 plain-Python test files (see §4)
scripts/check_parser.py Pre-commit parser regression guard (see §3)
.githooks/pre-commit     Runs check_parser.py when app.py or itself is staged
uploads/                 Uploaded resume/JD files (gitignored — contains real candidate PII)
```

No `src/`, no package structure, no blueprints — everything is a top-level function/route in
`app.py`. This is intentional for this project's size; don't start splitting it into modules
without discussing it first, since a lot of tooling (the pre-commit guard, this doc set) assumes
the single-file layout.

## 2. Coding conventions actually used in this codebase

- **DB access**: always `with db_conn() as conn: conn.execute(sql, params)`. Never build SQL by
  string-formatting user input — the one exception is a small, explicitly-validated allowlist
  (e.g. audit-trail sort column/direction), and even that is checked against a fixed set before
  interpolation.
- **`ensure_X_table(conn)` functions** (`CREATE TABLE IF NOT EXISTS ...`): called from *inside* the
  route that needs the table, every time — cheap, idempotent, no separate migration step to
  remember to run.
- **`_ensure_X_col()` functions** (`ALTER TABLE ... ADD COLUMN IF NOT EXISTS ...`): called **once,
  at module import time** (a bare function call at the bottom of the function definition, not
  inside a route) — this is how new columns get added to already-existing tables/deployments
  without a manual migration step. Wrap the whole thing in `try/except: pass` (matches existing
  pattern) since this runs on every process start and must never crash the app.
- **Single source of truth, don't duplicate**: `_tier_from_pct()` (score→label bucketing),
  `_CATEGORY_WEIGHTS`, `POSITION_STATUSES` (shared between JD and Requirement), `_evaluate_authenticity()`.
  If you need a similar bucketing/threshold function, check whether one already exists before
  writing a new one — this codebase used to have 3 separate copies of the tier-threshold logic at
  different cutoffs, which was a real, confusing bug.
- **Audit logging**: `log_audit(module, action, record_id=None, record_label=None, status="Success",
  details=None)` — call it explicitly at any business-meaningful action (Add/Edit/Delete/Login/
  Compare/Export/Schedule/etc.). A generic `@app.after_request` hook already auto-logs a "View"
  entry for any successful GET to an authenticated page not otherwise covered — so a brand-new
  route gets baseline audit coverage with zero extra code; only add an explicit call if the action
  needs a specific module/record label beyond the generic View.
- **Background work**: fire a `daemon=True` thread + in-process dedup set for anything slow that
  shouldn't block a request; if more than one server process could plausibly run at once, add a
  Postgres advisory lock too (see AI_PIPELINE.md §6) — an in-process lock alone doesn't protect
  against two OS processes racing.
- **Never trust a single Ollama response shape blindly** — this model (a small local LLM) has a
  demonstrated history of parroting prompt examples, nesting values unexpectedly, and omitting
  fields. Always validate defensively (see `_unwrap_category_score`, `_clamp_pct`,
  `_coerce_str_list`, `_KNOWN_POISONED_EXAMPLE_SCORES`) and have a safe fallback path.

## 3. How to add a new feature — worked patterns

### Add a new column to an existing table
1. Add it to the `CREATE TABLE IF NOT EXISTS` block (for new installs).
2. Add a `_ensure_<column>_col()` function (`ALTER TABLE ... ADD COLUMN IF NOT EXISTS`), called
   once at module level right after its definition (existing deployments).
3. Wire it into every route that reads/writes that table's form (`request.form.get(...)`, the
   INSERT/UPDATE SQL's column list *and* named-param list, and the SELECT the page reads back).
4. Add the form field to the relevant template.
5. If it needs to show up elsewhere (e.g. a dashboard summary table), update that SELECT too.

*Recent worked example*: adding `requirement.onsite_offshore` — schema + migration function +
`ONSITE_OFFSHORE_OPTIONS` constant + Add/Edit form dropdown + Dashboard's Top Open Requirements
SELECT + template columns, all in one self-contained change.

### Add a new route/page
Just add `@app.route(...)` + a view function + a template. No registration step, no blueprint. Add
a nav link in `base.html` (both the sidebar and the mobile menu block) if it should be reachable
from navigation. Pages gate on `_require_permission(key[, write=False])` (checks the current
session role against `ROLE_PERMISSIONS`) — see `app.py`'s RBAC section for the full 5-role matrix
(Admin/Recruiter/Hiring Manager/Interviewer/Viewer-Auditor). Add a new permission by adding one key
to `ROLE_PERMISSIONS`, not by hardcoding a role check inline.

### Add a new AI-scoring signal (like Authenticity Status or JD-Mirroring Risk)
Checklist, in order:
1. **Decide: LLM-derived or pure computation?** If pure computation (regex/string comparison), it
   can be computed fresh on every read and doesn't need `ai_match_cache` storage at all (see
   JD-Mirroring Risk). If LLM-derived, it rides on the *existing* `_llm_judge_match` call — don't
   add a second Ollama call per pair, extend the one prompt instead.
2. If LLM-derived: is it **required** (like the 9 category scores — omission invalidates the whole
   judgment) or **optional** (like each `consistency_checks` key — omission just means "not
   assessed")? Get this right; making an optional field required makes the whole feature far more
   fragile against this small model's occasional omissions. If the signal combines several
   sub-checks into one decision, consider a rule registry + a uniform failure-count threshold (see
   `_AUTHENTICITY_RULES` in AI_PIPELINE.md §5) rather than hardcoding the combination logic — it's
   what makes adding a 9th check later a one-line registry entry instead of a rewrite. Be wary of
   letting any single rule become an absolute override: `_AUTHENTICITY_RULES` originally had one
   (a hard gate), and live data showed it over-triggered badly (see §5's v7→v8 story) — a uniform
   threshold across all rules calibrates far more predictably than a special-cased override.
3. If it changes what gets cached in `extra` (JSONB): **bump `_HYBRID_ALGO_VERSION`** so every
   existing row self-heals. Skip this only if the field is computed fresh on every read (not
   persisted at all).
4. **Validate the prompt change directly against Ollama before wiring it into app.py** — write a
   standalone script calling `_ollama_chat`-equivalent logic with a few realistic (genuine vs.
   fabricated/shallow) test cases and eyeball the actual output. This model has broken prompt
   changes in non-obvious ways multiple times; catching it before it's live in the app is much
   cheaper than debugging it after.
5. Add the field to **every** place a `_hybrid_match`/cache-read result gets assembled:
   `_hybrid_match` (both the success path and the fallback path), `_read_cached_hybrid_match`,
   `_store_hybrid_match` (if persisted), `_rank_top_jd_matches` (both the cached and uncached
   branches — easy to forget one).
6. Add a UI badge (additive — don't restructure the existing card/page markup around it) to
   `profile.html` and `compare_result.html`. If it depends on the async LLM result, wire it into
   the JS `applyResult`/`render()` polling-update function too, not just the initial Jinja render.
7. Write regression tests (see §4) *before* considering it done: label-bucketing thresholds,
   missing-optional-field graceful degradation, cache round-trip, and (if LLM-derived) a mocked
   `_ollama_chat` test proving the feature flows through both the fallback and success paths of
   `_hybrid_match`.
8. Double-check every caller that reads `resume`/`job_description` rows for this feature fetches
   **full rows**, not a narrow `id, updated_at`-only SELECT — a partial row silently breaks any
   new field that reads other columns (this exact bug hit both `_resume_match_fingerprint` and
   `_jd_mirroring_risk` in `compare_ai_status`, on two separate occasions).

## 4. Testing

Five plain-Python test files (no `pytest` installed in this venv) — run directly:

```
python tests/test_match_consistency.py        # scoring model, tiering, caching, determinism
python tests/test_session_and_audit.py         # login/session timeout, audit logging
python tests/test_background_ai_analysis.py    # background jobs, polling endpoints, prompt fixes
python tests/test_pdf_parsing.py               # 2-column PDF layout / row-interleaving
python tests/test_upload_concurrency.py        # /api/parse-resume temp-file race condition
```

Each file is self-contained: a `check(label, condition)` helper appends failures to a module-level
list, `main()` runs every test function and prints a pass/fail summary, `sys.exit(1)` on any
failure. **Run all 5 after any change that touches `app.py`, not just the one you think is
relevant** — several bugs this project hit were regressions in a file that looked unrelated to the
change being made.

Conventions to follow when adding a test:
- Use `FAKE_RESUME_ID`/`FAKE_JD_ID` (large sentinel values, e.g. `999_999_001`/`999_999_002`) for
  synthetic `ai_match_cache` rows so they can never collide with real data — clean up in a
  `finally` block.
- For anything that needs a real resume row (FK constraints, `_rank_top_jd_matches`, etc.), create
  a **disposable** resume via the actual `/edit` POST route (capture the new id from the redirect
  `Location` header), and delete it — and any `ai_match_cache`/uploaded-file rows it created — in
  `finally`. Never reuse or mutate an existing real resume for a test.
- Mock `appmod._ollama_chat` (or `_llm_judge_match`/`parse_resume_with_llm_text` for
  higher-level tests) to keep tests fast and deterministic — don't make real Ollama calls in a test
  unless the test is specifically validating a real-model prompt behavior (and even then, prefer a
  one-off manual script over baking a slow live-model call into the automated suite).
- **Don't rely on `test_client()` + Python `threading` to prove real HTTP concurrency behavior** —
  Flask's test client has its own quirks under multi-threaded use that don't reliably reproduce
  real WSGI-server race conditions. For genuine concurrency bugs, prefer a *deterministic* proof
  (e.g. "every call gets a structurally unique resource, so a collision is impossible regardless of
  timing" — see `test_upload_concurrency.py`) over trying to catch a race in the act.
- The pre-commit hook (`scripts/check_parser.py`) runs automatically whenever `app.py` or itself is
  staged — an undefined-name static check plus a parsing smoke test over several fixture resume
  layouts (including two-column and roles/responsibilities-table fixtures). It runs on every commit
  that touches parsing-adjacent code; don't bypass it with `--no-verify` except in a genuine
  emergency.

## 5. Debugging tips

- **`app.run(..., use_reloader=False)`** — the server does **not** auto-reload on `app.py` edits.
  After any code change, you must manually: find the process on port 5001 (`netstat -ano | grep
  ":5001"` / PowerShell `Get-NetTCPConnection`), kill it (`Stop-Process -Id <pid> -Force`), and
  restart (`nohup .venv/Scripts/python.exe app.py > flask_out.log 2>&1 &` or equivalent). This has
  been the cause of "my fix isn't working" confusion multiple times — always verify the *running*
  process's start time is after your latest edit before concluding a fix doesn't work.
- Server stdout/stderr goes to `flask_out.log` in the project root — `tail` it after any action to
  see what actually happened (Ollama calls, cache hits/misses, background-trigger logging, etc.).
- **Ollama is CPU-only in this environment** — a real LLM inference call saturates the machine's
  CPU and can make *unrelated* lightweight requests feel slow system-wide while it runs. This is a
  hardware/resource constraint, not an application bug; don't chase phantom performance issues that
  are just CPU contention from a concurrent background AI job.
- Quick `ai_match_cache` inspection (useful for verifying scoring/caching behavior without going
  through the UI):
  ```python
  import psycopg2, psycopg2.extras
  conn = psycopg2.connect("<DATABASE_URL>")
  cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
  cur.execute("SELECT resume_id, jd_id, fit_percentage, extra->>'algo_version' AS algo, "
              "computed_at FROM ai_match_cache ORDER BY computed_at DESC LIMIT 20")
  for r in cur.fetchall(): print(r)
  ```
- `audit_log` is the fastest way to reconstruct "what actually happened and when" during a
  live-debugging session (e.g. figuring out that a JD was edited between two test resumes being
  created, which fully explained an apparent "matching is inconsistent" report that turned out to
  be correct behavior against genuinely different JD data).
- When a live-model prompt change is suspected of misbehaving, test it directly against Ollama in a
  standalone script first (bypass the whole app) — much faster to iterate on prompt wording this
  way than restarting the Flask server on every attempt.

## 6. Change impact — things that are easy to get subtly wrong

| If you change... | ...then you must also |
|---|---|
| `_CATEGORY_WEIGHTS`, the `_llm_judge_match` prompt's JSON contract, or how `fit_percentage` is computed | Bump `_HYBRID_ALGO_VERSION` (self-heals every cached row) |
| `_resume_match_fingerprint`'s field list | Understand this changes what counts as "the same resume" for cache-reuse purposes across every JD it's been matched against |
| Anything that reads `resume`/`job_description` for a new AI-related field | Confirm every caller SELECTs the full row, not a narrow `id, updated_at` projection (bit twice already — see AI_PIPELINE.md §3) |
| The `title` field's role in matching | Re-read AI_PIPELINE.md's determinism history first — title is *deliberately* excluded from AI scoring; re-adding it reintroduces a real, previously-reported bug |
| `POSITION_STATUSES` | Remember `REQUIREMENT_STATUSES` is the *same list* (`= POSITION_STATUSES`), not an independent copy |
| Any Save-path parsing behavior | Re-check whether the "skip re-parse if the form already looks filled in" optimization (AI_PIPELINE.md §1) still applies — this was a real, measured performance fix, easy to accidentally undo |
| A temp-file path anywhere in the upload/parse flow | Must be unique per request (`uuid4`), never a fixed filename — the exact bug class that broke concurrent resume previews |
| `app.run()`'s `threaded=True` | Don't remove it — a single slow AI call would otherwise block every other request behind it (this was itself a fix for a real "app looks frozen" bug) |
