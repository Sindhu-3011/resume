# PROJECT_CONTEXT.md — Resume Profile Pre-Screening App

Read this file for: overall architecture, routes/modules, database schema, workflows, config.
For resume parsing / matching / Ollama internals, see **AI_PIPELINE.md** instead.
For coding conventions, testing, and "how do I..." tasks, see **DEVELOPER_GUIDE.md** instead.

## 1. What this app is

An internal Flask app for a staffing/recruiting team: upload candidate resumes, upload job
descriptions (JDs), and get an AI-assisted match score between them. Tracks the hiring pipeline
(requirements, interviews, candidate status) and keeps a full audit trail of who did what.

## 2. Tech stack & architecture

- **Single-file Flask monolith**: `app.py` (~12,100 lines). No blueprints, no ORM.
- **Database**: PostgreSQL, accessed via raw `psycopg2` + `%s`-parameterized SQL. No SQLAlchemy.
  - `db_conn()` (context manager) opens a **fresh connection per call** — no pooling. Wraps it in
    `_PgConn`, a thin adapter giving `conn.execute(sql, params)` a sqlite3-style API returning a
    `RealDictCursor` (rows behave like dicts).
- **AI**: local **Ollama** server (`http://127.0.0.1:11434` by default), model configured via
  `OLLAMA_TEXT_MODEL` (currently `phi4-mini:latest` in `.env`). No external AI API calls — fully
  local/offline. Every Ollama call goes through the single `_ollama_chat()` helper, always with
  `temperature: 0` (deterministic decoding).
- **Server**: Werkzeug dev server, `threaded=True` (concurrent requests — needed since a real LLM
  judgment can take 30–90s+ on CPU-only hardware), `use_reloader=False`, port `5001`.
  **`use_reloader=False` means the server does NOT pick up app.py edits automatically — it must be
  manually killed and restarted after every code change.** See DEVELOPER_GUIDE.md.
- **Frontend**: server-rendered Jinja2 templates + vanilla JS (no build step, no framework). AJAX
  polling (not websockets) for anything that depends on a slow background AI computation.
- **Auth**: session-based login (`app_user` table), a single `@app.before_request` gate
  (`_require_login`) enforces login + a sliding-window inactivity timeout on every route.

## 3. Environment / config (`.env`)

| Variable | Purpose | Default |
|---|---|---|
| `DATABASE_URL` | Postgres connection string | — |
| `SECRET_KEY` | Flask session signing key | `"resume-profile-secret-key"` |
| `SESSION_TIMEOUT_MINUTES` | Inactivity logout window | `15` |
| `OLLAMA_BASE_URL` | Ollama server base URL | `http://localhost:11434` |
| `OLLAMA_TEXT_MODEL` | Model name for all LLM calls | `llama3.2:latest` (actual `.env`: `phi4-mini:latest`) |
| `OLLAMA_TEXT_TIMEOUT` | Per-call timeout (seconds) | `60` (actual `.env`: `300`) |
| `OLLAMA_BUDGET_SECS` | (secondary budget knob, some callers) | `10` |

## 4. Database schema

Base tables (`resume`, `resume_skill`) are created by **`create_tables.sql`** (run manually via
DBeaver/psql, not by the app). Everything else is created lazily by `ensure_X_table(conn)`
functions called from within routes, or via one-time `_ensure_X_col()` migration functions that
run at module import time (`ALTER TABLE ... ADD COLUMN IF NOT EXISTS`).

### `resume` (base: `create_tables.sql`, then extended by migrations)
`id, full_name, title, email, phone, linkedin, location, summary, skills, experience, education,
certifications, projects, slug, resume_file, created_at, updated_at` — then added later:
`status` (default `'New'`), `interview_status` (default `'To Be Interviewed'`), `l1_comments`,
`l2_comments`, `exp_yrs`, `department`.

- `title` is the **"Role & Group"** field in the UI — pure organizational metadata. It's
  deliberately **excluded** from AI scoring inputs (see AI_PIPELINE.md — this was the source of a
  determinism bug).
- `interview_status` (11-value pipeline stage: To Be Interviewed → ... → Joined Vaisesika, etc.,
  see `INTERVIEW_STATUSES`) is **not the same thing** as `interview_schedule.status` (4-value:
  Scheduled/Completed/Pending/Cancelled) — two genuinely different concepts, don't conflate them.

### `resume_skill`
`id, resume_id (FK, CASCADE), skill` — normalized one-row-per-skill table, rebuilt on every save
via `sync_skills()`. Not currently read by the matching engine (matching reads `resume.skills` as
one text blob) — exists for potential future skill-level querying.

### `job_description`
`id, title, role, category, responsibilities, requirements, skills, keywords, jd_file,
position_status (default 'Open'), created_at, updated_at`
- `POSITION_STATUSES = ["Open", "Fulfilled", "On Hold", "Closed"]` — single source of truth, also
  reused by `requirement.status` (see `REQUIREMENT_STATUSES = POSITION_STATUSES`).
- Matching **only ever considers `position_status = 'Open'`** JDs — filtered before scoring, not
  after, so a Fulfilled/Closed JD can never appear in Top 3 Matching Roles.

### `requirement`
`id, requirement_code, requirement_name, client, division, num_requirement, status (default
'Open'), profiles_shared, interviewed, offered, onsite_offshore (default ''), created_at,
updated_at`. `onsite_offshore` values: `ONSITE_OFFSHORE_OPTIONS = ["Onsite", "Offshore", "Hybrid"]`.

### `interview_schedule`
`id, resume_id (FK, CASCADE), position, interviewer, interview_date, status (default
'Scheduled'), created_at`. Status values: `INTERVIEW_SCHEDULE_STATUSES = ["Scheduled", "Completed",
"Pending", "Cancelled"]`. Candidate (`resume_id`) is immutable once created — editing only changes
status/date/interviewer.

### `ai_match_cache`
`resume_id, jd_id (composite PK), fit_percentage, verdict, rationale, extra (JSONB), 
resume_updated_at, jd_updated_at, computed_at`. The single source of truth for every AI-judged
resume/JD assessment — see AI_PIPELINE.md for the full caching model (this is the most
architecturally significant table in the app).

### `audit_log`
`id, user_id, username, role, created_at, module, action, record_id, record_label, ip_address,
user_agent, status (default 'Success'), details`. Indexed on `created_at DESC`, `user_id`,
`module`. Written by the reusable `log_audit()` service — see DEVELOPER_GUIDE.md.

### `app_user`
`id, username (unique), password_hash, full_name, email, role (default 'user'), is_active,
created_at, updated_at`. **`USER_ROLES = ["admin", "user"]`** — only two roles exist. (No
Viewer/Auditor role exists anywhere in the app, despite that being a plausible-sounding
requirement seen in past requests — confirmed by direct code search.)

### Legacy / unused tables
- **`role_match_cache`**: an embeddings-based match-cache table. **Nothing reads from it.**
  `rebuild_role_match_cache()` (the only writer) is also never called from anywhere — this was the
  original matching approach, superseded by the keyword+LLM hybrid system, and calling it used to
  make Save slow (fixed by removing the call sites, not the dead function itself). Safe to ignore;
  do not build new features on it.

## 5. Modules & routes

### Auth / session (`/login`, `/logout`, `/api/session/keepalive`)
Login sets `session['user_id']`, `role`, `last_active`. `_require_login` (global
`before_request`) enforces auth + the sliding inactivity timeout on every request except
`_LOGIN_EXEMPT_ENDPOINTS = {"login", "static"}`. Client-side warning JS: `static/session-timeout.js`.

### Dashboard (`/dashboard`)
KPI cards, department chart, candidate-status donut, interview-status donut, **Top Open
Requirements** table (Open-status requirements only — includes Onsite/Offshore, Interviewed,
Offered columns), **Recent Interviews** table (editable — see below), Candidate Details table
(search/filter/select/refresh over all resumes).

### Resume Profiles (`/edit`, `/edit/<id>`, `/profile/<id>`, `/profiles`, `/profile/<id>/delete`, ...)
- `/edit` (no id) = Add Profile; `/edit/<id>` = Edit Profile. Same route/template (`edit.html`)
  handles both (POST branches on whether `resume_id` is present).
- File upload (`resume_file` field) triggers **client-side** AJAX parsing on file-select
  (`/api/parse-resume`) which auto-fills the form fields for the user to review — the expensive
  OCR/LLM parse happens **once**, here. The final Save POST does **not** re-parse if the form
  already looks filled in (see AI_PIPELINE.md — this was a real perf bug, now fixed).
- On Save, a background thread computes AI match scores for the resume's estimated top-3 JDs —
  Save itself returns immediately (redirects to `/profile/<id>`), the AI scores populate a few
  seconds to a couple of minutes later via polling. See AI_PIPELINE.md §5.
- `/profile/<id>` (`profile.html`) is the "Top 3 Matching Roles" page — also shows the resume's raw
  parsed sections (Summary, Skills, Experience, Education, Certifications, Project Details) and an
  L1/L2 reviewer-comments box (`/profile/<id>/save-comments`).
- Export routes: `/profile/<id>/download-pdf` (raw resume), `/profile/<id>/export-rich-pdf` and
  `/profile/<id>/export-top-matches-pdf` (match-assessment PDFs), `/profile/<id>/export-top-matches`
  (JSON).

### JD Management (`/jd-management`, `/jd/add`, `/jd/<id>`, `/jd/<id>/edit`, `/jd/<id>/delete`, ...)
CRUD for job descriptions. `/jd/<id>/top-matches` shows the reverse view — top matching resumes
for one JD.

### Requirement Management (`/requirement-management`, `/requirement/add`, `/requirement/<id>`, ...)
CRUD for client staffing requirements (separate concept from JDs — a requirement tracks headcount/
client/division/pipeline metrics; a JD is the actual job posting text used for AI matching).

### Compare Resume (`/compare/<resume_id>`, `/compare/<resume_id>/<jd_id>`, ...)
Pick-a-JD flow, then the full AI Holistic Assessment page (category breakdown, Authenticity Status,
JD-Mirroring Risk — see AI_PIPELINE.md) for one resume/JD pair. Async-polled the same way as the
Top 3 cards.

### Interview scheduling (Dashboard's Recent Interviews widget)
`/api/schedule-interview` (create), `/api/interview/<id>/update` (edit status/date/interviewer —
candidate is immutable). Both are logged to the audit trail (`Interviews` module).

### User Management (`/users`, `/users/add`, `/users/<id>/edit`, `/users/<id>/delete`) — admin only
Two roles: `admin`, `user`.

### Audit Trail (`/audit-trail`, `/audit-trail/export.csv`) — admin only
Search/filter/sort/paginate/CSV-export over `audit_log`. See DEVELOPER_GUIDE.md for how logging
itself works (explicit `log_audit()` calls + a generic auto-logging hook for any new route).

### Bulk / misc
`/upload-files` (bulk multi-resume upload + ZIP extraction), `/api/bulk-compare` +
`/api/export-compare-pdf` (bulk resume-vs-JD comparison), `/groups` (role/category browsing),
`/api/candidates`, `/candidate-status`, `/candidate-interview-status` (dashboard table inline
edits), `/api/check-duplicate` (email/phone dedup check used by both Add-profile and inline blur
checks).

## 6. Templates inventory

| Template | Route(s) |
|---|---|
| `base.html` | Layout shell (sidebar nav, session-timeout script injection) for every page |
| `login.html` | `/login` |
| `dashboard.html` | `/dashboard` |
| `edit.html` | `/edit`, `/edit/<id>` |
| `profile.html` | `/profile/<id>` (Top 3 Matching Roles + raw resume sections) |
| `profile_list.html` | `/profiles` |
| `compare_select.html` | `/compare/<resume_id>` |
| `compare_result.html` | `/compare/<resume_id>/<jd_id>` |
| `jd_management.html`, `jd_form.html`, `jd_detail.html`, `jd_top_matches.html` | JD Management |
| `requirement_management.html`, `requirement_form.html`, `requirement_detail.html` | Requirement Management |
| `user_management.html`, `user_form.html` | User Management |
| `audit_trail.html` | `/audit-trail` |
| `groups.html` | `/groups`, `/groups/<role>` |
| `upload_files.html`, `_upload_files_body.html` | `/upload-files` (bulk upload) |

## 7. Cross-references

- Resume parsing internals, the AI matching/scoring pipeline, Authenticity Status, JD-Mirroring Risk,
  caching/determinism, background job architecture → **AI_PIPELINE.md**.
- Coding conventions, test suite, debugging, "how do I add X" → **DEVELOPER_GUIDE.md**.
