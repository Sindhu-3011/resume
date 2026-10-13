# AI_PIPELINE.md — Resume Parsing, Matching & Ollama Internals

Read this file for: resume parser internals, LLM prompt design, JD matching/scoring logic,
caching/determinism, background job architecture. For routes/DB/general architecture, see
**PROJECT_CONTEXT.md**. For coding conventions/testing, see **DEVELOPER_GUIDE.md**.

All Ollama calls go through one function: `_ollama_chat(prompt, *, as_json, num_predict, num_ctx,
timeout)` (app.py ~3823) — always `temperature: 0`. Model: `OLLAMA_TEXT_MODEL` (`.env`, currently
`phi4-mini:latest`, a **small** model — several of the fixes below exist specifically because
small models don't reliably follow complex JSON-format instructions).

## 1. Resume ingestion & parsing

### Entry points
- **Preview parse** (AJAX, fired the instant a user picks a file in Add/Edit Profile):
  `POST /api/parse-resume` → `parse_resume_api()` (app.py ~7962). Saves to a **unique
  per-request temp file** (`tmp-parse-<uuid4hex>.<ext>` in `UPLOAD_FOLDER`) — this used to be a
  *fixed* filename shared by every call, a real concurrency bug (two people previewing resumes at
  the same time could corrupt each other's parse — fixed 2026-08-05). Runs the full LLM parse,
  fills the form fields client-side, does not touch the database.
- **Save-time parse**: inside `edit_resume()` (app.py ~6096), only runs if the form does **not**
  already look filled in (`full_name` + `summary` + (`skills` or `experience`) all present) — if
  the preview already ran, Save skips re-parsing entirely and just persists the form + saves the
  file. This is a **perf-critical fix**: Save used to unconditionally re-run the full OCR+LLM parse
  a second time, taking as long as the preview itself (confirmed live: ~2.5 minutes). If the
  preview genuinely never ran (JS disabled, network error), Save falls back to the fast, non-AI
  `_parse_pdf_quick()` as a gap-filler only.
- **Bulk upload** (`/upload-files`, `/api/process-raw-files`): always uses the fast, non-AI
  `_parse_pdf_quick()` path — running the LLM on every file in a bulk ZIP would be far too slow.

### Text extraction
- PDF: `extract_text_from_pdf()` (~1428) → pymupdf (fastest) → pdfplumber → PyPDF2, in that order.
- DOCX: `extract_text_from_docx()` (~1798) — reads headers, body paragraphs, and tables (joining
  table cells with `\n`, not `" | "` or `""`, specifically to avoid fusing a sidebar cell's last
  line onto the next cell's first line). Falls back to `_extract_canvas_docx()` for floating
  text-box layouts if body extraction is empty.
- **2-column PDF layout handling** — the single most bug-prone part of this pipeline, fixed across
  several sessions (2026-08-03 through 2026-08-05):
  - `_detect_two_col_split()` / `_detect_col_gutter_words()` (~detect column boundary x-coordinate
    via block positions / word-gutter analysis).
  - `_pymupdf_page_text(page)` (~1292): if a column split is detected, classifies each **line**
    (grouped by PyMuPDF's own `(block, line)` id, never by individual word x-position — per-word
    classification let a wrapped line's last word drift into the wrong column and scramble
    sentences) into left/right.
  - For a genuine repeating-row table (e.g. "Client: / Project: / Description:" next to "Role and
    Responsibilities" bullets — common in consulting-style resumes), row boundaries are detected
    via `_split_left_lines_by_row_marker()` — anchored on a repeating label (`Client:`, `Company:`,
    `Employer:`, `Account:`, `Customer:`), **not** vertical gap size (gap size can't distinguish
    "still the same row" from "new row" — the gap between "Client:" and "Project:" in the *same*
    row is often identical to the gap between rows). Falls back to `_chunk_lines_by_gap()`
    (gap-based) for generic 2-column layouts (e.g. a sidebar) that aren't a repeating table.
  - Left/right rows are then **paired by y-range overlap** (greatest overlap wins, not first
    match — a header line merging with the next row's first bullet broke "first match" at a page
    break) so each project's description is followed immediately by *its own* responsibilities,
    not "all descriptions, then all responsibilities" dumped separately.
  - A **non-table 2-column page (e.g. main content + a Skills/Contact sidebar) must NOT be
    row-interleaved** — no repeating marker is found, so it falls back to the safe
    "whole left column, then whole right column" concatenation.
- Contact info (email/phone/name/location) fallback chain when the text layer is unreliable (scans,
  icon-font headers): `_ocr_contact_strip()` (EasyOCR, cached reader via `_get_ocr_reader()`) →
  `_pymupdf_rawdict_scan()` → `_tesseract_inverted_ocr_text()` → `_ollama_vision_contact_scan()` →
  `_pdfplumber_whitetext_scan()` / `_pdfplumber_contact_scan()`.

### Structured field extraction
- `parse_resume_with_llm_text(path)` (~4509): the "Resume Intelligence (AI)" path. Extracts raw
  text, detects section headings (`canonical_section_name()`/`SECTION_ALIASES`), slices verbatim
  content per heading, and asks the LLM (`_ollama_identity_extract`, and a per-field "AI fallback"
  call) to fill in anything the heading-based slicer couldn't find (commonly `skills`, sometimes
  `exp_yrs`). Returns `(parsed_dict, parser_used)`.
- `_parse_pdf_quick(path)` (~6079): fast, rule-based, **no AI** — font-based name extraction +
  `parse_resume_text()`. Used by Save's gap-filler and bulk compare.
- `_detect_years_of_experience()` (~7931): scans summary/experience text for "X+ years" phrases,
  falling back to estimating years from employment date ranges in the experience text
  (`_estimate_years_from_dates`).

## 2. Matching engine — two-stage design

**Stage 1 (cheap, instant, LLM-free)**: pick candidate JDs.
**Stage 2 (expensive, cached, LLM-judged)**: the real, displayed score for whichever JDs get shown.
The app is careful to **never let Stage 1's rough number be mistaken for Stage 2's real one** —
every card/page marks `is_ai_judged: True/False` explicitly.

### Stage 1 — `_quick_estimate_jd_match(resume_dict, jd_dict)` (~6398)
`70% × calculate_match_score() keyword match  +  20% × _holistic_match() whole-doc similarity  +
10% × exact role/title match bonus (_jd_role_matches)`. Used only to rank all Open JDs and pick the
top 3 candidates (`_rank_top_jd_matches`, ~6423) and to decide which JDs the Save-time background
job should compute first (`_top_n_jd_ids_by_estimate`, ~10663). **Never** the number actually shown
for a JD that's already been AI-judged — `_rank_top_jd_matches` always prefers the cached AI result
over the estimate once one exists.
- `calculate_match_score(resume_d, jd_d, include_title=True)` (~9931): literal keyword/requirement
  checklist (`_parse_jd_items`, `_skill_matches`). `include_title=False` (used by the AI path, see
  below) excludes the resume's `title` field from the corpus.
- `_holistic_match()` / `_whole_doc_similarity()` (~9852): cosine-ish term-frequency similarity
  between the two full documents as a whole (not line-by-line) — rescaled `similarity * 220` onto a
  0–100 range. Used as the pure-heuristic **fallback** display too when the LLM is unreachable
  (`_holistic_or_default`).

### Stage 2 — `_hybrid_match(resume_dict, jd_dict)` (~10320) — "the AI Holistic Assessment"
The one function that produces the final, cached, displayed score.

1. `keyword_result = calculate_match_score(..., include_title=False)` — grounding facts for the
   LLM prompt (matched/missing skill lists, keyword %).
2. `_jd_mirroring_risk()` — computed unconditionally, LLM-independent (see §4).
3. `llm_result = _llm_judge_match(...)` — the real AI judgment (see below). If `None` (Ollama
   down/timeout/unparseable), falls back to `_holistic_or_default()` with `is_ai_judged: False`,
   `confidence_label: "Low"` — **never cached** (a fallback caching itself would permanently freeze
   a pair at a low-confidence estimate even after Ollama recovers).
4. If the LLM succeeded: **`fit_percentage` is a deterministic Python-computed weighted sum of the
   9 category scores** — `_CATEGORY_WEIGHTS` (~10013):

   | Category | Weight | | Category | Weight |
   |---|---|---|---|---|
   | skills | 22% | | education | 7% |
   | experience | 15% | | regulatory | 8% |
   | domain | 10% | | tools | 10% |
   | responsibilities | 13% | | soft_skills | 7% |
   | certifications | 8% | | | |

   The LLM **never computes this number itself** — only the 9 independent category scores, summed
   in Python. This is why the score is auditable/reproducible from `category_scores` alone.
5. A separate **confidence** score (not the match %) blends: agreement between keyword-scan % and
   the weighted score (45%), data sufficiency/word-count (20%), the LLM's own self-reported
   confidence (35%) — `_CONFIDENCE_AGREEMENT_WEIGHT` / `_SUFFICIENCY_WEIGHT` / `_SELF_WEIGHT`.

### `_llm_judge_match(resume_dict, jd_dict, keyword_result)` (~10148) — the prompt
Returns `None` (never raises) on any failure — callers always have a safe fallback.

- **Title is deliberately excluded** from both `jd_text`'s resume half and the keyword scan feeding
  it. History: editing only "Role & Group" used to change an *unrelated* JD's cached score (same
  skills/experience, only title changed, score moved 58%→50%) — confirmed live 2026-08-04. Fixed by
  excluding title from every AI-scoring input, **and** switching cache invalidation from
  `resume.updated_at` to a content-only fingerprint (see §3) so a metadata-only edit doesn't even
  trigger a recompute.
- **Category score placeholders are bare `<int>`**, never a filled-in numeric example and never an
  inline descriptive hint next to the key. Two prior formats both failed against this small model:
  - `{"skills": <0-100 int, description>}` (hint inline) → model echoed the *description text* back
    as a nested object, e.g. `{"skills": {"required+preferred skills coverage": 60}}`.
  - A concrete filled-in example (`{"skills": 75, ...}`) → model **copied the example numbers back
    verbatim regardless of resume/JD content** (confirmed live: a strong match and a completely
    unrelated resume against the same JD both returned the *identical* example scores — this is
    why every score in the whole system once converged on ~69%). `_KNOWN_POISONED_EXAMPLE_SCORES`
    is a permanent tripwire: if a response exactly matches that specific 9-value combination, it's
    rejected as a copy, not a genuine judgment.
- `_unwrap_category_score(value)` (~10123): defensive parse for the nested-dict quirk above, kept
  as a second line of defense even with bare-`<int>` placeholders.
- **All 9 categories are strictly required** — if any is missing, `_llm_judge_match` returns `None`
  (triggers the fallback) rather than silently averaging over partial data.
- **`consistency_checks` is NOT required** — optional, supplementary field (see §5). A missing or
  malformed sub-key degrades that single check to "not assessed," never invalidates the whole
  judgment.
- `num_predict=850`, `num_ctx=8192` — bumped from 600 when the 7 consistency-check objects were
  added (previously bumped from 500 for the original evidence-depth fields, now removed).

## 3. Caching & determinism (`ai_match_cache`)

- Keyed `(resume_id, jd_id)`. `_read_cached_hybrid_match()` / `_store_hybrid_match()` /
  `_get_or_compute_hybrid_match()` (~10437, ~10511, ~10556) are the **only** functions that touch
  this table — every reader (Top 3 cards, Compare page, PDF exports) goes through the same read
  path so they can never disagree.
- A cache row counts as fresh only if **all** of: content fingerprint matches (see below), JD's
  `updated_at` unchanged, `extra.algo_version == _HYBRID_ALGO_VERSION`, `extra.is_ai_judged` is
  true (a legacy/pre-fix row without this flag self-heals as a miss).
- **`_resume_match_fingerprint(resume_dict)`** (~10054): SHA-256 over `summary, skills, experience,
  projects, certifications, education, exp_yrs` only — deliberately **excludes** `title`,
  `location`, contact info. Replaced a raw `resume.updated_at` timestamp comparison specifically so
  editing Role & Group (or any other pure-metadata field) can never invalidate an otherwise
  identical cached score.
- **`_HYBRID_ALGO_VERSION`** (~10051, currently `"hybrid-v8-authenticity-calibration"`): bump this
  **any time the scoring formula or prompt contract changes** — it forces every existing cache row
  to self-heal (recompute) the next time it's viewed, with zero manual data migration. Version
  history: `v2` (weighted categories) → `v3` (real category scoring, fixed the poisoned-example
  bug) → `v4` (title excluded) → `v5` (evidence depth added) → `v6` (deterministic
  quant-specificity cap on evidence depth) → `v7` (evidence depth's tri-bucket replaced entirely by
  the binary Authenticity Status system) → `v8` (achievements_measurable demoted from an absolute
  hard gate to one equal vote among 8 + prompt fixed to stop leaking JD-fit into the consistency
  checks, see §5 — v7 over-flagged most resumes). **Note:** a version bump forces one fresh LLM
  call per resume/JD pair with the new prompt text — a small model's specific category-score
  numbers can legitimately shift slightly as a side effect (not because of the new feature itself),
  since any prompt change can nudge other parts of the same response. This is a known, one-time
  cost of shipping a prompt/algo change, not an ongoing instability.
- **`jd_mirroring_*` fields are NOT stored in the cache at all** — computed fresh on every read
  inside `_read_cached_hybrid_match()` (and in `_hybrid_match`'s fallback path). It's pure string
  comparison, cheap enough to redo every time, so it never needs cache-invalidation logic and can
  never go stale. **Every caller of `_read_cached_hybrid_match`/`compare_ai_status` must fetch full
  `resume`/`job_description` rows, not just `id, updated_at`** — a partial JD row silently made
  `jd_mirroring_risk` see an empty JD and always report "no risk" (real bug, fixed 2026-08-06,
  caught because the Top 3 cards and the `/ai-status` endpoint disagreed for the same pair).

## 4. JD-Mirroring Risk (`_jd_mirroring_risk`, ~9771) — authenticity signal #1

Detects a resume whose wording closely echoes the JD's **own phrasing** — the tell of a resume
built by pasting/lightly-rewording a JD rather than describing genuine experience. Pure Python, no
LLM.

- `_extract_meaningful_ngrams(text, n=6)`: 6-word shingles, keeping only ones with ≥3 non-stopword
  tokens (`_DOC_STOPWORDS`).
- Overlap = `|JD shingles ∩ resume shingles| / |JD shingles|` (JD side pulled from
  `skills/requirements/responsibilities/keywords`; resume side from
  `summary/skills/experience/projects`).
- Thresholds (calibrated against real prose, not synthetic keyword lists): independently-written
  and legitimately-paraphrased resumes both scored **0%** against a real JD; a resume with one
  genuinely lifted sentence scored **~11%**; a near-verbatim copy scored **~69%**. →
  `_JD_MIRROR_SOME_THRESHOLD = 8`, `_JD_MIRROR_HIGH_THRESHOLD = 30`, plus a floor of
  `_JD_MIRROR_MIN_MATCHES = 3` matched phrases (avoids a single-coincidence false positive) and
  `_JD_MIRROR_MIN_JD_NGRAMS = 5` (skip JDs too short/generic to judge meaningfully).
- Computed **regardless of AI-judgment status** — available even on the Top 3 card's provisional
  estimate (before any LLM call runs) and even when Ollama is down entirely.

## 5. Authenticity Status (`_evaluate_authenticity`, `_AUTHENTICITY_RULES`, ~9787) — authenticity signal #2

Binary **"Likely Genuine" / "Needs Manual Verification"** signal, replacing an earlier tri-bucket
"Evidence Depth" system (Strong/Moderate/Shallow Evidence, driven by one LLM-judged 0–100 score).
That system collapsed almost every AI-polished resume into "Shallow," losing all ability to
differentiate candidates — the binary redesign is built on 8 explicit, itemized consistency checks
instead of one fuzzy number, and is explicitly **not** a "was this written by AI" judgment.

**Rule registry** (`_AUTHENTICITY_RULES`): each rule is `{key, label, source: "llm"|"deterministic"}`.
This is what makes the system genuinely modular — adding a new check later is one registry entry,
not a rewrite of the decision function (`_evaluate_authenticity`).

| Key | Source |
|---|---|
| `achievements_measurable` | deterministic (`_quantified_specificity_count`) |
| `skills_match_projects` | llm |
| `experience_matches_responsibilities` | llm |
| `timeline_consistent` | llm |
| `designation_matches_experience` | llm |
| `technologies_demonstrated` | llm |
| `projects_realistic` | llm |
| `no_major_contradictions` | llm |

**Decision logic** (`_evaluate_authenticity`) — a **uniform failure-count threshold across all 8**,
no rule overrides the others:
1. **Insufficient signal**: if fewer than `_AUTHENTICITY_MIN_CHECKS_PRESENT = 4` of the 8 checks
   came back usable (LLM omitted/malformed most of them, or the resume is too short to assess
   `achievements_measurable`), fail closed to "Needs Manual Verification" — absence of evidence
   isn't evidence of consistency.
2. **Threshold**: otherwise, `_AUTHENTICITY_FAIL_THRESHOLD = 2` or more of the *present* checks
   failing → "Needs Manual Verification" (the user's own spec: "multiple inconsistencies"). Fewer
   than 2 → "Likely Genuine".
- `reasons_positive`/`reasons_negative`: rule labels for passing rules (positive framing, e.g.
  "Skills are supported by project experience"); the specific concern text for failing rules (the
  LLM's own `reason`, or the deterministic explanation) rather than the label — reusing the
  positive-framed label with a warning icon would misleadingly read as a positive claim.
- `explanation`: a short templated 2–4 sentence summary composed from the reasons — **not** another
  free-text LLM field, so it stays auditable/reproducible from the stored rule results alone, same
  principle as `fit_percentage`.
- `rules`: the full per-rule `{pass, reason, source}` debug trail — stored in the cache `extra` blob
  (`authenticity_rules`) and logged (`logger.info` in `_hybrid_match`), not necessarily all surfaced
  in the UI, but available for debugging why a status was assigned.
- **Deliberately kept separate from `fit_percentage`/`category_scores`** — a flagged resume still
  gets an honest skill-match score; this is a flag for a recruiter to verify in interview, never an
  automatic score penalty.
- If the LLM didn't run at all (Ollama down), `authenticity_status` is `None` ("not assessed"),
  never guessed at — same graceful-degradation pattern the rest of the pipeline uses.

### Why achievements_measurable is one vote, not an absolute override (v7 → v8)

The regression this whole signal exists to catch: a resume that's structurally complete —
real-looking company names, date ranges, and a generic project title for every claimed skill,
grammatically smooth — but contains **zero** numbers or measurable outcomes anywhere. Found live
2026-08-07 (a real candidate resume, "Priya Narayanan"): under the old evidence-depth system, the
LLM alone scored this "Strong Evidence" (70–80/100). Rewording the prompt instructions shifted the
*same* text's score anywhere from 35 to 80 depending on phrasing — proof this particular judgment is
too unstable in a small model to trust on its own, unlike the 9 category scores which the model
handles far more consistently. This motivated `_quantified_specificity_count` as a deterministic,
un-gameable check.

**v7 shipped this as an unconditional hard gate** (any failure force-flagged the resume, regardless
of the other 7) on the theory that a well-fabricated resume would pass every semantic check too. In
practice this over-triggered badly: many genuinely-written resumes in non-metric-driven domains
(compliance/QA/audit work — the work itself is process/audit-oriented, not throughput/metric-driven)
have zero quantified claims and are completely genuine. The hard gate flagged most of them,
recreating the exact failure mode that killed the old tri-bucket Evidence Depth system in the first
place (everything collapsing into one bucket, no differentiation). Re-running the real fabricated
resume live also showed the theoretical worst case never actually happened: the LLM never gave it a
clean bill on every other check — it always independently caught at least one more real issue
alongside the missing metrics.

**v8 fix, two parts:**
1. `achievements_measurable` is now one equal vote among 8 (see decision logic above) — a resume can
   lack quantified achievements alone and still read as genuine, but that plus even one more red flag
   crosses the threshold.
2. **The consistency-check prompt was also leaking JD-fit into the judgment** — reasons like
   *"Experience slightly less than required"* and *"Projects... not specific to IVD"* (the JD's
   domain) were JD-fit gaps, not internal inconsistencies, but were failing `experience_matches_responsibilities`
   /`projects_realistic` anyway. The prompt now explicitly instructs the model: judge the resume's
   coherence with **itself**, never against this specific JD's bar — a JD-fit gap belongs in the
   category scores above, not here.

- `_QUANT_SPECIFICITY_RE`/`_quantified_specificity_count()` (~9769): matches a measurable,
  verifiable claim — a percentage, a team size, a before/after reduction (`"45min to 4min"`), a
  throughput scale (`"2M transactions/day"`), a named improvement with a number attached, etc.
  Calibrated against real text: the fabricated resume above scored **0** matches; a genuinely
  detailed resume scored **3+**.
- Skipped (marked "not assessed", not failed, excluded from both the failure count and the
  minimum-checks-present floor) for a stub-length profile — `_QUANT_SPECIFICITY_MIN_TEXT_LEN = 150`
  chars of `experience + projects` combined — so a resume right after Add, before any real content
  is entered, isn't penalized for being short.
- Mirrors the JD-Mirroring architecture: a deterministic, un-gameable check running *alongside* the
  LLM's semantic judgment, not a replacement for it.

## 6. Background execution & concurrency

- `trigger_background_ai_assessment(resume_id, jd_ids)` (~near `_get_or_compute_hybrid_match`):
  fire-and-forget — starts **one** `daemon=True` thread that processes the given JDs
  **sequentially** (never one thread per JD — concurrent local CPU-only LLM calls contend for the
  same CPU and are slower in aggregate than one at a time). In-process dedup via
  `_ai_assessment_in_flight` (a `set()` of `(resume_id, jd_id)` guarded by a `threading.Lock`).
- `_get_or_compute_hybrid_match()` additionally takes a **Postgres advisory lock**
  (`pg_advisory_lock(resume_id, jd_id)`) — not just an in-process lock — because this dev setup can
  end up with more than one server process bound to the same port (a stale process left running
  alongside a freshly restarted one). Session-scoped: releases automatically if the connection
  dies, so a crashed request can never leave a pair permanently locked.
- Two polling endpoints, **do not confuse them**:
  - `/api/compare/<r>/<j>/ai-status` (`compare_ai_status`): **read-only**, never computes, never
    blocks — what the frontend polls every few seconds.
  - `/api/compare/<r>/<j>/ai-assessment` (`compare_ai_assessment`): computes on cache miss
    (blocking, via `_get_or_compute_hybrid_match`) — still exists but current templates all use the
    polling endpoint instead; largely superseded.
- Frontend polling (`profile.html`, `compare_result.html`): `POLL_INTERVAL_MS = 4000`,
  `MAX_WAIT_MS = 360000` (6 minutes — deliberately exceeds `OLLAMA_TEXT_TIMEOUT`). Cards re-sort
  themselves by current score after each poll cycle (`resortCards()` in profile.html) so the
  visible order converges to what a full reload would show, without needing one.

## 7. Legacy / dead code — do not build on these

- `get_embedding()`, `rebuild_role_match_cache()`, `role_match_cache` table: an embeddings-based
  matching approach, fully superseded by the keyword+LLM hybrid system above. **No caller anywhere
  in the codebase invokes `rebuild_role_match_cache`** — confirmed by direct search. Removing the
  (already-removed) call sites from the Save path was what fixed an old "Save is slow" bug.
