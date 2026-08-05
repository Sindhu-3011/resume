"""Regression test for the 2-column PDF text-scrambling bug found live on
2026-08-03 while investigating incomplete resume parsing for a real user
(Arun Sivakumar's resume: a genuine 2-column "Project Details / Role and
Responsibilities" table).

Plain-Python, no pytest dependency — run directly:

    python tests/test_pdf_parsing.py

Root cause: _pymupdf_page_text classified each WORD independently by
comparing its own x0 against the detected column split point. A left-column
paragraph line that happens to wrap close to the column's natural right
margin can have its LAST word drift past the split point even though the
rest of the line is clearly left-column — that one trailing word then got
shaved off into the right-column bucket, scrambling sentences throughout
the document (confirmed live: "...providing greater visibility and
control" lost the word "control" this way, repeated across every project
description in the real resume).

Fix: classify by PyMuPDF's own (block, line) grouping — the line unit the
PDF's content stream actually rendered — using that line's own leftmost
word to decide the column, never any individual word's position. Two
column headers that legitimately share the same y (e.g. "Project Details"
/ "Role and Responsibilities") are already separate (block, line) pairs, so
they still split apart correctly.

This test does not depend on any real PDF file — it builds synthetic
PyMuPDF word tuples directly and monkeypatches column-split detection to a
known value, isolating the classification logic from the (separately
tested, real-file-verified) split-point detection heuristics.
"""

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


class _FakePage:
    """Minimal stand-in for a pymupdf Page — _pymupdf_page_text only ever
    calls page.get_text('words') on it once _detect_two_col_split is
    monkeypatched to skip real column detection."""

    def __init__(self, words):
        self._words = words

    def get_text(self, mode):
        assert mode == "words"
        return self._words


def test_line_that_wraps_past_split_point_stays_intact_in_its_own_column():
    # (x0, y0, x1, y1, text, block_no, line_no, word_no) — the real pymupdf
    # "words" tuple shape. Left column line 2's last word ("control") is
    # deliberately placed at x0=305, past split_x=300 — reproducing the
    # exact failure mode (a wrapped line's trailing word drifting past the
    # column boundary) without needing a real PDF file.
    words = [
        # Left column, line 0 (block=0, line=0): "Project Description: text"
        (29.0, 79.1, 60.0, 88.0, "Project", 0, 0, 0),
        (65.0, 79.1, 130.0, 88.0, "Description:", 0, 0, 1),
        (135.0, 79.1, 170.0, 88.0, "widget", 0, 0, 2),
        # Left column, line 1 (block=0, line=1): wraps close to the margin —
        # "control" (the last word) sits past split_x=300.
        (29.0, 94.1, 70.0, 103.0, "and", 0, 1, 0),
        (75.0, 94.1, 140.0, 103.0, "visibility", 0, 1, 1),
        (145.0, 94.1, 160.0, 103.0, "and", 0, 1, 2),
        (305.0, 94.1, 340.0, 103.0, "control", 0, 1, 3),
        # Right column, line 0 (block=1, line=0): bullet at an overlapping y
        (363.0, 89.1, 375.0, 98.0, "-", 1, 0, 0),
        (378.0, 89.1, 440.0, 98.0, "Validation", 1, 0, 1),
    ]

    real_split = appmod._detect_two_col_split
    try:
        appmod._detect_two_col_split = lambda page: 300.0
        result = appmod._pymupdf_page_text(_FakePage(words))
    finally:
        appmod._detect_two_col_split = real_split

    check(
        "the wrapped line's trailing word stays attached to its own sentence",
        "and control" in result,
    )
    check(
        "the trailing word was not shaved off into the right (bullet) column",
        "Validation" not in result.split("\n")[0] if "\n" in result else True,
    )
    # Locate the two column outputs by which one contains "Validation" —
    # order isn't the point here, correctness of grouping is.
    parts = result.split("\n")
    right_part = next((p for p in parts if "Validation" in p), "")
    check("the right-column bullet text has no stray 'control' bled into it", "control" not in right_part)


def test_same_y_column_headers_still_split_correctly():
    # Two headers sharing the exact same y (a real, common resume-table
    # pattern: "Project Details" | "Role and Responsibilities" side by
    # side) must still end up in their respective columns — distinguished
    # by (block, line), not by y-proximity.
    words = [
        (149.0, 133.0, 220.0, 142.0, "Project", 3, 0, 0),
        (225.0, 133.0, 260.0, 142.0, "Details", 3, 0, 1),
        (393.0, 133.0, 420.0, 142.0, "Role", 3, 1, 0),
        (425.0, 133.0, 460.0, 142.0, "and", 3, 1, 1),
        (465.0, 133.0, 520.0, 142.0, "Responsibilities", 3, 1, 2),
    ]

    real_split = appmod._detect_two_col_split
    try:
        appmod._detect_two_col_split = lambda page: 300.0
        result = appmod._pymupdf_page_text(_FakePage(words))
    finally:
        appmod._detect_two_col_split = real_split

    parts = result.split("\n")
    left_part = next((p for p in parts if "Project" in p), "")
    right_part = next((p for p in parts if "Responsibilities" in p), "")
    check("left header lands in the left column", "Details" in left_part)
    check("right header lands in the right column, not merged with the left", "Role and Responsibilities" in right_part)
    check("left column doesn't absorb the right header", "Responsibilities" not in left_part)


def test_each_project_paired_immediately_with_its_own_responsibilities():
    """Regression test for the bug found live on 2026-08-05: after fixing
    the column-scrambling bug above, the left and right columns were still
    each emitted as one giant blob — ALL project descriptions, then
    (separately, at the very end) ALL role/responsibility bullets — so a
    project's own responsibilities appeared to be missing until the very
    last section of the field. Root cause: within the left column, the
    vertical gap between "Client:" and "Project N:" (still the SAME row)
    is often identical to the gap between one row's last line and the
    next row's "Client:" line — both are just one paragraph-height apart —
    so pure gap-based row detection can't tell "still this row" from "new
    row" apart. Fixed by anchoring left-column row boundaries on the
    repeating "Client:" label instead, then pairing each row with its
    right-column content by vertical position (see
    _split_left_lines_by_row_marker / the overlap-assignment in
    _pymupdf_page_text). Expected: Project 1's own bullets appear right
    after Project 1's description, before Project 2 starts at all — not
    after every project description in the document.
    """
    def word(x0, y0, x1, y1, text, block, line, idx):
        return (x0, y0, x1, y1, text, block, line, idx)

    words = [
        # Left column — two "Client: ... / Project N: ... / Description"
        # rows, each internal line ~23pt apart (matches the real document:
        # same spacing whether it's still the current row or a new one).
        word(29, 100, 120, 109, "Client:", 0, 0, 0), word(125, 100, 200, 109, "Alpha", 0, 0, 1),
        word(205, 100, 240, 109, "Corp", 0, 0, 2),
        word(29, 123, 200, 132, "Project", 1, 0, 0), word(205, 123, 240, 132, "1:", 1, 0, 1),
        word(29, 146, 260, 155, "Description:", 2, 0, 0), word(265, 146, 300, 155, "Widget.", 2, 0, 1),

        word(29, 169, 120, 178, "Client:", 3, 0, 0), word(125, 169, 200, 178, "Beta", 3, 0, 1),
        word(205, 169, 240, 178, "Corp", 3, 0, 2),
        word(29, 192, 200, 201, "Project", 4, 0, 0), word(205, 192, 240, 201, "2:", 4, 0, 1),
        word(29, 215, 260, 224, "Description:", 5, 0, 0), word(265, 215, 300, 224, "Gadget.", 5, 0, 1),

        # Right column — column header merged close to row 1's own bullets
        # (mirrors the real page-break case that broke "first overlapping
        # chunk wins" and required picking the GREATEST overlap instead),
        # then row 2's bullets separated by a clear gap.
        word(363, 100, 550, 109, "Role", 6, 0, 0), word(555, 100, 600, 109, "Resp.", 6, 0, 1),
        word(363, 130, 500, 139, "Led", 6, 1, 0), word(505, 130, 560, 139, "widget", 6, 1, 1), word(565, 130, 610, 139, "team", 6, 1, 2),
        word(363, 141, 500, 150, "Delivered", 6, 2, 0), word(505, 141, 560, 150, "on", 6, 2, 1), word(565, 141, 610, 150, "time", 6, 2, 2),

        word(363, 200, 500, 209, "Led", 7, 0, 0), word(505, 200, 560, 209, "gadget", 7, 0, 1), word(565, 200, 610, 209, "team", 7, 0, 2),
        word(363, 211, 550, 220, "Migrated", 7, 1, 0), word(555, 211, 610, 220, "data", 7, 1, 1),
    ]

    real_split = appmod._detect_two_col_split
    try:
        appmod._detect_two_col_split = lambda page: 300.0
        result = appmod._pymupdf_page_text(_FakePage(words))
    finally:
        appmod._detect_two_col_split = real_split

    idx_widget_desc = result.find("Widget")
    idx_widget_bullet = result.find("widget team")
    idx_client_beta = result.find("Beta")
    idx_gadget_bullet = result.find("gadget team")

    check("all expected fragments are present", -1 not in
          (idx_widget_desc, idx_widget_bullet, idx_client_beta, idx_gadget_bullet))
    check(
        "project 1's own bullets appear right after project 1's description, "
        "before project 2 starts",
        idx_widget_desc < idx_widget_bullet < idx_client_beta,
    )
    check(
        "project 2's bullets appear after project 2 starts, not bundled at the top",
        idx_client_beta < idx_gadget_bullet,
    )


def test_generic_two_column_layout_is_not_row_interleaved():
    """A generic 2-column page (main content + a Skills/Contact sidebar
    that runs independently down the whole page) is NOT a row-paired
    table — there's no "Client:"-style repeating marker, so row-by-row
    interleaving must never kick in for it. Confirmed live against a real
    sidebar-style resume: applying gap-based chunking + y-overlap pairing
    unconditionally (instead of gating it behind a confirmed marker match)
    scrambled sidebar paragraphs into the middle of unrelated main-column
    content. The safe, pre-existing behavior — whole left column, then
    whole right column — must be preserved whenever no row marker is
    found.
    """
    def word(x0, y0, x1, y1, text, block, line, idx):
        return (x0, y0, x1, y1, text, block, line, idx)

    words = [
        # Left column: ordinary prose paragraphs, no "Client:"-style marker.
        word(29, 100, 120, 109, "About", 0, 0, 0), word(125, 100, 200, 109, "Me", 0, 0, 1),
        word(29, 123, 250, 132, "Mechanical", 1, 0, 0), word(255, 123, 300, 132, "engineer.", 1, 0, 1),
        word(29, 200, 150, 209, "Work", 2, 0, 0), word(155, 200, 250, 209, "Experience", 2, 0, 1),
        word(29, 223, 250, 232, "Best", 3, 0, 0), word(255, 223, 300, 232, "Enterprises.", 3, 0, 1),

        # Right column: an independent sidebar (Skills), unrelated row-for-
        # row to the left column's paragraphs.
        word(363, 100, 450, 109, "Skills", 4, 0, 0),
        word(363, 123, 450, 132, "AutoCAD", 5, 0, 0),
        word(363, 146, 450, 155, "CREO", 6, 0, 0),
    ]

    real_split = appmod._detect_two_col_split
    try:
        appmod._detect_two_col_split = lambda page: 300.0
        result = appmod._pymupdf_page_text(_FakePage(words))
    finally:
        appmod._detect_two_col_split = real_split

    idx_about = result.find("About")
    idx_work = result.find("Work Experience")
    idx_best = result.find("Best Enterprises")
    idx_skills = result.find("Skills")

    check("all expected fragments are present", -1 not in (idx_about, idx_work, idx_best, idx_skills))
    check(
        "the whole left column stays together — Best Enterprises still "
        "comes before the sidebar, not interrupted by it",
        idx_about < idx_work < idx_best < idx_skills,
    )


if __name__ == "__main__":
    tests = [
        test_line_that_wraps_past_split_point_stays_intact_in_its_own_column,
        test_same_y_column_headers_still_split_correctly,
        test_each_project_paired_immediately_with_its_own_responsibilities,
        test_generic_two_column_layout_is_not_row_interleaved,
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
    else:
        print("All PDF-parsing checks passed.")
