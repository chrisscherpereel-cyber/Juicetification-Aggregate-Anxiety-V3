"""Fixes from the first student review of V3: helper columns, circular references, formulas
that point at the wrong column, the lower-cost shortcut, spare columns and the receipt."""

import io

import pytest
from openpyxl import load_workbook

from aggplan import checker, guide, layout as L, recognize, report, workbook
from aggplan.checker import Cell
from tests.helpers import SID, book_from, fixture, scn7

HYB = L.SHEET_HYBRID


@pytest.fixture(scope="module")
def key_book():
    book, err = checker.read_book(fixture("key_seed7.xlsx"))
    assert err is None
    return book


def _cell(res, ref):
    return next(c for c in res.cells if c.ref == ref)


# ---- helper columns --------------------------------------------------------------------
@pytest.mark.parametrize("heading", ["original workers", "Helper: leftover with one worker fewer",
                                     "Workers (helper)", "scratch production", "Note"])
def test_headings_marked_as_helpers_are_not_read_as_plan_columns(heading):
    assert recognize.classify(heading, "hybrid") == recognize.HELPER


def test_second_workers_column_does_not_block_submitting(key_book):
    # the reviewer's "original workers" column, but with a heading read as Workers
    b = book_from(key_book, {(HYB, "W1"): Cell("Workers 2", None),
                             **{(HYB, f"W{r}"): Cell(5, None) for r in range(2, 14)}})
    res = checker.check_book(b, scn7(), "hybrid", SID)
    assert res.complete
    assert any(n.kind == "duplicate" and n.letters == ("E", "W") for n in res.notes)
    assert _cell(res, "W2").status == "helper"


# ---- circular references ---------------------------------------------------------------
def test_formula_refs_skip_other_sheets_and_text_and_expand_ranges():
    refs = checker.formula_refs("=SUM(K2:M2)+'Chase Plan'!P14+Scenario!B5+COUNTIF(B2:B3,\">0\")")
    assert (11, 2) in refs and (12, 2) in refs and (13, 2) in refs
    assert (16, 14) not in refs and (2, 5) not in refs
    assert (2, 2) in refs and (2, 3) in refs


def test_decision_that_uses_its_own_results_is_a_circular_reference(key_book):
    # Workers in March computed from Total Production in March, which depends on Workers'
    # regular production... here directly: E4 -> J4 -> G4 (a formula using E4)
    b = book_from(key_book, {(HYB, "E4"): Cell(0, "=J4/1000"),
                             (HYB, "G4"): Cell(0, "=E4*1000")})
    res = checker.check_book(b, scn7(), "hybrid", SID)
    assert not res.complete
    assert any("Circular reference" in m and "E4" in m for m in res.messages)
    assert {_cell(res, r).status for r in ("E4", "G4", "J4")} == {"error"}


def test_no_circular_reference_in_the_answer_key(key_book):
    for sheet in (L.SHEET_CHASE, L.SHEET_LEVEL, HYB, L.SHEET_SUMMARY):
        assert checker.circular_groups(key_book, sheet) == []


# ---- formulas pointing at the wrong column ------------------------------------------------
def test_hires_reading_the_overtime_column_is_wrong_even_when_the_number_fits(key_book):
    # the reviewer's K2 / L3: overtime is 0, so the values happen to be right
    b = book_from(key_book, {(HYB, "K3"): Cell(0, "=MAX(0,H3-E2)"),
                             (HYB, "L3"): Cell(0, "=MAX(0,H2-E3)")})
    res = checker.check_book(b, scn7(), "hybrid", SID)
    for ref in ("K3", "L3"):
        c = _cell(res, ref)
        assert c.status == "wrong" and "Overtime Production" in c.diag
    assert not res.complete


def test_legitimate_variations_are_not_flagged(key_book):
    b = book_from(key_book, {
        (HYB, "K3"): Cell(0, "=MAX(0,(F3-F2)/$C$17)"),        # via Regular Capacity
        (HYB, "N3"): Cell(key_book.cell(HYB, "N3").value, "=MAX(0,B3+D3-C3-J3)+0*M3"),
    })
    res = checker.check_book(b, scn7(), "hybrid", SID)
    assert _cell(res, "K3").status == "ok" and _cell(res, "N3").status == "ok"


# ---- Summary: the lower-cost answer must follow the totals ---------------------------------
def test_lower_cost_cell_pointing_at_a_heading_is_not_accepted(key_book):
    b = book_from(key_book, {(L.SHEET_SUMMARY, L.LOWER_COST_CELL): Cell("Chase", "=B3")})
    res = checker.check_book(b, scn7(), "summary", SID)
    c = _cell(res, L.LOWER_COST_CELL)
    assert c.status == "wrong" and "TOTAL PLAN COST" in c.diag


def test_lower_cost_formula_comparing_totals_is_accepted(key_book):
    res = checker.check_book(key_book, scn7(), "summary", SID)
    assert _cell(res, L.LOWER_COST_CELL).status == "ok"


# ---- the workbook ------------------------------------------------------------------------
def test_plan_sheet_marks_spare_columns_and_says_so():
    wb = load_workbook(io.BytesIO(workbook.build_workbook(scn7(), student_name="x",
                                                          student_id=SID)))
    ws = wb[L.SHEET_CHASE]
    last_main = L.FIRST_STUDENT_COL + len(L.REQUIRED["chase"]) + len(L.OPTIONAL["chase"]) - 1
    main, spare = ws.cell(2, last_main).fill.fgColor.rgb, ws.cell(2, last_main + 1).fill.fgColor.rgb
    assert main != spare
    notes = " ".join(str(c.value) for row in ws.iter_rows(min_row=15, max_row=25, max_col=1)
                     for c in row if c.value)
    assert "spare room" in notes
    help_text = " ".join(str(c.value) for row in wb[L.SHEET_HELP].iter_rows() for c in row if c.value)
    assert "ROUNDUP(need / per_unit, 0)" in help_text and "Circular reference" in help_text
    summ = " ".join(str(c.value) for row in wb[L.SHEET_SUMMARY].iter_rows() for c in row if c.value)
    assert 'COUNTIF(range, ">0")' in summ


def test_hybrid_start_adds_fixed_helper_columns_the_checker_ignores():
    scn = scn7()
    raw = workbook.build_workbook(scn, student_name="x", student_id=SID)
    out = workbook.add_hybrid_start(raw, scn, guide.starting_decision(scn, "chase"))
    ws = load_workbook(io.BytesIO(out))[HYB]
    heads = {ws.cell(1, j).value: j for j in range(1, 31) if ws.cell(1, j).value}
    assert "Helper: starting workers" in heads
    assert "Helper: leftover with one worker fewer" in heads
    j = heads["Helper: starting workers"]
    assert isinstance(ws.cell(2, j).value, (int, float))           # a fixed number
    book, _ = checker.read_book(out)
    assert checker.circular_groups(book, HYB) == []
    res = checker.check_book(book, scn, "hybrid", SID)
    assert not any(n.kind in ("duplicate", "missing") for n in res.notes)


# ---- receipt -----------------------------------------------------------------------------
def test_receipt_shows_first_and_last_attempt_and_every_upload():
    pdf, _ = report.build_report(
        scn7(), student_name="Ann", student_id=SID, identity_label="x",
        pieces=[{"title": "Level plan", "status": "Submitted", "attempts": 3,
                 "first": "0 / 128 correct", "last": "128 / 128 correct",
                 "history": "0/128 → 64/128 → 128/128", "help": 2,
                 "submitted_at": "t", "file": "f.xlsx", "sha": "a" * 64}],
        final=True, heading="Receipt")
    assert pdf[:4] == b"%PDF"


# ---- guided hybrid: an unchanged starting plan can't be submitted --------------------------
def _app_on_hybrid(plan):
    import os
    from streamlit.testing.v1 import AppTest
    from aggplan import engine
    app = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")
    at = AppTest.from_file(app, default_timeout=60)
    at.query_params["seed"] = "7"
    at.run()
    at.text_input[0].input("Ann")
    at.text_input[1].input(SID)
    next(b for b in at.button if b.label == "Start").click()
    at.run()
    s = scn7()
    ps = at.session_state["pieces"]
    for k in ("chase", "level"):
        ps[k]["attempts"] = [{"at": "t", "sha": k, "file": "w.xlsx", "ok": 1, "total": 1,
                              "complete": True, "fatal": False}]
        ps[k]["submitted"] = {"at": "t", "sha": k, "file": "w.xlsx", "feasibility": [],
                              "metrics": {k: engine.summarize(engine.reference_plan(s, k))}}
    answers = {guide.start_question(s).qid: {"tries": 1, "ok": True}}
    for step in guide.steps(s, "chase"):
        if hasattr(step, "qid"):
            answers[step.qid] = {"tries": 1, "ok": True}
    ps["hybrid"]["guide"] = {"start": "chase", "answers": answers}
    ps["hybrid"]["attempts"] = [{"at": "t", "sha": "h", "file": "w.xlsx", "ok": 1, "total": 1,
                                 "complete": True, "fatal": False}]
    res = checker.PieceResult("hybrid", "Hybrid plan", sha256="h", plan=plan,
                              feasibility=engine.check_feasibility(plan, s),
                              must_be_feasible=True,
                              metrics={"hybrid": engine.summarize(plan)})
    res.cells.append(checker.CellCheck(HYB, "J2", "Total Production", "x", "January", "J", "ok"))
    at.session_state["results"]["hybrid"] = res
    at.session_state["step"] = "hybrid"
    return at.run()


def test_guided_unchanged_starting_plan_cannot_be_submitted():
    from aggplan import engine
    at = _app_on_hybrid(engine.reference_plan(scn7(), "chase"))
    assert not at.exception
    assert not any(b.label.startswith("Submit the hybrid") for b in at.button)
    assert any("isn't better than your starting plan" in i.value for i in at.info)


def test_guided_improved_plan_can_be_submitted():
    from aggplan import engine
    s = scn7()
    base = guide.starting_decision(s, "chase")
    W, R, O = list(base.workers), list(base.regular), [0.0] * 12
    for m in guide._round_down_months(s):            # follow the checklist
        i = m["i"]
        W[i] -= 1
        O[i] = R[i] - W[i] * s.cap_month
        R[i] = W[i] * s.cap_month
    plan = engine.simulate(s, engine.Decision(W, R, O, [0.0] * 12))
    at = _app_on_hybrid(plan)
    assert not at.exception
    assert any(b.label.startswith("Submit the hybrid") for b in at.button)


# ---- formulas that leave out something their column depends on ---------------------------
def test_chase_formula_missing_a_zero_term_gets_a_notice_not_a_mistake(key_book):
    # the reviewer's chase Ending Inventory =C+E-B leaves out Beginning Backorders (always 0)
    b = book_from(key_book, {(L.SHEET_CHASE, "I5"): Cell(0, "=MAX(0,C5+E5-B5)")})
    res = checker.check_book(b, scn7(), "chase", SID)
    c = _cell(res, "I5")
    assert c.status == "ok" and "Beginning Backorders (D)" in c.notice
    assert res.complete                                  # a notice never blocks
    assert "Ending Inventory" in res.notice_items()


def test_level_formula_missing_a_term_is_wrong_even_when_the_number_fits(key_book):
    # May: no backorders yet, so leaving D out gives the right number — but not in general
    b = book_from(key_book, {(L.SHEET_LEVEL, "I5"): Cell(12100, "=MAX(0,C5+E5-B5)")})
    res = checker.check_book(b, scn7(), "level", SID)
    c = _cell(res, "I5")
    assert c.status == "wrong" and "Beginning Backorders (D)" in c.diag and not c.notice
    assert not res.complete


def test_hybrid_formula_missing_overtime_is_wrong(key_book):
    # Ending Inventory from Regular Production only: right while overtime is 0
    b = book_from(key_book, {(HYB, "M5"): Cell(14900, "=MAX(0,C5+G5-D5-B5)")})
    res = checker.check_book(b, scn7(), "hybrid", SID)
    assert _cell(res, "M5").status == "wrong"


def test_complete_formulas_written_differently_pass(key_book):
    b = book_from(key_book, {
        (HYB, "M5"): Cell(14900, "=MAX(0,C5+G5+H5+I5-D5-B5)"),          # parts of Total Production
        (L.SHEET_LEVEL, "M5"): Cell(0, "=LayoffCost*'Level Plan'!H5"),   # own sheet, by name
        (L.SHEET_LEVEL, "I5"): Cell(12100, "=MAX(0,C5-D5+E5-B5)"),
    })
    for piece, ref in (("hybrid", "M5"), ("level", "M5"), ("level", "I5")):
        c = _cell(checker.check_book(b, scn7(), piece, SID), ref)
        assert c.status == "ok" and not c.notice, (piece, ref, c.diag, c.notice)


def test_answer_key_has_no_notices():
    for piece in ("chase", "level", "hybrid"):
        res = checker.check_piece(fixture("key_seed7.xlsx"), scn7(), piece, SID)
        assert not res.notice_items(), piece
