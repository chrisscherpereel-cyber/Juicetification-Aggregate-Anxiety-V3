"""The checker against workbooks calculated by REAL Excel (tests/fixtures, made by
tests/make_fixtures.py), plus targeted cases built on top of them."""

import pytest

from aggplan import checker, layout as L, workbook
from aggplan.checker import Cell
from aggplan.scenario import resolve_scenario
from tests.helpers import SID, book_from, fixture, scn7

PIECES = ["chase", "level", "hybrid", "summary"]


@pytest.fixture(scope="module")
def key_book():
    book, err = checker.read_book(fixture("key_seed7.xlsx"))
    assert err is None
    return book


def _cell(res, ref, sheet=None):
    return next(c for c in res.cells if c.ref == ref and (sheet is None or c.sheet == sheet))


@pytest.mark.parametrize("piece", PIECES)
def test_excel_answer_key_completes_every_piece(piece):
    res = checker.check_piece(fixture("key_seed7.xlsx"), scn7(), piece, SID)
    assert res.complete, (res.notes, res.messages, res.counts())


@pytest.mark.parametrize("piece", PIECES)
def test_students_own_columns_and_headings_are_recognised(piece):
    """Chase columns in a different order, worded like a student would (Excel-calculated)."""
    res = checker.check_piece(fixture("shuffled_seed7.xlsx"), scn7(), piece, SID)
    assert res.complete, (res.notes, res.counts())
    if piece == "chase":
        read = {h: r for _, h, r in res.columns}
        assert read["Starting Stock"] == "Beginning Inventory" and read["Fired"] == "Layoffs"
        assert read["Regular Capacity"] == "Regular Capacity"          # optional, checked


def test_pieces_are_checked_separately():
    data = fixture("mistakes_seed7.xlsx")
    chase = checker.check_piece(data, scn7(), "chase", SID)
    level = checker.check_piece(data, scn7(), "level", SID)
    hybrid = checker.check_piece(data, scn7(), "hybrid", SID)
    assert {c.sheet for c in chase.cells} == {L.SHEET_CHASE}
    assert {c.sheet for c in level.cells} == {L.SHEET_LEVEL}
    assert hybrid.complete and not chase.complete and not level.complete


def test_typical_mistakes_are_diagnosed():
    data = fixture("mistakes_seed7.xlsx")
    chase = checker.check_piece(data, scn7(), "chase", SID)
    assert _cell(chase, "N3").status == "typed"
    assert set(chase.problem_items()) == {"Holding Cost"}
    level = checker.check_piece(data, scn7(), "level", SID)
    e3 = _cell(level, "E3")
    assert e3.status == "wrong" and "$C$20" in e3.diag
    assert _cell(level, "E2").status == "ok"
    assert _cell(level, "P3").status == "carried"
    assert set(level.problem_items()) == {"Regular Production"}     # carried cells excluded
    summ = checker.check_piece(data, scn7(), "summary", SID)
    assert _cell(summ, L.JUSTIFY_CELL).status == "wrong"


def test_missing_column_is_reported_not_graded(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "D1"): Cell(None, None)})   # drop Beginning Backorders
    res = checker.check_book(b, scn7(), "chase", SID)
    assert res.missing_columns == ["Beginning Backorders"] and not res.complete
    assert not any(c.item == "Beginning Backorders" for c in res.cells)


def test_unclear_heading_is_guessed_from_values(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "I1"): Cell("Column 7", None)})
    res = checker.check_book(b, scn7(), "chase", SID)
    assert [(n.kind, n.concept) for n in res.notes] == [("guessed", "Ending Inventory")]
    assert res.complete


def test_duplicate_columns_block_completion(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "Q1"): Cell("Closing inventory", None)})
    res = checker.check_book(b, scn7(), "chase", SID)
    assert any(n.kind == "duplicate" for n in res.notes) and not res.complete


def test_blank_workbook_counts_as_todo_not_mistakes():
    raw = workbook.build_workbook(scn7(), student_name="x", student_id=SID)
    res = checker.check_piece(raw, scn7(), "chase", SID)
    assert res.problem_items() == {}
    assert len(res.missing_columns) == len(L.REQUIRED["chase"])
    assert set(res.todo_items()) == {"plan:capacity", "plan:target"}


def test_file_never_opened_in_excel_is_reported_uncalculated():
    raw = workbook.build_workbook(scn7(), student_name="x", student_id=SID, solution=True)
    res = checker.check_piece(raw, scn7(), "chase", SID)
    assert res.uncalculated and not res.complete


def test_hybrid_must_be_feasible(key_book):
    # 9 workers making 9,000 a month can't meet demand: backorders pile up
    changes = {}
    for i in range(12):
        r = L.month_row(i)
        changes[(L.SHEET_HYBRID, f"E{r}")] = Cell(9.0, None)
        changes[(L.SHEET_HYBRID, f"G{r}")] = Cell(9000.0, None)
    res = checker.check_book(book_from(key_book, changes), scn7(), "hybrid", SID)
    assert res.feasibility and not res.feasible and not res.complete


@pytest.mark.parametrize("data,expect", [(b"not a zip", "not an Excel"), (b"", "empty")])
def test_garbage_is_rejected(data, expect):
    res = checker.check_piece(data, scn7(), "chase", SID)
    assert res.fatal and expect in res.fatal[0]


def test_other_students_workbook_is_rejected():
    res = checker.check_piece(fixture("key_seed7.xlsx"), scn7(), "chase", "SOMEONE-ELSE")
    assert res.fatal and "belongs to student ID" in res.fatal[0]


def test_other_scenario_is_rejected():
    res = checker.check_piece(fixture("key_seed7.xlsx"), resolve_scenario({}, 8), "chase", SID)
    assert res.fatal and "different scenario" in res.fatal[-1]


def test_presentation_settings_do_not_invalidate():
    res = checker.check_piece(fixture("key_seed7.xlsx"), scn7(min_justification_words=10),
                              "summary", SID)
    assert res.complete


def test_changed_demand_is_fatal(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "B5"): Cell(99.0, None)})
    res = checker.check_book(b, scn7(), "chase", SID)
    assert res.fatal and "demand for April was changed" in res.fatal[0]


def test_unrounded_workers_get_a_specific_observation(key_book):
    # January chase production for seed 7 is 4,600 bottles -> 4.6 workers
    b = book_from(key_book, {(L.SHEET_CHASE, "F2"): Cell(4.6, "=E2/$C$17")})
    c = _cell(checker.check_book(b, scn7(), "chase", SID), "F2")
    assert c.status == "wrong" and "part of a person" in c.diag


def test_typed_numbers_allowed_when_formulas_not_required(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "K2"): Cell(16000.0, None)})
    assert _cell(checker.check_book(b, scn7(), "chase", SID), "K2").status == "typed"
    assert checker.check_book(b, scn7(require_formulas=False), "chase", SID).complete


def test_lower_cost_answer_must_match(key_book):
    b = book_from(key_book, {(L.SHEET_SUMMARY, L.LOWER_COST_CELL):
                             Cell("Level", '=IF(B11>C11,"Chase","Level")')})
    res = checker.check_book(b, scn7(), "summary", SID)
    assert _cell(res, L.LOWER_COST_CELL).status == "wrong"


def test_grid_follows_the_students_layout():
    res = checker.check_piece(fixture("shuffled_seed7.xlsx"), scn7(), "chase", SID)
    g = checker.grid(res)
    assert list(g.index)[:12] == list(scn7().months) and g.index[-1] == "Total"
    assert g.columns[0] == "C · Starting Stock"
