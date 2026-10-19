"""The checker against workbooks calculated by REAL Excel (tests/fixtures, made by
tests/make_fixtures.py), plus targeted cases built on top of them."""

import pytest

from aggplan import checker, layout as L, workbook
from aggplan.checker import Cell
from aggplan.scenario import resolve_scenario
from tests.helpers import SID, book_from, fixture, scn7


@pytest.fixture(scope="module")
def key_book():
    book, err = checker.read_book(fixture("key_seed7.xlsx"))
    assert err is None
    return book


def _cell(res, sheet_key, ref):
    sheet = res.sheets[sheet_key]
    return next(c for c in sheet.cells if c.ref == ref and c.sheet == sheet.title)


def test_excel_answer_key_is_complete():
    res = checker.check_workbook(fixture("key_seed7.xlsx"), scn7(), SID)
    assert res.readable and res.complete, res.missing()
    assert res.plan_totals["chase"]["Total Cost"] == 426000
    assert res.sheets["hybrid"].complete
    assert all(c["ok"] for c in res.hybrid_feasibility)


def test_typical_mistakes_are_diagnosed():
    res = checker.check_workbook(fixture("mistakes_seed7.xlsx"), scn7(), SID)
    assert not res.complete
    assert _cell(res, "chase", "N3").status == "typed"
    e3 = _cell(res, "level", "E3")
    assert e3.status == "wrong" and "$S$6" in e3.message
    assert _cell(res, "level", "E2").status == "ok"
    assert _cell(res, "level", "P3").status == "carried"      # formula right, input wrong
    assert _cell(res, "summary", L.JUSTIFY_CELL).status == "wrong"
    todo, carried = res.fix_first()
    assert carried > 0 and all(c.status != "carried" for c, _ in todo)
    assert ("Level Plan", "E3") in [(c.sheet, c.ref) for c, _ in todo]


def test_file_never_opened_in_excel_is_reported_uncalculated():
    raw = workbook.build_workbook(scn7(), student_name="x", student_id=SID, solution=True)
    res = checker.check_workbook(raw, scn7(), SID)
    assert res.uncalculated and not res.complete


def test_blank_workbook_is_all_blank():
    raw = workbook.build_workbook(scn7(), student_name="x", student_id=SID)
    res = checker.check_workbook(raw, scn7(), SID)
    assert res.sheets["chase"].counts() == {"blank": res.sheets["chase"].n_total}
    assert not res.sheets["hybrid"].attempted and not res.sheets["hybrid"].required


@pytest.mark.parametrize("data,expect", [(b"not a zip", "not an Excel"), (b"", "empty")])
def test_garbage_is_rejected(data, expect):
    res = checker.check_workbook(data, scn7(), SID)
    assert res.fatal and expect in res.fatal[0]


def test_other_students_workbook_is_rejected():
    res = checker.check_workbook(fixture("key_seed7.xlsx"), scn7(), "SOMEONE-ELSE")
    assert res.fatal and "belongs to student ID" in res.fatal[0]


def test_other_scenario_is_rejected():
    res = checker.check_workbook(fixture("key_seed7.xlsx"), resolve_scenario({}, 8), SID)
    assert res.fatal and "different scenario" in res.fatal[-1]


def test_changed_instructor_parameter_is_rejected():
    assert checker.check_workbook(fixture("key_seed7.xlsx"), scn7(hiring_cost=700), SID).fatal


def test_presentation_settings_do_not_invalidate():
    res = checker.check_workbook(fixture("key_seed7.xlsx"), scn7(min_justification_words=10), SID)
    assert res.complete


def test_changed_demand_is_fatal(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "B5"): Cell(99.0, None)})
    res = checker.check_book(b, scn7(), SID)
    assert res.fatal and "demand for April was changed" in res.fatal[0]


def test_unrounded_workers_get_a_roundup_hint(key_book):
    # January chase production for seed 7 is 4,600 bottles -> 4.6 workers, needs 5
    b = book_from(key_book, {(L.SHEET_CHASE, "F2"): Cell(4.6, "=E2/$S$2")})
    c = _cell(checker.check_book(b, scn7(), SID), "chase", "F2")
    assert c.status == "wrong" and "ROUNDUP" in c.message


def test_negative_inventory_gets_a_max_hint(key_book):
    b = book_from(key_book, {(L.SHEET_LEVEL, "I8"): Cell(-500.0, "=C8+E8-D8-B8")})
    c = _cell(checker.check_book(b, scn7(), SID), "level", "I8")
    assert c.status == "wrong" and "MAX(0" in c.message


def test_excel_error_values(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "K4"): Cell("#NAME?", "=F4*LaborCst")})
    c = _cell(checker.check_book(b, scn7(), SID), "chase", "K4")
    assert c.status == "error" and "#NAME?" in c.message


def test_typed_numbers_allowed_when_formulas_not_required(key_book):
    b = book_from(key_book, {(L.SHEET_CHASE, "K2"): Cell(16000.0, None)})
    assert _cell(checker.check_book(b, scn7(), SID), "chase", "K2").status == "typed"
    assert checker.check_book(b, scn7(require_formulas=False), SID).complete


def test_lower_cost_answer_must_match(key_book):
    b = book_from(key_book, {(L.SHEET_SUMMARY, L.LOWER_COST_CELL):
                             Cell("Level", '=IF(B11>C11,"Chase","Level")')})
    res = checker.check_book(b, scn7(), SID)
    assert _cell(res, "summary", L.LOWER_COST_CELL).status == "wrong"


def test_recommending_hybrid_makes_it_required(key_book):
    b = book_from(key_book, {(L.SHEET_SUMMARY, L.RECOMMEND_CELL): Cell("Hybrid", None),
                             (L.SHEET_HYBRID, "V2"): Cell(1.0, "=1")})
    res = checker.check_book(b, scn7(), SID)
    assert res.sheets["hybrid"].required and not res.complete


def test_workbook_identifies_its_student(key_book):
    assert key_book.meta["student_id"] == SID and key_book.meta["seed"] == "7"


def test_grid_shape(key_book):
    res = checker.check_book(key_book, scn7(), SID)
    g = checker.grid(res.sheets["chase"])
    assert list(g.index)[:12] == list(scn7().months) and g.index[-1] == "Total"
    assert len(g.columns) == 14
