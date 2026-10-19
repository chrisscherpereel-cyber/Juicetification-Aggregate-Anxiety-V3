import io

from openpyxl import load_workbook

from aggplan import layout as L, workbook
from tests.helpers import scn7


def _wb(cfg=None, **kw):
    s = scn7(**(cfg or {}))
    data = workbook.build_workbook(s, student_name="Ann", student_id="S1", **kw)
    return s, load_workbook(io.BytesIO(data))


def test_sheets_names_and_meta():
    s, wb = _wb()
    assert wb.sheetnames == [L.SHEET_START, L.SHEET_SCENARIO, L.SHEET_CHASE, L.SHEET_LEVEL,
                             L.SHEET_HYBRID, L.SHEET_SUMMARY, L.SHEET_HELP, L.SHEET_META]
    meta_ws = wb[L.SHEET_META]
    assert meta_ws.sheet_state == "veryHidden"
    meta = {meta_ws.cell(r, 1).value: meta_ws.cell(r, 2).value for r in range(1, 11)}
    assert meta["seed"] == s.seed and meta["config_hash"] == s.config_hash
    for name in ("LaborCost", "BeginInv", "EndTarget", "BottlesPerHour", "SubCapacity"):
        assert name in wb.defined_names


def test_student_cells_are_empty_unlocked_and_given_cells_locked():
    s, wb = _wb()
    ws = wb[L.SHEET_CHASE]
    assert ws.protection.sheet
    for i in range(12):
        given = ws[f"B{L.month_row(i)}"]
        assert given.value == s.demand[i] and given.protection.locked
        for col in L.PLAN_COLUMNS["chase"][2:]:
            c = ws[L.cell_ref("chase", col.engine, i)]
            assert c.value is None and not c.protection.locked
    assert ws["S2"].value is None and not ws["S2"].protection.locked
    assert not wb[L.SHEET_SUMMARY][L.JUSTIFY_CELL].protection.locked


def test_hybrid_sheet_optional():
    _, wb = _wb(cfg={"include_hybrid": False})
    assert L.SHEET_HYBRID not in wb.sheetnames


def test_solution_writes_formulas():
    _, wb = _wb(solution=True)
    ws = wb[L.SHEET_LEVEL]
    assert ws["E5"].value == "=$S$5" and ws["S6"].value == "=ROUNDUP(S5/S2,0)"
    assert wb[L.SHEET_SUMMARY][L.LOWER_COST_CELL].value.startswith("=IF(")
    assert wb[L.SHEET_SUMMARY][L.RECOMMEND_CELL].value == "Chase"


def test_student_text_is_never_a_formula():
    data = workbook.build_workbook(scn7(), student_name="=HYPERLINK(1)", student_id="S1")
    wb = load_workbook(io.BytesIO(data))
    assert wb[L.SHEET_START][L.START_NAME_CELL].data_type == "s"
