"""Help with designing the columns: the planner, what's new per plan, the heading check, the
diagram and the printable planner on the Start Here sheet."""

import io

import pytest
from openpyxl import load_workbook

from aggplan import checker, columns, layout as L, workbook
from tests.helpers import SID, scn7

KINDS = ["chase", "level", "hybrid"]


@pytest.mark.parametrize("kind", KINDS)
def test_planner_covers_every_required_column_once(kind):
    listed = [c.concept for _, _, cols in columns.stages(kind) for c in cols]
    assert sorted(listed) == sorted(set(L.REQUIRED[kind]) | set(L.OPTIONAL[kind]))
    for _, _, cols in columns.stages(kind):
        for c in cols:
            assert c.question and c.why and c.comes_from, c.concept


@pytest.mark.parametrize("kind", KINDS)
def test_planner_ingredients_match_what_the_checker_requires(kind):
    """'Comes from' must name what the checker's left-out-input check requires, so the
    planner never tells a student less than the checker expects."""
    for concept, needs in checker._NEEDS.items():
        if concept not in columns.concepts(kind) or columns.is_decision(kind, concept):
            continue
        if concept in checker._CHASE_ONLY_NEEDS and kind != "chase":
            continue
        ins = set(columns.inputs(kind, concept))
        for alts in needs:
            want = {("Total Production" if kind == "hybrid" else "Regular Production")
                    if a == "Production" else a for a in alts}
            assert ins & want, (kind, concept, alts, ins)


def test_used_by_is_the_reverse_of_comes_from_and_uses_the_plans_titles():
    assert "Production" in columns.used_by("chase", "Beginning Inventory")
    assert "Hiring Cost" in columns.used_by("chase", "Hires")
    assert "next month's Beginning Inventory" in columns.used_by("chase", "Ending Inventory")
    assert "Total Production" in columns.used_by("hybrid", "Overtime Production")


def test_decisions_and_level_planning_cells_are_explained():
    assert "decision" in columns.comes_from("hybrid", "Workers").lower()
    assert "planning cell" in columns.comes_from("level", "Regular Production")
    assert "rounded UP" in columns.comes_from("chase", "Workers", whole_workers=True)
    assert "rounded" not in columns.comes_from("chase", "Workers", whole_workers=False)


def test_whats_new_does_not_give_the_new_columns_away():
    assert "14 questions" in columns.whats_new("chase")
    assert "Same 14" in columns.whats_new("level").replace("same", "Same")
    new = columns.whats_new("hybrid")
    assert "5 new questions" in new
    for c in ("Overtime Production", "Subcontract Production", "Total Production"):
        assert c not in new


def test_one_heading_at_a_time():
    found = {}
    r = columns.check_one("chase", "Starting Stock", found)
    assert r["status"] == "ok" and r["name"] == "Beginning Inventory"
    assert r["why"] and r["question"]                     # why it matters, once found
    found[r["concept"]] = "Starting Stock"
    assert columns.check_one("chase", "Opening inventory", found)["status"] == "repeat"
    assert columns.check_one("chase", "stuff", found) == {"status": "unknown"}   # no names
    assert columns.check_one("chase", "helper: notes", found)["status"] == "helper"
    assert columns.check_one("chase", "Overtime", found) == {"status": "other_plan",
                                                             "plans": ["hybrid"]}
    assert columns.check_one("hybrid", "Overtime", found)["status"] == "ok"
    assert columns.check_one("chase", "", found)["status"] == "empty"
    # a bare "Inventory" is read as the END of the month, and says so
    assert "end" in columns.check_one("chase", "Inventory", found)["note"]
    assert not columns.check_one("chase", "Closing stock", found)["note"]


def test_support_grows_with_each_miss():
    s1, s2, s3 = (columns.support(n) for n in (1, 2, 3))
    assert len({s1, s2, s3}) == 3 and len(s1) < len(s3)
    assert columns.support(9) == s3
    assert "inventory" not in s1.lower() and "inventory" in s3.lower()


def test_open_questions_list_questions_not_column_names():
    oq = columns.open_questions("hybrid", {"Beginning Inventory": "x"})
    flat = [(q, tag) for _, qs in oq for q, tag in qs]
    assert all("Inventory" not in q for q, _ in flat)
    assert ("How many bottles did you make on overtime?", "new") in flat
    assert ("How many bottles could your workforce make on regular time?", "optional") in flat
    assert "How many bottles are on the shelf when the month starts?" not in [q for q, _ in flat]


@pytest.mark.parametrize("kind", KINDS)
def test_diagram_is_valid_dot_with_every_column(kind):
    dot = columns.diagram(kind)
    assert dot.startswith("digraph") and dot.rstrip().endswith("}")
    assert dot.count("{") == dot.count("}")
    for c in columns.concepts(kind):
        assert f'"{c}"' in dot
    assert '"Ending Inventory" -> "Holding Cost"' in dot
    if kind == "level":
        assert '"Planning cells" -> "Regular Production"' in dot


def test_diagram_hides_columns_not_found_yet():
    dot = columns.diagram("chase", {"Workers": "Crew"})
    assert '"Workers" [fillcolor' in dot and 'label="Workers' in dot
    assert '"Hires" [fillcolor="#FFFFFF", label="?"' in dot


def test_start_here_has_a_printable_column_planner():
    ws = load_workbook(io.BytesIO(workbook.build_workbook(scn7(), student_name="x",
                                                          student_id=SID)))[L.SHEET_START]
    cells = {c.coordinate: c for row in ws.iter_rows() for c in row if c.value is not None}
    head = next(c for c in cells.values() if c.value == "My heading")
    r = head.row
    assert [ws[f"{x}{r}"].value for x in "BCDE"] == ["#", "Question it answers", "My heading",
                                                       "Needed in"]
    text_below = " ".join(str(ws[f"{x}{k}"].value) for k in range(r + 1, r + 30)
                          for x in "BCE" if ws[f"{x}{k}"].value is not None)
    assert "How many bottles did you make on overtime?" in text_below
    for name in ("Beginning Inventory", "Ending Backorders", "Labor Cost"):
        assert name not in text_below                         # questions, not answers
    row = next(k for k in range(r + 1, r + 30)
               if ws[f"C{k}"].value == "How many bottles did you make on overtime?")
    assert ws[f"E{row}"].value == "Hybrid"
    assert ws[f"D{row}"].protection.locked is False           # students can type their heading
    assert ws.row_breaks.brk                                    # prints on its own page
