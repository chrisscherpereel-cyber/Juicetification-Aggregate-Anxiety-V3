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


def test_whats_new_for_each_plan():
    assert "14 columns" in columns.whats_new("chase")
    assert "Same 14 columns" in columns.whats_new("level")
    new = columns.whats_new("hybrid")
    for c in ("Overtime Production", "Subcontract Production", "Total Production",
              "Overtime Cost", "Subcontract Cost"):
        assert c in new
    assert "5 new ones" in new


def test_heading_check_reads_like_the_checker_and_lists_what_is_missing():
    out = columns.check_headings("chase", ["Starting Stock", "Backlog at start", "Production",
                                           "Crew", "original workers", "Crew size", ""])
    read = {h: r for _, h, r in out["rows"]}
    assert read["Starting Stock"] == "Beginning Inventory"
    assert read["Backlog at start"] == "Beginning Backorders"
    assert read["Crew"] == "Workers" and read["original workers"].startswith("helper")
    assert read["Crew size"].startswith("duplicate")
    assert [r[0] for r in out["rows"]][:3] == ["C", "D", "E"]
    missing = [h for h, _ in out["missing"]]
    assert "Hires" in missing and "Total Cost" in missing and "Workers" not in missing
    full = columns.check_headings("chase", [h for _, h in workbook.default_columns("chase")])
    assert full["missing"] == []


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


def test_start_here_has_a_printable_column_planner():
    ws = load_workbook(io.BytesIO(workbook.build_workbook(scn7(), student_name="x",
                                                          student_id=SID)))[L.SHEET_START]
    cells = {c.coordinate: c for row in ws.iter_rows() for c in row if c.value is not None}
    head = next(c for c in cells.values() if c.value == "My heading")
    r = head.row
    assert ws[f"B{r}"].value == "Column" and ws[f"F{r}"].value == "Needed in"
    names = [ws[f"B{k}"].value for k in range(r + 1, r + 30)]
    for c in L.REQUIRED["hybrid"]:
        assert workbook.L.heading_for("hybrid", c) in names, c
    row = names.index("Overtime Production") + r + 1
    assert ws[f"F{row}"].value == "Hybrid"
    assert ws[f"G{row}"].protection.locked is False        # students can type their heading
    assert ws.row_breaks.brk                                 # planner prints on its own page
