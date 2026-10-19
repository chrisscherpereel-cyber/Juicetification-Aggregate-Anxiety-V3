"""Instructional text, generated from the active scenario so numbers are never stale.
One source for the app, the workbook and the checker's feedback."""

from __future__ import annotations

from typing import Dict, List, Tuple

from .scenario import PRESET_LABELS, Scenario


def n0(x) -> str:
    return f"{float(x):,.0f}"


def usd(x) -> str:
    x = float(x)
    return f"${x:,.0f}" if x.is_integer() else f"${x:,.2f}"


def md(s: str) -> str:
    """Escape $ for Streamlit markdown (a pair of $ would render as LaTeX)."""
    return s.replace("$", "\\$")


def param_rows(scn: Scenario) -> List[Tuple[str, str, str]]:
    """(label, value, unit) for the brief and the PDF."""
    p = scn.params
    rows = [
        ("Beginning inventory (before January)", n0(p["beginning_inventory"]), "bottles"),
        ("Safety stock", n0(p["safety_stock"]), "bottles"),
        ("Inventory target policy",
         "Maintain safety stock" if scn.maintains_safety else "No safety stock", ""),
        ("Storage capacity", n0(p["max_inventory"]), "bottles"),
        ("Output per worker-hour", n0(p["bottles_per_hour"]), "bottles/hour"),
        ("Paid hours per worker per day", n0(p["hours_per_day"]), "hours/day"),
        ("Working days per month", n0(p["working_days"]), "days"),
        ("Starting workforce", n0(p["starting_workforce"]), "workers"),
        ("Regular labor", usd(p["regular_labor_cost"]), "$/worker/month"),
        ("Hiring", usd(p["hiring_cost"]), "$/worker"),
        ("Layoff", usd(p["layoff_cost"]), "$/worker"),
        ("Holding", usd(p["holding_cost"]), "$/bottle/month"),
        ("Backorder", usd(p["backorder_cost"]), "$/bottle/month"),
        ("Overtime limit", f"{p['overtime_pct']:.0%}", "of regular capacity"),
        ("Overtime cost (extra)", usd(p["overtime_cost"]), "$/bottle"),
        ("Subcontract cost", usd(p["subcontract_cost"]), "$/bottle"),
        ("Subcontract capacity", n0(p["subcontract_capacity"]), "bottles/month"),
    ]
    if scn.preset != "standard":
        rows.append(("Scenario", PRESET_LABELS.get(scn.preset, scn.preset), ""))
    return rows


def target_sentence(scn: Scenario) -> str:
    if scn.maintains_safety:
        return (f"Your company keeps {n0(scn['safety_stock'])} bottles of safety stock: each "
                f"chase month aims to END with {n0(scn.target)} bottles, and the level plan "
                f"must end December with at least {n0(scn.target)}.")
    return ("Your company carries no safety stock: the chase plan uses up the beginning "
            "inventory in January and then ends each month with 0 bottles; the level plan "
            "aims to end December at 0.")


EXCEL_TIPS = [
    ("Start every calculation with =", "Type =C2+E2 rather than the number you worked out. "
     "The checker wants formulas, because a formula updates itself when an input changes."),
    ("Fill down", "Write February's formula once, select it and drag the small square at "
     "the bottom-right corner down to December (or double-click it)."),
    ("Absolute references with $", "When every month must point at ONE cell (for example "
     "a planning cell such as $C$17), lock it with $ signs or use a name. Without the $, "
     "filling down slides the reference onto empty cells."),
    ("Use the names", "Every given number has a name (LaborCost, HireCost, BeginInv, ...). "
     "=F2*LaborCost is easier to read and check than =F2*Scenario!$B$22."),
    ("Link, don't retype", "On the Summary sheet, point at your plan's totals: type = and "
     "click the total on the plan sheet. If your plan changes, the summary follows."),
    ("Save, then upload", "Save the file as .xlsx in Excel (or Google Sheets → Download → "
     ".xlsx) so the calculated values are stored. Then upload it in the app."),
]

EXCEL_FUNCTIONS = [
    ("SUM", "=SUM(B2:B13)", "adds a range of cells"),
    ("MAX", "=MAX(0, A5-B5)", "the larger value: with 0 first, a result can never go below zero"),
    ("MIN", "=MIN(B2:B13)", "the smallest value"),
    ("ROUNDUP", "=ROUNDUP(7.2, 0)", "rounds up to a whole number: 7.2 becomes 8"),
    ("IF", '=IF(A1<B1,"A","B")', "picks one of two results based on a test"),
    ("COUNTIF", '=COUNTIF(A2:A13,">0")', "counts the cells that meet a condition"),
    ("ABS", "=ABS(A1-B1)", "size of a difference, without the minus sign"),
]

EXCEL_ERRORS = [
    ("#DIV/0!", "you divided by zero or by an empty cell — check what is right of the /"),
    ("#NAME?", "Excel does not know a word in the formula — check the spelling of a name "
               "or function (ROUNDUP, LaborCost)"),
    ("#REF!", "the formula points at a cell that was deleted — undo (Ctrl+Z) and retype"),
    ("#VALUE!", "text where a number was expected — a cell holds words or a stray space"),
    ("####", "not an error: the column is too narrow; widen it"),
]
