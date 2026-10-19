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


# Rule for each calculated column (in words, using the Excel names). Shown as cell comments
# in the workbook, on the Start Here sheet, in the app guide and in the checker feedback.
def column_rules(kind: str) -> Dict[str, str]:
    prod = {
        "chase": "MAX(0, Demand + EndTarget − Beginning Inventory + Beginning Backorders): "
                 "make exactly what this month needs.",
        "level": "the level production per month from your planning cell (same every month).",
        "hybrid": "your decision.",
    }[kind]
    workers = {
        "chase": "ROUNDUP(Production ÷ capacity per worker, 0): whole workers, so round UP.",
        "level": "the level workforce from your planning cell (same every month).",
        "hybrid": "your decision (whole workers).",
    }[kind]
    total_prod = "Regular + Overtime + Subcontract production." if kind == "hybrid" else ""
    rules = {
        "Beginning Inventory": "January = BeginInv. Later months = last month's Ending Inventory.",
        "Beginning Backorders": "January = 0. Later months = last month's Ending Backorders.",
        "Regular Production": prod,
        "Workers": workers,
        "Regular Capacity": "Workers × capacity per worker.",
        "Total Production": total_prod,
        "Hires": "MAX(0, Workers − last month's Workers). January compares with StartWorkers.",
        "Layoffs": "MAX(0, last month's Workers − Workers). January compares with StartWorkers.",
        "Ending Inventory": "MAX(0, Beginning Inventory + Production − Beginning Backorders − "
                            "Demand). Inventory is never negative.",
        "Ending Backorders": "MAX(0, Demand + Beginning Backorders − Beginning Inventory − "
                             "Production). Unfilled orders carried to next month.",
        "Labor Cost": "Workers × LaborCost (workers are paid whether or not they are busy).",
        "Hiring Cost": "Hires × HireCost.",
        "Layoff Cost": "Layoffs × LayoffCost.",
        "Overtime Cost": "Overtime Production × OTCost.",
        "Subcontract Cost": "Subcontract Production × SubCost.",
        "Holding Cost": "Ending Inventory × HoldCost.",
        "Backorder Cost": "Ending Backorders × BackorderCost.",
        "Total Cost": ("Labor + Hiring + Layoff + Overtime + Subcontract + Holding + Backorder."
                       if kind == "hybrid" else "Labor + Hiring + Layoff + Holding + Backorder."),
    }
    if kind == "hybrid":
        rules["Total Production"] = total_prod
        rules["Ending Inventory"] = rules["Ending Inventory"].replace("Production", "Total Production")
        rules["Ending Backorders"] = rules["Ending Backorders"].replace("Production", "Total Production")
    return rules


PLANNING_RULES = {
    "capacity": "BottlesPerHour × HoursPerDay × DaysPerMonth.",
    "target": "=EndTarget (the inventory each month should end with).",
    "total_demand": "SUM of the twelve Demand cells.",
    "level_rate": "(Total demand + ending inventory target − BeginInv) ÷ 12.",
    "level_workers": "ROUNDUP(Level production ÷ capacity per worker, 0).",
    "ot_limit": "Capacity per worker × OTLimitPct.",
}

TOTAL_RULE = "SUM of the twelve monthly values above it (row 14)."

SUMMARY_RULES = {
    "link": "Link to the matching total on the plan sheet, e.g. ='Chase Plan'!P14.",
    "Peak Workforce": "MAX of the plan's Workers column.",
    "Peak Inventory": "MAX of the plan's Ending Inventory column.",
    "Months With Backorders": 'COUNTIF of the Ending Backorders column, criteria ">0".',
    "lower": 'IF(chase total cost < level total cost, "Chase", "Level").',
    "savings": "ABS(chase total cost − level total cost).",
}


EXCEL_TIPS = [
    ("Start every calculation with =", "Type =C2+E2 rather than the number you worked out. "
     "The checker wants formulas, because a formula updates itself when an input changes."),
    ("Fill down", "Write February's formula once, select it and drag the small square at "
     "the bottom-right corner down to December (or double-click it)."),
    ("Absolute references with $", "When every month must point at ONE cell (for example "
     "capacity per worker in $S$2), lock it with $ signs or use a name. Without the $, "
     "filling down slides the reference onto empty cells."),
    ("Use the names", "Every given number has a name (LaborCost, HireCost, BeginInv, ...). "
     "=F2*LaborCost is easier to read and check than =F2*Scenario!$B$22."),
    ("Link, don't retype", "On the Summary sheet, point at the plan totals: ='Chase Plan'!P14. "
     "If your plan changes, the summary follows."),
    ("Save, then upload", "Save the file as .xlsx in Excel (or Google Sheets → Download → "
     ".xlsx) so the calculated values are stored. Then upload it in the app."),
]

EXCEL_FUNCTIONS = [
    ("SUM", "=SUM(B2:B13)", "adds a range: total demand, total cost"),
    ("MAX", "=MAX(0, C2+E2-D2-B2)", "the larger value: keeps inventory from going negative"),
    ("MIN", "=MIN(B2:B13)", "the smallest value"),
    ("ROUNDUP", "=ROUNDUP(E2/$S$2, 0)", "rounds up: 10.2 workers means you need 11"),
    ("IF", '=IF(B11<C11,"Chase","Level")', "picks one of two results based on a test"),
    ("COUNTIF", '=COUNTIF(J2:J13,">0")', "counts the cells that meet a condition"),
    ("ABS", "=ABS(B11-C11)", "size of a difference, without the minus sign"),
]

EXCEL_ERRORS = [
    ("#DIV/0!", "you divided by zero or by an empty cell — check what is right of the /"),
    ("#NAME?", "Excel does not know a word in the formula — check the spelling of a name "
               "or function (ROUNDUP, LaborCost)"),
    ("#REF!", "the formula points at a cell that was deleted — undo (Ctrl+Z) and retype"),
    ("#VALUE!", "text where a number was expected — a cell holds words or a stray space"),
    ("####", "not an error: the column is too narrow; widen it"),
]
