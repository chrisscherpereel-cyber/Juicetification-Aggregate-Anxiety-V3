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
        ("Beginning inventory (on hand before January)", n0(p["beginning_inventory"]), "bottles"),
        ("Beginning backorders (orders still owed before January)",
         n0(p.get("beginning_backorders", 0)), "bottles"),
        ("Safety stock policy",
         f"Keep {n0(p['safety_stock'])} bottles" if scn.maintains_safety else "None", ""),
        ("Ending-inventory target (EndTarget)", n0(scn.target), "bottles"),
        ("Workforce policy",
         "Whole workers only" if scn.whole_workers else "Fractional workers allowed", ""),
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
        ("Overtime cost", usd(p["overtime_cost"]), "$/bottle made on overtime"),
        ("Subcontract cost", usd(p["subcontract_cost"]), "$/bottle"),
        ("Subcontract capacity", n0(p["subcontract_capacity"]), "bottles/month"),
    ]
    if scn.preset != "standard":
        rows.append(("Scenario", PRESET_LABELS.get(scn.preset, scn.preset), ""))
    return rows


def policies(scn: Scenario) -> List[Tuple[str, str, List[str]]]:
    """The company's planning policies: (title, one-line summary, details). These are the
    business rules every plan must follow — stated as policy, not as formulas."""
    p = scn.params
    s, b = n0(p["safety_stock"]), n0(p["beginning_inventory"])
    bb = n0(p.get("beginning_backorders", 0))
    if scn.maintains_safety:
        inv = ("Safety stock: KEEP " + s + " bottles", [
            f"Your company keeps a safety stock of {s} bottles as a cushion against surprises, "
            f"so EndTarget = {s}.",
            f"Chase plan: every month aims to END with {s} bottles on the shelf — no more, no less.",
            f"Level plan: one steady rate, chosen so that December ends with {s} bottles.",
            f"Hybrid plan: December must end with at least {s} bottles.",
            "Safety stock is a target, not a locked reserve: if demand is higher than planned, "
            "customers are served first.",
            "Safety-stock bottles are held like any others, so they cost HoldCost every month "
            "they sit on the shelf.",
            f"Beginning inventory (BeginInv = {b}) is simply the stock on the shelf on January 1. "
            "It is not the safety stock, even if the two numbers happen to be equal.",
        ])
    else:
        inv = ("Safety stock: NONE", [
            "Your company does NOT keep safety stock, so EndTarget = 0.",
            "Chase plan: use up the beginning inventory, then aim to end every month with 0 "
            "bottles on the shelf.",
            "Level plan: one steady rate, chosen so that December ends with 0 bottles.",
            "Hybrid plan: December may end with any stock of 0 or more.",
            f"Beginning inventory (BeginInv = {b}) is simply the stock on the shelf on January 1 "
            "— use it up; don't keep it as a reserve.",
        ])
    bo = ("Backorders: unfilled orders carry forward", [
        "An order you can't fill is not lost: it becomes a backorder and is still owed next month.",
        f"You start the year owing BeginBackorders = {bb} bottles of orders.",
        "When stock is available, old orders are shipped first, then this month's demand.",
        "Every bottle still on backorder at the end of a month costs BackorderCost for that month.",
        "Stock is never negative: a shortfall is recorded as backorders, not as negative inventory.",
        f"Hybrid plan: no more than {n0(p['terminal_backlog_max'])} bottles may still be on "
        f"backorder at the end of December"
        + (f", and at least {p['service_requirement_pct']:g}% of demand must ship in the month "
           "it is ordered." if p["service_requirement_pct"] > 0 else "."),
    ])
    if scn.whole_workers:
        wk = ("Workforce: whole workers only", [
            "You employ whole people — there are no fractions of a worker.",
            "If a month's production needs, say, 10.2 workers' worth of capacity, you must employ "
            "11. The extra capacity sits idle, but those workers are still paid.",
            "Workers are paid LaborCost for every month they are employed, busy or not.",
            f"You start the year with StartWorkers = {n0(p['starting_workforce'])} workers. "
            "Hiring and layoffs are counted against the previous month's workforce (January "
            "against StartWorkers): each hire costs HireCost, each layoff LayoffCost.",
            "One worker's monthly capacity comes from their output per hour, paid hours per day "
            "and working days per month (Scenario sheet).",
        ])
    else:
        wk = ("Workforce: fractional workers allowed", [
            "Your company can staff in fractions of a worker (part-time hours, shared staff), so "
            "capacity can be added or removed in small increments instead of a whole worker at "
            "a time. Think of the Workers column as worker-equivalents.",
            "Employ exactly the workforce production needs — 10.2 workers is fine. Do not round, "
            "so no capacity sits idle just because of rounding.",
            "Labor, hiring and layoff costs are paid in proportion: half a worker hired costs "
            "half of HireCost.",
            f"You start the year with StartWorkers = {n0(p['starting_workforce'])} workers. "
            "Hiring and layoffs are counted against the previous month's workforce (January "
            "against StartWorkers).",
            "One worker's monthly capacity comes from their output per hour, paid hours per day "
            "and working days per month (Scenario sheet).",
        ])
    return [(t, d[0], d[1:]) for t, d in (inv, bo, wk)]


def target_sentence(scn: Scenario) -> str:
    """One-line summary of all three policies."""
    return " ".join(f"{t}." for t, _, _ in policies(scn))


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

# Patterns: how a planning question becomes a formula (the shape, never the answer).
EXCEL_PATTERNS = [
    ("Never below zero", "=MAX(0, have - need)",
     "Stock, backorders, hires and layoffs can't be negative. MAX(0, …) keeps the result at 0 "
     "when the subtraction would go below zero."),
    ("Whole things to cover a need", "=ROUNDUP(need / per_unit, 0)",
     "How many whole workers make enough bottles? Divide what you need by what one worker "
     "makes, then round UP — rounding down leaves you short."),
    ("This month starts where last month ended", "=last month's ending cell",
     "Each month's beginning value points at the row above. January has no row above: it uses "
     "the given starting value (a name such as BeginInv)."),
    ("A change between months", "=MAX(0, this month - last month)",
     "An increase only (hires); swap the two for a decrease only (layoffs). January compares "
     "with the given starting value."),
    ("Stock balance", "start + made - owed - demand",
     "What you have after the month: bottles on the shelf, plus what you made, minus orders "
     "still owed from before, minus this month's demand. If it is above 0 it's ending "
     "inventory; below 0 the shortfall is ending backorders — two MAX(0, …) formulas, one "
     "each way."),
    ("Count the months that…", '=COUNTIF(range, ">0")',
     "Counts the cells in a range that meet a condition, e.g. months that end with "
     "backorders."),
    ("Pick a word with a test", '=IF(test, "this", "that")',
     "For an answer that should follow your numbers, e.g. which plan costs less."),
]

EXCEL_ERRORS = [
    ("#DIV/0!", "you divided by zero or by an empty cell — check what is right of the /"),
    ("#NAME?", "Excel does not know a word in the formula — check the spelling of a name "
               "or function (ROUNDUP, LaborCost)"),
    ("#REF!", "the formula points at a cell that was deleted — undo (Ctrl+Z) and retype"),
    ("#VALUE!", "text where a number was expected — a cell holds words or a stray space"),
    ("####", "not an error: the column is too narrow; widen it"),
    ("Circular reference", "a formula uses its own result, directly or through other cells — "
                           "Excel then shows 0. Follow the arrows (Formulas ▸ Error Checking ▸ "
                           "Circular References) and break the loop. On the hybrid sheet, type "
                           "your decisions as numbers."),
]
