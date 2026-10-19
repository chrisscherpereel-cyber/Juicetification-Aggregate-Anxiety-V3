"""Where everything lives in the student workbook.

Plan sheets give only Month (column A) and Demand (column B). The student designs the rest
of the worksheet: they type their own column headings in row 1 (columns C onward) and their
formulas in rows 2-13 (January-December), with yearly totals in row 14. The checker finds
each column by its heading (recognize.py), so any order and any clear wording works.

Below the grid, a short block of planning cells (labels in A:B, the student's formula in
column C) holds the one-off calculations each plan needs (capacity per worker, level rate…).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List

FIRST_ROW = 2                         # January
N_MONTHS = 12
LAST_ROW = FIRST_ROW + N_MONTHS - 1   # December = 13
TOTAL_ROW = LAST_ROW + 1              # 14
FIRST_STUDENT_COL = 3                 # C
LAST_STUDENT_COL = 30                 # AD — room for helper columns
PLANNING_FIRST_ROW = 17
PLANNING_VALUE_COL = "C"

SHEET_START = "Start Here"
SHEET_SCENARIO = "Scenario"
SHEET_CHASE = "Chase Plan"
SHEET_LEVEL = "Level Plan"
SHEET_HYBRID = "Hybrid Plan"
SHEET_SUMMARY = "Summary"
SHEET_HELP = "Excel Help"
SHEET_META = "_meta"

PLAN_SHEETS = {"chase": SHEET_CHASE, "level": SHEET_LEVEL, "hybrid": SHEET_HYBRID}
PLAN_LABEL = {"chase": "Chase", "level": "Level", "hybrid": "Hybrid"}

# --------------------------------------------------------------------------- #
# Concepts: the columns a plan needs. Concept ids are engine.simulate() column names.
# --------------------------------------------------------------------------- #
COST_CONCEPTS = ["Labor Cost", "Hiring Cost", "Layoff Cost", "Overtime Cost",
                 "Subcontract Cost", "Holding Cost", "Backorder Cost", "Total Cost"]

REQUIRED: Dict[str, List[str]] = {
    "chase": ["Beginning Inventory", "Beginning Backorders", "Regular Production", "Workers",
              "Hires", "Layoffs", "Ending Inventory", "Ending Backorders", "Labor Cost",
              "Hiring Cost", "Layoff Cost", "Holding Cost", "Backorder Cost", "Total Cost"],
    "hybrid": ["Beginning Inventory", "Beginning Backorders", "Workers", "Regular Production",
               "Overtime Production", "Subcontract Production", "Total Production", "Hires",
               "Layoffs", "Ending Inventory", "Ending Backorders", "Labor Cost", "Hiring Cost",
               "Layoff Cost", "Overtime Cost", "Subcontract Cost", "Holding Cost",
               "Backorder Cost", "Total Cost"],
}
REQUIRED["level"] = list(REQUIRED["chase"])

# Recognised and checked when present, never required.
OPTIONAL: Dict[str, List[str]] = {"chase": ["Regular Capacity"], "level": ["Regular Capacity"],
                                  "hybrid": ["Regular Capacity"]}

# Hybrid columns that are the student's own choices (not calculated).
DECISIONS = ["Workers", "Regular Production", "Overtime Production", "Subcontract Production"]

# Row-14 totals: required under every cost column, checked when present under these.
TOTAL_OPTIONAL = ["Regular Production", "Hires", "Layoffs", "Overtime Production",
                  "Subcontract Production", "Total Production", "Regular Capacity"]

# Headings used by the answer key (and suggested at the last hint level).
CANONICAL = {
    "Beginning Inventory": "Beginning Inventory", "Beginning Backorders": "Beginning Backorders",
    "Regular Production": "Production", "Workers": "Workers", "Hires": "Hires",
    "Layoffs": "Layoffs", "Ending Inventory": "Ending Inventory",
    "Ending Backorders": "Ending Backorders", "Labor Cost": "Labor Cost",
    "Hiring Cost": "Hiring Cost", "Layoff Cost": "Layoff Cost", "Holding Cost": "Holding Cost",
    "Backorder Cost": "Backorder Cost", "Total Cost": "Total Cost",
    "Overtime Production": "Overtime Production", "Subcontract Production": "Subcontract Production",
    "Total Production": "Total Production", "Regular Capacity": "Regular Capacity",
    "Overtime Cost": "Overtime Cost", "Subcontract Cost": "Subcontract Cost",
}


def heading_for(kind: str, concept: str) -> str:
    if kind == "hybrid" and concept == "Regular Production":
        return "Regular Production"
    return CANONICAL[concept]


# --------------------------------------------------------------------------- #
# Planning cells (below the grid)
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class PlanningCell:
    key: str
    label: str
    unit: str


PLANNING: Dict[str, List[PlanningCell]] = {
    "chase": [
        PlanningCell("capacity", "Capacity per worker", "bottles/worker/month"),
        PlanningCell("target", "Ending inventory target", "bottles"),
    ],
    "level": [
        PlanningCell("capacity", "Capacity per worker", "bottles/worker/month"),
        PlanningCell("target", "Ending inventory target for December", "bottles"),
        PlanningCell("total_demand", "Total annual demand", "bottles"),
        PlanningCell("level_rate", "Level production per month", "bottles/month"),
        PlanningCell("level_workers", "Level workforce", "workers"),
    ],
    "hybrid": [
        PlanningCell("capacity", "Capacity per worker", "bottles/worker/month"),
        PlanningCell("ot_limit", "Overtime limit per worker", "bottles/worker/month"),
    ],
}


def planning_ref(kind: str, key: str) -> str:
    for i, c in enumerate(PLANNING[kind]):
        if c.key == key:
            return f"{PLANNING_VALUE_COL}{PLANNING_FIRST_ROW + i}"
    raise KeyError(key)


def month_row(i: int) -> int:
    return FIRST_ROW + i


# --------------------------------------------------------------------------- #
# Summary sheet (fixed layout: it is the comparison, not the modelling)
# --------------------------------------------------------------------------- #
MONEY = "$#,##0.00"
SUMMARY_HEAD_ROW = 3
SUMMARY_PLAN_COL = {"chase": "B", "level": "C", "hybrid": "D"}


@dataclass(frozen=True)
class SummaryRow:
    metric: str            # key in engine.summarize()
    label: str
    hint: str
    fmt: str = "#,##0"
    hybrid_only: bool = False


SUMMARY_ROWS: List[SummaryRow] = [
    SummaryRow("Labor Cost", "Total labor cost", "from the plan's yearly totals", MONEY),
    SummaryRow("Hiring Cost", "Total hiring cost", "from the plan's yearly totals", MONEY),
    SummaryRow("Layoff Cost", "Total layoff cost", "from the plan's yearly totals", MONEY),
    SummaryRow("Overtime Cost", "Total overtime cost", "hybrid only", MONEY, True),
    SummaryRow("Subcontract Cost", "Total subcontract cost", "hybrid only", MONEY, True),
    SummaryRow("Holding Cost", "Total holding cost", "from the plan's yearly totals", MONEY),
    SummaryRow("Backorder Cost", "Total backorder cost", "from the plan's yearly totals", MONEY),
    SummaryRow("Total Cost", "TOTAL PLAN COST", "from the plan's yearly totals", MONEY),
    SummaryRow("Total Hires", "Total workers hired", "over the whole year"),
    SummaryRow("Total Layoffs", "Total workers laid off", "over the whole year"),
    SummaryRow("Peak Workforce", "Largest workforce", "the biggest month"),
    SummaryRow("Peak Inventory", "Largest ending inventory (storage needed)", "the biggest month"),
    SummaryRow("Months With Backorders", "Months ending with backorders", "a count of months"),
]
SUMMARY_FIRST_ROW = SUMMARY_HEAD_ROW + 1


def summary_row(metric: str) -> int:
    for i, r in enumerate(SUMMARY_ROWS):
        if r.metric == metric:
            return SUMMARY_FIRST_ROW + i
    raise KeyError(metric)


SUMMARY_LAST_METRIC_ROW = SUMMARY_FIRST_ROW + len(SUMMARY_ROWS) - 1
LOWER_COST_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 2}"
SAVINGS_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 3}"
RECOMMEND_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 5}"
JUSTIFY_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 6}"
JUSTIFY_LAST_ROW = SUMMARY_LAST_METRIC_ROW + 14

START_NAME_CELL = "C4"
START_ID_CELL = "C5"
