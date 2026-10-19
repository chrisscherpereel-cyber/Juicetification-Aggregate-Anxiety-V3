"""Where everything lives in the student workbook.

The workbook builder (workbook.py) writes these cells and the checker (checker.py) reads
them, so both sides always agree. Every plan sheet uses the same simple addressing:
row 1 = headings, January = row 2 ... December = row 13, row 14 = totals."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter

FIRST_ROW = 2                     # January
N_MONTHS = 12
LAST_ROW = FIRST_ROW + N_MONTHS - 1   # December = 13
TOTAL_ROW = LAST_ROW + 1              # 14

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


@dataclass(frozen=True)
class Column:
    heading: str          # shown in row 1 of the sheet
    engine: str           # column in engine.simulate() output
    unit: str
    kind: str             # "given" | "calc" | "decision"
    total: bool = False   # row 14 holds a SUM the student must build
    fmt: str = "#,##0"


MONEY = "$#,##0.00"


def _plan_columns(production_heading: str) -> List[Column]:
    return [
        Column("Month", "Month", "", "given", fmt="@"),
        Column("Demand", "Demand", "bottles", "given"),
        Column("Beginning Inventory", "Beginning Inventory", "bottles", "calc"),
        Column("Beginning Backorders", "Beginning Backorders", "bottles", "calc"),
        Column(production_heading, "Regular Production", "bottles", "calc", total=True),
        Column("Workers", "Workers", "workers", "calc", fmt="#,##0"),
        Column("Hires", "Hires", "workers", "calc", total=True),
        Column("Layoffs", "Layoffs", "workers", "calc", total=True),
        Column("Ending Inventory", "Ending Inventory", "bottles", "calc"),
        Column("Ending Backorders", "Ending Backorders", "bottles", "calc"),
        Column("Labor Cost", "Labor Cost", "$", "calc", True, MONEY),
        Column("Hiring Cost", "Hiring Cost", "$", "calc", True, MONEY),
        Column("Layoff Cost", "Layoff Cost", "$", "calc", True, MONEY),
        Column("Holding Cost", "Holding Cost", "$", "calc", True, MONEY),
        Column("Backorder Cost", "Backorder Cost", "$", "calc", True, MONEY),
        Column("Total Cost", "Total Cost", "$", "calc", True, MONEY),
    ]


PLAN_COLUMNS: Dict[str, List[Column]] = {
    "chase": _plan_columns("Production"),
    "level": _plan_columns("Production"),
    "hybrid": [
        Column("Month", "Month", "", "given", fmt="@"),
        Column("Demand", "Demand", "bottles", "given"),
        Column("Beginning Inventory", "Beginning Inventory", "bottles", "calc"),
        Column("Beginning Backorders", "Beginning Backorders", "bottles", "calc"),
        Column("Workers ✎", "Workers", "workers (your decision)", "decision"),
        Column("Regular Capacity", "Regular Capacity", "bottles", "calc", total=True),
        Column("Regular Production ✎", "Regular Production", "bottles (your decision)",
               "decision", total=True),
        Column("Overtime Production ✎", "Overtime Production", "bottles (your decision)",
               "decision", total=True),
        Column("Subcontract Production ✎", "Subcontract Production",
               "bottles (your decision)", "decision", total=True),
        Column("Total Production", "Total Production", "bottles", "calc", total=True),
        Column("Hires", "Hires", "workers", "calc", total=True),
        Column("Layoffs", "Layoffs", "workers", "calc", total=True),
        Column("Ending Inventory", "Ending Inventory", "bottles", "calc"),
        Column("Ending Backorders", "Ending Backorders", "bottles", "calc"),
        Column("Labor Cost", "Labor Cost", "$", "calc", True, MONEY),
        Column("Hiring Cost", "Hiring Cost", "$", "calc", True, MONEY),
        Column("Layoff Cost", "Layoff Cost", "$", "calc", True, MONEY),
        Column("Overtime Cost", "Overtime Cost", "$", "calc", True, MONEY),
        Column("Subcontract Cost", "Subcontract Cost", "$", "calc", True, MONEY),
        Column("Holding Cost", "Holding Cost", "$", "calc", True, MONEY),
        Column("Backorder Cost", "Backorder Cost", "$", "calc", True, MONEY),
        Column("Total Cost", "Total Cost", "$", "calc", True, MONEY),
    ],
}


def col_letter(kind: str, engine_name: str) -> str:
    for j, c in enumerate(PLAN_COLUMNS[kind]):
        if c.engine == engine_name:
            return get_column_letter(j + 1)
    raise KeyError(engine_name)


@dataclass(frozen=True)
class PlanningCell:
    key: str
    label: str
    unit: str
    hint: str


# Planning calculations block beside each plan grid: label in column `label_col`,
# the student's formula in column `value_col`, starting at row 2.
PLANNING: Dict[str, Tuple[str, str, List[PlanningCell]]] = {
    "chase": ("R", "S", [
        PlanningCell("capacity", "Capacity per worker", "bottles/worker/month",
                     "=BottlesPerHour*HoursPerDay*DaysPerMonth"),
        PlanningCell("target", "Ending inventory target", "bottles", "=EndTarget"),
    ]),
    "level": ("R", "S", [
        PlanningCell("capacity", "Capacity per worker", "bottles/worker/month",
                     "=BottlesPerHour*HoursPerDay*DaysPerMonth"),
        PlanningCell("target", "Ending inventory target (December)", "bottles", "=EndTarget"),
        PlanningCell("total_demand", "Total annual demand", "bottles", "SUM of the Demand column"),
        PlanningCell("level_rate", "Level production per month", "bottles/month",
                     "(total demand + target − BeginInv) ÷ 12"),
        PlanningCell("level_workers", "Level workforce", "workers",
                     "ROUNDUP(level production ÷ capacity per worker, 0)"),
    ]),
    "hybrid": ("X", "Y", [
        PlanningCell("capacity", "Capacity per worker", "bottles/worker/month",
                     "=BottlesPerHour*HoursPerDay*DaysPerMonth"),
        PlanningCell("ot_limit", "Overtime limit per worker", "bottles/worker/month",
                     "capacity per worker × OTLimitPct"),
    ]),
}


def planning_ref(kind: str, key: str) -> str:
    _, vcol, cells = PLANNING[kind]
    for i, c in enumerate(cells):
        if c.key == key:
            return f"{vcol}{FIRST_ROW + i}"
    raise KeyError(key)


# --------------------------------------------------------------------------- #
# Summary sheet
# --------------------------------------------------------------------------- #
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
    SummaryRow("Labor Cost", "Total labor cost", "link to the plan's Labor Cost total (row 14)", MONEY),
    SummaryRow("Hiring Cost", "Total hiring cost", "link to the Hiring Cost total", MONEY),
    SummaryRow("Layoff Cost", "Total layoff cost", "link to the Layoff Cost total", MONEY),
    SummaryRow("Overtime Cost", "Total overtime cost", "hybrid only", MONEY, True),
    SummaryRow("Subcontract Cost", "Total subcontract cost", "hybrid only", MONEY, True),
    SummaryRow("Holding Cost", "Total holding cost", "link to the Holding Cost total", MONEY),
    SummaryRow("Backorder Cost", "Total backorder cost", "link to the Backorder Cost total", MONEY),
    SummaryRow("Total Cost", "TOTAL PLAN COST", "link to the Total Cost total", MONEY),
    SummaryRow("Total Hires", "Total workers hired", "link to the Hires total"),
    SummaryRow("Total Layoffs", "Total workers laid off", "link to the Layoffs total"),
    SummaryRow("Peak Workforce", "Largest workforce", "MAX of the Workers column"),
    SummaryRow("Peak Inventory", "Largest ending inventory (storage needed)",
               "MAX of the Ending Inventory column"),
    SummaryRow("Months With Backorders", "Months ending with backorders",
               'COUNTIF(Ending Backorders column,">0")'),
]
SUMMARY_FIRST_ROW = SUMMARY_HEAD_ROW + 1


def summary_row(metric: str) -> int:
    for i, r in enumerate(SUMMARY_ROWS):
        if r.metric == metric:
            return SUMMARY_FIRST_ROW + i
    raise KeyError(metric)


SUMMARY_LAST_METRIC_ROW = SUMMARY_FIRST_ROW + len(SUMMARY_ROWS) - 1
LOWER_COST_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 2}"      # "Chase" or "Level" via IF
SAVINGS_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 3}"         # ABS(chase − level)
RECOMMEND_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 5}"       # dropdown
JUSTIFY_CELL = f"B{SUMMARY_LAST_METRIC_ROW + 6}"         # free text (merged area)
JUSTIFY_LAST_ROW = SUMMARY_LAST_METRIC_ROW + 14

# Start Here: identification cells (pre-filled, informational)
START_NAME_CELL = "C4"
START_ID_CELL = "C5"


def month_row(i: int) -> int:
    return FIRST_ROW + i


def cell_ref(kind: str, engine_name: str, i: int) -> str:
    return f"{col_letter(kind, engine_name)}{month_row(i)}"


def total_ref(kind: str, engine_name: str) -> str:
    return f"{col_letter(kind, engine_name)}{TOTAL_ROW}"


def column_by_letter(kind: str, letter: str) -> Optional[Column]:
    from openpyxl.utils import column_index_from_string
    j = column_index_from_string(letter) - 1
    cols = PLAN_COLUMNS[kind]
    return cols[j] if 0 <= j < len(cols) else None
