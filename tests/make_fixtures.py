"""Regenerate the test fixtures (Windows + Excel only).

The checker reads the values EXCEL saved with a file, so the fixtures must be calculated by
real Excel, not by openpyxl. This script writes the workbooks with openpyxl and then runs
tools/recalc_in_excel.ps1 on each:

  python tests/make_fixtures.py

  key_seed7.xlsx       complete answer key incl. a hybrid plan
  mistakes_seed7.xlsx  typical slips: a typed number, a missing $, a short justification
  shuffled_seed7.xlsx  chase columns in a student's own order with their own headings
"""

from __future__ import annotations

import io
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from aggplan import layout as L, workbook  # noqa: E402
from aggplan.engine import Decision  # noqa: E402
from aggplan.scenario import resolve_scenario  # noqa: E402

SEED, SID, NAME = 7, "S1", "Test Student"
HYBRID = Decision([11] * 12, [11000] * 12, [0] * 5 + [1000, 2000] + [0] * 5, [0] * 12)
FIX = os.path.join(HERE, "fixtures")

# A student's own design: different order, different words, one helper column.
SHUFFLED_CHASE = [
    ("Beginning Inventory", "Starting Stock"), ("Beginning Backorders", "Backlog Carried In"),
    ("Regular Production", "Bottles Made"), ("Workers", "Crew"),
    ("Regular Capacity", "Regular Capacity"),
    ("Hires", "New Hires"), ("Layoffs", "Fired"),
    ("Ending Backorders", "Unfilled Orders"), ("Ending Inventory", "Closing Stock"),
    ("Labor Cost", "Wages"), ("Holding Cost", "Storage Cost"),
    ("Backorder Cost", "Shortage Penalty"), ("Hiring Cost", "Hiring Expense"),
    ("Layoff Cost", "Severance"), ("Total Cost", "Monthly Total Cost"),
]


def recalc(src: str, dst: str) -> None:
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                    os.path.join(ROOT, "tools", "recalc_in_excel.ps1"), "-src", src, "-dst", dst],
                   check=True)


def main() -> None:
    os.makedirs(FIX, exist_ok=True)
    scn = resolve_scenario({}, SEED)
    raw = os.path.join(FIX, "_raw.xlsx")

    key = workbook.build_workbook(scn, student_name=NAME, student_id=SID, solution=True,
                                  hybrid_decision=HYBRID)
    open(raw, "wb").write(key)
    recalc(raw, os.path.join(FIX, "key_seed7.xlsx"))

    wb = load_workbook(io.BytesIO(key))
    wb[L.SHEET_CHASE]["N3"].value = 0                                   # typed, not calculated
    rate = L.planning_ref("level", "level_rate")                       # =C20 filled down, no $
    for i in range(L.N_MONTHS):
        wb[L.SHEET_LEVEL][f"E{L.month_row(i)}"].value = f"=C{int(rate[1:]) + i}"
    wb[L.SHEET_SUMMARY][L.JUSTIFY_CELL].value = "Too short."
    wb.save(raw)
    recalc(raw, os.path.join(FIX, "mistakes_seed7.xlsx"))

    shuffled = workbook.build_workbook(scn, student_name=NAME, student_id=SID, solution=True,
                                       hybrid_decision=HYBRID, columns={"chase": SHUFFLED_CHASE})
    open(raw, "wb").write(shuffled)
    recalc(raw, os.path.join(FIX, "shuffled_seed7.xlsx"))
    os.remove(raw)


if __name__ == "__main__":
    main()
