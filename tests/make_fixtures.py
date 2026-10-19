"""Regenerate the test fixtures (Windows + Excel only).

The checker reads the values EXCEL saved with a file, so the fixtures must be calculated by
real Excel, not by openpyxl. This script writes the workbooks with openpyxl and then runs
tools/recalc_in_excel.ps1 on each:

  python tests/make_fixtures.py

  key_seed7.xlsx       complete answer key incl. a hybrid plan (Excel-calculated)
  mistakes_seed7.xlsx  typical slips: a typed number, a missing $, a short justification
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

from aggplan import workbook  # noqa: E402
from aggplan.engine import Decision  # noqa: E402
from aggplan.scenario import resolve_scenario  # noqa: E402

SEED, SID, NAME = 7, "S1", "Test Student"
HYBRID = Decision([11] * 12, [11000] * 12, [0] * 5 + [1000, 2000] + [0] * 5, [0] * 12)
FIX = os.path.join(HERE, "fixtures")


def recalc(src: str, dst: str) -> None:
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                    os.path.join(ROOT, "tools", "recalc_in_excel.ps1"), "-src", src, "-dst", dst],
                   check=True)


def main() -> None:
    os.makedirs(FIX, exist_ok=True)
    scn = resolve_scenario({}, SEED)
    key = workbook.build_workbook(scn, student_name=NAME, student_id=SID, solution=True,
                                  hybrid_decision=HYBRID)
    raw = os.path.join(FIX, "_raw_key.xlsx")
    open(raw, "wb").write(key)
    recalc(raw, os.path.join(FIX, "key_seed7.xlsx"))

    wb = load_workbook(io.BytesIO(key))
    wb["Chase Plan"]["N3"].value = 0                                   # typed, not calculated
    for r in range(2, 14):                                             # =S5 filled down, no $
        wb["Level Plan"][f"E{r}"].value = f"=S{3 + r}"
    wb["Summary"]["B22"].value = "Too short."
    wb.save(raw)
    recalc(raw, os.path.join(FIX, "mistakes_seed7.xlsx"))
    os.remove(raw)


if __name__ == "__main__":
    main()
