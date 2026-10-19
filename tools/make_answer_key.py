#!/usr/bin/env python
"""Write an instructor ANSWER KEY workbook: the student layout filled with reference
formulas. Open it in Excel and every value calculates.

  python tools/make_answer_key.py --sid A1234567 [--game OPS301-F26] [--cfg cfg.json]
  python tools/make_answer_key.py --seed 207659

--sid (with the same --game as the student's link) reproduces that student's scenario;
--cfg is the instructor configuration as JSON (the same parameters as manifest.py).
The file is written next to the current directory unless --out is given."""

from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import student_store as store  # noqa: E402
from aggplan import workbook  # noqa: E402
from aggplan.scenario import resolve_scenario  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--sid", help="student ID (scenario derived exactly as in the app)")
    g.add_argument("--seed", type=int, help="scenario number shown in the app / workbook")
    ap.add_argument("--game", default=None, help="game code from the student's link, if any")
    ap.add_argument("--cfg", help="JSON file with instructor parameters")
    ap.add_argument("--out", help="output .xlsx path")
    a = ap.parse_args()
    cfg = json.load(open(a.cfg, encoding="utf-8")) if a.cfg else {}
    seed = a.seed if a.seed is not None else store.derive_seed(a.game, a.sid, lo=1, hi=10 ** 6)
    scn = resolve_scenario(cfg, seed)
    data = workbook.build_workbook(scn, student_name="ANSWER KEY", student_id=a.sid or "",
                                   game=a.game, solution=True)
    out = a.out or workbook.file_name(a.sid or f"scenario_{seed}", "", solution=True)
    with open(out, "wb") as f:
        f.write(data)
    print(f"Scenario #{seed} -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
