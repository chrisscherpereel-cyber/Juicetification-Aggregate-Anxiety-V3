"""Guided hybrid walkthrough: a starting plan plus questions that lead to a good solution.

The student first chooses which of their two plans to start from (the cheaper one), gets that
plan pre-built on the Hybrid Plan sheet, and is then walked — one question at a time — to the
improvement that plan allows. Every number in a question comes from the student's own
scenario; the student does the arithmetic and makes the changes in Excel.

Tracks (chosen from the scenario, not hard-coded):
  chase     the chase plan pays for whole workers who are only partly used. In months where
            the bottles beyond whole workers are few, one worker fewer plus overtime is cheaper
            (overtime cost < a month of labor) — tested to beat both pure plans.
  no_gain   nothing to fix that way (fractional workers, or no month where it pays): the
            lesson is WHY — a regular-time bottle costs less than an overtime bottle — and the
            chase plan is already the best mix.
  level     the level plan leaves customers waiting in the peak: use the paid-but-idle
            capacity first, then the cheapest extra source (usually overtime), then
            subcontracting, only as much as needed.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Union

import pandas as pd

from . import engine
from .scenario import Scenario


@dataclass
class Question:
    qid: str
    prompt: str
    kind: str                                  # "choice" | "number" | "months"
    answer: Union[str, float, frozenset]
    why: str                                   # shown once answered (or revealed)
    nudge: str = ""                            # shown after a wrong try
    options: Sequence[str] = ()
    tol: float = 0.5


@dataclass
class Task:
    qid: str
    title: str
    text: str


Step = Union[Question, Task]


def _money(x: float) -> str:
    return f"${x:,.2f}" if abs(x - round(x)) > 1e-9 else f"${x:,.0f}"


def _n(x: float) -> str:
    return f"{x:,.2f}".rstrip("0").rstrip(".") if abs(x - round(x)) > 1e-9 else f"{x:,.0f}"


# --------------------------------------------------------------------------- #
# Starting point
# --------------------------------------------------------------------------- #
def plan_costs(scn: Scenario) -> dict:
    return {k: float(engine.reference_plan(scn, k)["Total Cost"].sum()) for k in ("chase", "level")}


def best_start(scn: Scenario) -> str:
    c = plan_costs(scn)
    return "chase" if c["chase"] <= c["level"] else "level"


def start_question(scn: Scenario) -> Question:
    c = plan_costs(scn)
    best = best_start(scn)
    other = "level" if best == "chase" else "chase"
    return Question(
        "start", "Which of your two plans should the hybrid start from?", "choice",
        best.title(),
        why=(f"Start from the cheaper plan — your {best} plan ({_money(c[best])} vs "
             f"{_money(c[other])}). A hybrid improves a plan by fixing its most expensive "
             "habit, so the closer you start to the best answer, the less there is to fix."),
        nudge="Look at the two total costs. A hybrid improves on a plan — which one is closer "
              "to the best answer already?",
        options=("Chase", "Level"))


def starting_decision(scn: Scenario, start: str) -> engine.Decision:
    d = engine.reference_chase(scn) if start == "chase" else engine.reference_level(scn)
    n = len(scn.months)
    return engine.Decision(list(d.workers), list(d.regular), [0.0] * n, [0.0] * n)


# --------------------------------------------------------------------------- #
# Tracks
# --------------------------------------------------------------------------- #
def _round_down_months(scn: Scenario) -> List[dict]:
    """Months of the chase plan where one worker fewer + overtime is cheaper and allowed."""
    p, cap = scn.params, scn.cap_month
    if not scn.whole_workers:
        return []
    df = engine.reference_plan(scn, "chase")
    out = []
    for i, r in df.iterrows():
        w, prod = float(r["Workers"]), float(r["Regular Production"])
        if w < 1 or prod <= 0:
            continue
        rem = prod - (w - 1) * cap
        if rem <= 0 or rem > (w - 1) * cap * p["overtime_pct"] + 1e-9:
            continue
        saving = p["regular_labor_cost"] - rem * p["overtime_cost"]
        if saving > 0:
            out.append({"i": i, "month": scn.months[i], "workers": w, "prod": prod,
                        "idle": w * cap - prod, "rem": rem, "ot_cost": rem * p["overtime_cost"],
                        "saving": saving})
    return out


def track_name(scn: Scenario, start: str) -> str:
    if start == "level":
        return "level"
    return "chase" if _round_down_months(scn) else "no_gain"


def steps(scn: Scenario, start: str) -> List[Step]:
    p, cap = scn.params, scn.cap_month
    labor, ot, sub = p["regular_labor_cost"], p["overtime_cost"], p["subcontract_cost"]
    track = track_name(scn, start)
    out: List[Step] = []
    if track == "chase":
        months = _round_down_months(scn)
        m = max(months, key=lambda r: r["saving"])
        w, prod = m["workers"], m["prod"]
        out += [
            Question("c_idle",
                     f"In {m['month']} your starting plan makes {_n(prod)} bottles with "
                     f"{_n(w)} workers, and one worker can make {_n(cap)} bottles a month. "
                     "How many bottles of paid capacity sit idle that month?",
                     "number", w * cap - prod,
                     why=f"{_n(w)} × {_n(cap)} = {_n(w * cap)} bottles of capacity for "
                         f"{_n(prod)} bottles of work: the last worker is only partly busy but "
                         f"is paid a full month ({_money(labor)}).",
                     nudge="Capacity = workers × bottles per worker. Idle = capacity − production."),
            Question("c_otcost",
                     f"Suppose you employ {_n(w - 1)} workers in {m['month']} instead and make "
                     "the bottles they can't cover on overtime. What would those overtime "
                     "bottles cost in total ($)?",
                     "number", m["ot_cost"],
                     why=f"{_n(w - 1)} workers make {_n((w - 1) * cap)}; the other "
                         f"{_n(m['rem'])} bottles on overtime cost {_n(m['rem'])} × "
                         f"{_money(ot)} = {_money(m['ot_cost'])}.",
                     nudge="Bottles still needed = production − (workers − 1) × capacity. Each "
                           "costs OTCost on overtime."),
            Question("c_compare", f"Which is cheaper for {m['month']}?", "choice",
                     "Overtime instead",
                     why=f"Overtime costs {_money(m['ot_cost'])}, a month of the extra worker "
                         f"{_money(labor)} — you save {_money(m['saving'])}, and employing one "
                         "worker fewer can also save a hire and a layoff around that month.",
                     nudge="Compare the overtime cost you just worked out with one month of "
                           "regular labor (LaborCost).",
                     options=(f"Keep worker #{_n(w)}", "Overtime instead")),
            Question("c_rule", "So when is overtime cheaper than employing one more worker?",
                     "choice", "few",
                     why=f"Overtime wins while (bottles needed beyond whole workers) × "
                         f"{_money(ot)} < {_money(labor)} — fewer than about "
                         f"{_n(math.floor(labor / ot))} bottles — and only if they fit the "
                         f"overtime limit ({p['overtime_pct']:.0%} of the smaller crew's "
                         "capacity). Beyond that, the extra worker is cheaper.",
                     nudge="Overtime costs more per bottle than regular labor. When can it "
                           "still be the cheaper choice?",
                     options=("few", "always", "never", "peak")),
            Task("c_apply", "Make the change in Excel",
                 "On the Hybrid Plan sheet, go month by month. Where the bottles needed beyond "
                 f"whole workers are fewer than about {_n(math.floor(labor / ot))} AND fit "
                 "the overtime limit: set Workers one lower, Regular Production = that crew's "
                 "full capacity, and Overtime = the bottles still needed. Leave the other "
                 "months alone. Save in Excel and upload."),
        ]
    elif track == "no_gain":
        per = labor / cap
        out += [
            Question("n_bottle",
                     f"A worker costs {_money(labor)} a month and makes {_n(cap)} bottles. What "
                     "does one regular-time bottle cost in labor ($)?",
                     "number", per, tol=0.01,
                     why=f"{_money(labor)} ÷ {_n(cap)} = {_money(per)} per bottle.",
                     nudge="Divide a month of labor by the bottles one worker makes."),
            Question("n_cheaper",
                     f"An overtime bottle costs {_money(ot)}. Which is cheaper per bottle?",
                     "choice", "Regular time" if per <= ot else "Overtime",
                     why=("Regular time is cheaper per bottle, so overtime only pays when it "
                          "saves paying for a mostly idle worker. "
                          + ("With fractional workers your chase plan never pays for idle "
                             "capacity" if not scn.whole_workers else
                             "Your chase plan has no month where that happens")
                          + " — it is already the best mix. Submitting it as your hybrid is a "
                            "good solution; try one experiment first and watch the cost rise."),
                     nudge="Compare the two per-bottle costs.",
                     options=("Regular time", "Overtime")),
            Task("n_apply", "Check it in Excel",
                 "Upload the starting plan as it is (it is feasible). If you like, first try "
                 "moving some regular production to overtime in one month and see the total "
                 "cost go up; then put it back."),
        ]
    else:                                                          # level
        df = engine.reference_plan(scn, "level")
        w, rate = float(df["Workers"].iloc[0]), float(df["Regular Production"].iloc[0])
        short = frozenset(scn.months[i] for i in range(len(df))
                          if df["Ending Backorders"].iloc[i] > 1e-6)
        hire = (p["hiring_cost"] + p["layoff_cost"] + labor) / cap
        srcs = {"Overtime": ot, "Subcontract": sub, "Hire a worker for a month": hire}
        cheapest = min(srcs, key=srcs.get)
        out += [
            Question("l_months", "Which months end with customers still waiting (backorders) "
                                 "in your starting plan?", "months", short,
                     why="Those months are where the level plan breaks the on-time requirement "
                         "— the hybrid's first job is to fix them.",
                     nudge="Look at your level plan's Ending Backorders column."),
            Question("l_spare",
                     f"Your level crew of {_n(w)} could make {_n(w * cap)} bottles a month but "
                     f"makes only {_n(rate)}. How many more bottles a month could they make on "
                     "regular time?", "number", w * cap - rate,
                     why="That capacity is already paid for, so using it adds no labor cost — "
                         "always use it before buying extra bottles.",
                     nudge="Capacity = workers × bottles per worker; compare with production."),
            Question("l_next", "If that isn't enough, which is the cheapest source of more "
                               "bottles?", "choice", cheapest,
                     why=" · ".join(f"{k}: {_money(v)}/bottle" for k, v in srcs.items())
                         + f". Use {cheapest.lower()} first, then the next cheapest.",
                     nudge="Compare the cost of one more bottle each way (hiring for a month = "
                           "hire + layoff + a month of labor, spread over one worker's bottles).",
                     options=tuple(srcs)),
            Question("l_otlimit", "How many bottles may your crew make on overtime in one "
                                  "month?", "number", w * cap * p["overtime_pct"],
                     why=f"{_n(w)} × {_n(cap)} × {p['overtime_pct']:.0%} = "
                         f"{_n(w * cap * p['overtime_pct'])} bottles a month, and only once "
                         "regular capacity is full.",
                     nudge="Overtime limit = workers × capacity per worker × OTLimitPct."),
            Task("l_apply", "Make the change in Excel",
                 "On the Hybrid Plan sheet, starting with the first month that ends with "
                 "backorders: raise Regular Production up to full capacity; if that month is "
                 "still short, add overtime (up to the limit), then subcontracting. Add only "
                 "as much as needed to stop the backorders. Save in Excel and upload."),
        ]
    return out


# --------------------------------------------------------------------------- #
# Evaluating the student's current hybrid
# --------------------------------------------------------------------------- #
@dataclass
class Verdict:
    feasible: bool
    cost: float
    start_cost: float
    chase: float
    level: float
    good: bool
    notes: List[str] = field(default_factory=list)


def evaluate(scn: Scenario, start: str, plan: pd.DataFrame) -> Verdict:
    c = plan_costs(scn)
    checks = engine.check_feasibility(plan, scn)
    feasible = all(x["ok"] for x in checks)
    cost = float(plan["Total Cost"].sum())
    start_cost = c[start]
    improvable = track_name(scn, start) != "no_gain"
    good = (feasible and cost <= min(c.values()) + 0.5
            and (not improvable or cost < start_cost - 0.5))
    notes = [f"✖ {x['label']}: {x['detail']}" for x in checks if not x["ok"]]
    if feasible and improvable and abs(cost - start_cost) <= 0.5:
        notes.append("This is still your starting plan's cost — make the change described in "
                     "step 4, save, and upload again.")
    if feasible and not good:
        base = engine.reference_plan(scn, start)
        diffs = {k: float(plan[k].sum() - base[k].sum()) for k in
                 ("Labor Cost", "Hiring Cost", "Layoff Cost", "Overtime Cost",
                  "Subcontract Cost", "Holding Cost", "Backorder Cost")}
        up = max(diffs, key=diffs.get)
        if diffs[up] > 0.5:
            notes.append(f"Compared with your starting plan, {up.lower()} went up most "
                         f"(+{_money(diffs[up])}). Is each of those changes paying for itself?")
    return Verdict(feasible, cost, start_cost, c["chase"], c["level"], good, notes)


def reflection(start: str, track: str) -> Question:
    right = {"chase": "It stopped paying a full month of labor for a worker needed for only a "
                      "few hundred bottles.",
             "no_gain": "Regular time is the cheapest way to make a bottle, and the chase plan "
                        "already wastes almost none of it.",
             "level": "It used capacity it was already paying for, then the cheapest extra "
                      "bottles, instead of leaving customers waiting."}[track]
    import hashlib
    opts = tuple(sorted((right, "Overtime is always cheaper than regular labor.",
                         "It holds more inventory, and inventory is free.",
                         "It hires more workers so nobody works overtime."),
                        key=lambda o: hashlib.sha256(o.encode()).hexdigest()))
    return Question("reflect", "Why is your hybrid a good solution?", "choice", right,
                    why="That's the trade-off a hybrid exploits: use each source of capacity "
                        "only where it is the cheapest one.",
                    nudge="Think about which cost your changes actually reduced.",
                    options=opts)


def check_answer(q: Question, value) -> bool:
    if value is None:
        return False
    if q.kind == "number":
        try:
            return abs(float(value) - float(q.answer)) <= max(q.tol, 1e-6 * abs(float(q.answer)))
        except (TypeError, ValueError):
            return False
    if q.kind == "months":
        return frozenset(value) == q.answer
    return str(value) == str(q.answer)


RULE_LABELS = {
    "few": "When only a few bottles are needed beyond whole workers (their overtime cost is "
           "less than a month of labor) — and they fit the overtime limit",
    "always": "Always — overtime needs no hiring",
    "never": "Never — regular labor is cheaper per bottle",
    "peak": "Only in the peak month",
}
