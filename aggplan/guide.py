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


def _chase_checklist(scn: Scenario, m: dict) -> str:
    p, cap = scn.params, scn.cap_month
    brk = math.floor(p["regular_labor_cost"] / p["overtime_cost"])
    w, prod, rem = m["workers"], m["prod"], m["rem"]
    lim = (w - 1) * cap * p["overtime_pct"]
    return (
        "On the **Hybrid Plan** sheet, check every month.\n\n"
        "**Change only the green cells, and type the new numbers in.** A formula in a green "
        "cell that uses a column calculated from the green cells (Total Production, Ending "
        "Inventory, costs …) makes a circular reference (Excel then "
        "shows 0). Don't add a Leftover column inside your plan either — the grey helper "
        "columns on the right hold your starting plan's fixed numbers for that.\n\n"
        "1. **Leftover** = Regular Production − (Workers − 1) × capacity per worker — the "
        "bottles you keep the last worker for. It's already worked out in the grey column "
        "**Helper: leftover with one worker fewer** (from your starting plan, so it doesn't "
        "change as you edit).\n"
        f"2. If the leftover is **under about {_n(brk)} bottles** AND **no more than the overtime "
        f"limit for one worker fewer** ((Workers − 1) × {_n(cap)} × {p['overtime_pct']:.0%}), "
        "change **three cells** in that month:\n"
        "   - **Workers** → one fewer\n"
        f"   - **Regular Production** → the new Workers × {_n(cap)} (a smaller crew makes less "
        "on regular time — this is the step that's easy to forget)\n"
        "   - **Overtime Production** → the leftover\n"
        "3. Otherwise leave the month exactly as it is.\n\n"
        f"**Worked example — {m['month']}:** {_n(w)} workers, {_n(prod)} bottles. Leftover = "
        f"{_n(prod)} − {_n(w - 1)} × {_n(cap)} = **{_n(rem)}** — under {_n(brk)}, and within "
        f"the limit of {_n(lim)}. So Workers {_n(w)} → **{_n(w - 1)}**, Regular Production "
        f"{_n(prod)} → **{_n((w - 1) * cap)}**, Overtime 0 → **{_n(rem)}**.\n\n"
        "Save in Excel and upload. The app then checks every month and tells you what's left.")


def _level_checklist(scn: Scenario, w: float, rate: float) -> str:
    p, cap = scn.params, scn.cap_month
    return (
        "On the **Hybrid Plan** sheet, go through the months **in order** (stock and "
        "backorders carry forward, so a change in one month affects the next).\n\n"
        "**Change only the green cells, and type the new numbers in.** A formula in a green "
        "cell that uses a column calculated from the green cells (Total Production, Ending "
        "Inventory, costs …) makes a circular reference (Excel then "
        "shows 0). The grey helper columns on the right keep your starting plan for "
        "comparison.\n\n"
        "1. Find the first month that ends with backorders.\n"
        f"2. Raise its **Regular Production** — up to the crew's capacity of {_n(w * cap)} "
        f"(you make {_n(rate)}, so up to {_n(w * cap - rate)} more, already paid for).\n"
        "3. Still short? Add **Overtime Production** — only once Regular Production is at "
        f"capacity, and at most {_n(w * cap * p['overtime_pct'])} a month.\n"
        "4. Still short? Add **Subcontract Production**, at most "
        f"{_n(p['subcontract_capacity'])} a month.\n"
        "5. Add only what that month needs to end with no backorders, then move to the next "
        "month that still ends with backorders.\n\n"
        "Save in Excel and upload. The app then checks every month and tells you what's left.")


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
            Task("c_apply", "Make the change in Excel", _chase_checklist(scn, m)),
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
            Task("l_apply", "Make the change in Excel", _level_checklist(scn, w, rate)),
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


def coach(scn: Scenario, start: str, plan: pd.DataFrame, level: int = 1):
    """Month-by-month help on the student's current hybrid.

    Returns (fixes, opportunities). `fixes` names every month that breaks a limit and what to
    change there, in terms of the student's own numbers — always shown, because a plan that
    breaks a limit cannot be submitted. `opportunities` (only once the plan meets every limit
    but isn't a good solution yet) say where money is still left, more specifically as
    `level` rises: 1 a nudge, 2 which months, 3 the change for each month."""
    p, cap, eps = scn.params, scn.cap_month, 1e-6
    pct = float(p["overtime_pct"])
    fixes: List[str] = []
    for i, r in plan.iterrows():
        m = scn.months[i]
        w, reg = float(r["Workers"]), float(r["Regular Production"])
        ot, sub = float(r["Overtime Production"]), float(r["Subcontract Production"])
        capw = w * cap
        if scn.whole_workers and abs(w - round(w)) > eps:
            fixes.append(f"**{m}:** {_n(w)} workers — your policy is whole workers only.")
        if reg > capw + eps:
            fixes.append(
                f"**{m}:** {_n(w)} workers can make only {_n(capw)} bottles on regular time, but "
                f"Regular Production is {_n(reg)}. Set Regular Production to **{_n(capw)}** and "
                f"Overtime to **{_n(ot + reg - capw)}** (the {_n(reg - capw)} bottles the crew "
                "can't make on regular time).")
            continue                       # the overtime checks below follow from this one
        if ot > eps and capw - reg > eps:
            fixes.append(
                f"**{m}:** {_n(ot)} bottles on overtime while {_n(capw - reg)} bottles of regular "
                f"capacity sit idle. Make them on regular time first: Regular Production up to "
                f"{_n(capw)}, overtime down by the same amount.")
        if ot > capw * pct + eps:
            fixes.append(
                f"**{m}:** overtime is {_n(ot)} but the limit is {_n(capw * pct)} "
                f"({pct:.0%} of {_n(capw)}). Keep one more worker this month (Workers "
                f"{_n(w + 1)}, Regular Production up to {_n((w + 1) * cap)}) and cut overtime to "
                "what is still needed.")
        if sub > p["subcontract_capacity"] + eps:
            fixes.append(f"**{m}:** subcontracting {_n(sub)} is over the limit of "
                         f"{_n(p['subcontract_capacity'])}.")
        if float(r["Ending Inventory"]) > p["max_inventory"] + eps:
            fixes.append(f"**{m}:** {_n(r['Ending Inventory'])} bottles in stock — more than the "
                         f"{_n(p['max_inventory'])} you can store. Make less in or before {m}.")
    checks = {c["label"]: c for c in engine.check_feasibility(plan, scn)}
    late = [scn.months[i] for i in range(len(plan)) if plan["Ending Backorders"].iloc[i] > eps]
    service_bad = any(not c["ok"] for k, c in checks.items()
                      if k.startswith("On-time") or k.startswith("December backorders"))
    if late and (service_bad or not fixes):
        fixes.append(f"**Customers left waiting** at the end of {', '.join(late)}. Those months "
                     "need more bottles — from idle regular capacity first, then overtime.")
    last = plan.iloc[-1]
    if float(last["Ending Inventory"]) < scn.target - eps:
        fixes.append(f"**December** ends with {_n(last['Ending Inventory'])} bottles; your "
                     f"safety-stock policy needs at least {_n(scn.target)}.")

    opps: List[str] = []
    v_ok = all(c["ok"] for c in checks.values())
    if not v_ok or fixes:
        return fixes, opps
    cost = float(plan["Total Cost"].sum())
    costs = plan_costs(scn)
    track = track_name(scn, start)
    if track == "chase":
        base = engine.reference_plan(scn, "chase")
        qual = {r["i"]: r for r in _round_down_months(scn)}
        todo = [i for i in qual if abs(plan["Workers"].iloc[i] - base["Workers"].iloc[i]) < eps]
        bad = [i for i in range(len(plan)) if plan["Workers"].iloc[i] < base["Workers"].iloc[i]
               - eps and i not in qual]
        if cost < min(costs.values()) - 0.5 and not todo and not bad:
            return fixes, opps
        if level <= 1:
            if todo:
                opps.append(f"{len(todo)} month{'s' if len(todo) != 1 else ''} still qualif"
                            f"{'y' if len(todo) != 1 else 'ies'} for the rule. Work out the "
                            "leftover for every month again — including months with low demand.")
            if bad:
                opps.append(f"In {len(bad)} month{'s' if len(bad) != 1 else ''} you dropped a "
                            "worker where it doesn't pay: the leftover is too big for overtime "
                            "to be cheaper, or over the limit.")
        elif level == 2:
            if todo:
                opps.append("Still to change: " + ", ".join(scn.months[i] for i in todo) + ".")
            if bad:
                opps.append("Put the worker back in: " + ", ".join(scn.months[i] for i in bad)
                            + ".")
        else:
            for i in todo:
                q = qual[i]
                opps.append(f"**{q['month']}:** Workers {_n(q['workers'])} → "
                            f"{_n(q['workers'] - 1)}, Regular Production {_n(q['prod'])} → "
                            f"{_n((q['workers'] - 1) * cap)}, Overtime → {_n(q['rem'])}.")
            for i in bad:
                opps.append(f"**{scn.months[i]}:** Workers back to "
                            f"{_n(base['Workers'].iloc[i])}, Regular Production back to "
                            f"{_n(base['Regular Production'].iloc[i])}, Overtime → 0.")
    elif track == "level":
        extra = [i for i in range(len(plan))
                 if plan["Overtime Production"].iloc[i] + plan["Subcontract Production"].iloc[i]
                 > eps and plan["Ending Inventory"].iloc[i] > scn.target + eps]
        if extra:
            if level <= 1:
                opps.append("Some months buy more extra bottles than they need — the leftover "
                            "stock costs holding, and the extra bottles cost overtime or "
                            "subcontracting.")
            elif level == 2:
                opps.append("Buying more than needed in: "
                            + ", ".join(scn.months[i] for i in extra) + ".")
            else:
                for i in extra:
                    cut = min(plan["Overtime Production"].iloc[i]
                              + plan["Subcontract Production"].iloc[i],
                              plan["Ending Inventory"].iloc[i] - scn.target)
                    opps.append(f"**{scn.months[i]}:** cut subcontracting first, then overtime, "
                                f"by {_n(cut)} bottles in total.")
        elif cost >= costs["level"] - 0.5:
            opps.append("Your plan still costs as much as the level plan: check the months "
                        "with idle regular capacity before the peak — making bottles there "
                        "and holding them is cheaper than overtime later.")
    else:                                                          # no_gain
        base = engine.reference_plan(scn, "chase")
        diff = [i for i in range(len(plan))
                if abs(plan["Workers"].iloc[i] - base["Workers"].iloc[i]) > eps
                or plan["Overtime Production"].iloc[i] > eps
                or plan["Subcontract Production"].iloc[i] > eps]
        if cost > min(costs.values()) + 0.5 and diff:
            opps.append("Your experiment costs more than the chase plan — that's the lesson. "
                        "Put these months back to the chase plan's workers and production, "
                        "with no overtime or subcontracting: "
                        + ", ".join(scn.months[i] for i in diff) + ".")
    return fixes, opps


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
