"""Calculation engine: the answer key the uploaded workbook is checked against.

Students do every calculation in Excel. This module computes the same plans in Python so
the checker can compare. The rules are exactly the ones printed on the workbook's
"Start Here" sheet:

    Beginning Inventory   = last month's Ending Inventory   (January = BeginInv)
    Beginning Backorders  = last month's Ending Backorders  (January = BeginBackorders)
    Ending Inventory      = MAX(0, Beg Inv + Production - Beg Backorders - Demand)
    Ending Backorders     = MAX(0, Demand + Beg Backorders - Beg Inv - Production)
    Hires                 = MAX(0, Workers - last month's Workers)   (January vs StartWorkers)
    Layoffs               = MAX(0, last month's Workers - Workers)
    costs                 = quantity x rate; Total = sum of the month's costs

Chase:  Production = MAX(0, Demand + EndTarget - Beg Inv + Beg Backorders);
        Workers = ROUNDUP(Production / capacity per worker, 0)
Level:  Production = (Total demand + EndTarget - BeginInv + BeginBackorders) / 12 every month;
        Workers = ROUNDUP(level production / capacity per worker, 0) every month
Workers are paid for capacity (Workers x LaborCost) whether or not it is fully used.
Under worker_policy=partial the ROUNDUPs are dropped (fractional worker-equivalents).
"""

from __future__ import annotations

import math
import threading
from collections import OrderedDict
from dataclasses import dataclass
from typing import Dict, List, Sequence

import pandas as pd

from .scenario import Scenario

EPS = 1e-6

PLAN_COLUMNS = [
    "Month", "Demand", "Beginning Inventory", "Beginning Backorders", "Workers",
    "Regular Capacity", "Regular Production", "Overtime Production", "Subcontract Production",
    "Total Production", "Hires", "Layoffs", "Ending Inventory", "Ending Backorders",
    "On-Time Shipments", "Labor Cost", "Hiring Cost", "Layoff Cost", "Overtime Cost",
    "Subcontract Cost", "Holding Cost", "Backorder Cost", "Total Cost",
]


@dataclass(frozen=True)
class Decision:
    """The decisions that define a plan (everything else is calculated)."""
    workers: Sequence[float]
    regular: Sequence[float]
    overtime: Sequence[float] = ()
    subcontract: Sequence[float] = ()


def _z(seq, n):
    return [float(x) for x in seq] if len(seq) else [0.0] * n


def simulate(scn: Scenario, dec: Decision) -> pd.DataFrame:
    """Month-by-month plan for explicit decisions."""
    p = scn.params
    n = len(scn.months)
    D = [float(x) for x in scn.demand]
    W, R = [float(x) for x in dec.workers], [float(x) for x in dec.regular]
    O, S = _z(dec.overtime, n), _z(dec.subcontract, n)
    if not (len(W) == len(R) == len(O) == len(S) == n):
        raise ValueError(f"Every decision list must have {n} entries (one per month).")
    prev_w = float(p["starting_workforce"])
    inv, bk = float(p["beginning_inventory"]), float(p.get("beginning_backorders", 0))
    rows = []
    for i, m in enumerate(scn.months):
        w, reg, ot, sub = W[i], R[i], O[i], S[i]
        prod = reg + ot + sub
        beg_inv, beg_bk = inv, bk
        avail = inv + prod
        cleared = min(bk, avail)
        ontime = min(D[i], avail - cleared)
        inv = max(0.0, beg_inv + prod - beg_bk - D[i])
        bk = max(0.0, D[i] + beg_bk - beg_inv - prod)
        hires, layoffs = max(0.0, w - prev_w), max(0.0, prev_w - w)
        costs = {
            "Labor Cost": w * p["regular_labor_cost"],
            "Hiring Cost": hires * p["hiring_cost"],
            "Layoff Cost": layoffs * p["layoff_cost"],
            "Overtime Cost": ot * p["overtime_cost"],
            "Subcontract Cost": sub * p["subcontract_cost"],
            "Holding Cost": inv * p["holding_cost"],
            "Backorder Cost": bk * p["backorder_cost"],
        }
        rows.append({
            "Month": m, "Demand": D[i], "Beginning Inventory": beg_inv,
            "Beginning Backorders": beg_bk, "Workers": w, "Regular Capacity": w * scn.cap_month,
            "Regular Production": reg, "Overtime Production": ot,
            "Subcontract Production": sub, "Total Production": prod, "Hires": hires,
            "Layoffs": layoffs, "Ending Inventory": inv, "Ending Backorders": bk,
            "On-Time Shipments": ontime, **costs, "Total Cost": sum(costs.values()),
        })
        prev_w = w
    return pd.DataFrame(rows, columns=PLAN_COLUMNS)


def workers_for(scn: Scenario, req: float) -> float:
    """Workforce needed to make `req` bottles: whole workers round UP (the extra capacity is
    idle but paid); under the partial-worker policy the fraction is kept."""
    need = max(0.0, req) / scn.cap_month
    return float(math.ceil(need - 1e-9)) if scn.whole_workers else need


def reference_chase(scn: Scenario) -> Decision:
    """Chase: each month produce exactly what is needed to end at the target."""
    t = scn.target
    net = float(scn.params["beginning_inventory"]) - float(scn.params.get("beginning_backorders", 0))
    workers, prod = [], []
    for d in scn.demand:
        req = max(0.0, d + t - net)
        workers.append(workers_for(scn, req))
        prod.append(req)
        net = net + req - d
    return Decision(workers, prod)


def level_rate(scn: Scenario) -> float:
    """(total demand + ending-inventory target - beginning inventory + beginning backorders)
    / months, never < 0."""
    p = scn.params
    return max(0.0, (scn.total_demand + scn.target - p["beginning_inventory"]
                     + p.get("beginning_backorders", 0)) / len(scn.months))


def reference_level(scn: Scenario) -> Decision:
    rate = level_rate(scn)
    n = len(scn.months)
    return Decision([workers_for(scn, rate)] * n, [rate] * n)


_REF_CACHE: "OrderedDict[tuple, pd.DataFrame]" = OrderedDict()
_REF_LOCK = threading.Lock()


def reference_plan(scn: Scenario, kind: str) -> pd.DataFrame:
    """The answer-key plan. Every check, guide question and verdict asks for it, so it is
    computed once per scenario and kept (bounded); callers get their own copy."""
    key = (scn.config_hash, scn.seed, kind)
    with _REF_LOCK:
        if key in _REF_CACHE:
            _REF_CACHE.move_to_end(key)
            return _REF_CACHE[key].copy()
    df = simulate(scn, reference_chase(scn) if kind == "chase" else reference_level(scn))
    with _REF_LOCK:
        _REF_CACHE[key] = df
        while len(_REF_CACHE) > 512:
            _REF_CACHE.popitem(last=False)
    return df.copy()


# --------------------------------------------------------------------------- #
# Summary metrics (the rows of the workbook's Summary sheet)
# --------------------------------------------------------------------------- #
def summarize(df: pd.DataFrame) -> Dict[str, float]:
    return {
        "Labor Cost": float(df["Labor Cost"].sum()),
        "Hiring Cost": float(df["Hiring Cost"].sum()),
        "Layoff Cost": float(df["Layoff Cost"].sum()),
        "Overtime Cost": float(df["Overtime Cost"].sum()),
        "Subcontract Cost": float(df["Subcontract Cost"].sum()),
        "Holding Cost": float(df["Holding Cost"].sum()),
        "Backorder Cost": float(df["Backorder Cost"].sum()),
        "Total Cost": float(df["Total Cost"].sum()),
        "Total Hires": float(df["Hires"].sum()),
        "Total Layoffs": float(df["Layoffs"].sum()),
        "Peak Workforce": float(df["Workers"].max()),
        "Peak Inventory": float(df["Ending Inventory"].max()),
        "Months With Backorders": float((df["Ending Backorders"] > EPS).sum()),
        "On-Time Fulfillment %": (100.0 * float(df["On-Time Shipments"].sum())
                                  / max(EPS, float(df["Demand"].sum()))),
    }


# --------------------------------------------------------------------------- #
# Feasibility (used for the optional hybrid plan)
# --------------------------------------------------------------------------- #
def check_feasibility(df: pd.DataFrame, scn: Scenario) -> List[dict]:
    """Each check: {label, ok, detail}. A plan is feasible when every check is ok."""
    p = scn.params
    months = list(df["Month"])
    checks: List[dict] = []

    def add(label, mask, ok_text, fail_prefix):
        bad = [months[i] for i, v in enumerate(mask) if v]
        checks.append({"label": label, "ok": not bad,
                       "detail": ok_text if not bad else f"{fail_prefix}: {', '.join(bad)}"})

    add("Regular production within capacity",
        (df["Regular Production"] > df["Regular Capacity"] + EPS).tolist(),
        "Regular production ≤ Workers × capacity per worker every month.",
        "Regular production exceeds capacity in")
    if scn.whole_workers:
        add("Whole-number workforce",
            [abs(w - round(w)) > 1e-6 for w in df["Workers"]],
            "Workers are whole numbers every month.", "Fractional workers in")
    ot = df["Overtime Production"]
    add(f"Overtime within {p['overtime_pct']:.0%} of regular capacity",
        (ot > df["Regular Capacity"] * p["overtime_pct"] + EPS).tolist(),
        "Overtime is within the limit every month.", "Overtime exceeds the limit in")
    add("Overtime only when regular capacity is fully used",
        ((ot > EPS) & (df["Regular Capacity"] - df["Regular Production"] > EPS)).tolist(),
        "No overtime while regular capacity sits idle.",
        "Overtime used while regular capacity is idle in")
    add(f"Subcontracting within {p['subcontract_capacity']:,.0f} bottles/month",
        (df["Subcontract Production"] > p["subcontract_capacity"] + EPS).tolist(),
        "Subcontracting is within its limit.", "Subcontract limit exceeded in")
    add(f"Storage limit ({p['max_inventory']:,.0f} bottles)",
        (df["Ending Inventory"] > p["max_inventory"] + EPS).tolist(),
        "Ending inventory never exceeds storage.", "Storage limit exceeded in")
    if p["service_requirement_pct"] > 0:
        ont = summarize(df)["On-Time Fulfillment %"]
        checks.append({"label": f"On-time fulfillment ≥ {p['service_requirement_pct']:g}%",
                       "ok": bool(ont + 1e-9 >= p["service_requirement_pct"]),
                       "detail": f"On-time fulfillment is {ont:.1f}%."})
    last = df.iloc[-1]
    checks.append({"label": f"December backorders ≤ {p['terminal_backlog_max']:,.0f}",
                   "ok": bool(last["Ending Backorders"] <= p["terminal_backlog_max"] + EPS),
                   "detail": f"December ending backorders: {last['Ending Backorders']:,.0f}."})
    checks.append({"label": f"December inventory ≥ {scn.target:,.0f}",
                   "ok": bool(last["Ending Inventory"] >= scn.target - EPS),
                   "detail": f"December ending inventory: {last['Ending Inventory']:,.0f}."})
    return checks
