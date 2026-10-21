"""Help with DESIGNING the columns: which ones a plan needs, why, what each is calculated from
and what uses it. One source for the app's column planner, the "how they connect" diagram,
the heading check and the printable planner on the workbook's Start Here sheet.

It explains the business logic, never a formula: "comes from" names the ingredients (the
level of hint 1-2), and the formula itself stays with the hint ladder (hints.py)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from . import layout as L
from .recognize import HELPER, REQ_PRODUCTION, classify
from .hints import QUESTIONS

DEMAND, PLANNING = "Demand", "Planning cells"

STAGES: List[Tuple[str, str]] = [
    ("Start of the month", "What you begin the month with."),
    ("Your plan for the month", "How many bottles you make, and with how many people."),
    ("What happens", "The results of the plan: the crew changes, and what's left over or owed."),
    ("What it costs", "One cost column for each thing that costs money, then the month's total."),
]
_STAGE_OF = {
    "Beginning Inventory": 0, "Beginning Backorders": 0,
    "Regular Production": 1, "Workers": 1, "Regular Capacity": 1, "Overtime Production": 1,
    "Subcontract Production": 1, "Total Production": 1,
    "Hires": 2, "Layoffs": 2, "Ending Inventory": 2, "Ending Backorders": 2,
}
_ORDER = ["Beginning Inventory", "Beginning Backorders", "Regular Production", "Workers",
          "Regular Capacity", "Overtime Production", "Subcontract Production", "Total Production",
          "Hires", "Layoffs", "Ending Inventory", "Ending Backorders", "Labor Cost",
          "Hiring Cost", "Layoff Cost", "Overtime Cost", "Subcontract Cost", "Holding Cost",
          "Backorder Cost", "Total Cost"]

_WHY = {
    "Beginning Inventory": "Bottles already on the shelf cover part of this month's demand, so "
                           "you need to make fewer.",
    "Beginning Backorders": "Orders you still owe must be filled on top of this month's demand "
                            "— and every month they wait costs money.",
    "Regular Production": "The heart of the plan: how many bottles your crew makes on regular "
                          "time.",
    "Workers": "Bottles need people. The crew size sets the labor cost, and every change in it "
               "means hiring or layoffs.",
    "Regular Capacity": "What your crew could make on regular time. It shows idle capacity, "
                        "and overtime may only start once it's full. (Optional column.)",
    "Overtime Production": "Extra bottles from your own crew beyond regular time — up to the "
                           "overtime limit, and only once regular capacity is full.",
    "Subcontract Production": "Bottles bought from a partner, up to its monthly capacity.",
    "Total Production": "Everything that comes in this month, from all three sources.",
    "Hires": "Adding workers costs money (a hiring cost per worker).",
    "Layoffs": "Letting workers go costs money too (a layoff cost per worker).",
    "Ending Inventory": "Bottles left over carry into next month — and cost money to store.",
    "Ending Backorders": "Orders you couldn't fill carry into next month — and each one "
                         "waiting costs money.",
    "Labor Cost": "What your crew is paid this month.",
    "Hiring Cost": "What adding workers cost this month.",
    "Layoff Cost": "What letting workers go cost this month.",
    "Overtime Cost": "What the overtime bottles cost (the full cost of each one).",
    "Subcontract Cost": "What the subcontracted bottles cost.",
    "Holding Cost": "What storing the leftover bottles cost.",
    "Backorder Cost": "What the unfilled orders cost.",
    "Total Cost": "One number per month to compare plans; its yearly total is the plan's cost.",
}
_RATE = {"Labor Cost": "LaborCost", "Hiring Cost": "HireCost", "Layoff Cost": "LayoffCost",
         "Overtime Cost": "OTCost", "Subcontract Cost": "SubCost", "Holding Cost": "HoldCost",
         "Backorder Cost": "BackorderCost"}
_QTY = {"Labor Cost": "Workers", "Hiring Cost": "Hires", "Layoff Cost": "Layoffs",
        "Overtime Cost": "Overtime Production", "Subcontract Cost": "Subcontract Production",
        "Holding Cost": "Ending Inventory", "Backorder Cost": "Ending Backorders"}


def concepts(kind: str) -> List[str]:
    """The columns a plan needs (plus the optional Regular Capacity), in working order."""
    have = set(L.REQUIRED[kind]) | set(L.OPTIONAL[kind])
    return [c for c in _ORDER if c in have]


def is_optional(kind: str, concept: str) -> bool:
    return concept in L.OPTIONAL[kind] and concept not in L.REQUIRED[kind]


def is_decision(kind: str, concept: str) -> bool:
    return kind == "hybrid" and concept in L.DECISIONS


def inputs(kind: str, concept: str) -> List[str]:
    """What this column is calculated from in the SAME month (Demand / Planning cells / other
    columns). Last month's values are in `from_last_month`."""
    prod = "Total Production" if kind == "hybrid" else "Regular Production"
    if is_decision(kind, concept):
        return []
    if concept == "Regular Production":
        return [DEMAND, "Beginning Inventory", "Beginning Backorders"] if kind == "chase" else [PLANNING]
    if concept == "Workers":
        return ["Regular Production"] if kind == "chase" else [PLANNING]
    if concept == "Regular Capacity":
        return ["Workers"]
    if concept == "Total Production":
        return ["Regular Production", "Overtime Production", "Subcontract Production"]
    if concept in ("Hires", "Layoffs"):
        return ["Workers"]
    if concept in ("Ending Inventory", "Ending Backorders"):
        return ["Beginning Inventory", "Beginning Backorders", prod, DEMAND]
    if concept in _QTY:
        return [_QTY[concept]]
    if concept == "Total Cost":
        return [c for c in concepts(kind) if c in _QTY]
    return []


def from_last_month(concept: str) -> Optional[str]:
    return {"Beginning Inventory": "Ending Inventory", "Beginning Backorders": "Ending Backorders",
            "Hires": "Workers", "Layoffs": "Workers"}.get(concept)


def used_by(kind: str, concept: str) -> List[str]:
    """The columns that use this one (by the plan's own heading idea)."""
    out = [heading_idea(kind, c) for c in concepts(kind) if concept in inputs(kind, c)]
    nxt = [heading_idea(kind, c) for c in concepts(kind) if from_last_month(c) == concept]
    return out + [f"next month's {c}" for c in nxt]


def comes_from(kind: str, concept: str, whole_workers: bool = True) -> str:
    if is_decision(kind, concept):
        return "Your decision — type the number you choose."
    if concept == "Beginning Inventory":
        return "Last month's Ending Inventory (January: the given starting inventory, BeginInv)."
    if concept == "Beginning Backorders":
        return ("Last month's Ending Backorders (January: the given starting backorders, "
                "BeginBackorders).")
    if concept == "Regular Production":
        if kind == "level":
            return "Your level production rate (a planning cell) — the same every month."
        return ("This month's Demand, your ending-inventory target (a planning cell), Beginning "
                "Inventory and Beginning Backorders.")
    if concept == "Workers":
        if kind == "level":
            return "Your level workforce (a planning cell) — the same every month."
        return ("Production and what one worker makes in a month (a planning cell)"
                + (" — rounded UP to whole workers." if whole_workers else "."))
    if concept == "Regular Capacity":
        return "Workers and what one worker makes in a month (a planning cell)."
    if concept == "Total Production":
        return "Regular, Overtime and Subcontract Production."
    if concept in ("Hires", "Layoffs"):
        return ("This month's Workers and last month's (January: the starting workforce, "
                "StartWorkers).")
    if concept in ("Ending Inventory", "Ending Backorders"):
        prod = "Total Production" if kind == "hybrid" else "Production"
        return f"Beginning Inventory, Beginning Backorders, {prod} and this month's Demand."
    if concept in _QTY:
        return f"{_QTY[concept]} and its cost rate ({_RATE[concept]} on the Scenario sheet)."
    if concept == "Total Cost":
        return "Every cost column to its left."
    return ""


def heading_idea(kind: str, concept: str) -> str:
    return L.heading_for(kind, concept)


@dataclass
class ColumnInfo:
    concept: str
    heading: str
    question: str
    why: str
    comes_from: str
    used_by: List[str]
    optional: bool
    decision: bool


def stages(kind: str, whole_workers: bool = True) -> List[Tuple[str, str, List[ColumnInfo]]]:
    """[(stage title, stage blurb, [ColumnInfo])] for the column planner."""
    out = [(t, b, []) for t, b in STAGES]
    for c in concepts(kind):
        info = ColumnInfo(c, heading_idea(kind, c), QUESTIONS[c], _WHY[c],
                          comes_from(kind, c, whole_workers), used_by(kind, c),
                          is_optional(kind, c), is_decision(kind, c))
        out[_STAGE_OF.get(c, 3)][2].append(info)
    return out


def whats_new(kind: str) -> str:
    """How this plan's columns differ from the plan before (markdown, without giving the new
    columns away)."""
    n = len(L.REQUIRED[kind])
    if kind == "chase":
        return (f"**Your first plan: {n} questions, one column each.** Build them once, "
                "carefully — the level plan needs the same columns.")
    if kind == "level":
        return (f"**The same {n} columns as your chase plan** — copy your chase headings. What "
                "changes: **Production** and **Workers** now come from your level planning "
                "cells instead of following demand.")
    new = [c for c in L.REQUIRED["hybrid"] if c not in L.REQUIRED["chase"]]
    return (f"**Your chase columns plus {len(new)} new questions** (marked *new* below). Four "
            "columns become decisions you type instead of calculating.")


_END_WORDS = ("end", "ending", "closing", "close", "left", "remaining", "final", "eom")
_SUPPORT = [
    "That heading isn't recognised yet. Say in plain words what the column measures this "
    "month.",
    "Pick one of the open questions and turn its key words into a heading: name the thing "
    "it counts or costs and, where it matters, *when* (the start or the end of the month).",
    "Headings the checker understands usually name a thing — stock or inventory, orders "
    "owed or backorders, workers or crew, hires, layoffs, bottles made or production, or a "
    "… cost — plus a time word such as beginning/starting or ending/closing.",
]


def support(misses: int) -> str:
    """Help after `misses` unrecognised headings (1, 2, 3+)."""
    return _SUPPORT[min(max(misses, 1), len(_SUPPORT)) - 1]


def check_one(kind: str, heading: str, found: Dict[str, str]) -> dict:
    """Check ONE planned heading. `found`: {concept: the student's heading} found so far.
    Returns {"status": ok | repeat | other_plan | helper | unknown | empty, "concept",
    "name", "question", "why", "note"} — the column's name and why it matters only on ok."""
    h = (heading or "").strip()
    if not h:
        return {"status": "empty"}
    c = classify(h, kind)
    if c == REQ_PRODUCTION or (kind != "hybrid" and c == "Total Production"):
        c = "Regular Production"
    if c == HELPER:
        return {"status": "helper"}
    if c is None:
        return {"status": "unknown"}
    if c not in concepts(kind):
        plans = [k for k in ("chase", "level", "hybrid") if c in concepts(k)]
        return {"status": "other_plan", "plans": plans}
    out = {"concept": c, "name": heading_idea(kind, c), "question": QUESTIONS[c]}
    if c in found:
        return {**out, "status": "repeat", "yours": found[c]}
    note = ""
    words = h.lower()
    if c in ("Ending Inventory", "Ending Backorders") and not any(w in words for w in _END_WORDS):
        note = ("This reads as the **end** of the month. If you meant the start, say so "
                "(e.g. “Beginning …”).")
    return {**out, "status": "ok", "why": _WHY[c], "note": note,
            "optional": is_optional(kind, c), "decision": is_decision(kind, c)}


def open_questions(kind: str, found: Dict[str, str]) -> List[Tuple[str, List[Tuple[str, str]]]]:
    """[(stage title, [(question, tag)])] still to find, in month order; tag is "new" (new in
    the hybrid), "optional" or ""."""
    new = set(concepts(kind)) - set(concepts("chase")) if kind == "hybrid" else set()
    out = []
    for title, _, cols in stages(kind):
        qs = [(c.question, "new" if c.concept in new else "optional" if c.optional else "")
              for c in cols if c.concept not in found]
        if qs:
            out.append((title, qs))
    return out


# ---- diagram ----------------------------------------------------------------------------------
_COLORS = {"given": "#FFF2CC", "decision": "#E2EFDA", "calc": "#DDEBF7", "cost": "#FCE4D6",
           "planning": "#EDEDED"}


def _q(s: str) -> str:
    return '"' + s.replace('"', "'") + '"'


def diagram(kind: str, found: Optional[Dict[str, str]] = None) -> str:
    """Graphviz DOT: how the columns feed each other within a month. What carries over from
    last month is written in the box ("+ last month's Ending Inventory") rather than drawn,
    which would loop arrows back across the whole picture. With `found`, a column not found
    yet is drawn as "?" — the picture fills in as the student finds each one."""
    cs = concepts(kind)
    lines = ["digraph G {", 'rankdir=TB; nodesep=0.2; ranksep=0.35;',   # tall: fits a column
             'node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=13, '
             'color="#8C8C8C"];', 'edge [color="#7F7F7F", arrowsize=0.7];',
             f'{_q(DEMAND)} [fillcolor="{_COLORS["given"]}", label="Demand\\n(given)"];']
    if any(PLANNING in inputs(kind, c) for c in cs):
        lines.append(f'{_q(PLANNING)} [fillcolor="{_COLORS["planning"]}", '
                     'label="Planning cells\\n(under the grid)"];')
    for k, (title, _) in enumerate(STAGES):
        members = [c for c in cs if _STAGE_OF.get(c, 3) == k]
        lines.append(f'subgraph cluster_{k} {{ label={_q(f"{k + 1}. {title}")}; fontname='
                     '"Helvetica-Bold"; fontsize=12; color="#BFBFBF"; style="rounded";')
        for c in members:
            fill = (_COLORS["decision"] if is_decision(kind, c) else
                    _COLORS["cost"] if c in L.COST_CONCEPTS else _COLORS["calc"])
            extra = ', style="rounded,filled,dashed"' if is_optional(kind, c) else ""
            prev = from_last_month(c)
            if found is not None and c not in found:
                lines.append(f'{_q(c)} [fillcolor="#FFFFFF", label="?", style="rounded,dashed"];')
                continue
            label = heading_idea(kind, c) + (
                "\\n(your decision)" if is_decision(kind, c) else
                "\\n(optional)" if is_optional(kind, c) else
                f"\\n(+ last month's {heading_idea(kind, prev)})" if prev else "")
            lines.append(f'{_q(c)} [fillcolor="{fill}", label={_q(label)}{extra}];')
        lines.append("}")
    for c in cs:
        for src in inputs(kind, c):
            if src == DEMAND or src == PLANNING or src in cs:
                lines.append(f"{_q(src)} -> {_q(c)};")
    lines.append("}")
    return "\n".join(lines)
