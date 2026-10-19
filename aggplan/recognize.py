"""Work out what each student-built column is, from its heading.

Students design their own worksheet, so "Starting Stock", "Beg. Inv" and "Opening
inventory" must all be read as Beginning Inventory, and "Crew", "Fired" and "Wages" as
Workers, Layoffs and Labor Cost. classify() maps one heading to a concept (an
engine.simulate() column name), a helper marker, or None. map_columns() maps a whole sheet.

When a required concept has no recognisable heading, map_columns() can fall back on the
values: an unrecognised column whose numbers match that concept in at least 10 of 12
months is taken as that column, and the student is told to give it a clearer heading."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from openpyxl.utils import get_column_letter

from . import layout as L

HELPER = "_helper"            # recognised as a helper column (fine, not graded)
REQ_PRODUCTION = "_required_production"

_BEGIN = ("begin", "beg", "start", "starting", "opening", "open", "initial", "carried",
          "previous", "prior", "bom", "incoming")
_BACKORDER = ("backorder", "back order", "backlog", "back log", "shortage", "short",
              "unfilled", "stockout", "stock out", "unmet", "owed", "late order", "waiting")
_INVENTORY = ("inventory", "inv", "stock", "on hand", "onhand")
_WORKERS = ("worker", "workforce", "employee", "staff", "headcount", "head count", "crew",
            "people", "labor force", "labour force", "team")
_PRODUCTION = ("production", "produce", "produced", "output", "made", "make", "bottles",
               "units", "manufactur")
_HIRE = ("hire", "hiring", "hired", "new worker", "added worker", "workers added")
_LAYOFF = ("layoff", "lay off", "laid off", "lay offs", "fire", "fired", "firing", "let go",
           "separation", "dismiss", "workers cut", "cut")
_OVERTIME = ("overtime", "over time")
_SUBCONTRACT = ("subcontract", "sub contract", "outsourc", "contracted", "sub con")
_HOLD = ("hold", "carry", "carrying", "inventory", "storage", "stock", "inv")
_LABOR = ("labor", "labour", "wage", "payroll", "salary", "salaries", "worker", "regular",
          "employee", "staff", "crew", "pay")


def _norm(text: str) -> str:
    t = re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()
    return f" {t} "


def _has(h: str, words) -> bool:
    for w in words:
        if len(w) <= 3:                       # short words must be whole tokens (inv, ot)
            if f" {w} " in h:
                return True
        elif w in h:
            return True
    return False


def classify(heading, kind: str) -> Optional[str]:
    """Concept id, HELPER, REQ_PRODUCTION or None (unrecognised)."""
    if heading is None:
        return None
    raw = str(heading)
    h = _norm(raw)
    if not h.strip():
        return None
    tokens = h.split()
    costy = "$" in raw or _has(h, ("cost", "expense", "wage", "payroll", "salary",
                                   "salaries", "penalty", "charge", "spend",
                                   "severance"))
    if _has(h, ("idle", "unused", "cumulative", "demand", "forecast", "month", "shipped",
                "shipment", "sales", "net position")) and not costy:
        return HELPER
    if costy:
        if _has(h, _HIRE):
            return "Hiring Cost"
        if _has(h, _LAYOFF) or _has(h, ("severance",)):
            return "Layoff Cost"
        if _has(h, _OVERTIME) or "ot" in tokens:
            return "Overtime Cost"
        if _has(h, _SUBCONTRACT):
            return "Subcontract Cost"
        if _has(h, _BACKORDER):
            return "Backorder Cost"
        if _has(h, _HOLD):
            return "Holding Cost"
        if _has(h, _LABOR):
            return "Labor Cost"
        if _has(h, ("total", "monthly", "all", "month")) or tokens == ["cost"] or tokens == ["costs"]:
            return "Total Cost"
        return None
    if _has(h, ("capacity",)):
        return "Regular Capacity"
    if _has(h, _OVERTIME) or "ot" in tokens:
        return "Overtime Production"
    if _has(h, _SUBCONTRACT):
        return "Subcontract Production"
    if _has(h, _HIRE):
        return "Hires"
    if _has(h, _LAYOFF):
        return "Layoffs"
    begin = any(t in _BEGIN or t.startswith("begin") for t in tokens) or _has(h, ("carried in", "from last"))
    if _has(h, _BACKORDER):
        return "Beginning Backorders" if begin else "Ending Backorders"
    if _has(h, _INVENTORY):
        return "Beginning Inventory" if begin else "Ending Inventory"
    if _has(h, ("required", "requirement", "needed", "need", "net req")) and not _has(h, _WORKERS):
        return REQ_PRODUCTION
    if _has(h, _WORKERS) and not _has(h, _PRODUCTION):
        return "Workers"
    if _has(h, _PRODUCTION):
        if _has(h, ("total",)):
            return "Total Production" if kind == "hybrid" else "Regular Production"
        return "Regular Production"
    return None


@dataclass
class ColumnMap:
    by_concept: Dict[str, str] = field(default_factory=dict)       # concept -> letter
    headings: Dict[str, str] = field(default_factory=dict)         # letter -> heading text
    read_as: Dict[str, Optional[str]] = field(default_factory=dict)  # letter -> concept/None
    duplicates: Dict[str, List[str]] = field(default_factory=dict)  # concept -> letters
    guessed: Dict[str, str] = field(default_factory=dict)          # concept -> letter
    unknown: List[str] = field(default_factory=list)               # unrecognised letters

    def missing(self, kind: str) -> List[str]:
        return [c for c in L.REQUIRED[kind] if c not in self.by_concept]

    @property
    def n_headings(self) -> int:
        return len(self.headings)


def map_columns(headings: Dict[int, str], kind: str,
                guess: Optional[Callable[[str, str], bool]] = None) -> ColumnMap:
    """`headings`: {column index: heading text} for the student's columns.
    `guess(concept, letter)` -> True if that column's values match the concept (fallback)."""
    cm = ColumnMap()
    req_prod: List[str] = []
    for j in sorted(headings):
        text = str(headings[j]).strip()
        if not text:
            continue
        letter = get_column_letter(j)
        cm.headings[letter] = text
        concept = classify(text, kind)
        if kind != "hybrid" and concept in ("Total Production",):
            concept = "Regular Production"
        if concept == REQ_PRODUCTION:
            req_prod.append(letter)
            cm.read_as[letter] = None
            continue
        if concept is None:
            cm.unknown.append(letter)
            cm.read_as[letter] = None
            continue
        if concept == HELPER:
            cm.read_as[letter] = HELPER
            continue
        if concept in cm.by_concept:
            cm.duplicates.setdefault(concept, [cm.by_concept[concept]]).append(letter)
            cm.read_as[letter] = None
            continue
        cm.by_concept[concept] = letter
        cm.read_as[letter] = concept
    if req_prod:                                      # "Required Production" as production
        if "Regular Production" not in cm.by_concept:
            cm.by_concept["Regular Production"] = req_prod[0]
            cm.read_as[req_prod[0]] = "Regular Production"
        for letter in req_prod:
            if cm.read_as.get(letter) is None:
                cm.read_as[letter] = HELPER
    if guess:
        for concept in cm.missing(kind):
            if kind == "hybrid" and concept in L.DECISIONS:
                continue                              # decisions are never guessed
            for letter in list(cm.unknown):
                if guess(concept, letter):
                    cm.by_concept[concept] = letter
                    cm.guessed[concept] = letter
                    cm.read_as[letter] = concept
                    cm.unknown.remove(letter)
                    break
    return cm
