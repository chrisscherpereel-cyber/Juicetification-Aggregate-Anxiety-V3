"""Check an uploaded workbook against the answer key.

The checker never evaluates a formula itself. It reads the file twice with openpyxl:
once for the formulas the student typed and once for the values Excel calculated and saved
with the file. Each checked cell gets one status:

  ok          right value, calculated by a formula
  typed       right value, but typed in as a number (only when formulas are required)
  carried     the cell follows the rule, but an earlier cell it uses is wrong
  wrong       the value is wrong
  blank       nothing there yet
  error       Excel shows an error (#DIV/0!, #NAME? ...)
  text        words where a number belongs
  uncalc      a formula with no saved value (the file was not saved by a spreadsheet app)

"carried" is the useful one for learning: it separates "your formula is wrong" from
"your formula is right but is fed a wrong number", so the student fixes the root cause.
"""

from __future__ import annotations

import hashlib
import io
import math
import re
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd

from . import MODEL_VERSION, engine, layout as L, text
from .engine import Decision
from .scenario import Scenario
from .workbook import APP_ID

MAX_BYTES = 5 * 1024 * 1024
MAX_UNCOMPRESSED = 60 * 1024 * 1024
MAX_MEMBERS = 2000
READ_ROWS, READ_COLS = 40, 30

OK_STATUSES = {"ok"}
STATUS_LABEL = {
    "ok": "✔ Correct", "typed": "✎ Typed, not calculated", "carried": "↪ Follows an earlier mistake",
    "wrong": "✖ Not right yet", "blank": "○ Empty", "error": "⚠ Excel error",
    "text": "⚠ Text, not a number", "uncalc": "⚠ Not calculated", "decision": "◆ Your decision",
}
STATUS_HELP = {
    "typed": "The number is right but it was typed in. Replace it with a formula so Excel "
             "calculates it.",
    "carried": "Your formula follows the rule, but it uses an earlier cell that is wrong. Fix "
               "that cell first and this one will correct itself.",
    "blank": "This cell still needs a formula.",
    "error": "Excel could not calculate this cell. See 'What an error means' on the Excel "
             "Help sheet.",
    "text": "This cell holds text. Delete it and enter a formula that starts with =.",
    "uncalc": "The formula has no saved result. Open the file in Excel, let it calculate, "
              "save as .xlsx and upload again.",
}
_ERR_RE = re.compile(r"^#(DIV/0!|NAME\?|REF!|VALUE!|N/A|NUM!|NULL!|SPILL!|CALC!)")


# =========================================================================================
# Reading
# =========================================================================================
@dataclass
class Cell:
    value: object = None          # value saved by Excel (None if blank)
    formula: Optional[str] = None  # text of the formula, if the cell holds one

    @property
    def blank(self) -> bool:
        return self.formula is None and (self.value is None or
                                         (isinstance(self.value, str) and not self.value.strip()))


@dataclass
class Book:
    sheets: Dict[str, Dict[str, Cell]]
    meta: Dict[str, str]
    sha256: str

    def cell(self, sheet: str, ref: str) -> Cell:
        return self.sheets.get(sheet, {}).get(ref, Cell())


def _container_problem(data: bytes) -> Optional[str]:
    if not data:
        return "The file is empty."
    if len(data) > MAX_BYTES:
        return "The file is larger than 5 MB. Upload the workbook you downloaded from the app."
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            infos = z.infolist()
            if len(infos) > MAX_MEMBERS or sum(i.file_size for i in infos) > MAX_UNCOMPRESSED:
                return "The file is too large inside to be this workbook."
            names = {i.filename for i in infos}
            if "xl/workbook.xml" not in names:
                return "This is not an Excel .xlsx workbook."
    except zipfile.BadZipFile:
        return ("This is not an Excel .xlsx workbook. If you used another program, save or "
                "download it as an Excel Workbook (.xlsx).")
    return None


def read_book(data: bytes) -> Tuple[Optional[Book], Optional[str]]:
    """(Book, None) or (None, a message the student can act on). Never raises."""
    problem = _container_problem(data)
    if problem:
        return None, problem
    try:
        from openpyxl import load_workbook
        wf = load_workbook(io.BytesIO(data), data_only=False)
        wv = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:                                   # noqa: BLE001
        return None, f"The workbook could not be opened ({type(e).__name__}). Save it again as .xlsx."
    sheets: Dict[str, Dict[str, Cell]] = {}
    for name in wf.sheetnames:
        f_ws, v_ws = wf[name], wv[name]
        cells: Dict[str, Cell] = {}
        for row_f, row_v in zip(f_ws.iter_rows(min_row=1, max_row=READ_ROWS, max_col=READ_COLS),
                                v_ws.iter_rows(min_row=1, max_row=READ_ROWS, max_col=READ_COLS)):
            for cf, cv in zip(row_f, row_v):
                if not hasattr(cf, "coordinate"):
                    continue                                   # merged-cell placeholder
                raw = cf.value
                formula = None
                if isinstance(raw, str) and raw.startswith("=") and cf.data_type == "f":
                    formula = raw
                elif type(raw).__name__ in ("ArrayFormula", "DataTableFormula"):
                    formula = getattr(raw, "text", None) or "="
                val = cv.value
                if formula is None and val is None and raw is not None:
                    val = raw
                if formula is not None or val is not None:
                    cells[cf.coordinate] = Cell(val, formula)
        sheets[name] = cells
    meta: Dict[str, str] = {}
    if L.SHEET_META in wv.sheetnames:
        ws = wv[L.SHEET_META]
        for r in range(1, 30):
            k, v = ws.cell(r, 1).value, ws.cell(r, 2).value
            if k:
                meta[str(k)] = "" if v is None else str(v)
    return Book(sheets, meta, hashlib.sha256(data).hexdigest()), None


# =========================================================================================
# Results
# =========================================================================================
@dataclass
class CellCheck:
    sheet: str
    ref: str
    label: str                 # e.g. "Ending Inventory, March"
    group: str                 # column heading / planning label (for grids)
    row_label: str             # month / "Total" / "Planning"
    status: str
    message: str = ""
    shown: str = ""            # what the student's cell shows

    @property
    def needs_work(self) -> bool:
        return self.status not in ("ok", "decision")


@dataclass
class SheetResult:
    key: str
    title: str
    cells: List[CellCheck] = field(default_factory=list)
    required: bool = True
    attempted: bool = True
    notes: List[str] = field(default_factory=list)

    @property
    def graded(self) -> List[CellCheck]:
        return [c for c in self.cells if c.status != "decision"]

    @property
    def n_ok(self) -> int:
        return sum(c.status == "ok" for c in self.graded)

    @property
    def n_total(self) -> int:
        return len(self.graded)

    @property
    def complete(self) -> bool:
        return self.attempted and all(c.status == "ok" for c in self.graded)

    def counts(self) -> Dict[str, int]:
        out: Dict[str, int] = {}
        for c in self.graded:
            out[c.status] = out.get(c.status, 0) + 1
        return out


@dataclass
class CheckResult:
    fatal: List[str] = field(default_factory=list)
    sheets: Dict[str, SheetResult] = field(default_factory=dict)
    meta: Dict[str, str] = field(default_factory=dict)
    sha256: str = ""
    recommendation: str = ""
    justification: str = ""
    words: int = 0
    min_words: int = 0
    plan_totals: Dict[str, Dict[str, float]] = field(default_factory=dict)
    hybrid_feasibility: List[dict] = field(default_factory=list)
    uncalculated: bool = False

    @property
    def readable(self) -> bool:
        return not self.fatal

    @property
    def complete(self) -> bool:
        return self.readable and all(s.complete for s in self.sheets.values() if s.required)

    def missing(self) -> List[str]:
        if self.fatal:
            return list(self.fatal)
        out = []
        for s in self.sheets.values():
            if s.required and not s.complete:
                left = s.n_total - s.n_ok
                out.append(f"{s.title}: {left} cell{'s' if left != 1 else ''} still to fix")
        return out

    def fix_first(self, limit: int = 8) -> Tuple[List[Tuple[CellCheck, int]], int]:
        """What to fix next: ([(first cell, how many similar cells)], number of cells that
        only carry an earlier mistake). Problems are grouped by sheet, column and status, so
        twelve months with the same $-sign slip show up once; carried cells are left out
        because they correct themselves once their root cause is fixed."""
        groups: Dict[Tuple[str, str, str], List[CellCheck]] = {}
        carried = 0
        for s in self.sheets.values():
            if not s.attempted:
                continue
            for c in s.cells:
                if c.status == "carried":
                    carried += 1
                elif c.needs_work:
                    groups.setdefault((c.sheet, c.group, c.status), []).append(c)
        return [(cs[0], len(cs)) for cs in groups.values()][:limit], carried


# =========================================================================================
# Helpers
# =========================================================================================
def _num(cell: Cell) -> Optional[float]:
    v = cell.value
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else f
    return None


def _shown(cell: Cell) -> str:
    v = cell.value
    if v is None:
        return "" if cell.formula is None else "(no value)"
    if isinstance(v, float):
        return f"{v:,.2f}".rstrip("0").rstrip(".") if not v.is_integer() else f"{v:,.0f}"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)[:40]


def _tol(engine_name: str) -> float:
    return 0.01 if engine_name in ("Workers", "level_workers", "Peak Workforce") else 0.5


def _close(a: Optional[float], b: Optional[float], tol: float) -> bool:
    return a is not None and b is not None and abs(a - b) <= max(tol, 1e-9 * abs(b))


_LOCK_RE = re.compile(r"(?<![$A-Z])([SY])(\d+)\b")


def _judge(cell: Cell, key: Optional[float], local: Optional[float], tol: float,
           require_formula: bool, rule: str, extra: Optional[Callable[[float], str]] = None
           ) -> Tuple[str, str]:
    """(status, message) for one numeric cell."""
    if cell.blank:
        return "blank", STATUS_HELP["blank"] + (f" Rule: {rule}" if rule else "")
    if cell.formula is not None and cell.value is None:
        return "uncalc", STATUS_HELP["uncalc"]
    if isinstance(cell.value, str):
        if _ERR_RE.match(cell.value.strip()):
            return "error", f"Excel shows {cell.value.strip()}. " + STATUS_HELP["error"]
        return "text", STATUS_HELP["text"]
    v = _num(cell)
    if v is None:
        return "text", STATUS_HELP["text"]
    if _close(v, key, tol):
        if require_formula and cell.formula is None:
            return "typed", STATUS_HELP["typed"]
        return "ok", ""
    if _close(v, local, tol):
        return "carried", STATUS_HELP["carried"]
    msg = f"Rule: {rule}" if rule else "Check this value."
    if extra:
        hint = extra(v)
        if hint:
            msg = hint + " " + msg
    return "wrong", msg


# =========================================================================================
# Plan sheets
# =========================================================================================
def _plan_values(book: Book, kind: str) -> Dict[str, List[Optional[float]]]:
    sheet = L.PLAN_SHEETS[kind]
    out = {}
    for col in L.PLAN_COLUMNS[kind]:
        if col.engine == "Month":
            continue
        out[col.engine] = [_num(book.cell(sheet, L.cell_ref(kind, col.engine, i)))
                           for i in range(L.N_MONTHS)]
    return out


def _planning_values(book: Book, kind: str) -> Dict[str, Optional[float]]:
    sheet = L.PLAN_SHEETS[kind]
    return {pc.key: _num(book.cell(sheet, L.planning_ref(kind, pc.key)))
            for pc in L.PLANNING[kind][2]}


def _f(*xs) -> bool:
    return all(x is not None for x in xs)


def _local(kind: str, col: str, i: int, v: Dict[str, List[Optional[float]]],
           pv: Dict[str, Optional[float]], scn: Scenario) -> Optional[float]:
    """What this cell SHOULD show given the student's own upstream cells."""
    p = scn.params
    d = float(scn.demand[i])
    g = lambda name: v[name][i]                                   # noqa: E731
    prev = lambda name: v[name][i - 1]                             # noqa: E731
    prod_col = "Total Production" if kind == "hybrid" else "Regular Production"
    if col == "Beginning Inventory":
        return float(p["beginning_inventory"]) if i == 0 else prev("Ending Inventory")
    if col == "Beginning Backorders":
        return 0.0 if i == 0 else prev("Ending Backorders")
    if col == "Regular Production":
        if kind == "chase":
            t = pv.get("target")
            bi, bb = g("Beginning Inventory"), g("Beginning Backorders")
            return max(0.0, d + t - bi + bb) if _f(t, bi, bb) else None
        if kind == "level":
            return pv.get("level_rate")
        return None
    if col == "Workers":
        if kind == "chase":
            cap, pr = pv.get("capacity"), g("Regular Production")
            return float(math.ceil(pr / cap - 1e-9)) if _f(cap, pr) and cap else None
        if kind == "level":
            return pv.get("level_workers")
        return None
    if col == "Regular Capacity":
        cap, w = pv.get("capacity"), g("Workers")
        return w * cap if _f(cap, w) else None
    if col == "Total Production":
        parts = [g("Regular Production"), g("Overtime Production") or 0.0,
                 g("Subcontract Production") or 0.0]
        return sum(parts) if _f(*parts) else None
    if col in ("Hires", "Layoffs"):
        w = g("Workers")
        wp = float(p["starting_workforce"]) if i == 0 else prev("Workers")
        if not _f(w, wp):
            return None
        return max(0.0, w - wp) if col == "Hires" else max(0.0, wp - w)
    if col in ("Ending Inventory", "Ending Backorders"):
        bi, bb, pr = g("Beginning Inventory"), g("Beginning Backorders"), g(prod_col)
        if not _f(bi, bb, pr):
            return None
        return max(0.0, bi + pr - bb - d) if col == "Ending Inventory" else max(0.0, d + bb - bi - pr)
    rate_of = {"Labor Cost": ("Workers", "regular_labor_cost"), "Hiring Cost": ("Hires", "hiring_cost"),
               "Layoff Cost": ("Layoffs", "layoff_cost"),
               "Overtime Cost": ("Overtime Production", "overtime_cost"),
               "Subcontract Cost": ("Subcontract Production", "subcontract_cost"),
               "Holding Cost": ("Ending Inventory", "holding_cost"),
               "Backorder Cost": ("Ending Backorders", "backorder_cost")}
    if col in rate_of:
        q, k = rate_of[col]
        x = g(q) if q in v else 0.0
        if x is None and q in ("Overtime Production", "Subcontract Production"):
            x = 0.0
        return x * float(p[k]) if x is not None else None
    if col == "Total Cost":
        parts = [g(c) for c in ("Labor Cost", "Hiring Cost", "Layoff Cost", "Overtime Cost",
                                "Subcontract Cost", "Holding Cost", "Backorder Cost") if c in v]
        return sum(parts) if _f(*parts) else None
    return None


def _extra_hint(kind: str, col: str, i: int, v, pv, scn: Scenario, cell: Cell):
    """Targeted hints for the most common Excel/planning mistakes."""
    def hint(x: float) -> str:
        f = (cell.formula or "").upper()
        if i > 0 and f and _LOCK_RE.search(f):
            m = _LOCK_RE.search(f)
            return (f"Your formula uses {m.group(0)} without $ signs, so filling down moved it. "
                    f"Lock it: ${m.group(1)}${m.group(2)}.")
        if col == "Workers" and kind == "chase":
            pr, cap = v["Regular Production"][i], pv.get("capacity")
            if _f(pr, cap) and cap:
                raw = pr / cap
                if abs(x - raw) < 0.01 and abs(raw - round(raw)) > 0.01:
                    return "Workers must be whole: round UP with ROUNDUP(…, 0)."
                if abs(x - math.floor(raw)) < 0.01 and raw > math.floor(raw) + 1e-9:
                    return "Rounding down leaves you short of capacity; use ROUNDUP."
        if col in ("Ending Inventory", "Ending Backorders") and x < -1e-9:
            return ("This can't be negative. Wrap it in MAX(0, …) — a shortage belongs in "
                    "Ending Backorders, extra stock in Ending Inventory.")
        if col in ("Hires", "Layoffs") and i == 0:
            return f"January compares with StartWorkers ({scn['starting_workforce']:,.0f})."
        if col == "Beginning Inventory" and i == 0:
            return "January starts with BeginInv."
        return ""
    return hint


def _grade_plan(book: Book, scn: Scenario, kind: str, key_df: pd.DataFrame,
                require_formula: bool) -> SheetResult:
    sheet = L.PLAN_SHEETS[kind]
    res = SheetResult(kind, sheet)
    v = _plan_values(book, kind)
    pv = _planning_values(book, kind)
    rules = text.column_rules(kind)

    # planning cells
    key_pv = {"capacity": scn.cap_month, "target": scn.target,
              "total_demand": scn.total_demand, "level_rate": engine.level_rate(scn),
              "level_workers": engine.workers_for(scn, engine.level_rate(scn)),
              "ot_limit": scn.cap_month * float(scn["overtime_pct"])}
    for pc in L.PLANNING[kind][2]:
        ref = L.planning_ref(kind, pc.key)
        cell = book.cell(sheet, ref)
        local = None
        if pc.key == "level_rate" and _f(pv.get("total_demand"), pv.get("target")):
            local = (pv["total_demand"] + pv["target"] - float(scn["beginning_inventory"])) / 12
        elif pc.key == "level_workers" and _f(pv.get("level_rate"), pv.get("capacity")) and pv["capacity"]:
            local = float(math.ceil(pv["level_rate"] / pv["capacity"] - 1e-9))
        elif pc.key == "ot_limit" and pv.get("capacity") is not None:
            local = pv["capacity"] * float(scn["overtime_pct"])
        st, msg = _judge(cell, key_pv[pc.key], local, _tol(pc.key), require_formula,
                         text.PLANNING_RULES[pc.key])
        res.cells.append(CellCheck(sheet, ref, pc.label, pc.label, "Planning", st, msg,
                                   _shown(cell)))

    # monthly grid, month by month so the first problem listed is the earliest one
    for i in range(L.N_MONTHS):
        for col in L.PLAN_COLUMNS[kind]:
            if col.kind == "given":
                continue
            ref = L.cell_ref(kind, col.engine, i)
            cell = book.cell(sheet, ref)
            label = f"{col.heading.rstrip(' ✎')}, {scn.months[i]}"
            if col.kind == "decision":
                st, msg = _judge_decision(cell, col.engine)
            else:
                key = float(key_df[col.engine].iloc[i])
                local = _local(kind, col.engine, i, v, pv, scn)
                st, msg = _judge(cell, key, local, _tol(col.engine), require_formula,
                                 rules.get(col.engine, ""),
                                 _extra_hint(kind, col.engine, i, v, pv, scn, cell))
            res.cells.append(CellCheck(sheet, ref, label, col.heading, scn.months[i], st, msg,
                                       _shown(cell)))
    # totals
    for col in L.PLAN_COLUMNS[kind]:
        if not col.total:
            continue
        ref = L.total_ref(kind, col.engine)
        cell = book.cell(sheet, ref)
        vals = v[col.engine]
        if col.kind == "decision":
            vals = [x if x is not None else (0.0 if col.engine != "Regular Production" else None)
                    for x in vals]
        local = sum(vals) if _f(*vals) else None
        st, msg = _judge(cell, float(key_df[col.engine].sum()), local, 0.5, require_formula,
                         text.TOTAL_RULE)
        res.cells.append(CellCheck(sheet, ref, f"{col.heading.rstrip(' ✎')}, total", col.heading,
                                   "Total", st, msg, _shown(cell)))
    return res


def _judge_decision(cell: Cell, engine_name: str) -> Tuple[str, str]:
    if cell.blank:
        if engine_name in ("Overtime Production", "Subcontract Production"):
            return "decision", ""                         # blank = none used
        return "blank", "Enter your decision for this month."
    if cell.formula is not None and cell.value is None:
        return "uncalc", STATUS_HELP["uncalc"]
    x = _num(cell)
    if x is None:
        return "text", "Your decision must be a number."
    if x < 0:
        return "wrong", "A decision can't be negative."
    return "decision", ""


def _grade_decisions(book: Book, scn: Scenario) -> List[CellCheck]:
    out = []
    for i in range(L.N_MONTHS):
        for col in L.PLAN_COLUMNS["hybrid"]:
            if col.kind != "decision":
                continue
            ref = L.cell_ref("hybrid", col.engine, i)
            cell = book.cell(L.SHEET_HYBRID, ref)
            st, msg = _judge_decision(cell, col.engine)
            out.append(CellCheck(L.SHEET_HYBRID, ref, f"{col.heading.rstrip(' ✎')}, "
                                 f"{scn.months[i]}", col.heading, scn.months[i], st, msg,
                                 _shown(cell)))
    return out


def _hybrid_decision(book: Book) -> Optional[Decision]:
    v = _plan_values(book, "hybrid")
    w, r = v["Workers"], v["Regular Production"]
    if not (_f(*w) and _f(*r)):
        return None
    ot = [x or 0.0 for x in v["Overtime Production"]]
    sub = [x or 0.0 for x in v["Subcontract Production"]]
    if any(x < 0 for x in w + r + ot + sub):
        return None
    return Decision(w, r, ot, sub)


def _hybrid_attempted(book: Book) -> bool:
    sheet = L.SHEET_HYBRID
    for name in ("Workers", "Regular Production", "Overtime Production", "Subcontract Production"):
        for i in range(L.N_MONTHS):
            if not book.cell(sheet, L.cell_ref("hybrid", name, i)).blank:
                return True
    return False


# =========================================================================================
# Summary sheet
# =========================================================================================
def _student_metric(kind: str, metric: str, book: Book) -> Optional[float]:
    """The value this Summary cell should show given the student's own plan sheet."""
    sheet = L.PLAN_SHEETS[kind]
    v = _plan_values(book, kind)
    if metric == "Peak Workforce":
        return max(v["Workers"]) if _f(*v["Workers"]) else None
    if metric == "Peak Inventory":
        return max(v["Ending Inventory"]) if _f(*v["Ending Inventory"]) else None
    if metric == "Months With Backorders":
        b = v["Ending Backorders"]
        return float(sum(x > 1e-6 for x in b)) if _f(*b) else None
    col = {"Total Hires": "Hires", "Total Layoffs": "Layoffs"}.get(metric, metric)
    return _num(book.cell(sheet, L.total_ref(kind, col)))


def _grade_summary(book: Book, scn: Scenario, keys: Dict[str, Dict[str, float]],
                   kinds: List[str], require_formula: bool, result: CheckResult) -> SheetResult:
    sheet = L.SHEET_SUMMARY
    res = SheetResult("summary", sheet)
    for row in L.SUMMARY_ROWS:
        r = L.summary_row(row.metric)
        for kind in kinds:
            if row.hybrid_only and kind != "hybrid":
                continue
            ref = f"{L.SUMMARY_PLAN_COL[kind]}{r}"
            cell = book.cell(sheet, ref)
            rule = text.SUMMARY_RULES.get(row.metric, text.SUMMARY_RULES["link"])
            st, msg = _judge(cell, keys[kind][row.metric], _student_metric(kind, row.metric, book),
                             _tol(row.metric), require_formula, rule)
            res.cells.append(CellCheck(sheet, ref, f"{row.label} ({L.PLAN_LABEL[kind]})",
                                       L.PLAN_LABEL[kind], row.label, st, msg, _shown(cell)))
    tc = L.summary_row("Total Cost")
    kc, kl = keys["chase"]["Total Cost"], keys["level"]["Total Cost"]
    sc, sl = _num(book.cell(sheet, f"B{tc}")), _num(book.cell(sheet, f"C{tc}"))

    # lower-cost plan (text from an IF formula)
    cell = book.cell(sheet, L.LOWER_COST_CELL)
    want = {"chase", "level"} if abs(kc - kl) < 0.5 else ({"chase"} if kc < kl else {"level"})
    local = None
    if _f(sc, sl):
        local = {"chase", "level"} if abs(sc - sl) < 0.5 else ({"chase"} if sc < sl else {"level"})
    if cell.blank:
        st, msg = "blank", "Use " + text.SUMMARY_RULES["lower"]
    elif cell.formula is not None and cell.value is None:
        st, msg = "uncalc", STATUS_HELP["uncalc"]
    else:
        got = str(cell.value).strip().lower()
        if _ERR_RE.match(got.upper()):
            st, msg = "error", STATUS_HELP["error"]
        elif got in want:
            st, msg = ("typed", "Right answer, but typed. Let Excel decide: "
                       + text.SUMMARY_RULES["lower"]) if require_formula and cell.formula is None else ("ok", "")
        elif local and got in local:
            st, msg = "carried", STATUS_HELP["carried"]
        else:
            st, msg = "wrong", "Show the plan with the LOWER total cost: " + text.SUMMARY_RULES["lower"]
    res.cells.append(CellCheck(sheet, L.LOWER_COST_CELL, "Lower-cost plan", "Answer",
                               "Lower-cost plan", st, msg, _shown(cell)))

    cell = book.cell(sheet, L.SAVINGS_CELL)
    st, msg = _judge(cell, abs(kc - kl), abs(sc - sl) if _f(sc, sl) else None, 0.5,
                     require_formula, text.SUMMARY_RULES["savings"])
    res.cells.append(CellCheck(sheet, L.SAVINGS_CELL, "Savings of the lower-cost plan", "Answer",
                               "Savings", st, msg, _shown(cell)))

    # recommendation + justification (judgement, not graded for content)
    allowed = ["Chase", "Level"] + (["Hybrid"] if scn.params.get("include_hybrid", True) else [])
    cell = book.cell(sheet, L.RECOMMEND_CELL)
    rec = str(cell.value or "").strip()
    match = next((a for a in allowed if a.lower() == rec.lower()), "")
    result.recommendation = match
    if cell.blank:
        st, msg = "blank", "Choose your recommendation from the drop-down list."
    elif not match:
        st, msg = "wrong", f"Choose one of: {', '.join(allowed)}."
    else:
        st, msg = "ok", ""
    res.cells.append(CellCheck(sheet, L.RECOMMEND_CELL, "Your recommendation", "Answer",
                               "Recommendation", st, msg, rec))

    cell = book.cell(sheet, L.JUSTIFY_CELL)
    just = str(cell.value or "").strip() if not isinstance(cell.value, (int, float)) else ""
    words = len(re.findall(r"[A-Za-z0-9$%][\w$%.,'’-]*", just))
    result.justification, result.words = just, words
    need = int(scn.params.get("min_justification_words", 0) or 0)
    result.min_words = need
    if not just:
        st, msg = "blank", "Write your justification in the large box."
    elif words < need:
        st, msg = "wrong", f"{words} words so far; write at least {need}."
    else:
        st, msg = "ok", ""
    res.cells.append(CellCheck(sheet, L.JUSTIFY_CELL, "Justification", "Answer",
                               "Justification", st, msg, f"{words} words"))
    return res


# =========================================================================================
# Public API
# =========================================================================================
def check_identity(book: Book, scn: Scenario, expected_sid: Optional[str]) -> List[str]:
    m = book.meta
    if m.get("app") != APP_ID:
        return ["This doesn't look like an Aggregate Anxiety workbook from this app (its hidden "
                "identification sheet is missing). Upload the workbook you downloaded in step 2, "
                "or download a fresh one and copy your formulas into it."]
    out = []
    if m.get("model_version", "").split(".")[0] != MODEL_VERSION.split(".")[0]:
        out.append("This workbook was made by a different version of the app. Download a fresh "
                   "workbook in step 2 and copy your work into it.")
    if expected_sid and m.get("student_id", "") != expected_sid:
        out.append(f"This workbook belongs to student ID '{m.get('student_id', '')}', not "
                   f"'{expected_sid}'. Upload your own workbook.")
    if str(m.get("seed", "")) != str(scn.seed) or m.get("config_hash") != scn.config_hash:
        out.append("This workbook was made for a different scenario than the one you are signed "
                   "in with (or your instructor changed the assignment). Download your workbook "
                   "again in step 2 and copy your formulas into it.")
    return out


def _check_given(book: Book, scn: Scenario, kinds: List[str]) -> List[str]:
    out = []
    for kind in kinds:
        sheet = L.PLAN_SHEETS[kind]
        if sheet not in book.sheets:
            out.append(f"The '{sheet}' sheet is missing or was renamed. Keep the sheet names as "
                       "they were.")
            continue
        for i, d in enumerate(scn.demand):
            ref = f"B{L.month_row(i)}"
            if not _close(_num(book.cell(sheet, ref)), float(d), 1e-6):
                out.append(f"{sheet} {ref}: the demand for {scn.months[i]} was changed. Put back "
                           f"{d:,.0f} (given data must not change).")
                break
    if L.SHEET_SUMMARY not in book.sheets:
        out.append("The 'Summary' sheet is missing or was renamed.")
    return out


def check_book(book: Book, scn: Scenario, expected_sid: Optional[str] = None) -> CheckResult:
    res = CheckResult(meta=dict(book.meta), sha256=book.sha256)
    res.fatal = check_identity(book, scn, expected_sid)
    include_hybrid = bool(scn.params.get("include_hybrid", True))
    kinds = ["chase", "level"]
    if not res.fatal:
        res.fatal = _check_given(book, scn, kinds)
    if res.fatal:
        return res
    require = bool(scn.params.get("require_formulas", True))

    keys: Dict[str, Dict[str, float]] = {}
    for kind in kinds:
        df = engine.reference_plan(scn, kind)
        keys[kind] = engine.summarize(df)
        res.sheets[kind] = _grade_plan(book, scn, kind, df, require)

    sum_kinds = list(kinds)
    if include_hybrid and L.SHEET_HYBRID in book.sheets and _hybrid_attempted(book):
        dec = _hybrid_decision(book)
        if dec is None:
            hs = SheetResult("hybrid", L.SHEET_HYBRID, required=False)
            hs.notes.append("Fill every month's Workers and Regular Production (green cells) "
                            "with numbers ≥ 0; then the rest of the sheet can be checked.")
            hs.cells = [c for c in _grade_decisions(book, scn)]
            res.sheets["hybrid"] = hs
        else:
            df = engine.simulate(scn, dec)
            keys["hybrid"] = engine.summarize(df)
            hs = _grade_plan(book, scn, "hybrid", df, require)
            hs.required = False
            res.hybrid_feasibility = engine.check_feasibility(df, scn)
            res.sheets["hybrid"] = hs
            sum_kinds.append("hybrid")
    elif include_hybrid:
        res.sheets["hybrid"] = SheetResult("hybrid", L.SHEET_HYBRID, required=False,
                                           attempted=False,
                                           notes=["Optional — not started."])
    summ = _grade_summary(book, scn, keys, sum_kinds, require, res)
    if "hybrid" in sum_kinds:          # optional work must not block the required Summary
        hyb = [c for c in summ.cells if c.group == L.PLAN_LABEL["hybrid"]]
        summ.cells = [c for c in summ.cells if c.group != L.PLAN_LABEL["hybrid"]]
        res.sheets["hybrid"].cells.extend(hyb)
    res.sheets["summary"] = summ
    if res.recommendation == "Hybrid" and "hybrid" in res.sheets:
        res.sheets["hybrid"].required = True
        res.sheets["hybrid"].notes.append("You recommend the hybrid plan, so it must be complete "
                                          "and correct too.")
    res.plan_totals = keys
    res.uncalculated = any(c.status == "uncalc" for s in res.sheets.values() for c in s.cells)
    return res


def check_workbook(data: bytes, scn: Scenario, expected_sid: Optional[str] = None) -> CheckResult:
    book, err = read_book(data)
    if err:
        return CheckResult(fatal=[err], sha256=hashlib.sha256(data or b"").hexdigest())
    return check_book(book, scn, expected_sid)


def grid(sheet: SheetResult) -> pd.DataFrame:
    """Status grid for display: rows = months/total/planning, columns = headings."""
    rows: Dict[str, Dict[str, str]] = {}
    order: List[str] = []
    cols: List[str] = []
    for c in sheet.cells:
        if c.row_label == "Planning" or c.sheet != sheet.title:
            continue
        if c.row_label not in rows:
            rows[c.row_label] = {}
            order.append(c.row_label)
        if c.group not in cols:
            cols.append(c.group)
        rows[c.row_label][c.group] = c.status
    df = pd.DataFrame([rows[r] for r in order], index=order, columns=cols)
    return df.fillna("")
