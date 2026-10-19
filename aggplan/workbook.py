"""Build the student's personal Excel workbook.

The workbook carries only GIVEN data (demand, parameters with Excel names) and empty,
colour-coded cells. Every calculation is the student's own Excel formula. With
`solution=True` the same layout is filled with reference formulas: that is the instructor
answer key (tools/make_answer_key.py) and the end-to-end test fixture.

Layout: see layout.py. Sheets are protected WITHOUT a password, only to stop accidental
edits of the given data; the cells a student fills are unlocked."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
from typing import Optional

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import MODEL_VERSION, layout as L, text
from .engine import Decision, reference_plan
from .scenario import NAME_KEYS, Scenario

SCHEMA = 1
APP_ID = "aggregate-anxiety-excel"

# --- styles ----------------------------------------------------------------------------
GIVEN_FONT = Font(color="1F4E9E")
GIVEN_FILL = PatternFill("solid", fgColor="FFF2CC")       # given data: blue on light yellow
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")       # your formulas: light blue
DECIDE_FILL = PatternFill("solid", fgColor="E2EFDA")      # your decisions: light green
NA_FILL = PatternFill("solid", fgColor="EDEDED")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOTAL_FILL = PatternFill("solid", fgColor="BDD7EE")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=16, color="1F3864")
H2 = Font(bold=True, size=12, color="1F3864")
NOTE = Font(italic=True, color="595959")
CODE = Font(name="Consolas", color="1F3864")
THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
UNLOCKED = Protection(locked=False)
RED_FILL = PatternFill("solid", bgColor="FFC7CE")
RED_FONT = Font(color="9C0006")
AMBER_FILL = PatternFill("solid", bgColor="FFEB9C")


def _put(ws, ref, value, *, font=None, fill=None, fmt=None, wrap=False, border=False,
         align=None):
    """Write one cell. Strings are stored as TEXT even when they start with '='."""
    c = ws[ref]
    c.value = value
    if isinstance(value, str):
        c.data_type = "s"
    if font is not None:
        c.font = font
    if fill is not None:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if wrap:
        c.alignment = WRAP
    if align is not None:
        c.alignment = align
    if border:
        c.border = BOX
    return c


def _formula(ws, ref, f):
    """A real Excel formula (only ever our own text, never student input)."""
    ws[ref].value = f
    return ws[ref]


def _input(ws, ref, fmt=None, fill=INPUT_FILL):
    c = ws[ref]
    c.fill, c.border, c.protection = fill, BOX, UNLOCKED
    if fmt:
        c.number_format = fmt
    return c


def _protect(ws):
    ws.protection.sheet = True
    ws.protection.formatColumns = False          # False = allowed
    ws.protection.formatRows = False
    ws.protection.formatCells = False


def _comment(t: str) -> Comment:
    c = Comment(t, "Aggregate Anxiety")
    c.width, c.height = 300, 120
    return c


def _sheet_ref(sheet: str) -> str:
    return f"'{sheet}'" if " " in sheet else sheet


# =========================================================================================
# Sheets
# =========================================================================================
def _start_sheet(wb, scn: Scenario, name: str, sid: str):
    ws = wb.active
    ws.title = L.SHEET_START
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 62
    ws.column_dimensions["D"].width = 62
    _put(ws, "B1", str(scn.params["assignment_name"]), font=TITLE)
    _put(ws, "B2", "Juicetification Inc. — aggregate production planning, done in Excel.",
         font=NOTE)
    _put(ws, "B4", "Student name", font=BOLD)
    _put(ws, L.START_NAME_CELL, name or "", border=True, fill=GIVEN_FILL, font=GIVEN_FONT)
    _put(ws, "B5", "Student ID", font=BOLD)
    _put(ws, L.START_ID_CELL, sid or "", border=True, fill=GIVEN_FILL, font=GIVEN_FONT)
    _put(ws, "B6", "Scenario number", font=BOLD)
    _put(ws, "C6", scn.seed, border=True, fill=GIVEN_FILL, font=GIVEN_FONT, fmt="0",
         align=Alignment(horizontal="left"))
    _put(ws, "D4", "This workbook was made for you. Your demand and numbers differ from "
                   "your classmates', so upload only your own file.", font=NOTE, wrap=True)
    ws.merge_cells("D4:D6")

    r = 8
    _put(ws, f"B{r}", "What you will do", font=H2)
    steps = [
        ("1. Scenario", "Read your demand forecast and the given data. Every given number has "
                        "an Excel NAME you can use in formulas (e.g. LaborCost)."),
        ("2. Chase Plan", "Fill the planning cells (column S), then every blue cell for "
                          "January–December, then the totals in row 14. Use formulas only."),
        ("3. Level Plan", "Same as the chase plan, but production and workers are constant: "
                          "work them out in the planning cells first (S2:S6)."),
        ("4. Summary", "Link each plan's totals, find the cheaper plan with IF, then choose "
                       "and justify your recommendation in your own words."),
        ("5. Hybrid Plan", "Optional challenge: choose workers, overtime and subcontracting "
                           "(green cells) and calculate the rest. Try to beat both plans."),
        ("6. Check", "Save as .xlsx and upload it in the app. It tells you which cells to fix. "
                     "Fix, save, upload again — as often as you like."),
    ]
    if not scn.params.get("include_hybrid", True):
        steps = [s for s in steps if not s[0].startswith("5.")]
    for k, (a, b) in enumerate(steps, start=r + 1):
        _put(ws, f"B{k}", a, font=BOLD, wrap=True)
        _put(ws, f"C{k}", b, wrap=True)
        ws.merge_cells(f"C{k}:D{k}")
    r = r + len(steps) + 2

    _put(ws, f"B{r}", "Colour key", font=H2)
    legend = [(GIVEN_FILL, "Given data — do not change"),
              (INPUT_FILL, "Your formula goes here"),
              (DECIDE_FILL, "Your decision (Hybrid Plan only)"),
              (TOTAL_FILL, "Totals row — your SUM formula goes here")]
    for k, (fill, label) in enumerate(legend, start=r + 1):
        _put(ws, f"B{k}", "", fill=fill, border=True)
        _put(ws, f"C{k}", label)
    r = r + len(legend) + 2

    _put(ws, f"B{r}", "The rules of the model", font=H2)
    _put(ws, f"B{r + 1}", text.target_sentence(scn), wrap=True)
    ws.merge_cells(f"B{r + 1}:D{r + 1}")
    ws.row_dimensions[r + 1].height = 32
    r += 2
    _put(ws, f"B{r}", "Column", font=HEAD_FONT, fill=HEAD_FILL)
    _put(ws, f"C{r}", "Chase plan", font=HEAD_FONT, fill=HEAD_FILL)
    _put(ws, f"D{r}", "Level plan (where different)", font=HEAD_FONT, fill=HEAD_FILL)
    chase, level = text.column_rules("chase"), text.column_rules("level")
    for col in L.PLAN_COLUMNS["chase"]:
        if col.kind != "calc":
            continue
        r += 1
        _put(ws, f"B{r}", col.heading, font=BOLD, border=True, wrap=True)
        _put(ws, f"C{r}", chase[col.engine], border=True, wrap=True)
        lv = level[col.engine]
        _put(ws, f"D{r}", lv if lv != chase[col.engine] else "same", border=True, wrap=True)
    r += 1
    _put(ws, f"B{r}", "Row 14 totals", font=BOLD, border=True)
    _put(ws, f"C{r}", text.TOTAL_RULE, border=True, wrap=True)
    _put(ws, f"D{r}", "same", border=True)
    r += 2
    _put(ws, f"B{r}", "Stuck on Excel itself? See the Excel Help sheet.", font=NOTE)
    _protect(ws)


def _scenario_sheet(wb, scn: Scenario):
    ws = wb.create_sheet(L.SHEET_SCENARIO)
    ws.sheet_view.showGridLines = False
    _put(ws, "A1", "Scenario: your given data", font=TITLE)
    _put(ws, "A2", "Use the Excel NAMES in column G in your formulas, e.g. =F2*LaborCost. "
                   "Never retype these numbers.", font=NOTE)
    _put(ws, "A4", "Demand forecast", font=H2)
    for j, h in enumerate(["Month", "Demand (bottles)"]):
        _put(ws, f"{get_column_letter(j + 1)}5", h, font=HEAD_FONT, fill=HEAD_FILL, border=True)
    for i, (m, d) in enumerate(zip(scn.months, scn.demand)):
        _put(ws, f"A{6 + i}", m, font=GIVEN_FONT, fill=GIVEN_FILL, border=True)
        _put(ws, f"B{6 + i}", float(d), font=GIVEN_FONT, fill=GIVEN_FILL, border=True,
             fmt="#,##0")

    _put(ws, "E4", "Given data and their Excel names", font=H2)
    for j, h in enumerate(["Given data", "Value", "Excel name", "Unit"]):
        _put(ws, f"{get_column_letter(5 + j)}5", h, font=HEAD_FONT, fill=HEAD_FILL,
             border=True)
    for i, (nm, key, meaning, unit) in enumerate(NAME_KEYS):
        r = 6 + i
        v = scn.name_value(key)
        fmt = "0%" if key == "overtime_pct" else ("$#,##0.00" if unit.startswith("$")
                                                  else "#,##0")
        _put(ws, f"E{r}", meaning, border=True)
        _put(ws, f"F{r}", v, font=GIVEN_FONT, fill=GIVEN_FILL, border=True, fmt=fmt)
        _put(ws, f"G{r}", nm, font=CODE, border=True)
        _put(ws, f"H{r}", unit, border=True)
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"{L.SHEET_SCENARIO}!$F${r}")
    for col, w in zip("ABCDEFGH", [13, 17, 3, 3, 66, 13, 17, 16]):
        ws.column_dimensions[col].width = w
    _protect(ws)


def _plan_sheet(wb, scn: Scenario, kind: str):
    ws = wb.create_sheet(L.PLAN_SHEETS[kind])
    cols = L.PLAN_COLUMNS[kind]
    rules = text.column_rules(kind)
    ws.row_dimensions[1].height = 48
    for j, c in enumerate(cols):
        letter = get_column_letter(j + 1)
        _put(ws, f"{letter}1", c.heading, font=HEAD_FONT, fill=HEAD_FILL, border=True,
             align=CENTER)
        tip = f"{c.heading}\nUnit: {c.unit}" if c.unit else c.heading
        if c.engine in rules and c.kind != "given":
            tip += f"\n\n{rules[c.engine]}"
        ws[f"{letter}1"].comment = _comment(tip)
        ws.column_dimensions[letter].width = 13 if j else 11
        for i in range(L.N_MONTHS):
            r = L.month_row(i)
            if j == 0:
                _put(ws, f"A{r}", scn.months[i], font=GIVEN_FONT, fill=GIVEN_FILL, border=True)
            elif j == 1:
                _put(ws, f"B{r}", float(scn.demand[i]), font=GIVEN_FONT, fill=GIVEN_FILL,
                     border=True, fmt="#,##0")
            else:
                _input(ws, f"{letter}{r}", c.fmt,
                       DECIDE_FILL if c.kind == "decision" else INPUT_FILL)
        tr = L.TOTAL_ROW
        if j == 0:
            _put(ws, f"A{tr}", "Total", font=BOLD, fill=TOTAL_FILL, border=True)
        elif c.total:
            _input(ws, f"{letter}{tr}", c.fmt, TOTAL_FILL)
            ws[f"{letter}{tr}"].font = BOLD
        else:
            _put(ws, f"{letter}{tr}", None, fill=NA_FILL, border=True)
    ws.freeze_panes = "C2"

    # planning block
    lcol, vcol, cells = L.PLANNING[kind]
    ws.column_dimensions[lcol].width = 52
    ws.column_dimensions[vcol].width = 14
    _put(ws, f"{lcol}1", "Planning calculations", font=HEAD_FONT, fill=HEAD_FILL, border=True,
         align=CENTER)
    _put(ws, f"{vcol}1", "Your formula", font=HEAD_FONT, fill=HEAD_FILL, border=True,
         align=CENTER)
    for i, pc in enumerate(cells):
        r = L.FIRST_ROW + i
        _put(ws, f"{lcol}{r}", f"{pc.label} ({pc.unit})", border=True)
        _input(ws, f"{vcol}{r}", "#,##0.##")
        ws[f"{vcol}{r}"].comment = _comment(f"{pc.label}\n\n{text.PLANNING_RULES[pc.key]}")
    r = L.FIRST_ROW + len(cells) + 1
    tips = {
        "chase": ["Tip: every month uses the SAME capacity cell, so lock it: $S$2.",
                  "Do January first (it uses BeginInv and StartWorkers),",
                  "then February, then fill February down to December."],
        "level": ["Tip: every month uses the SAME level production and workforce:",
                  "=$S$5 and =$S$6 (the $ signs keep them fixed when you fill down).",
                  "Hires and Layoffs happen only in January."],
        "hybrid": ["Green cells are YOUR decisions (type numbers or formulas).",
                   "Everything else is calculated, exactly like the other plans.",
                   "Limits: overtime ≤ OTLimitPct of regular capacity and only once regular",
                   "capacity is full; subcontracting ≤ SubCapacity; inventory ≤ MaxInventory;",
                   "no backorders left in December; December inventory ≥ EndTarget."],
    }[kind]
    for k, t in enumerate(tips):
        _put(ws, f"{lcol}{r + k}", t, font=NOTE)

    # conditional formatting: backorders red, inventory over storage amber
    bl = L.col_letter(kind, "Ending Backorders")
    il = L.col_letter(kind, "Ending Inventory")
    ws.conditional_formatting.add(
        f"{bl}{L.FIRST_ROW}:{bl}{L.LAST_ROW}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL, font=RED_FONT))
    ws.conditional_formatting.add(
        f"{il}{L.FIRST_ROW}:{il}{L.LAST_ROW}",
        FormulaRule(formula=[f"AND(ISNUMBER({il}{L.FIRST_ROW}),{il}{L.FIRST_ROW}>MaxInventory)"],
                    fill=AMBER_FILL))
    leg = r + len(tips) + 1
    _put(ws, f"{lcol}{leg}", "Red backorders = customers waiting. Amber inventory = over "
                             "storage capacity.", font=NOTE)

    # chart: demand vs production
    pl = L.col_letter(kind, "Total Production" if kind == "hybrid" else "Regular Production")
    ch = LineChart()
    ch.title = "Demand vs. production"
    ch.y_axis.title = "bottles"
    ch.height, ch.width = 7.5, 16
    for letter in ("B", pl):
        idx = ord(letter) - 64
        ch.add_data(Reference(ws, min_col=idx, min_row=1, max_row=L.LAST_ROW), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=L.FIRST_ROW, max_row=L.LAST_ROW))
    ws.add_chart(ch, f"A{L.TOTAL_ROW + 3}")
    _protect(ws)
    return ws


def _summary_sheet(wb, scn: Scenario, include_hybrid: bool):
    ws = wb.create_sheet(L.SHEET_SUMMARY)
    ws.sheet_view.showGridLines = False
    _put(ws, "A1", "Summary: compare the plans and recommend one", font=TITLE)
    h = L.SUMMARY_HEAD_ROW
    heads = ["Measure", "Chase", "Level", "Hybrid (optional)" if include_hybrid else "",
             "How to calculate it"]
    for j, t in enumerate(heads):
        if t:
            _put(ws, f"{get_column_letter(j + 1)}{h}", t, font=HEAD_FONT, fill=HEAD_FILL,
                 border=True, align=CENTER)
    for row in L.SUMMARY_ROWS:
        r = L.summary_row(row.metric)
        bold = BOLD if row.metric == "Total Cost" else None
        _put(ws, f"A{r}", row.label, border=True, font=bold)
        for kind, letter in L.SUMMARY_PLAN_COL.items():
            if kind == "hybrid" and not include_hybrid:
                continue
            if row.hybrid_only and kind != "hybrid":
                _put(ws, f"{letter}{r}", "n/a", fill=NA_FILL, border=True,
                     align=Alignment(horizontal="center"))
            else:
                _input(ws, f"{letter}{r}", row.fmt)
                if bold:
                    ws[f"{letter}{r}"].font = bold
        _put(ws, f"E{r}", row.hint, font=NOTE)
    lc = L.LOWER_COST_CELL
    r = int(lc[1:])
    _put(ws, f"A{r}", "Lower-cost plan: Chase or Level?", font=BOLD, border=True)
    _input(ws, lc)
    _put(ws, f"E{r}", text.SUMMARY_RULES["lower"], font=NOTE)
    r = int(L.SAVINGS_CELL[1:])
    _put(ws, f"A{r}", "Savings of the lower-cost plan ($)", font=BOLD, border=True)
    _input(ws, L.SAVINGS_CELL, "$#,##0.00")
    _put(ws, f"E{r}", text.SUMMARY_RULES["savings"], font=NOTE)

    r = int(L.RECOMMEND_CELL[1:])
    _put(ws, f"A{r}", "Your recommendation", font=BOLD, border=True)
    _input(ws, L.RECOMMEND_CELL)
    choices = '"Chase,Level,Hybrid"' if include_hybrid else '"Chase,Level"'
    dv = DataValidation(type="list", formula1=choices, allow_blank=True,
                        prompt="Pick the plan you recommend", promptTitle="Recommendation")
    ws.add_data_validation(dv)
    dv.add(L.RECOMMEND_CELL)
    _put(ws, f"E{r}", "Pick from the drop-down list.", font=NOTE)
    r = int(L.JUSTIFY_CELL[1:])
    _put(ws, f"A{r}", "Why? Justify it in your own words.", font=BOLD, border=True, wrap=True)
    words = int(scn.params.get("min_justification_words", 0) or 0)
    _put(ws, f"A{r + 1}", (f"At least {words} words. " if words else "")
         + "Use your numbers: cost, workforce changes, inventory, backorders, and the risks "
           "your plan takes.", font=NOTE, wrap=True)
    ws.merge_cells(f"A{r + 1}:A{r + 5}")
    rng = f"{L.JUSTIFY_CELL}:E{L.JUSTIFY_LAST_ROW}"
    ws.merge_cells(rng)
    c = _input(ws, L.JUSTIFY_CELL)
    c.alignment = WRAP
    for col, w in zip("ABCDE", [40, 17, 17, 19, 52]):
        ws.column_dimensions[col].width = w
    _protect(ws)


def _help_sheet(wb):
    ws = wb.create_sheet(L.SHEET_HELP)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 70
    _put(ws, "A1", "Excel help", font=TITLE)
    r = 3
    _put(ws, f"A{r}", "Habits that make this easy", font=H2)
    for t, d in text.EXCEL_TIPS:
        r += 1
        _put(ws, f"A{r}", t, font=BOLD, wrap=True)
        _put(ws, f"B{r}", d, wrap=True)
        ws.merge_cells(f"B{r}:C{r}")
        ws.row_dimensions[r].height = 45
    r += 2
    _put(ws, f"A{r}", "Functions you will use", font=H2)
    for fn, ex, d in text.EXCEL_FUNCTIONS:
        r += 1
        _put(ws, f"A{r}", fn, font=BOLD)
        _put(ws, f"B{r}", ex, font=CODE)
        _put(ws, f"C{r}", d, wrap=True)
    r += 2
    _put(ws, f"A{r}", "What an error means", font=H2)
    for e, d in text.EXCEL_ERRORS:
        r += 1
        _put(ws, f"A{r}", e, font=BOLD)
        _put(ws, f"B{r}", d, wrap=True)
        ws.merge_cells(f"B{r}:C{r}")
        ws.row_dimensions[r].height = 30
    r += 2
    _put(ws, f"A{r}", "Practice: absolute references", font=H2)
    _put(ws, f"A{r + 1}", "Holding cost = bottles × the ONE rate in B" + str(r + 3)
         + ". Type a formula in C" + str(r + 4) + " and fill it down. Without a $ sign, "
         "the rows below show 0 — that is the lesson.", wrap=True)
    ws.merge_cells(f"A{r + 1}:C{r + 1}")
    ws.row_dimensions[r + 1].height = 32
    _put(ws, f"A{r + 3}", "Rate ($/bottle)", font=BOLD)
    _put(ws, f"B{r + 3}", 0.25, fill=GIVEN_FILL, font=GIVEN_FONT, fmt="$0.00")
    for k, q in enumerate([1200, 850, 1500, 2300, 975, 1840]):
        _put(ws, f"A{r + 4 + k}", f"Month {k + 1}")
        _put(ws, f"B{r + 4 + k}", q, fill=GIVEN_FILL, font=GIVEN_FONT, fmt="#,##0")
        _input(ws, f"C{r + 4 + k}", "$#,##0.00")
    _protect(ws)


def _meta_sheet(wb, scn: Scenario, name: str, sid: str, game: Optional[str]):
    ws = wb.create_sheet(L.SHEET_META)
    created = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    wid = hashlib.sha256(f"{scn.seed}|{sid}|{created}".encode()).hexdigest()[:12]
    rows = [("app", APP_ID), ("schema", SCHEMA), ("model_version", MODEL_VERSION),
            ("seed", scn.seed), ("config_hash", scn.config_hash), ("student_id", sid or ""),
            ("student_name", name or ""), ("game", game or ""), ("created_utc", created),
            ("workbook_id", wid)]
    for i, (k, v) in enumerate(rows, start=1):
        _put(ws, f"A{i}", k)
        _put(ws, f"B{i}", v if not isinstance(v, str) else str(v))
    ws.sheet_state = "veryHidden"
    _protect(ws)


# =========================================================================================
# Answer key formulas
# =========================================================================================
def _fill_solution(wb, scn: Scenario, hybrid: Optional[Decision], include_hybrid: bool):
    for kind in ("chase", "level") + (("hybrid",) if hybrid is not None else ()):
        ws = wb[L.PLAN_SHEETS[kind]]
        c = lambda name: L.col_letter(kind, name)              # noqa: E731
        if kind != "hybrid":
            _formula(ws, L.planning_ref(kind, "capacity"), "=BottlesPerHour*HoursPerDay*DaysPerMonth")
            _formula(ws, L.planning_ref(kind, "target"), "=EndTarget")
        else:
            _formula(ws, L.planning_ref(kind, "capacity"), "=BottlesPerHour*HoursPerDay*DaysPerMonth")
            _formula(ws, L.planning_ref(kind, "ot_limit"), f"={L.planning_ref(kind, 'capacity')}*OTLimitPct")
        if kind == "level":
            _formula(ws, L.planning_ref(kind, "total_demand"), "=SUM(B2:B13)")
            _formula(ws, L.planning_ref(kind, "level_rate"), "=(S4+S3-BeginInv)/12")
            _formula(ws, L.planning_ref(kind, "level_workers"), "=ROUNDUP(S5/S2,0)")
        for i in range(L.N_MONTHS):
            r, p = L.month_row(i), L.month_row(i) - 1
            prod = c("Total Production") if kind == "hybrid" else c("Regular Production")
            f = {}
            f["Beginning Inventory"] = "=BeginInv" if i == 0 else f"={c('Ending Inventory')}{p}"
            f["Beginning Backorders"] = "=0" if i == 0 else f"={c('Ending Backorders')}{p}"
            w, wp = f"{c('Workers')}{r}", ("StartWorkers" if i == 0 else f"{c('Workers')}{p}")
            if kind == "chase":
                f["Regular Production"] = f"=MAX(0,B{r}+$S$3-C{r}+D{r})"
                f["Workers"] = f"=ROUNDUP(E{r}/$S$2,0)"
            elif kind == "level":
                f["Regular Production"] = "=$S$5"
                f["Workers"] = "=$S$6"
            else:
                f["Workers"] = hybrid.workers[i]
                f["Regular Production"] = hybrid.regular[i]
                f["Overtime Production"] = hybrid.overtime[i] if hybrid.overtime else 0
                f["Subcontract Production"] = hybrid.subcontract[i] if hybrid.subcontract else 0
                f["Regular Capacity"] = f"={w}*$Y$2"
                f["Total Production"] = (f"={c('Regular Production')}{r}+"
                                         f"{c('Overtime Production')}{r}+"
                                         f"{c('Subcontract Production')}{r}")
                f["Overtime Cost"] = f"={c('Overtime Production')}{r}*OTCost"
                f["Subcontract Cost"] = f"={c('Subcontract Production')}{r}*SubCost"
            f["Hires"] = f"=MAX(0,{w}-{wp})"
            f["Layoffs"] = f"=MAX(0,{wp}-{w})"
            f["Ending Inventory"] = f"=MAX(0,C{r}+{prod}{r}-D{r}-B{r})"
            f["Ending Backorders"] = f"=MAX(0,B{r}+D{r}-C{r}-{prod}{r})"
            f["Labor Cost"] = f"={w}*LaborCost"
            f["Hiring Cost"] = f"={c('Hires')}{r}*HireCost"
            f["Layoff Cost"] = f"={c('Layoffs')}{r}*LayoffCost"
            f["Holding Cost"] = f"={c('Ending Inventory')}{r}*HoldCost"
            f["Backorder Cost"] = f"={c('Ending Backorders')}{r}*BackorderCost"
            f["Total Cost"] = f"=SUM({c('Labor Cost')}{r}:{c('Backorder Cost')}{r})"
            for name, v in f.items():
                ws[f"{c(name)}{r}"].value = v
        for col in L.PLAN_COLUMNS[kind]:
            if col.total:
                x = c(col.engine)
                _formula(ws, f"{x}{L.TOTAL_ROW}", f"=SUM({x}{L.FIRST_ROW}:{x}{L.LAST_ROW})")

    ws = wb[L.SHEET_SUMMARY]
    for kind, letter in L.SUMMARY_PLAN_COL.items():
        if kind == "hybrid" and hybrid is None:
            continue
        sh = _sheet_ref(L.PLAN_SHEETS[kind])
        for row in L.SUMMARY_ROWS:
            if row.hybrid_only and kind != "hybrid":
                continue
            r = L.summary_row(row.metric)
            if row.metric == "Total Hires":
                f = f"={sh}!{L.total_ref(kind, 'Hires')}"
            elif row.metric == "Total Layoffs":
                f = f"={sh}!{L.total_ref(kind, 'Layoffs')}"
            elif row.metric == "Peak Workforce":
                x = L.col_letter(kind, "Workers")
                f = f"=MAX({sh}!{x}2:{x}13)"
            elif row.metric == "Peak Inventory":
                x = L.col_letter(kind, "Ending Inventory")
                f = f"=MAX({sh}!{x}2:{x}13)"
            elif row.metric == "Months With Backorders":
                x = L.col_letter(kind, "Ending Backorders")
                f = f'=COUNTIF({sh}!{x}2:{x}13,">0")'
            else:
                f = f"={sh}!{L.total_ref(kind, row.metric)}"
            _formula(ws, f"{letter}{r}", f)
    tc = L.summary_row("Total Cost")
    _formula(ws, L.LOWER_COST_CELL, f'=IF(B{tc}<C{tc},"Chase","Level")')
    _formula(ws, L.SAVINGS_CELL, f"=ABS(B{tc}-C{tc})")
    chase_cost = reference_plan(scn, "chase")["Total Cost"].sum()
    level_cost = reference_plan(scn, "level")["Total Cost"].sum()
    ws[L.RECOMMEND_CELL].value = "Chase" if chase_cost <= level_cost else "Level"
    ws[L.JUSTIFY_CELL].value = (
        "ANSWER KEY. The recommendation and its justification are the student's own "
        "judgement and are graded by the instructor, not by the checker. " * 3).strip()


# =========================================================================================
# Public API
# =========================================================================================
def build_workbook(scn: Scenario, *, student_name: str = "", student_id: str = "",
                   game: Optional[str] = None, solution: bool = False,
                   hybrid_decision: Optional[Decision] = None) -> bytes:
    include_hybrid = bool(scn.params.get("include_hybrid", True))
    wb = Workbook()
    _start_sheet(wb, scn, student_name, student_id)
    _scenario_sheet(wb, scn)
    _plan_sheet(wb, scn, "chase")
    _plan_sheet(wb, scn, "level")
    if include_hybrid:
        _plan_sheet(wb, scn, "hybrid")
    _summary_sheet(wb, scn, include_hybrid)
    _help_sheet(wb)
    _meta_sheet(wb, scn, student_name, student_id, game)
    if solution:
        _fill_solution(wb, scn, hybrid_decision if include_hybrid else None, include_hybrid)
    wb.active = 0
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def file_name(student_name: str, student_id: str, solution: bool = False) -> str:
    base = "".join(ch for ch in (student_name or student_id or "student")
                   if ch.isalnum() or ch in " -_").strip().replace(" ", "_") or "student"
    return f"Aggregate_Anxiety_{base}{'_ANSWER_KEY' if solution else ''}.xlsx"
