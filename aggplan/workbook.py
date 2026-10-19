"""Build the student's personal Excel workbook.

The workbook carries only GIVEN data (demand, parameters with Excel names) and empty,
colour-coded space. On each plan sheet the student designs the worksheet: they type their
own column headings in row 1, their formulas for January–December below, and the yearly
totals in row 14. Planning cells below the grid hold one-off calculations.

With `solution=True` the same workbook is filled with headings and reference formulas:
the instructor answer key (tools/make_answer_key.py) and the end-to-end test fixtures.
`columns` lets a test put the columns in any order with any headings.

Sheets are protected WITHOUT a password, only to stop accidental edits of the given data;
every cell a student fills is unlocked."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import MODEL_VERSION, hints, layout as L, text
from .engine import Decision, reference_plan
from .scenario import Scenario

SCHEMA = 2
APP_ID = "aggregate-anxiety-excel"

# --- styles ----------------------------------------------------------------------------
GIVEN_FONT = Font(color="1F4E9E")
GIVEN_FILL = PatternFill("solid", fgColor="FFF2CC")       # given data: blue on light yellow
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")       # your formulas: light blue
HEADING_FILL = PatternFill("solid", fgColor="FCE4D6")     # your column headings: peach
NA_FILL = PatternFill("solid", fgColor="EDEDED")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(bold=True, color="FFFFFF")
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


def _sheet_ref(sheet: str) -> str:
    return f"'{sheet}'" if " " in sheet else sheet


def _pieces(scn: Scenario) -> List[str]:
    return ["chase", "level"] + (["hybrid"] if scn.params.get("include_hybrid", True) else [])


# =========================================================================================
# Sheets
# =========================================================================================
def _start_sheet(wb, scn: Scenario, name: str, sid: str):
    ws = wb.active
    ws.title = L.SHEET_START
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 60
    _put(ws, "B1", str(scn.params["assignment_name"]), font=TITLE)
    _put(ws, "B2", "Juicetification Inc. — aggregate production planning, built by you in Excel.",
         font=NOTE)
    _put(ws, "B4", "Student name", font=BOLD)
    _put(ws, L.START_NAME_CELL, name or "", border=True, fill=GIVEN_FILL, font=GIVEN_FONT)
    _put(ws, "B5", "Student ID", font=BOLD)
    _put(ws, L.START_ID_CELL, sid or "", border=True, fill=GIVEN_FILL, font=GIVEN_FONT)
    _put(ws, "B6", "Scenario number", font=BOLD)
    _put(ws, "C6", scn.seed, border=True, fill=GIVEN_FILL, font=GIVEN_FONT, fmt="0",
         align=Alignment(horizontal="left"))
    _put(ws, "D4", "This workbook was made for you: your demand and numbers differ from your "
                   "classmates'. Upload only your own file.", font=NOTE, wrap=True)
    ws.merge_cells("D4:D6")

    r = 8
    _put(ws, f"B{r}", "How this works", font=H2)
    steps = [("Pieces", "You build and submit this workbook one piece at a time in the app: "
                        "the chase plan, then the level plan"
                        + (", then the hybrid plan" if "hybrid" in _pieces(scn) else "")
                        + ", then the comparison on the Summary sheet."),
             ("Design it", "On each plan sheet only Month and Demand are given. YOU decide "
                           "which columns the plan needs: type a heading in row 1 for each one "
                           "(peach cells), then the formulas for January–December below it, "
                           "and the yearly totals in row 14."),
             ("Plan first", "Fill the planning cells under the grid first — the one-off numbers "
                            "the months will use."),
             ("Use formulas", "Every number you work out must be an Excel formula (start with =). "
                              "Use the names on the Scenario sheet, e.g. =…*LaborCost."),
             ("Check", "Save as .xlsx and upload it at that piece's step in the app. You'll see "
                       "which cells are right; each new attempt unlocks more help on what is "
                       "still wrong. When a piece is all correct, submit it and move on.")]
    for k, (a, b) in enumerate(steps, start=r + 1):
        _put(ws, f"B{k}", a, font=BOLD, wrap=True)
        _put(ws, f"C{k}", b, wrap=True)
        ws.merge_cells(f"C{k}:D{k}")
        ws.row_dimensions[k].height = 32           # Excel never auto-fits merged cells
    r = r + len(steps) + 2

    _put(ws, f"B{r}", "Colour key", font=H2)
    legend = [(GIVEN_FILL, "Given data — do not change"),
              (HEADING_FILL, "Your column heading goes here (row 1)"),
              (INPUT_FILL, "Your formula goes here"),
              (TOTAL_FILL, "Yearly totals row (row 14)")]
    for k, (fill, label) in enumerate(legend, start=r + 1):
        _put(ws, f"B{k}", "", fill=fill, border=True)
        _put(ws, f"C{k}", label)
    r = r + len(legend) + 2

    _put(ws, f"B{r}", "Your company's planning policies (every plan must follow them)", font=H2)
    r += 1
    for title, summary, details in text.policies(scn):
        _put(ws, f"B{r}", title, font=BOLD, wrap=True)
        for line in [summary] + details:
            _put(ws, f"C{r}", f"• {line}", wrap=True)
            ws.merge_cells(f"C{r}:D{r}")
            ws.row_dimensions[r].height = 15 * (1 + len(line) // 110)
            r += 1
        r += 1
    r += 1
    _put(ws, f"B{r}", "Questions every plan must answer, for each month", font=H2)
    for k, q in enumerate(hints.build_questions("chase"), start=1):
        _put(ws, f"B{r + k}", f"{k}.", align=Alignment(horizontal="right"))
        _put(ws, f"C{r + k}", q)
    r += len(hints.build_questions("chase")) + 2
    _put(ws, f"B{r}", "Each question becomes a column. Which columns does each one depend on? "
                      "That is your formula.", font=NOTE)
    _put(ws, f"B{r + 1}", "Stuck on Excel itself? See the Excel Help sheet.", font=NOTE)
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
    for i, (nm, key, meaning, unit) in enumerate(scn.name_keys()):
        r = 6 + i
        fmt = "0%" if key == "overtime_pct" else ("$#,##0.00" if unit.startswith("$")
                                                  else "#,##0")
        _put(ws, f"E{r}", meaning, border=True)
        _put(ws, f"F{r}", scn.name_value(key), font=GIVEN_FONT, fill=GIVEN_FILL, border=True,
             fmt=fmt)
        _put(ws, f"G{r}", nm, font=CODE, border=True)
        _put(ws, f"H{r}", unit, border=True)
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"{L.SHEET_SCENARIO}!$F${r}")
    r = 6 + len(scn.name_keys()) + 2
    _put(ws, f"A{r}", "Planning policies — how your company handles safety stock, backorders "
                      "and workers", font=H2)
    r += 1
    for title, summary, details in text.policies(scn):
        _put(ws, f"A{r}", title, font=HEAD_FONT, fill=HEAD_FILL)
        ws.merge_cells(f"A{r}:H{r}")
        r += 1
        for line in [summary] + details:
            _put(ws, f"A{r}", f"• {line}", wrap=True)
            ws.merge_cells(f"A{r}:H{r}")
            ws.row_dimensions[r].height = 15 * (1 + len(line) // 140)   # merged: no auto-fit
            r += 1
        r += 1
    for col, w in zip("ABCDEFGH", [13, 17, 3, 3, 66, 13, 17, 16]):
        ws.column_dimensions[col].width = w
    _protect(ws)


_PLAN_INTRO = {
    "chase": "CHASE PLAN — change the workforce every month so production follows demand.",
    "level": "LEVEL PLAN — one steady production rate and workforce all year.",
    "hybrid": "HYBRID PLAN — your own strategy: choose workers, regular production, overtime "
              "and subcontracting each month, then calculate the results.",
}


def _plan_sheet(wb, scn: Scenario, kind: str):
    ws = wb.create_sheet(L.PLAN_SHEETS[kind])
    ws.row_dimensions[1].height = 48
    _put(ws, "A1", "Month", font=HEAD_FONT, fill=HEAD_FILL, border=True, align=CENTER)
    _put(ws, "B1", "Demand", font=HEAD_FONT, fill=HEAD_FILL, border=True, align=CENTER)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws["C1"].comment = Comment("Type your first column heading here, then continue to the right. "
                               "One column for each question your plan must answer.",
                               "Aggregate Anxiety")
    for j in range(L.FIRST_STUDENT_COL, L.LAST_STUDENT_COL + 1):
        letter = get_column_letter(j)
        ws.column_dimensions[letter].width = 13
        h = _input(ws, f"{letter}1", fill=HEADING_FILL)
        h.alignment, h.font = CENTER, BOLD
        for i in range(L.N_MONTHS):
            _input(ws, f"{letter}{L.month_row(i)}", "#,##0.##")
        _input(ws, f"{letter}{L.TOTAL_ROW}", "#,##0.##", TOTAL_FILL).font = BOLD
    for i in range(L.N_MONTHS):
        r = L.month_row(i)
        _put(ws, f"A{r}", scn.months[i], font=GIVEN_FONT, fill=GIVEN_FILL, border=True)
        _put(ws, f"B{r}", float(scn.demand[i]), font=GIVEN_FONT, fill=GIVEN_FILL, border=True,
             fmt="#,##0")
    _put(ws, f"A{L.TOTAL_ROW}", "Year total", font=BOLD, fill=TOTAL_FILL, border=True)
    _put(ws, f"B{L.TOTAL_ROW}", None, fill=NA_FILL, border=True)
    ws.freeze_panes = "C2"

    p0 = L.PLANNING_FIRST_ROW
    _put(ws, f"A{p0 - 1}", "Planning calculations — do these first", font=H2)
    for i, pc in enumerate(L.PLANNING[kind]):
        r = p0 + i
        _put(ws, f"A{r}", f"{pc.label} ({pc.unit})", border=True, wrap=True)
        ws.merge_cells(f"A{r}:B{r}")
        _input(ws, L.planning_ref(kind, pc.key), "#,##0.##")
        ws.row_dimensions[r].height = 30
    r = p0 + len(L.PLANNING[kind]) + 1
    _put(ws, f"A{r}", _PLAN_INTRO[kind], font=BOLD)
    _put(ws, f"A{r + 1}", "Row 1: your headings. Rows 2–13: your formulas for January–December. "
                          "Row 14: yearly totals. Add helper columns if they help you think.",
         font=NOTE)
    if kind == "hybrid":
        _put(ws, f"A{r + 2}", "Limits: overtime only once regular capacity is full and at most "
                              "OTLimitPct of it; subcontracting at most SubCapacity; inventory at "
                              "most MaxInventory; no backorders left in December; December "
                              "inventory at least EndTarget.", font=NOTE)
    _protect(ws)


def _summary_sheet(wb, scn: Scenario):
    include_hybrid = "hybrid" in _pieces(scn)
    ws = wb.create_sheet(L.SHEET_SUMMARY)
    ws.sheet_view.showGridLines = False
    _put(ws, "A1", "Summary: compare the plans and recommend one", font=TITLE)
    h = L.SUMMARY_HEAD_ROW
    heads = ["Measure", "Chase", "Level", "Hybrid" if include_hybrid else "", "Where it comes from"]
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
    r = int(L.LOWER_COST_CELL[1:])
    _put(ws, f"A{r}", "Lower-cost plan: Chase or Level?", font=BOLD, border=True)
    _input(ws, L.LOWER_COST_CELL)
    _put(ws, f"E{r}", "let Excel decide with a formula", font=NOTE)
    r = int(L.SAVINGS_CELL[1:])
    _put(ws, f"A{r}", "Savings of the lower-cost plan ($)", font=BOLD, border=True)
    _input(ws, L.SAVINGS_CELL, "$#,##0.00")
    _put(ws, f"E{r}", "a positive number", font=NOTE)
    r = int(L.RECOMMEND_CELL[1:])
    _put(ws, f"A{r}", "Your recommendation", font=BOLD, border=True)
    _input(ws, L.RECOMMEND_CELL)
    choices = '"Chase,Level,Hybrid"' if include_hybrid else '"Chase,Level"'
    dv = DataValidation(type="list", formula1=choices, allow_blank=True,
                        prompt="Pick the plan you recommend", promptTitle="Recommendation")
    ws.add_data_validation(dv)
    dv.add(L.RECOMMEND_CELL)
    _put(ws, f"E{r}", "pick from the drop-down list", font=NOTE)
    r = int(L.JUSTIFY_CELL[1:])
    _put(ws, f"A{r}", "Why? Justify it in your own words.", font=BOLD, border=True, wrap=True)
    words = int(scn.params.get("min_justification_words", 0) or 0)
    _put(ws, f"A{r + 1}", (f"At least {words} words. " if words else "")
         + "Use your numbers: cost, workforce changes, inventory, backorders, and the risks "
           "your plan takes.", font=NOTE, wrap=True)
    ws.merge_cells(f"A{r + 1}:A{r + 5}")
    ws.merge_cells(f"{L.JUSTIFY_CELL}:E{L.JUSTIFY_LAST_ROW}")
    _input(ws, L.JUSTIFY_CELL).alignment = WRAP
    for col, w in zip("ABCDE", [40, 17, 17, 17, 40]):
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
    _put(ws, f"A{r}", "Functions you may need", font=H2)
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
    _put(ws, f"A{r + 1}", f"Cost = bottles × the ONE rate in B{r + 3}. Type a formula in "
                          f"C{r + 4} and fill it down. Without a $ sign the rows below show 0 "
                          "— that is the lesson.", wrap=True)
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
            ("workbook_id", wid),
            ("safety_stock_policy", scn.params.get("safety_stock_policy", "none")),
            ("worker_policy", scn.params.get("worker_policy", "whole"))]
    for i, (k, v) in enumerate(rows, start=1):
        _put(ws, f"A{i}", k)
        _put(ws, f"B{i}", v)
    ws.sheet_state = "veryHidden"
    _protect(ws)


# =========================================================================================
# Answer key
# =========================================================================================
ColumnSpec = Sequence[Tuple[str, str]]          # [(concept, heading)] in column order


def default_columns(kind: str) -> List[Tuple[str, str]]:
    order = L.REQUIRED[kind]
    if kind == "hybrid":
        order = ["Beginning Inventory", "Beginning Backorders", "Workers", "Regular Capacity",
                 "Regular Production", "Overtime Production", "Subcontract Production",
                 "Total Production", "Hires", "Layoffs", "Ending Inventory",
                 "Ending Backorders", "Labor Cost", "Hiring Cost", "Layoff Cost",
                 "Overtime Cost", "Subcontract Cost", "Holding Cost", "Backorder Cost",
                 "Total Cost"]
    return [(c, L.heading_for(kind, c)) for c in order]


def _fill_plan(ws, scn: Scenario, kind: str, columns: ColumnSpec, hybrid: Optional[Decision]):
    col = {c: get_column_letter(L.FIRST_STUDENT_COL + j) for j, (c, _) in enumerate(columns)}
    for j, (_, heading) in enumerate(columns):
        ws.cell(row=1, column=L.FIRST_STUDENT_COL + j).value = heading
    cap = "$C$" + L.planning_ref(kind, "capacity")[1:]
    ws[L.planning_ref(kind, "capacity")] = "=BottlesPerHour*HoursPerDay*DaysPerMonth"
    if kind in ("chase", "level"):
        ws[L.planning_ref(kind, "target")] = "=EndTarget"
    if kind == "level":
        td, tg, lr = (L.planning_ref(kind, k) for k in ("total_demand", "target", "level_rate"))
        ws[td] = "=SUM(B2:B13)"
        ws[lr] = f"=({td}+{tg}-BeginInv+BeginBackorders)/12"
        need = f"{lr}/{L.planning_ref(kind, 'capacity')}"
        ws[L.planning_ref(kind, "level_workers")] = (f"=ROUNDUP({need},0)" if scn.whole_workers
                                                     else f"={need}")
    if kind == "hybrid":
        ws[L.planning_ref(kind, "ot_limit")] = f"={L.planning_ref(kind, 'capacity')}*OTLimitPct"
    absref = lambda ref: f"${ref[0]}${ref[1:]}"                          # noqa: E731
    prod = col.get("Total Production") if kind == "hybrid" else col["Regular Production"]
    for i in range(L.N_MONTHS):
        r, p = L.month_row(i), L.month_row(i) - 1
        g = lambda c: f"{col[c]}{r}"                                       # noqa: E731
        f: Dict[str, object] = {}
        f["Beginning Inventory"] = "=BeginInv" if i == 0 else f"={col['Ending Inventory']}{p}"
        f["Beginning Backorders"] = ("=BeginBackorders" if i == 0
                                     else f"={col['Ending Backorders']}{p}")
        wprev = "StartWorkers" if i == 0 else f"{col['Workers']}{p}"
        if kind == "chase":
            f["Regular Production"] = (f"=MAX(0,B{r}+{absref(L.planning_ref(kind, 'target'))}"
                                       f"-{g('Beginning Inventory')}+{g('Beginning Backorders')})")
            need = f"{g('Regular Production')}/{cap}"
            f["Workers"] = f"=ROUNDUP({need},0)" if scn.whole_workers else f"={need}"
        elif kind == "level":
            f["Regular Production"] = f"={absref(L.planning_ref(kind, 'level_rate'))}"
            f["Workers"] = f"={absref(L.planning_ref(kind, 'level_workers'))}"
        else:
            f["Workers"] = hybrid.workers[i]
            f["Regular Production"] = hybrid.regular[i]
            f["Overtime Production"] = hybrid.overtime[i] if hybrid.overtime else 0
            f["Subcontract Production"] = hybrid.subcontract[i] if hybrid.subcontract else 0
            f["Total Production"] = (f"={g('Regular Production')}+{g('Overtime Production')}"
                                     f"+{g('Subcontract Production')}")
            f["Overtime Cost"] = f"={g('Overtime Production')}*OTCost"
            f["Subcontract Cost"] = f"={g('Subcontract Production')}*SubCost"
        if "Regular Capacity" in col:
            f["Regular Capacity"] = f"={g('Workers')}*{cap}"
        f["Hires"] = f"=MAX(0,{g('Workers')}-{wprev})"
        f["Layoffs"] = f"=MAX(0,{wprev}-{g('Workers')})"
        f["Ending Inventory"] = (f"=MAX(0,{g('Beginning Inventory')}+{prod}{r}"
                                 f"-{g('Beginning Backorders')}-B{r})")
        f["Ending Backorders"] = (f"=MAX(0,B{r}+{g('Beginning Backorders')}"
                                  f"-{g('Beginning Inventory')}-{prod}{r})")
        f["Labor Cost"] = f"={g('Workers')}*LaborCost"
        f["Hiring Cost"] = f"={g('Hires')}*HireCost"
        f["Layoff Cost"] = f"={g('Layoffs')}*LayoffCost"
        f["Holding Cost"] = f"={g('Ending Inventory')}*HoldCost"
        f["Backorder Cost"] = f"={g('Ending Backorders')}*BackorderCost"
        costs = [c for c in L.COST_CONCEPTS if c != "Total Cost" and c in col]
        f["Total Cost"] = "=" + "+".join(g(c) for c in costs)
        for concept, v in f.items():
            if concept in col:
                ws[g(concept)].value = v
    for concept in col:
        if concept in L.COST_CONCEPTS or concept in L.TOTAL_OPTIONAL:
            x = col[concept]
            ws[f"{x}{L.TOTAL_ROW}"].value = f"=SUM({x}{L.FIRST_ROW}:{x}{L.LAST_ROW})"
    return col


def _fill_summary(wb, scn: Scenario, letters: Dict[str, Dict[str, str]]):
    ws = wb[L.SHEET_SUMMARY]
    for kind, col in letters.items():
        sh = _sheet_ref(L.PLAN_SHEETS[kind])
        letter = L.SUMMARY_PLAN_COL[kind]
        for row in L.SUMMARY_ROWS:
            if row.hybrid_only and kind != "hybrid":
                continue
            src = {"Total Hires": "Hires", "Total Layoffs": "Layoffs"}.get(row.metric, row.metric)
            if row.metric == "Peak Workforce":
                f = f"=MAX({sh}!{col['Workers']}2:{col['Workers']}13)"
            elif row.metric == "Peak Inventory":
                x = col["Ending Inventory"]
                f = f"=MAX({sh}!{x}2:{x}13)"
            elif row.metric == "Months With Backorders":
                x = col["Ending Backorders"]
                f = f'=COUNTIF({sh}!{x}2:{x}13,">0")'
            else:
                f = f"={sh}!{col[src]}{L.TOTAL_ROW}"
            ws[f"{letter}{L.summary_row(row.metric)}"] = f
    tc = L.summary_row("Total Cost")
    ws[L.LOWER_COST_CELL] = f'=IF(B{tc}<C{tc},"Chase","Level")'
    ws[L.SAVINGS_CELL] = f"=ABS(B{tc}-C{tc})"
    chase_cost = reference_plan(scn, "chase")["Total Cost"].sum()
    level_cost = reference_plan(scn, "level")["Total Cost"].sum()
    ws[L.RECOMMEND_CELL] = "Chase" if chase_cost <= level_cost else "Level"
    ws[L.JUSTIFY_CELL] = ("ANSWER KEY. The recommendation and its justification are the "
                          "student's own judgement and are graded by the instructor, not by "
                          "the checker. " * 3).strip()


# =========================================================================================
# Public API
# =========================================================================================
def build_workbook(scn: Scenario, *, student_name: str = "", student_id: str = "",
                   game: Optional[str] = None, solution: bool = False,
                   hybrid_decision: Optional[Decision] = None,
                   columns: Optional[Dict[str, ColumnSpec]] = None) -> bytes:
    pieces = _pieces(scn)
    wb = Workbook()
    _start_sheet(wb, scn, student_name, student_id)
    _scenario_sheet(wb, scn)
    for kind in pieces:
        _plan_sheet(wb, scn, kind)
    _summary_sheet(wb, scn)
    _help_sheet(wb)
    _meta_sheet(wb, scn, student_name, student_id, game)
    if solution:
        letters = {}
        for kind in pieces:
            if kind == "hybrid" and hybrid_decision is None:
                continue
            spec = (columns or {}).get(kind) or default_columns(kind)
            letters[kind] = _fill_plan(wb[L.PLAN_SHEETS[kind]], scn, kind, spec, hybrid_decision)
        _fill_summary(wb, scn, letters)
    wb.active = 0
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def file_name(student_name: str, student_id: str, solution: bool = False) -> str:
    base = "".join(ch for ch in (student_name or student_id or "student")
                   if ch.isalnum() or ch in " -_").strip().replace(" ", "_") or "student"
    return f"Aggregate_Anxiety_{base}{'_ANSWER_KEY' if solution else ''}.xlsx"
