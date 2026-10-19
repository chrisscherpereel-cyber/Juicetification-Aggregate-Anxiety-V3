# Juicetification: Aggregate Anxiety — V3 (Excel edition)

A [Streamlit](https://streamlit.io) companion for learning **aggregate production planning in
Excel**. Every calculation happens in the student's own spreadsheet; the app hands out a
personal workbook, checks the completed file cell by cell, and produces a PDF check report.

You play the operations planner for **Juicetification Inc.**, a juice bottler with a summer
peak. Each student ID produces its own reproducible demand forecast, so answers can't be
copied — the method is the same, the numbers are yours.

> V3 replaces V2's in-app worksheets, drills and multiple-choice stages with one loop:
> **download → build in Excel → upload → fix → upload again**. The planning model, the
> instructor parameters and the Director integration carry over from V2.

---

## The four steps (always visible in the sidebar)

| Step | What the student does |
|------|-----------------------|
| **1 · Meet your scenario** | Enters name and student ID; sees the demand forecast, a chart and the given data. |
| **2 · Build it in Excel** | Downloads a personal `.xlsx` and follows the sheet-by-sheet guide (also inside the workbook). |
| **3 · Upload & check** | Uploads the saved workbook. Gets a progress card per sheet, a short **Fix these next** list and a colour-coded grid laid out like the sheet (hover a cell for its feedback). Re-uploads as often as needed. |
| **4 · Submit** | Once every required cell is correct, downloads a PDF check report and submits **both** the workbook and the PDF to the LMS. A DRAFT report is available any time. |

Back/Next buttons at the bottom of every step name their destination; the sidebar shows ✔ for
finished steps and the last check's progress. A returning student can go straight to
**Upload & check**: the workbook identifies its owner, so no re-entry is needed.

## The workbook

| Sheet | Contents |
|-------|----------|
| **Start Here** | Name/ID, the six steps, colour key, and the rule for every column (chase and level). |
| **Scenario** | Demand forecast and every given number, each with an Excel **name** (`LaborCost`, `HireCost`, `BeginInv`, `EndTarget`, `BottlesPerHour`, …). Locked. |
| **Chase Plan** | Row 1 headings, January = row 2 … December = row 13, row 14 totals. Planning cells S2–S3 (capacity per worker, ending-inventory target). |
| **Level Plan** | Same grid. Planning cells S2–S6: capacity, target, total demand, level production, level workforce. |
| **Hybrid Plan** *(optional)* | Green decision cells (workers, regular, overtime, subcontract); everything else calculated. Checked for correctness **and** feasibility. |
| **Summary** | Links to each plan's totals, MAX/COUNTIF measures, `IF` for the cheaper plan, `ABS` for the savings, a recommendation drop-down and a justification box. |
| **Excel Help** | Habits (formulas, fill down, `$` references, names), the functions used, what each error means, and a small `$` practice block. |

Colour key: yellow = given, blue = your formula, green = your decision, darker blue = your
`SUM` totals. Column-heading comments repeat each rule. Sheets are protected **without a
password** (to prevent accidental edits of given data); every cell a student fills is unlocked.
A demand-vs-production chart fills in as the student works, backorders turn red and
inventory over storage capacity turns amber.

### Model rules (identical in the workbook, the app and the checker)

```
Beginning Inventory  = last month's Ending Inventory     (January = BeginInv)
Beginning Backorders = last month's Ending Backorders    (January = 0)
Ending Inventory     = MAX(0, Beg Inv + Production − Beg Backorders − Demand)
Ending Backorders    = MAX(0, Demand + Beg Backorders − Beg Inv − Production)
Hires   = MAX(0, Workers − last month's Workers)          (January vs StartWorkers)
Layoffs = MAX(0, last month's Workers − Workers)
Costs   = quantity × rate;  Total Cost = sum of the month's costs
Chase:  Production = MAX(0, Demand + EndTarget − Beg Inv + Beg Backorders)
        Workers    = ROUNDUP(Production ÷ capacity per worker, 0)
Level:  Production = (total demand + EndTarget − BeginInv) ÷ 12, every month
        Workers    = ROUNDUP(level production ÷ capacity per worker, 0), every month
```

`EndTarget` is 0, or the safety stock when the instructor sets `safety_stock_policy=maintain`.
Workers are whole and paid for capacity whether or not it is used.

## How checking works

The checker never evaluates a formula. It reads the file twice — once for the formulas the
student typed and once for the values **Excel calculated and saved** — and compares each value
with a Python answer key ([aggplan/engine.py](aggplan/engine.py)). Every checked cell gets one
status:

| | Status | Meaning |
|---|---|---|
| ✔ | Correct | right value, calculated by a formula |
| ✖ | Not right yet | wrong value, with the column's rule and, where it applies, a targeted hint |
| ↪ | Follows an earlier mistake | the cell's formula follows the rule given the student's *own* upstream cells; fix the root cause and it corrects itself |
| ✎ | Typed, not calculated | right number, typed in (only when `require_formulas` is on) |
| ○ | Empty | still to do |
| ⚠ | Excel error / text / not calculated | `#DIV/0!`, `#NAME?` …; words in a number cell; or a file never saved by a spreadsheet program |

The ↪ status is what keeps feedback short: one slip in February no longer turns the rest of
the sheet red. **Fix these next** lists root causes only, grouped per column ("and 10 more like
it in this column"), and counts the ↪ cells separately. Targeted hints cover the common Excel
mistakes: a planning-cell reference without `$` that slid when filled down ("Lock it: `$S$6`"),
workers not rounded up, negative inventory without `MAX(0, …)`, and January's special cases.

The checker also refuses — with a message the student can act on — a file that isn't this
app's workbook, belongs to another student ID, was made for another scenario or an older
configuration, has altered demand, or isn't an `.xlsx` (size and zip-bomb limits apply).

A workbook is **complete** when every cell on Chase Plan, Level Plan and Summary is ✔, a
recommendation is chosen and the justification has at least `min_justification_words` words.
The hybrid plan is optional unless the student recommends it.

Spreadsheet programs: Excel (Windows, Mac, web) is verified. Google Sheets works via
*File → Download → Microsoft Excel (.xlsx)*, which stores calculated values. A file that was
never opened and saved by a spreadsheet program has formulas but no values; the checker says so.

## Instructor configuration (optional)

Standalone practice needs no setup. Parameters live in [manifest.py](manifest.py) and can be set
from the Juicetification Director form, or by URL: `?cfg=<base64 json>`, `?game=<code>`,
`?seed=<int>` (one scenario for everyone), `?manifest=1` (emit the schema).

| Parameter | Default | Meaning |
|---|---|---|
| demand, costs, capacity, inventory | as V2 | the planning data (see manifest.py) |
| `safety_stock_policy` | `none` | `maintain` makes `EndTarget` = safety stock for both plans |
| `scenario_preset` | `standard` | `strong_seasonal`, `stability_favored`, `limited_flex`, `tight_storage` |
| `include_hybrid` | true | include the optional Hybrid Plan sheet |
| `require_formulas` | true | a typed number in a calculated cell is not accepted |
| `min_justification_words` | 50 | minimum length of the written justification |
| `service_requirement_pct`, `terminal_backlog_max` | 95, 0 | hybrid feasibility limits |
| `assignment_name`, `report_timezone` | … | shown in the workbook and the PDF |

`include_hybrid`, `require_formulas`, `min_justification_words`, `assignment_name` and
`report_timezone` do not change any correct answer, so they can be changed mid-course without
invalidating workbooks already handed out. Changing anything else (costs, demand, capacity,
policy) invalidates them by design: students then download a fresh workbook.

### Answer keys

```bash
python tools/make_answer_key.py --sid A1234567 --game OPS301-F26
```

writes the student's workbook layout filled with reference formulas (open it in Excel and
everything calculates). Use `--seed` for a scenario number and `--cfg cfg.json` for a non-default
configuration.

### Identity, signed links and completion records

Unchanged from V2. A typed student ID is a label; for graded use set `ASSIGNMENT_SIGNING_KEY`
and mint per-student links with `tools/make_assignment_links.py`. With the Dropbox store
configured (`DB_ENCRYPTION_KEY` + Dropbox credentials, see [student_store.py](student_store.py)),
the app records progress after each check and a completion record (attempt ID, workbook SHA-256,
seed, recommendation, plan costs) when a final report is created, so the Director roster fills
itself. Without a store everything still works and nothing is recorded.

The PDF report prints the workbook's SHA-256, so an instructor can confirm the submitted
workbook is the one that was checked. Plan cost is a result, not a grade; the recommendation
and justification are for the instructor's rubric.

## Run it

Python 3.9+.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tests: `pip install -r requirements-dev.txt && python -m pytest tests -q`.

The checker tests run against workbooks **calculated by real Excel** (`tests/fixtures`). After
changing the workbook layout or the model, regenerate them on Windows with Excel installed:

```bash
python tests/make_fixtures.py
```

## Architecture

```
app.py                 the four steps (Streamlit)
aggplan/               no Streamlit import; unit-testable
  scenario.py          Scenario: parameters, demand, Excel names
  engine.py            answer key: simulate(), chase/level reference plans, feasibility
  layout.py            where every cell lives (shared by builder and checker)
  text.py              rules and help text, generated from the active scenario
  workbook.py          builds the student workbook (and the answer key)
  checker.py           reads an uploaded workbook and grades every cell
  report.py            PDF check report
  identity.py          signed assignment links (from V2)
student_store.py, juice_director.py   shared Director modules (unchanged)
tools/                 make_answer_key.py, make_assignment_links.py, recalc_in_excel.ps1
tests/                 pytest suite + Excel-calculated fixtures
```

## Known limitations

* A browser refresh keeps the student ID (it is in the URL) but clears the last check; the
  workbook itself is the saved work, so re-uploading restores everything.
* The checker can't see *how* a value was computed beyond "formula or typed", so an exotic but
  correct formula passes, and two errors that cancel out can pass a cell.
* LibreOffice and Apple Numbers exports are not verified.

## License

MIT — see [LICENSE](LICENSE).
