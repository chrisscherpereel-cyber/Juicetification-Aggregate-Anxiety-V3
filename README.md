# Juicetification: Aggregate Anxiety — V3 (Excel edition)

A [Streamlit](https://streamlit.io) companion for learning **aggregate production planning in
Excel**. Students design and build the plans in their own spreadsheet — **choosing the columns
as well as the formulas** — and submit the work **in pieces**. The app checks each piece on its
own and gives more guidance with every upload, but never the answer.

You play the operations planner for **Juicetification Inc.**, a juice bottler with a summer
peak. Each student ID produces its own reproducible demand forecast, so answers can't be copied.

---

## The steps (always visible in the sidebar)

| Step | What the student does |
|------|-----------------------|
| **1 · Scenario & workbook** | Signs in, studies the demand forecast and given data, downloads a personal `.xlsx`. |
| **2 · Chase plan** | Designs the chase worksheet, uploads, fixes, re-uploads; submits when it is all correct. |
| **3 · Level plan** | Opens once the chase plan is submitted. Same loop. |
| **4 · Hybrid plan** | Opens after the level plan. The student chooses workers, regular production, overtime and subcontracting; the plan must be correct **and** respect every limit. |
| **5 · Compare & recommend** | The Summary sheet: each plan's yearly figures, the cheaper plan (`IF`), the savings (`ABS`), a recommendation and a written justification. |
| **6 · Final report** | A PDF listing every piece. A DRAFT is available at any time. |

Pieces open one at a time (🔒 until the previous one is submitted). Each submitted piece gets a
one-page **receipt PDF**, so an instructor can collect pieces on separate due dates. The student
keeps working in the same workbook throughout and uploads it at each step; only that step's
piece is checked.

## Students build the columns

On each plan sheet only **Month** (column A) and **Demand** (column B) are given. The student
types a heading in row 1 for every column the plan needs (peach cells, from column C on), the
formulas for January–December below, and yearly totals in row 14. Planning cells under the
grid hold the one-off calculations (capacity per worker, level rate, level workforce, …).

The app never lists the columns. It asks the questions a plan must answer each month —
*"How many bottles are on the shelf when the month starts?"*, *"How many orders are still
unfilled at the end of the month?"* — and the student turns each one into a column.

The checker reads the student's headings ([aggplan/recognize.py](aggplan/recognize.py)), so
**any order and any clear wording works**: "Starting Stock", "Beg Inv" and "Opening inventory"
are all Beginning Inventory; "Crew", "Fired", "Wages", "Shortage Penalty" are Workers, Layoffs,
Labor Cost and Backorder Cost. After each upload the student sees how each heading was read:

* a **missing** column is reported as the question the plan can't answer yet (more specific on
  later uploads, ending with a suggested heading);
* two columns that look the same are flagged as **duplicates**;
* an unclear heading whose numbers clearly match a column is **checked as that column**, with a
  request for a clearer heading;
* unrecognised extra columns are treated as **helper columns** and left alone.

## Help that grows with each attempt — never the answer

Every column, planning cell and Summary item has a four-step hint ladder
([aggplan/hints.py](aggplan/hints.py)):

| Step | Example — Ending Inventory |
|---|---|
| 1 · Think about it | *What is left on the shelf after this month's orders go out?* |
| 2 · The idea | Start with what you had, add what you made, ship the orders you owed and this month's demand; stock can never be negative. |
| 3 · In Excel | One formula with Beginning Inventory, Production, Beginning Backorders and Demand, wrapped in `MAX(0, …)`. |
| 4 · The rule | Ending Inventory = MAX(0, Beginning Inventory + Production − Beginning Backorders − Demand). |

Each upload in which an item is **still wrong** moves that item up one step, so help is
focused on what the student is stuck on. Cells that are merely **empty**, or wrong only because
an earlier cell is wrong (↪ "follows an earlier mistake"), never unlock more help — uploading an
unfinished workbook does not reveal the method. From step 2, the student also sees a specific
observation about their own cell when one applies: a reference that slid because it lacks `$`
("lock it in January's formula, probably `$C$20`"), a fraction of a worker, a negative stock,
January's special case. **No hint ever contains a number from the scenario** (tested).

Re-uploading the same file is not a new attempt. The final report records, per piece, the number
of uploads, the first-attempt score and the highest hint step used — evidence of process for the
instructor, not a grade.

### Cell statuses

✔ correct (and calculated by a formula) · ✖ not right yet · ↪ follows an earlier mistake ·
✎ right number but typed in · ○ empty · ⚠ Excel error / text / not calculated ·
◆ hybrid decision (the student's choice)

The results show a "What to work on" list (root causes only), the column check, and a grid laid
out like the student's own sheet.

## Model rules

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
Hybrid: Workers, Regular, Overtime and Subcontract production are decisions;
        Total Production = Regular + Overtime + Subcontract
```

These rules are in this README for instructors; students reach them only through the hint
ladder. `EndTarget` is 0, or the safety stock when `safety_stock_policy=maintain`.

## How checking works

The checker never evaluates a formula. It reads the file twice — the formulas the student typed
and the values **Excel calculated and saved** — and compares each value with a Python answer key
([aggplan/engine.py](aggplan/engine.py)); for the hybrid, the key is built from the student's
own decisions and the plan is also tested against every limit (capacity, overtime, subcontract,
storage, on-time fulfillment, December backorders and inventory). It refuses, with a message the
student can act on, a file that isn't this app's workbook, belongs to another student ID, was made
for another scenario or configuration, has altered demand, or isn't an `.xlsx`.

Excel (Windows, Mac, web) is verified. Google Sheets works via *File → Download → Microsoft Excel
(.xlsx)*. A file never saved by a spreadsheet program has no calculated values; the app says so.

## Instructor configuration (optional)

Parameters live in [manifest.py](manifest.py) and can be set from the Juicetification Director
form or by URL: `?cfg=<base64 json>`, `?game=<code>`, `?seed=<int>`, `?manifest=1`.

| Parameter | Default | Meaning |
|---|---|---|
| demand, costs, capacity, inventory | as V2 | the planning data |
| `safety_stock_policy` | `none` | `maintain` makes `EndTarget` = safety stock |
| `scenario_preset` | `standard` | `strong_seasonal`, `stability_favored`, `limited_flex`, `tight_storage` |
| `include_hybrid` | true | include the Hybrid Plan piece |
| `require_formulas` | true | a typed number in a calculated cell is not accepted |
| `min_justification_words` | 50 | minimum length of the justification |
| `service_requirement_pct`, `terminal_backlog_max` | 95, 0 | hybrid feasibility limits |
| `assignment_name`, `report_timezone` | … | shown in the workbook and PDFs |

`include_hybrid`, `require_formulas`, `min_justification_words`, `assignment_name` and
`report_timezone` don't change any correct answer and can be changed mid-course; anything else
invalidates workbooks already handed out, by design.

**Answer keys:** `python tools/make_answer_key.py --sid A1234567 --game OPS301-F26` writes that
student's workbook filled with headings and reference formulas (open it in Excel to calculate).

**Identity and records** (unchanged from V2): signed links via `ASSIGNMENT_SIGNING_KEY` and
`tools/make_assignment_links.py`. With the Dropbox store configured (see
[student_store.py](student_store.py)) the app saves each student's pieces, attempts and hint
levels after every upload — so a student can close the browser and resume — and writes a
completion record (report ID, per-piece uploads and help, workbook SHA-256s, plan costs) with the
final report. Without a store everything works within a browser session.

## Run it

Python 3.9+.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tests: `pip install -r requirements-dev.txt && python -m pytest tests -q`. The checker tests run
against workbooks **calculated by real Excel** in `tests/fixtures` — the answer key, a workbook
with typical mistakes, and one whose chase columns are in a student's own order and wording.
After changing the workbook layout or the model, regenerate them on Windows with Excel:
`python tests/make_fixtures.py`.

## Architecture

```
app.py                 the six steps (Streamlit): pieces, uploads, hint levels, receipts
aggplan/               no Streamlit import; unit-testable
  scenario.py          Scenario: parameters, demand, Excel names
  engine.py            answer key: simulate(), chase/level reference plans, feasibility
  layout.py            fixed cells: given columns, totals row, planning cells, Summary
  recognize.py         reads the student's headings -> columns
  checker.py           grades one piece of an uploaded workbook
  hints.py             the four-step hint ladders and the column questions
  text.py              scenario text and Excel help
  workbook.py          builds the student workbook (and the answer key)
  report.py            piece receipts and the final report (PDF)
  identity.py          signed assignment links (from V2)
student_store.py, juice_director.py   shared Director modules (unchanged)
tools/                 make_answer_key.py, make_assignment_links.py, recalc_in_excel.ps1
tests/                 pytest suite + Excel-calculated fixtures
```

## Known limitations

* Without the Dropbox store, attempt counts and submitted pieces live in the browser session; a
  refresh keeps the student ID but clears them (the work itself is safe in the workbook, and a
  correct piece re-submits on its first upload).
* Heading recognition covers common wording; an unusual heading is matched by its numbers when
  they are right, otherwise the student is asked for a clearer heading.
* Two errors that cancel out can pass a cell; the checker sees values, and only "formula or typed".
* LibreOffice and Apple Numbers exports are not verified.

## License

MIT — see [LICENSE](LICENSE).
