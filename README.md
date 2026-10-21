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
| **1 · Scenario & workbook** | Signs in, studies the demand forecast and the planning policies, and downloads a personal `.xlsx` built for the policies set in the sidebar. |
| **2 · Chase plan** | Designs the chase worksheet, uploads, fixes, re-uploads; submits when it is all correct. |
| **3 · Level plan** | Opens once the chase plan is submitted. Same loop. |
| **4 · Hybrid plan** | Opens after the level plan. **Guided** (default): a starting plan and questions that lead to a good solution. **Build it yourself**: the student designs the sheet. Either way the plan must be correct **and** respect every limit. Skipped when the hybrid is switched **off**. |
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

Each column answers one question about the month — *"How many bottles are on the shelf when
the month starts?"*, *"How many orders are still unfilled at the end of the month?"* — and the
student turns each one into a column. Help with designing them
([aggplan/columns.py](aggplan/columns.py)), in the step's **① Design your columns** panel:

Nothing about a column is given away before the student finds it:

* **What's new** for this plan — a count, not a list: 14 questions for the chase plan; the
  level plan uses the same columns (only Production and Workers change); the hybrid adds 5 new
  questions (marked *new*) and makes four columns decisions.
* **Still to find** — the open questions, grouped in the order a month happens (start of the
  month → your plan → what happens → what it costs). Questions only, no column names.
* **Check one heading** — the student types one planned heading at a time:
  * found → ticked off, with the column's name, the question it answers and **why it matters**
    (and a note if, say, a bare "Inventory" will be read as the *end* of the month);
  * same question as a heading already found → said so; a helper column → said so; a column
    from another plan (e.g. overtime on the chase plan) → said so;
  * not recognised → help that grows with each miss: say what it measures → turn an open
    question's key words into a heading (and "try this one") → the kinds of words the
    checker understands.
  Found columns and misses are saved with the student's work; columns already read from an
  upload count as found.
  The panel reads top down, newest first: the latest check, *How they connect*, *Still to
  find*, then the earlier finds (most recent first) at the bottom. After each check the page scrolls
  so the heading box and its reply are in view.
* **Column order**, once every column is found — judged on the workbook's own order once it
  has been uploaded with every column, otherwise on the order the headings were checked. If a
  column comes before a column it uses (e.g. Total Cost before the costs it adds up), the
  panel names those, recommends an order (with column letters, the way a month happens) and
  says why it's better: each formula only looks left; a row reads like the month; mistakes
  flow one way, so fixing the leftmost one fixes the rest. Excel and the checker accept any
  order.
* **How they connect** — a diagram of which column feeds which; a column not found yet is a
  **?**, so the picture fills in as the student goes.

The workbook's **Start Here** sheet has the questions as a printable table (on its own page) —
question, *My heading* to fill in, and which plans need it — without the answers.

The checker reads the student's headings ([aggplan/recognize.py](aggplan/recognize.py)), so
**any order and any clear wording works**: "Starting Stock", "Beg Inv" and "Opening inventory"
are all Beginning Inventory; "Crew", "Fired", "Wages", "Shortage Penalty" are Workers, Layoffs,
Labor Cost and Backorder Cost. After each upload the student sees how each heading was read:

* a **missing** column is reported as the question the plan can't answer yet (more specific on
  later uploads, ending with a suggested heading);
* two columns that look the same are noted as **duplicates**: the first is checked, the others
  are shown as helper columns (they don't block submitting);
* a heading with *helper*, *original*, *scratch* or *note* in it is always a helper column;
* an unclear heading whose numbers clearly match a column is **checked as that column**, with a
  request for a clearer heading;
* unrecognised extra columns are treated as **helper columns** and left alone.

The plan sheets colour exactly enough columns for a complete plan (every required column plus
the optional Regular Capacity); the paler columns to the right are marked as spare room for
helper columns. The **Excel Help** sheet has a *Patterns* section — never below zero
(`MAX(0, …)`), whole workers (`ROUNDUP`), this month starts where last month ended, changes
between months, the stock balance, `COUNTIF`, `IF` — as shapes, not answers.

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
out like the student's own sheet: every column they built (helper columns shown as "not checked"),
the totals row and the planning cells. Wide sheets such as the hybrid switch to vertical headings
so all columns fit; hybrid columns wait (…) until the decisions are filled in.

## Settings (sidebar)

Three switches sit in the sidebar under **Settings**:

| Switch | Options (default first) |
|---|---|
| **Safety stock** | *None* (EndTarget = 0) · *Keep the safety stock* (EndTarget = the safety-stock quantity) |
| **Workforce** | *Whole workers* (round up; idle capacity is paid) · *Fractional workers* (capacity grows in small increments — part-time hours — with no rounding) |
| **Hybrid plan** | *Guided walkthrough* · *Build it yourself* · *Off* (chase and level only) |

The workbook, answer key, checker and hints follow the two policies. Those lock once the first
piece is submitted (every piece must follow the same rules); changing them earlier asks the
student to download the workbook again, and an upload made for other policies is refused with a
message naming the switch. When the policy is *no safety stock*, the safety-stock figure is not
shown anywhere.

The hybrid switch can change until the hybrid (or the comparison) is submitted. **Off** removes
the Hybrid step entirely: a newly downloaded workbook has no Hybrid Plan sheet or Hybrid column,
the Summary is checked for chase and level only (an existing Hybrid column is ignored), the
recommendation must be Chase or Level, and the final report leaves the hybrid out.

Instructors who want fixed policies set `student_sets_policies=false`; `hybrid_mode` sets the
hybrid default (`include_hybrid=false` removes the hybrid option altogether).

## Planning policies (stated to students)

The Scenario sheet, the Start Here sheet and step 1 of the app spell out the company's policies
as business rules, not formulas:

* **Safety stock** — either *none* (EndTarget = 0: the chase plan uses up the beginning stock
  and ends each month at 0; the level plan ends December at 0) or *keep S bottles*
  (EndTarget = S: every chase month aims to end with S; the level plan ends December with S;
  the hybrid must end December with at least S). It also states that beginning inventory is
  not the safety stock, that safety stock is a target customers can still draw on, and that it
  costs holding like any other bottle.
* **Backorders** — unfilled orders carry forward and are filled oldest-first; the year starts
  owing `BeginBackorders` (given data with its own Excel name; 0 by default); every bottle
  still owed at month end costs BackorderCost; stock is never negative.
* **Workforce** — either *whole workers only* (a need of 10.2 workers means employing 11; the
  idle capacity is still paid) or *fractional workers* (worker-equivalents, no rounding; costs
  in proportion), plus how hires and layoffs are counted from StartWorkers.
* **Overtime** costs OTCost for each bottle made on overtime — the full cost of that bottle;
  regular labor pays only for regular-time capacity.

## The hybrid plan

**Where to start — the cheaper of the two plans.** A hybrid improves a plan by fixing its most
expensive habit. Simulating every default scenario shows why the starting point matters: with
cheap hiring and layoffs ($600/$900) against $3,200 a month of labor, a steady-workforce hybrid
built from the level plan is *more* expensive than both pure plans, while the chase plan with
the rounded-up extra worker replaced by overtime beats both. When hiring and layoffs are
expensive (`stability_favored`), the level plan can be the cheaper one and is the right start.
Both modes recommend the student's cheaper plan, with their own cost breakdown as evidence.

**Guided walkthrough** ([aggplan/guide.py](aggplan/guide.py)):

1. *Pick the starting point* — the student chooses, from their own two plans (answer: the
   cheaper one).
2. *Get the starting plan* — the student uploads their workbook; the app fills the Hybrid Plan
   sheet with that plan (every column, formula and total built; Workers and Regular Production
   as green decisions; no overtime or subcontracting) and hands it back, with grey **helper
   columns** after a gap: the starting workers and regular production as fixed numbers and
   (whole workers) the leftover with one worker fewer. Because they never change, formulas can
   use them without a circular reference. Other sheets are untouched. The student opens it in
   Excel, saves, and uploads.
3. *Questions*, with numbers from the student's scenario, checked as they answer (a nudge after
   a wrong try; the answer can be revealed after three):
   * **chase track** — "In February you make 8,500 bottles with 9 workers… how much paid
     capacity sits idle?", "what would the missing bottles cost on overtime with 8 workers?",
     "which is cheaper?", "so when is overtime cheaper than one more worker?";
   * **level track** — which months end with backorders, how much paid capacity is idle, which
     extra source is cheapest per bottle, the overtime limit;
   * **no-gain track** (fractional workers) — the cost of a regular-time vs an overtime bottle,
     and why the chase plan is already the best mix.
4. *Make the change in Excel* — a numbered checklist with a worked example from the
   student's own numbers (e.g. February: Workers 9 → 8, Regular Production 8,500 → 8,000,
   Overtime 0 → 500), stressing the step that is easy to forget: a smaller crew makes less on
   regular time. It says to change only the green cells by typing numbers, and not to add a
   Leftover column inside the plan.
5. *Month-by-month coach* — after every upload, each month that breaks a limit is named with the
   exact cells to change, in the student's own numbers ("February: 8 workers can make only
   8,000 on regular time, but Regular Production is 8,500 — set it to 8,000 and Overtime to
   500"). Once the plan meets every limit, help on where money is still left grows with each
   upload: a nudge, then which months, then the change for each month. Tests replay a
   half-finished plan through the coach to a good solution.
6. *How good is it?* — each upload compares the hybrid with both plans. A **good solution** is
   feasible, at least as cheap as both plans, and (where an improvement exists) cheaper than the
   starting plan; then a short "why is it better?" question. **Submit appears only for a good
   solution**: an unchanged starting plan has every formula right ("Formulas correct 209 / 209")
   but is told to keep improving. Tests confirm that doing what the walkthrough says reaches a
   good solution in every scenario family.

**Build it yourself:** the student designs the Hybrid Plan sheet from scratch, with a "where to
start" panel (their cost breakdown, the recommended starting plan and why, steps, and what one
more bottle costs each way). The checker, feasibility tests and cell-by-cell grid are the same;
a hybrid that costs no less than the cheaper pure plan can be submitted, with a warning.

## Model rules

```
Beginning Inventory  = last month's Ending Inventory     (January = BeginInv)
Beginning Backorders = last month's Ending Backorders    (January = BeginBackorders)
Ending Inventory     = MAX(0, Beg Inv + Production − Beg Backorders − Demand)
Ending Backorders    = MAX(0, Demand + Beg Backorders − Beg Inv − Production)
Hires   = MAX(0, Workers − last month's Workers)          (January vs StartWorkers)
Layoffs = MAX(0, last month's Workers − Workers)
Costs   = quantity × rate;  Total Cost = sum of the month's costs
Chase:  Production = MAX(0, Demand + EndTarget − Beg Inv + Beg Backorders)
        Workers    = ROUNDUP(Production ÷ capacity per worker, 0)
Level:  Production = (total demand + EndTarget − BeginInv + BeginBackorders) ÷ 12, every month
        Workers    = ROUNDUP(level production ÷ capacity per worker, 0), every month
(worker_policy = partial: no ROUNDUP — fractional worker-equivalents)
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

Three checks look at the formulas themselves:

* **Circular references** on a plan sheet (cells whose formulas use each other in a loop — Excel
  then saves 0) are named, every cell in the loop is marked, and the piece can't be completed
  until the loop is broken. On the hybrid the message explains the usual cause: a formula in a
  green decision cell that uses a column calculated from it.
* **Wrong-column references.** Each column may use only what it depends on in the same month
  (and, for carried-over columns, the month before), e.g. Hires may use Workers or Regular
  Capacity but not Overtime. A formula that points elsewhere is marked wrong *even when its
  number happens to fit* (e.g. `=MAX(0,H2-E3)` while overtime is 0). Month, Demand and helper
  columns are always allowed. On the Summary, the lower-cost answer must be a formula that
  compares the two TOTAL PLAN COST cells (`=B3`, which shows the word "Chase", is not accepted).
* **Left-out inputs.** Each column's formula must use what it depends on — e.g. Ending
  Inventory needs Demand, Beginning Inventory, Beginning Backorders and production (on the
  hybrid, Total Production or all three of its parts). A formula that leaves one out can still
  give the right number when that term is 0. On the **chase plan**, where backorders are 0, that
  is a **notice** ("Works now, but…", ✔ with a small !) that doesn't block submitting; on the
  **level and hybrid plans**, where the term is not always 0, the cell is marked **wrong**.
  References that name the cell's own sheet (`'Level Plan'!H3`) count as same-sheet; a formula
  that goes through a helper column isn't judged.

The receipt and final report show each piece's first **and** last upload score and the score of
every upload in between.

Excel (Windows, Mac, web) is verified. Google Sheets works via *File → Download → Microsoft Excel
(.xlsx)*. A file never saved by a spreadsheet program has no calculated values; the app says so.

## Instructor configuration (optional)

Parameters live in [manifest.py](manifest.py) and can be set from the Juicetification Director
form or by URL: `?cfg=<base64 json>`, `?game=<code>`, `?seed=<int>`, `?manifest=1`.

| Parameter | Default | Meaning |
|---|---|---|
| demand, costs, capacity, inventory | as V2 | the planning data |
| `safety_stock_policy` | `none` | `maintain` makes `EndTarget` = safety stock |
| `beginning_backorders` | 0 | orders still owed before January (`BeginBackorders`) |
| `worker_policy` | `whole` | `partial` allows fractional workers (no rounding) |
| `student_sets_policies` | true | students set the two policies in the sidebar |
| `hybrid_mode` | `guided` | default hybrid mode (`guided`, `build` or `off`); students can switch |
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

## No work is lost

**Between screens:** everything stays in the session — uploads and their feedback, hint
levels, guided answers, settings and the step you're on.

**On another computer:** a student's saved work is everything needed to continue anywhere
([aggplan/persistence.py](aggplan/persistence.py)): name, policies, hybrid mode, every piece
(uploads, hint levels, guided answers, submissions), the step they were on, and **their
latest uploaded workbook**. Signing in with the same student ID brings it all back — the
feedback grids are rebuilt by re-checking that workbook — and the sidebar's
**⬇ My latest workbook** lets them carry on in Excel on the new computer.

* **With the server store configured** (`DB_ENCRYPTION_KEY` + Dropbox credentials): saved
  automatically after every upload, answer, setting change and step change, in the
  background, encrypted. The sidebar shows *Saving… / Saved at hh:mm / last save failed*.
* **Without it:** the sidebar says so plainly and offers **⬇ Progress file** (a JSON file with
  everything, including the workbook). Uploading it on step 1 of any computer restores the
  work. The file carries a checksum — and an HMAC signature when `ASSIGNMENT_SIGNING_KEY` is
  set — so a damaged or edited file, or another student's, is refused.

Safeguards: nothing is saved for a student until their saved work has been looked for and
loaded, so a fresh session on a second computer can't overwrite it. If saved work can't be
loaded (a network error) or was made for a different setup (the instructor changed the
assignment), saving pauses and the student is told, with a download of their saved workbook
and an explicit "Start fresh" choice. Saves for one student are sent in order, newest last.

## Running a class of 30+ on Streamlit Community Cloud

Verified with a real Streamlit server and 30 simulated browsers (`tools/load_test.py`); full
method and numbers in [docs/PERFORMANCE.md](docs/PERFORMANCE.md).

| 30 students, a click every ~10 s each | Before tuning | Now |
|---|---|---|
| Median click | 233 ms | 176 ms |
| 95% of clicks within | 26.7 s | 0.54 s |
| Server CPU | 43% of a core | 28% |
| Memory | 162 MB | 156 MB |

What keeps it fast: uploads are read in openpyxl's read-only mode and only for the sheets that
step needs (and cached by content), each student's workbook is built once, receipts once per
submission, the guided starting plan once per file, and progress saves run in the
background. 30 different uploads in the same second are all checked in ~1.6 s on the test
machine. `.streamlit/config.toml` turns off the file watcher and usage stats for Cloud.

If a class is much larger, or the Cloud container is throttled to its minimum CPU, clicks start
to queue: ask Streamlit for more resources, or deploy the app twice and split sections.

## Run it

Python 3.9+.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tests: `pip install -r requirements-dev.txt && python -m pytest tests -q`.
Capacity checks: `python tools/bench.py` and `python tools/load_test.py --students 30`. The checker tests run
against workbooks **calculated by real Excel** in `tests/fixtures` — the answer key, a workbook
with typical mistakes, one whose chase columns are in a student's own order and wording,
and one with every policy switched (safety stock kept, fractional workers, orders owed).
After changing the workbook layout or the model, regenerate them on Windows with Excel:
`python tests/make_fixtures.py`.

## Architecture

```
app.py                 the six steps (Streamlit): settings, pieces, uploads, hints, receipts
aggplan/               no Streamlit import; unit-testable
  scenario.py          Scenario: parameters, demand, Excel names
  engine.py            answer key: simulate(), chase/level reference plans, feasibility
  layout.py            fixed cells: given columns, totals row, planning cells, Summary
  recognize.py         reads the student's headings -> columns
  columns.py           one-heading check, open questions, what's new, diagram
  checker.py           grades one piece of an uploaded workbook
  hints.py             the four-step hint ladders and the column questions
  guide.py             the guided hybrid: starting point, questions, verdict
  text.py              scenario text and Excel help
  workbook.py          builds the student workbook (and the answer key)
  report.py            piece receipts and the final report (PDF)
  saving.py            background, per-student-ordered progress saves
  persistence.py       saved-work snapshot and the progress file
  identity.py          signed assignment links (from V2)
student_store.py, juice_director.py   shared Director modules (unchanged)
tools/                 make_answer_key.py, make_assignment_links.py, recalc_in_excel.ps1,
                       bench.py, load_test.py (+ load_app.py) for capacity checks
tests/                 pytest suite + Excel-calculated fixtures
```

## Known limitations

* Without the Dropbox store, work moves between computers only through the progress file the
  student downloads; a closed tab without one starts over (the workbook itself is safe on the
  student's computer).
* Heading recognition covers common wording; an unusual heading is matched by its numbers when
  they are right, otherwise the student is asked for a clearer heading.
* Two errors that cancel out can still pass a cell: the checker compares values and checks
  which columns a formula uses (and leaves out), but doesn't evaluate the formula with other
  inputs, so a wrong sign or operator on a 0 term isn't caught.
* LibreOffice and Apple Numbers exports are not verified.

## License

MIT — see [LICENSE](LICENSE).
