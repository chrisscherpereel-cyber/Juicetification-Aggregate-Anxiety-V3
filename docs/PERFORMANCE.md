# Can one Streamlit Community Cloud app serve a class of 30+? — capacity analysis

**Short answer: yes, with room to spare at a normal class pace — after the changes below.**
Before them, the same test produced 27–35-second freezes for some students whenever
uploads and page loads coincided.

Everything here can be re-run: `python tools/bench.py` (component costs) and
`python tools/load_test.py --students 30` (a real Streamlit server with 30 simulated
browsers).

## Second pass (November 2026): 30–40 students, with uploads

Measured on a 2-vCPU Linux container, Python 3.13, Streamlit 1.67, with a real Streamlit
server and N simulated browsers over websockets (`tools/load_test.py`). Before = `main` at
05a3216; after = this branch. Same harness, same machine, back to back.

### What changed in the test itself (so the numbers mean something)

* **Real uploads.** `--upload-every k` makes every k-th action an upload exactly as a browser
  does it: ask for an upload URL over the websocket, PUT the file over HTTP, rerun with the
  uploader's state. Each upload is a workbook the server has never seen (no cache hits).
  Transfer and check-and-show time are reported separately.
* **Realistic students.** Simulated students now have a downloaded and uploaded workbook, as a
  real student mid-course does (that turns on the sidebar's "My latest workbook" and
  "Progress file" work, which the old test never exercised).
* **Errors are counted.** The old test only timed page loads, so a page that crashed looked
  fast. Every run below: **0 errors**.
* **No harness overhead.** The test entry script recompiled all of `app.py` on every click
  (~15 ms) — a cost a deployed app never pays (Streamlit caches the main script). Fixed;
  both columns below use the fixed harness.
* `bench.py` section 7 (AppTest sessions on a thread pool) no longer works with newer
  Streamlit (AppTest isn't thread-safe), so it was retired in favour of the real-server test.

### Results (measured)

| Scenario | | Before | After |
|---|---|---|---|
| **30 students, normal pace** (a click every ~10 s each) | median / p95 / max click | 81 / 189 / 530 ms | **40 / 95 / 271 ms** |
| | server CPU per click | 78 ms | **38 ms** |
| **40 students, normal pace** | median / p95 / max click | 81 / 188 / 340 ms | **44 / 94 / 214 ms** |
| | server load | 23% of one core | **11%** |
| **30 students, stress** (a click every ~2 s each) | median / p95 / max click | 292 / 1,665 / 2,387 ms | **39 / 223 / 630 ms** |
| **40 students, stress** | median / p95 / max click | 1,115 / 2,489 / 3,012 ms | **62 / 770 / 1,673 ms** |
| | server load | 74% of one core | **49%** |
| **30 students, uploading** (every 3rd action an upload, ~3 s apart) | upload check & show, median / p95 / max | 1,009 / 3,143 / 3,305 ms | **192 / 632 / 755 ms** |
| | slowest click | 3,864 ms | **816 ms** |
| **40 students, uploading** | upload check & show, median / p95 / max | 2,279 / 4,528 / 4,722 ms | **327 / 691 / 830 ms** |
| | slowest click | 6,534 ms | **1,036 ms** |
| | server load | 84% of one core | **62%** |
| Memory, 40 sessions open | peak | 252 MB | **241 MB** |
| Upload transfer (browser → server, 41 KB workbook) | median | 64–382 ms (queued) | **10–18 ms** |

At 1–20 students (measured with the earlier harness) clicks were already fast and flat; the
gains above are where the class is large or busy.

### What was changed, by impact

1. **Closed sections are no longer sent on every click** (Priority 2). Streamlit re-sends every
   element of the page on every click, including what sits inside a closed expander. About half
   of a step's elements were in sections students rarely open — the finished "Find your
   columns" panel (~48 elements), "Found so far" (~29), policies, the Excel-macro help, Excel
   tips, hints so far. These now use tracked expanders (`on_change="rerun"`, `.open`): closed,
   they cost one element; opening one reruns the page with its contents. The level step went
   from 133 elements to 74. This is most of the CPU saving — Streamlit's own work packing and
   sending elements was about two-thirds of each click.
2. **Sidebar downloads are built only when clicked** (Priority 2). "Progress file" JSON-encoded
   and signed the whole saved state (with the base64 workbook inside) on *every* click
   anywhere, and "My latest workbook" decoded and re-hashed the workbook every time; Streamlit
   then hashed and stored both. They now pass a function to `st.download_button` (deferred
   generation), given a private copy of the state taken at render time so no rerun can race it.
3. **Download clicks don't rerun the page** (`on_click="ignore"`) where nothing needs to change:
   receipts, the report PDF, the sidebar files, the saved-workbook download. The workbook
   download that records "downloaded" still reruns, on purpose.
4. **The results grid is built once per check, without pandas** (Priority 3). It did ~180
   DataFrame cell lookups on every click of a checked step; the HTML is now kept on the
   (immutable) check result.
5. **Garbage-collector tuning, once per process** (Priority 3). Checking an upload creates
   hundreds of thousands of short-lived objects; automatic collections took ~15% of a check
   (5.6 of 35.6 ms) re-scanning long-lived objects. `gc.freeze()` after start-up plus a larger
   youngest generation: 3.2 ms (measured in-process, 80 checks).
6. **Deprecated Streamlit calls replaced** (Priority 1 for reliability, not speed).
   `st.components.v1.html` (scroll-to-top, scroll-to-reply) was announced for removal after
   2026-06-01 and `use_container_width` after 2025-12-31. With an unpinned Streamlit, a
   redeploy could have broken the app. Now `st.iframe` (falls back to the component on older
   versions) and `width="stretch"`. Verified in a browser: new steps open at the top, the
   heading box stays in view after a check.
7. **`requirements.txt`: `streamlit>=1.52,<2`** (was `>=1.46`) — deferred downloads need 1.52.
   Streamlit recommends pinning; the upper bound keeps a future major version from arriving
   unannounced.

### Community Cloud: what's code, configuration, platform

* **Code (fixed above):** per-click work in the app and in what it asks Streamlit to send.
* **Configuration (already in place):** no file watcher, no usage stats, `fastReruns`, 5 MB
  upload limit (`.streamlit/config.toml`).
* **Platform (can't be changed from code):** one Python process serves every student, so
  CPU-heavy work is serialised by the GIL — the useful budget is about one core. Streamlit's
  documented limits (approximate, "may change at any time without notice"): CPU 0.078–2
  cores, memory 690 MB–2.7 GB. Actual CPU speed and how much of a core the app gets at a
  given moment are **unknown**; apps sleep after 12 hours without traffic (the first visit
  then waits for a cold start — open the app a few minutes before class).

### Can it serve 30–40 students?

On this test machine: **yes**. At 40 students in a busy upload burst the server used 62% of
one core and the slowest click was ~1 s. A Cloud core is slower and shared, so expect these
latencies to be roughly 1.5–3× higher there (an **estimate** — no load was put on the deployed
app). That still keeps a normal class pace well under a second, and an "everyone upload now"
moment around 1–3 s for the last student. Memory (~240 MB peak for 40 sessions) is far below
the 690 MB minimum.

### Not measured / still open

* The deployed app on Community Cloud (no load test against production).
* The Dropbox progress store under load (no credentials here). Saves run in the background,
  but on Cloud they still cost CPU (encoding and encrypting the state with the workbook) once
  per upload, answer or setting change.
* Browser rendering time.
* **Possible next step, not done:** read each upload in one pass instead of two. openpyxl is
  opened twice (formulas, then saved values); a direct one-pass XML reader could roughly halve
  the ~25 ms parse. Not done because it means re-implementing openpyxl's handling of shared
  formulas and date-formatted cells on the path that grades every upload — a correctness risk
  for a gain that the measurements above no longer need.

### Re-running

```bash
python tools/load_test.py --students 40 --clicks 8 --think 10          # normal pace
python tools/load_test.py --students 40 --clicks 12 --think 2          # stress
python tools/load_test.py --students 40 --clicks 9 --think 3 --upload-every 3   # uploads
```

---

## First pass (earlier): the original analysis

## How Streamlit Cloud runs the app (what matters)

* **One Python process serves every student.** Each click reruns `app.py` in a thread of that
  process; CPU-heavy Python work is serialised by the GIL, so the useful budget is about
  **one core**, however many vCPUs the container has.
* Community Cloud's documented limits are roughly 0.078–2 vCPU and 690 MB–2.7 GB of memory.
  A Cloud core is also slower than a desktop core and is shared.
* So the question is not "is one click fast?" but "when 30 students click and upload in the
  same few seconds, how long does the last one wait?"

## What was measured

Test machine: Intel i9-14900K, Windows 11, Python 3.9.10, Streamlit 1.50, openpyxl 3.1.5.
Two kinds of test:

1. **Components** (`tools/bench.py`): CPU per action, measured in-process.
2. **Real server** (`tools/load_test.py`): starts `streamlit run`, opens N websocket sessions
   exactly as browsers do (Streamlit's own protocol), and has every student click repeatedly,
   cycling through the steps. Each simulated student starts mid-course: chase plan submitted
   (receipt on screen), level plan checked (results grid on screen) — the busiest pages. The
   server process's CPU and memory are sampled with psutil.

Paces: **realistic** = a click about every 10 s per student (most of the time is spent in
Excel); **stress** = a click about every 2 s per student, all 30 at once.

## Results, before → after (same machine, run back to back)

### Real server, 30 students

| | Realistic pace — before | after | Stress pace — before | after |
|---|---|---|---|---|
| Median click | 233 ms | **176 ms** | 2,502 ms | **340 ms** |
| 95th percentile click | **26,694 ms** | **535 ms** | **33,450 ms** | 2,449 ms |
| Slowest click | 31,590 ms | 963 ms | 35,097 ms | 2,975 ms |
| Server CPU | 43% of a core | **28%** | 81% | 68% |
| Server memory | 162 MB (peak 226) | **156 MB (peak 202)** | 161 MB | 154 MB |

The before-numbers' long tail is the real problem: a student who clicked while classmates
were uploading or opening their workbook waited half a minute.

### Components (CPU per action)

| Action | Before | After |
|---|---|---|
| Check an uploaded workbook (new file) | ~80 ms | **~23 ms** |
| Check the same file again (e.g. next step) | ~80 ms | **~3 ms** |
| Memory while checking | 10.7 MB peak | **2.5 MB peak** |
| 30 different uploads in the same second | 3.0 s until the last is done | **1.6 s** |
| Build a student's workbook | ~110 ms, every visit / refresh | ~95 ms once per student, then **cached** |
| Receipt PDF on a submitted step | ~20 ms on **every** click | once per submission |
| Guided starting plan | ~55 ms on every click while the file sat in the uploader | once per file |
| Saving progress (Dropbox store) | network wait inside the click | **background** |

## What was changed

1. **Uploads are read in openpyxl's read-only mode, and only the sheets a step needs**
   (`checker.read_book`). A full load builds every style object of every sheet, twice; the
   streaming reader is ~3–7× cheaper. A test checks the fast reader returns exactly what the
   full reader did on every fixture.
2. **Parsed uploads are cached per process** (bounded, 96 files) by content hash, so the same
   workbook uploaded at the next step — or still sitting in the uploader on a rerun — isn't
   parsed again.
3. **Each student's blank workbook is cached** (bounded, 256), so refreshing the page or
   opening a second tab doesn't rebuild it; the build itself is faster because identical cell
   styles are copied instead of re-hashed (~1,200 cells per workbook).
4. **Answer-key plans are cached** per scenario (callers get copies).
5. **Receipts are built once per submission**, not on every click of a submitted step (this also
   fixes a side effect: the receipt's report ID used to change on every click).
6. **The guided starting plan is built once per uploaded file.**
7. **Progress saves run in the background** (`aggplan/saving.py`): the click returns at once;
   saves for one student are sent in order and coalesced to the newest state.
8. **reportlab loads on first use**, not at every cold start.
9. **`.streamlit/config.toml`** for Cloud: no file watcher, no usage stats, fast reruns.

## What this means on Community Cloud

* **Realistic class (30 students working in Excel and uploading as they go):** 28% of one
  i9 core here. Even if a Cloud core is three times slower, that is under one core, and the
  95th-percentile click stays well under the ~1 s a person notices as lag on this machine.
  **Comfortable.**
* **Everyone clicks at once (e.g. "everybody upload now"):** 30 uploads in the same second
  finish in 1.6 s here; budget roughly 3–5 s on Cloud for the last student. Acceptable, and a
  one-off.
* **Memory:** ~155 MB with 30 sessions open (~200 MB peak), far below the 690 MB minimum.
* **Classes larger than ~60 at stress pace,** or a container throttled to its minimum CPU,
  would queue clicks. If you see that, ask Streamlit for more resources or split sections
  across two deployments of the app (each is its own process).

## What was not measured

* Streamlit Cloud itself — no load was put on a deployed app; Cloud CPU speed is estimated.
* The Dropbox store under load (no credentials here). Saves now run in the background, so
  network time no longer delays clicks, but Dropbox's own rate limits still apply: roughly one
  save per upload, answer or setting change per student.
* Browser-side rendering time (the page itself is light: tables and HTML, no heavy charts
  except the demand chart on step 1).
* Windows note: openpyxl writes each sheet through a temporary file while saving a workbook;
  that costs ~65 ms of the 95 ms build here (antivirus scans temp files) and much less on
  Cloud's Linux.

## Re-running

```bash
python tools/bench.py                          # components + 30-student in-process tests
python tools/load_test.py --students 30 --think 10   # real server, realistic pace
python tools/load_test.py --students 30 --think 2    # real server, stress pace
```
