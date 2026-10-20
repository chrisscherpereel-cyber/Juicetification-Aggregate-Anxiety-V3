# Can one Streamlit Community Cloud app serve a class of 30+? — capacity analysis

**Short answer: yes, with room to spare at a normal class pace — after the changes below.**
Before them, the same test produced 27–35-second freezes for some students whenever
uploads and page loads coincided.

Everything here can be re-run: `python tools/bench.py` (component costs) and
`python tools/load_test.py --students 30` (a real Streamlit server with 30 simulated
browsers).

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
