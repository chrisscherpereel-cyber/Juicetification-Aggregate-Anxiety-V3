#!/usr/bin/env python
"""Capacity benchmark: can one Streamlit Community Cloud app serve a class of 30+?

  python tools/bench.py            # everything
  python tools/bench.py --quick    # skip the 30-student concurrency run

Measures, in this process:
  1. cold import time of the app's modules
  2. CPU cost of each expensive action (workbook build, checking each piece, PDF reports,
     the guided starting plan)
  3. CPU cost of an ordinary page rerun (a click) on every step, via Streamlit's AppTest
     (includes AppTest's own overhead, so it is an upper bound)
  4. memory one student's session holds
  5. 30 simulated students doing a whole course session at once on a thread pool —
     Streamlit runs every session in ONE process, so this is the shape of the real load
"""

from __future__ import annotations

import argparse
import gc
import os
import pickle
import statistics
import sys
import time
import tracemalloc
from concurrent.futures import ThreadPoolExecutor

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
FIX = os.path.join(ROOT, "tests", "fixtures")


def cpu_ms(fn, n=10):
    """Mean CPU milliseconds of fn() over n runs after one warm-up. (Windows' process clock
    ticks every 15.6 ms, so single runs must not be timed one by one.)"""
    fn()
    t = time.process_time()
    for _ in range(n):
        fn()
    return (time.process_time() - t) * 1000 / n


def fresh(fn, *a):
    """Call fn with the parsed-upload cache cleared, i.e. a never-seen file."""
    from aggplan import checker
    getattr(checker, "_BOOK_CACHE", {}).clear()
    return fn(*a)


def section(title):
    print(f"\n== {title} " + "=" * max(0, 70 - len(title)))


def imports():
    section("1. cold import")
    import subprocess
    code = ("import time;t=time.perf_counter();"
            "import aggplan.checker, aggplan.workbook, aggplan.guide, aggplan.text;"
            "import streamlit, altair;"
            "print(round((time.perf_counter()-t)*1000))")
    ms = [int(subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True,
                             text=True).stdout.strip()) for _ in range(3)]
    print(f"modules the app loads at start: {statistics.median(ms)} ms (median of 3 processes)")


def actions():
    section("2. expensive actions (CPU ms, median)")
    from aggplan import checker, engine, guide, report, workbook
    from aggplan.scenario import resolve_scenario
    s = resolve_scenario({}, 7)
    key = open(os.path.join(FIX, "key_seed7.xlsx"), "rb").read()
    rows = {
        "resolve_scenario": lambda: resolve_scenario({}, 7),
        "reference plans (chase+level)": lambda: (engine.reference_plan(s, "chase"),
                                                  engine.reference_plan(s, "level")),
        "build student workbook": lambda: workbook.build_workbook(s, student_name="A", student_id="S1"),
        "read uploaded workbook (all sheets)": lambda: fresh(checker.read_book, key),
    }
    for piece in ("chase", "level", "hybrid", "summary"):
        rows[f"check upload: {piece}"] = (lambda p=piece: fresh(checker.check_piece, key, s, p, "S1"))
    rows["check again (same file, cached)"] = lambda: checker.check_piece(key, s, "chase", "S1")
    res = checker.check_piece(key, s, "summary", "S1")
    pieces = [{"title": "Chase plan", "status": "Submitted", "attempts": 2, "first": "1 / 2",
               "submitted_at": "now", "help": 1, "file": "f", "sha": "a" * 64}]
    rows["PDF receipt / report"] = lambda: report.build_report(
        s, student_name="A", student_id="S1", identity_label="x", pieces=pieces, final=True,
        heading="Receipt", comparison=res.metrics, recommendation="Chase",
        justification="Because.", words=1)
    rows["guided: add starting plan"] = lambda: workbook.add_hybrid_start(
        key, s, guide.starting_decision(s, "chase"))
    rows["guided: plan costs + steps"] = lambda: (guide.plan_costs(s), guide.steps(s, "chase"))
    out = {}
    for name, fn in rows.items():
        out[name] = cpu_ms(fn)
        print(f"{name:34s} {out[name]:8.1f}")
    return out


# Runs app.py and records the CPU time of the script's own thread — what the server spends
# on a click (building and serialising the page), without the test harness's own work.
_TIMED_APP = f'''
import os, sys, time, streamlit as st
sys.path.insert(0, {ROOT!r})
os.chdir({ROOT!r})
_t = time.thread_time()
try:
    exec(compile(open("app.py", encoding="utf-8").read(), "app.py", "exec"), globals())
finally:
    st.session_state["_bench_ms"] = (time.thread_time() - _t) * 1000
'''


def reruns():
    section("3. page reruns (CPU ms per click, the app's own thread)")
    from streamlit.testing.v1 import AppTest
    from aggplan import checker
    from aggplan.scenario import resolve_scenario
    at = AppTest.from_string(_TIMED_APP, default_timeout=60).run()
    at.text_input[0].input("Ann")
    at.text_input[1].input("S1")
    next(b for b in at.button if b.label == "Start").click()
    at = at.run()
    s = resolve_scenario({}, store_seed("S1"))
    key = open(os.path.join(FIX, "key_seed7.xlsx"), "rb").read()
    # a mid-course student: chase submitted (receipt on screen), level checked with
    # results on screen, hybrid not started
    at.session_state["pieces"]["chase"]["submitted"] = {
        "at": "now", "sha": "x", "file": "f", "metrics": {}, "feasibility": []}
    at.session_state["pieces"]["level"]["attempts"] = [{
        "at": "now", "sha": "y", "file": "f", "ok": 1, "total": 2, "complete": False,
        "fatal": False}]
    at.session_state["results"]["level"] = checker.check_piece(key, resolve_scenario({}, 7),
                                                               "level", None)
    out = {}
    for step in ("start", "chase", "level", "report"):
        at.session_state["step"] = step
        at.run()
        ms = []
        for _ in range(24):                     # the clock ticks in 15.6 ms steps on Windows,
            at.run()                            # so average many clicks
            ms.append(at.session_state["_bench_ms"])
        out[step] = statistics.mean(ms)
        print(f"step {step:10s} {out[step]:8.1f}")
    del s
    return out, at


def store_seed(sid):
    import student_store
    return student_store.derive_seed(None, sid, lo=1, hi=10 ** 6)


def memory(at):
    section("4. memory per student session")
    state = {k: at.session_state[k] for k in ("pieces", "results", "wb_bytes", "student",
                                              "policy") if k in at.session_state}
    from aggplan import checker
    from aggplan.scenario import resolve_scenario
    s = resolve_scenario({}, store_seed("S1"))
    data = open(os.path.join(FIX, "key_seed7.xlsx"), "rb").read()
    tracemalloc.start()
    results = {p: checker.check_piece(data, resolve_scenario({}, 7), p, "S1")
               for p in ("chase", "level", "hybrid", "summary")}
    cur, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    blob = len(pickle.dumps(results))
    print(f"4 check results kept in a session: ~{blob / 1024:.0f} KB pickled, "
          f"{cur / 1024 / 1024:.1f} MB live, peak while checking {peak / 1024 / 1024:.1f} MB")
    print(f"workbook bytes kept: {len(state.get('wb_bytes', b'')) / 1024:.0f} KB")
    return blob, peak


def concurrency(n_students=30):
    section(f"5. {n_students} students at once (one process, thread pool)")
    from aggplan import checker, report, workbook
    from aggplan.scenario import resolve_scenario
    key = open(os.path.join(FIX, "key_seed7.xlsx"), "rb").read()

    def one_student(i):
        """Everything heavy one student does in a whole session: download the workbook,
        upload 2 attempts per piece, get 4 receipts and a final report."""
        s = resolve_scenario({}, 7)
        t0 = time.perf_counter()
        workbook.build_workbook(s, student_name=f"S{i}", student_id="S1")
        for p in ("chase", "level", "hybrid", "summary"):
            for _ in range(2):                              # two different files
                r = fresh(checker.check_piece, key, s, p, "S1")
        pieces = [{"title": p, "status": "Submitted", "attempts": 2, "first": "",
                   "submitted_at": "now", "help": 0, "file": "f", "sha": "a" * 64}
                  for p in ("chase", "level", "hybrid", "summary")]
        for _ in range(5):
            report.build_report(s, student_name="A", student_id="S1", identity_label="x",
                                pieces=pieces, final=True, heading="R", comparison=r.metrics)
        return time.perf_counter() - t0

    one_student(0)                                          # warm caches / imports
    gc.collect()
    t_cpu, t_wall = time.process_time(), time.perf_counter()
    with ThreadPoolExecutor(max_workers=n_students) as ex:
        lat = list(ex.map(one_student, range(n_students)))
    wall, cpu = time.perf_counter() - t_wall, time.process_time() - t_cpu
    print(f"{n_students} full sessions of heavy work: {wall:.1f} s wall, {cpu:.1f} s CPU "
          f"({cpu / n_students * 1000:.0f} ms CPU per student)")
    print(f"worst student waited {max(lat):.1f} s for all of it at once")
    return wall, cpu


def burst(n=30):
    section(f"6. burst: {n} different workbooks uploaded in the same second")
    from aggplan import checker
    from aggplan.scenario import resolve_scenario
    key = open(os.path.join(FIX, "key_seed7.xlsx"), "rb").read()
    s = resolve_scenario({}, 7)
    size = getattr(checker, "_BOOK_CACHE_SIZE", None)
    if size is not None:
        checker._BOOK_CACHE_SIZE = 0                  # every file is new: no cache hits
    try:
        def one(_):
            t = time.perf_counter()
            checker.check_piece(key, s, "chase", "S1")
            return time.perf_counter() - t
        with ThreadPoolExecutor(max_workers=n) as ex:
            t = time.perf_counter()
            lat = sorted(ex.map(one, range(n)))
            wall = time.perf_counter() - t
    finally:
        if size is not None:
            checker._BOOK_CACHE_SIZE = size
    print(f"all {n} checked in {wall:.2f} s; median wait {statistics.median(lat):.2f} s, "
          f"slowest {lat[-1]:.2f} s")
    return wall, lat


def sessions(n=30, clicks=12):
    section(f"7. {n} live app sessions clicking at once (real script runner, one process)")
    from streamlit.testing.v1 import AppTest

    def student(i):
        at = AppTest.from_string(_TIMED_APP, default_timeout=120).run()
        at.text_input[0].input(f"Student {i}")
        at.text_input[1].input(f"S{i}")
        next(b for b in at.button if b.label == "Start").click()
        at = at.run()
        lat, cpu = [], []
        steps = ["chase", "start", "chase", "report"]
        for k in range(clicks):
            at.session_state["step"] = steps[k % len(steps)]
            t = time.perf_counter()
            at.run()
            lat.append(time.perf_counter() - t)
            cpu.append(at.session_state["_bench_ms"])
        return lat, cpu

    t = time.perf_counter()
    with ThreadPoolExecutor(max_workers=n) as ex:
        res = list(ex.map(student, range(n)))
    wall = time.perf_counter() - t
    lat = sorted(x for r in res for x in r[0])
    cpu = [x for r in res for x in r[1]]
    p95 = lat[int(len(lat) * .95)]
    print(f"{n} sessions x {clicks} clicks in {wall:.1f} s; click CPU {statistics.mean(cpu):.1f} ms; "
          f"click latency median {statistics.median(lat) * 1000:.0f} ms, p95 {p95 * 1000:.0f} ms "
          "(includes the test harness)")
    return lat


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    a = ap.parse_args()
    imports()
    actions()
    r, at = reruns()
    memory(at)
    if not a.quick:
        concurrency()
        burst()
        sessions()
