"""Load-test entry point: the real app, with each simulated student already mid-course.

  streamlit run tools/load_app.py --server.port 8599

Opened as ?sid=S7&step=level, a session starts signed in as that student with the chase
plan submitted and a checked level plan on screen (the busiest page), then runs app.py
unchanged. Used only by tools/load_test.py; never deployed."""

import os
import sys

import streamlit as st

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

qp = st.query_params
if qp.get("sid") and "_load_ready" not in st.session_state:
    from aggplan import checker
    from aggplan.scenario import resolve_scenario
    st.session_state["_load_ready"] = True
    st.session_state["student"] = {"name": f"Load {qp['sid']}", "sid": qp["sid"]}
    pieces = {k: {"attempts": [], "help": {}, "colhelp": {}, "submitted": None, "found": {},
                  "head_misses": 0}
              for k in ("chase", "level", "hybrid", "summary")}
    pieces["chase"]["submitted"] = {"at": "now", "sha": "x", "file": "f", "metrics": {},
                                    "feasibility": []}
    pieces["level"]["attempts"] = [{"at": "now", "sha": "y", "file": "f", "ok": 1,
                                    "total": 2, "complete": False, "fatal": False}]
    st.session_state["pieces"] = pieces
    key = open(os.path.join(ROOT, "tests", "fixtures", "key_seed7.xlsx"), "rb").read()
    st.session_state["results"] = {
        "level": checker.check_piece(key, resolve_scenario({}, 7), "level", None)}
    # mid-course, a student has downloaded the workbook and uploaded it at least once; the
    # app keeps the latest upload for the sidebar's "My latest workbook" and progress file
    from aggplan import persistence
    st.session_state["last_wb"] = persistence.workbook_entry("workbook.xlsx", key)
    st.session_state["downloaded_hash"] = resolve_scenario({}, 7).config_hash
if qp.get("step"):
    st.session_state["step"] = qp["step"]

# Compile app.py once per process, as Streamlit does for a deployed main script (it caches
# the bytecode); recompiling it on every click would add ~15 ms that a real class never pays.
_cache = sys.modules.setdefault("_aa_load_cache", type(sys)("_aa_load_cache"))
if not hasattr(_cache, "code"):
    _cache.code = compile(open(os.path.join(ROOT, "app.py"), encoding="utf-8").read(),
                          "app.py", "exec")
if os.environ.get("AA_PROFILE"):                 # profile the app's own run (this thread)
    import cProfile
    _cache.prof = getattr(_cache, "prof", None) or cProfile.Profile()
    _cache.prof.enable()
    try:
        exec(_cache.code, globals())
    finally:
        _cache.prof.disable()
        _cache.prof.dump_stats(os.environ["AA_PROFILE"])
else:
    exec(_cache.code, globals())
