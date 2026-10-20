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
    pieces = {k: {"attempts": [], "help": {}, "colhelp": {}, "submitted": None}
              for k in ("chase", "level", "hybrid", "summary")}
    pieces["chase"]["submitted"] = {"at": "now", "sha": "x", "file": "f", "metrics": {},
                                    "feasibility": []}
    pieces["level"]["attempts"] = [{"at": "now", "sha": "y", "file": "f", "ok": 1,
                                    "total": 2, "complete": False, "fatal": False}]
    st.session_state["pieces"] = pieces
    key = open(os.path.join(ROOT, "tests", "fixtures", "key_seed7.xlsx"), "rb").read()
    st.session_state["results"] = {
        "level": checker.check_piece(key, resolve_scenario({}, 7), "level", None)}
if qp.get("step"):
    st.session_state["step"] = qp["step"]

exec(compile(open(os.path.join(ROOT, "app.py"), encoding="utf-8").read(), "app.py", "exec"),
     globals())
