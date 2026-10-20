"""
Juicetification: Aggregate Anxiety — Excel edition (V3)

Students build an aggregate plan in their own Excel workbook — choosing the columns as well
as the formulas — and submit it in pieces: the chase plan, the level plan, the hybrid plan,
and finally the comparison and recommendation. Each piece is uploaded and checked on its
own; every new upload unlocks more help on whatever is still wrong, but never the answer.

Steps (always visible in the sidebar):
  1 Scenario & workbook   2 Chase plan   3 Level plan   4 Hybrid plan
  5 Compare & recommend   6 Final report

Run:  streamlit run app.py        (see README.md for instructor configuration)
"""

from __future__ import annotations

import datetime as dt
import hashlib
from html import escape
from typing import Dict, List

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import student_store as store
from juice_director import resolve_config, serve_manifest_if_requested
from manifest import MANIFEST

from aggplan import checker, engine, guide, hints, layout as L, persistence, saving, text, workbook
from aggplan import identity as ident
from aggplan.recognize import HELPER
from aggplan.scenario import MONTHS as L_MONTHS, resolve_scenario

st.set_page_config(page_title="Aggregate Anxiety — Excel", page_icon="🧃", layout="wide")
SS = st.session_state
serve_manifest_if_requested(MANIFEST)
CFG, CTX = resolve_config(MANIFEST)

st.markdown("""
<style>
  div[data-testid="stSidebarContent"] .stButton button {justify-content: flex-start; text-align: left;}
  .aa-card {border: 1px solid rgba(128,128,128,.35); border-radius: 10px; padding: .8rem 1rem; margin-bottom: .6rem;}
  .aa-step {font-size: .85rem; opacity: .75; margin-bottom: -.4rem;}
  .aa-legend span {display: inline-block; margin: 0 .9rem .3rem 0; white-space: nowrap;}
  .aa-grid {overflow-x: auto;}
  .aa-grid table {border-collapse: collapse; font-size: .8rem; margin-bottom: .8rem;}
  .aa-grid th, .aa-grid td {border: 1px solid rgba(128,128,128,.45); padding: .2rem .35rem;}
  .aa-grid thead th {font-weight: 600; vertical-align: bottom; max-width: 5.6rem;
                     white-space: normal; line-height: 1.1; font-size: .72rem;}
  .aa-grid thead th .ltr {display: block; opacity: .6; font-weight: 400;}
  .aa-grid.narrow thead th {max-width: none; padding: .3rem .15rem;}
  .aa-grid.narrow thead th .hd {writing-mode: vertical-rl; transform: rotate(180deg);
                                white-space: nowrap; display: inline-block;}
  .aa-grid.narrow td {min-width: 1.45rem; padding: .15rem .1rem;}
  .aa-grid tbody th {text-align: left; font-weight: 600; white-space: nowrap;}
  .aa-grid td {text-align: center; color: #222; min-width: 1.7rem; cursor: default;}
  .aa-hint {border-left: 4px solid #1F3864; padding: .1rem 0 .1rem .8rem; margin: .2rem 0 .9rem 0;}
  .aa-hint .lvl {font-size: .75rem; text-transform: uppercase; letter-spacing: .04em; opacity: .7;}
</style>""", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
# Identity
# --------------------------------------------------------------------------- #
GAME = store.game_code()
URL_SID = store.get_student_id()
PRACTICE = bool(SS.get("practice"))
identity = ident.resolve_identity(GAME, URL_SID, st.query_params.get("tok"), store.enabled(),
                                  practice=PRACTICE)
if identity.mode == "blocked":
    st.title("🧃 Aggregate Anxiety")
    st.error("**This assignment needs your personal link.** " + identity.message)
    st.write("Open the link your instructor gave you. To practise without recording "
             "anything, continue below.")
    if st.button("Continue in practice mode"):
        SS["practice"] = True
        st.rerun()
    st.stop()

INCLUDE_HYBRID = bool(resolve_scenario(CFG, 1).params.get("include_hybrid", True))
ALL_PIECES = ["chase", "level", "hybrid", "summary"]
GOALS = {
    "chase": "Change the workforce every month so that production follows demand. You'll "
             "carry almost no inventory, but you'll hire and lay off often.",
    "level": "Keep one steady production rate and one workforce all year. Inventory builds "
             "up in slow months and is used up in busy ones — and if it runs out, customers wait.",
    "hybrid": "Design your own strategy. You choose the workers, regular production, overtime "
              "and subcontracting each month; your formulas work out the rest. It must respect "
              "every limit — and the goal is to beat both the chase and the level plan.",
    "summary": "Bring the plans together on the Summary sheet, let Excel find the cheaper one, "
               "and recommend a plan in your own words.",
}

SS.setdefault("step", "start")
SS.setdefault("student", None)          # {"name", "sid"}
SS.setdefault("results", {})            # piece -> latest PieceResult (not persisted)


def new_piece() -> dict:
    return {"attempts": [], "help": {}, "colhelp": {}, "submitted": None}


SS.setdefault("pieces", {k: new_piece() for k in ALL_PIECES})


def seed_for(sid: str) -> int:
    if CTX.get("seed") is not None:
        return int(CTX["seed"])
    return store.derive_seed(GAME, sid, lo=1, hi=10 ** 6)


POLICY_KEYS = ("safety_stock_policy", "worker_policy")
STUDENT_SETS_POLICIES = bool(CFG.get("student_sets_policies", True))
SS.setdefault("policy", {"safety_stock_policy": CFG.get("safety_stock_policy") or "none",
                         "worker_policy": CFG.get("worker_policy") or "whole"})
SS.setdefault("hybrid_mode", CFG.get("hybrid_mode") or "guided")

# The hybrid plan is a piece only when the instructor allows it and the student hasn't
# switched it off in the sidebar. Off = no Hybrid step, no hybrid on the Summary sheet,
# no hybrid in the recommendation or the report.
HYBRID_ON = INCLUDE_HYBRID and SS["hybrid_mode"] != "off"
PIECES = ["chase", "level"] + (["hybrid"] if HYBRID_ON else []) + ["summary"]
STEPS = [("start", "Scenario & workbook")] + [(k, checker.PIECE_TITLES[k]) for k in PIECES] \
    + [("report", "Final report")]
STEP_KEYS = [k for k, _ in STEPS]


def policy_cfg() -> dict:
    """Instructor configuration plus the student's two policy switches (when allowed)."""
    cfg = {**CFG, **SS["policy"]} if STUDENT_SETS_POLICIES else dict(CFG)
    cfg["include_hybrid"] = HYBRID_ON
    return cfg


def scenario():
    stu = SS["student"]
    return resolve_scenario(policy_cfg(), seed_for(stu["sid"])) if stu else None


def policies_locked() -> bool:
    """Policies can change only until the first piece is submitted (every piece must be
    checked against the same rules)."""
    return not STUDENT_SETS_POLICIES or any(submitted(k) for k in PIECES)


def submitted(key: str) -> bool:
    return bool(SS["pieces"].get(key, {}).get("submitted"))


def unlocked(key: str) -> bool:
    if key == "start":
        return True
    if not SS["student"]:
        return False
    if key == "report":
        return True
    i = PIECES.index(key)
    return i == 0 or submitted(PIECES[i - 1])


def go(step: str) -> None:
    SS["step"] = step
    save_state("nav")                      # the step you're on is part of your saved work


def _report():
    """The PDF module (and reportlab) load on first use, not at every cold start."""
    from aggplan import report
    return report


# ---- saving and restoring --------------------------------------------------------------
# Saved work = persistence.snapshot(): everything needed to continue on any computer,
# including the latest uploaded workbook. It goes to the server store when one is configured
# (in the background), and can always be downloaded as a progress file.
#
# Nothing is ever saved for a student until their saved work has been looked for and loaded:
# otherwise a fresh, empty session on a second computer could overwrite it. If saved work
# exists but can't be used (network error, or the assignment's settings changed), saving
# stays paused until the student chooses what to do.
SS.setdefault("last_wb", None)             # latest uploaded workbook (persistence entry)
SS.setdefault("restore", None)             # {"sid", "status", "msg", "saved"}


def server_saving() -> bool:
    return bool(identity.recorded and store.enabled())


def current_snapshot() -> dict:
    stu = SS["student"]
    scn = scenario()
    return persistence.snapshot(
        sid=stu["sid"], name=stu["name"], seed=scn.seed, config_hash=scn.config_hash,
        pieces=SS["pieces"], policy=SS["policy"], hybrid_mode=SS["hybrid_mode"],
        step=SS["step"], last_wb=SS.get("last_wb"), downloaded_hash=SS.get("downloaded_hash"))


def save_state(step: str) -> None:
    stu = SS.get("student")
    if not (stu and server_saving()):
        return
    r = SS.get("restore") or {}
    if r.get("sid") != stu["sid"] or r.get("status") not in ("loaded", "none", "fresh"):
        return                              # saved work not loaded yet: never overwrite it
    done = sum(submitted(k) for k in PIECES)
    # uploaded in the background so the student's screen never waits on the network
    saving.save_later(f"{GAME}|{stu['sid']}", store.save, current_snapshot(), GAME, stu["sid"],
                      progress=done / len(PIECES), step=step)


def apply_saved(saved: dict, sid: str, source: str) -> str:
    """Load saved work into this session. Returns "loaded" or "incompatible"."""
    policy = dict(SS["policy"])
    if STUDENT_SETS_POLICIES and isinstance(saved.get("policy"), dict):
        policy.update({k: v for k, v in saved["policy"].items() if k in POLICY_KEYS})
    mode = saved.get("hybrid_mode") if saved.get("hybrid_mode") in ("guided", "build", "off") \
        else SS["hybrid_mode"]
    cfg = {**CFG, **policy} if STUDENT_SETS_POLICIES else dict(CFG)
    cfg["include_hybrid"] = INCLUDE_HYBRID and mode != "off"
    scn = resolve_scenario(cfg, seed_for(sid))
    if saved.get("config_hash") != scn.config_hash or not isinstance(saved.get("pieces"), dict):
        return "incompatible"
    SS["policy"], SS["hybrid_mode"] = policy, mode
    name = saved.get("name") or (SS["student"] or {}).get("name") or sid
    SS["student"] = {"name": name, "sid": sid}
    for k in ALL_PIECES:
        SS["pieces"][k] = {**new_piece(), **(saved["pieces"].get(k) or {})}
    if saved.get("step"):
        SS["step"] = saved["step"]
    SS["last_wb"] = saved.get("last_wb")
    if saved.get("downloaded_hash"):
        SS["downloaded_hash"] = saved["downloaded_hash"]
    # bring back the last feedback by re-checking the latest workbook
    SS["results"] = {}
    data = persistence.workbook_bytes(SS["last_wb"])
    if data:
        for k in ALL_PIECES:
            ps = SS["pieces"][k]
            if ps["attempts"] and not ps["submitted"]:
                SS["results"][k] = checker.check_piece(data, scn, k, expected_sid=sid)
    SS["restore"] = {"sid": sid, "status": "loaded", "source": source}
    return "loaded"


def restore_state(sid: str) -> None:
    """Look for this student's saved work once per session (per student)."""
    r = SS.get("restore") or {}
    if r.get("sid") == sid:
        return
    if not server_saving():
        SS["restore"] = {"sid": sid, "status": "none"}
        return
    try:
        saved = store.load(GAME, sid) or {}
    except Exception as e:                                            # noqa: BLE001
        SS["restore"] = {"sid": sid, "status": "error", "msg": type(e).__name__}
        return
    if not saved:
        SS["restore"] = {"sid": sid, "status": "none"}
        return
    before = (SS["hybrid_mode"], SS["step"])
    if apply_saved(saved, sid, "server") == "incompatible":
        SS["restore"] = {"sid": sid, "status": "incompatible", "saved": saved}
        return
    if before != (SS["hybrid_mode"], SS["step"]):
        st.rerun()                          # the steps (and sidebar) depend on what was loaded


if URL_SID:
    restore_state(URL_SID)


# --------------------------------------------------------------------------- #
# Sidebar
# --------------------------------------------------------------------------- #
def _set_policy(key: str, widget: str) -> None:
    SS["policy"][key] = SS[widget]
    save_state("settings")


def _set_hybrid_mode() -> None:
    SS["hybrid_mode"] = SS["set_hybrid"]
    save_state("settings")


def sidebar_settings() -> None:
    """Three switches: safety stock, workforce, and how the hybrid plan is done."""
    st.markdown("**Settings**")
    locked = policies_locked()
    ss_val = f"{float(CFG.get('safety_stock') or 0):,.0f}"
    pol = SS["policy"] if STUDENT_SETS_POLICIES else {k: CFG.get(k) for k in POLICY_KEYS}
    ss_opts = {"none": "None (end-of-month target 0)", "maintain": f"Keep {ss_val} bottles"}
    st.radio("Safety stock", list(ss_opts), format_func=ss_opts.get,
             index=list(ss_opts).index(pol.get("safety_stock_policy") or "none"),
             key="set_ss", disabled=locked, on_change=_set_policy,
             args=("safety_stock_policy", "set_ss"),
             help="None: plans aim to end with 0 bottles on the shelf. Keep: every chase month "
                  "(and the level plan's December) aims to end with the safety stock.")
    wk_opts = {"whole": "Whole workers (round up)", "partial": "Fractional workers"}
    st.radio("Workforce", list(wk_opts), format_func=wk_opts.get,
             index=list(wk_opts).index(pol.get("worker_policy") or "whole"),
             key="set_wk", disabled=locked, on_change=_set_policy,
             args=("worker_policy", "set_wk"),
             help="Whole: a need of 10.2 workers means employing 11 (the idle capacity is paid). "
                  "Fractional: capacity grows in small increments, such as part-time hours.")
    if locked:
        st.caption("Policies are locked: " + ("set by your instructor." if not STUDENT_SETS_POLICIES
                                              else "you've submitted a piece."))
    else:
        st.caption("Set the policies before you download your workbook — it is built for them.")
    if INCLUDE_HYBRID:
        hm = {"guided": "Guided walkthrough", "build": "Build it yourself",
              "off": "Off — chase and level only"}
        locked_h = submitted("hybrid") or submitted("summary")
        st.radio("Hybrid plan", list(hm), format_func=hm.get,
                 index=list(hm).index(SS["hybrid_mode"]), key="set_hybrid",
                 disabled=locked_h, on_change=_set_hybrid_mode,
                 help="Guided: you get a starting plan and questions that lead to a good "
                      "solution. Build: you design the hybrid sheet yourself. Off: no hybrid "
                      "plan — the comparison and recommendation cover chase and level only.")
        if locked_h:
            st.caption("Locked: your comparison is built on this choice."
                       if submitted("summary") else "Locked: the hybrid plan is submitted.")


def your_work() -> None:
    """Sidebar: is my work safe, and how do I take it to another computer?"""
    stu = SS["student"]
    st.markdown("**Your work**")
    r = SS.get("restore") or {}
    if server_saving():
        stt = saving.status(f"{GAME}|{stu['sid']}")
        if r.get("status") in ("error", "incompatible"):
            st.caption("⚠ Saving is paused — see the message on step 1.")
        elif stt["state"] == "saving":
            st.caption("⟳ Saving…")
        elif stt["state"] == "failed":
            st.caption(f"⚠ Last save failed ({stt.get('error')}). It retries on your next "
                       "action; download a progress file to be safe.")
        elif stt["state"] == "saved":
            st.caption("✓ Saved to the server at "
                       + dt.datetime.fromtimestamp(stt["at"]).strftime("%H:%M:%S")
                       + ". Sign in with the same ID on any computer to continue.")
        else:
            st.caption("✓ Saved to the server as you work. Sign in with the same ID on any "
                       "computer to continue.")
    else:
        st.caption("⚠ Server saving is off for this app. Before you close this tab or switch "
                   "computers, download your progress file (and your workbook).")
    wb = persistence.workbook_bytes(SS.get("last_wb"))
    if wb:
        st.download_button("⬇ My latest workbook", wb, file_name=SS["last_wb"]["name"],
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet", width="stretch", key="dl_last_wb",
                           help="The last workbook you uploaded — continue in it on any "
                                "computer.")
    st.download_button("⬇ Progress file", persistence.encode_file(current_snapshot(),
                                                                  ident.signing_key()),
                       file_name=f"Aggregate_Anxiety_progress_{stu['sid']}.json",
                       mime="application/json", width="stretch", key="dl_progress",
                       help="Everything you've done, including your latest workbook. Upload it "
                            "on step 1 of any computer to continue.")


with st.sidebar:
    st.markdown("## 🧃 Aggregate Anxiety")
    st.caption("Aggregate planning, built by you in Excel")
    stu = SS["student"]
    if stu:
        st.markdown(f"**{stu['name']}**  \nID {stu['sid']} · Scenario #{scenario().seed}")
        st.caption(identity.label)
    st.markdown("---")
    for n, (key, label) in enumerate(STEPS, start=1):
        if key in PIECES and submitted(key):
            mark = "✔"
        elif not unlocked(key):
            mark = "🔒"
        else:
            mark = str(n)
        st.button(f"{mark}  {label}", key=f"nav_{key}", width="stretch",
                  type="primary" if SS["step"] == key else "secondary",
                  disabled=not unlocked(key), on_click=go, args=(key,))
    if stu:
        done = sum(submitted(k) for k in PIECES)
        st.markdown("---")
        st.caption(f"{done} of {len(PIECES)} pieces submitted")
        st.progress(done / len(PIECES))
    if stu:
        st.markdown("---")
        your_work()
    st.markdown("---")
    sidebar_settings()
    st.markdown("---")
    with st.expander("How this works"):
        st.write("Build the workbook one piece at a time. For each piece: design the columns, "
                 "write the formulas, save, upload. The app shows what's right; every new upload "
                 "unlocks more help on what's still wrong. When a piece is all correct, submit "
                 "it and the next one opens.")


def scroll_to_top(token: str) -> None:
    """Streamlit keeps the scroll position across reruns; a new step should start at the top."""
    if SS.get("_scrolled_for") == token:
        return
    SS["_scrolled_for"] = token
    components.html(f"""<div style="display:none">{token}</div><script>(function(){{
      try{{const d=window.parent.document;
      const up=()=>{{for(const s of ['[data-testid="stMain"]','[data-testid="stAppViewContainer"]',
                      'section.main']){{const e=d.querySelector(s); if(e){{e.scrollTop=0;}}}}
                    window.parent.scrollTo(0,0);}};
      up(); setTimeout(up,100); setTimeout(up,400);}}catch(e){{}}}})();</script>""", height=0)


def header(key: str) -> None:
    scroll_to_top(key)
    n = STEP_KEYS.index(key) + 1
    st.markdown(f"<div class='aa-step'>Step {n} of {len(STEPS)}</div>", unsafe_allow_html=True)
    st.title(dict(STEPS)[key])


def nav_buttons(key: str, next_label: str = "", next_disabled: bool = False) -> None:
    i = STEP_KEYS.index(key)
    st.markdown("")
    c1, c3 = st.columns(2)
    if i > 0:
        c1.button(f"← {STEPS[i - 1][1]}", on_click=go, args=(STEP_KEYS[i - 1],),
                  width="stretch", key=f"back_{key}")
    if i < len(STEPS) - 1:
        nxt = STEP_KEYS[i + 1]
        c3.button(next_label or f"{STEPS[i + 1][1]} →", type="primary", width="stretch",
                  on_click=go, args=(nxt,), disabled=next_disabled or not unlocked(nxt),
                  key=f"next_{key}")


def md_table(df: pd.DataFrame) -> None:
    """A wrapped, full-width table for text-heavy content (st.dataframe truncates text)."""
    esc = lambda x: text.md(str(x)).replace("|", "\\|").replace("\n", " ")  # noqa: E731
    lines = ["| " + " | ".join(df.columns) + " |", "|" + "---|" * len(df.columns)]
    lines += ["| " + " | ".join(esc(v) for v in row) + " |" for row in df.itertuples(index=False)]
    st.markdown("\n".join(lines))


# --------------------------------------------------------------------------- #
# Step 1: scenario & workbook
# --------------------------------------------------------------------------- #
def demand_chart(scn) -> alt.Chart:
    df = pd.DataFrame({"Month": [m[:3] for m in scn.months], "Demand": scn.demand})
    return (alt.Chart(df).mark_bar(color="#F28C28").encode(
        x=alt.X("Month", sort=None, title=None), y=alt.Y("Demand", title="bottles"),
        tooltip=["Month", alt.Tooltip("Demand", format=",.0f")]).properties(height=240))


def show_policies(scn, expanded: bool = False) -> None:
    """Safety stock, backorders and workforce rules — one card each."""
    cols = st.columns(3) if expanded else [st.container()] * 3
    for col, (title, summary, details) in zip(cols, text.policies(scn)):
        with col:
            if expanded:
                with st.container(border=True):
                    st.markdown(f"**{text.md(title)}**")
                    st.markdown(text.md(summary))
                    st.markdown("\n".join(f"- {text.md(d)}" for d in details))
            else:
                st.markdown(f"**{text.md(title)}** — {text.md(summary)}")
                st.markdown("\n".join(f"- {text.md(d)}" for d in details))


def restore_panel() -> None:
    """Saved work that couldn't be used, and continuing from a progress file."""
    r = SS.get("restore") or {}
    if r.get("status") == "error":
        st.error(f"**Your saved work couldn't be loaded** ({r.get('msg')}). Saving is paused "
                 "so nothing is overwritten. Try again in a moment.")
        if st.button("Try loading again", type="primary"):
            SS["restore"] = None
            restore_state(SS["student"]["sid"] if SS["student"] else URL_SID)
            st.rerun()
    elif r.get("status") == "incompatible":
        saved = r.get("saved") or {}
        st.warning("**You have saved work, but it was made for a different setup** — your "
                   "instructor changed the assignment, or you chose different policies. It "
                   "hasn't been loaded, and saving is paused so it isn't overwritten.")
        wb = persistence.workbook_bytes(saved.get("last_wb"))
        c1, c2 = st.columns(2)
        if wb:
            c1.download_button("⬇ Download my saved workbook", wb,
                               file_name=saved["last_wb"]["name"], width="stretch")
        if c2.button("Start fresh (replaces the saved work)", width="stretch"):
            SS["restore"] = {"sid": r["sid"], "status": "fresh"}
            save_state("fresh")
            st.rerun()
    with st.expander("Continue from a progress file", expanded=not server_saving() and not SS["student"]):
        st.caption("Downloaded from the sidebar (**Your work ▸ Progress file**) on another "
                   "computer. It brings back everything, including your latest workbook.")
        up = st.file_uploader("Progress file (.json)", type=["json"], key="progress_up")
        if up is not None:
            state, err = persistence.decode_file(up.getvalue(), ident.signing_key())
            sid = (state or {}).get("sid")
            if err:
                st.error(err)
            elif (URL_SID and sid != URL_SID) or (SS["student"] and sid != SS["student"]["sid"]):
                st.error(f"This progress file belongs to student ID '{sid}', not yours.")
            elif SS.get("_progress_loaded") != up.file_id:
                SS["_progress_loaded"] = up.file_id
                if apply_saved(state, sid, "file") == "incompatible":
                    st.error("This progress file was made for a different setup of the "
                             "assignment, so it can't be loaded here.")
                else:
                    if not URL_SID and not PRACTICE:
                        store.set_student_id(sid)
                    save_state("restored")
                    st.success("Your work is back.")
                    st.rerun()


def page_start() -> None:
    header("start")
    restore_panel()
    if not SS["student"]:
        st.write("You are the operations planner at **Juicetification Inc.**, a juice bottler "
                 "with a strong summer peak. You'll plan a year of production in Excel — "
                 "designing the worksheet yourself. Your student ID gives you your own scenario.")
        with st.form("who"):
            name = st.text_input("Your name")
            sid = st.text_input("Student ID", value=URL_SID or "", disabled=bool(URL_SID),
                                help="Your scenario is made from this ID: use the same ID every time.")
            if st.form_submit_button("Start", type="primary"):
                sid = (URL_SID or sid).strip()
                if not name.strip() or not sid:
                    st.error("Enter your name and your student ID.")
                else:
                    SS["student"] = {"name": name.strip(), "sid": sid}
                    if not URL_SID and not PRACTICE:
                        store.set_student_id(sid)   # survives a refresh; enables recording
                    else:
                        restore_state(sid)          # load saved work before anything is saved
                    st.rerun()
        return

    scn = scenario()
    stu = SS["student"]
    for note in scn.notices:
        st.warning(note)
    st.write("Demand for bottled juice swings with the seasons. Your job is to decide how many "
             "workers to employ and how much to make each month so the year's demand is met at "
             "the lowest total cost.")
    m = st.columns(4)
    m[0].metric("Annual demand", f"{scn.total_demand:,.0f}")
    m[1].metric("Peak month", scn.months[scn.demand.index(max(scn.demand))])
    m[2].metric("Peak demand", f"{max(scn.demand):,.0f}")
    m[3].metric("Lowest demand", f"{min(scn.demand):,.0f}")
    st.altair_chart(demand_chart(scn), use_container_width=True)
    st.subheader("Your company's planning policies")
    st.caption("Every plan must follow these; they are also on the Scenario sheet of your "
               "workbook. " + ("Change them under **Settings** in the sidebar before you "
                               "download your workbook." if not policies_locked() else ""))
    show_policies(scn, expanded=True)
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.subheader("Given data")
        md_table(pd.DataFrame(text.param_rows(scn), columns=["Item", "Value", "Unit"]))
    with c2:
        st.subheader("Your workbook")
        if SS.get("wb_for") != (stu["sid"], scn.config_hash, HYBRID_ON):
            SS["wb_bytes"] = workbook.build_workbook(scn, student_name=stu["name"],
                                                     student_id=stu["sid"], game=GAME)
            SS["wb_for"] = (stu["sid"], scn.config_hash, HYBRID_ON)
        got = SS.get("downloaded_hash")
        if got and got != scn.config_hash:
            st.warning("You changed a policy after downloading. **Download the workbook again** "
                       "— the old one was built for the other policy and won't be accepted. "
                       "Copy across any work you already did.")

        def _downloaded(h=scn.config_hash):
            SS["downloaded_hash"] = h

        st.download_button("⬇ Download my Excel workbook", SS["wb_bytes"],
                           file_name=workbook.file_name(stu["name"], stu["sid"]),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", on_click=_downloaded)
        st.caption("Download it once and keep working in the same file for every piece. Works "
                   "in Excel and Google Sheets (File → Download → .xlsx).")
        st.markdown("**You'll submit it in pieces:**")
        st.markdown("\n".join(f"{i}. **{checker.PIECE_TITLES[k]}** — {GOALS[k].split('.')[0]}."
                              for i, k in enumerate(PIECES, start=1)))
        st.markdown("On each plan sheet only **Month** and **Demand** are given. You decide "
                    "which columns the plan needs, name them, and write the formulas.")
        with st.expander("Excel tips"):
            for t, d in text.EXCEL_TIPS:
                st.markdown(f"**{t}.** {text.md(d)}")
            st.markdown("**Names for the given data:** "
                        + ", ".join(f"`{n}`" for n, _, _, _ in scn.name_keys()))
    nav_buttons("start", "Start the chase plan →")


# --------------------------------------------------------------------------- #
# Piece pages
# --------------------------------------------------------------------------- #
COLORS = {"ok": "#C6EFCE", "typed": "#FFEB9C", "carried": "#E4DFEC", "wrong": "#FFC7CE",
          "blank": "#F2F2F2", "error": "#F8CBAD", "text": "#F8CBAD", "uncalc": "#F8CBAD",
          "decision": "#DDEBF7", "waiting": "#FFFFFF", "helper": "#F7F7F7"}
SYMBOL = {"ok": "✔", "typed": "✎", "carried": "↪", "wrong": "✖", "blank": "○",
          "error": "⚠", "text": "⚠", "uncalc": "⚠", "decision": "◆",
          "waiting": "…", "helper": "·"}


def item_name(kind: str, item: str, res: checker.PieceResult) -> str:
    if item.startswith("plan:"):
        return next(pc.label for pc in L.PLANNING[kind] if pc.key == item[5:]) + " (planning cell)"
    if item == "total":
        return "Yearly totals (row 14)"
    if item.startswith("summary:"):
        return {"summary:link": "Totals from your plans", "summary:lower": "Lower-cost plan",
                "summary:savings": "Savings", "summary:recommend": "Recommendation",
                "summary:justify": "Justification"}.get(item, item[8:])
    for letter, heading, read_as in res.columns:
        if read_as == item:
            return f"{heading} (column {letter})"
    return item


def hint_block(kind: str, item: str, level: int) -> str:
    steps = hints.hint_text(kind, item, level, scenario().whole_workers)
    parts = []
    for i, s in enumerate(steps, start=1):
        parts.append(f"<div class='lvl'>Hint {i} · {hints.LEVEL_NAMES[i]}</div>"
                     f"<div>{escape(s)}</div>")
    more = ("" if level >= hints.MAX_LEVEL else
            "<div class='lvl' style='margin-top:.3rem'>More help unlocks if this is still wrong "
            "on your next upload.</div>")
    return f"<div class='aa-hint'>{''.join(parts)}{more}</div>"


def grid_html(res: checker.PieceResult) -> str:
    g = checker.grid(res)
    if g.empty:
        return ""
    def _th(c):
        letter, _, heading = c.partition(" · ")
        return (f"<th title='{escape(c)}'><span class='ltr'>{escape(letter)}</span>"
                f"<span class='hd'>{escape(heading)}</span></th>")
    head = "".join(_th(c) for c in g.columns)
    body = []
    for row in g.index:
        tds = []
        for col in g.columns:
            c = g.at[row, col]
            if not hasattr(c, "status"):            # empty grid position
                tds.append("<td></td>")
                continue
            tip = escape(f"{c.ref}: {checker.STATUS_LABEL[c.status]}"
                         + (f" (shows {c.shown})" if c.shown else "")
                         + (f" — note: {c.notice}" if c.notice else ""))
            mark = "<sup>!</sup>" if c.notice else ""
            tds.append(f"<td title=\"{tip}\" style=\"background:{COLORS[c.status]}\">"
                       f"{SYMBOL[c.status]}{mark}</td>")
        body.append(f"<tr><th>{escape(row)}</th>{''.join(tds)}</tr>")
    narrow = " narrow" if len(g.columns) > 12 else ""
    return (f"<div class='aa-grid{narrow}'><table><thead><tr><th></th>{head}</tr></thead>"
            f"<tbody>{''.join(body)}</tbody></table></div>")


def legend() -> None:
    st.markdown("<div class='aa-legend'>" + "".join(
        f"<span><span style='background:{COLORS[k]};color:#222;padding:0 .35rem;border-radius:3px'>"
        f"{SYMBOL[k]}</span> {checker.STATUS_LABEL[k].split(' ', 1)[1]}</span>"
        for k in ("ok", "wrong", "carried", "typed", "blank", "error", "decision", "waiting",
                  "helper"))
        + f"<span><span style='background:{COLORS['ok']};color:#222;padding:0 .35rem;"
          "border-radius:3px'>✔<sup>!</sup></span> correct, but see the note</span>"
        + "</div>", unsafe_allow_html=True)


def run_check(key: str, up) -> None:
    data = up.getvalue()
    sha = hashlib.sha256(data).hexdigest()
    ps = SS["pieces"][key]
    if ps["attempts"] and ps["attempts"][-1]["sha"] == sha and key in SS["results"]:
        return
    res = checker.check_piece(data, scenario(), key, expected_sid=SS["student"]["sid"])
    SS["results"][key] = res
    if not res.fatal:
        SS["last_wb"] = persistence.workbook_entry(up.name, data) or SS.get("last_wb")
    if ps["attempts"] and ps["attempts"][-1]["sha"] == sha:
        return                                   # same file again: not a new attempt
    ps["attempts"].append({"at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"), "sha": sha,
                           "file": up.name, "ok": res.n_ok, "total": res.n_total,
                           "complete": res.complete, "fatal": bool(res.fatal)})
    if not res.fatal:
        # more help only where a real mistake persists (blank/carried cells never count)
        for item in res.problem_items():
            ps["help"][item] = min(hints.MAX_LEVEL, ps["help"].get(item, 0) + 1)
        if len(res.columns) >= 3:
            for concept in res.missing_columns:
                ps["colhelp"][concept] = min(hints.MAX_COLUMN_LEVEL,
                                             ps["colhelp"].get(concept, 0) + 1)
    save_state(key)
    st.rerun()


def piece_record(key: str) -> dict:
    ps = SS["pieces"][key]
    sub = ps.get("submitted")
    ok = [a for a in ps["attempts"] if not a.get("fatal")]
    first, last = (ok[0], ok[-1]) if ok else (None, None)
    return {"title": checker.PIECE_TITLES[key],
            "status": "Submitted" if sub else ("In progress" if ps["attempts"] else "Not started"),
            "attempts": len(ok),
            "first": f"{first['ok']} / {first['total']} correct" if first else "",
            "last": f"{last['ok']} / {last['total']} correct" if last else "",
            "history": " → ".join(f"{a['ok']}/{a['total']}" for a in ok),
            "submitted_at": sub["at"] if sub else "",
            "help": max(list(ps["help"].values()) + list(ps["colhelp"].values()) + [0]),
            "file": sub["file"] if sub else "", "sha": sub["sha"] if sub else ""}


def submit_piece(key: str, res: checker.PieceResult) -> None:
    ps = SS["pieces"][key]
    last = ps["attempts"][-1]
    ps["submitted"] = {"at": _report().now_in(str(scenario().params.get("report_timezone") or ""))
                       .strftime("%Y-%m-%d %H:%M %Z").strip(),
                       "sha": res.sha256, "file": last["file"], "metrics": res.metrics,
                       "feasibility": res.feasibility, "recommendation": res.recommendation,
                       "justification": res.justification, "words": res.words}
    save_state(f"submitted:{key}")


def _plan_table(scn) -> None:
    """The student's two submitted plans side by side."""
    mets = {k: (SS["pieces"][k]["submitted"] or {}).get("metrics", {}).get(k)
            for k in ("chase", "level")}
    if not all(mets.values()):
        mets = {k: engine.summarize(engine.reference_plan(scn, k)) for k in ("chase", "level")}
    rows = []
    for m, lab in (("Labor Cost", "Labor"), ("Hiring Cost", "Hiring"),
                   ("Layoff Cost", "Layoffs"), ("Holding Cost", "Holding"),
                   ("Backorder Cost", "Backorders"), ("Total Cost", "TOTAL")):
        rows.append({"Your plans": lab, "Chase": f"${mets['chase'][m]:,.0f}",
                     "Level": f"${mets['level'][m]:,.0f}"})
    rows.append({"Your plans": "On-time fulfillment",
                 "Chase": f"{mets['chase']['On-Time Fulfillment %']:.1f}%",
                 "Level": f"{mets['level']['On-Time Fulfillment %']:.1f}%"})
    md_table(pd.DataFrame(rows))


def _bottle_costs(scn) -> None:
    p, cap = scn.params, scn.cap_month
    hire_cycle = (p["hiring_cost"] + p["layoff_cost"] + p["regular_labor_cost"]) / cap
    st.markdown("**What one more bottle in a busy month costs** (from your given data):")
    md_table(pd.DataFrame([
        {"Option": "Regular time (workers already employed and idle)",
         "Cost per bottle": "$0 more — already paid", "Limits": "only the idle capacity"},
        {"Option": "Regular time (a worker paid for the whole month)",
         "Cost per bottle": f"${p['regular_labor_cost'] / cap:,.2f}",
         "Limits": f"comes in steps of {cap:,.0f} bottles" + (" (whole workers)"
                                                              if scn.whole_workers else "")},
        {"Option": "Overtime", "Cost per bottle": f"${p['overtime_cost']:,.2f}",
         "Limits": f"only when regular capacity is full; up to {p['overtime_pct']:.0%} of it"},
        {"Option": "Subcontract", "Cost per bottle": f"${p['subcontract_cost']:,.2f}",
         "Limits": f"up to {p['subcontract_capacity']:,.0f} bottles a month"},
        {"Option": "Make it earlier and hold it",
         "Cost per bottle": f"${p['holding_cost']:,.2f} per month held",
         "Limits": "needs spare capacity in an earlier month"},
        {"Option": "Hire a worker for one month", "Cost per bottle": f"about ${hire_cycle:,.2f}",
         "Limits": f"hire + layoff + a month's labor over {cap:,.0f} bottles"},
        {"Option": "Let the customer wait",
         "Cost per bottle": f"${p['backorder_cost']:,.2f} per month late",
         "Limits": "hurts on-time fulfillment; nothing may be owed in December"},
    ]))


def hybrid_start_guide(scn, expanded: bool) -> None:
    """Build-it-yourself mode: where to start, and why — from the student's own plans."""
    best = guide.best_start(scn)
    with st.expander(f"Where to start: from your {best.upper()} plan (and why)",
                     expanded=expanded):
        _plan_table(scn)
        if best == "chase":
            st.markdown(
                "**Start from your chase plan — it is the cheaper one.** A hybrid improves a "
                "plan by fixing its most expensive habit. Here regular labor is cheap compared "
                "with holding a workforce you don't need, so following demand already does "
                "well; the waste is in months where you pay a whole extra worker for only a "
                "few hundred bottles. Overtime and subcontracting can cover those bottles "
                "instead — check which is cheaper with the table below.")
        else:
            st.markdown(
                "**Start from your level plan — it is the cheaper one.** Hiring and layoffs are "
                "expensive here, so a steady workforce pays. Its weak spots are the months where "
                "customers wait: use the capacity you already pay for first, then the cheapest "
                "extra bottles (table below), only as much as needed.")
        st.markdown(
            "**Steps**\n"
            f"1. Copy your {best.title()} Plan's columns onto the Hybrid Plan sheet. Paste "
            "**Workers** and **Regular Production** as *values* (Paste Special → Values) — they "
            "are now your decisions, not formulas pointing at the old planning cells.\n"
            "2. Add the columns a hybrid also needs (overtime, subcontracting, total production "
            "and their costs) and fill the hybrid planning cells.\n"
            "3. Upload: the app checks the formulas and every limit.\n"
            "4. Change one thing at a time where a cheaper source of bottles exists, and "
            "upload again — your cost is shown next to the chase and level plans.")
        _bottle_costs(scn)


# ----------------------------------------------------------------------------------
# Guided walkthrough
# ----------------------------------------------------------------------------------
def _ask(q, g: dict) -> bool:
    """Render one question; True once it is answered (or revealed)."""
    state = g["answers"].setdefault(q.qid, {"tries": 0, "ok": False})
    fmt = (lambda v: guide.RULE_LABELS.get(v, v)) if q.qid == "c_rule" else str
    if state["ok"]:
        shown = fmt(q.answer) if q.kind == "choice" else (
            ", ".join(m for m in L_MONTHS if m in q.answer) if q.kind == "months"
            else f"{float(q.answer):,.2f}".rstrip("0").rstrip("."))
        st.markdown(f"<div class='aa-hint'><div class='lvl'>"
                    f"{'Revealed' if state.get('revealed') else '✔ Answered'}</div>"
                    f"<b>{escape(q.prompt)}</b><br>{escape(shown)}<br>"
                    f"<span style='opacity:.85'>{escape(q.why)}</span></div>",
                    unsafe_allow_html=True)
        return True
    with st.container(border=True):
        st.markdown(f"**{q.prompt}**")
        if q.kind == "choice":
            val = st.radio("Your answer", list(q.options), format_func=fmt, index=None,
                           key=f"g_{q.qid}", label_visibility="collapsed")
        elif q.kind == "months":
            val = st.multiselect("Months", L_MONTHS, key=f"g_{q.qid}",
                                 label_visibility="collapsed")
        else:
            val = st.number_input("Your answer", value=None, step=1.0, format="%.2f",
                                  key=f"g_{q.qid}", label_visibility="collapsed")
        c1, c2 = st.columns([1, 3])
        if c1.button("Check answer", key=f"gc_{q.qid}", type="primary"):
            if guide.check_answer(q, val):
                state["ok"] = True
            else:
                state["tries"] += 1
            save_state("guide")
            st.rerun()
        if state["tries"]:
            st.warning(f"Not quite ({state['tries']} tr{'y' if state['tries'] == 1 else 'ies'}). "
                       + q.nudge)
        if state["tries"] >= 3 and c2.button("Show me the answer", key=f"gr_{q.qid}"):
            state.update(ok=True, revealed=True)
            save_state("guide")
            st.rerun()
    return False


def guided_hybrid(scn, ps: dict) -> bool:
    """The guided hybrid. Returns True once the student has reached the Excel task, so the
    upload section and the verdict should be shown."""
    g = ps.setdefault("guide", {"start": None, "answers": {}})
    st.markdown("<div class='aa-card'><b>Guided hybrid.</b> You'll get a starting plan with "
                "every column already built, then answer a few questions that lead you to a "
                "better plan. You make each change yourself in Excel. (Prefer to build it all "
                "yourself? Switch <i>Hybrid plan</i> in the sidebar.)</div>",
                unsafe_allow_html=True)
    st.markdown("#### 1 · Pick your starting point")
    _plan_table(scn)
    q0 = guide.start_question(scn)
    if not _ask(q0, g):
        return False
    g["start"] = q0.answer.lower()
    start = g["start"]

    st.markdown("#### 2 · Get your starting plan")
    st.write(f"Upload your **current** workbook here (the one with your submitted chase and "
             f"level plans). The app fills the **Hybrid Plan** sheet with your {start} plan: all "
             "columns, formulas and totals, with Workers and Regular Production as your "
             "decisions (green) and no overtime or subcontracting yet. Grey **helper columns** "
             "on the right keep your starting plan's numbers fixed, so you can compare against "
             "them as you change the green cells. Anything already on the Hybrid Plan sheet is "
             "replaced; other sheets are untouched.")
    seed = st.file_uploader("Workbook to add the starting plan to (.xlsx)", type=["xlsx"],
                            key="hyb_seed")
    if seed is not None:
        data = seed.getvalue()
        sig = (hashlib.sha256(data).hexdigest(), start, scn.config_hash)
        cached = SS.get("hyb_seed_out")
        if not cached or cached[0] != sig:          # build once per file, not every click
            book, err = checker.read_book(data, ())
            problems = [err] if err else checker.check_identity(book, scn, SS["student"]["sid"])
            out = None if problems else workbook.add_hybrid_start(
                data, scn, guide.starting_decision(scn, start))
            SS["hyb_seed_out"] = cached = (sig, problems, out)
        _, problems, out = cached
        if problems:
            for m in problems:
                st.error(m)
        else:
            st.download_button("⬇ Download workbook with the starting plan", out,
                               file_name=seed.name.replace(".xlsx", "") + "_hybrid_start.xlsx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet", type="primary")
            st.info("**Open it in Excel and save it** (so Excel calculates every formula) "
                    "before you upload it below. Keep working in this file from now on.")

    st.markdown("#### 3 · Questions")
    track = guide.track_name(scn, start)
    for step in guide.steps(scn, start):
        if isinstance(step, guide.Task):
            st.markdown(f"#### 4 · {step.title}")
            with st.container(border=True):
                st.markdown(text.md(step.text))
            g["track"] = track
            return True
        if not _ask(step, g):
            return False
    return True


def guided_verdict(scn, ps: dict, res):
    g = ps.get("guide") or {}
    if not g.get("start") or res.plan is None:
        return None
    v = guide.evaluate(scn, g["start"], res.plan)
    st.markdown("#### 5 · How good is your hybrid?")
    md_table(pd.DataFrame([{"Plan": "Your hybrid", "Total cost": f"${v.cost:,.0f}",
                            "Feasible": "✔ yes" if v.feasible else "✖ no"},
                           {"Plan": "Chase", "Total cost": f"${v.chase:,.0f}", "Feasible": ""},
                           {"Plan": "Level", "Total cost": f"${v.level:,.0f}", "Feasible": ""}]))
    # how much opportunity help: grows with each upload that is feasible but not good yet
    if v.feasible and not v.good and g.get("coach_sha") != res.sha256:
        g["coach_sha"] = res.sha256
        g["coach_level"] = min(3, g.get("coach_level", 0) + 1)
        save_state("guide")
    fixes, opps = guide.coach(scn, g["start"], res.plan, max(1, g.get("coach_level", 1)))
    if fixes:
        st.markdown("**Fix these months** — your plan breaks a limit here:")
        st.markdown("\n".join(f"- {text.md(f)}" for f in fixes))
    for n in v.notes:
        if not n.startswith("✖"):
            st.markdown(text.md(n))
    if opps:
        lvl = g.get("coach_level", 1)
        st.markdown(f"**Where money is still left** (help step {lvl} of 3 — more detail on "
                    "each upload):")
        st.markdown("\n".join(f"- {text.md(o)}" for o in opps))
    if v.good:
        saved = v.start_cost - v.cost
        st.success("🎯 A good solution: feasible and "
                   + (f"${saved:,.0f} cheaper than your starting plan — and cheaper than both "
                      "pure strategies." if saved > 0.5 else
                      "as cheap as the best pure strategy (nothing cheaper is possible here)."))
        if _ask(guide.reflection(g["start"], guide.track_name(scn, g["start"])), g):
            st.caption("Submit your hybrid plan below.")
    elif v.feasible and not opps:
        st.info("Feasible — now make it cheaper than both plans. Re-read the checklist above "
                "and check each month; upload after each change.")
    return v


def show_guidance(key: str, scn) -> None:
    ps = SS["pieces"][key]
    st.markdown(f"<div class='aa-card'><b>The goal.</b> {escape(GOALS[key])}</div>",
                unsafe_allow_html=True)
    if key == "summary":
        st.markdown("Fill the **Summary** sheet: each plan's yearly figures, which plan costs "
                    "less, the savings, your recommendation and why. Use formulas that point "
                    "at your plan sheets.")
        if INCLUDE_HYBRID and not HYBRID_ON:
            st.info("The hybrid plan is **off**, so the comparison covers the chase and level "
                    "plans only. If your Summary sheet has a Hybrid column, leave it empty — it "
                    "isn't checked — and recommend **Chase** or **Level**.")
    else:
        with st.expander("① Design your columns", expanded=not ps["attempts"]):
            st.write(f"On the **{L.PLAN_SHEETS[key]}** sheet, only Month and Demand are given. "
                     "Every month your plan must answer the questions below. Turn each one into "
                     "a column: type a clear heading in row 1 (from column C on), then the "
                     "formulas for January–December, then the yearly totals in row 14 under "
                     "every cost column. Put them in whatever order makes sense to you.")
            qs = [hints.QUESTIONS[c] for c in L.REQUIRED[key]]
            st.markdown("\n".join(f"- {q}" for q in qs))
            if key == "hybrid":
                st.caption("Four of these are your decisions (workers, regular production, "
                           "overtime, subcontracting): type them as numbers or formulas. "
                           "Everything else is calculated.")
            st.caption("Ask yourself for each one: which other columns or given values does it "
                       "depend on? That's your formula. Helper columns are welcome.")
        with st.expander("② Planning cells (do these first)", expanded=not ps["attempts"]):
            st.write(f"Under the grid on the **{L.PLAN_SHEETS[key]}** sheet:")
            md_table(pd.DataFrame([{"Cell": L.planning_ref(key, pc.key),
                                    "Work out": f"{pc.label} ({pc.unit})",
                                    "Think about": hints.ladder(key, f"plan:{pc.key}",
                                                                scn.whole_workers)[0]}
                                   for pc in L.PLANNING[key]]))
    if key != "summary":
        with st.expander("Your company's policies: safety stock, backorders, workers"):
            show_policies(scn)
    if key == "hybrid":
        hybrid_start_guide(scn, expanded=not ps["attempts"])
    if ps["help"] or ps["colhelp"]:
        with st.expander(f"Your hints so far ({len(ps['help']) + len(ps['colhelp'])})"):
            res = SS["results"].get(key) or checker.PieceResult(key, "")
            for item, lvl in ps["help"].items():
                st.markdown(f"**{escape(item_name(key, item, res))}**")
                st.markdown(hint_block(key, item, lvl), unsafe_allow_html=True)
            for concept, lvl in ps["colhelp"].items():
                st.markdown(f"- {escape(hints.column_hint(key, concept, lvl))}")


def show_columns(key: str, res: checker.PieceResult, ps: dict) -> None:
    st.markdown("**Your columns**")
    if not res.columns:
        st.info(f"No column headings yet. Type your headings in row 1 of the "
                f"{L.PLAN_SHEETS[key]} sheet, starting in column C.")
    else:
        rows = []
        for letter, heading, read_as in res.columns:
            if read_as == HELPER:
                read = "helper column (fine — not checked)"
            elif read_as:
                guessed = any(n.kind == "guessed" and letter in n.letters for n in res.notes)
                read = f"✔ {read_as}" + (" (guessed from its numbers)" if guessed else "")
            else:
                read = "? not recognised"
            rows.append({"Column": letter, "Your heading": heading, "Read as": read})
        md_table(pd.DataFrame(rows))
    for n in res.notes:
        if n.kind == "missing":
            lvl = max(1, ps["colhelp"].get(n.concept, 0))
            st.warning(hints.column_hint(key, n.concept, lvl))
        elif n.kind == "duplicate":
            first, extra = n.letters[0], ", ".join(n.letters[1:])
            st.info(f"Columns {' and '.join(n.letters)} all look like **{n.concept}**. Column "
                    f"{first} is checked as {n.concept}; {extra} is treated as a helper column "
                    "(shown, not checked). If that's the wrong way round, put the word "
                    "**helper** in the extra column's heading.")
        elif n.kind == "guessed":
            st.info(f"Column {n.letters[0]} (“{n.heading}”) behaves like **{n.concept}**, so it "
                    "was checked as that. Give it a clearer heading so a reader knows too.")
        elif n.kind == "unknown":
            st.caption(f"Column {n.letters[0]} (“{n.heading}”) wasn't recognised. If it's a "
                       "helper column, that's fine; if it's one of the questions above, use a "
                       "heading that says what it measures.")


def show_feedback(key: str, res: checker.PieceResult, ps: dict) -> None:
    problems = res.problem_items()
    todo = res.todo_items()
    carried = sum(c.status == "carried" for c in res.cells)
    if problems:
        st.subheader("What to work on")
        order = {c.item: n for n, c in enumerate(res.cells)}
        for item in sorted(problems, key=lambda it: order[it]):
            cells = problems[item]
            first = cells[0]
            refs = ", ".join(c.ref for c in cells[:6]) + (" …" if len(cells) > 6 else "")
            statuses = sorted({checker.STATUS_LABEL[c.status] for c in cells})
            lvl = max(1, ps["help"].get(item, 1))
            st.markdown(f"**{escape(item_name(key, item, res))}** — {', '.join(statuses)} in "
                        f"{len(cells)} cell{'s' if len(cells) != 1 else ''}: {refs}")
            extra = []
            if any(c.status == "typed" for c in cells):
                extra.append(checker.STATUS_HELP["typed"])
            if any(c.status in ("error", "text", "uncalc") for c in cells):
                extra.append(checker.STATUS_HELP[next(c.status for c in cells
                                                      if c.status in ("error", "text", "uncalc"))])
            if lvl >= 2 and first.diag:
                extra.append(f"In {first.ref}: {first.diag}")
            for e in extra:
                st.caption(text.md(e))
            if any(c.status == "wrong" for c in cells):
                st.markdown(hint_block(key, item, lvl), unsafe_allow_html=True)
    notices = res.notice_items()
    if notices:
        st.markdown("**Works now, but…** (doesn't stop you submitting)")
        for item, cells in notices.items():
            refs = ", ".join(c.ref for c in cells[:6]) + (" …" if len(cells) > 6 else "")
            st.caption(text.md(f"**{item_name(key, item, res)}** ({refs}): {cells[0].notice}"))
    if carried:
        st.caption(f"↪ {carried} more cell{'s' if carried != 1 else ''} only follow an earlier "
                   "mistake: their formulas look right and should correct themselves once the "
                   "cells above are fixed.")
    if todo:
        names = [f"{item_name(key, it, res)} ({len(cs)})" for it, cs in todo.items()]
        st.markdown("**Still empty:** " + text.md(", ".join(names)))


def _receipt(key: str, scn, sub: dict):
    """The one-page receipt for a submitted piece: (pdf bytes, report id)."""
    return _report().build_report(
        scn, student_name=SS["student"]["name"], student_id=SS["student"]["sid"],
        identity_label=identity.label, pieces=[piece_record(key)], final=True,
        heading=f"Receipt — {checker.PIECE_TITLES[key]} submitted",
        comparison=sub.get("metrics") or None, feasibility=sub.get("feasibility") or None,
        recommendation=sub.get("recommendation", ""),
        justification=sub.get("justification") if key == "summary" else None,
        words=sub.get("words", 0))


def page_piece(key: str) -> None:
    header(key)
    scn = scenario()
    ps = SS["pieces"][key]
    if not unlocked(key):
        prev = PIECES[PIECES.index(key) - 1]
        st.info(f"Submit the **{checker.PIECE_TITLES[prev]}** first — the pieces open one at a time.")
        st.button(f"Go to {checker.PIECE_TITLES[prev]}", on_click=go, args=(prev,))
        return
    if ps["submitted"]:
        sub = ps["submitted"]
        n = len([a for a in ps["attempts"] if not a.get("fatal")])
        st.success(f"✔ Submitted {sub['at']} after {n} upload{'s' if n != 1 else ''}.")
        receipts = SS.setdefault("receipts", {})
        sig = (sub.get("sha"), sub.get("at"))
        if receipts.get(key, {}).get("sig") != sig:   # build once per submission
            receipts[key] = {"sig": sig, "pdf": None}
        if receipts[key]["pdf"] is None:
            receipts[key]["pdf"], receipts[key]["rid"] = _receipt(key, scn, sub)
        pdf, rid = receipts[key]["pdf"], receipts[key]["rid"]
        st.download_button("⬇ Download receipt (PDF)", pdf,
                           file_name=_report().report_file_name(SS["student"]["name"], key, rid),
                           mime="application/pdf")
        st.caption("Upload the receipt to your LMS if your instructor collects each piece.")
        nav_buttons(key)
        return

    guided = key == "hybrid" and SS["hybrid_mode"] == "guided"
    if guided:
        if not guided_hybrid(scn, ps):
            nav_buttons(key, next_disabled=True)
            return
    else:
        show_guidance(key, scn)
    st.subheader("Upload & check")
    st.write("Save your workbook in Excel, then upload it here. Only the "
             f"**{checker.PIECE_TITLES[key].lower()}** is checked at this step.")
    up = st.file_uploader("Your workbook (.xlsx)", type=["xlsx"], key=f"up_{key}")
    if up is not None:
        run_check(key, up)
    res = SS["results"].get(key)
    if res is None:
        nav_buttons(key, next_disabled=True)
        return
    n = len([a for a in ps["attempts"] if not a.get("fatal")])
    st.caption(f"Upload #{n} for this piece · {ps['attempts'][-1]['file']} · "
               f"{ps['attempts'][-1]['at']}")
    if res.fatal:
        for msg in res.fatal:
            st.error(msg)
        nav_buttons(key, next_disabled=True)
        return
    if res.uncalculated:
        st.warning("**Some formulas have no saved results.** Open the file in Excel, save it as "
                   ".xlsx, and upload again.")
    for msg in res.messages:
        st.warning(msg)
    c1, c2 = st.columns([1, 2])
    c1.metric("Formulas correct" if key == "hybrid" else "Cells correct",
              f"{res.n_ok} / {res.n_total}")
    c2.progress(res.n_ok / max(1, res.n_total))
    if key == "hybrid":
        c2.caption("This counts cells calculated correctly from your decisions. Whether the "
                   "plan itself is good — cheaper than chase and level, within every limit — "
                   "is checked below.")
    if key != "summary":
        show_columns(key, res, ps)
    if res.feasibility:
        st.markdown("**Does your hybrid plan respect every limit?**")
        for chk in res.feasibility:
            st.markdown(f"{'✔' if chk['ok'] else '✖'} **{chk['label']}** — {chk['detail']}")
        others = {k: SS["pieces"][k]["submitted"]["metrics"][k]["Total Cost"]
                  for k in ("chase", "level") if submitted(k)}
        if "hybrid" in res.metrics and others:
            mine = res.metrics["hybrid"]["Total Cost"]
            st.caption(text.md(f"Your hybrid costs ${mine:,.0f} (if your formulas are right): "
                               + " · ".join(f"{k} ${v:,.0f}" for k, v in others.items())))
    verdict = guided_verdict(scn, ps, res) if guided else None
    show_feedback(key, res, ps)
    if key != "summary" and res.cells:
        st.markdown("**Cell by cell** — every column you built, laid out like your sheet; hover "
                    "a cell for details")
        legend()
        g = checker.grid(res)
        if not g.empty:
            st.markdown(grid_html(res), unsafe_allow_html=True)
        plan = [c for c in res.cells if c.row_label == "Planning"]
        st.markdown("Planning cells: " + " · ".join(
            f"{c.ref} {SYMBOL[c.status]} {escape(c.label)}" for c in plan))
    elif key == "summary":
        md_table(pd.DataFrame([{"Cell": c.ref, "Item": c.label,
                                "Status": checker.STATUS_LABEL[c.status],
                                "Your value": c.shown or "—"} for c in res.cells]))

    improve = (res.complete and guided and verdict is not None and not verdict.good)
    if improve:
        st.info("Every formula is right, but this plan isn't better than your starting plan "
                f"yet (${verdict.cost:,.0f} vs ${verdict.start_cost:,.0f}). Follow the "
                "checklist above, change the green cells, and upload again. You can submit "
                "once your hybrid beats both pure strategies.")
    elif res.complete:
        st.success(f"🎉 Your **{checker.PIECE_TITLES[key].lower()}** is complete and correct.")
        if key == "hybrid" and not guided and "hybrid" in res.metrics:
            best = min((SS["pieces"][k]["submitted"]["metrics"][k]["Total Cost"]
                        for k in ("chase", "level") if submitted(k)), default=None)
            if best is not None and res.metrics["hybrid"]["Total Cost"] >= best - 0.5:
                st.warning("Your hybrid costs no less than your cheaper pure plan. You can "
                           "submit it, but a hybrid should mix the levers to beat both.")
        if st.button(f"Submit the {checker.PIECE_TITLES[key].lower()}", type="primary"):
            submit_piece(key, res)
            st.rerun()
    elif res.feasibility and not res.feasible and all(c.status == "ok" for c in res.graded):
        st.info("Your formulas are all correct, but the plan breaks a limit. Change your "
                "decisions (workers, regular production, overtime, subcontracting) and upload again.")
    nav_buttons(key, next_disabled=True)


# --------------------------------------------------------------------------- #
# Final report
# --------------------------------------------------------------------------- #
def page_report() -> None:
    header("report")
    scn = scenario()
    stu = SS["student"]
    recs = [piece_record(k) for k in PIECES]
    md_table(pd.DataFrame([{"Piece": r["title"], "Status": r["status"],
                            "Uploads": r["attempts"], "First attempt": r["first"] or "—",
                            "Last attempt": r["last"] or "—",
                            "Most help used": (f"{r['help']} · {hints.LEVEL_NAMES[r['help']]}"
                                               if r["help"] else "none"),
                            "Submitted": r["submitted_at"] or "—"} for r in recs]))
    final = all(submitted(k) for k in PIECES)
    summ = SS["pieces"]["summary"]["submitted"] or {}
    if final:
        st.success("Every piece is submitted. Download your final report and submit it with "
                   "your workbook.")
    else:
        st.info("Not every piece is submitted yet. You can download a DRAFT report to show "
                "your progress.")
    if st.button("Create my final report (PDF)" if final else "Create a DRAFT report (PDF)",
                 type="primary" if final else "secondary"):
        hyb = (SS["pieces"]["hybrid"].get("submitted") or {}) if HYBRID_ON else {}
        pdf, rid = _report().build_report(
            scn, student_name=stu["name"], student_id=stu["sid"], identity_label=identity.label,
            pieces=recs, final=final, heading="Final report" if final else "Progress report",
            comparison=summ.get("metrics") or None, feasibility=hyb.get("feasibility") or None,
            recommendation=summ.get("recommendation", ""),
            justification=summ.get("justification") if summ else None,
            words=summ.get("words", 0))
        SS["final_pdf"] = {"pdf": pdf, "rid": rid, "final": final}
        if final and identity.recorded and store.enabled():
            try:
                store.record_completion(GAME, stu["sid"], completion_code=rid, score=None, extra={
                    "seed": scn.seed, "config_hash": scn.config_hash,
                    "recommendation": summ.get("recommendation"),
                    "pieces": {k: {"uploads": r["attempts"], "help": r["help"],
                                   "workbook_sha256": r["sha"]} for k, r in zip(PIECES, recs)},
                    "plan_costs": {k: v["Total Cost"] for k, v in (summ.get("metrics") or {}).items()}})
                SS["final_pdf"]["recorded"] = True
            except Exception as e:                                     # noqa: BLE001
                SS["final_pdf"]["recorded"] = type(e).__name__
        st.rerun()
    fp = SS.get("final_pdf")
    if fp and fp["final"] == final:
        st.download_button("⬇ Download report PDF", fp["pdf"],
                           file_name=_report().report_file_name(stu["name"], "final" if final
                                                             else "draft", fp["rid"]),
                           mime="application/pdf", type="primary")
        if fp.get("recorded") is True:
            st.caption(f"Completion recorded (report {fp['rid']}).")
        elif fp.get("recorded"):
            st.warning(f"The completion could not be recorded on the server ({fp['recorded']}). "
                       "Your PDF is fine — submit it as usual.")
        if final:
            st.markdown("#### Submit to your LMS\n1. **Your Excel workbook** (the version you "
                        "submitted last)\n2. **This report**")
    nav_buttons("report")


step = SS["step"]
if step not in STEP_KEYS or not unlocked(step):
    step = SS["step"] = "start" if not SS["student"] else next(
        (k for k in PIECES if not submitted(k)), "report")
if step == "start":
    page_start()
elif step == "report":
    page_report()
else:
    page_piece(step)
