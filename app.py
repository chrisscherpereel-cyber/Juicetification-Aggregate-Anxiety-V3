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

from aggplan import checker, hints, layout as L, report, text, workbook
from aggplan import identity as ident
from aggplan.recognize import HELPER
from aggplan.scenario import NAME_KEYS, resolve_scenario

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
  .aa-grid thead th {font-weight: 600; vertical-align: bottom; max-width: 7rem;
                     white-space: normal; line-height: 1.1;}
  .aa-grid tbody th {text-align: left; font-weight: 600; white-space: nowrap;}
  .aa-grid td {text-align: center; color: #222; min-width: 2rem; cursor: default;}
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
PIECES = ["chase", "level"] + (["hybrid"] if INCLUDE_HYBRID else []) + ["summary"]
STEPS = [("start", "Scenario & workbook")] + [(k, checker.PIECE_TITLES[k]) for k in PIECES] \
    + [("report", "Final report")]
STEP_KEYS = [k for k, _ in STEPS]
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


SS.setdefault("pieces", {k: new_piece() for k in PIECES})


def seed_for(sid: str) -> int:
    if CTX.get("seed") is not None:
        return int(CTX["seed"])
    return store.derive_seed(GAME, sid, lo=1, hi=10 ** 6)


def scenario():
    stu = SS["student"]
    return resolve_scenario(CFG, seed_for(stu["sid"])) if stu else None


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


# ---- saving (optional Dropbox store; lets a student resume and feeds the roster) ----
def save_state(step: str) -> None:
    stu = SS.get("student")
    if not (identity.recorded and store.enabled() and stu):
        return
    scn = scenario()
    done = sum(submitted(k) for k in PIECES)
    try:
        store.save(GAME, stu["sid"], {"name": stu["name"], "seed": scn.seed,
                                      "config_hash": scn.config_hash, "pieces": SS["pieces"]},
                   progress=done / len(PIECES), step=step)
    except Exception:                                                  # noqa: BLE001
        pass


def restore_state() -> None:
    if SS.get("_restored") or not (identity.recorded and store.enabled() and URL_SID):
        return
    SS["_restored"] = True
    try:
        saved = store.load(GAME, URL_SID) or {}
    except Exception:                                                  # noqa: BLE001
        return
    scn = resolve_scenario(CFG, seed_for(URL_SID))
    if saved.get("config_hash") == scn.config_hash and isinstance(saved.get("pieces"), dict):
        SS["student"] = {"name": saved.get("name") or URL_SID, "sid": URL_SID}
        for k in PIECES:
            SS["pieces"][k] = {**new_piece(), **saved["pieces"].get(k, {})}


restore_state()


# --------------------------------------------------------------------------- #
# Sidebar
# --------------------------------------------------------------------------- #
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


def page_start() -> None:
    header("start")
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
                    save_state("start")
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
    st.caption("Every plan must follow these. They are also on the Scenario sheet of your workbook.")
    show_policies(scn, expanded=True)
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.subheader("Given data")
        md_table(pd.DataFrame(text.param_rows(scn), columns=["Item", "Value", "Unit"]))
    with c2:
        st.subheader("Your workbook")
        if SS.get("wb_for") != (stu["sid"], scn.config_hash):
            SS["wb_bytes"] = workbook.build_workbook(scn, student_name=stu["name"],
                                                     student_id=stu["sid"], game=GAME)
            SS["wb_for"] = (stu["sid"], scn.config_hash)
        st.download_button("⬇ Download my Excel workbook", SS["wb_bytes"],
                           file_name=workbook.file_name(stu["name"], stu["sid"]),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")
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
                        + ", ".join(f"`{n}`" for n, _, _, _ in NAME_KEYS))
    nav_buttons("start", "Start the chase plan →")


# --------------------------------------------------------------------------- #
# Piece pages
# --------------------------------------------------------------------------- #
COLORS = {"ok": "#C6EFCE", "typed": "#FFEB9C", "carried": "#E4DFEC", "wrong": "#FFC7CE",
          "blank": "#F2F2F2", "error": "#F8CBAD", "text": "#F8CBAD", "uncalc": "#F8CBAD",
          "decision": "#DDEBF7"}
SYMBOL = {"ok": "✔", "typed": "✎", "carried": "↪", "wrong": "✖", "blank": "○",
          "error": "⚠", "text": "⚠", "uncalc": "⚠", "decision": "◆"}


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
    head = "".join(f"<th>{escape(c)}</th>" for c in g.columns)
    body = []
    for row in g.index:
        tds = []
        for col in g.columns:
            c = g.at[row, col]
            if not isinstance(c, checker.CellCheck):
                tds.append("<td></td>")
                continue
            tip = escape(f"{c.ref}: {checker.STATUS_LABEL[c.status]}"
                         + (f" (shows {c.shown})" if c.shown else ""))
            tds.append(f"<td title=\"{tip}\" style=\"background:{COLORS[c.status]}\">"
                       f"{SYMBOL[c.status]}</td>")
        body.append(f"<tr><th>{escape(row)}</th>{''.join(tds)}</tr>")
    return (f"<div class='aa-grid'><table><thead><tr><th></th>{head}</tr></thead>"
            f"<tbody>{''.join(body)}</tbody></table></div>")


def legend() -> None:
    st.markdown("<div class='aa-legend'>" + "".join(
        f"<span><span style='background:{COLORS[k]};color:#222;padding:0 .35rem;border-radius:3px'>"
        f"{SYMBOL[k]}</span> {checker.STATUS_LABEL[k].split(' ', 1)[1]}</span>"
        for k in ("ok", "wrong", "carried", "typed", "blank", "error", "decision"))
        + "</div>", unsafe_allow_html=True)


def run_check(key: str, up) -> None:
    data = up.getvalue()
    sha = hashlib.sha256(data).hexdigest()
    ps = SS["pieces"][key]
    if ps["attempts"] and ps["attempts"][-1]["sha"] == sha and key in SS["results"]:
        return
    res = checker.check_piece(data, scenario(), key, expected_sid=SS["student"]["sid"])
    SS["results"][key] = res
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
    first = next((a for a in ps["attempts"] if not a.get("fatal")), None)
    return {"title": checker.PIECE_TITLES[key],
            "status": "Submitted" if sub else ("In progress" if ps["attempts"] else "Not started"),
            "attempts": len([a for a in ps["attempts"] if not a.get("fatal")]),
            "first": f"{first['ok']} / {first['total']} correct" if first else "",
            "submitted_at": sub["at"] if sub else "",
            "help": max(list(ps["help"].values()) + list(ps["colhelp"].values()) + [0]),
            "file": sub["file"] if sub else "", "sha": sub["sha"] if sub else ""}


def submit_piece(key: str, res: checker.PieceResult) -> None:
    ps = SS["pieces"][key]
    last = ps["attempts"][-1]
    ps["submitted"] = {"at": report.now_in(str(scenario().params.get("report_timezone") or ""))
                       .strftime("%Y-%m-%d %H:%M %Z").strip(),
                       "sha": res.sha256, "file": last["file"], "metrics": res.metrics,
                       "feasibility": res.feasibility, "recommendation": res.recommendation,
                       "justification": res.justification, "words": res.words}
    save_state(f"submitted:{key}")


def show_guidance(key: str, scn) -> None:
    ps = SS["pieces"][key]
    st.markdown(f"<div class='aa-card'><b>The goal.</b> {escape(GOALS[key])}</div>",
                unsafe_allow_html=True)
    if key == "summary":
        st.markdown("Fill the **Summary** sheet: each plan's yearly figures, which plan costs "
                    "less, the savings, your recommendation and why. Use formulas that point "
                    "at your plan sheets.")
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
            st.warning(f"Columns {' and '.join(n.letters)} both look like **{n.concept}**. Keep "
                       "one, or give the other a heading that says what makes it different.")
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
    if carried:
        st.caption(f"↪ {carried} more cell{'s' if carried != 1 else ''} only follow an earlier "
                   "mistake: their formulas look right and should correct themselves once the "
                   "cells above are fixed.")
    if todo:
        names = [f"{item_name(key, it, res)} ({len(cs)})" for it, cs in todo.items()]
        st.markdown("**Still empty:** " + text.md(", ".join(names)))


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
        rec = piece_record(key)
        pdf, rid = report.build_report(
            scn, student_name=SS["student"]["name"], student_id=SS["student"]["sid"],
            identity_label=identity.label, pieces=[rec], final=True,
            heading=f"Receipt — {checker.PIECE_TITLES[key]} submitted",
            comparison=sub.get("metrics") or None, feasibility=sub.get("feasibility") or None,
            recommendation=sub.get("recommendation", ""),
            justification=sub.get("justification") if key == "summary" else None,
            words=sub.get("words", 0))
        st.download_button("⬇ Download receipt (PDF)", pdf,
                           file_name=report.report_file_name(SS["student"]["name"], key, rid),
                           mime="application/pdf")
        st.caption("Upload the receipt to your LMS if your instructor collects each piece.")
        nav_buttons(key)
        return

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
    c1.metric("Cells correct", f"{res.n_ok} / {res.n_total}")
    c2.progress(res.n_ok / max(1, res.n_total))
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
    show_feedback(key, res, ps)
    if key != "summary" and not checker.grid(res).empty:
        st.markdown("**Cell by cell** — same layout as your sheet; hover a cell for details")
        legend()
        st.markdown(grid_html(res), unsafe_allow_html=True)
    elif key == "summary":
        md_table(pd.DataFrame([{"Cell": c.ref, "Item": c.label,
                                "Status": checker.STATUS_LABEL[c.status],
                                "Your value": c.shown or "—"} for c in res.cells]))

    if res.complete:
        st.success(f"🎉 Your **{checker.PIECE_TITLES[key].lower()}** is complete and correct.")
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
        hyb = SS["pieces"].get("hybrid", {}).get("submitted") or {}
        pdf, rid = report.build_report(
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
                           file_name=report.report_file_name(stu["name"], "final" if final
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
