"""
Juicetification: Aggregate Anxiety — Excel edition (V3)

Students download a personal Excel workbook, build a chase plan and a level plan (and an
optional hybrid) with their own Excel formulas, upload the file, and get cell-by-cell
feedback until everything checks out. Then they download a PDF check report and submit
it with the workbook.

Four steps, always visible in the sidebar:
  1 Meet your scenario   2 Build it in Excel   3 Upload & check   4 Submit

Run:  streamlit run app.py        (see README.md for instructor configuration)
"""

from __future__ import annotations

import datetime as dt
import hashlib

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import student_store as store
from juice_director import resolve_config, serve_manifest_if_requested
from manifest import MANIFEST

from aggplan import checker, layout as L, report, text, workbook
from aggplan import identity as ident
from aggplan.scenario import NAME_KEYS, resolve_scenario

st.set_page_config(page_title="Aggregate Anxiety — Excel", page_icon="🧃", layout="wide")
SS = st.session_state
serve_manifest_if_requested(MANIFEST)
CFG, CTX = resolve_config(MANIFEST)

STEPS = [("scenario", "Meet your scenario"), ("build", "Build it in Excel"),
         ("check", "Upload & check"), ("submit", "Submit")]
STEP_KEYS = [k for k, _ in STEPS]

st.markdown("""
<style>
  div[data-testid="stSidebarContent"] .stButton button {justify-content: flex-start; text-align: left;}
  .aa-card {border: 1px solid rgba(128,128,128,.35); border-radius: 10px; padding: .8rem 1rem; margin-bottom: .6rem;}
  .aa-step {font-size: .85rem; opacity: .75; margin-bottom: -.4rem;}
  .aa-legend span {display: inline-block; margin: 0 .9rem .3rem 0; white-space: nowrap;}
  .aa-grid {overflow-x: auto;}
  .aa-grid table {border-collapse: collapse; font-size: .8rem; margin-bottom: .8rem;}
  .aa-grid th, .aa-grid td {border: 1px solid rgba(128,128,128,.45); padding: .2rem .35rem;}
  .aa-grid thead th {font-weight: 600; vertical-align: bottom; max-width: 6.5rem;
                     white-space: normal; line-height: 1.1;}
  .aa-grid tbody th {text-align: left; font-weight: 600; white-space: nowrap;}
  .aa-grid td {text-align: center; color: #222; min-width: 2rem; cursor: default;}
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

SS.setdefault("step", "scenario")
SS.setdefault("student", None)          # {"name", "sid"}
SS.setdefault("history", [])            # one entry per check


def seed_for(sid: str) -> int:
    if CTX.get("seed") is not None:
        return int(CTX["seed"])
    return store.derive_seed(GAME, sid, lo=1, hi=10 ** 6)


def scenario():
    stu = SS["student"]
    return resolve_scenario(CFG, seed_for(stu["sid"])) if stu else None


def go(step: str) -> None:
    SS["step"] = step


def save_progress(step: str, fraction: float) -> None:
    """Best effort: lets the Director's roster show where each student is."""
    stu = SS.get("student")
    if not (identity.recorded and store.enabled() and stu):
        return
    try:
        store.save(GAME, stu["sid"], {"name": stu["name"]}, progress=fraction, step=step)
    except Exception:                                                  # noqa: BLE001
        pass


def latest():
    return SS["history"][-1] if SS["history"] else None


def step_done(key: str) -> bool:
    last = latest()
    return {"scenario": SS["student"] is not None,
            "build": bool(SS.get("downloaded")) or last is not None,
            "check": bool(last and last["result"].complete),
            "submit": bool(SS.get("report") and SS["report"].get("final"))}[key]


# --------------------------------------------------------------------------- #
# Sidebar: who you are + the four steps
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.markdown("## 🧃 Aggregate Anxiety")
    st.caption("Aggregate planning, calculated in Excel")
    stu = SS["student"]
    if stu:
        scn_ = scenario()
        st.markdown(f"**{stu['name']}**  \nID {stu['sid']} · Scenario #{scn_.seed}")
        st.caption(identity.label)
    st.markdown("---")
    for n, (key, label) in enumerate(STEPS, start=1):
        mark = "✔" if step_done(key) else f"{n}"
        st.button(f"{mark}  {label}", key=f"nav_{key}", width="stretch",
                  type="primary" if SS["step"] == key else "secondary",
                  disabled=(key != "scenario" and not SS["student"] and key != "check"),
                  on_click=go, args=(key,))
    last = latest()
    if last and not last["result"].fatal:
        r = last["result"]
        req = [s for s in r.sheets.values() if s.required]
        ok = sum(s.n_ok for s in req)
        tot = sum(s.n_total for s in req) or 1
        st.markdown("---")
        st.caption(f"Last check: {ok} of {tot} required cells correct")
        st.progress(ok / tot)
    st.markdown("---")
    with st.expander("How this works"):
        st.write("1. Get your personal scenario.  \n2. Download your workbook and do every "
                 "calculation in Excel.  \n3. Upload it; fix what is flagged; upload again "
                 "as often as you like.  \n4. Download your check report and submit it with "
                 "your workbook.")


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
        c3.button(next_label or f"{STEPS[i + 1][1]} →", type="primary", width="stretch",
                  on_click=go, args=(STEP_KEYS[i + 1],), disabled=next_disabled,
                  key=f"next_{key}")


# --------------------------------------------------------------------------- #
# Step 1: scenario
# --------------------------------------------------------------------------- #
def demand_chart(scn) -> alt.Chart:
    df = pd.DataFrame({"Month": [m[:3] for m in scn.months], "Demand": scn.demand})
    return (alt.Chart(df).mark_bar(color="#F28C28").encode(
        x=alt.X("Month", sort=None, title=None),
        y=alt.Y("Demand", title="bottles"),
        tooltip=["Month", alt.Tooltip("Demand", format=",.0f")]).properties(height=260))


def page_scenario() -> None:
    header("scenario")
    if not SS["student"]:
        st.write("You are the operations planner at **Juicetification Inc.**, a juice bottler "
                 "with a strong summer peak. You will plan a year of production in Excel. "
                 "First, tell us who you are — your student ID gives you your own scenario.")
        with st.form("who"):
            name = st.text_input("Your name")
            sid = st.text_input("Student ID", value=URL_SID or "", disabled=bool(URL_SID),
                                help="Your scenario is made from this ID, so use the same ID "
                                     "every time.")
            if st.form_submit_button("Start", type="primary"):
                sid = (URL_SID or sid).strip()
                if not name.strip() or not sid:
                    st.error("Enter your name and your student ID.")
                else:
                    SS["student"] = {"name": name.strip(), "sid": sid}
                    if not URL_SID and not PRACTICE:
                        store.set_student_id(sid)   # survives a refresh; enables recording
                    save_progress("scenario", 0.05)
                    st.rerun()
        st.info("**Coming back?** If you already have your workbook, go straight to "
                "**Upload & check** in the sidebar — the app recognises your file.")
        return

    scn = scenario()
    for note in scn.notices:
        st.warning(note)
    st.write("Demand for bottled juice swings with the seasons. Your job is to decide how many "
             "workers to employ and how much to produce each month so that the year's demand "
             "is met at the lowest total cost. You will build and compare two classic plans:")
    c1, c2 = st.columns(2)
    c1.markdown("<div class='aa-card'><b>Chase plan</b><br>Change the workforce every month so "
                "production follows demand. Little inventory, lots of hiring and layoffs."
                "</div>", unsafe_allow_html=True)
    c2.markdown("<div class='aa-card'><b>Level plan</b><br>Keep one steady workforce and "
                "production rate all year. Inventory builds in slow months and is used up "
                "in busy ones.</div>", unsafe_allow_html=True)
    st.subheader("Your demand forecast")
    m = st.columns(4)
    m[0].metric("Annual demand", f"{scn.total_demand:,.0f}")
    m[1].metric("Peak month", f"{scn.months[scn.demand.index(max(scn.demand))]}")
    m[2].metric("Peak demand", f"{max(scn.demand):,.0f}")
    m[3].metric("Lowest demand", f"{min(scn.demand):,.0f}")
    st.altair_chart(demand_chart(scn), use_container_width=True)
    st.subheader("Given data")
    st.caption(text.md(text.target_sentence(scn)))
    rows = pd.DataFrame(text.param_rows(scn), columns=["Item", "Value", "Unit"])
    md_table(rows)
    nav_buttons("scenario", "Next: get your Excel workbook →")


# --------------------------------------------------------------------------- #
# Step 2: build in Excel
# --------------------------------------------------------------------------- #
def md_table(df: pd.DataFrame) -> None:
    """A wrapped, full-width table for text-heavy content (st.dataframe truncates text)."""
    esc = lambda x: text.md(str(x)).replace("|", "\\|")                # noqa: E731
    lines = ["| " + " | ".join(df.columns) + " |", "|" + "---|" * len(df.columns)]
    lines += ["| " + " | ".join(esc(v) for v in row) + " |" for row in df.itertuples(index=False)]
    st.markdown("\n".join(lines))


def rules_table(kind: str) -> pd.DataFrame:
    rules = text.column_rules(kind)
    rows = []
    for pc in L.PLANNING[kind][2]:
        rows.append({"Cell": L.planning_ref(kind, pc.key), "What": pc.label,
                     "How to calculate": text.PLANNING_RULES[pc.key]})
    for j, col in enumerate(L.PLAN_COLUMNS[kind]):
        if col.kind == "given":
            continue
        letter = chr(65 + j)
        rows.append({"Cell": f"{letter}2:{letter}13", "What": col.heading.rstrip(" ✎"),
                     "How to calculate": rules.get(col.engine, "")})
    rows.append({"Cell": "row 14", "What": "Totals", "How to calculate": text.TOTAL_RULE})
    return pd.DataFrame(rows)


def page_build() -> None:
    header("build")
    scn = scenario()
    stu = SS["student"]
    if SS.get("wb_for") != (stu["sid"], scn.config_hash):
        SS["wb_bytes"] = workbook.build_workbook(scn, student_name=stu["name"],
                                                 student_id=stu["sid"], game=GAME)
        SS["wb_for"] = (stu["sid"], scn.config_hash)

    c1, c2 = st.columns([1, 1.4])
    with c1:
        st.markdown("<div class='aa-card'><b>Your personal workbook</b><br>Your demand and "
                    "numbers are already in it. Every calculation is yours to build with "
                    "Excel formulas.</div>", unsafe_allow_html=True)

        def _downloaded():
            SS["downloaded"] = True
            save_progress("build", 0.15)

        st.download_button("⬇ Download my Excel workbook", SS["wb_bytes"],
                           file_name=workbook.file_name(stu["name"], stu["sid"]),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", on_click=_downloaded)
        st.caption("Works in Excel (Windows/Mac/web) and Google Sheets (then File → Download → "
                   "Microsoft Excel .xlsx).")
    with c2:
        st.markdown("**Colour key in the workbook**")
        st.markdown("🟨 Given data — don't change  \n🟦 Your formula goes here  \n"
                    "🟩 Your decision (Hybrid Plan only)  \n🔷 Totals row — your SUM formula")

    st.subheader("Work through the sheets in this order")
    st.caption("The same instructions are on the workbook's Start Here sheet and in the "
               "comment on every column heading.")
    tabs = st.tabs(["1 · Chase Plan", "2 · Level Plan", "3 · Summary"]
                   + (["4 · Hybrid (optional)"] if scn.params.get("include_hybrid", True) else [])
                   + ["Excel tips"])
    with tabs[0]:
        st.write("Make exactly what each month needs, changing the workforce to match.")
        st.markdown("1. Fill the **planning cells** S2–S3.\n2. Fill **January** (row 2) — it "
                    "uses BeginInv and StartWorkers.\n3. Fill **February** (row 3) — it uses "
                    "January's results.\n4. Select February's formulas and **fill down** to "
                    "December.\n5. Add the **totals** in row 14.")
        md_table(rules_table("chase"))
    with tabs[1]:
        st.write("Same production and workforce every month; inventory absorbs the swings.")
        st.markdown("1. Fill the **planning cells** S2–S6 (capacity, target, total demand, "
                    "level production, level workforce).\n2. Point every month's Production "
                    "and Workers at those cells — with **$ signs** (=$S$5) so they stay put when "
                    "you fill down.\n3. The rest works exactly like the chase plan.")
        md_table(rules_table("level"))
    with tabs[2]:
        st.write("Compare the plans and make a recommendation.")
        rows = [{"Rows": "Costs and totals", "How": text.SUMMARY_RULES["link"]},
                {"Rows": "Largest workforce / inventory", "How": "MAX of the plan's column."},
                {"Rows": "Months with backorders", "How": text.SUMMARY_RULES["Months With Backorders"]},
                {"Rows": f"Lower-cost plan ({L.LOWER_COST_CELL})", "How": text.SUMMARY_RULES["lower"]},
                {"Rows": f"Savings ({L.SAVINGS_CELL})", "How": text.SUMMARY_RULES["savings"]},
                {"Rows": f"Recommendation ({L.RECOMMEND_CELL})", "How": "Pick from the drop-down."},
                {"Rows": "Justification", "How": "Your own words, using your numbers: cost, "
                 "workforce changes, inventory, backorders and risk."}]
        md_table(pd.DataFrame(rows))
    i = 3
    if scn.params.get("include_hybrid", True):
        with tabs[3]:
            st.write("Optional challenge: mix strategies to beat both plans. The green cells "
                     "(Workers, Regular Production, Overtime, Subcontract) are your decisions; "
                     "calculate everything else with formulas, as before.")
            md_table(rules_table("hybrid"))
            st.caption("The checker also tests your plan against the limits: capacity, "
                       "overtime, subcontracting, storage, service level and December "
                       "backorders.")
        i = 4
    with tabs[i]:
        for t, d in text.EXCEL_TIPS:
            st.markdown(f"**{t}.** {text.md(d)}")
        st.markdown("**Functions you'll use:** " + ", ".join(f"`{f}`" for f, _, _ in text.EXCEL_FUNCTIONS))
        st.markdown("**Names for the given data:** " + ", ".join(f"`{n}`" for n, _, _, _ in NAME_KEYS))
    nav_buttons("build", "I've saved my workbook — check it →")


# --------------------------------------------------------------------------- #
# Step 3: upload & check
# --------------------------------------------------------------------------- #
COLORS = {"ok": "#C6EFCE", "typed": "#FFEB9C", "carried": "#E4DFEC", "wrong": "#FFC7CE",
          "blank": "#F2F2F2", "error": "#F8CBAD", "text": "#F8CBAD", "uncalc": "#F8CBAD",
          "decision": "#DDEBF7"}
SYMBOL = {"ok": "✔", "typed": "✎", "carried": "↪", "wrong": "✖", "blank": "○",
          "error": "⚠", "text": "⚠", "uncalc": "⚠", "decision": "◆", "": ""}


def grid_html(sheet: checker.SheetResult) -> str:
    """The monthly grid as a plain HTML table, laid out like the Excel sheet. Each cell shows
    a symbol on a colour (never colour alone) and its feedback as a hover tooltip."""
    from html import escape
    cells = {(c.row_label, c.group): c for c in sheet.cells
             if c.row_label != "Planning" and c.sheet == sheet.title}
    g = checker.grid(sheet)
    head = "".join(f"<th>{escape(c.rstrip(' ✎'))}</th>" for c in g.columns)
    body = []
    for row in g.index:
        tds = []
        for col in g.columns:
            c = cells.get((row, col))
            if c is None:
                tds.append("<td></td>")
                continue
            tip = escape(f"{c.ref}: {checker.STATUS_LABEL[c.status]}"
                         + (f" (shows {c.shown})" if c.shown else "")
                         + (f" — {c.message}" if c.message else ""))
            tds.append(f"<td title=\"{tip}\" style=\"background:{COLORS[c.status]}\">"
                       f"{SYMBOL[c.status]}</td>")
        body.append(f"<tr><th>{escape(row)}</th>{''.join(tds)}</tr>")
    return (f"<div class='aa-grid'><table><thead><tr><th></th>{head}</tr></thead>"
            f"<tbody>{''.join(body)}</tbody></table></div>")


def cell_table(cells) -> None:
    md_table(pd.DataFrame([{"Cell": c.ref, "Item": c.label,
                            "Status": checker.STATUS_LABEL[c.status],
                            "Your value": c.shown or "—", "Feedback": c.message or ""}
                           for c in cells]))


def legend() -> None:
    st.markdown("<div class='aa-legend'>" + "".join(
        f"<span><span style='background:{COLORS[k]};color:#222;padding:0 .35rem;border-radius:3px'>"
        f"{SYMBOL[k]}</span> {checker.STATUS_LABEL[k].split(' ', 1)[1]}</span>"
        for k in ("ok", "wrong", "carried", "typed", "blank", "error", "decision"))
        + "</div>", unsafe_allow_html=True)


def adopt_student_from(book: checker.Book) -> bool:
    """Coming back with only the workbook: take name/ID from it (unless the URL fixes the ID)."""
    sid = book.meta.get("student_id", "").strip()
    if not sid or (URL_SID and sid != URL_SID):
        return False
    SS["student"] = {"name": book.meta.get("student_name", "").strip() or sid, "sid": sid}
    return True


def run_check(up) -> None:
    data = up.getvalue()
    sha = hashlib.sha256(data).hexdigest()
    if latest() and latest()["sha"] == sha:
        return
    if not SS["student"]:
        book, err = checker.read_book(data)
        if err or not adopt_student_from(book):
            st.error(err or "Start on **Meet your scenario** first (enter your name and ID), "
                            "then upload.")
            return
    scn = scenario()
    res = checker.check_workbook(data, scn, expected_sid=SS["student"]["sid"])
    SS["history"].append({"sha": sha, "result": res, "file": up.name,
                          "at": dt.datetime.now().strftime("%H:%M:%S")})
    SS.pop("report", None)
    if not res.fatal:
        req = [s for s in res.sheets.values() if s.required]
        frac = sum(s.n_ok for s in req) / max(1, sum(s.n_total for s in req))
        save_progress("complete" if res.complete else "check", 0.2 + 0.7 * frac)
    st.rerun()                      # refresh the sidebar (student card, progress)


def page_check() -> None:
    header("check")
    st.write("Save your workbook in Excel, then drop it here. Fix what is flagged, save, and "
             "upload again — as many times as you need.")
    up = st.file_uploader("Your completed workbook (.xlsx)", type=["xlsx"], key="upload")
    if up is not None:
        run_check(up)
    last = latest()
    if not last:
        st.info("No workbook checked yet.")
        nav_buttons("check", next_disabled=True)
        return
    res: checker.CheckResult = last["result"]
    st.caption(f"Checked **{last['file']}** at {last['at']} · check #{len(SS['history'])}")
    if res.fatal:
        for msg in res.fatal:
            st.error(msg)
        nav_buttons("check", next_disabled=True)
        return
    if res.uncalculated:
        st.warning("**Some formulas have no saved results.** This happens when a file was "
                   "never opened and saved in a spreadsheet program. Open it in Excel, save "
                   "it as .xlsx, and upload again.")
    if res.complete:
        st.success("🎉 **Everything checks out.** Every required cell is correct and calculated "
                   "by a formula. Go on to **Submit**.")

    cols = st.columns(len(res.sheets))
    for c, s in zip(cols, res.sheets.values()):
        with c:
            with st.container(border=True):
                if not s.attempted:
                    st.markdown(f"**{s.title}**")
                    st.caption("Optional · not started")
                    continue
                badge = "✔ " if s.complete else ""
                st.markdown(f"**{badge}{s.title}**" + ("" if s.required else " *(optional)*"))
                st.progress(s.n_ok / max(1, s.n_total))
                st.caption(f"{s.n_ok} of {s.n_total} correct")

    todo, carried = res.fix_first(8)
    if todo or carried:
        st.subheader("Fix these next")
        for k, (c, more) in enumerate(todo, start=1):
            extra = f" — and {more - 1} more like it in this column" if more > 1 else ""
            st.markdown(f"**{k}. {c.sheet} · {c.ref}** ({c.label}) — "
                        f"{checker.STATUS_LABEL[c.status]}{extra}  \n"
                        f"<span style='opacity:.85'>{text.md(c.message)}</span>",
                        unsafe_allow_html=True)
        if carried:
            st.caption(f"↪ {carried} other cell{'s' if carried != 1 else ''} only follow"
                       f"{'' if carried != 1 else 's'} these mistakes: their formulas look right "
                       "and should correct themselves once the cells above are fixed.")

    st.subheader("Cell by cell")
    legend()
    shown = [s for s in res.sheets.values() if s.attempted]
    tabs = st.tabs([s.title for s in shown])
    for tab, s in zip(tabs, shown):
        with tab:
            for note in s.notes:
                st.info(note)
            if s.key == "summary":
                cell_table(s.cells)
                continue
            plan = [c for c in s.cells if c.row_label == "Planning"]
            if plan:
                st.markdown("**Planning cells**")
                cell_table(plan)
            if any(c.row_label != "Planning" and c.sheet == s.title for c in s.cells):
                st.markdown("**Monthly grid** — same layout as your sheet; hover a cell for "
                            "its feedback")
                st.markdown(grid_html(s), unsafe_allow_html=True)
            bad = [c for c in s.cells if c.needs_work and c.row_label != "Planning"]
            if bad:
                with st.expander(f"Feedback for the {len(bad)} cells that need work"):
                    cell_table(bad)
            if s.key == "hybrid" and res.hybrid_feasibility:
                st.markdown("**Is your hybrid plan feasible?**")
                for chk in res.hybrid_feasibility:
                    st.markdown(f"{'✔' if chk['ok'] else '✖'} **{chk['label']}** — {chk['detail']}")
                if "hybrid" in res.plan_totals:
                    t = {k: v["Total Cost"] for k, v in res.plan_totals.items()}
                    st.caption("Total cost from your decisions: " + " · ".join(
                        f"{k.title()} ${v:,.0f}" for k, v in t.items()).replace("$", "\\$"))
    nav_buttons("check", "Next: submit →", next_disabled=not res.complete)
    if not res.complete:
        st.caption("Submit unlocks when every required cell is correct. You can still download a "
                   "draft report on the Submit step.")


# --------------------------------------------------------------------------- #
# Step 4: submit
# --------------------------------------------------------------------------- #
def page_submit() -> None:
    header("submit")
    last = latest()
    if not last or last["result"].fatal:
        st.info("Upload your workbook on **Upload & check** first.")
        nav_buttons("submit")
        return
    res: checker.CheckResult = last["result"]
    scn = scenario()
    stu = SS["student"]
    if res.complete:
        st.success("Your workbook is complete and correct.")
        kinds = [k for k in ("chase", "level", "hybrid") if k in res.plan_totals
                 and (k != "hybrid" or res.sheets["hybrid"].complete)]
        rows = {}
        for m in ("Total Cost", "Labor Cost", "Hiring Cost", "Layoff Cost", "Holding Cost",
                  "Backorder Cost", "Total Hires", "Total Layoffs", "Peak Inventory",
                  "Months With Backorders"):
            rows[m] = [(f"${res.plan_totals[k][m]:,.0f}" if "Cost" in m
                        else f"{res.plan_totals[k][m]:,.0f}") for k in kinds]
        md_table(pd.DataFrame({"Measure": list(rows), **{k.title(): [rows[m][j] for m in rows]
                                                        for j, k in enumerate(kinds)}}))
        st.markdown(f"**Your recommendation:** {res.recommendation} · {res.words} words of "
                    "justification")
    else:
        st.warning("Your workbook isn't complete yet:\n\n" + "\n".join(f"- {m}" for m in res.missing()))
        st.caption("You can download a DRAFT report to show your progress, but submit the final "
                   "report once everything checks out.")

    rep = SS.get("report")
    label = "Create my check report (PDF)" if res.complete else "Create a DRAFT report (PDF)"
    if st.button(label, type="primary" if res.complete else "secondary"):
        pdf, attempt = report.build_report(res, scn, student_name=stu["name"],
                                           student_id=stu["sid"], identity_label=identity.label,
                                           file_name=last["file"])
        rep = {"pdf": pdf, "attempt": attempt, "final": res.complete, "sha": res.sha256}
        SS["report"] = rep
        if res.complete and identity.recorded and store.enabled():
            try:
                store.record_completion(GAME, stu["sid"], completion_code=attempt, score=None,
                                        extra={"workbook_sha256": res.sha256,
                                               "seed": scn.seed, "config_hash": scn.config_hash,
                                               "recommendation": res.recommendation,
                                               "plan_costs": {k: v["Total Cost"] for k, v in
                                                              res.plan_totals.items()}})
                rep["recorded"] = True
            except Exception as e:                                     # noqa: BLE001
                rep["recorded"] = False
                rep["record_error"] = type(e).__name__
        st.rerun()                  # refresh the sidebar tick
    if rep and rep.get("sha") == res.sha256:
        st.download_button("⬇ Download report PDF", rep["pdf"],
                           file_name=report.report_file_name(stu["name"], rep["attempt"]),
                           mime="application/pdf", type="primary")
        if rep.get("recorded"):
            st.caption(f"Completion recorded (attempt {rep['attempt']}).")
        elif rep.get("recorded") is False:
            st.warning("The completion could not be recorded on the server "
                       f"({rep.get('record_error')}). Your PDF is fine — submit it as usual.")
        if rep["final"]:
            st.markdown("#### Submit both files to your LMS")
            st.markdown(f"1. **Your Excel workbook** — the same file you checked: `{last['file']}`  \n"
                        f"2. **This report** — attempt `{rep['attempt']}`")
            st.caption("The report records the workbook's fingerprint (SHA-256), so your "
                       "instructor can confirm the two files belong together.")
    nav_buttons("submit")


PAGES = {"scenario": page_scenario, "build": page_build, "check": page_check,
         "submit": page_submit}
step = SS["step"]
if step not in ("scenario", "check") and not SS["student"]:
    step = SS["step"] = "scenario"
PAGES[step]()
