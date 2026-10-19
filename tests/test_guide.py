"""The guided hybrid: the right starting point, sensible questions, and a path that really
leads to a plan at least as cheap as both pure strategies."""

import math

import pytest

from aggplan import checker, engine, guide, workbook
from aggplan.scenario import resolve_scenario
from tests.helpers import SID, fixture, scn7

CASES = [({}, s) for s in (7, 1, 99, 2024, 207659)] + \
        [({"safety_stock_policy": "maintain"}, 7), ({"worker_policy": "partial"}, 7),
         ({"scenario_preset": "stability_favored"}, 1),
         ({"scenario_preset": "stability_favored"}, 99)]


def follow(s, start):
    """Do what the walkthrough's Excel task tells the student to do."""
    p, cap = s.params, s.cap_month
    track = guide.track_name(s, start)
    base = guide.starting_decision(s, start)
    if track == "no_gain":
        return base
    if track == "chase":
        W, R, O = list(base.workers), list(base.regular), [0.0] * 12
        for m in guide._round_down_months(s):
            i = m["i"]
            W[i] -= 1
            R[i] = W[i] * cap
            O[i] = m["rem"]
        return engine.Decision(W, R, O, [0.0] * 12)
    # level: use idle capacity, then overtime, then subcontracting, only where short
    w = base.workers[0]
    net = p["beginning_inventory"] - p["beginning_backorders"]
    R, O, S = [], [], []
    for i, d in enumerate(s.demand):
        r = base.regular[i]
        short = max(0.0, d - (net + r))
        add = min(short, w * cap - r)
        r += add
        short -= add
        o = min(short, w * cap * p["overtime_pct"])
        short -= o
        sub = min(short, p["subcontract_capacity"])
        R.append(r), O.append(o), S.append(sub)
        net += r + o + sub - d
    return engine.Decision([w] * 12, R, O, S)


@pytest.mark.parametrize("cfg,seed", CASES)
def test_following_the_walkthrough_reaches_a_good_solution(cfg, seed):
    s = resolve_scenario(cfg, seed)
    start = guide.best_start(s)
    v = guide.evaluate(s, start, engine.simulate(s, follow(s, start)))
    assert v.feasible and v.good, (cfg, seed, start, v)


def test_starting_point_is_the_cheaper_plan():
    s = scn7()
    q = guide.start_question(s)
    assert q.answer == "Chase" and guide.check_answer(q, "Chase")
    assert not guide.check_answer(q, "Level")
    s = resolve_scenario({"scenario_preset": "stability_favored"}, 1)
    assert guide.best_start(s) == "level"


def test_tracks_match_the_policies():
    assert guide.track_name(scn7(), "chase") == "chase"
    assert guide.track_name(scn7(worker_policy="partial"), "chase") == "no_gain"
    assert guide.track_name(scn7(), "level") == "level"


def test_chase_track_questions_use_the_students_numbers():
    s = scn7()
    qs = {q.qid: q for q in guide.steps(s, "chase") if isinstance(q, guide.Question)}
    assert qs["c_idle"].answer == 500 and "April" in qs["c_idle"].prompt
    assert qs["c_otcost"].answer == 2250 and qs["c_compare"].answer == "Overtime instead"
    assert guide.check_answer(qs["c_idle"], 500.2) and not guide.check_answer(qs["c_idle"], 510)
    assert isinstance(guide.steps(s, "chase")[-1], guide.Task)


def test_level_track_months_question():
    s = resolve_scenario({"scenario_preset": "stability_favored"}, 1)
    q = next(q for q in guide.steps(s, "level") if q.qid == "l_months")
    assert guide.check_answer(q, list(q.answer))
    assert not guide.check_answer(q, list(q.answer)[:1])


def test_starting_plan_is_added_to_the_students_workbook():
    s = scn7()
    out = workbook.add_hybrid_start(fixture("mistakes_seed7.xlsx"), s,
                                    guide.starting_decision(s, "chase"))
    book, err = checker.read_book(out)
    assert err is None and not checker.check_identity(book, s, SID)
    res = checker.check_book(book, s, "hybrid", SID)
    assert not res.missing_columns                  # every column is built
    assert res.uncalculated                         # Excel must calculate it once
    assert book.cell("Level Plan", "E3").formula == "=C21"   # the student's own work kept


def test_no_question_states_the_overtime_cost_as_extra_on_top_of_labor():
    for cfg, seed in CASES:
        s = resolve_scenario(cfg, seed)
        for st in guide.steps(s, guide.best_start(s)):
            body = getattr(st, "prompt", "") + getattr(st, "why", "") + getattr(st, "text", "")
            assert "on top of" not in body


def test_unchanged_starting_plan_is_not_yet_good_when_it_can_be_improved():
    s = scn7()
    v = guide.evaluate(s, "chase", engine.reference_plan(s, "chase"))
    assert v.feasible and not v.good and any("starting plan" in n for n in v.notes)
    p = scn7(worker_policy="partial")
    assert guide.evaluate(p, "chase", engine.reference_plan(p, "chase")).good


@pytest.mark.parametrize("seed", [7, 1, 207659])
def test_coach_finishes_the_job_from_the_screenshot_mistake(seed):
    """Workers lowered but Regular Production left as it was (the plan breaks capacity).
    Doing what the coach says — month by month — must end in a good solution."""
    s = resolve_scenario({}, seed)
    cap = s.cap_month
    base = guide.starting_decision(s, "chase")
    qual = {r["i"]: r for r in guide._round_down_months(s)}
    W, R = list(base.workers), list(base.regular)
    for i in qual:
        W[i] -= 1                                        # the mistake
    plan = engine.simulate(s, engine.Decision(W, R, [0.0] * 12, [0.0] * 12))
    fixes, opps = guide.coach(s, "chase", plan, 1)
    assert len(fixes) == len(qual) and not opps
    assert all("Regular Production to" in f and "Overtime to" in f for f in fixes)
    O = [0.0] * 12
    for i in qual:                                       # do what each fix says
        O[i] = R[i] - W[i] * cap
        R[i] = W[i] * cap
    plan = engine.simulate(s, engine.Decision(W, R, O, [0.0] * 12))
    assert guide.coach(s, "chase", plan, 1) == ([], [])
    assert guide.evaluate(s, "chase", plan).good


def test_coach_opportunities_get_more_specific_and_lead_to_the_answer():
    s = resolve_scenario({}, 207659)
    cap = s.cap_month
    base = guide.starting_decision(s, "chase")
    qual = guide._round_down_months(s)
    W, R, O = list(base.workers), list(base.regular), [0.0] * 12
    for q in qual[:1]:                                    # only one month done
        W[q["i"]] -= 1
        R[q["i"]] = W[q["i"]] * cap
        O[q["i"]] = q["rem"]
    plan = engine.simulate(s, engine.Decision(W, R, O, [0.0] * 12))
    v = guide.evaluate(s, "chase", plan)
    assert v.feasible and not v.good
    one, two, three = (guide.coach(s, "chase", plan, k)[1] for k in (1, 2, 3))
    assert "still qualif" in one[0] and "Still to change" in two[0]
    assert len(three) == len(qual) - 1 and "Workers" in three[0]
    for q in qual[1:]:                                     # do what step 3 says
        W[q["i"]] -= 1
        R[q["i"]] = W[q["i"]] * cap
        O[q["i"]] = q["rem"]
    assert guide.evaluate(s, "chase", engine.simulate(s, engine.Decision(W, R, O, [0.0] * 12))).good


def test_level_coach_flags_overbuying():
    s = resolve_scenario({"scenario_preset": "stability_favored"}, 1)
    good = __import__("tests.test_guide", fromlist=["follow"]).follow(s, "level")
    S = [x + (500 if i == 6 else 0) for i, x in enumerate(good.subcontract)]   # 500 too many
    plan = engine.simulate(s, engine.Decision(good.workers, good.regular, good.overtime, S))
    fixes, opps = guide.coach(s, "level", plan, 2)
    assert not fixes and any("July" in o for o in opps)


def test_task_checklist_has_a_worked_example():
    s = scn7()
    task = guide.steps(s, "chase")[-1]
    assert "Worked example" in task.text and "three cells" in task.text
    assert "Regular Production" in task.text and "easy to forget" in task.text
