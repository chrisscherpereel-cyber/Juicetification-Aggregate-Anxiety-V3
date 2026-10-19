import math

from aggplan import engine
from aggplan.scenario import generate_forecast, resolve_scenario
from tests.helpers import scn7


def test_seed7_reference_totals_match_excel():
    """$426,000 / $489,375 are the totals Excel 16 calculated from the answer-key formulas."""
    s = scn7()
    assert engine.reference_plan(s, "chase")["Total Cost"].sum() == 426000
    assert engine.reference_plan(s, "level")["Total Cost"].sum() == 489375


def test_forecast_is_reproducible_and_round():
    a, b = generate_forecast(123), generate_forecast(123)
    assert a == b and len(a) == 12
    assert all(x % 500 == 0 for x in a) and sum(a) % 6000 == 0


def test_inventory_and_backorders_conserve():
    for seed in (1, 7, 99, 2024):
        s = resolve_scenario({}, seed)
        for kind in ("chase", "level"):
            df = engine.reference_plan(s, kind)
            for i in range(12):
                r = df.iloc[i]
                net_end = r["Ending Inventory"] - r["Ending Backorders"]
                net_beg = r["Beginning Inventory"] - r["Beginning Backorders"]
                assert math.isclose(net_end, net_beg + r["Total Production"] - r["Demand"])
                assert r["Ending Inventory"] >= 0 and r["Ending Backorders"] >= 0
                assert not (r["Ending Inventory"] > 0 and r["Ending Backorders"] > 0)


def test_chase_meets_target_with_whole_workers():
    s = resolve_scenario({"safety_stock_policy": "maintain"}, 7)
    df = engine.reference_plan(s, "chase")
    assert (df["Ending Inventory"] == s.target).all()
    assert (df["Regular Production"] <= df["Regular Capacity"] + 1e-9).all()
    assert all(float(w).is_integer() for w in df["Workers"])


def test_level_plan_constant_and_hits_december_target():
    for policy in ("none", "maintain"):
        s = resolve_scenario({"safety_stock_policy": policy}, 7)
        df = engine.reference_plan(s, "level")
        assert df["Regular Production"].nunique() == 1 and df["Workers"].nunique() == 1
        net_dec = df["Ending Inventory"].iloc[-1] - df["Ending Backorders"].iloc[-1]
        assert math.isclose(net_dec, s.target)
        assert df["Hires"].iloc[1:].sum() == 0 and df["Layoffs"].iloc[1:].sum() == 0


def test_feasibility_flags_overtime_with_idle_capacity():
    s = scn7()
    dec = engine.Decision([12] * 12, [10000] * 12, [500] + [0] * 11, [0] * 12)
    checks = {c["label"]: c["ok"] for c in engine.check_feasibility(engine.simulate(s, dec), s)}
    assert checks["Overtime only when regular capacity is fully used"] is False


def test_preset_respects_instructor_values():
    s = resolve_scenario({"scenario_preset": "stability_favored", "hiring_cost": 999}, 7)
    assert s["hiring_cost"] == 999 and s["layoff_cost"] == 2700


def test_bad_demand_config_is_ignored_with_notice():
    s = resolve_scenario({"forecast_demand": [1, 2, 3]}, 7)
    assert s.notices and len(s.demand) == 12
