import re

from aggplan import hints, layout as L
from tests.helpers import scn7

ITEMS = {kind: L.REQUIRED[kind] + [f"plan:{pc.key}" for pc in L.PLANNING[kind]] + ["total"]
         for kind in ("chase", "level", "hybrid")}


def test_every_item_has_four_growing_steps():
    for kind, items in ITEMS.items():
        for item in items:
            steps = hints.ladder(kind, item)
            assert len(steps) == hints.MAX_LEVEL and all(steps), (kind, item)
            for lvl in range(1, 5):
                assert hints.hint_text(kind, item, lvl) == steps[:lvl]
    for key in ("link", "Peak Workforce", "Months With Backorders", "lower", "savings"):
        assert len(hints.ladder("summary", f"summary:{key}")) == 4


def test_first_hint_is_a_question_and_never_contains_a_number():
    for kind, items in ITEMS.items():
        for item in items:
            assert hints.ladder(kind, item)[0].endswith("?"), (kind, item)


def test_hints_never_contain_scenario_numbers():
    """Hints teach the method; the numbers are always the student's to calculate."""
    s = scn7()
    banned = {f"{v:,.0f}" for v in list(s.demand) + [s.total_demand]} | {"426,000", "489,375"}
    for kind, items in ITEMS.items():
        for item in items:
            for step in hints.ladder(kind, item):
                found = set(re.findall(r"\d[\d,]*", step))
                assert not (found & banned), (kind, item, step)


def test_column_hint_ends_with_a_heading():
    assert "Your plan can't answer" in hints.column_hint("chase", "Ending Backorders", 1)
    assert "“Ending Backorders”" in hints.column_hint("chase", "Ending Backorders", 3)
    assert "“Regular Production”" in hints.column_hint("hybrid", "Regular Production", 3)
