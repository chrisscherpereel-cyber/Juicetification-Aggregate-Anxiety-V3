"""The 'Design your columns' panel in the app: one heading at a time, help that grows with
each miss, and a column's name and reason shown only once it is found."""

import os

import pytest

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


def chase_page():
    at = st_testing.AppTest.from_file(APP, default_timeout=60)
    at.query_params["seed"] = "7"
    at.run()
    at.text_input[0].input("Ann")
    at.text_input[1].input("S1")
    next(b for b in at.button if b.label == "Start").click()
    at.run()
    at.session_state["step"] = "chase"
    return at.run()


def check(at, heading):
    next(t for t in at.text_input if t.label == "Check one heading").input(heading)
    next(b for b in at.button if b.label == "Check").click()
    return at.run()


def texts(at):
    return " ".join([m.value for m in at.markdown] + [m.value for m in at.caption]
                    + [m.value for m in at.success] + [m.value for m in at.warning]
                    + [m.value for m in at.info])


def test_nothing_is_given_away_before_it_is_found():
    at = chase_page()
    assert not at.exception
    body = texts(at)
    assert "How many bottles are on the shelf when the month starts?" in body
    assert "Beginning Inventory" not in body and "Ending Backorders" not in body
    assert "0 of 14 columns found" in " ".join(p.value.get("text", "") if isinstance(p.value, dict)
                                                else str(p.proto.text) for p in at.get("progress"))


def test_misses_get_more_help_and_a_find_explains_why():
    at = chase_page()
    at = check(at, "stuff")
    w1 = at.warning[-1].value
    assert "isn't recognised" in w1 and "Beginning" not in w1
    at = check(at, "more stuff")
    assert "open questions" in at.warning[-1].value
    assert any("Try this one" in c.value for c in at.caption)
    at = check(at, "Starting Stock")
    ok = at.success[-1].value
    assert "Beginning Inventory" in ok and "Why it matters" in ok
    assert at.session_state["pieces"]["chase"]["found"] == {"Beginning Inventory": "Starting Stock"}
    assert at.session_state["pieces"]["chase"]["head_misses"] == 2
    assert "1 of 14 columns found" in str(at.get("progress")[0].proto)
    at = check(at, "Opening inventory")
    assert "same question" in at.info[-1].value
    assert "How many bottles are on the shelf when the month starts?" not in texts(at)
