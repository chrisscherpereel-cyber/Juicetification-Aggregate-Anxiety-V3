"""Drive the Streamlit app headlessly: sign in, then check the pieces open one at a time."""

import os

import pytest

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


def _button(at, label):
    for b in at.button:
        if label in b.label:
            return b
    raise AssertionError(f"no button {label!r}: {[b.label for b in at.button]}")


def test_pieces_unlock_in_order():
    at = st_testing.AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception and at.title[0].value == "Scenario & workbook"
    at.text_input[0].input("Ann Student")
    at.text_input[1].input("S1")
    _button(at, "Start").click()
    at = at.run()
    assert not at.exception
    assert any("Your workbook comes after your columns" in i.value for i in at.info)
    assert _button(at, "Level plan").disabled and _button(at, "Hybrid plan").disabled
    _button(at, "Start the chase plan").click()
    at = at.run()
    assert at.title[0].value == "Chase plan" and not at.exception
    # pretend the chase plan was submitted: the level plan opens, the hybrid stays locked
    at.session_state["pieces"]["chase"]["submitted"] = {"at": "now", "sha": "x", "file": "f",
                                                        "metrics": {}, "feasibility": []}
    at = at.run()
    assert not _button(at, "Level plan").disabled and _button(at, "Hybrid plan").disabled
    _button(at, "Final report").click()
    at = at.run()
    assert at.title[0].value == "Final report" and not at.exception
