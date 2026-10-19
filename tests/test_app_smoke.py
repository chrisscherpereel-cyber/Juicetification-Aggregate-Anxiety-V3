"""Drive the Streamlit app headlessly through its four steps."""

import os

import pytest

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


def _click(at, label):
    for b in at.button:
        if label in b.label:
            b.click()
            return at.run()
    raise AssertionError(f"no button {label!r}: {[b.label for b in at.button]}")


def test_four_steps_navigate():
    at = st_testing.AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    assert at.title[0].value == "Meet your scenario"
    at.text_input[0].input("Ann Student")
    at.text_input[1].input("S1")
    at = _click(at, "Start")
    assert not at.exception and at.session_state["student"]["sid"] == "S1"
    at = _click(at, "Next: get your Excel workbook")
    assert at.title[0].value == "Build it in Excel"
    assert at.session_state["wb_bytes"][:2] == b"PK"
    at = _click(at, "I've saved my workbook")
    assert at.title[0].value == "Upload & check" and not at.exception
    at = _click(at, "Submit")
    assert at.title[0].value == "Submit" and not at.exception
