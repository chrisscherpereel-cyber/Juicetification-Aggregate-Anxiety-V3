"""The stage strip on a plan page: one 'Do this next', sections numbered to match, only the
current section open, and the upload only once the student has a workbook. Plus what the
'Checking' switch tells the student."""

import os

import pytest

from aggplan import columns, text

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")
SUBMITTED = {"at": "t", "sha": "c", "file": "w", "metrics": {}, "feasibility": []}


def page(step="chase", **state):
    at = st_testing.AppTest.from_file(APP, default_timeout=60)
    at.query_params["seed"] = "7"
    at.run()
    at.text_input[0].input("Ann")
    at.text_input[1].input("S1")
    next(b for b in at.button if b.label == "Start").click()
    at.run()
    for k, v in state.items():
        at.session_state[k] = v
    at.session_state["step"] = step
    return at.run()


def strip(at):
    return next(m.value for m in at.markdown if "<div class='aa-stages'>" in m.value)


def next_step(at):
    return next(m.value for m in at.markdown if "<div class='aa-next'>" in m.value)


def expanders(at):
    return {e.label: e.proto.expanded for e in at.expander}


def test_chase_starts_on_the_columns_with_the_upload_closed():
    at = page()
    assert not at.exception
    s = strip(at)
    assert s.index("Find your columns") < s.index("Get your workbook") < s.index("Fill it in") \
        < s.index("Upload, check") < s.index("Submit")
    assert "now'>① Find your columns" in s
    assert "Find your columns" in next_step(at)
    ex = expanders(at)
    assert ex["① Find your columns"] and not ex["② Get your workbook"] \
        and not ex["③ Fill it in, in Excel"]
    assert any(h.value == "④ Upload, check & fix" for h in at.subheader)
    assert any("Upload opens once you have your workbook" in i.value for i in at.info)
    assert not at.get("file_uploader")


def test_with_a_workbook_the_upload_opens_and_the_next_step_is_excel():
    at = page(downloaded_hash="x")
    at.session_state["pieces"]["chase"]["skipped_design"] = True
    at = at.run()
    assert at.get("file_uploader")
    assert "now'>③ Fill it in" in strip(at) and "✔ Get your workbook" in strip(at)
    assert "Chase Plan" in next_step(at) and "planning cells first" in next_step(at)
    assert expanders(at)["③ Fill it in, in Excel"]


def test_level_plan_has_no_workbook_stage_and_its_columns_carry_over():
    at = page(downloaded_hash="x")
    at.session_state["pieces"]["chase"]["found"] = {c: c for c in columns.concepts("chase")}
    at.session_state["pieces"]["chase"]["submitted"] = SUBMITTED
    at.session_state["step"] = "level"
    at = at.run()
    assert not at.exception
    s = strip(at)
    assert "Get your workbook" not in s and "③ Upload, check" in s
    assert "✔ Find your columns" in s and "now'>② Fill it in" in s


def test_checking_switch_changes_what_the_student_is_told():
    at = page(downloaded_hash="x")
    body = " ".join(c.value for c in at.caption)
    assert text.checking_sentence(False).split("—")[0].replace("$", "\\$") in body
    radio = next(r for r in at.sidebar.radio if r.label == "Checking")
    radio.set_value("formulas")
    at = at.run()
    body = " ".join(c.value for c in at.caption)
    assert "values and your formulas" in body
