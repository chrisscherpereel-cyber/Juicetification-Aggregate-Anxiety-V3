"""No work is lost: moving between screens, or signing in again on another computer.

The server store (Dropbox in production) is replaced by an in-memory dict; everything else —
the app, its background saver, restore — is the real code."""

import copy
import json
import os
import time

import pytest

import student_store
from aggplan import persistence, saving
from tests.helpers import fixture

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


class FakeStore:
    def __init__(self):
        self.data, self.saves, self.fail_load = {}, [], False

    def save(self, game, sid, state, progress=None, step=None, score=None):
        self.saves.append(copy.deepcopy(state))
        self.data[(game, sid)] = copy.deepcopy(state)
        return True

    def load(self, game, sid):
        if self.fail_load:
            raise ConnectionError("network down")
        return copy.deepcopy(self.data.get((game, sid), {}))


@pytest.fixture
def server(monkeypatch):
    fake = FakeStore()
    monkeypatch.setattr(student_store, "enabled", lambda: True)
    monkeypatch.setattr(student_store, "save", fake.save)
    monkeypatch.setattr(student_store, "load", fake.load)
    return fake


def wait_for_saves():
    for _ in range(200):
        if saving.pending() == 0:
            return
        time.sleep(0.02)
    raise AssertionError("background saves did not finish")


def computer(sid=None, seed=7):
    """A fresh browser session (another computer)."""
    at = st_testing.AppTest.from_file(APP, default_timeout=60)
    at.query_params["seed"] = str(seed)
    if sid:
        at.query_params["sid"] = sid
    return at.run()


def sign_in(at, name="Ann", sid="S1"):
    at.text_input[0].input(name)
    if sid is not None and not at.text_input[1].disabled:
        at.text_input[1].input(sid)
    next(b for b in at.button if b.label == "Start").click()
    return at.run()


def do_some_work(at):
    """Progress like a real student's: chase submitted, level uploaded twice with hints,
    guided hybrid answers — and the latest uploaded workbook."""
    ps = at.session_state["pieces"]
    ps["chase"]["attempts"] = [{"at": "t1", "sha": "a", "file": "wb.xlsx", "ok": 150,
                                "total": 179, "complete": False, "fatal": False}]
    ps["chase"]["submitted"] = {"at": "t2", "sha": "a", "file": "wb.xlsx", "metrics": {},
                                "feasibility": []}
    ps["level"]["attempts"] = [{"at": "t3", "sha": "b", "file": "wb.xlsx", "ok": 100,
                                "total": 182, "complete": False, "fatal": False}] * 2
    ps["level"]["help"] = {"Regular Production": 2}
    ps["hybrid"]["guide"] = {"start": "chase", "answers": {"start": {"tries": 1, "ok": True}}}
    at.session_state["last_wb"] = persistence.workbook_entry("wb.xlsx", fixture("key_seed7.xlsx"))
    return at


def click(at, label):
    next(b for b in at.sidebar.button if label in b.label).click()
    return at.run()


# --------------------------------------------------------------------------------------
def test_moving_between_screens_keeps_everything():
    at = sign_in(computer())
    do_some_work(at)
    before = copy.deepcopy(dict(at.session_state["pieces"]))
    from aggplan import checker
    from tests.helpers import scn7
    feedback = checker.check_piece(fixture("mistakes_seed7.xlsx"), scn7(), "level", "S1")
    at.session_state["results"]["level"] = feedback
    at = at.run()
    for label in ("Level plan", "Chase plan", "Final report", "Scenario & workbook",
                  "Hybrid plan", "Level plan"):
        at = click(at, label)
        assert not at.exception
    assert dict(at.session_state["pieces"]) == before
    assert at.session_state["results"]["level"] is feedback
    assert at.session_state["last_wb"]["name"] == "wb.xlsx"
    assert at.session_state["step"] == "level"


def test_second_computer_gets_everything_back(server):
    a = sign_in(computer("S1"), sid=None)
    do_some_work(a)
    a = click(a, "Level plan")                      # navigating saves (in the background)
    wait_for_saves()
    saved = server.data[(None, "S1")]
    assert saved["pieces"]["level"]["help"] == {"Regular Production": 2}

    b = computer("S1")                             # another computer, same link
    assert not b.exception
    assert b.session_state["student"] == {"name": "Ann", "sid": "S1"}
    assert dict(b.session_state["pieces"]) == json.loads(json.dumps(dict(a.session_state["pieces"])))
    assert b.session_state["step"] == "level"
    assert persistence.workbook_bytes(b.session_state["last_wb"]) == fixture("key_seed7.xlsx")
    res = b.session_state["results"]["level"]       # feedback rebuilt from the workbook
    assert res.complete and "chase" not in b.session_state["results"]
    assert any("My latest workbook" in d.label for d in b.sidebar.get("download_button"))


def test_typing_the_id_on_a_new_computer_never_overwrites_saved_work(server):
    a = sign_in(computer("S1"), sid=None)
    do_some_work(a)
    a = click(a, "Level plan")
    wait_for_saves()
    n_before = len(server.saves)

    b = sign_in(computer(), name="Ann on the lab PC", sid="S1")   # no ID in the link
    b = b.run()
    wait_for_saves()
    for st in server.saves[n_before:]:              # nothing empty was ever written
        assert st["pieces"]["level"]["attempts"], "saved work was overwritten"
    assert b.session_state["pieces"]["chase"]["submitted"]


def test_saved_work_for_a_different_setup_is_kept_not_overwritten(server):
    a = sign_in(computer("S1"), sid=None)
    do_some_work(a)
    a = click(a, "Level plan")
    wait_for_saves()
    server.data[(None, "S1")]["config_hash"] = "something-else"
    n = len(server.saves)
    b = computer("S1")
    assert b.session_state["restore"]["status"] == "incompatible"
    b = sign_in(b, sid=None)
    b = click(b, "Chase plan")
    wait_for_saves()
    assert len(server.saves) == n                   # saving paused
    b = click(b, "Scenario & workbook")
    assert any("different setup" in w.value for w in b.warning)


def test_a_load_error_pauses_saving(server):
    server.data[(None, "S1")] = {"config_hash": "x", "pieces": {}}
    server.fail_load = True
    b = computer("S1")
    assert b.session_state["restore"]["status"] == "error"
    n = len(server.saves)
    b = sign_in(b, sid=None)
    wait_for_saves()
    assert len(server.saves) == n


def test_progress_file_round_trip_and_tamper_checks():
    state = {"sid": "S1", "pieces": {"chase": {"attempts": [1]}}, "last_wb": None}
    blob = persistence.encode_file(state, b"k")
    got, err = persistence.decode_file(blob, b"k")
    assert got == state and not err
    doc = json.loads(blob)
    doc["state"]["pieces"]["chase"]["attempts"] = []            # a student edits it
    assert persistence.decode_file(json.dumps(doc).encode(), b"k")[0] is None
    assert persistence.decode_file(blob, b"other-key")[0] is None
    assert persistence.decode_file(b"not json")[0] is None


def test_workbook_entry_round_trip_and_size_limit():
    data = fixture("key_seed7.xlsx")
    e = persistence.workbook_entry("w.xlsx", data)
    assert persistence.workbook_bytes(e) == data
    e["b64"] = e["b64"][:-8] + "AAAAAAAA"
    assert persistence.workbook_bytes(e) is None                 # damaged -> refused
    assert persistence.workbook_entry("big.xlsx", b"x" * (persistence.MAX_WORKBOOK_BYTES + 1)) is None
