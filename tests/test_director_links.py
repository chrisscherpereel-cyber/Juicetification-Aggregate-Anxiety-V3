"""Links from the Juicetification Director: settings in the link (?cfg=…) or, when the
Director's store is on, just ?game=CODE with the settings published to the store."""

import os

import pytest

import student_store
from juice_director import encode_cfg

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")
SETTINGS = {"hybrid_mode": "off", "worker_policy": "partial", "student_sets_policies": False}


def open_app(**query):
    import streamlit as st
    st.cache_data.clear()
    at = st_testing.AppTest.from_file(APP, default_timeout=60)
    for k, v in query.items():
        at.query_params[k] = v
    return at.run()


def labels(at):
    return [b.label for b in at.sidebar.button]


@pytest.fixture
def published(monkeypatch):
    calls = []

    def load(code):
        calls.append(code)
        return {"params": SETTINGS, "seed": 4242} if code == "AA3TEST" else None

    monkeypatch.setattr(student_store, "load_game_config", load)
    return calls


def test_short_game_link_uses_the_published_settings(published):
    at = open_app(game="AA3TEST", sid="S1")
    assert not at.exception
    assert published == ["AA3TEST"]
    assert at.session_state["hybrid_mode"] == "off"                  # hybrid switched off
    assert not any("Hybrid plan" in b for b in labels(at))
    assert at.session_state["policy"]["worker_policy"] == "partial"
    at.text_input[0].input("Ann")
    next(b for b in at.button if b.label == "Start").click()
    at = at.run()
    sidebar = " ".join(m.value for m in at.sidebar.markdown) + " ".join(
        c.value for c in at.sidebar.caption)
    assert "4242" in sidebar                                          # the game's fixed seed


def test_settings_in_the_link_win_and_need_no_lookup(published):
    at = open_app(game="AA3TEST", cfg=encode_cfg({"hybrid_mode": "build"}), sid="S1")
    assert not at.exception
    assert published == []
    assert at.session_state["hybrid_mode"] == "build"


def test_unknown_game_falls_back_to_defaults(published):
    at = open_app(game="NOSUCH", sid="S1")
    assert not at.exception
    assert at.session_state["hybrid_mode"] == "guided"
