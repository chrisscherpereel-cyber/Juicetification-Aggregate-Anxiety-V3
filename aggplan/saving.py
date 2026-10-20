"""Background progress saving, shared by every session in the process.

Saving to the Dropbox store is a network round trip (typically 0.3-1 s). Done inline it
freezes the student's screen on every upload, answer and setting change, and with a class of
30 those waits pile up. Here a save only records the latest state and returns; a small pool
uploads it. Saves for the same student coalesce: if three changes happen before the upload
starts, only the newest state is sent.

This module lives outside app.py on purpose: Streamlit re-executes app.py on every rerun,
so anything defined there would be rebuilt each click; a module is imported once per process.
"""

from __future__ import annotations

import json
import threading
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, Dict, Tuple

_POOL = ThreadPoolExecutor(max_workers=2, thread_name_prefix="aa-save")
_PENDING: Dict[str, Tuple[Callable, tuple, dict]] = {}
_RUNNING: set = set()            # students whose save is in flight — one at a time each
_LOCK = threading.Lock()


def _snapshot(obj: Any) -> Any:
    """A JSON-safe deep copy taken NOW, so later edits to session state can't leak in."""
    return json.loads(json.dumps(obj, default=str))


def _flush(key: str) -> None:
    """Upload this student's newest state; repeat while newer state arrives meanwhile.
    Saves for one student never overlap, so an older state can't land after a newer one."""
    while True:
        with _LOCK:
            job = _PENDING.pop(key, None)
            if job is None:
                _RUNNING.discard(key)
                return
        fn, args, kwargs = job
        try:
            fn(*args, **kwargs)
        except Exception:                                # noqa: BLE001 — saving is best effort
            pass


def save_later(key: str, fn: Callable, state: dict, *args, **kwargs) -> None:
    """Queue fn(*args, snapshot_of_state, **kwargs) for `key` (one student)."""
    job = (fn, args + (_snapshot(state),), kwargs)
    with _LOCK:
        _PENDING[key] = job                                # newest state wins
        if key in _RUNNING:
            return                                         # the running flush will pick it up
        _RUNNING.add(key)
    _POOL.submit(_flush, key)


def pending() -> int:
    with _LOCK:
        return len(_PENDING) + len(_RUNNING)
