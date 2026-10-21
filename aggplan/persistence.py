"""What a student's saved work is, and the progress file that carries it between computers.

A snapshot holds everything needed to continue exactly where the student left off on any
computer: name, policies, hybrid mode, every piece (uploads, hint levels, guided answers,
submissions), the step they were on, and their latest uploaded workbook (so they can
download it again on a computer that doesn't have it). Check results are not stored; they
are rebuilt by re-checking that workbook.

It is saved to the server store when one is configured, and can always be downloaded as a
progress file — a JSON document with a checksum (and an HMAC signature when the deployment
has ASSIGNMENT_SIGNING_KEY) so a damaged or edited file is refused."""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
from typing import Any, Dict, Optional, Tuple

APP = "aggregate-anxiety-progress"
SCHEMA = 1
MAX_WORKBOOK_BYTES = 1_500_000      # larger uploads aren't kept (the 5 MB limit is for safety)


def workbook_entry(name: str, data: bytes) -> Optional[dict]:
    if not data or len(data) > MAX_WORKBOOK_BYTES:
        return None
    return {"name": name, "sha": hashlib.sha256(data).hexdigest(),
            "b64": base64.b64encode(data).decode("ascii")}


def workbook_bytes(entry: Optional[dict]) -> Optional[bytes]:
    if not entry or not entry.get("b64"):
        return None
    try:
        data = base64.b64decode(entry["b64"])
    except Exception:                                     # noqa: BLE001
        return None
    return data if hashlib.sha256(data).hexdigest() == entry.get("sha") else None


def snapshot(*, sid: str, name: str, seed: int, config_hash: str, pieces: dict, policy: dict,
             hybrid_mode: str, step: str, last_wb: Optional[dict],
             downloaded_hash: Optional[str], restarts: int = 0,
             restart_log: Optional[list] = None) -> Dict[str, Any]:
    return {"schema": SCHEMA, "sid": sid, "name": name, "seed": seed,
            "config_hash": config_hash, "pieces": pieces, "policy": policy,
            "hybrid_mode": hybrid_mode, "step": step, "last_wb": last_wb,
            "downloaded_hash": downloaded_hash, "restarts": restarts,
            "restart_log": restart_log or []}


def _canonical(state: dict) -> bytes:
    return json.dumps(state, sort_keys=True, separators=(",", ":"), default=str).encode()


def encode_file(state: dict, key: Optional[bytes] = None) -> bytes:
    body = _canonical(state)
    doc = {"app": APP, "schema": SCHEMA, "state": state,
           "sum": hashlib.sha256(body).hexdigest()}
    if key:
        doc["sig"] = hmac.new(key, body, hashlib.sha256).hexdigest()
    return json.dumps(doc, separators=(",", ":"), default=str).encode()


def decode_file(data: bytes, key: Optional[bytes] = None) -> Tuple[Optional[dict], str]:
    """(state, "") or (None, a message the student can act on)."""
    try:
        doc = json.loads(data.decode("utf-8"))
    except Exception:                                     # noqa: BLE001
        return None, "That isn't a progress file from this app."
    if not isinstance(doc, dict) or doc.get("app") != APP or not isinstance(doc.get("state"), dict):
        return None, "That isn't a progress file from this app."
    body = _canonical(doc["state"])
    if hashlib.sha256(body).hexdigest() != doc.get("sum"):
        return None, "This progress file is damaged or was edited, so it can't be used."
    if key and not hmac.compare_digest(hmac.new(key, body, hashlib.sha256).hexdigest(),
                                       str(doc.get("sig", ""))):
        return None, "This progress file wasn't made by this assignment's app."
    return doc["state"], ""
