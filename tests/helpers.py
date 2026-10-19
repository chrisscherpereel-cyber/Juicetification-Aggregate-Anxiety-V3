"""Shared test helpers."""

import os

from aggplan.checker import Book
from aggplan.scenario import resolve_scenario

FIX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")
SEED, SID = 7, "S1"


def scn7(**cfg):
    return resolve_scenario(cfg, SEED)


def fixture(name: str) -> bytes:
    with open(os.path.join(FIX, name), "rb") as f:
        return f.read()


def book_from(base: Book, changes: dict) -> Book:
    """Copy a Book and overwrite cells: {(sheet, ref): Cell(value, formula)}."""
    sheets = {k: dict(v) for k, v in base.sheets.items()}
    for (sheet, ref), cell in changes.items():
        sheets.setdefault(sheet, {})[ref] = cell
    return Book(sheets, dict(base.meta), base.sha256)
