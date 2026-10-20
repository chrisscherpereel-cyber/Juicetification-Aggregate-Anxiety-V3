"""PDF receipts for submitted pieces and the final report.

Each piece (chase, level, hybrid, comparison) gets a one-page receipt when it is submitted;
the final report lists every piece — when it was submitted, how many uploads it took, how
the first attempt went and how far up the hint ladder the student went — plus the plan
comparison and the recommendation and justification verbatim. Anything incomplete is
marked DRAFT on every page."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import os
import unicodedata
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from . import APP_NAME, MODEL_VERSION
from .hints import LEVEL_NAMES
from .scenario import Scenario
from .text import n0, param_rows

_FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
_FONTS = {"DejaVu": "DejaVuSans.ttf", "DejaVu-Bold": "DejaVuSans-Bold.ttf",
          "DejaVu-Oblique": "DejaVuSans-Oblique.ttf",
          "DejaVu-BoldOblique": "DejaVuSans-BoldOblique.ttf"}
NAVY = colors.HexColor("#1F3864")
METRICS = ["Labor Cost", "Hiring Cost", "Layoff Cost", "Overtime Cost", "Subcontract Cost",
           "Holding Cost", "Backorder Cost", "Total Cost", "Total Hires", "Total Layoffs",
           "Peak Workforce", "Peak Inventory", "Months With Backorders", "On-Time Fulfillment %"]


def _register_fonts() -> set:
    if "DejaVu" not in pdfmetrics.getRegisteredFontNames():
        for name, fn in _FONTS.items():
            pdfmetrics.registerFont(TTFont(name, os.path.join(_FONT_DIR, fn)))
        pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold",
                                      italic="DejaVu-Oblique", boldItalic="DejaVu-BoldOblique")
    return set(pdfmetrics.getFont("DejaVu").face.charToGlyph.keys())


def _clean(s, cmap: set) -> str:
    """Escape for Paragraph markup; characters the font lacks become □."""
    s = unicodedata.normalize("NFC", "" if s is None else str(s))
    out = []
    for ch in s.replace("\r\n", "\n").replace("\r", "\n"):
        if ch == "\n" or ord(ch) in cmap:
            out.append(ch)
        elif ord(ch) >= 32:
            out.append("□")
    return escape("".join(out)).replace("\n", "<br/>")


def now_in(tz_name: str) -> dt.datetime:
    if tz_name:
        try:
            from zoneinfo import ZoneInfo
            return dt.datetime.now(ZoneInfo(tz_name))
        except Exception:                                     # noqa: BLE001
            pass
    return dt.datetime.now().astimezone()


def build_report(scn: Scenario, *, student_name: str, student_id: str, identity_label: str,
                 pieces: List[dict], final: bool, heading: str,
                 comparison: Optional[Dict[str, Dict[str, float]]] = None,
                 feasibility: Optional[List[dict]] = None, recommendation: str = "",
                 justification: Optional[str] = None, words: int = 0) -> Tuple[bytes, str]:
    """(pdf bytes, report id). `pieces`: dicts with title, status, attempts, first,
    submitted_at, help, file, sha."""
    cmap = _register_fonts()
    now = now_in(str(scn.params.get("report_timezone") or ""))
    seed = "|".join(p.get("sha", "") for p in pieces) + f"|{student_id}|{now.isoformat()}"
    rid = hashlib.sha256(seed.encode()).hexdigest()[:10].upper()

    base = ParagraphStyle("b", fontName="DejaVu", fontSize=9.5, leading=13)
    small = ParagraphStyle("s", parent=base, fontSize=8, leading=10.5,
                           textColor=colors.HexColor("#444444"))
    h1 = ParagraphStyle("h1", parent=base, fontName="DejaVu-Bold", fontSize=16, leading=20,
                        textColor=NAVY)
    h2 = ParagraphStyle("h2", parent=base, fontName="DejaVu-Bold", fontSize=11.5, leading=15,
                        textColor=NAVY, spaceBefore=10, spaceAfter=4)
    cell = ParagraphStyle("c", parent=base, fontSize=8.5, leading=11)
    P = lambda t, st=base: Paragraph(t, st)                                # noqa: E731
    C = lambda t: Paragraph(_clean(t, cmap), cell)                          # noqa: E731

    def table(rows, widths, head=True):
        body = ([rows[0]] + [[C(x) for x in r] for r in rows[1:]]) if head else \
            [[C(x) for x in r] for r in rows]
        t = Table(body, colWidths=widths, repeatRows=1 if head else 0)
        style = [("FONTNAME", (0, 0), (-1, -1), "DejaVu"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                 ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A6A6A6")),
                 ("VALIGN", (0, 0), (-1, -1), "TOP")]
        if head:
            style += [("BACKGROUND", (0, 0), (-1, 0), NAVY),
                      ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                      ("FONTNAME", (0, 0), (-1, 0), "DejaVu-Bold")]
        t.setStyle(TableStyle(style))
        return t

    story = [P(_clean(str(scn.params.get("assignment_name") or APP_NAME), cmap), h1),
             P(_clean(heading + ("" if final else " — DRAFT"), cmap),
               ParagraphStyle("st", parent=h2, textColor=(colors.HexColor("#1E7B34") if final
                                                          else colors.HexColor("#B00020")))),
             Spacer(1, 4)]
    info = [["Student", student_name or "—"], ["Student ID", student_id or "—"],
            ["Identity", identity_label], ["Created", now.strftime("%Y-%m-%d %H:%M %Z").strip()],
            ["Report ID", rid],
            ["Scenario", f"#{scn.seed}  ·  config {scn.config_hash}  ·  model {MODEL_VERSION}"]]
    story.append(table(info, [1.4 * inch, 5.6 * inch], head=False))

    story.append(P("Pieces", h2))
    rows = [["Piece", "Status", "Uploads", "First attempt", "Last attempt", "Most help used",
             "Submitted"]]
    for p in pieces:
        lvl = int(p.get("help") or 0)
        rows.append([p["title"], p["status"], str(p.get("attempts", 0)), p.get("first") or "—",
                     p.get("last") or "—", f"{lvl} · {LEVEL_NAMES[lvl]}" if lvl else "none",
                     p.get("submitted_at") or "—"])
    story.append(table(rows, [1.1 * inch, 0.8 * inch, 0.6 * inch, 0.95 * inch, 0.95 * inch,
                              1.2 * inch, 1.4 * inch]))
    for p in pieces:
        if p.get("history") and int(p.get("attempts", 0)) > 1:
            story.append(P(_clean(f"{p['title']}, every upload (cells correct): "
                                  f"{p['history']}", cmap), small))
    for p in pieces:
        if p.get("sha"):
            story.append(P(_clean(f"{p['title']}: {p.get('file', '')} · SHA-256 "
                                  f"{p['sha'][:24]}…", cmap), small))
    story.append(P("“Most help used” is the highest hint step unlocked by uploading again after "
                   "a mistake (1 a question … 4 the full rule). Uploads count every changed "
                   "file checked for that piece.", small))

    if comparison:
        story.append(P("Plan comparison (verified values)", h2))
        kinds = [k for k in ("chase", "level", "hybrid") if k in comparison]
        rows = [["Measure"] + [k.title() for k in kinds]]
        for m in METRICS:
            vals = [comparison[k][m] for k in kinds]
            if m.endswith("Cost"):
                vals = [f"${v:,.2f}" for v in vals]
            elif m.endswith("%"):
                vals = [f"{v:.1f}%" for v in vals]
            else:
                vals = [n0(v) for v in vals]
            rows.append([m] + vals)
        w = 5.0 / len(kinds)
        story.append(table(rows, [2.0 * inch] + [w * inch] * len(kinds)))
    if feasibility:
        ok = all(c["ok"] for c in feasibility)
        story.append(Spacer(1, 4))
        story.append(P(_clean("Hybrid plan is " + ("FEASIBLE: every limit is met." if ok else
                       "NOT feasible: " + "; ".join(c["detail"] for c in feasibility
                                                    if not c["ok"])), cmap), small))
    if justification is not None:
        story.append(P("Recommendation", h2))
        story.append(P(_clean(recommendation or "(none chosen)", cmap),
                       ParagraphStyle("r", parent=base, fontName="DejaVu-Bold")))
        story.append(Spacer(1, 4))
        story.append(P(_clean(justification or "(no justification written)", cmap)))
        story.append(P(_clean(f"{words} words", cmap), small))

    story.append(P("Scenario data", h2))
    story.append(table([[a, f"{b} {c}".strip()] for a, b, c in param_rows(scn)],
                       [3.2 * inch, 3.8 * inch], head=False))
    story.append(Spacer(1, 4))
    story.append(P(_clean("Demand: " + ", ".join(f"{m[:3]} {n0(d)}" for m, d in
                                                  zip(scn.months, scn.demand)), cmap), small))
    story.append(Spacer(1, 8))
    story.append(P("Plan cost is a result of the plan, not an academic grade. The recommendation "
                   "and justification are graded by your instructor.", small))

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont("DejaVu", 7.5)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(0.75 * inch, 0.5 * inch, f"{APP_NAME} · report {rid} · "
                          f"{'FINAL' if final else 'DRAFT'} · page {doc.page}")
        if not final:
            canvas.setFont("DejaVu-Bold", 60)
            canvas.setFillColor(colors.Color(0.85, 0.1, 0.1, alpha=0.10))
            canvas.translate(4.25 * inch, 5.5 * inch)
            canvas.rotate(35)
            canvas.drawCentredString(0, 0, "DRAFT")
        canvas.restoreState()

    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter, leftMargin=0.75 * inch, rightMargin=0.75 * inch,
                            topMargin=0.7 * inch, bottomMargin=0.75 * inch,
                            title="Aggregate Anxiety report", author=student_name or student_id)
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return bio.getvalue(), rid


def report_file_name(student_name: str, what: str, rid: str) -> str:
    base = "".join(ch for ch in (student_name or "student") if ch.isalnum() or ch in " -_")
    what = "".join(ch for ch in what.title() if ch.isalnum()) or "Report"
    return f"Aggregate_Anxiety_{what}_{base.strip().replace(' ', '_') or 'student'}_{rid}.pdf"
