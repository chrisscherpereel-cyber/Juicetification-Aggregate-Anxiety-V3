"""The PDF check report a student submits to the LMS next to their workbook.

It records WHICH workbook was checked (file name + SHA-256), when, the per-sheet result,
the plan comparison and the student's recommendation and justification verbatim. A report
for an incomplete workbook is marked DRAFT on every page."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import os
import unicodedata
from typing import Optional, Tuple
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from . import APP_NAME, MODEL_VERSION
from .checker import CheckResult
from .scenario import Scenario
from .text import n0, param_rows

_FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
_FONTS = {"DejaVu": "DejaVuSans.ttf", "DejaVu-Bold": "DejaVuSans-Bold.ttf",
          "DejaVu-Oblique": "DejaVuSans-Oblique.ttf",
          "DejaVu-BoldOblique": "DejaVuSans-BoldOblique.ttf"}
NAVY = colors.HexColor("#1F3864")


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


def _now(tz_name: str) -> dt.datetime:
    if tz_name:
        try:
            from zoneinfo import ZoneInfo
            return dt.datetime.now(ZoneInfo(tz_name))
        except Exception:                                     # noqa: BLE001
            pass
    return dt.datetime.now().astimezone()


def _money(x) -> str:
    return f"${float(x):,.2f}"


def build_report(res: CheckResult, scn: Scenario, *, student_name: str, student_id: str,
                 identity_label: str, file_name: str) -> Tuple[bytes, str]:
    """(pdf bytes, attempt id)."""
    cmap = _register_fonts()
    final = res.complete
    now = _now(str(scn.params.get("report_timezone") or ""))
    attempt = hashlib.sha256(f"{res.sha256}|{student_id}|{now.isoformat()}".encode()).hexdigest()[:10].upper()

    base = ParagraphStyle("b", fontName="DejaVu", fontSize=9.5, leading=13)
    small = ParagraphStyle("s", parent=base, fontSize=8, leading=10.5, textColor=colors.HexColor("#444444"))
    h1 = ParagraphStyle("h1", parent=base, fontName="DejaVu-Bold", fontSize=16, leading=20, textColor=NAVY)
    h2 = ParagraphStyle("h2", parent=base, fontName="DejaVu-Bold", fontSize=11.5, leading=15,
                        textColor=NAVY, spaceBefore=10, spaceAfter=4)
    cell = ParagraphStyle("c", parent=base, fontSize=8.5, leading=11)
    P = lambda t, st=base: Paragraph(t, st)                                # noqa: E731
    C = lambda t: Paragraph(_clean(t, cmap), cell)                          # noqa: E731

    def table(rows, widths, head=True):
        t = Table(rows, colWidths=widths, repeatRows=1 if head else 0)
        style = [("FONTNAME", (0, 0), (-1, -1), "DejaVu"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                 ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A6A6A6")),
                 ("VALIGN", (0, 0), (-1, -1), "TOP")]
        if head:
            style += [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                      ("FONTNAME", (0, 0), (-1, 0), "DejaVu-Bold")]
        t.setStyle(TableStyle(style))
        return t

    story = [P(_clean(str(scn.params.get("assignment_name") or APP_NAME), cmap), h1),
             P(("CHECK REPORT — COMPLETE" if final else "DRAFT — the workbook is not complete yet"),
               ParagraphStyle("st", parent=h2, textColor=(colors.HexColor("#1E7B34") if final
                                                          else colors.HexColor("#B00020")))),
             Spacer(1, 4)]
    info = [["Student", student_name or "—"],
            ["Student ID", student_id or "—"],
            ["Identity", identity_label],
            ["Checked", now.strftime("%Y-%m-%d %H:%M %Z").strip()],
            ["Attempt ID", attempt],
            ["Workbook", file_name],
            ["Workbook SHA-256", res.sha256[:32] + "…"],
            ["Scenario", f"#{scn.seed}  ·  config {scn.config_hash}  ·  model {MODEL_VERSION}"]]
    story.append(table([[C(a), C(b)] for a, b in info], [1.5 * inch, 5.5 * inch], head=False))

    story.append(P("Check results", h2))
    rows = [["Sheet", "Cells correct", "Status"]]
    for s in res.sheets.values():
        if not s.attempted:
            rows.append([s.title, "—", "optional, not attempted"])
            continue
        status = "complete" if s.complete else ("needs work" if s.required else "optional — needs work")
        rows.append([s.title, f"{s.n_ok} / {s.n_total}", status])
    story.append(table([rows[0]] + [[C(x) for x in r] for r in rows[1:]], [2.2 * inch, 1.6 * inch, 3.2 * inch]))
    if not final:
        story.append(Spacer(1, 4))
        story.append(P(_clean("Still to do: " + "; ".join(res.missing()), cmap), small))

    if res.plan_totals and final:
        story.append(P("Plan comparison (verified values)", h2))
        kinds = [k for k in ("chase", "level", "hybrid") if k in res.plan_totals
                 and (k != "hybrid" or res.sheets.get("hybrid") and res.sheets["hybrid"].complete)]
        metrics = ["Labor Cost", "Hiring Cost", "Layoff Cost", "Overtime Cost", "Subcontract Cost",
                   "Holding Cost", "Backorder Cost", "Total Cost", "Total Hires", "Total Layoffs",
                   "Peak Workforce", "Peak Inventory", "Months With Backorders",
                   "On-Time Fulfillment %"]
        rows = [["Measure"] + [k.title() for k in kinds]]
        for m in metrics:
            vals = [res.plan_totals[k][m] for k in kinds]
            if m.endswith("Cost"):
                vals = [_money(v) for v in vals]
            elif m.endswith("%"):
                vals = [f"{v:.1f}%" for v in vals]
            else:
                vals = [n0(v) for v in vals]
            rows.append([m] + vals)
        w = 6.0 / len(kinds)
        story.append(table([rows[0]] + [[C(x) for x in r] for r in rows[1:]], [1.9 * inch] + [w * 0.85 * inch] * len(kinds)))
        if res.hybrid_feasibility and "hybrid" in kinds:
            ok = all(c["ok"] for c in res.hybrid_feasibility)
            story.append(Spacer(1, 4))
            story.append(P(_clean("Hybrid plan is " + ("FEASIBLE: every limit is met." if ok else
                           "NOT feasible: " + "; ".join(c["detail"] for c in res.hybrid_feasibility
                                                        if not c["ok"])), cmap), small))

    story.append(P("Recommendation", h2))
    story.append(P(_clean(res.recommendation or "(none chosen)", cmap),
                   ParagraphStyle("r", parent=base, fontName="DejaVu-Bold")))
    story.append(Spacer(1, 4))
    story.append(P(_clean(res.justification or "(no justification written)", cmap)))
    story.append(P(_clean(f"{res.words} words", cmap), small))

    story.append(P("Scenario data", h2))
    story.append(table([[C(a), C(f"{b} {c}".strip())] for a, b, c in param_rows(scn)],
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
        canvas.drawString(0.75 * inch, 0.5 * inch, f"{APP_NAME} · attempt {attempt} · "
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
                            title="Aggregate Anxiety check report", author=student_name or student_id)
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return bio.getvalue(), attempt


def report_file_name(student_name: str, attempt: str) -> str:
    base = "".join(ch for ch in (student_name or "student") if ch.isalnum() or ch in " -_")
    return f"Aggregate_Anxiety_Report_{base.strip().replace(' ', '_') or 'student'}_{attempt}.pdf"
