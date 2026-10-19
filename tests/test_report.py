from aggplan import checker, report
from tests.helpers import SID, fixture, scn7


def _pdf_text(pdf: bytes) -> str:
    try:
        import fitz                                   # optional (pymupdf)
    except ImportError:
        return ""
    return "".join(p.get_text() for p in fitz.open(stream=pdf, filetype="pdf"))


def test_receipt_and_final_report():
    s = scn7()
    summ = checker.check_piece(fixture("key_seed7.xlsx"), s, "summary", SID)
    pieces = [{"title": "Chase plan", "status": "Submitted", "attempts": 3, "first": "120 / 179",
               "submitted_at": "2026-10-19 10:00", "help": 2, "file": "wb.xlsx", "sha": "ab" * 32},
              {"title": "Level plan", "status": "In progress", "attempts": 1, "first": "", "help": 0}]
    for final in (True, False):
        pdf, rid = report.build_report(s, student_name="Zoë O'Brien & 李", student_id=SID,
                                       identity_label="test", pieces=pieces, final=final,
                                       heading="Final report", comparison=summ.metrics,
                                       recommendation="Chase", justification="Because.", words=1)
        assert pdf.startswith(b"%PDF") and len(rid) == 10
        text = _pdf_text(pdf)
        if text:
            assert ("DRAFT" in text) == (not final)
            assert "Zoë O'Brien" in text and "2 · The idea" in text and "$426,000.00" in text
