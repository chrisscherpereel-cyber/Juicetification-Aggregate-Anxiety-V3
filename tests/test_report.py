from aggplan import checker, report
from tests.helpers import SID, fixture, scn7


def _pdf_text(pdf: bytes) -> str:
    try:
        import fitz                                   # optional (pymupdf)
    except ImportError:
        return ""
    return "".join(p.get_text() for p in fitz.open(stream=pdf, filetype="pdf"))


def test_final_and_draft_reports():
    s = scn7()
    for name, final in (("key_seed7.xlsx", True), ("mistakes_seed7.xlsx", False)):
        res = checker.check_workbook(fixture(name), s, SID)
        pdf, attempt = report.build_report(res, s, student_name="Zoë O'Brien & 李",
                                           student_id=SID, identity_label="test", file_name=name)
        assert pdf.startswith(b"%PDF") and len(attempt) == 10
        text = _pdf_text(pdf)
        if text:
            assert ("COMPLETE" in text) == final
            assert ("DRAFT" in text) == (not final)
            assert "Zoë O'Brien" in text
