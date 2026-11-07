"""aggplan — scenario, answer key, workbook builder, workbook checker and PDF report for
Juicetification: Aggregate Anxiety (Excel edition). Nothing here imports Streamlit at
module load, so it is unit-testable."""

from manifest import MODEL_VERSION  # single source of truth for the model version

APP_NAME = "Juicetification: Aggregate Anxiety"
__all__ = ["MODEL_VERSION", "APP_NAME"]

_GC_TUNED = False


def tune_gc_once() -> None:
    """Once per server process: move everything loaded so far (Streamlit, pandas, openpyxl,
    the app's modules) out of the garbage collector's reach with gc.freeze(), and let the
    youngest generation grow larger before a collection. Checking an upload builds hundreds of
    thousands of short-lived objects; with the defaults, collections took ~15% of a check,
    re-scanning the same long-lived objects every time. Memory is unaffected in practice:
    frozen objects are the ones that live for the whole process anyway."""
    global _GC_TUNED
    if _GC_TUNED:
        return
    import gc
    gc.collect()
    gc.freeze()
    gc.set_threshold(10_000, 20, 20)
    _GC_TUNED = True
