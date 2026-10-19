"""aggplan — scenario, answer key, workbook builder, workbook checker and PDF report for
Juicetification: Aggregate Anxiety (Excel edition). Nothing here imports Streamlit at
module load, so it is unit-testable."""

from manifest import MODEL_VERSION  # single source of truth for the model version

APP_NAME = "Juicetification: Aggregate Anxiety"
__all__ = ["MODEL_VERSION", "APP_NAME"]
