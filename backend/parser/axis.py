"""
Axis Bank statement parser.

NOT YET VERIFIED against a real Axis statement — built as a generic
table-based parser using Axis's commonly documented column headers
(Tran Date, Chq No, Particulars, Debit, Credit, Balance). Send a real
Axis PDF to calibrate if this returns zero or incorrect results.
"""

from .generic_table import parse_generic

AXIS_COLUMN_ALIASES = {
    "date": ["tran date"],
    "narration": ["particulars"],
    "debit": ["debit"],
    "credit": ["credit"],
    "balance": ["balance"],
}


def parse(pdf_path, password=None):
    return parse_generic(pdf_path, column_aliases=AXIS_COLUMN_ALIASES, password=password)
