"""
HDFC Bank statement parser.

NOT YET VERIFIED against a real HDFC statement — built as a generic
table-based parser using HDFC's commonly documented column headers
(Date, Narration, Chq/Ref No., Value Dt, Withdrawal Amt., Deposit
Amt., Closing Balance). Send a real HDFC PDF to calibrate if this
returns zero or incorrect results.
"""

from .generic_table import parse_generic

HDFC_COLUMN_ALIASES = {
    "date": ["date", "value dt"],
    "narration": ["narration"],
    "debit": ["withdrawal amt.", "withdrawal amt"],
    "credit": ["deposit amt.", "deposit amt"],
    "balance": ["closing balance"],
}


def parse(pdf_path, password=None):
    return parse_generic(pdf_path, column_aliases=HDFC_COLUMN_ALIASES, password=password)
