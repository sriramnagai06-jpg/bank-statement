"""
State Bank of India (SBI) statement parser.

NOT YET VERIFIED against a real SBI statement — built as a generic
table-based parser using SBI's commonly documented column headers
(Txn Date, Value Date, Description, Ref No./Cheque No., Debit,
Credit, Balance). Send a real SBI PDF to calibrate if this returns
zero or incorrect results.
"""

from .generic_table import parse_generic

SBI_COLUMN_ALIASES = {
    "date": ["txn date", "value date"],
    "narration": ["description"],
    "debit": ["debit"],
    "credit": ["credit"],
    "balance": ["balance"],
}


def parse(pdf_path, password=None):
    return parse_generic(pdf_path, column_aliases=SBI_COLUMN_ALIASES, password=password)
