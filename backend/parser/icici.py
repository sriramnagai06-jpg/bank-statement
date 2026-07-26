"""
ICICI Bank statement parser.

NOT YET VERIFIED against a real ICICI statement — built as a generic
table-based parser using ICICI's commonly documented column headers
(Transaction Date, Value Date, Transaction Remarks, Withdrawal Amount
(INR), Deposit Amount (INR), Balance (INR)). Send a real ICICI PDF to
calibrate if this returns zero or incorrect results.
"""

from .generic_table import parse_generic

ICICI_COLUMN_ALIASES = {
    "date": ["transaction date", "value date"],
    "narration": ["transaction remarks"],
    "debit": ["withdrawal amount (inr )", "withdrawal amount (inr)"],
    "credit": ["deposit amount (inr )", "deposit amount (inr)"],
    "balance": ["balance (inr )", "balance (inr)"],
}


def parse(pdf_path, password=None):
    return parse_generic(pdf_path, column_aliases=ICICI_COLUMN_ALIASES, password=password)
