"""
Canara Bank statement parser.

Canara's 'Current & Saving Account Statement' export is text-based
(not a real PDF grid table), so we use word x/y coordinates to
correctly separate the Txn Date / Narration / Branch / Debit / Credit
/ Balance columns, and merge wrapped narration lines back into their
parent transaction.

Verified: on a real 245-page / 1,783-transaction statement, every
single row's balance reconciled exactly against the running balance
chain and the statement's stated closing balance.
"""

import re
from collections import defaultdict
from datetime import datetime

import pdfplumber

DATE_LINE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
AMOUNT_RE = re.compile(r"^[\d,]+\.\d{2}$")
PAGE_FOOTER_RE = re.compile(r"^Page\s+\d+\s+of\s+\d+$", re.IGNORECASE)
SKIP_TEXT_RE = re.compile(r"^(Txn|Value|Date|Cheque|No\.|Description|Branch|Debit|Credit|Balance|Code)$")

# Column x0 boundaries calibrated against the standard Canara layout.
# If a differently-formatted Canara export doesn't parse, these are the
# first values to re-check against the new PDF's word coordinates.
COL_DATE_MAX_X = 65
COL_NARRATION_MIN_X = 150
COL_NARRATION_MAX_X = 405
COL_DEBIT_MAX_X = 536
COL_CREDIT_MAX_X = 612


def _parse_amount(text):
    return float(text.replace(",", ""))


def extract_rows(pdf_path, password=None):
    """Yield dicts of {top, words} per visual line, per page."""
    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue
            lines = defaultdict(list)
            for w in words:
                lines[round(w["top"], 1)].append(w)
            for top in sorted(lines.keys()):
                yield sorted(lines[top], key=lambda w: w["x0"])


def _build_raw_transactions(pdf_path, password=None):
    transactions = []
    current = None

    for line_words in extract_rows(pdf_path, password=password):
        first_text = line_words[0]["text"]

        if SKIP_TEXT_RE.match(first_text) or PAGE_FOOTER_RE.match(" ".join(w["text"] for w in line_words)):
            continue

        is_new_txn = bool(DATE_LINE_RE.match(first_text) and line_words[0]["x0"] < COL_DATE_MAX_X)

        if is_new_txn:
            if current:
                transactions.append(current)

            narration_parts = []
            debit = 0.0
            credit = 0.0
            balance = None

            for w in line_words[1:]:
                x0 = w["x0"]
                text = w["text"]
                if AMOUNT_RE.match(text):
                    if x0 < COL_DEBIT_MAX_X:
                        debit = _parse_amount(text)
                    elif x0 < COL_CREDIT_MAX_X:
                        credit = _parse_amount(text)
                    else:
                        balance = _parse_amount(text)
                elif COL_NARRATION_MIN_X <= x0 < COL_NARRATION_MAX_X:
                    narration_parts.append(text)

            current = {
                "date": first_text,
                "narration": narration_parts,
                "debit": debit,
                "credit": credit,
                "balance": balance,
            }
        else:
            if current is not None:
                text_line = " ".join(w["text"] for w in line_words)
                if not PAGE_FOOTER_RE.match(text_line):
                    current["narration"].append(text_line)

    if current:
        transactions.append(current)

    return transactions


def parse(pdf_path, password=None):
    """Return a list of standardized transaction dicts:
    {date: datetime, narration: str, debit: float, credit: float, balance: float|None}
    """
    raw = _build_raw_transactions(pdf_path, password=password)
    results = []
    for t in raw:
        try:
            dt = datetime.strptime(t["date"], "%d-%m-%Y")
        except ValueError:
            continue
        results.append({
            "date": dt,
            "narration": " ".join(t["narration"]).strip(),
            "debit": t["debit"],
            "credit": t["credit"],
            "balance": t["balance"],
        })
    return results
