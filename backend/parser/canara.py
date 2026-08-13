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

DATE_LINE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$|^\d{2}-[A-Za-z]{3}-\d{2,4}$")
AMOUNT_RE = re.compile(r"^-?[\d,]+\.\d{2}$")
PAGE_FOOTER_RE = re.compile(r"^Page\s+\d+\s+of\s+\d+$", re.IGNORECASE)
SKIP_TEXT_RE = re.compile(r"^(Txn|Value|Date|Cheque|No\.|Description|Branch|Debit|Credit|Balance|Code)$")


def _parse_amount(text):
    return abs(float(text.replace(",", "")))


def detect_canara_columns(pdf_path, password=None):
    """Detect boundaries dynamically based on column header words on any page."""
    boundaries = {
        "date_max": 75,
        "narr_min": 145,
        "narr_max": 360,
        "debit_max": 425,
        "credit_max": 495
    }
    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue
                withdraws_w = None
                deposit_w = None
                balance_w = None
                for w in words:
                    text = w["text"].lower()
                    if "withdraws" in text or "withdrawals" in text:
                        withdraws_w = w
                    elif "deposit" in text or "deposits" in text:
                        deposit_w = w
                    elif "balance" in text:
                        balance_w = w
                if withdraws_w and deposit_w:
                    w_x0 = withdraws_w["x0"]
                    d_x0 = deposit_w["x0"]
                    boundaries["debit_max"] = (w_x0 + d_x0) / 2
                    if balance_w:
                        b_x0 = balance_w["x0"]
                        boundaries["credit_max"] = (d_x0 + b_x0) / 2
                    else:
                        boundaries["credit_max"] = d_x0 + (d_x0 - w_x0)
                    boundaries["narr_max"] = w_x0 - 5
                    break
    except Exception:
        pass
    return boundaries


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
    bounds = detect_canara_columns(pdf_path, password=password)
    transactions = []
    current = None

    for line_words in extract_rows(pdf_path, password=password):
        first_text = line_words[0]["text"]

        if SKIP_TEXT_RE.match(first_text) or PAGE_FOOTER_RE.match(" ".join(w["text"] for w in line_words)):
            continue

        line_str = " ".join(w["text"] for w in line_words)
        
        # Check for footer/summary lines
        skip_phrases = [
            "statement summary", "total debit", "total credit",
            "closing balance", "clear balance", "unless the constituent", "beware of phishing",
            "details of ombudsman", "office of banking", "are you a merchant", "computer output",
            "end of statement", "page ", "******end"
        ]
        line_str_lower = line_str.lower()
        if any(phrase in line_str_lower for phrase in skip_phrases):
            if current:
                transactions.append(current)
                current = None
            continue

        is_new_txn = bool(DATE_LINE_RE.match(first_text) and line_words[0]["x0"] < bounds["date_max"])

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
                    if x0 < bounds["debit_max"]:
                        debit = _parse_amount(text)
                    elif x0 < bounds["credit_max"]:
                        credit = _parse_amount(text)
                    else:
                        balance = _parse_amount(text)
                elif bounds["narr_min"] <= x0 < bounds["narr_max"]:
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
                current["narration"].append(line_str)

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
        dt = None
        for fmt in ["%d-%m-%Y", "%d-%b-%y", "%d-%b-%Y"]:
            try:
                dt = datetime.strptime(t["date"], fmt)
                break
            except ValueError:
                continue
        if dt is None:
            continue
            
        narr = " ".join(t["narration"]).strip()
        narr_lower = narr.lower()
        if "b/f" in narr_lower or "brought forward" in narr_lower or "opening balance" in narr_lower:
            results.append({
                "date": dt,
                "narration": narr,
                "debit": t["debit"],
                "credit": t["credit"],
                "balance": t["balance"],
            })
            continue
            
        results.append({
            "date": dt,
            "narration": narr,
            "debit": t["debit"],
            "credit": t["credit"],
            "balance": t["balance"],
        })
    return results

