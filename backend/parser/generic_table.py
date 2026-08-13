"""
Generic grid-table statement parser.

Used as the base for banks whose PDFs export as real PDF tables
(rows/columns pdfplumber can detect directly), unlike Canara's
text-based layout. This covers most other major Indian banks, but
each bank's exact header wording differs, so `column_aliases` can be
extended per-bank when a real sample statement is available to test
against.

IMPORTANT: this generic parser has NOT been verified against a real
SBI / HDFC / ICICI / Axis statement (no sample was available at build
time). If it returns zero transactions or misclassifies debit/credit,
send a real sample statement so the column aliases and layout
assumptions can be corrected.
"""

import re
from datetime import datetime

import pdfplumber

DEFAULT_COLUMN_ALIASES = {
    "date": [
        "date", "txn date", "txn. date", "tran date", "tran. date", "transaction date",
        "value date", "value dt", "posting date", "date & time", "post date",
    ],
    "narration": [
        "narration", "description", "particulars", "details", "remarks",
        "transaction remarks", "transaction details", "narration/description",
        "narration / description", "transaction description",
    ],
    "debit": [
        "debit", "withdrawal", "withdrawals", "withdrawal amt", "withdrawal amount", "withdrawal amt.",
        "dr", "debit amount", "withdrawal amount (inr )", "withdrawal amount (inr)",
        "debit (inr)", "debit(inr)", "withdrawal(dr)", "withdrawal (dr)", "debit (rs.)",
    ],
    "credit": [
        "credit", "deposit", "deposits", "deposit amt", "deposit amount", "deposit amt.",
        "cr", "credit amount", "deposit amount (inr )", "deposit amount (inr)",
        "credit (inr)", "credit(inr)", "deposit(cr)", "deposit (cr)", "credit (rs.)",
    ],
    "balance": [
        "balance", "closing balance", "available balance", "balance amt", "balance amount",
        "balance (inr )", "balance (inr)", "balance(inr)", "balance (rs.)", "closing bal",
    ],
}

DATE_FORMATS = [
    "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
    "%d %b %Y", "%d-%b-%Y", "%d/%b/%Y", "%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y",
    "%d %b %y", "%d-%b-%y", "%d/%b/%y",
    "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %I:%M:%S %p",
]


def _normalize_header(cell):
    if cell is None:
        return ""
    return re.sub(r"\s+", " ", str(cell)).strip().lower()


def _map_columns(header_row, column_aliases):
    mapping = {}
    for idx, cell in enumerate(header_row):
        norm = _normalize_header(cell)
        for std_name, aliases in column_aliases.items():
            if norm in aliases or any(a in norm for a in aliases):
                mapping[std_name] = idx
                break
    return mapping


def _parse_amount(value):
    if value is None:
        return 0.0
    s = str(value).replace(",", "").replace("\u20b9", "").replace("Rs.", "").replace("INR", "").strip()
    s = re.sub(r"\b(Dr|Cr|DR|CR)\b", "", s).strip()
    if not s or s in ("", "-", "--"):
        return 0.0

    # Handle parenthesis format e.g. (1234.50)
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1].strip()

    try:
        val = float(s)
        return abs(val)
    except ValueError:
        return 0.0


def _parse_date(value):
    if value is None:
        return None
    s = str(value).strip()
    # If date contains time, strip time portion if split by space
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def parse_generic(pdf_path, column_aliases=None, password=None):
    """Return a list of standardized transaction dicts using pdfplumber's
    table detection. column_aliases can extend/override DEFAULT_COLUMN_ALIASES
    for a specific bank's header wording."""
    aliases = {**DEFAULT_COLUMN_ALIASES}
    if column_aliases:
        for k, v in column_aliases.items():
            aliases[k] = list(set(aliases.get(k, []) + v))

    results = []
    col_map = None

    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                start_row = 0
                header_map = _map_columns(table[0], aliases)
                if "date" in header_map and ("debit" in header_map or "credit" in header_map):
                    col_map = header_map
                    start_row = 1

                if col_map is None:
                    continue

                for row in table[start_row:]:
                    if not row or len(row) <= max(col_map.values()):
                        continue
                    dt = _parse_date(row[col_map.get("date")]) if "date" in col_map else None
                    if dt is None:
                        continue

                    narration = row[col_map["narration"]] if "narration" in col_map else ""
                    debit = _parse_amount(row[col_map["debit"]]) if "debit" in col_map else 0.0
                    credit = _parse_amount(row[col_map["credit"]]) if "credit" in col_map else 0.0
                    balance = _parse_amount(row[col_map["balance"]]) if "balance" in col_map else None

                    results.append({
                        "date": dt,
                        "narration": (narration or "").strip().replace("\n", " "),
                        "debit": debit,
                        "credit": credit,
                        "balance": balance,
                    })

    return results
