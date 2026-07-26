"""
Dispatches a statement PDF to the correct bank-specific parser based
on auto-detection.
"""

from . import canara, sbi, hdfc, icici, axis, generic
from .detector import detect_bank

PARSERS = {
    "canara": canara,
    "sbi": sbi,
    "hdfc": hdfc,
    "icici": icici,
    "axis": axis,
    "pnb": generic,
    "bob": generic,
    "kotak": generic,
    "indusind": generic,
    "union": generic,
    "idfc": generic,
    "yes": generic,
    "boi": generic,
    "cbi": generic,
    "iob": generic,
    "uco": generic,
    "federal": generic,
    "southindian": generic,
    "indian": generic,
    "other": generic,
}

BANK_DISPLAY_NAMES = {
    "canara": "Canara Bank",
    "sbi": "State Bank of India",
    "hdfc": "HDFC Bank",
    "icici": "ICICI Bank",
    "axis": "Axis Bank",
    "pnb": "Punjab National Bank",
    "bob": "Bank of Baroda",
    "kotak": "Kotak Mahindra Bank",
    "indusind": "IndusInd Bank",
    "union": "Union Bank of India",
    "idfc": "IDFC First Bank",
    "yes": "Yes Bank",
    "boi": "Bank of India",
    "cbi": "Central Bank of India",
    "iob": "Indian Overseas Bank",
    "uco": "UCO Bank",
    "federal": "Federal Bank",
    "southindian": "South Indian Bank",
    "indian": "Indian Bank",
    "other": "Other Indian Bank",
    "unknown": "Unknown Bank",
}


class UnsupportedBankError(Exception):
    pass


def parse_statement(pdf_path, forced_bank=None, password=None):
    """Detect the bank (unless forced_bank is given) and parse the PDF.

    Returns (bank_key, bank_display_name, transactions).
    Raises UnsupportedBankError if the bank can't be identified or has
    no parser.
    """
    bank_key = forced_bank or detect_bank(pdf_path, password=password)

    if bank_key not in PARSERS:
        if forced_bank == "other":
            bank_key = "other"
        else:
            raise UnsupportedBankError(
                f"Could not identify the bank for this statement "
                f"(detected: '{bank_key}'). Supported banks: "
                f"{', '.join(PARSERS.keys())}."
            )

    transactions = PARSERS[bank_key].parse(pdf_path, password=password)
    return bank_key, BANK_DISPLAY_NAMES.get(bank_key, bank_key), transactions
