"""
Dispatches a statement PDF to the correct bank-specific parser based
on auto-detection.
"""

from . import canara, sbi, hdfc, icici, axis, generic, kvb
from .detector import detect_bank

PARSERS = {
    "canara": canara,
    "sbi": sbi,
    "hdfc": hdfc,
    "icici": icici,
    "axis": axis,
    "kvb": kvb,
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
    "kvb": "Karur Vysya Bank",
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
    Uses a multi-tier fallback pipeline so any bank statement format works.
    Returns (bank_key, bank_display_name, transactions).
    """
    detected_key = detect_bank(pdf_path, password=password)
    bank_key = forced_bank or detected_key

    # 1. Primary Attempt: Use requested/detected parser
    if bank_key in PARSERS:
        try:
            txns = PARSERS[bank_key].parse(pdf_path, password=password)
            if txns:
                return bank_key, BANK_DISPLAY_NAMES.get(bank_key, bank_key), txns
        except Exception:
            pass

    # 2. Fallback Attempt 1: Canara text-based parser
    try:
        txns = canara.parse(pdf_path, password=password)
        if txns:
            key = bank_key if bank_key != "unknown" else "canara"
            return key, BANK_DISPLAY_NAMES.get(key, "Canara Bank"), txns
    except Exception:
        pass

    # 3. Fallback Attempt 2: Generic grid-table parser
    try:
        txns = generic.parse(pdf_path, password=password)
        if txns:
            key = bank_key if bank_key != "unknown" else "other"
            return key, BANK_DISPLAY_NAMES.get(key, "Bank Statement"), txns
    except Exception:
        pass

    display_bank = BANK_DISPLAY_NAMES.get(bank_key, bank_key)
    raise UnsupportedBankError(
        f"Could not extract transactions from this statement (Bank: {display_bank}). "
        f"If the file is password-protected, please enter the password and try again."
    )


def parse_text_statement(text_content, forced_bank=None):
    """Detect bank from text and parse raw copy-pasted statement text.
    Returns (bank_key, bank_display_name, transactions).
    """
    from .copypaste import parse_text
    from .detector import detect_bank_from_text

    detected_key = detect_bank_from_text(text_content)
    bank_key = forced_bank or detected_key

    transactions = parse_text(text_content)
    return bank_key, BANK_DISPLAY_NAMES.get(bank_key, "Unknown Bank"), transactions

