"""
Detects which bank issued a statement PDF by scanning the first couple
of pages for identifying text (bank name, IFSC prefix, etc).
"""

import re

import pdfplumber

# Order matters only in that first match wins; keywords are chosen to
# avoid collisions with each other.
BANK_SIGNATURES = [
    ("canara", [r"canara\s+bank", r"\bcnrb\b"]),
    ("sbi", [r"state\s+bank\s+of\s+india", r"\bsbin\b", r"\bsbi\b"]),
    ("hdfc", [r"hdfc\s+bank", r"\bhdfc\d", r"\bhdfc0\b"]),
    ("icici", [r"icici\s+bank", r"\bicic\b"]),
    ("axis", [r"axis\s+bank", r"\butib\b"]),
    ("pnb", [r"punjab\s+national\s+bank", r"\bpunb\b"]),
    ("bob", [r"bank\s+of\s+baroda", r"\bbarb\b"]),
    ("kotak", [r"kotak\s+mahindra", r"\bkkbk\b"]),
    ("indusind", [r"indusind\s+bank", r"\bindb\b"]),
    ("union", [r"union\s+bank\s+of\s+india", r"\bubin\b"]),
    ("idfc", [r"idfc\s+first\s+bank", r"\bidfb\b"]),
    ("yes", [r"yes\s+bank", r"\byesb\b"]),
    ("boi", [r"bank\s+of\s+india", r"\bbkid\b"]),
    ("cbi", [r"central\s+bank\s+of\s+india", r"\bcbin\b"]),
    ("iob", [r"indian\s+overseas\s+bank", r"\bioba\b"]),
    ("uco", [r"uco\s+bank", r"\bucba\b"]),
    ("federal", [r"federal\s+bank", r"\bfdrl\b"]),
    ("southindian", [r"south\s+indian\s+bank", r"\bsibl\b"]),
    ("indian", [r"indian\s+bank", r"\bidib\b"]),
]


def detect_bank(pdf_path, password=None):
    """Return one of 'canara', 'sbi', 'hdfc', 'icici', 'axis', or
    'unknown' based on text found in the first 2 pages of the PDF."""
    text = ""
    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page in pdf.pages[:2]:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    text_lower = text.lower()

    for bank_name, patterns in BANK_SIGNATURES:
        for pattern in patterns:
            if re.search(pattern, text_lower):
                return bank_name

    return "unknown"
