"""
Generic fallback parser.

Routes to the generic table parser without overriding any column aliases.
This is used for all newly added Indian banks where we don't have a 
specific sample to verify column headers.
"""

from .generic_table import parse_generic

def parse(pdf_path, password=None):
    return parse_generic(pdf_path, password=password)
