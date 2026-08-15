"""
Parser for copy-pasted bank statement text.
Supports both structured table columns (tab/double-space separated) 
and line-by-line heuristic parsing.
"""

import re
from datetime import datetime

# Regexes for dates
DATE_PATTERNS = [
    (re.compile(r'\b(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{2,4})\b'), 'numeric'),
    (re.compile(r'\b(\d{1,2})[-/\.\s]+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-zA-Z]*[-/\.\s]+(\d{2,4})\b', re.IGNORECASE), 'alphabetic'),
    (re.compile(r'\b(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})\b'), 'iso'),
    (re.compile(r'\b(\d{1,2})[-/\.\s]*(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)', re.IGNORECASE), 'alphabetic_no_year')
]

MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}

def extract_date_and_clean_line(line):
    parsed_date = None
    cleaned_line = line
    
    for pattern, pat_type in DATE_PATTERNS:
        match = pattern.search(cleaned_line)
        if match:
            try:
                if pat_type == 'numeric':
                    d, m, y = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    if y < 100:
                        y += 2000 if y > 50 else 1900
                    parsed_date = datetime(y, m, d)
                elif pat_type == 'alphabetic':
                    d = int(match.group(1))
                    m_str = match.group(2).lower()[:3]
                    m = MONTH_MAP.get(m_str, 1)
                    y = int(match.group(3))
                    if y < 100:
                        y += 2000
                    parsed_date = datetime(y, m, d)
                elif pat_type == 'iso':
                    y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    parsed_date = datetime(y, m, d)
                elif pat_type == 'alphabetic_no_year':
                    d = int(match.group(1))
                    m_str = match.group(2).lower()[:3]
                    m = MONTH_MAP.get(m_str, 1)
                    y = datetime.now().year
                    parsed_date = datetime(y, m, d)
                
                # Remove all matching dates from the line to clean narration
                cleaned_line = pattern.sub('', cleaned_line)
                break
            except Exception:
                continue
    return parsed_date, cleaned_line

def _clean_amount_str(s):
    s = s.replace(',', '').replace('₹', '').replace('Rs.', '').replace('INR', '').strip()
    s = re.sub(r'\b(Dr|Cr|DR|CR)\b', '', s).strip()
    if not s or s in ('-', '--'):
        return 0.0
    
    is_neg = False
    if s.startswith('(') and s.endswith(')'):
        is_neg = True
        s = s[1:-1].strip()
        
    try:
        val = float(s)
        return -val if is_neg else abs(val)
    except ValueError:
        return 0.0

def _parse_date_basic(s):
    s = s.strip()
    from .generic_table import DATE_FORMATS
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None

def try_parse_as_table(lines):
    from .generic_table import DEFAULT_COLUMN_ALIASES
    
    header_idx = -1
    col_mapping = None
    delimiter = None
    
    for idx, line in enumerate(lines[:10]):
        delims = []
        if '\t' in line:
            delims.append('\t')
        if '  ' in line:
            delims.append(r' {2,}')
        if ',' in line:
            delims.append(',')
            
        for delim in delims:
            parts = [p.strip().lower() for p in re.split(delim, line) if p.strip()]
            if len(parts) >= 3:
                has_date = False
                has_narr = False
                has_amt = False
                for p in parts:
                    if any(a in p for a in DEFAULT_COLUMN_ALIASES['date']):
                        has_date = True
                    if any(a in p for a in DEFAULT_COLUMN_ALIASES['narration']):
                        has_narr = True
                    if any(a in p for a in DEFAULT_COLUMN_ALIASES['debit'] + DEFAULT_COLUMN_ALIASES['credit']):
                        has_amt = True
                        
                if has_date and (has_narr or has_amt):
                    header_idx = idx
                    delimiter = delim
                    col_mapping = {}
                    for col_name, aliases in DEFAULT_COLUMN_ALIASES.items():
                        for col_idx, part in enumerate(parts):
                            if part == col_name or part in aliases or any(a in part for a in aliases):
                                col_mapping[col_name] = col_idx
                                break
                    break
        if header_idx != -1:
            break
            
    if header_idx == -1 or col_mapping is None or 'date' not in col_mapping:
        return None
        
    transactions = []
    for idx, line in enumerate(lines):
        if idx == header_idx:
            continue
            
        parts = [p.strip() for p in re.split(delimiter, line)]
        max_required_idx = max(col_mapping.values())
        if len(parts) <= max_required_idx:
            continue
            
        dt_str = parts[col_mapping['date']]
        dt = _parse_date_basic(dt_str)
        if not dt:
            parsed_dt, _ = extract_date_and_clean_line(dt_str)
            if parsed_dt:
                dt = parsed_dt
            else:
                continue
                
        narration = parts[col_mapping['narration']] if 'narration' in col_mapping else ""
        debit = _clean_amount_str(parts[col_mapping['debit']]) if 'debit' in col_mapping else 0.0
        credit = _clean_amount_str(parts[col_mapping['credit']]) if 'credit' in col_mapping else 0.0
        balance = _clean_amount_str(parts[col_mapping['balance']]) if 'balance' in col_mapping else None
        
        transactions.append({
            "date": dt,
            "narration": narration.strip(),
            "debit": debit,
            "credit": credit,
            "balance": balance
        })
        
    return transactions

def parse_line_heuristically(line):
    original_line = line  # keep for Dr/Cr detection
    dt, cleaned = extract_date_and_clean_line(line)
    if not dt:
        return None
        
    cleaned = cleaned.strip()
    
    has_delimiters = '\t' in cleaned or '  ' in cleaned
    if has_delimiters:
        parts = [p.strip() for p in re.split(r'\t| {2,}', cleaned) if p.strip()]
        amounts = []
        narration_parts = []
        for part in parts:
            cleaned_part = re.sub(r'[₹Rs\$+]', '', part).strip()
            if re.match(r'^-?\d+(?:,\d+)*(?:\.\d{1,2})?$', cleaned_part):
                val_str = cleaned_part.replace(',', '')
                if '.' not in val_str and len(val_str) >= 8:
                    narration_parts.append(part)
                else:
                    try:
                        val = float(val_str)
                        is_dr = bool(re.search(r'\b(dr|debit|wdl|sent|paid)\b', part, re.IGNORECASE))
                        is_cr = bool(re.search(r'\b(cr|credit|dep|rec|interest)\b', part, re.IGNORECASE))
                        amounts.append((val, is_dr, is_cr, part))
                    except ValueError:
                        narration_parts.append(part)
            else:
                narration_parts.append(part)
                
        if len(amounts) > 0:
            narration = " ".join(narration_parts).strip()
            debit = 0.0
            credit = 0.0
            balance = None
            
            if len(amounts) == 1:
                val, is_dr, is_cr, _ = amounts[0]
                narr_lower = narration.lower()
                if is_dr or any(w in narr_lower for w in ['debit', 'withdrawal', 'to', 'chg', 'charge', 'paid', 'sent', 'transfer to']):
                    debit = val
                elif is_cr or any(w in narr_lower for w in ['credit', 'deposit', 'from', 'int', 'interest', 'received']):
                    credit = val
                else:
                    debit = val
            elif len(amounts) == 2:
                val1, _, cr1, _ = amounts[0]
                val2, _, _, _ = amounts[1]
                narr_lower = narration.lower()
                if cr1 or any(w in narr_lower for w in ['credit', 'deposit', 'from', 'int', 'interest', 'received']):
                    credit = val1
                else:
                    debit = val1
                balance = val2
            else:
                debit = amounts[0][0]
                credit = amounts[1][0]
                balance = amounts[2][0]
                
            return {
                "date": dt,
                "narration": narration,
                "debit": debit,
                "credit": credit,
                "balance": balance
            }

    matches = list(re.finditer(r'\b\d+(?:,\d+)*(?:\.\d+)?\b', cleaned))
    if not matches:
        return {
            "date": dt,
            "narration": cleaned.strip(),
            "debit": 0.0,
            "credit": 0.0,
            "balance": None
        }
        
    # Traverse matches from right to left to collect amounts, stopping at letters
    valid_amounts = []
    last_idx = len(cleaned)
    for m in reversed(matches):
        start, end = m.start(), m.end()
        interstitial = cleaned[end:last_idx]
        if re.search(r'[a-zA-Z]', interstitial):
            break
            
        s = m.group(0).replace(',', '')
        if '.' not in s and len(s) >= 8:
            break
            
        try:
            val = float(s)
            valid_amounts.insert(0, (val, start, end))
            last_idx = start
        except ValueError:
            break
            
    if not valid_amounts:
        return {
            "date": dt,
            "narration": cleaned.strip(),
            "debit": 0.0,
            "credit": 0.0,
            "balance": None
        }
        
    first_amount_start = valid_amounts[0][1]
    narration = cleaned[:first_amount_start].strip()
    
    amount_vals = [x[0] for x in valid_amounts]
    debit = 0.0
    credit = 0.0
    balance = None
    
    # If the line represents a balance row, don't treat the balance as a transaction amount
    narr_lower = narration.lower()
    orig_lower = original_line.lower()
    is_bal_row = any(k in narr_lower for k in ['closing balance', 'opening balance', 'brought forward', 'b/f', 'balance b/d', 'balance c/f', 'balance c/d', 'bal b/f'])
    
    # Check explicit Dr/Cr labels anywhere in the original line
    has_explicit_cr = bool(re.search(r'\b(cr|credit|credited|deposit|received|int|interest|\+)\b', orig_lower))
    has_explicit_dr = bool(re.search(r'\b(dr|debit|debited|withdrawal|wdl|paid|sent|transfer out|-)\b', orig_lower))
    
    if len(amount_vals) == 1:
        val = amount_vals[0]
        if is_bal_row:
            balance = val
        elif has_explicit_cr and not has_explicit_dr:
            credit = val
        elif has_explicit_dr and not has_explicit_cr:
            debit = val
        elif any(w in narr_lower for w in ['credit', 'deposit', 'from', 'int', 'interest', 'received']):
            credit = val
        else:
            debit = val
    elif len(amount_vals) == 2:
        val1, val2 = amount_vals[0], amount_vals[1]
        if is_bal_row:
            balance = val2
        else:
            if (has_explicit_cr and not has_explicit_dr) or any(w in narr_lower for w in ['credit', 'deposit', 'from', 'int', 'interest', 'received']):
                credit = val1
            else:
                debit = val1
            balance = val2
    elif len(amount_vals) >= 3:
        debit = amount_vals[0]
        credit = amount_vals[1]
        balance = amount_vals[2]
        
    return {
        "date": dt,
        "narration": narration,
        "debit": debit,
        "credit": credit,
        "balance": balance
    }

def preprocess_raw_text(text):
    if not text:
        return ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if len(lines) <= 2:
        # Match dates in format: DD-Month
        date_pattern = re.compile(
            r'\b\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|\b\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}', 
            re.IGNORECASE
        )
        # If not matched with word boundary due to direct concatenation, search without boundary
        matches = list(date_pattern.finditer(text))
        if len(matches) < 2:
            no_boundary_pattern = re.compile(
                r'\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}',
                re.IGNORECASE
            )
            matches = list(no_boundary_pattern.finditer(text))
            
        if len(matches) >= 2:
            split_parts = []
            last_pos = 0
            for m in matches:
                start_pos = m.start()
                if start_pos > last_pos:
                    split_parts.append(text[last_pos:start_pos].strip())
                last_pos = start_pos
            if last_pos < len(text):
                split_parts.append(text[last_pos:].strip())
            return "\n".join(split_parts)
    return text


def adjust_transaction_years(transactions):
    if not transactions:
        return transactions
        
    # Standardize years so they form a continuous chronological chain
    current_year = datetime.now().year
    last_month = transactions[-1]["date"].month
    current_month = datetime.now().month
    
    base_last_year = current_year
    if last_month > current_month:
        base_last_year = current_year - 1
        
    curr_yr = base_last_year
    transactions[-1]["date"] = transactions[-1]["date"].replace(year=curr_yr)
    
    for idx in range(len(transactions) - 2, -1, -1):
        prev_m = transactions[idx]["date"].month
        curr_m = transactions[idx + 1]["date"].month
        if prev_m > curr_m:
            curr_yr -= 1
        transactions[idx]["date"] = transactions[idx]["date"].replace(year=curr_yr)
        
    return transactions


def parse_text(raw_text):
    """Parse raw text blocks into standardized transactions list."""
    if not raw_text or not raw_text.strip():
        return []
        
    # Preprocess single-line pastes that contain multiple transactions
    raw_text = preprocess_raw_text(raw_text)
        
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    
    table_txns = try_parse_as_table(lines)
    if table_txns:
        return adjust_transaction_years(table_txns)
        
    transactions = []
    for line in lines:
        txn = parse_line_heuristically(line)
        if txn:
            transactions.append(txn)
            
    return adjust_transaction_years(transactions)

def parse(file_path, password=None):
    """File-based wrapper for copypaste parser."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read()
        return parse_text(content)
    except Exception as e:
        raise ValueError(f"Failed to read raw text statement: {e}")
