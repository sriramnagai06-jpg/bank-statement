"""
Bank Statement Analyzer - Flask backend.

Endpoints:
  GET  /                         -> serves the frontend
  POST /api/analyze              -> upload a PDF, get back bank name,
                                     month-wise summary JSON, and a
                                     download link for the Excel file
  GET  /api/download/<filename>  -> download a generated Excel file

Run:
  cd backend
  pip install -r requirements.txt
  python app.py
Then open http://localhost:5000 in a browser.
"""

import os
import time
import uuid

from flask import Flask, jsonify, request, send_from_directory

from excel.exporter import export_to_excel, summarize
from parser.parser_manager import UnsupportedBankError, parse_statement, parse_text_statement, PARSERS

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "frontend")

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path="")

MAX_UPLOAD_MB = 50
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


# ---- CORS headers (needed for tunnel/public URL access from phones) ----

@app.after_request
def add_cors_and_security_headers(response):
    """Add CORS headers so the API works when accessed via public tunnel,
    different Wi-Fi, or from a phone browser pointing at a public URL."""
    origin = request.headers.get("Origin", "*")
    response.headers["Access-Control-Allow-Origin"] = origin
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, bypass-tunnel-reminder, ngrok-skip-browser-warning, serveo-skip-browser-warning"
    response.headers["Access-Control-Max-Age"] = "3600"
    # Prevent MIME-type sniffing
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({
        "error": f"File is too large. Maximum allowed size is {MAX_UPLOAD_MB} MB. Please upload a smaller file."
    }), 413


@app.errorhandler(500)
def handle_500_error(error):
    """Ensure all internal server errors return a clean JSON response instead of HTML."""
    msg = str(error.original_exception) if hasattr(error, 'original_exception') else str(error)
    return jsonify({
        "error": f"Server processing error: {msg}"
    }), 500


# ---- Handle CORS preflight requests ----

@app.route("/api/analyze", methods=["OPTIONS"])
@app.route("/api/download/<path:filename>", methods=["OPTIONS"])
def cors_preflight(*args, **kwargs):
    response = jsonify({"status": "ok"})
    return response, 200


def cleanup_old_uploads(max_age_seconds=3600):
    """Delete uploaded/generated files older than max_age_seconds."""
    try:
        now = time.time()
        for fname in os.listdir(UPLOAD_DIR):
            fpath = os.path.join(UPLOAD_DIR, fname)
            if os.path.isfile(fpath):
                if now - os.path.getmtime(fpath) > max_age_seconds:
                    try:
                        os.remove(fpath)
                    except OSError:
                        pass
    except Exception:
        pass


def filter_and_recalculate_transactions(transactions, skip_keyword_intercompany=False):
    import re
    
    def is_charge_transaction(narration):
        if not narration:
            return False
        narration_lower = narration.lower()
        keywords = ["chg", "charge", "fee", "tax", "gst", "commission", "folio amt"]
        if any(k in narration_lower for k in keywords):
            return True
        if re.search(r"\bsc\b", narration_lower):
            return True
        return False

    def is_inter_company_transaction(narration):
        if not narration:
            return False
        narration_lower = narration.lower()
        inter_company_keywords = ["lakshmi", "senthil", "mahalakshmi"]
        return any(k in narration_lower for k in inter_company_keywords)

    filtered = []
    for t in transactions:
        narration = t.get("narration", "")
        debit = t.get("debit") or 0.0
        credit = t.get("credit") or 0.0
        
        # Remove charges
        if is_charge_transaction(narration):
            continue
            
        # Inter-company logic
        if not skip_keyword_intercompany and is_inter_company_transaction(narration):
            # Deposits are accepted (kept)
            if credit > 0:
                filtered.append(t)
            # Withdrawals are removed (skipped)
            elif debit > 0:
                continue
            else:
                filtered.append(t)
        else:
            filtered.append(t)

    # Recalculate running balance
    if filtered:
        # Find first transaction with a balance to act as the starting point
        start_idx = -1
        for i, t in enumerate(filtered):
            if t.get("balance") is not None:
                start_idx = i
                break
        
        if start_idx != -1:
            curr_balance = filtered[start_idx]["balance"]
            for idx in range(start_idx + 1, len(filtered)):
                t = filtered[idx]
                debit = t.get("debit") or 0.0
                credit = t.get("credit") or 0.0
                curr_balance = round(curr_balance - debit + credit, 2)
                t["balance"] = curr_balance

    return filtered


def extract_printed_totals_from_pdf(pdf_path):
    """
    Search for printed totals in the statement PDF text (e.g. Canara Bank style).
    """
    import pdfplumber
    import re
    
    totals = {"debit": None, "credit": None}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
                    
        # 1. Look for Statement Summary pattern in Canara Bank style
        if "statement summary" in full_text.lower():
            lines = full_text.splitlines()
            for idx, line in enumerate(lines):
                if "statement summary" in line.lower():
                    # The values are typically 2 or 3 lines below
                    for offset in range(1, 6):
                        if idx + offset >= len(lines):
                            break
                        cand_line = lines[idx + offset].strip()
                        parts = cand_line.split()
                        amounts = []
                        for p in parts:
                            p_clean = p.replace(",", "").strip()
                            if re.match(r"^\d+\.\d{2}$", p_clean):
                                amounts.append(float(p_clean))
                        if len(amounts) >= 3:
                            totals["debit"] = amounts[1]
                            totals["credit"] = amounts[2]
                            return totals
                            
        # 2. General search
        deb_match = re.search(r"total\s+debits?\s*(?:amount)?\s*:?\s*([\d,]+\.\d{2})", full_text, re.IGNORECASE)
        if deb_match:
            totals["debit"] = float(deb_match.group(1).replace(",", ""))
            
        cred_match = re.search(r"total\s+credits?\s*(?:amount)?\s*:?\s*([\d,]+\.\d{2})", full_text, re.IGNORECASE)
        if cred_match:
            totals["credit"] = float(cred_match.group(1).replace(",", ""))
            
    except Exception:
        pass
    return totals


def reconcile_transactions(transactions):
    """
    Determine Debit/Credit based on running balance differences.
    Also checks for opening and closing balances correctness.
    """
    if not transactions:
        return [], "pass", [], ""

    txns = sorted(transactions, key=lambda x: x["date"])

    first_txn = txns[0]
    opening_balance = 0.0
    opening_known = False  # True only if we have an explicit B/F / opening balance row

    narr_lower = (first_txn.get("narration") or "").lower()
    is_bf = any(k in narr_lower for k in ["b/f", "brought forward", "opening balance", "balance b/d", "bal b/f"])

    if is_bf and first_txn.get("balance") is not None:
        opening_balance = first_txn["balance"]
        opening_known = True
        first_txn["debit"] = 0.0
        first_txn["credit"] = 0.0
        start_idx = 1
    else:
        curr_bal = first_txn.get("balance")
        deb = first_txn.get("debit") or 0.0
        cred = first_txn.get("credit") or 0.0
        if curr_bal is not None:
            # Derive opening from first txn — not guaranteed to be the real opening
            opening_balance = round(curr_bal + deb - cred, 2)
        start_idx = 0

    running_balance = opening_balance
    unreconciled_logs = []
    status = "pass"

    for idx in range(start_idx, len(txns)):
        t = txns[idx]
        curr_bal = t.get("balance")
        deb = t.get("debit") or 0.0
        cred = t.get("credit") or 0.0
        amount = deb or cred

        # If it's a balance row (e.g. closing balance row), skip
        t_narr_lower = (t.get("narration") or "").lower()
        if any(k in t_narr_lower for k in ["closing balance", "balance c/f", "balance c/d"]):
            t["debit"] = 0.0
            t["credit"] = 0.0
            if curr_bal is not None:
                running_balance = curr_bal
            continue

        if curr_bal is not None:
            expected_credit_bal = round(running_balance + amount, 2)
            expected_debit_bal = round(running_balance - amount, 2)
            actual_bal = round(curr_bal, 2)

            # Check Credit
            if abs(expected_credit_bal - actual_bal) <= 0.02:
                t["credit"] = amount
                t["debit"] = 0.0
                running_balance = curr_bal
            # Check Debit
            elif abs(expected_debit_bal - actual_bal) <= 0.02:
                t["debit"] = amount
                t["credit"] = 0.0
                running_balance = curr_bal
            # Check 0-amount rows
            elif amount == 0.0 and abs(running_balance - actual_bal) <= 0.02:
                t["debit"] = 0.0
                t["credit"] = 0.0
                running_balance = curr_bal
            else:
                # Could be a gap in statement (weekends, missing pages)
                # Only flag as fail if opening was explicitly known
                if opening_known:
                    status = "fail"
                unreconciled_logs.append(
                    f"UNRECONCILED: [{t['date'].strftime('%d-%m-%Y')}, {t['narration'][:40]}, Amount: {amount:.2f}]"
                )
                # Reset running balance to the stated balance to continue the chain
                running_balance = curr_bal
        else:
            unreconciled_logs.append(
                f"UNRECONCILED: [{t['date'].strftime('%d-%m-%Y')}, {t['narration'][:40]}, No Balance]"
            )

    closing_balance = running_balance
    total_credits = sum(t.get("credit") or 0.0 for t in txns[start_idx:])
    total_debits = sum(t.get("debit") or 0.0 for t in txns[start_idx:])
    expected_closing = round(opening_balance + total_credits - total_debits, 2)

    verification_warning = ""
    # Only do the summary math check when we have a known opening balance (B/F row)
    if opening_known and abs(expected_closing - closing_balance) > 0.02:
        status = "fail"
        verification_warning = (
            f"Math verification failed: Opening ({opening_balance:.2f}) + "
            f"Credits ({total_credits:.2f}) - Debits ({total_debits:.2f}) = "
            f"{expected_closing:.2f}, but Closing Balance is {closing_balance:.2f}"
        )
    elif not opening_known and unreconciled_logs:
        # Warn but don't fail — likely statement gaps, not parsing errors
        verification_warning = (
            f"Note: {len(unreconciled_logs)} row(s) could not be matched by balance chain "
            f"(may be due to missing pages or statement gaps). "
            f"Chain integrity: Opening (derived {opening_balance:.2f}) → Closing {closing_balance:.2f}"
        )

    return txns, status, unreconciled_logs, verification_warning


def find_inter_company_transactions(txns_A, txns_B):
    """
    Matches inter-company transfers between Statement A and Statement B.
    """
    matched_pairs = []
    
    # Initialize fields
    for t in txns_A:
        t["inter_company"] = 0.0
        t["inter_company_ref"] = None
    for t in txns_B:
        t["inter_company"] = 0.0
        t["inter_company_ref"] = None
        
    matched_B_indices = set()
    
    for i, t_A in enumerate(txns_A):
        deb_A = t_A.get("debit") or 0.0
        cred_A = t_A.get("credit") or 0.0
        
        if deb_A == 0.0 and cred_A == 0.0:
            continue
            
        for j, t_B in enumerate(txns_B):
            if j in matched_B_indices:
                continue
                
            deb_B = t_B.get("debit") or 0.0
            cred_B = t_B.get("credit") or 0.0
            
            # Check opposite directions and matching amount
            if (deb_A > 0 and cred_B == deb_A) or (cred_A > 0 and deb_B == cred_A):
                # Date window +/- 2 days
                days_diff = abs((t_A["date"] - t_B["date"]).days)
                if days_diff <= 2:
                    matched_B_indices.add(j)
                    
                    amount = deb_A if deb_A > 0 else cred_A
                    t_A["inter_company"] = amount
                    t_B["inter_company"] = amount
                    
                    ref_A = f"StmtA_{t_A['date'].strftime('%d%b')}_{amount:.0f}"
                    ref_B = f"StmtB_{t_B['date'].strftime('%d%b')}_{amount:.0f}"
                    
                    t_A["inter_company_ref"] = ref_B
                    t_B["inter_company_ref"] = ref_A
                    
                    desc_A = t_A['narration'][:25].replace('\n', ' ')
                    desc_B = t_B['narration'][:25].replace('\n', ' ')
                    
                    matched_pairs.append(
                        f"Pair: [{t_A['date'].strftime('%d-%m-%Y')} StmtA ({desc_A}) ↔ {t_B['date'].strftime('%d-%m-%Y')} StmtB ({desc_B}) Amount: {amount:.2f}]"
                    )
                    break
                    
    return txns_A, txns_B, matched_pairs


def summarize_combined(txns_A, txns_B=None):
    """
    Generate a combined summary table. Month-wise external credits and debits exclude inter-company transactions.
    """
    import pandas as pd
    from excel.exporter import to_dataframe
    
    df_A = to_dataframe(txns_A)
    if txns_B:
        df_B = to_dataframe(txns_B)
        combined_df = pd.concat([df_A, df_B], ignore_index=True)
    else:
        combined_df = df_A
        
    if combined_df.empty:
        return []
        
    month_labels = (
        combined_df[["Year", "Month", "MonthLabel"]]
        .drop_duplicates()
        .sort_values(["Year", "Month"])["MonthLabel"]
        .tolist()
    )
    
    summary_rows = []
    for label in month_labels:
        month_df = combined_df[combined_df["MonthLabel"] == label]
        
        def sum_type(t, col):
            val = month_df.loc[month_df["Type"] == t, col].sum()
            return round(float(val), 2)
            
        raw_credit = sum_type("Credit", "Credit")
        raw_debit = sum_type("Debit", "Debit")
        
        # Inter-company logic
        ic_credit_sum = 0.0
        ic_debit_sum = 0.0
        if "inter_company" in month_df.columns:
            ic_credit_sum = round(float(month_df.loc[month_df["Type"] == "Credit", "inter_company"].sum()), 2)
            ic_debit_sum = round(float(month_df.loc[month_df["Type"] == "Debit", "inter_company"].sum()), 2)
            
        ic_total = max(ic_credit_sum, ic_debit_sum)
        ext_credit = max(0.0, round(raw_credit - ic_credit_sum, 2))
        ext_debit = max(0.0, round(raw_debit - ic_debit_sum, 2))
        
        summary_rows.append({
            "Month": label,
            "Credit": ext_credit,
            "Debit": ext_debit,
            "Credit Charge": sum_type("Credit Charge", "Credit"),
            "Debit Charge": sum_type("Debit Charge", "Debit"),
            "Credit Interest": sum_type("Credit Interest", "Credit"),
            "Debit Interest": sum_type("Debit Interest", "Debit"),
            "Inter-Company Transactions": ic_total,
            "Transaction Count": len(month_df),
        })
        
    totals = {"Month": "TOTAL (FY)"}
    for key in ["Credit", "Debit", "Credit Charge", "Debit Charge", "Credit Interest", "Debit Interest", "Inter-Company Transactions"]:
        totals[key] = round(sum(r[key] for r in summary_rows), 2)
    totals["Transaction Count"] = sum(r["Transaction Count"] for r in summary_rows)
    summary_rows.append(totals)
    
    return summary_rows


@app.route("/")
def index():
    return app.send_static_file("index.html")


# Serve manifest.json explicitly
@app.route("/manifest.json")
def manifest():
    return app.send_static_file("manifest.json")


@app.route("/api/analyze", methods=["POST"])
def analyze():
    cleanup_old_uploads()

    forced_bank = request.form.get("bank")
    if forced_bank == "auto":
        forced_bank = None
    if forced_bank and forced_bank not in PARSERS:
        return jsonify({"error": f"Unsupported bank '{forced_bank}'."}), 400

    text_content = request.form.get("text_content")
    job_id = uuid.uuid4().hex[:10]

    # We will collect statements to process
    statements_data = [] # List of tuples: (bank_key, bank_display_name, raw_txns, printed_totals)
    unreconciled_logs_all = []
    reconciliation_status = "pass"
    verification_warnings = []

    if text_content and text_content.strip():
        try:
            bank_key, bank_display_name, txns = parse_text_statement(text_content, forced_bank=forced_bank)
            if not txns:
                return jsonify({
                    "error": (
                        "No transactions could be parsed from the pasted text. "
                        "Please make sure each line starts with a date (e.g. 12/08/2026 or 13-Aug-2026) "
                        "followed by a narration and amount. "
                        "Use Dr / Cr labels for best accuracy. "
                        "Supported formats:\n"
                        "  • Date  Narration  Debit  Credit  Balance (space/tab-separated)\n"
                        "  • Date Narration Dr/Cr Amount\n"
                        "  • Tab-separated with a header row (Date, Narration, Debit, Credit, Balance)"
                    )
                }), 422
            txns, status, logs, warn = reconcile_transactions(txns)
            if status == "fail":
                reconciliation_status = "fail"
            if warn:
                verification_warnings.append(warn)
            unreconciled_logs_all.extend(logs)
            
            statements_data.append((bank_key, bank_display_name, txns, {"debit": None, "credit": None}))
        except Exception as e:
            return jsonify({"error": f"Failed to parse text: {str(e)}"}), 500
    else:
        files = request.files.getlist("statement")
        if not files or len(files) == 0 or (len(files) == 1 and files[0].filename == ""):
            return jsonify({"error": "No statement file or text was provided."}), 400

        if len(files) > 2:
            return jsonify({"error": "Maximum of 2 statement files can be processed together."}), 400

        password = request.form.get("password")
        for file in files:
            if file.filename == "":
                continue
            file_ext = os.path.splitext(file.filename)[1].lower()
            if file_ext not in [".pdf", ".xlsx", ".xls"]:
                file_ext = ".pdf"
            
            file_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex[:10]}{file_ext}")
            file.save(file_path)

            try:
                # Verify header bytes
                with open(file_path, "rb") as f:
                    header = f.read(1024)
                if file_ext == ".pdf":
                    if b"%PDF-" not in header:
                        os.remove(file_path)
                        return jsonify({"error": f"The file '{file.filename}' is corrupted or not a valid PDF statement."}), 400
                else:
                    if b"PK\x03\x04" not in header and not file_path.endswith(".xls"):
                        os.remove(file_path)
                        return jsonify({"error": f"The file '{file.filename}' is corrupted or not a valid Excel statement."}), 400
                
                # Parse
                bank_key, bank_display_name, txns = parse_statement(file_path, forced_bank=forced_bank, password=password)
                
                # Extract printed totals
                printed_totals = {"debit": None, "credit": None}
                if file_ext == ".pdf":
                    printed_totals = extract_printed_totals_from_pdf(file_path)
                
                # Reconcile using balance math
                txns, status, logs, warn = reconcile_transactions(txns)
                if status == "fail":
                    reconciliation_status = "fail"
                if warn:
                    verification_warnings.append(f"File '{file.filename}': {warn}")
                unreconciled_logs_all.extend(logs)
                
                # Printed totals cross check
                # Exclude the B/F row credit if it exists
                first_narr = (txns[0].get("narration") or "").lower() if txns else ""
                is_bf = any(k in first_narr for k in ["b/f", "brought forward", "opening balance", "balance b/d", "bal b/f"])
                start_check_idx = 1 if is_bf else 0
                
                raw_credit_sum = sum(t.get("credit") or 0.0 for t in txns[start_check_idx:])
                raw_debit_sum = sum(t.get("debit") or 0.0 for t in txns[start_check_idx:])
                
                if printed_totals["credit"] is not None:
                    if abs(raw_credit_sum - printed_totals["credit"]) > 0.02:
                        verification_warnings.append(
                            f"File '{file.filename}': Printed Total Credits ({printed_totals['credit']:.2f}) does not match parsed credits sum ({raw_credit_sum:.2f})"
                        )
                if printed_totals["debit"] is not None:
                    if abs(raw_debit_sum - printed_totals["debit"]) > 0.02:
                        verification_warnings.append(
                            f"File '{file.filename}': Printed Total Debits ({printed_totals['debit']:.2f}) does not match parsed debits sum ({raw_debit_sum:.2f})"
                        )
                
                statements_data.append((bank_key, bank_display_name, txns, printed_totals))
            except UnsupportedBankError as e:
                return jsonify({"error": f"File '{file.filename}': {str(e)}"}), 422
            except Exception as e:
                err_msg = str(e)
                if "password" in err_msg.lower() or "encrypted" in err_msg.lower() or "protected" in err_msg.lower():
                    return jsonify({"error": f"File '{file.filename}' is password protected. Please enter the correct password and try again."}), 401
                return jsonify({"error": f"Failed to parse '{file.filename}': {err_msg}"}), 500
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)

    if not statements_data:
        return jsonify({"error": "No statement was parsed successfully."}), 400

    # Match inter-company transactions if there are 2 statements
    matched_pairs = []
    if len(statements_data) == 2:
        txns_A = statements_data[0][2]
        txns_B = statements_data[1][2]
        txns_A, txns_B, matched_pairs = find_inter_company_transactions(txns_A, txns_B)
        
        # Apply filters (GST/charges), skipping keyword-based intercompany filtering
        filtered_A = filter_and_recalculate_transactions(txns_A, skip_keyword_intercompany=True)
        filtered_B = filter_and_recalculate_transactions(txns_B, skip_keyword_intercompany=True)
        
        if not filtered_A and not filtered_B:
            return jsonify({"error": "All transactions in both statements were filtered out."}), 422
            
        combined_summary = summarize_combined(filtered_A, filtered_B)
        
        xlsx_name = f"{job_id}_consolidated.xlsx"
        xlsx_path = os.path.join(UPLOAD_DIR, xlsx_name)
        export_to_excel(filtered_A, xlsx_path, transactions_B=filtered_B, matched_pairs=matched_pairs, unreconciled_logs=unreconciled_logs_all)
        
        bank_name = f"Consolidated ({statements_data[0][1]} & {statements_data[1][1]})"
        total_txns_count = len(filtered_A) + len(filtered_B)
    else:
        # Single statement
        bank_key, bank_display_name, txns, _ = statements_data[0]
        filtered_txns = filter_and_recalculate_transactions(txns, skip_keyword_intercompany=False)
        if not filtered_txns:
            return jsonify({"error": "All transactions were filtered out based on the removal rules."}), 422
            
        combined_summary = summarize_combined(filtered_txns)
        
        xlsx_name = f"{job_id}_monthwise.xlsx"
        xlsx_path = os.path.join(UPLOAD_DIR, xlsx_name)
        export_to_excel(filtered_txns, xlsx_path, matched_pairs=[], unreconciled_logs=unreconciled_logs_all)
        
        bank_name = bank_display_name
        total_txns_count = len(filtered_txns)

    # Serialize transactions for frontend display
    def serialize_txns(txns):
        result = []
        for t in txns:
            result.append({
                "date": t["date"].strftime("%d-%m-%Y") if hasattr(t.get("date"), "strftime") else str(t.get("date", "")),
                "narration": t.get("narration") or "",
                "debit": t.get("debit") or 0.0,
                "credit": t.get("credit") or 0.0,
                "balance": t.get("balance"),
            })
        return result

    if len(statements_data) == 2:
        all_txns_serialized = serialize_txns(filtered_A) + serialize_txns(filtered_B)
    else:
        all_txns_serialized = serialize_txns(filtered_txns)

    return jsonify({
        "bank": bank_name,
        "transaction_count": total_txns_count,
        "download_url": f"/api/download/{xlsx_name}",
        "summary": combined_summary,
        "transactions": all_txns_serialized,
        "reconciliation_status": reconciliation_status,
        "unreconciled_transactions": unreconciled_logs_all,
        "inter_company_matches": matched_pairs,
        "verification_warnings": verification_warnings
    })


@app.route("/api/download/<path:filename>")
def download(filename):
    if not filename.endswith(".xlsx") or "/" in filename or "\\" in filename:
        return jsonify({"error": "Invalid filename."}), 400
    response = send_from_directory(UPLOAD_DIR, filename, as_attachment=True,
                                    download_name="statement_monthwise.xlsx")
    # Ensure proper Content-Disposition for mobile downloads
    response.headers["Content-Disposition"] = 'attachment; filename="statement_monthwise.xlsx"'
    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return response


if __name__ == "__main__":
    app.run(host="0.0.0.0", debug=True, port=5000)
