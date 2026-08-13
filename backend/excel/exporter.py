"""
Takes a standardized list of transaction dicts (as returned by any
bank parser) and:
  1. Classifies each row as Debit/Credit x plain/Charge/Interest
     (narration contains 'CHG' -> Charge, contains 'INT' -> Interest)
  2. Writes an Excel workbook with one sheet per month + a Summary
     sheet with FY totals.
"""

import re

import pandas as pd
from openpyxl.utils import get_column_letter

CHARGE_RE = re.compile(r"\bCHG\b|CHARGE", re.IGNORECASE)
INTEREST_RE = re.compile(r"\bINT\b|INTEREST", re.IGNORECASE)


def classify_row(narration, debit, credit):
    side = "Debit" if debit > 0 else ("Credit" if credit > 0 else None)
    if side is None:
        return "Other"

    narration = narration or ""
    if INTEREST_RE.search(narration):
        category = "Interest"
    elif CHARGE_RE.search(narration):
        category = "Charge"
    else:
        category = None

    return f"{side} {category}" if category else side


def to_dataframe(transactions):
    df = pd.DataFrame(transactions)
    if df.empty:
        return df
    df["Type"] = df.apply(lambda r: classify_row(r["narration"], r["debit"], r["credit"]), axis=1)
    df["Month"] = df["date"].dt.month
    df["Year"] = df["date"].dt.year
    df["MonthLabel"] = df["date"].dt.strftime("%b-%Y")
    df = df.sort_values("date", kind="mergesort").reset_index(drop=True)
    df = df.rename(columns={
        "date": "Date", "narration": "Narration", "debit": "Debit",
        "credit": "Credit", "balance": "Balance",
    })
    return df


def _autofit(ws, df):
    for i, col in enumerate(df.columns, start=1):
        col_letter = get_column_letter(i)
        max_len = max(df[col].astype(str).map(len).max() if len(df) else 0, len(str(col)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

        # Apply number formatting for numeric columns (Debit, Credit, Balance, Charges, Interest, Inter-Company)
        if col in ["Debit", "Credit", "Balance", "Credit Charge", "Debit Charge", "Credit Interest", "Debit Interest", "Inter-Company Transactions"]:
            for cell in ws[col_letter][1:]:  # skip header
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"


def _month_summary_rows(df, month_labels):
    summary_rows = []
    for label in month_labels:
        month_df = df[df["MonthLabel"] == label]

        def sum_type(t, col):
            val = month_df.loc[month_df["Type"] == t, col].sum()
            return round(float(val), 2)

        summary_rows.append({
            "Month": label,
            "Credit": sum_type("Credit", "Credit"),
            "Debit": sum_type("Debit", "Debit"),
            "Credit Charge": sum_type("Credit Charge", "Credit"),
            "Debit Charge": sum_type("Debit Charge", "Debit"),
            "Credit Interest": sum_type("Credit Interest", "Credit"),
            "Debit Interest": sum_type("Debit Interest", "Debit"),
            "Transaction Count": len(month_df),
        })
    return summary_rows


def summarize(transactions):
    """Return JSON-serializable month-wise summary rows (for API responses)."""
    df = to_dataframe(transactions)
    if df.empty:
        return []

    month_labels = (
        df[["Year", "Month", "MonthLabel"]]
        .drop_duplicates()
        .sort_values(["Year", "Month"])["MonthLabel"]
        .tolist()
    )
    rows = _month_summary_rows(df, month_labels)

    totals = {"Month": "TOTAL (FY)"}
    for key in ["Credit", "Debit", "Credit Charge", "Debit Charge", "Credit Interest", "Debit Interest"]:
        totals[key] = round(sum(r[key] for r in rows), 2)
    totals["Transaction Count"] = sum(r["Transaction Count"] for r in rows)
    rows.append(totals)

    return rows


def summarize_combined(txns_A, txns_B=None):
    """
    Generate a combined summary table. Month-wise external credits and debits exclude inter-company transactions.
    """
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


def verify_balance_chain(transactions):
    """
    Verify the running balance chain of the transactions.
    Supports both ascending and descending chronological orders.
    Returns (status, mismatch_count) where status is one of:
      - "verified": balance chain is fully consistent in one direction.
      - "mismatches": running balance mismatches are detected.
      - "no_balances": statement has no balance data.
    """
    if not transactions:
        return "no_balances", 0

    has_any_balance = any(t.get("balance") is not None for t in transactions)
    if not has_any_balance:
        return "no_balances", 0

    # 1. Forward direction test
    forward_mismatches = 0
    prev_bal = None
    for t in transactions:
        curr_bal = t.get("balance")
        debit = t.get("debit") or 0.0
        credit = t.get("credit") or 0.0
        if curr_bal is not None:
            if prev_bal is not None:
                expected = round(prev_bal - debit + credit, 2)
                actual = round(curr_bal, 2)
                if abs(expected - actual) > 0.02:
                    forward_mismatches += 1
            prev_bal = curr_bal

    # 2. Backward direction test (reverse chronological order)
    backward_mismatches = 0
    prev_bal = None
    for t in reversed(transactions):
        curr_bal = t.get("balance")
        debit = t.get("debit") or 0.0
        credit = t.get("credit") or 0.0
        if curr_bal is not None:
            if prev_bal is not None:
                expected = round(prev_bal - debit + credit, 2)
                actual = round(curr_bal, 2)
                if abs(expected - actual) > 0.02:
                    backward_mismatches += 1
            prev_bal = curr_bal

    mismatches = min(forward_mismatches, backward_mismatches)
    if mismatches > 0:
        return "mismatches", mismatches
    return "verified", 0


def export_to_excel(transactions, output_path, transactions_B=None, matched_pairs=None, unreconciled_logs=None):
    df_A = to_dataframe(transactions)
    if df_A.empty:
        raise ValueError("No transactions to export.")
        
    if matched_pairs is None:
        matched_pairs = []
    if unreconciled_logs is None:
        unreconciled_logs = []

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if transactions_B is not None:
            # Dual Statement Mode
            df_B = to_dataframe(transactions_B)
            
            # 1. Summary Sheet
            summary_rows = summarize_combined(transactions, transactions_B)
            summary_df = pd.DataFrame(summary_rows)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            _autofit(writer.sheets["Summary"], summary_df)
            
            # 2. Statement A Sheet
            sheet_A_df = df_A[["Date", "Narration", "Type", "Debit", "Credit", "Balance"]].copy()
            if "inter_company" in df_A.columns:
                sheet_A_df["Inter-Company Transaction"] = df_A.apply(
                    lambda r: r["inter_company"] if r["inter_company"] > 0 else "No", axis=1
                )
            sheet_A_df["Date"] = sheet_A_df["Date"].dt.strftime("%d-%m-%Y")
            
            debit_sum_A = sheet_A_df["Debit"].sum()
            credit_sum_A = sheet_A_df["Credit"].sum()
            total_row_A = pd.DataFrame([{
                "Date": "TOTAL",
                "Narration": "",
                "Type": "",
                "Debit": debit_sum_A,
                "Credit": credit_sum_A,
                "Balance": None,
                "Inter-Company Transaction": ""
            }])
            sheet_A_df = pd.concat([sheet_A_df, total_row_A], ignore_index=True)
            sheet_A_df.to_excel(writer, sheet_name="Statement A", index=False)
            _autofit(writer.sheets["Statement A"], sheet_A_df)
            
            # 3. Statement B Sheet
            sheet_B_df = df_B[["Date", "Narration", "Type", "Debit", "Credit", "Balance"]].copy()
            if "inter_company" in df_B.columns:
                sheet_B_df["Inter-Company Transaction"] = df_B.apply(
                    lambda r: r["inter_company"] if r["inter_company"] > 0 else "No", axis=1
                )
            sheet_B_df["Date"] = sheet_B_df["Date"].dt.strftime("%d-%m-%Y")
            
            debit_sum_B = sheet_B_df["Debit"].sum()
            credit_sum_B = sheet_B_df["Credit"].sum()
            total_row_B = pd.DataFrame([{
                "Date": "TOTAL",
                "Narration": "",
                "Type": "",
                "Debit": debit_sum_B,
                "Credit": credit_sum_B,
                "Balance": None,
                "Inter-Company Transaction": ""
            }])
            sheet_B_df = pd.concat([sheet_B_df, total_row_B], ignore_index=True)
            sheet_B_df.to_excel(writer, sheet_name="Statement B", index=False)
            _autofit(writer.sheets["Statement B"], sheet_B_df)
            
            # Write matched pairs and unreconciled logs in Summary sheet
            ws = writer.sheets["Summary"]
            start_row = len(summary_df) + 4
            
            ws.cell(row=start_row, column=1, value="System Balance Verification Status:")
            status_A, _ = verify_balance_chain(transactions)
            status_B, _ = verify_balance_chain(transactions_B)
            ws.cell(row=start_row, column=2, value=f"Stmt A: {status_A.upper()} | Stmt B: {status_B.upper()}")
            
            start_row += 2
            ws.cell(row=start_row, column=1, value="Inter-Company Matched Pairs:")
            if matched_pairs:
                for pair in matched_pairs:
                    start_row += 1
                    ws.cell(row=start_row, column=1, value=pair)
            else:
                start_row += 1
                ws.cell(row=start_row, column=1, value="No matching inter-company transactions found.")
                
            if unreconciled_logs:
                start_row += 2
                ws.cell(row=start_row, column=1, value="Unreconciled Transactions:")
                for log in unreconciled_logs:
                    start_row += 1
                    ws.cell(row=start_row, column=1, value=log)
        else:
            # Single Statement Mode
            month_labels = (
                df_A[["Year", "Month", "MonthLabel"]]
                .drop_duplicates()
                .sort_values(["Year", "Month"])["MonthLabel"]
                .tolist()
            )
            for label in month_labels:
                month_df = df_A[df_A["MonthLabel"] == label].copy()
                sheet_df = month_df[["Date", "Narration", "Type", "Debit", "Credit", "Balance"]].copy()
                sheet_df["Date"] = sheet_df["Date"].dt.strftime("%d-%m-%Y")
                
                debit_sum = sheet_df["Debit"].sum()
                credit_sum = sheet_df["Credit"].sum()
                total_row = pd.DataFrame([{
                    "Date": "TOTAL",
                    "Narration": "",
                    "Type": "",
                    "Debit": debit_sum,
                    "Credit": credit_sum,
                    "Balance": None
                }])
                sheet_df = pd.concat([sheet_df, total_row], ignore_index=True)
                
                sheet_name = label[:31]
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                _autofit(writer.sheets[sheet_name], sheet_df)
                
            all_txns_df = df_A[["Date", "Narration", "Type", "Debit", "Credit", "Balance"]].copy()
            all_txns_df["Date"] = all_txns_df["Date"].dt.strftime("%d-%m-%Y")
            
            grand_debit_sum = all_txns_df["Debit"].sum()
            grand_credit_sum = all_txns_df["Credit"].sum()
            grand_total_row = pd.DataFrame([{
                "Date": "TOTAL",
                "Narration": "",
                "Type": "",
                "Debit": grand_debit_sum,
                "Credit": grand_credit_sum,
                "Balance": None
            }])
            all_txns_df = pd.concat([all_txns_df, grand_total_row], ignore_index=True)
            all_txns_df.to_excel(writer, sheet_name="All Transactions", index=False)
            _autofit(writer.sheets["All Transactions"], all_txns_df)
            
            summary_rows = _month_summary_rows(df_A, month_labels)
            summary_df = pd.DataFrame(summary_rows)
            totals = summary_df.drop(columns="Month").sum(numeric_only=True)
            totals["Month"] = "TOTAL (FY)"
            summary_df = pd.concat([summary_df, pd.DataFrame([totals])], ignore_index=True)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            _autofit(writer.sheets["Summary"], summary_df)
            
            writer.book.move_sheet("Summary", offset=-(len(writer.book.sheetnames) - 1))
            writer.book.move_sheet("All Transactions", offset=-(len(writer.book.sheetnames) - 2))
            
            # Write balance verification
            status, mismatch_count = verify_balance_chain(transactions)
            ws = writer.sheets["Summary"]
            start_row = len(summary_df) + 4
            
            ws.cell(row=start_row, column=1, value="System Balance Verification Status:")
            if status == "verified":
                val_str = "VERIFIED (Running balance chain matches perfectly)"
            elif status == "mismatches":
                val_str = f"MISMATCH DETECTED ({mismatch_count} rows mismatched. Check column alignments)"
            else:
                val_str = "UNVERIFIED (No running balances found in statements)"
            ws.cell(row=start_row, column=2, value=val_str)
            
            if unreconciled_logs:
                start_row += 2
                ws.cell(row=start_row, column=1, value="Unreconciled Transactions:")
                for log in unreconciled_logs:
                    start_row += 1
                    ws.cell(row=start_row, column=1, value=log)
                    
    return output_path
