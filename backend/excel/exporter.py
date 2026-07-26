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
        max_len = max(df[col].astype(str).map(len).max() if len(df) else 0, len(str(col)))
        ws.column_dimensions[get_column_letter(i)].width = min(max_len + 3, 60)


def _month_summary_rows(df, month_labels):
    summary_rows = []
    for label in month_labels:
        month_df = df[df["MonthLabel"] == label]

        def sum_type(t, col):
            return month_df.loc[month_df["Type"] == t, col].sum()

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
    for key in ["Credit", "Debit", "Credit Charge", "Debit Charge", "Credit Interest", "Debit Interest", "Transaction Count"]:
        totals[key] = sum(r[key] for r in rows)
    rows.append(totals)

    return rows


def export_to_excel(transactions, output_path):
    df = to_dataframe(transactions)
    if df.empty:
        raise ValueError("No transactions to export.")

    month_labels = (
        df[["Year", "Month", "MonthLabel"]]
        .drop_duplicates()
        .sort_values(["Year", "Month"])["MonthLabel"]
        .tolist()
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for label in month_labels:
            month_df = df[df["MonthLabel"] == label].copy()
            sheet_df = month_df[["Date", "Narration", "Type", "Debit", "Credit", "Balance"]].copy()
            sheet_df["Date"] = sheet_df["Date"].dt.strftime("%d-%m-%Y")

            sheet_name = label[:31]
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            _autofit(writer.sheets[sheet_name], sheet_df)

        summary_rows = _month_summary_rows(df, month_labels)
        summary_df = pd.DataFrame(summary_rows)
        totals = summary_df.drop(columns="Month").sum(numeric_only=True)
        totals["Month"] = "TOTAL (FY)"
        summary_df = pd.concat([summary_df, pd.DataFrame([totals])], ignore_index=True)

        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        _autofit(writer.sheets["Summary"], summary_df)
        writer.book.move_sheet("Summary", offset=-(len(writer.book.sheetnames) - 1))

    return output_path
