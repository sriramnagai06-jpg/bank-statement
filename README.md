# Bank Statement Analyzer & Reconciliation Engine

A complete, production-grade Bank Statement Analyzer that converts bank statement PDFs, Excel files, and copy-pasted statement rows into clean, verified, and audited financial ledgers.

---

## Key Features

- **ONE Source of Truth:** The exact same verified transaction dataset powers the executive summary, dashboard tables, reconciliation checks, inter-company detection, and Excel exports.
- **Executive Summary & Month-Wise Breakdown:** Computes FY totals, Opening/Closing balance, Credit/Debit breakdowns, Charges (tax/GST/fees), and Interest across all months.
- **Complete All Transactions Detail Table:** Displayed directly on the web app below the summary cards, complete with `S.No`, `Date`, `Particulars`, `Type`, `Debit`, `Credit`, `Balance`, client-side real-time filtering, and a bottom `TOTAL` row.
- **Dedicated Inter-Company Transfer Detection:** Transferred funds between linked entities/partners are classified and presented in a dedicated section with totals, without removing them from the master transaction list.
- **Comprehensive Multi-Sheet Excel Export (`Bank_Statement_Analysis.xlsx`):**
  - `Sheet 1 — Summary`: Bank name, statement period, total transaction count, total debit/credit, opening/closing balance, reconciliation status, and month-wise breakdown table.
  - `Sheet 2 — All Transactions`: Every single transaction with `S.No`, `Date`, `Particulars`, `Type`, `Debit`, `Credit`, `Balance`, `Reference`, and bottom `TOTAL` row.
  - `Sheet 3 — Inter-Company Transactions`: All inter-company transactions with `S.No`, `Date`, `Particulars`, `Type`, `Debit`, `Credit`, `Balance`, `Classification`, and bottom `TOTAL` row.
  - `Sheets 4+ — Monthly Sheets (Apr-2025, May-2025...)`: Month-by-month transaction ledgers with `S.No` and bottom `TOTAL` rows.
- **Professional Formatting:** Bold headers, freeze panes, auto-filters, borders, and currency number formatting (`#,##0.00`) on all sheets.
- **Multi-Format Input:**
  - PDF Statement uploads (Canara, SBI, HDFC, ICICI, Axis, etc.)
  - Excel Statement uploads (Karur Vysya Bank / KVB formats A, B, C)
  - Copy-Pasted text mode with automatic delimiter and date detection.
- **Dual Statement Inter-Company Reconciliation:** Automatically reconciles two companion statements (Statement A ↔ Statement B) and highlights matched pairs.

---

## Supported Banks & Parsers

| Bank / Source | Input Format | Parser Strategy |
|---|---|---|
| **Karur Vysya Bank (KVB)** | `.xlsx`, `.xls` | Coordinate & text prefix parser handling all KVB format variants, repeated headers, and multiline narration |
| **Canara Bank** | `.pdf` | High-precision visual layout word coordinate parser |
| **SBI, HDFC, ICICI, Axis** | `.pdf` | Grid table parser with header alias detection |
| **Copy-Paste Text** | Plain text / TSV | Heuristic tokenizer supporting single-amount, two-column, and Dr/Cr labeled text |

---

## Installation & Running Locally

### Prerequisites
- Python 3.9+
- `pip`

### Step 1: Install Dependencies
```bash
pip install -r backend/requirements.txt
```

### Step 2: Start the Server
```bash
python run.py
```
*(Or directly from `backend/`: `python app.py`)*

### Step 3: Open in Browser
- **Local machine:** [http://127.0.0.1:5000](http://127.0.0.1:5000)
- **Same Wi-Fi network:** `http://<your-ip-address>:5000`

---

## Project Structure

```
bank-statement-analyzer/
├── backend/
│   ├── app.py                 # Flask REST API & web server
│   ├── requirements.txt       # Python package dependencies
│   ├── excel/
│   │   └── exporter.py        # Multi-sheet openpyxl Excel workbook generator
│   ├── parser/
│   │   ├── kvb.py             # Karur Vysya Bank Excel parser
│   │   ├── canara.py          # Canara Bank coordinate-based parser
│   │   ├── copypaste.py       # Copy-paste text heuristic parser
│   │   ├── generic_table.py   # Table-based PDF extraction
│   │   ├── detector.py        # Auto-detects bank from statement text
│   │   └── parser_manager.py  # Central dispatcher
│   └── uploads/               # Temporary storage for uploads & generated workbooks
├── frontend/
│   ├── index.html             # Single-page interface
│   ├── script.js              # UI interaction, dynamic rendering, and filters
│   └── style.css              # Financial ledger design system
├── tests/                     # Test suite
├── run.py                     # Universal runner with local IP detection
└── README.md
```

---

## API Reference

### 1. Analyze Statement
`POST /api/analyze`

**Request (`multipart/form-data`):**
- `file`: PDF or Excel statement file (up to 2 files for dual-statement reconciliation)
- `bank`: Bank key (`auto`, `kvb`, `canara`, `sbi`, `hdfc`, `icici`, `axis`)
- `password`: (Optional) PDF decryption password
- `paste_text`: (Optional) Raw statement text if using copy-paste mode

**Response (`application/json`):**
```json
{
  "bank": "Karur Vysya Bank",
  "transaction_count": 1193,
  "opening_balance": 1133.34,
  "closing_balance": 5083.34,
  "reconciliation_status": "pass",
  "download_url": "/api/download/<job_id>_monthwise.xlsx",
  "summary": [
    { "Month": "Apr-2025", "Credit": 10000.0, "Debit": 4300.0, "Transaction Count": 5 }
  ],
  "transactions": [
    { "s_no": 1, "date": "04-04-2025", "narration": "UPI/CR/...", "type": "Credit", "debit": 0.0, "credit": 10000.0, "balance": 11133.34 }
  ],
  "inter_company_transactions": [ ... ],
  "inter_company_summary": { "count": 12, "total_debit": 250000.0, "total_credit": 250000.0 }
}
```

### 2. Download Complete Workbook
`GET /api/download/<filename>`

Returns the complete `Bank_Statement_Analysis.xlsx` spreadsheet.

---

## Running Tests

Run the test suite:
```bash
python -m unittest discover tests
```

---

## Security & Privacy
- Uploaded statement files are strictly processed in memory or ephemeral storage and removed immediately after processing.
- No sensitive banking credentials or API keys are required.
- CORS protection and filename sanitization enabled.
