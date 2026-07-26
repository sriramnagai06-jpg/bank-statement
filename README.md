# Bank Statement Analyzer

Upload a bank statement PDF, get back an Excel workbook split month-by-month
into **Credit / Debit / Charges / Interest**, for the whole financial year.

Supported banks: **Canara** (fully verified against a real 245-page / 1,783-
transaction statement — zero balance-chain mismatches), plus **SBI, HDFC,
ICICI, Axis** (generic table-based parsers, not yet tested against real
statements from those banks — see "Known limitations" below).

## Project structure

```
bank-statement-analyzer/
├── backend/
│   ├── app.py                 # Flask API (serves frontend + /api routes)
│   ├── requirements.txt
│   ├── parser/
│   │   ├── canara.py          # verified, coordinate-based parser
│   │   ├── sbi.py              # generic table parser, needs a real sample
│   │   ├── hdfc.py             # generic table parser, needs a real sample
│   │   ├── icici.py            # generic table parser, needs a real sample
│   │   ├── axis.py             # generic table parser, needs a real sample
│   │   ├── generic_table.py    # shared table-extraction logic
│   │   ├── detector.py         # auto-detects bank from PDF text
│   │   └── parser_manager.py   # dispatches to the right parser
│   ├── excel/
│   │   └── exporter.py         # classification + month-wise .xlsx writer
│   └── uploads/                # temp storage for uploaded/generated files
│
└── frontend/
    ├── index.html
    ├── style.css
    └── script.js
```

## Running it on ANY Laptop & Phone

### Option A: One-Click Launchers
- **Windows Laptop:** Double-click [`run.bat`](file:///c:/Users/acer/Downloads/bank-statement-analyzer/bank-statement-analyzer/run.bat) (or run `python run.py`)
- **Mac / Linux Laptop:** Run `./run.sh` (or `python run.py`)

### Option B: Manual Command Line
1. Install dependencies:
   ```bash
   pip install -r backend/requirements.txt
   ```
2. Start the server:
   ```bash
   python run.py
   ```

### Accessing the App:
- **On the host laptop browser:** `http://localhost:5000`
- **On SAME Wi-Fi network:** `http://192.168.29.50:5000`

---

## Free Cloud Deployment (Access from ANY Phone / Laptop on DIFFERENT Wi-Fi / Mobile Data)

To make this app available 24/7 on **all phones, laptops, and mobile networks worldwide** without needing your laptop to stay on:

### Option 1: Deploy on Render.com (Recommended - 100% Free)
1. Push this folder/repository to GitHub.
2. Sign up at [render.com](https://render.com) (Free).
3. Click **New +** -> **Web Service**.
4. Connect your GitHub repository.
5. Render will automatically detect [`render.yaml`](file:///c:/Users/acer/Downloads/bank-statement-analyzer/bank-statement-analyzer/render.yaml) and deploy your app.
6. You will get a permanent public link like: `https://bank-statement-analyzer.onrender.com`

### Option 2: Deploy using Docker
Build and run the container on any cloud platform (AWS, GCP, DigitalOcean, Railway):
```bash
docker build -t bank-statement-analyzer .
docker run -p 5000:5000 bank-statement-analyzer
```


## How classification works

- Every row is tagged **Debit** or **Credit** based on which statement
  column the amount came from.
- If the narration contains **"CHG"** → tagged as a **Charge**.
- If the narration contains **"INT"** → tagged as **Interest**.
- Combined, each row becomes one of: Credit, Debit, Credit Charge,
  Debit Charge, Credit Interest, Debit Interest.

This is a simple substring rule (as requested) — it's fast and works well
in practice, but can occasionally misfire if unrelated text happens to
contain "CHG" or "INT" (e.g. a payer's name or note). Spot-check the
Summary sheet if exact charge/interest totals matter for filing purposes.

## Known limitations

- **SBI / HDFC / ICICI / Axis parsers are unverified.** They use
  `pdfplumber`'s table detection with each bank's commonly documented
  column headers, but real exports vary. If a statement returns "No
  transactions could be extracted" or wrong figures, send a real sample
  PDF from that bank so the parser can be corrected.
- The Canara parser is calibrated to one specific export layout (word
  x-coordinates for column boundaries). A different Canara export style
  (e.g. from a different branch/app version) may need the boundary
  constants in `parser/canara.py` re-checked.
- Uploaded PDFs are deleted from `backend/uploads/` immediately after
  parsing; generated Excel files remain there until manually cleared.
