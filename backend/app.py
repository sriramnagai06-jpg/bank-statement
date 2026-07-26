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
import uuid

from flask import Flask, jsonify, request, send_from_directory

from excel.exporter import export_to_excel, summarize
from parser.parser_manager import UnsupportedBankError, parse_statement, PARSERS
import pdfminer.pdfdocument

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "frontend")

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path="")

MAX_UPLOAD_MB = 50
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.route("/")
def index():
    return app.send_static_file("index.html")


@app.route("/api/analyze", methods=["POST"])
def analyze():
    file = request.files.get("statement")
    if not file or file.filename == "":
        return jsonify({"error": "No file was uploaded."}), 400

    forced_bank = request.form.get("bank")
    if forced_bank == "auto":
        forced_bank = None
    if forced_bank and forced_bank not in PARSERS:
        return jsonify({"error": f"Unsupported bank '{forced_bank}'."}), 400

    password = request.form.get("password")

    job_id = uuid.uuid4().hex[:10]
    pdf_path = os.path.join(UPLOAD_DIR, f"{job_id}.pdf")
    file.save(pdf_path)

    try:
        bank_key, bank_display_name, transactions = parse_statement(pdf_path, forced_bank=forced_bank, password=password)
    except UnsupportedBankError as e:
        return jsonify({"error": str(e)}), 422
    except pdfminer.pdfdocument.PDFPasswordIncorrect:
        return jsonify({"error": "This PDF is password protected. Please enter the correct password."}), 401
    except Exception as e:
        return jsonify({"error": f"Failed to parse the statement: {e}"}), 500
    finally:
        # keep the uploaded PDF only as long as needed; remove after parsing
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    if not transactions:
        return jsonify({
            "error": f"No transactions could be extracted (detected bank: {bank_display_name}). "
                     f"This statement's layout may not match the parser yet."
        }), 422

    xlsx_name = f"{job_id}_monthwise.xlsx"
    xlsx_path = os.path.join(UPLOAD_DIR, xlsx_name)
    export_to_excel(transactions, xlsx_path)

    return jsonify({
        "bank_key": bank_key,
        "bank": bank_display_name,
        "transaction_count": len(transactions),
        "download_url": f"/api/download/{xlsx_name}",
        "summary": summarize(transactions),
    })


@app.route("/api/download/<path:filename>")
def download(filename):
    if not filename.endswith(".xlsx") or "/" in filename or "\\" in filename:
        return jsonify({"error": "Invalid filename."}), 400
    return send_from_directory(UPLOAD_DIR, filename, as_attachment=True,
                                download_name="statement_monthwise.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", debug=True, port=5000)
