import logging
import os
import re
import tempfile
import traceback
from io import BytesIO
from typing import Dict, List

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

# ─────────────────────────  AJUSTA SOLO ESTO SI CAMBIA TU PDF  ───────────────
COL_BOUNDS: Dict[str, tuple] = {
    "ref":   (  0,  70),
    "desc":  ( 70, 340),
    "upc":   (340, 430),
    "ctry":  (430, 465),
    "hs":    (465, 535),
    "qty":   (535, 585),
    "unit":  (585, 635),
    "total": (635, 725),
}
# ─────────────────────────────────────────────────────────────────────────────

OUTPUT_COLS = [
    "Reference", "Code EAN", "Custom Code", "Description", "Origin",
    "Quantity", "Unit Price", "Total Price", "Invoice Number",
]

REF_PAT  = re.compile(r"^\d{5,6}[A-Z]?$")
UPC_PAT  = re.compile(r"^\d{12,14}$")
NUM_PAT  = re.compile(r"[0-9]")

SKIP_SNIPPETS = {
    "No. Description", "Total before", "Bill To Ship", "CIF CHILE",
    "Invoice", "Ship From", "Ship To", "VAT/Tax", "Shipping Te"
}

def clean(t: str) -> str:
    return t.replace("\u202f", " ").strip()

def to_float(t: str) -> float:
    t = t.replace("\u202f", "").replace(" ", "")
    if t.count(",") == 1 and t.count(".") == 0:
        t = t.replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    return float(t or 0)

def to_int(t: str) -> int:
    return int(t.replace(",", "").replace(".", "") or 0)

def rows_from_page(page) -> List[Dict[str, str]]:
    rows = []
    grouped: Dict[float, List[dict]] = {}
    for ch in page.chars:
        y = round(ch["top"], 1)
        grouped.setdefault(y, []).append(ch)

    for _, chs in sorted(grouped.items()):
        text = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if not text.strip() or any(s in text for s in SKIP_SNIPPETS):
            continue

        cols = {k: "" for k in COL_BOUNDS}
        for ch in sorted(chs, key=lambda c: c["x0"]):
            xm = (ch["x0"] + ch["x1"]) / 2
            for key, (x0, x1) in COL_BOUNDS.items():
                if x0 <= xm < x1:
                    cols[key] += ch["text"]; break
        cols = {k: clean(v) for k, v in cols.items()}

        if not cols["ref"]:
            if rows: rows[-1]["desc"] += " " + cols["desc"]
            continue
        if not REF_PAT.match(cols["ref"]) or not UPC_PAT.match(cols["upc"]):
            continue
        if not NUM_PAT.search(cols["qty"]):
            continue
        rows.append(cols)
    return rows

@app.route("/", methods=["POST"])
@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded", 400

        extracted = []
        for f in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                f.save(tmp.name)
                inv = re.search(r"SIP(\\d+)", f.filename or "")
                inv = inv.group(1) if inv else ""
                with pdfplumber.open(tmp.name) as pdf:
                    for p in pdf.pages:
                        for r in rows_from_page(p):
                            extracted.append({
                                "Reference": r["ref"], "Code EAN": r["upc"],
                                "Custom Code": r["hs"], "Description": r["desc"],
                                "Origin": r["ctry"], "Quantity": to_int(r["qty"]),
                                "Unit Price": to_float(r["unit"]),
                                "Total Price": to_float(r["total"]),
                                "Invoice Number": inv,
                            })
            os.unlink(tmp.name)

        if not extracted:
            return "Sin registros extraídos", 400

        wb = Workbook(); ws = wb.active; ws.append(OUTPUT_COLS)
        for r in extracted: ws.append([r[c] for c in OUTPUT_COLS])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name="extracted_data.xlsx",
                         mimetype=("application/vnd.openxmlformats-"
                                   "officedocument.spreadsheetml.sheet"))
    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")



