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

###############################################################################
#  SIMPLE, PURE‑PYTHON  TABLE EXTRACTOR  (no Ghostscript / Java required)     #
#  Designed for Interparfums invoices uploaded to Vercel / GitHub workflows  #
###############################################################################
#  ‣ NO external system packages                                             #
#  ‣ Uses x‑coordinate “slices” once per column                              #
#  ‣ Filters rows with a strict numeric pattern → cabeceras quedan fuera     #
###############################################################################

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

# ──────────────────────────────  COLUMN BOUNDS  ──────────────────────────────
# Ajusta estos valores si tu PDF tiene márgenes distintos. Medidos en puntos
# PDF (1 pt ≈ 0.353 mm). Basta afinarlos una sola vez.
COL_BOUNDS: Dict[str, tuple] = {
    "ref":   (  0,  70),   # Reference (5‑6 dígitos, opcional sufijo A/B)
    "desc":  ( 70, 340),   # Descripción completa
    "upc":   (340, 430),   # UPC / Code EAN (12‑14 dígitos)
    "ctry":  (430, 465),   # Country of origin (2‑3 letras)
    "hs":    (465, 535),   # HS Code
    "qty":   (535, 585),   # Quantity
    "unit":  (585, 635),   # Unit price
    "total": (635, 725),   # Line amount
}

OUTPUT_COLS = [
    "Reference", "Code EAN", "Custom Code", "Description", "Origin",
    "Quantity", "Unit Price", "Total Price", "Invoice Number",
]

# ───────────────────────────── REGEX DE VALIDACIÓN ───────────────────────────
REF_PAT  = re.compile(r"^\d{5,6}[A-Z]?$")        # 12345, 26714A...
UPC_PAT  = re.compile(r"^\d{12,14}$")            # 085715169587
NUM_PAT  = re.compile(r"[0-9]" )                 # al menos un dígito

# Palabras que marcan cabeceras / totales que queremos descartar
SKIP_SNIPPETS = {"No. Description", "Total before", "Bill To Ship", "CIF CHILE"}

# ─────────────────────────────  UTILIDADES  ─────────────────────────────────-

def clean(txt: str) -> str:
    return txt.replace("\u202f", " ").strip()


def to_float(txt: str) -> float:
    txt = txt.replace("\u202f", "").replace(" ", "")
    if txt.count(",") == 1 and txt.count(".") == 0:
        txt = txt.replace(",", ".")          # 15,13 → 15.13
    elif txt.count(".") > 1:                   # 1.234.567 → 1234567
        txt = txt.replace(".", "")
    return float(txt or 0)


def to_int(txt: str) -> int:
    return int(txt.replace(",", "").replace(".", "") or 0)

# ─────────────────────────────  EXTRACCIÓN  ─────────────────────────────────-

def rows_from_page(page) -> List[Dict[str, str]]:
    """Extrae filas válidas de una página usando coordenadas."""
    rows: List[Dict[str, str]] = []

    # Agrupar caracteres por línea Y (baseline) con tolerancia
    grouped: Dict[float, List[dict]] = {}
    for ch in page.chars:
        y_key = round(ch["top"], 1)  # 0.1 pt de precisión
        grouped.setdefault(y_key, []).append(ch)

    for y, chs in sorted(grouped.items()):
        line_text = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if not line_text.strip():
            continue
        if any(snippet in line_text for snippet in SKIP_SNIPPETS):
            continue

        cols = {k: "" for k in COL_BOUNDS}
        for ch in sorted(chs, key=lambda c: c["x0"]):
            x_mid = (ch["x0"] + ch["x1"]) / 2
            for key, (x0, x1) in COL_BOUNDS.items():
                if x0 <= x_mid < x1:
                    cols[key] += ch["text"]
                    break
        # Limpieza básica
        cols = {k: clean(v) for k, v in cols.items()}

        # Si la línea es continuación de descripción (ref vacío) ⇒ concatenar
        if not cols["ref"]:
            if rows:
                rows[-1]["desc"] += " " + cols["desc"]
            continue

        # Validación fuerte de fila
        if not REF_PAT.match(cols["ref"]):
            continue
        if not UPC_PAT.match(cols["upc"]):
            continue
        if not NUM_PAT.search(cols["qty"]):
            continue

        rows.append(cols)
    return rows

# ─────────────────────────────  FLASK ROUTES  ───────────────────────────────

@app.route("/", methods=["POST"])
@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded", 400

        extracted = []
        for pdf_file in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf_file.save(tmp.name)
                invoice_no = re.search(r"SIP(\d+)", pdf_file.filename or "")
                invoice_no = invoice_no.group(1) if invoice_no else ""

                with pdfplumber.open(tmp.name) as pdf:
                    for page in pdf.pages:
                        for r in rows_from_page(page):
                            extracted.append(
                                {
                                    "Reference":      r["ref"],
                                    "Code EAN":       r["upc"],
                                    "Custom Code":    r["hs"],
                                    "Description":    r["desc"],
                                    "Origin":         r["ctry"],
                                    "Quantity":       to_int(r["qty"]),
                                    "Unit Price":     to_float(r["unit"]),
                                    "Total Price":    to_float(r["total"]),
                                    "Invoice Number": invoice_no,
                                }
                            )
                os.unlink(tmp.name)

        if not extracted:
            return "Sin registros extraídos", 400

        # ─── Generar Excel ─────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.append(OUTPUT_COLS)
        for row in extracted:
            ws.append([row[c] for c in OUTPUT_COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")




