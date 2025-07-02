import logging
import tempfile
import os
import traceback
from io import BytesIO
from typing import List, Dict

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ────────────────── CONFIGURACIÓN DE COLUMNAS (en puntos PDF) ──────────────────
# Los x‑rangos están medidos sobre un invoice Interparfums tamaño carta (612×792 pt).
# Si tu PDF viene con márgenes diferentes, ajústalos aquí una sola vez.

COL_BOUNDS = {
    "ref":   (  0,  65),   # Reference 5‑6 dígitos (puede llevar sufijo A/B)
    "desc":  ( 70, 330),   # Descripción completa (puede ocupar varias líneas)
    "upc":   (330, 420),   # UPC / Code EAN 12‑14 dígitos
    "ctry":  (420, 455),   # Country of origin (2‑3 letras)
    "hs":    (455, 525),   # HS Code
    "qty":   (525, 570),   # Quantity
    "unit":  (570, 615),   # Unit Price
    "total": (615, 705),   # Line Amount
}

COL_ORDER = ["ref", "upc", "hs", "desc", "ctry", "qty", "unit", "total"]

OUTPUT_COLS = [
    "Reference", "Code EAN", "Custom Code", "Description",
    "Origin", "Quantity", "Unit Price", "Total Price", "Invoice Number",
]

HEADER_KEYS = {"No.", "Description", "UPC"}
SUMMARY_SNIPPET = "Total before discount"

# ──────────────────────────── AUXILIARES ───────────────────────────────────────

def clean(s: str) -> str:
    return s.replace("\u202f", " ").strip()


def parse_number(num: str) -> float:
    """Convierte texto con separadores (coma/punto) en float."""
    num = num.replace(" ", "").replace("\u202f", "")
    if num.count(",") == 1 and num.count(".") == 0:
        num = num.replace(",", ".")           # formato 1,234 → 1.234
    elif num.count(".") > 1:                    # 1.234.567 → 1234567
        num = num.replace(".", "")
    return float(num)


def parse_int(num: str) -> int:
    return int(num.replace(",", "").replace(".", ""))


# ───────────────────────── EXTRACCIÓN POR COORDENADAS ─────────────────────────

def extract_rows(page) -> List[Dict[str, str]]:
    """Devuelve lista de dicts con columnas crudas para una página."""
    rows = []
    # extrae caracteres y los agrupa por línea (y0 ~ baseline)
    chars = page.chars
    # agrupamos por y0 redondeado a 1 décima para robustez
    lines: Dict[float, List[dict]] = {}
    for ch in chars:
        y_key = round(ch["top"], 1)
        lines.setdefault(y_key, []).append(ch)

    for y, chs in sorted(lines.items()):
        text_line = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if not text_line.strip():
            continue
        # saltar cabeceras y resumen
        if any(h in text_line for h in HEADER_KEYS) or SUMMARY_SNIPPET in text_line:
            continue
        # mapear a columnas
        cols = {k: "" for k in COL_BOUNDS}
        for cell in sorted(chs, key=lambda c: c["x0"]):
            x_center = (cell["x0"] + cell["x1"]) / 2
            for key, (x0, x1) in COL_BOUNDS.items():
                if x0 <= x_center < x1:
                    cols[key] += cell["text"]
                    break
        # limpiar
        for k in cols:
            cols[k] = clean(cols[k])
        # desc puede extenderse a varias líneas: si esta línea no tiene referencia
        if not cols["ref"]:
            # se concatena a la desc de la fila previa
            if rows:
                rows[-1]["desc"] += " " + cols["desc"]
            continue
        rows.append(cols)
    return rows

# ──────────────────────────── FLASK ROUTE ─────────────────────────────────────

@app.route("/", methods=["POST"])
@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded", 400

        all_rows = []
        for pdf_file in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf_file.save(tmp.name)
                with pdfplumber.open(tmp.name) as pdf:
                    # naive invoice number (look first page for SIP...)
                    first_text = pdf.pages[0].extract_text() or ""
                    inv_match = re.search(r"SIP(\d{8})", first_text)
                    invoice_no = inv_match.group(1) if inv_match else ""

                    for page in pdf.pages:
                        page_rows = extract_rows(page)
                        for r in page_rows:
                            all_rows.append({
                                "Reference":   r["ref"],
                                "Code EAN":    r["upc"],
                                "Custom Code": r["hs"],
                                "Description": r["desc"],
                                "Origin":      r["ctry"],
                                "Quantity":    parse_int(r["qty"] or "0"),
                                "Unit Price":  parse_number(r["unit"] or "0"),
                                "Total Price": parse_number(r["total"] or "0"),
                                "Invoice Number": invoice_no,
                            })
                os.unlink(tmp.name)

        if not all_rows:
            return "Sin registros extraídos", 400

        # ─── Generar Excel ─────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.append(OUTPUT_COLS)
        for r in all_rows:
            ws.append([r[c] for c in OUTPUT_COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf,
                         as_attachment=True,
                         download_name="extracted_data.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")



