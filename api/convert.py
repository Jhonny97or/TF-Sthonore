"""
Función serverless para Vercel que convierte un PDF (factura o proforma)
en un Excel descargable.  Rutas aceptadas:
  • POST /api/convert  (producción en Vercel)
  • POST /            (útil para pruebas locales)
"""

import logging
import tempfile
import traceback
import re
from io import BytesIO

from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

# Suprimir warnings ruidosos de pdfminer
logging.getLogger("pdfminer").setLevel(logging.ERROR)

app = Flask(__name__)

# ──────────────────────────── Helpers ─────────────────────────────
def _num(s: str) -> float:
    """Convierte '1.234,56'  →  1234.56"""
    s = (s or "").strip()
    return float(s.replace('.', '').replace(',', '.')) if s else 0.0


def _factura_regex(text: str) -> str | None:
    """Extrae el número de factura de la primera página (6+ dígitos)."""
    m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', text.upper())
    return m.group(1) if m else None


# ─────────────────────────── Endpoints ────────────────────────────
@app.route("/", methods=["GET", "POST"])
@app.route("/api/convert", methods=["GET", "POST"])
def convert():
    # GET simple para evitar 404 al navegar manualmente
    if request.method == "GET":
        return "Convertidor PDF → Excel activo. Envía un POST con tu PDF.", 200

    # 1) Validar archivo
    if "file" not in request.files:
        return "No file uploaded", 400

    uploaded = request.files["file"]
    doc_type = request.form.get("type", "auto").lower()

    try:
        # 2) Guardar PDF en /tmp (sistema de archivos temporal de Vercel)
        tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        uploaded.save(tmp_pdf.name)
        pdf_path = tmp_pdf.name

        # 3) Leer texto de la 1ª página para detectar tipo
        first_page_text = extract_text(pdf_path, page_numbers=[0]) or ""
        upper_first = first_page_text.upper()

        if doc_type == "auto":
            if ("ACCUSE" in upper_first and "RECEPTION" in upper_first) or "ACKNOWLEDGE" in upper_first:
                doc_type = "proforma"
            elif "FACTURE" in upper_first or "INVOICE" in upper_first:
                doc_type = "factura"
            else:
                return "No pude determinar si es factura o proforma", 400

        # 4) Extraer líneas
        records: list[dict] = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""

                if doc_type == "factura":
                    # ref  EAN  custom  qty  unit  total
                    pat = re.compile(
                        r'^([A-Z]\d{5,7})\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$'
                    )
                    for ln in txt.split("\n"):
                        m = pat.match(ln.strip())
                        if m:
                            ref, ean, custom, qty_s, unit_s, tot_s = m.groups()
                            records.append({
                                "Reference": ref,
                                "Code EAN": ean,
                                "Custom Code": custom,
                                "Quantity": int(qty_s),
                                "Unit Price": _num(unit_s),
                                "Total Price": _num(tot_s),
                            })
                else:  # proforma
                    # ref  EAN  price  qty
                    pat = re.compile(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')
                    lines = txt.split("\n")
                    i = 0
                    while i < len(lines):
                        m = pat.search(lines[i])
                        if m:
                            ref, ean, price_s, qty_s = m.groups()
                            qty = int(qty_s.replace('.', '').replace(',', ''))
                            unit = _num(price_s)
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            records.append({
                                "Reference": ref,
                                "Code EAN": ean,
                                "Description": desc,
                                "Quantity": qty,
                                "Unit Price": unit,
                                "Total Price": unit * qty,
                            })
                            i += 2
                        else:
                            i += 1

        if not records:
            return "Sin registros extraídos – verifica el PDF", 400

        # 5) Generar Excel en memoria
        headers = list(records[0].keys())
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for row in records:
            ws.append([row[h] for h in headers])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        # 6) Responder con el archivo
        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception:
        tb = traceback.format_exc()
        return f"❌ Error interno:\n{tb}", 500
