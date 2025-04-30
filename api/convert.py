"""
Función serverless (Vercel) que detecta automáticamente si el PDF es
FACTURA o PROFORMA y genera un Excel con las columnas completas:

• Factura  → Reference, Code EAN, Custom Code, Description, Origin,
             Quantity, Unit Price, Total Price, Invoice Number
• Proforma → Reference, Code EAN, Description, Origin,
             Quantity, Unit Price, Total Price
"""

import logging, re, tempfile, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)

app = Flask(__name__)

# ───────────────────────── helpers ─────────────────────────
def parse_num(s: str) -> float:
    """'1.234,56' → 1234.56"""
    return float(s.replace('.', '').replace(',', '.').strip()) if s else 0.0


def detect_doc_type(first_page_text: str) -> str:
    up = first_page_text.upper()
    if ('ACCUSE' in up and 'RECEPTION' in up) or 'ACKNOWLEDGE' in up or 'PROFORMA' in up:
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('No pude determinar si es factura o proforma.')


# ───────────────────────── endpoint ─────────────────────────
@app.route("/", methods=["POST"])            # útil localmente
@app.route("/api/convert", methods=["POST"])  # ruta en Vercel
def convert():
    try:
        if 'file' not in request.files:
            return "No file uploaded", 400

        uploaded = request.files['file']

        # Guardar PDF temporal
        tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        uploaded.save(tmp_pdf.name)

        # Detectar tipo
        first_txt = extract_text(tmp_pdf.name, page_numbers=[0]) or ""
        doc_type  = detect_doc_type(first_txt)

        records = []
        with pdfplumber.open(tmp_pdf.name) as pdf:
            if doc_type == 'factura':
                # ——— Invoice base ———
                m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,50}?(\d{6,})', first_txt, re.I)
                invoice_base = m.group(1) if m else None
                if not invoice_base:
                    cands = re.findall(r'(\d{8,})', first_txt)
                    invoice_base = cands[0] if cands else None
                if not invoice_base:
                    return "No encontré número de factura", 400

                origin = ''
                detail_pat = re.compile(
                    r'^([A-Z]\d+)\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$'
                )

                for page in pdf.pages:
                    txt = page.extract_text() or ''
                    lines = txt.split('\n')
                    up   = txt.upper()
                    inv_num = invoice_base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else invoice_base

                    # país de origen, si existe
                    for L in lines:
                        if "PAYS D'ORIGINE" in L:
                            origin = L.split(':',1)[-1].strip()

                    i = 0
                    while i < len(lines):
                        mm = detail_pat.match(lines[i].strip())
                        if mm:
                            ref, ean, custom, qty_s, unit_s, tot_s = mm.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            records.append({
                                'Reference':      ref,
                               

