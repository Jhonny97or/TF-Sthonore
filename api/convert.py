"""
Función serverless (Vercel) que detecta automáticamente si el PDF es
FACTURA o PROFORMA y genera un Excel con todas las columnas:
Reference, Code EAN, Custom Code, Description, Origin, Quantity,
Unit Price, Total Price, Invoice Number (solo facturas).
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
    """Convierte '1.234,56' ➜ 1234.56"""
    return float(s.replace('.', '').replace(',', '.').strip()) if s else 0.0


def detect_doc_type(first_page_text: str) -> str:
    txt = first_page_text.upper()
    if ('ACCUSE' in txt and 'RECEPTION' in txt) or 'ACKNOWLEDGE' in txt or 'PROFORMA' in txt:
        return 'proforma'
    if 'FACTURE' in txt or 'INVOICE' in txt:
        return 'factura'
    raise ValueError('No pude determinar si es factura o proforma.')


# ───────────────────────── endpoint ─────────────────────────
@app.route("/", methods=["POST"])           # útil si haces curl local
@app.route("/api/convert", methods=["POST"]) # ruta oficial en Vercel
def convert():
    try:
        if 'file' not in request.files:
            return "No file uploaded", 400

        uploaded = request.files['file']

        # Guarda PDF en /tmp (lectura más rápida)
        pdf_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        uploaded.save(pdf_tmp.name)

        # Detectar tipo
        first_txt = (extract_text(pdf_tmp.name, page_numbers=[0]) or "")
        doc_type  = detect_doc_type(first_txt)
        print('🛈 Documento detectado como:', doc_type)

        records = []
        with pdfplumber.open(pdf_tmp.name) as pdf:
            if doc_type == 'factura':
                # 1) Invoice base
                m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,50}?(\d{6,})', first_txt, re.I)
                invoice_base = m.group(1) if m else None
                if not invoice_base:
                    cands = re.findall(r'(\d{8,})', first_txt)
                    invoice_base = cands[0] if cands else None
                if not invoice_base:
                    raise ValueError('No encontré número de factura.')

                # 2) Recorrido de páginas
                origin = ''
                detail_pat = re.compile(
                    r'^([A-Z]\d+)\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$'
                )
                for page in pdf.pages:
                    txt  = page.extract_text() or ''
                    lines = txt.split('\n')
                    up   = txt.upper()
                    inv_num = invoice_base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else invoice_base

                    # PAYS D'ORIGINE
                    for L in lines:
                        if "PAYS D'ORIGINE" in L:
                            parts = L.split(':',1)
                            if len(parts)==2: origin = parts[1].strip()

                    i = 0
                    while i < len(lines):
                        m2 = detail_pat.match(lines[i].strip())
                        if m2:
                            ref, ean, custom, qty_s, unit_s, tot_s = m2.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            records.append({
                                'Reference':      ref,
                                'Code EAN':       ean,
                                'Custom Code':    custom,
                                'Description':    desc,
                                'Origin':         origin,
                                'Quantity':       int(qty_s),
                                'Unit Price':     parse_num(unit_s),
                                'Total Price':    parse_num(tot_s),
                                'Invoice Number': inv_num
                            })
                            i += 1  # saltar descripción
                        i += 1

            else:  # PROFORMA
                origin = ''
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    # PAYS D'ORIGINE
                    for L in lines:
                        if "PAYS D'ORIGINE" in L:
                            parts = L.split(':',1)
                            if len(parts)==2: origin = parts[1].strip()

                    for idx, line in enumerate(lines):
                        m = re.search(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)', line)
                        if m:
                            ref, ean, unit_s, qty_s = m.groups()
                            desc = lines[idx+1].strip() if idx+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',','').strip())
                            unit = parse_num(unit_s)
                            records.append({
                                'Reference':   ref,
                                'Code EAN':    ean,
                                'Description': desc,
                                'Origin':      origin,
                                'Quantity':    qty,
                                'Unit Price':  unit,
                                'Total Price': unit*qty
                            })

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        # Ordenar columnas según tipo
        cols_fact = ['Reference','Code EAN','Custom Code','Description','Origin',
                     'Quantity','Unit Price','Total Price','Invoice Number']
        cols_prof = ['Reference','Code EAN','Description','Origin',
                     'Quantity','Unit Price','Total Price']

        headers = cols_fact if 'Invoice Number' in records[0] else cols_prof

        # Generar Excel en memoria
        wb = Workbook(); ws = wb.active; ws.append(headers)
        for r in records: ws.append([r.get(h,'') for h in headers])
        buf = BytesIO(); wb.save(buf); buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
       
