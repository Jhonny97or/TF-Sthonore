"""
Serverless para Vercel: convierte FACTURA o PROFORMA en Excel
con todas las columnas que necesitas.

• Factura  → Reference, Code EAN, Custom Code, Description,
             Origin, Quantity, Unit Price, Total Price, Invoice Number
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
    s = (s or '').strip()
    return float(s.replace('.', '').replace(',', '.')) if s else 0.0


def detect_doc_type(first_page: str) -> str:
    up = first_page.upper()
    if ('ACCUSE' in up and 'RECEPTION' in up) or 'ACKNOWLEDGE' in up or 'PROFORMA' in up:
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('No pude determinar si es factura o proforma.')


# ───────────────────────── endpoint ─────────────────────────
@app.route('/', methods=['POST'])             # útil local
@app.route('/api/convert', methods=['POST'])  # ruta Vercel
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        pdf_file = request.files['file']

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        pdf_file.save(tmp.name)

        first_text = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type   = detect_doc_type(first_text)

        records: list[dict] = []

        with pdfplumber.open(tmp.name) as pdf:
            # ────────────── FACTURA ──────────────
            if doc_type == 'factura':
                # Invoice base
                m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,50}(\d{6,})', first_text, re.I)
                invoice_base = m.group(1) if m else None
                if not invoice_base:
                    cands = re.findall(r'\d{8,}', first_text)
                    invoice_base = cands[0] if cands else None
                if not invoice_base:
                    return 'No invoice number found', 400

                pattern = re.compile(
                    r'^([A-Z]\d+)\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$'
                )
                origin = ''

                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')
                    up    = txt.upper()
                    inv   = invoice_base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else invoice_base

                    # origen, si aparece
                    for L in lines:
                        if "PAYS D'ORIGINE" in L:
                            origin = L.split(':', 1)[-1].strip()

                    i = 0
                    while i < len(lines):
                        mm = pattern.match(lines[i].strip())
                        if mm:
                            ref, ean, custom, qty_s, unit_s, tot_s = mm.groups()
                            desc = lines[i + 1].strip() if i + 1 < len(lines) else ''
                            records.append({
                                'Reference':      ref,
                                'Code EAN':       ean,
                                'Custom Code':    custom,
                                'Description':    desc,
                                'Origin':         origin,
                                'Quantity':       int(qty_s),
                                'Unit Price':     parse_num(unit_s),
                                'Total Price':    parse_num(tot_s),
                                'Invoice Number': inv
                            })
                            i += 1          # salta descripción
                        i += 1

            # ────────────── PROFORMA ──────────────
            else:
                origin = ''
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    for L in lines:
                        if "PAYS D'ORIGINE" in L:
                            origin = L.split(':', 1)[-1].strip()

                    for idx, line in enumerate(lines):
                        m = re.search(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)', line)
                        if m:
                            ref, ean, unit_s, qty_s = m.groups()
                            desc = lines[idx + 1].strip() if idx + 1 < len(lines) else ''
                            qty  = int(qty_s.replace('.', '').replace(',', '').strip())
                            unit = parse_num(unit_s)
                            records.append({
                                'Reference':   ref,
                                'Code EAN':    ean,
                                'Description': desc,
                                'Origin':      origin,
                                'Quantity':    qty,
                                'Unit Price':  unit,
                                'Total Price': unit * qty
                            })

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        # columnas finales
        cols_fact = ['Reference', 'Code EAN', 'Custom Code', 'Description',
                     'Origin', 'Quantity', 'Unit Price', 'Total Price', 'Invoice Number']
        cols_prof = ['Reference', 'Code EAN', 'Description', 'Origin',
                     'Quantity', 'Unit Price', 'Total Price']
        headers = cols_fact if 'Invoice Number' in records[0] else cols_prof

        wb = Workbook(); ws = wb.active; ws.append(headers)
        for r in records:
            ws.append([r.get(h, '') for h in headers])

        buf = BytesIO(); wb.save(buf); buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error interno:\n{traceback.format_exc()}', 500


