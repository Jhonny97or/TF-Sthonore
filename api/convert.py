"""
Serverless para Vercel – convierte FACTURA o PROFORMA en Excel.
Corrige duplicados de Invoice Number y asigna correctamente el Origin.

Salidas finales
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

# ───────────────── helpers ─────────────────
def parse_num(s: str) -> float:
    s = (s or '').strip()
    return float(s.replace('.', '').replace(',', '.')) if s else 0.0


def detect_doc_type(first_page: str) -> str:
    up = first_page.upper()
    if ('ACCUSE' in up and 'RECEPTION' in up) or 'ACKNOWLEDGE' in up or 'PROFORMA' in up:
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('No pude determinar si es factura o proforma.')


# ───────────────── endpoint ─────────────────
@app.route('/', methods=['POST'])            # para curl local
@app.route('/api/convert', methods=['POST'])  # ruta Vercel
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        pdf_file = request.files['file']

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        pdf_file.save(tmp.name)

        first_txt = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type  = detect_doc_type(first_txt)

        records: list[dict] = []

        with pdfplumber.open(tmp.name) as pdf:
            # ───────── FACTURA ─────────
            if doc_type == 'factura':
                # base num factura
                m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,50}(\d{6,})', first_txt, re.I)
                invoice_base = m.group(1) if m else None
                if not invoice_base:
                    hits = re.findall(r'\d{8,}', first_txt)
                    invoice_base = hits[0] if hits else None
                if not invoice_base:
                    return 'No invoice number found', 400

                row_pat = re.compile(
                    r'^([A-Z]\d{5,7})\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s*$'
                )

                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')
                    up    = txt.upper()

                    # ¿lleva sufijo PLV?
                    inv_num = invoice_base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else invoice_base

                    # detectar origin presente en la propia página
                    origin_page = ''
                    match_origin = re.search(r"PAYS D'ORIGINE\s*:\s*(.+)", up)
                    if match_origin:
                        origin_page = match_origin.group(1).strip().title()

                    i = 0
                    while i < len(lines):
                        mo = row_pat.match(lines[i].strip())
                        if mo:
                            ref, ean, custom, qty_s, unit_s, tot_s = mo.groups()

                            # descripción = línea siguiente si no es otra fila
                            desc = ''
                            if i + 1 < len(lines) and not row_pat.match(lines[i+1].strip()):
                                desc = lines[i+1].strip()

                            # si la descripción incluye país podemos sobrescribir origin
                            origin_inline = ''
                            for word in ('Union Européenne', 'China'):
                                if word.lower() in desc.lower():
                                    origin_inline = word
                                    break

                            records.append({
                                'Reference':      ref,
                                'Code EAN':       ean,
                                'Custom Code':    custom,
                                'Description':    desc,
                                'Origin':         origin_inline or origin_page,
                                'Quantity':       int(qty_s),
                                'Unit Price':     parse_num(unit_s),
                                'Total Price':    parse_num(tot_s),
                                'Invoice Number': inv_num
                            })
                            i += 1          # saltar desc
                        i += 1

            # ───────── PROFORMA ─────────
            else:
                row_pat = re.compile(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    # origin de la página
                    origin_page = ''
                    match_origin = re.search(r"PAYS D'ORIGINE\s*:\s*(.+)", txt, re.I)
                    if match_origin:
                        origin_page = match_origin.group(1).strip()

                    for idx, line in enumerate(lines):
                        m = row_pat.search(line)
                        if m:
                            ref, ean, unit_s, qty_s = m.groups()
                            desc = lines[idx+1].strip() if idx+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.', '').replace(',', '').strip())
                            unit = parse_num(unit_s)
                            records.append({
                                'Reference':   ref,
                                'Code EAN':    ean,
                                'Description': desc,
                                'Origin':      origin_page,
                                'Quantity':    qty,
                                'Unit Price':  unit,
                                'Total Price': unit*qty
                            })

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        # columnas finales
        cols_fact = ['Reference','Code EAN','Custom Code','Description','Origin',
                     'Quantity','Unit Price','Total Price','Invoice Number']
        cols_prof = ['Reference','Code EAN','Description','Origin',
                     'Quantity','Unit Price','Total Price']
        headers = cols_fact if 'Invoice Number' in records[0] else cols_prof

        wb = Workbook(); ws = wb.active; ws.append(headers)
        for r in records:
            ws.append([r.get(h,'') for h in headers])

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


