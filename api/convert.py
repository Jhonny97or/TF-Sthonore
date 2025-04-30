"""PDF → Excel para facturas/proformas LVMH & Dior (serverless Vercel)."""

import logging, re, tempfile, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ───────────────── helpers ─────────────────
def num(s: str) -> float:
    s = (s or '').strip()
    return float(s.replace('.', '').replace(',', '.')) if s else 0.0

def detect_type(txt: str) -> str:
    up = txt.upper()
    if any(k in up for k in ('ACKNOWLEDGE', 'ACCUSE', 'PROFORMA')):  return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:                           return 'factura'
    raise ValueError('No pude determinar tipo (factura/proforma).')

ORIGIN_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

ROW_FACT   = re.compile(
    r'^([A-Z]\d{4,12})\s+'      # referencia (≥5)
    r'(\d{12,14})\s+'           # EAN
    r'(\d{6,9})\s+'             # nomenclature
    r'(\d[\d.,]*)\s+'           # qty
    r'([\d.,]+)\s+'             # unit
    r'([\d.,]+)\s*$',           # total
)

ROW_PROF   = re.compile(
    r'([A-Z]\d{4,12})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

# ───────────────── endpoint ─────────────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        pdf_file = request.files['file']
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        pdf_file.save(tmp.name)

        first = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type = detect_type(first)

        records = []
        with pdfplumber.open(tmp.name) as pdf:

            # ─────── FACTURA ───────
            if doc_type == 'factura':

                # n° base factura
                base = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', first, re.I)
                inv_base = base.group(1) if base else None
                if not inv_base:                        # respaldo
                    m = re.search(r'\b(\d{8,})\b', first)
                    inv_base = m.group(1) if m else ''
                if not inv_base:
                    return 'No invoice number found', 400

                for pg in pdf.pages:
                    txt   = pg.extract_text() or ''
                    lines = txt.split('\n')
                    up    = txt.upper()
                    inv   = inv_base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else inv_base

                    # Origin cabecera
                    origin_page = ''
                    m_org = ORIGIN_PAT.search(txt)
                    if m_org: origin_page = m_org.group(1).strip()

                    i = 0
                    while i < len(lines):
                        mrow = ROW_FACT.match(lines[i].strip())
                        if mrow:
                            ref, ean, custom, qty_s, unit_s, tot_s = mrow.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]) else ''

                            # si descripción tiene “China” o “Union Européenne…”
                            origin_inline = ''
                            if 'CHINA' in desc.upper():
                                origin_inline = 'China'
                            elif 'UNION EUROP' in desc.upper():
                                origin_inline = 'Union Européenne/European Union'

                            records.append({
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : custom,
                                'Description'   : desc,
                                'Origin'        : origin_inline or origin_page,
                                'Quantity'      : int(qty_s.replace('.','').replace(',','')),
                                'Unit Price'    : num(unit_s),
                                'Total Price'   : num(tot_s),
                                'Invoice Number': inv
                            })
                            i += 1        # salta desc
                        i += 1

            # ─────── PROFORMA ───────
            else:
                for pg in pdf.pages:
                    txt   = pg.extract_text() or ''
                    lines = txt.split('\n')

                    origin_page = ''
                    m_org = ORIGIN_PAT.search(txt)
                    if m_org: origin_page = m_org.group(1).strip()

                    for idx, line in enumerate(lines):
                        m = ROW_PROF.search(line)
                        if m:
                            ref, ean, unit_s, qty_s = m.groups()
                            desc = lines[idx+1].strip() if idx+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = num(unit_s)
                            records.append({
                                'Reference'  : ref,
                                'Code EAN'   : ean,
                                'Description': desc,
                                'Origin'     : origin_page,
                                'Quantity'   : qty,
                                'Unit Price' : unit,
                                'Total Price': unit*qty
                            })

        # ───── fallback: texto completo ─────
        if not records:
            full = extract_text(tmp.name) or ''
            for m in ROW_FACT.finditer(full):
                ref, ean, custom, qty_s, unit_s, tot_s = m.groups()
                records.append({
                    'Reference': ref, 'Code EAN': ean, 'Custom Code': custom,
                    'Quantity': int(qty_s), 'Unit Price': num(unit_s),
                    'Total Price': num(tot_s), 'Description':'', 'Origin':'', 'Invoice Number':''
                })
        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        # columnas
        cols_fact = ['Reference','Code EAN','Custom Code','Description','Origin',
                     'Quantity','Unit Price','Total Price','Invoice Number']
        cols_prof = ['Reference','Code EAN','Description','Origin',
                     'Quantity','Unit Price','Total Price']
        headers = cols_fact if 'Invoice Number' in records[0] else cols_prof

        wb, ws = Workbook(), Workbook().active
        ws.append(headers)
        for r in records: ws.append([r.get(h,'') for h in headers])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error interno:\n{traceback.format_exc()}', 500
