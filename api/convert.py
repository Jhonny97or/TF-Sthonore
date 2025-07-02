import logging
import re
import tempfile
import os
import traceback
from io import BytesIO

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook
from collections import defaultdict

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── PATTERNS ──────────────────────────────────────────────────────
INV_PAT      = re.compile(r'(?:FACTURE|INVOICE)\D*(\d{6,})', re.I)
PROF_PAT     = re.compile(r'PROFORMA[\s\S]*?(\d{6,})', re.I)
ORDER_PAT_EN = re.compile(r'ORDER\s+NUMBER\D*(\d{6,})', re.I)
ORDER_PAT_FR = re.compile(r'N°\s*DE\s*COMMANDE\D*(\d{6,})', re.I)
PLV_PAT      = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT      = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)
HEADER_PAT   = re.compile(r'^No\.\s+Description\b', re.I)
SUMMARY_PAT  = re.compile(r'^\s*Total\s+before\s+discount', re.I)

# Line item patterns
ROW_FACT     = re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)$')
ROW_PROF_DIOR= re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,10})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)$')
ROW_PROF     = re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)$')
# Updated ROW_INV2: allow thousands separators in quantity and more flexible whitespace
ROW_INV2     = re.compile(
    r'^(\d+)\s+(.+?)\s+(\d{12,14})\s+([A-Z]{2})\s+([\d.]+\.[\d.]+\.[\d.]+)\s+([\d.,]+)\s+[^\d]*?([\d.,]+)\s+[\d.,-]*\s+([\d.,]+)$'
)
# groups: ref, desc, upc, country, hs, qty, unit, total

COLS = ['Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number']


def fnum(s: str) -> float:
    s = s.strip()
    if not s:
        return 0.0
    if ',' in s and '.' in s:
        return float(s.replace(',', '')) if s.index(',') < s.index('.') else float(s.replace('.', '').replace(',', '.'))
    return float(s.replace(',', '').replace(' ', ''))


def doc_kind(text: str) -> str:
    up = text.upper()
    return 'proforma' if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up) else 'factura'

@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        pdfs = request.files.getlist('file')
        if not pdfs:
            return 'No file(s) uploaded', 400

        rows = []
        for pdf_file in pdfs:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                all_txt = '\n'.join(page.extract_text() or '' for page in pdf.pages)
                kind = doc_kind(all_txt)

                inv_no = ''
                if kind == 'factura':
                    m = INV_PAT.search(all_txt)
                    inv_no = m.group(1) if m else ''
                else:
                    m = PROF_PAT.search(all_txt) or ORDER_PAT_EN.search(all_txt) or ORDER_PAT_FR.search(all_txt)
                    inv_no = m.group(1) if m else ''
                invoice_full = inv_no + ('PLV' if PLV_PAT.search(all_txt) else '')

                origin_global = ''
                for page in pdf.pages:
                    txt = page.extract_text() or ''
                    if SUMMARY_PAT.search(txt):
                        break
                    lines = txt.split('\n')
                    for ln in lines:
                        if mo := ORG_PAT.search(ln):
                            origin_global = mo.group(1).strip() or origin_global

                    capturing = False
                    for idx, ln in enumerate(lines):
                        ln_strip = ln.strip()
                        if not capturing and HEADER_PAT.match(ln_strip):
                            capturing = True
                            continue
                        if not capturing or not ln_strip:
                            continue

                        # Attempt full-line match first
                        if kind == 'factura' and (m2 := ROW_INV2.match(ln_strip)):
                            ref, desc, upc, ctry, hs, qty_s, unit_s, tot_s = m2.groups()
                            rows.append({
                                'Reference': ref,
                                'Code EAN': upc,
                                'Custom Code': hs,
                                'Description': desc,
                                'Origin': ctry or origin_global,
                                'Quantity': int(qty_s.replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })
                            continue
                        # If line starts with number but not matched yet, try merge with next line
                        if kind == 'factura' and ln_strip[0].isdigit() and idx+1 < len(lines):
                            merged = ln_strip + ' ' + lines[idx+1].strip()
                            if m2 := ROW_INV2.match(merged):
                                ref, desc, upc, ctry, hs, qty_s, unit_s, tot_s = m2.groups()
                                rows.append({
                                    'Reference': ref,
                                    'Code EAN': upc,
                                    'Custom Code': hs,
                                    'Description': desc,
                                    'Origin': ctry or origin_global,
                                    'Quantity': int(qty_s.replace(',', '')),
                                    'Unit Price': fnum(unit_s),
                                    'Total Price': fnum(tot_s),
                                    'Invoice Number': invoice_full
                                })
                                continue

                os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in rows:
            ws.append([r[c] for c in COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='extracted_data.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.exception('Error')
        return f'<pre>{traceback.format_exc()}</pre>', 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')

