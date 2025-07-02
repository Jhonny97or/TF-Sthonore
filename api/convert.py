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
SUMMARY_PAT  = re.compile(r'Total\s+before\s+discount', re.I)

ROW_INV2 = re.compile(
    r'^(\d+)\s+(.+?)\s+(\d{12,14})\s+([A-Z]{2})\s+([\d.]+\.[\d.]+\.[\d.]+)\s+([\d.,]+)\s+[^\d]*?([\d.,]+)\s+[\d.,\-]*\s+([\d.,]+)$'
)

COLS = ['Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number']

def fnum(s: str) -> float:
    s = s.strip().replace('\u202f', '')  # remove NBSP if any
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
                text_all = "\n".join(pg.extract_text() or '' for pg in pdf.pages)
                kind = doc_kind(text_all)
                inv_no = (INV_PAT if kind=='factura' else PROF_PAT).search(text_all)
                inv_no = inv_no.group(1) if inv_no else ''
                invoice_full = inv_no + ('PLV' if PLV_PAT.search(text_all) else '')

                origin_global = ''
                stop_capture = False
                for page in pdf.pages:
                    page_lines = (page.extract_text() or '').split('\n')
                    # update global origin
                    for line in page_lines:
                        if mo := ORG_PAT.search(line):
                            origin_global = mo.group(1).strip() or origin_global
                    # iterate lines
                    capturing = False
                    i = 0
                    while i < len(page_lines) and not stop_capture:
                        ln = page_lines[i].strip()
                        if SUMMARY_PAT.search(ln):
                            stop_capture = True
                            break
                        if not capturing and HEADER_PAT.match(ln):
                            capturing = True
                            i += 1
                            continue
                        if not capturing:
                            i += 1
                            continue
                        # if line empty skip
                        if not ln:
                            i += 1; continue
                        # build up to 3‑line candidate for robustness
                        candidate = ln
                        for extra in range(1,3):
                            if ROW_INV2.match(candidate):
                                break
                            if i+extra < len(page_lines):
                                candidate += ' ' + page_lines[i+extra].strip()
                        if m := ROW_INV2.match(candidate):
                            ref, desc, upc, ctry, hs, qty_s, unit_s, tot_s = m.groups()
                            rows.append({
                                'Reference': ref,
                                'Code EAN': upc,
                                'Custom Code': hs,
                                'Description': desc,
                                'Origin': ctry or origin_global,
                                'Quantity': int(qty_s.replace(',', '').replace('.', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })
                        i += 1
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
        return send_file(buf, as_attachment=True, download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.exception('Error')
        return f'<pre>{traceback.format_exc()}</pre>', 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
