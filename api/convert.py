import logging
import re
import tempfile
import os
import traceback
from io import BytesIO

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── REGEX PATRONES ───────────────────────────────────────────────
INV_PAT      = re.compile(r'(?:FACTURE|INVOICE)\D*(\d{6,})', re.I)
PROF_PAT     = re.compile(r'PROFORMA[\s\S]*?(\d{6,})', re.I)
ORDER_PAT_EN = re.compile(r'ORDER\s+NUMBER\D*(\d{6,})', re.I)
ORDER_PAT_FR = re.compile(r'N°\s*DE\s*COMMANDE\D*(\d{6,})', re.I)
PLV_PAT      = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT      = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)
HEADER_PAT   = re.compile(r'^No\.\s+Description', re.I)
SUMMARY_PAT  = re.compile(r'^\s*Total\s+before\s+discount', re.I)

# Nuevo: permite palabra "Each" y campo POSM opcional
ROW_LINE = re.compile(
    r'^(?P<ref>\d+)\s+(?P<desc>.+?)\s+(?P<upc>\d{12,14})\s+(?P<ctry>[A-Z]{2})\s+'
    r'(?P<hs>\d{4}\.\d{2}\.\d{4})\s+(?P<qty>[\d,.]+)\s+Each\s+'
    r'(?P<unit>[\d.]+)\s+(?:-|[\d.,]+)\s+(?P<total>[\d.,]+)$'
)

COLS = [
    'Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number'
]

# ─── UTIL ─────────────────────────────────────────────────────────

def fnum(s: str) -> float:
    s = s.replace('\u202f', '').strip()
    if not s:
        return 0.0
    if ',' in s and '.' in s:
        return float(s.replace(',', '')) if s.find(',') < s.find('.') else float(s.replace('.', '').replace(',', '.'))
    return float(s.replace(',', '').replace('.', '').replace(' ', ''))


def doc_kind(text: str) -> str:
    up = text.upper()
    return 'proforma' if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up) else 'factura'

# ─── ENDPOINT ─────────────────────────────────────────────────────

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
                full_txt = "\n".join(p.extract_text() or '' for p in pdf.pages)
                kind = doc_kind(full_txt)
                inv_no = ''
                if kind == 'factura':
                    if m := INV_PAT.search(full_txt):
                        inv_no = m.group(1)
                else:
                    m = PROF_PAT.search(full_txt) or ORDER_PAT_EN.search(full_txt) or ORDER_PAT_FR.search(full_txt)
                    inv_no = m.group(1) if m else ''
                invoice_full = inv_no + ('PLV' if PLV_PAT.search(full_txt) else '')

                origin_global = ''
                stop = False
                for pg in pdf.pages:
                    txt_lines = (pg.extract_text() or '').split('\n')
                    # actualizar origen
                    for t in txt_lines:
                        if mo := ORG_PAT.search(t):
                            origin_global = mo.group(1).strip()
                    capturing = False
                    i = 0
                    while i < len(txt_lines) and not stop:
                        line = txt_lines[i].strip()
                        if SUMMARY_PAT.search(line):
                            stop = True
                            break
                        if not capturing and HEADER_PAT.match(line):
                            capturing = True
                            i += 1
                            continue
                        if not capturing:
                            i += 1
                            continue
                        if not line:
                            i += 1
                            continue

                        merged = line
                        # unir hasta 3 líneas mientras no matchea
                        for extra in range(1,4):
                            if ROW_LINE.match(merged):
                                break
                            if i+extra < len(txt_lines):
                                merged += ' ' + txt_lines[i+extra].strip()
                        if mrow := ROW_LINE.match(merged):
                            gd = mrow.groupdict()
                            rows.append({
                                'Reference': gd['ref'],
                                'Code EAN': gd['upc'],
                                'Custom Code': gd['hs'],
                                'Description': gd['desc'],
                                'Origin': gd['ctry'] or origin_global,
                                'Quantity': int(gd['qty'].replace(',', '').replace('.', '')),
                                'Unit Price': fnum(gd['unit']),
                                'Total Price': fnum(gd['total']),
                                'Invoice Number': invoice_full
                            })
                        i += 1
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ─── Excel ────────────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in rows:
            ws.append([r[c] for c in COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf,
                         as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.exception('Error')
        return f'<pre>{traceback.format_exc()}</pre>', 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
