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

# ─── REGEX ─────────────────────────────────────────────────────────
INV_PAT        = re.compile(r'(?:FACTURE|INVOICE)[\s\S]{0,1000}?(\d{6,})', re.I)
PLV_PAT        = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT        = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)

ORDER_PAT_EN   = re.compile(r'ORDER\s+NUMBER\s*/?\s*:?\s*(\d{6,})', re.I)
ORDER_PAT_FR   = re.compile(r'N°\s*DE\s*COMMANDE[^\d]*(\d{6,})', re.I)
PROF_NUM_PAT   = re.compile(r'PROFORMA[^\d]{0,20}?(\d{6,})', re.I)

ROW_FACT       = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$')
ROW_PROF_DIOR  = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,10})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$')
ROW_PROF       = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)\s*$')

COLS = [
    'Reference','Code EAN','Custom Code','Description',
    'Origin','Quantity','Unit Price','Total Price','Invoice Number'
]

def fnum(s: str) -> float:
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def doc_kind(text: str) -> str:
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    return 'factura'

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
                # 1) Determinar tipo y extraer invoice/proforma número de **todas** las páginas
                first_txt = pdf.pages[0].extract_text() or ''
                kind = doc_kind(first_txt)

                inv_global = ''
                plv_global = False

                for page in pdf.pages:
                    txt = page.extract_text() or ''

                    if kind == 'factura':
                        if m := INV_PAT.search(txt):
                            inv_global = m.group(1)
                        if PLV_PAT.search(txt):
                            plv_global = True
                    else:  # proforma
                        if p := PROF_NUM_PAT.search(txt):
                            inv_global = p.group(1)
                        elif e := ORDER_PAT_EN.search(txt):
                            inv_global = e.group(1)
                        elif f := ORDER_PAT_FR.search(txt):
                            inv_global = f.group(1)

                invoice_full = inv_global + ('PLV' if plv_global else '')

                # 2) Segundopaso: extraer filas usando el invoice_full fijo
                org_global = ''
                for page in pdf.pages:
                    txt = page.extract_text() or ''
                    lines = txt.split('\n')

                    # actualizar país de origen
                    for ln in lines:
                        if mo := ORG_PAT.search(ln):
                            val = mo.group(1).strip()
                            if val:
                                org_global = val

                    # extraer líneas de detalle
                    for i, raw in enumerate(lines):
                        ln = raw.strip()

                        # factura
                        if kind == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = lines[i+1].strip() if i+1<len(lines) and not ROW_FACT.match(lines[i+1]) else ''
                            rows.append({
                                'Reference': ref, 'Code EAN': ean, 'Custom Code': custom,
                                'Description': desc, 'Origin': org_global,
                                'Quantity': int(qty_s.replace('.','').replace(',','')),
                                'Unit Price': fnum(unit_s), 'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })

                        # proforma Dior (6 cols)
                        elif kind == 'proforma' and (mpd := ROW_PROF_DIOR.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mpd.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            rows.append({
                                'Reference': ref, 'Code EAN': ean, 'Custom Code': custom,
                                'Description': desc, 'Origin': org_global,
                                'Quantity': int(qty_s.replace('.','').replace(',','')),
                                'Unit Price': fnum(unit_s), 'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })

                        # proforma genérica (4 cols)
                        elif kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            rows.append({
                                'Reference': ref, 'Code EAN': ean, 'Custom Code': '',
                                'Description': desc, 'Origin': org_global,
                                'Quantity': qty, 'Unit Price': unit,
                                'Total Price': unit * qty,
                                'Invoice Number': invoice_full
                            })

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # rellenar origen faltante si solo hay uno por invoice
        inv2org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv2org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv2org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv2org[r['Invoice Number']]))

        # generar Excel
        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in rows:
            ws.append([r[c] for c in COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.exception("Error in /convert")
        return f'<pre>{traceback.format_exc()}</pre>', 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
