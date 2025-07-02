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

# Line item patterns for different invoice types
ROW_FACT     = re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)$')
ROW_PROF_DIOR= re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,10})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)$')
ROW_PROF     = re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)$')
# New format: No., Description, UPC, Country, HS Code, Quantity, U.of M., Unit Price, POSM/FOC, Line Amount
ROW_INV2     = re.compile(
    r'^(\d+)\s+(.+?)\s*(\d{12,14})\s+([A-Z]{2})\s+([\d.]+\.[\d.]+\.[\d.]+)\s+(\d+)\s+([^\s]+)\s+([\d.,]+)\s+([\-\d.,]+)\s+([\d.,]+)$'
)

COLS = [
    'Reference','Code EAN','Custom Code','Description',
    'Origin','Quantity','Unit Price','Total Price','Invoice Number'
]

def fnum(s: str) -> float:
    s = s.strip()
    if not s:
        return 0.0
    # Detect US vs EU style
    if ',' in s and '.' in s:
        if s.index(',') < s.index('.'):
            return float(s.replace(',', ''))
        return float(s.replace('.', '').replace(',', '.'))
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

                inv_global = ''
                plv_global = False
                if kind == 'factura':
                    if m := INV_PAT.search(all_txt): inv_global = m.group(1)
                    if PLV_PAT.search(all_txt): plv_global = True
                else:
                    if m := PROF_PAT.search(all_txt): inv_global = m.group(1)
                    elif m := ORDER_PAT_EN.search(all_txt): inv_global = m.group(1)
                    elif m := ORDER_PAT_FR.search(all_txt): inv_global = m.group(1)

                invoice_full = inv_global + ('PLV' if plv_global else '')
                org_global = ''

                # Iterate all pages fully (no premature break)
                for page in pdf.pages:
                    lines = (page.extract_text() or '').split('\n')
                    # Update origin if found
                    for ln in lines:
                        if mo := ORG_PAT.search(ln):
                            org_global = mo.group(1).strip() or org_global

                    # Parse each line, including merges for multi-line descriptions
                    i = 0
                    while i < len(lines):
                        ln = lines[i].strip()
                        merged = ''
                        # Try merge with next line
                        if kind == 'factura' and ln and ln[0].isdigit() and not ROW_INV2.match(ln) and i+1 < len(lines):
                            cand = ln + ' ' + lines[i+1].strip()
                            if ROW_INV2.match(cand):
                                merged, i = cand, i+1
                        # Try merge with previous line
                        if not merged and kind == 'factura' and ln and ln[0].isdigit() and not ROW_INV2.match(ln) and i > 0:
                            cand = lines[i-1].strip() + ' ' + ln
                            if ROW_INV2.match(cand):
                                merged = cand

                        target = merged or ln
                        # Match standard invoice
                        if kind == 'factura' and (mf := ROW_FACT.match(target)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if not merged and i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': org_global,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })
                        # Match new format
                        elif kind == 'factura' and (mn2 := ROW_INV2.match(target)):
                            no_s, desc, upc, orig_code, hs_code, qty_s, uom, unit_s, posm_s, tot_s = mn2.groups()
                            rows.append({
                                'Reference': no_s,
                                'Code EAN': upc,
                                'Custom Code': hs_code,
                                'Description': desc,
                                'Origin': orig_code or org_global,
                                'Quantity': int(qty_s.replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })
                        # Match proforma DIOR
                        elif kind == 'proforma' and (mpd := ROW_PROF_DIOR.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mpd.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': org_global,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })
                        # Match generic proforma
                        elif kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty_val = int(qty_s.replace('.', '').replace(',', ''))
                            unit_val = fnum(unit_s)
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': desc,
                                'Origin': org_global,
                                'Quantity': qty_val,
                                'Unit Price': unit_val,
                                'Total Price': unit_val * qty_val,
                                'Invoice Number': invoice_full
                            })
                        i += 1
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # Complete origin if unique per invoice
        inv2org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv2org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv2org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv2org[r['Invoice Number']]))

        # Build Excel
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
        logging.exception("Error in /convert")
        return f'<pre>{traceback.format_exc()}</pre>', 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')


