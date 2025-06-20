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

# ─── PATRONES ──────────────────────────────────────────────────────
# Factura/invoice con al menos 6 dígitos
INV_PAT      = re.compile(r'(?:FACTURE|INVOICE)\D*(\d{6,})', re.I)
# Proforma: “PROFORMA” seguido de cualquier texto (incluyendo salto de línea) hasta el número
PROF_PAT     = re.compile(r'PROFORMA[\s\S]*?(\d{6,})', re.I)
# Pedidos EN/FR
ORDER_PAT_EN = re.compile(r'ORDER\s+NUMBER\D*(\d{6,})', re.I)
ORDER_PAT_FR = re.compile(r'N°\s*DE\s*COMMANDE\D*(\d{6,})', re.I)
# Sufijo “sin pago”
PLV_PAT      = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
# País de origen
ORG_PAT      = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

# Líneas de detalle
ROW_FACT     = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF_DIOR= re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,10})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF     = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)\s*$'
)

COLS = [
    'Reference','Code EAN','Custom Code','Description',
    'Origin','Quantity','Unit Price','Total Price','Invoice Number'
]

def fnum(s: str) -> float:
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

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
                # ─── 1) Prima pasada: combinar todo el texto y extraer un único número ─────────
                all_txt = "\n".join(page.extract_text() or '' for page in pdf.pages)
                kind    = doc_kind(all_txt)

                inv_global = ''
                plv_global = False

                if kind == 'factura':
                    if m := INV_PAT.search(all_txt):
                        inv_global = m.group(1)
                    if PLV_PAT.search(all_txt):
                        plv_global = True
                else:  # proforma
                    if m := PROF_PAT.search(all_txt):
                        inv_global = m.group(1)
                    elif m := ORDER_PAT_EN.search(all_txt):
                        inv_global = m.group(1)
                    elif m := ORDER_PAT_FR.search(all_txt):
                        inv_global = m.group(1)

                invoice_full = inv_global + ('PLV' if plv_global else '')

                # ─── 2) Segunda pasada: recorrer páginas y extraer filas con ese invoice_full ──
                org_global = ''
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    # actualizar país de origen
                    for ln in lines:
                        if mo := ORG_PAT.search(ln):
                            val = mo.group(1).strip()
                            if val:
                                org_global = val

                    # extraer cada línea de detalle
                    for i, raw in enumerate(lines):
                        ln = raw.strip()

                        if kind == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append({
                                'Reference':      ref,
                                'Code EAN':       ean,
                                'Custom Code':    custom,
                                'Description':    desc,
                                'Origin':         org_global,
                                'Quantity':       int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price':     fnum(unit_s),
                                'Total Price':    fnum(tot_s),
                                'Invoice Number': invoice_full
                            })

                        elif kind == 'proforma' and (mpd := ROW_PROF_DIOR.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mpd.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            rows.append({
                                'Reference':      ref,
                                'Code EAN':       ean,
                                'Custom Code':    custom,
                                'Description':    desc,
                                'Origin':         org_global,
                                'Quantity':       int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price':     fnum(unit_s),
                                'Total Price':    fnum(tot_s),
                                'Invoice Number': invoice_full
                            })

                        elif kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.', '').replace(',', ''))
                            unit = fnum(unit_s)
                            rows.append({
                                'Reference':      ref,
                                'Code EAN':       ean,
                                'Custom Code':    '',
                                'Description':    desc,
                                'Origin':         org_global,
                                'Quantity':       qty,
                                'Unit Price':     unit,
                                'Total Price':    unit * qty,
                                'Invoice Number': invoice_full
                            })

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # completar 'Origin' si sólo hay uno por invoice
        inv2org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv2org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv2org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv2org[r['Invoice Number']]))

        # generar Excel en memoria
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
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.exception("Error in /convert")
        return f'<pre>{traceback.format_exc()}</pre>', 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')

