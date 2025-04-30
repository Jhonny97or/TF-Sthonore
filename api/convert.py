"""PDF → Excel (factura / proforma) – función /api/convert para Vercel"""

# ─── 1) IMPORTS ────────────────────────────────────────────────────
import logging, re, tempfile, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 2) HELPERS ────────────────────────────────────────────────────
def fnum(s: str) -> float:
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

def detect_doc_type(txt: str) -> str:
    up = txt.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('Tipo de documento no reconocido.')

# ─── 3) REGEX ──────────────────────────────────────────────────────
INV_PAT  = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', re.I)
ORIG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)
ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

# ─── 4) ENDPOINT ───────────────────────────────────────────────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        request.files['file'].save(tmp.name)

        first = extract_text(tmp.name, page_numbers=[0]) or ''
        kind  = detect_doc_type(first)

        # contexto inicial
        inv_base = (INV_PAT.search(first).group(1)
                    if INV_PAT.search(first) else '')
        add_plv  = 'FACTURE SANS PAIEMENT' in first.upper()
        origin   = (ORIG_PAT.search(first).group(1).strip()
                    if ORIG_PAT.search(first) else '')

        rows = []

        with pdfplumber.open(tmp.name) as pdf:
            for page in pdf.pages:
                txt   = page.extract_text() or ''
                lines = txt.split('\n')
                up    = txt.upper()

                if (m := INV_PAT.search(txt)):
                    inv_base = m.group(1)
                    add_plv  = 'FACTURE SANS PAIEMENT' in up
                if (o := ORIG_PAT.search(txt)):
                    origin = o.group(1).strip()

                i = 0
                while i < len(lines):
                    line = lines[i].strip()

                    if kind == 'factura':
                        if (mo := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mo.groups()
                            desc = (lines[i+1].strip()
                                    if i+1 < len(lines) and not ROW_FACT.match(lines[i+1])
                                    else '')
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': origin,
                                'Quantity': qty,
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': inv_base + ('PLV' if add_plv else '')
                            })
                            i += 1
                    else:  # proforma
                        if (mp := ROW_PROF.search(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Description': desc,
                                'Origin': origin,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit*qty
                            })
                    i += 1

        if not rows:
            return 'Sin registros extraídos', 400

        cols = (['Reference','Code EAN','Custom Code','Description','Origin',
                 'Quantity','Unit Price','Total Price','Invoice Number']
                if 'Invoice Number' in rows[0] else
                ['Reference','Code EAN','Description','Origin',
                 'Quantity','Unit Price','Total Price'])

        wb = Workbook()
        ws = wb.active
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c,'') for c in cols])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

