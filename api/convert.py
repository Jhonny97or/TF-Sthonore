# ─── 1) IMPORTS ────────────────────────────────────────────────────
import logging, re, tempfile, traceback, os
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
        files = request.files.getlist('file')
        if not files:
            return 'No file(s) uploaded', 400

        rows = []

        # ── Procesar cada PDF ──
        for uploaded in files:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            uploaded.save(tmp.name)

            first = extract_text(tmp.name, page_numbers=[0]) or ''
            kind  = detect_doc_type(first)

            inv_base = INV_PAT.search(first).group(1) if INV_PAT.search(first) else ''
            add_plv  = 'FACTURE SANS PAIEMENT' in first.upper()

            origin = ''         # se actualizará línea por línea

            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    lines = (page.extract_text() or '').split('\n')

                    i = 0
                    while i < len(lines):
                        line = lines[i].strip()

                        # ① ¿Es línea que declara origen?
                        if (mo := ORIG_PAT.match(line)):
                            origin = mo.group(1).strip()
                            i += 1
                            continue   # siguiente línea

                        # ② ¿Fila de FACTURA?
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]) else ''
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

                        # ③ ¿Fila de PROFORMA?
                        elif kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': desc,
                                'Origin': origin,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit * qty,
                                'Invoice Number': inv_base + ('PLV' if add_plv else '')
                            })
                        i += 1

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ── Generar Excel ──
        cols = [
            'Reference','Code EAN','Custom Code','Description',
            'Origin','Quantity','Unit Price','Total Price','Invoice Number'
        ]
        wb = Workbook(); ws = wb.active
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c, '') for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

