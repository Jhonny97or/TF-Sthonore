# ── 1) IMPORTS ─────────────────────────────────────────────────────
import logging, re, tempfile, traceback, os
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ── 2) HELPERS ─────────────────────────────────────────────────────
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

# ── 3) REGEX ───────────────────────────────────────────────────────
INV_PAT  = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', re.I)
ORIG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)

ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

# ── 4) ENDPOINT ────────────────────────────────────────────────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        files = request.files.getlist('file')
        if not files:
            return 'No file(s) uploaded', 400

        rows = []

        for uploaded in files:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            uploaded.save(tmp.name)

            # ———————————————————————————————
            # ① PASADA 1: índice factura → origen
            # ———————————————————————————————
            invoice_origin = {}          # { '102555271': 'Union …' }

            with pdfplumber.open(tmp.name) as pdf:
                cur_invoice = ''
                for page in pdf.pages:
                    txt = page.extract_text() or ''

                    if m_inv := INV_PAT.search(txt):
                        cur_invoice = m_inv.group(1)

                    if m_ori := ORIG_PAT.search(txt):
                        val = m_ori.group(1).strip()
                        if val and cur_invoice and cur_invoice not in invoice_origin:
                            invoice_origin[cur_invoice] = val

            # ———————————————————————————————
            # ② PASADA 2: extraer filas
            # ———————————————————————————————
            first = extract_text(tmp.name, page_numbers=[0]) or ''
            kind  = detect_doc_type(first)

            inv_base = ''    # irá cambiando página a página
            add_plv  = False

            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ''

                    if m_inv := INV_PAT.search(txt):
                        inv_base = m_inv.group(1)
                        add_plv  = 'FACTURE SANS PAIEMENT' in txt.upper()

                    origin = invoice_origin.get(inv_base, '')

                    for line in txt.split('\n'):
                        line = line.strip()

                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': '',   # se podría recuperar como antes si lo necesitas
                                'Origin': origin,
                                'Quantity': int(qty_s.replace('.','').replace(',','')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': inv_base + ('PLV' if add_plv else '')
                            })

                        elif kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            unit = fnum(unit_s)
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': '',
                                'Origin': origin,
                                'Quantity': int(qty_s.replace('.','').replace(',','')),
                                'Unit Price': unit,
                                'Total Price': unit*int(qty_s.replace('.','').replace(',','')),
                                'Invoice Number': inv_base + ('PLV' if add_plv else '')
                            })

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ——— Generar Excel ———
        cols = [
            'Reference','Code EAN','Custom Code','Description',
            'Origin','Quantity','Unit Price','Total Price','Invoice Number'
        ]
        wb = Workbook(); ws = wb.active
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c,'') for c in cols])

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
