# ───  Imports  ──────────────────────────────────────────────
import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ───  Patrones  ─────────────────────────────────────────────
INV_PAT  = re.compile(
    r'(?:FACTURE|INVOICE)[^\d]{0,800}?(\d{6,})',   # ← 800 caracteres y DOTALL
    re.I | re.S
)
PLV_PAT  = re.compile(r'FACTURE\s+SANS\s+PAIEMENT', re.I)   # ← tolera múltiples espacios

ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

def fnum(s):
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def doc_kind(text):
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('Tipo de documento no reconocido.')

# ───  Endpoint  ─────────────────────────────────────────────
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

            kind = doc_kind(pdfplumber.open(tmp.name).pages[0].extract_text() or '')

            current_inv  = ''
            add_plv      = False
            current_org  = ''
            pending_rows = []

            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    txt   = page.extract_text() or ''

                    # 1️⃣  Actualizar nº factura / modo PLV de la página
                    if m := INV_PAT.search(txt):
                        current_inv = m.group(1)
                        add_plv     = bool(PLV_PAT.search(txt))
                        # completar filas pendientes
                        for r in pending_rows:
                            r['Invoice Number'] = current_inv + ('PLV' if add_plv else '')
                        pending_rows.clear()

                    lines = txt.split('\n')
                    i = 0
                    while i < len(lines):
                        line = lines[i].strip()

                        # 2️⃣ ¿Cambia el país de origen justo antes?
                        if "PAYS D'ORIGINE" in line.upper():
                            after = line.split(':',1)[1].strip() if ':' in line else ''
                            if not after and i+1 < len(lines):
                                after = lines[i+1].strip()
                                i += 1        # saltamos la línea usada
                            if after:         # siempre actualizamos; sin break
                                current_org = after
                            i += 1
                            continue

                        # 3️⃣ Rows FACTURA
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            row = {
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : custom,
                                'Description'   : desc,
                                'Origin'        : current_org,
                                'Quantity'      : int(qty_s.replace('.','').replace(',','')),
                                'Unit Price'    : fnum(unit_s),
                                'Total Price'   : fnum(tot_s),
                                'Invoice Number': current_inv + ('PLV' if add_plv else '') if current_inv else ''
                            }
                            rows.append(row)
                            if not current_inv:
                                pending_rows.append(row)
                            i += 1
                            continue

                        # 4️⃣ Rows PROFORMA
                        if kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if i+1 < len(lines):
                                desc = lines[i+1].strip()
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            row = {
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : '',
                                'Description'   : desc,
                                'Origin'        : current_org,
                                'Quantity'      : qty,
                                'Unit Price'    : unit,
                                'Total Price'   : unit*qty,
                                'Invoice Number': current_inv + ('PLV' if add_plv else '') if current_inv else ''
                            }
                            rows.append(row)
                            if not current_inv:
                                pending_rows.append(row)
                        i += 1

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # 5️⃣ Rellenar origen si es único por factura
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])

        for r in rows:
            if not r['Origin']:
                orgs = inv_to_org.get(r['Invoice Number'], set())
                if len(orgs) == 1:
                    r['Origin'] = next(iter(orgs))

        # 6️⃣ Excel
        cols = ['Reference','Code EAN','Custom Code','Description',
                'Origin','Quantity','Unit Price','Total Price','Invoice Number']
        wb = Workbook(); ws = wb.active; ws.append(cols)
        for r in rows:
            ws.append([r.get(c,'') for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

