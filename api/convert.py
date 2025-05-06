import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 1) PATRONES ─────────────────────────────────────────────
INV_PAT = re.compile(
    r'(?:FACTURE|INVOICE)[^\d]{0,800}?(\d{6,})',
    re.I | re.S
)
PLV_PAT = re.compile(
    r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT',
    re.I
)

ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
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

# ─── 2) ENDPOINT ─────────────────────────────────────────────
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

            # Definimos estado por PDF
            kind         = doc_kind(pdfplumber.open(tmp.name).pages[0].extract_text() or '')
            current_inv  = ''
            add_plv      = False
            current_org  = ''
            pending_rows = []

            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ''

                    # ─── A) A nivel de página detecto invoice y PLV ─────
                    if m := INV_PAT.search(txt):
                        current_inv = m.group(1)
                        add_plv     = bool(PLV_PAT.search(txt))
                        # relleno las filas que quedaron “pendientes”
                        for r in pending_rows:
                            r['Invoice Number'] = current_inv + ('PLV' if add_plv else '')
                        pending_rows.clear()

                    # ─── B) Extraigo líneas y actualizo país cada vez que aparece
                    lines = txt.split('\n')
                    for idx, ln in enumerate(lines):
                        if "PAYS D'ORIGINE" in ln.upper():
                            after = ln.split(':',1)[1].strip() if ':' in ln else ''
                            if not after and idx+1 < len(lines):
                                after = lines[idx+1].strip()
                            if after:
                                current_org = after

                    # ─── C) Recorro líneas para extraer filas
                    for i, raw in enumerate(lines):
                        line = raw.strip()

                        # FACTURA
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            inv_full = current_inv + ('PLV' if add_plv else '')
                            row = {
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : custom,
                                'Description'   : desc,
                                'Origin'        : current_org,
                                'Quantity'      : int(qty_s.replace('.','').replace(',','')),
                                'Unit Price'    : fnum(unit_s),
                                'Total Price'   : fnum(tot_s),
                                'Invoice Number': inv_full
                            }
                            rows.append(row)
                            if not current_inv:  # si aún no había header
                                pending_rows.append(row)
                            continue

                        # PROFORMA
                        if kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if i+1 < len(lines):
                                desc = lines[i+1].strip()
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            inv_full = current_inv + ('PLV' if add_plv else '')
                            row = {
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : '',
                                'Description'   : desc,
                                'Origin'        : current_org,
                                'Quantity'      : qty,
                                'Unit Price'    : unit,
                                'Total Price'   : unit*qty,
                                'Invoice Number': inv_full
                            }
                            rows.append(row)
                            if not current_inv:
                                pending_rows.append(row)

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ─── 3) Rellenar Origin vacío si es único ────────────────
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin']:
                orgs = inv_to_org.get(r['Invoice Number'], set())
                if len(orgs) == 1:
                    r['Origin'] = next(iter(orgs))

        # ─── 4) Exportar a Excel ────────────────────────────────
        cols = ['Reference','Code EAN','Custom Code','Description',
                'Origin','Quantity','Unit Price','Total Price','Invoice Number']
        wb = Workbook(); ws = wb.active; ws.append(cols)
        for r in rows:
            ws.append([r.get(c,'') for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

