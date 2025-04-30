import logging, re, tempfile, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ───────── helpers ─────────
def fnum(s: str) -> float:
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

TYPE_PAT  = re.compile(r'(FACTURE SANS PAIEMENT|FACTURE|INVOICE|PROFORMA)', re.I)
INV_PAT   = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,50}(\d{6,})', re.I)
ORIG_PAT  = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

ROW_FACT  = re.compile(
    r'^([A-Z]\w{3,11})\s+'        # Reference
    r'(\d{12,14})\s+'             # EAN
    r'(\d{6,9})\s+'               # Custom/Nomenclature
    r'(\d[\d.,]*)\s+'             # Qty
    r'([\d.,]+)\s+'               # Unit
    r'([\d.,]+)\s*$'              # Total
)
ROW_PROF  = re.compile(
    r'([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

def detect_doc_type(first_page: str) -> str:
    up = first_page.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('No pude determinar tipo (factura / proforma).')

# ───────── endpoint ─────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        request.files['file'].save(tmp.name)

        first_txt = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type  = detect_doc_type(first_txt)

        current_origin  = ''
        current_invbase = ''
        add_plv         = False
        records         = []

        with pdfplumber.open(tmp.name) as pdf:
            for pg in pdf.pages:
                text  = pg.extract_text() or ''
                lines = text.split('\n')
                up    = text.upper()

                # — actualizar invoice number si en esta página hay cabecera —
                head = INV_PAT.search(text)
                if head:
                    current_invbase = head.group(1)
                    add_plv = 'FACTURE SANS PAIEMENT' in up

                # — actualizar origin si la página lo declara —
                o = ORIG_PAT.search(text)
                if o:
                    current_origin = o.group(1).strip()

                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    if doc_type == 'factura':
                        mo = ROW_FACT.match(line)
                        if mo:
                            ref, ean, custom, qty_s, unit_s, tot_s = mo.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            qty  = int(qty_s.replace('.', '').replace(',', ''))
                            rec  = {
                                'Reference'   : ref,
                                'Code EAN'    : ean,
                                'Custom Code' : custom,
                                'Description' : desc,
                                'Origin'      : current_origin,
                                'Quantity'    : qty,
                                'Unit Price'  : fnum(unit_s),
                                'Total Price' : fnum(tot_s),
                                'Invoice Number': (current_invbase + 'PLV') if add_plv else current_invbase
                            }
                            records.append(rec)
                            i += 1          # saltar descripción
                    else:  # proforma
                        mp = ROW_PROF.search(line)
                        if mp:
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.', '').replace(',', ''))
                            rec  = {
                                'Reference'  : ref, 'Code EAN': ean,
                                'Description': desc, 'Origin': current_origin,
                                'Quantity'   : qty, 'Unit Price': fnum(unit_s),
                                'Total Price': fnum(unit_s)*qty
                            }
                            records.append(rec)
                    i += 1

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        headers = (['Reference','Code EAN','Custom Code','Description','Origin',
                    'Quantity','Unit Price','Total Price','Invoice Number']
                   if 'Invoice Number' in records[0]
                   else ['Reference','Code EAN','Description','Origin',
                         'Quantity','Unit Price','Total Price'])

        wb = Workbook(); ws = wb.active; ws.append(headers)
        for r in records: ws.append([r.get(h,'') for h in headers])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error interno:\n{traceback.format_exc()}', 500

