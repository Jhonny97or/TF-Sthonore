import logging, re, tempfile, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ───────── helpers ─────────
def num(s: str) -> float:
    s = (s or '').strip()
    return float(s.replace('.', '').replace(',', '.')) if s else 0.0

def detect_type(text: str) -> str:
    up = text.upper()
    if any(k in up for k in ('ACKNOWLEDGE', 'ACCUSE', 'PROFORMA')):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('No pude determinar tipo.')

ORIGIN_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

ROW_FACT = re.compile(
    r'^([A-Z]\d{4,12})\s+'      # referencia
    r'(\d{12,14})\s+'           # EAN
    r'(\d{6,9})\s+'             # nomenclature
    r'(\d[\d.,]*)\s+'           # qty
    r'([\d.,]+)\s+'             # unit
    r'([\d.,]+)\s*$'            # total
)
ROW_PROF = re.compile(
    r'([A-Z]\d{4,12})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

# ───────── endpoint ─────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        pdf_file = request.files['file']
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        pdf_file.save(tmp.name)

        first_txt = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type  = detect_type(first_txt)

        records = []

        with pdfplumber.open(tmp.name) as pdf:

            # ─── FACTURA ───
            if doc_type == 'factura':
                m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', first_txt, re.I)
                base = m.group(1) if m else re.search(r'\b(\d{8,})\b', first_txt).group(1)
                for pg in pdf.pages:
                    txt   = pg.extract_text() or ''
                    lines = txt.split('\n')
                    up    = txt.upper()
                    inv   = base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else base
                    origin_page = ORIGIN_PAT.search(txt)
                    origin_page = origin_page.group(1).strip() if origin_page else ''

                    i = 0
                    while i < len(lines):
                        mo = ROW_FACT.match(lines[i].strip())
                        if mo:
                            ref, ean, custom, qty_s, unit_s, tot_s = mo.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]) else ''
                            origin_inline = 'China' if 'CHINA' in desc.upper() else \
                                            'Union Européenne/European Union' if 'UNION EUROP' in desc.upper() else ''
                            records.append({
                                'Reference': ref, 'Code EAN': ean, 'Custom Code': custom,
                                'Description': desc,
                                'Origin': origin_inline or origin_page,
                                'Quantity': int(qty_s.replace('.','').replace(',','')),
                                'Unit Price': num(unit_s),
                                'Total Price': num(tot_s),
                                'Invoice Number': inv
                            })
                            i += 1
                        i += 1

            # ─── PROFORMA ───
            else:
                for pg in pdf.pages:
                    txt   = pg.extract_text() or ''
                    lines = txt.split('\n')
                    origin_page = ORIGIN_PAT.search(txt)
                    origin_page = origin_page.group(1).strip() if origin_page else ''

                    for idx, line in enumerate(lines):
                        m = ROW_PROF.search(line)
                        if m:
                            ref, ean, unit_s, qty_s = m.groups()
                            desc = lines[idx+1].strip() if idx+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = num(unit_s)
                            records.append({
                                'Reference': ref, 'Code EAN': ean,
                                'Description': desc, 'Origin': origin_page,
                                'Quantity': qty, 'Unit Price': unit,
                                'Total Price': unit*qty
                            })

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        cols_fact = ['Reference','Code EAN','Custom Code','Description','Origin',
                     'Quantity','Unit Price','Total Price','Invoice Number']
        cols_prof = ['Reference','Code EAN','Description','Origin',
                     'Quantity','Unit Price','Total Price']
        headers = cols_fact if 'Invoice Number' in records[0] else cols_prof

        wb = Workbook()           # ← libro único
        ws = wb.active
        ws.append(headers)
        for r in records:
            ws.append([r.get(h,'') for h in headers])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf, as_attachment=True, download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error interno:\n{traceback.format_exc()}', 500

