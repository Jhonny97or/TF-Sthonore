import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 1) PATRONES ─────────────────────────────────────────────
# Captura INVOICE/FACTURE con sufijo hasta 6 dígitos, permitiendo saltos de línea
INV_PAT = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,800}?(\d{6,})', re.I | re.S)
# Detecta PLV en factura sin pago
PLV_PAT = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)

# Filas de factura vs proforma
ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

def fnum(s: str) -> float:
    """Convierte cadenas numéricas con ',' y '.'"""
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

def doc_kind(text: str) -> str:
    """Determina si es factura o proforma"""
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
        current_inv = ''
        add_plv = False
        current_org = ''

        for pdf_file in pdfs:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            # Abrimos todas las páginas
            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ''

                    # 1️⃣ Detectar invoice + si es PLV en todo el texto de la página
                    if (m := INV_PAT.search(text)):
                        current_inv = m.group(1)
                        add_plv = bool(PLV_PAT.search(text))

                    # 2️⃣ Extraer origen varias veces: la última prevalece
                    for idx, line in enumerate(text.split('\n')):
                        if "PAYS D'ORIGINE" in line.upper():
                            part = line.split(':', 1)
                            after = part[1].strip() if len(part) > 1 else ''
                            if not after and idx + 1 < len(text.split('\n')):
                                after = text.split('\n')[idx + 1].strip()
                            if after:
                                current_org = after

                    # 3️⃣ Extraer filas según el tipo
                    for idx, raw in enumerate(text.split('\n')):
                        line = raw.strip()
                        # FACTURA
                        if doc_kind(text) == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            nxt = text.split('\n')[idx + 1] if idx + 1 < len(text.split('\n')) else ''
                            if not ROW_FACT.match(nxt):
                                desc = nxt.strip()
                            invoice_full = current_inv + ('PLV' if add_plv else '')
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': current_org,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            })

                        # PROFORMA
                        elif doc_kind(text) == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            nxt = text.split('\n')[idx + 1] if idx + 1 < len(text.split('\n')) else ''
                            desc = nxt.strip()
                            qty = int(qty_s.replace('.', '').replace(',', ''))
                            unit = fnum(unit_s)
                            invoice_full = current_inv + ('PLV' if add_plv else '')
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': desc,
                                'Origin': current_org,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit * qty,
                                'Invoice Number': invoice_full
                            })

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # 4️⃣ Rellenar origen si es único por factura
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_to_org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv_to_org[r['Invoice Number']]))

        # 5️⃣ Generar Excel
        cols = ['Reference','Code EAN','Custom Code','Description',
                'Origin','Quantity','Unit Price','Total Price','Invoice Number']
        wb = Workbook()
        ws = wb.active
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c, '') for c in cols])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

