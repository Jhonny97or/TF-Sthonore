import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 1) PATRONES ─────────────────────────────────────────────
# Busca invoice/facture con dígitos (hasta 800 chars después)
INV_PAT = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,800}?(\d{6,})', re.I | re.S)
# Detecta PLV en el texto
PLV_PAT = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)

# Filas de factura vs proforma
ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

def fnum(s: str) -> float:
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

# Determina si es factura o proforma
def doc_kind(text: str) -> str:
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
        current_org = ''

        for pdf_file in pdfs:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                # Inicial invoice y PLV para todo el PDF
                current_inv = ''
                add_plv = False

                for page in pdf.pages:
                    text = page.extract_text() or ''

                    # Detectar invoice + PLV en todo el texto de la página
                    m = INV_PAT.search(text)
                    if m:
                        current_inv = m.group(1)
                        add_plv = bool(PLV_PAT.search(text))

                    # Nuevos países donde aparezcan
                    for idx, line in enumerate(text.split('\n')):
                        if "PAYS D'ORIGINE" in line.upper():
                            parts = line.split(':', 1)
                            after = parts[1].strip() if len(parts) > 1 else ''
                            if not after and idx + 1 < len(text.split('\n')):
                                after = text.split('\n')[idx+1].strip()
                            if after:
                                current_org = after

                    kind = doc_kind(text)
                    for idx, raw in enumerate(text.split('\n')):
                        line = raw.strip()

                        # Fila FACTURA
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            nxt = text.split('\n')[idx+1] if idx+1 < len(text.split('\n')) else ''
                            if not ROW_FACT.match(nxt):
                                desc = nxt.strip()
                            inv_full = current_inv + ('PLV' if add_plv else '')
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': current_org,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': inv_full
                            })
                            continue

                        # Fila PROFORMA
                        if kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = text.split('\n')[idx+1].strip() if idx+1 < len(text.split('\n')) else ''
                            qty = int(qty_s.replace('.', '').replace(',', ''))
                            unit = fnum(unit_s)
                            inv_full = current_inv + ('PLV' if add_plv else '')
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': desc,
                                'Origin': current_org,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit * qty,
                                'Invoice Number': inv_full
                            })
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ─── 3) Forward-fill de Invoice Number faltantes ───────────
        prev = ''
        for r in rows:
            if r['Invoice Number']:
                prev = r['Invoice Number']
            else:
                r['Invoice Number'] = prev

        # ─── 4) Rellenar Origin si es único ───────────────────────
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_to_org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv_to_org[r['Invoice Number']]))

        # ─── 5) Exportar a Excel ─────────────────────────────────
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
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

