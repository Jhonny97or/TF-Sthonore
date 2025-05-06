import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 1) PATRONES ────────────────────────────────────────────
# Captura FACTURE/INVOICE con posible PLV y número
INV_PAT = re.compile(
    r'(?:FACTURE|INVOICE)'                   # palabra clave
    r'(?:\s+(SANS\s+PAIEMENT|WITHOUT\s+PAYMENT))?'  # grupo opcional PLV
    r'[^\d]{0,800}?'                         # hasta 800 chars no dígito
    r'(\d{6,})',                            # captura nº de invoice
    re.I | re.S
)
# Patrón para detectar PLV si no aparece junto al invoice
PLV_INLINE = re.compile(r'(SANS\s+PAIEMENT|WITHOUT\s+PAYMENT)', re.I)

# Patrones de filas
ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+'      # Reference
    r'(\d{12,14})\s+'           # EAN
    r'(\d{6,9})\s+'             # Custom Code
    r'(\d[\d.,]*)\s+'          # Quantity
    r'([\d.,]+)\s+'             # Unit Price
    r'([\d.,]+)\s*$'            # Total Price
)
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+'      # Reference
    r'(\d{12,14})\s+'           # EAN
    r'([\d.,]+)\s+'             # Unit Price
    r'([\d.,]+)'                 # Quantity
)

def fnum(s: str) -> float:
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

# Determina tipo de documento (usado solo al inicio)
def doc_kind(text: str) -> str:
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('Tipo de documento no reconocido.')

# ─── 2) ENDPOINT ────────────────────────────────────────────
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
                # Determinamos tipo a partir de primera página
                first_text = pdf.pages[0].extract_text() or ''
                kind = doc_kind(first_text)

                # Estado por PDF
                current_inv = ''
                add_plv = False
                current_org = ''
                pending_rows = []

                # Recorremos todo el PDF línea a línea
                for page in pdf.pages:
                    text = page.extract_text() or ''
                    lines = text.split('\n')

                    for idx, raw in enumerate(lines):
                        line = raw.strip()

                        # 1) Detección de invoice + PLV
                        if m := INV_PAT.search(line):
                            # grupo2 = número, grupo1 = PLV si existía
                            new_inv = m.group(2)
                            new_plv = bool(m.group(1))
                            # si no venía inline, buscamos cerca
                            if not new_plv and PLV_INLINE.search(line) is None:
                                # chequeamos siguiente línea
                                if idx+1 < len(lines) and PLV_INLINE.search(lines[idx+1]):
                                    new_plv = True
                            current_inv = new_inv
                            add_plv = new_plv
                            # completamos las filas pendientes
                            for r in pending_rows:
                                r['Invoice Number'] = current_inv + ('PLV' if add_plv else '')
                            pending_rows.clear()
                            continue

                        # 2) Detección de país
                        if "PAYS D'ORIGINE" in line.upper():
                            part = line.split(':',1)
                            after = part[1].strip() if len(part)>1 else ''
                            if not after and idx+1 < len(lines):
                                after = lines[idx+1].strip()
                            if after:
                                current_org = after
                            continue

                        # 3) Filas FACTURA
                        if kind=='factura' and (mf:=ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            # posible descripción en la siguiente línea
                            if idx+1 < len(lines) and not ROW_FACT.match(lines[idx+1]):
                                desc = lines[idx+1].strip()
                            invoice_full = current_inv + ('PLV' if add_plv else '') if current_inv else ''
                            row = {
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': current_org,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            }
                            rows.append(row)
                            # si aún no detectamos invoice, lo dejamos pendiente
                            if not current_inv:
                                pending_rows.append(row)
                            continue

                        # 4) Filas PROFORMA
                        if kind=='proforma' and (mp:=ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if idx+1 < len(lines):
                                desc = lines[idx+1].strip()
                            qty = int(qty_s.replace('.', '').replace(',', ''))
                            unit = fnum(unit_s)
                            invoice_full = current_inv + ('PLV' if add_plv else '') if current_inv else ''
                            row = {
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': desc,
                                'Origin': current_org,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit*qty,
                                'Invoice Number': invoice_full
                            }
                            rows.append(row)
                            if not current_inv:
                                pending_rows.append(row)

            os.unlink(tmp.name)

        # 5) Forward-fill invoice faltantes
        prev_inv = ''
        for r in rows:
            if r['Invoice Number']:
                prev_inv = r['Invoice Number']
            else:
                r['Invoice Number'] = prev_inv

        # 6) Completar Origin si es único por factura
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_to_org[r['Invoice Number']])==1:
                r['Origin']=next(iter(inv_to_org[r['Invoice Number']]))

        # 7) Generar Excel
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

