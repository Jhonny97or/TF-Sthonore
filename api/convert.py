import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 1) PATRONES ────────────────────────────────────────────
# Captura número de factura y opcionalmente 'PLV' inline
INV_PAT = re.compile(
    r'(?:FACTURE|INVOICE)'                  # palabra clave
    r'(?:\s+(SANS\s+PAIEMENT|WITHOUT\s+PAYMENT))?'  # sufijo PLV inline opcional
    r'[^
\d]{0,800}? '                      # hasta 800 chars no dígito ni newline
    r'(\d{6,})',                           # captura número
    re.I | re.S
)
# Patrón para detectar texto PLV en página
PLV_PAT = re.compile(r'(?:FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT)', re.I)

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

# Conversión de texto a float
def fnum(s: str) -> float:
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

# Detecta tipo de documento
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
        current_org = ''
        current_inv = ''
        add_plv = False

        for pdf_file in pdfs:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                kind = doc_kind(pdf.pages[0].extract_text() or '')

                for page in pdf.pages:
                    text = page.extract_text() or ''
                    lines = text.split('\n')

                    # 1) Detectar invoice y PLV a nivel de página
                    m = INV_PAT.search(text)
                    if m:
                        current_inv = m.group(2)
                        # PLV inline o búsqueda general
                        add_plv = bool(m.group(1)) or bool(PLV_PAT.search(text))

                    # 2) Actualizar país línea a línea
                    for idx, ln in enumerate(lines):
                        if "PAYS D'ORIGINE" in ln.upper():
                            parts = ln.split(':', 1)
                            after = parts[1].strip() if len(parts) > 1 else ''
                            if not after and idx+1 < len(lines):
                                after = lines[idx+1].strip()
                            if after:
                                current_org = after

                    # 3) Extraer filas
                    for idx, raw in enumerate(lines):
                        line = raw.strip()

                        # Factura
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            nxt = lines[idx+1] if idx+1 < len(lines) else ''
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
                        # Proforma
                        elif kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[idx+1].strip() if idx+1 < len(lines) else ''
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
                                'Total Price': unit*qty,
                                'Invoice Number': invoice_full
                            })
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # 4) Forward-fill Invoice Number
        last = ''
        for r in rows:
            if r['Invoice Number']:
                last = r['Invoice Number']
            else:
                r['Invoice Number'] = last

        # 5) Rellenar Origin si es único
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_to_org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv_to_org[r['Invoice Number']]))

        # 6) Exportar a Excel
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

