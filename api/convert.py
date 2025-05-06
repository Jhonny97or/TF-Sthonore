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
        # Obtenemos lista de archivos bajo el campo 'file'
        files = request.files.getlist('file')
        if not files:
            return 'No file(s) uploaded', 400

        rows = []
        # Procesamos cada PDF por separado
        for uploaded in files:
            # Guardar temporalmente
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            uploaded.save(tmp.name)

            # Lectura inicial de la primera página
            first = extract_text(tmp.name, page_numbers=[0]) or ''
            kind  = detect_doc_type(first)

            # Valores base por documento
            inv_base = INV_PAT.search(first).group(1) if INV_PAT.search(first) else ''
            add_plv  = 'FACTURE SANS PAIEMENT' in first.upper()
            origin   = ORIG_PAT.search(first).group(1).strip() if ORIG_PAT.search(first) else ''

            # Abrir con pdfplumber y extraer líneas
            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')
                    up    = txt.upper()

                    # Actualizar inv_base / origin si aparecen en páginas posteriores
                    if m := INV_PAT.search(txt):
                        inv_base = m.group(1)
                        add_plv  = 'FACTURE SANS PAIEMENT' in up
                    if o := ORIG_PAT.search(txt):
                        origin = o.group(1).strip()

                    i = 0
                    while i < len(lines):
                        line = lines[i].strip()
                        # Fila de FACTURA
                        if kind == 'factura' and (mo := ROW_FACT.match(line)):
                            ref, ean, custom_code, qty_s, unit_s, tot_s = mo.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom_code,
                                'Description': desc,
                                'Origin': origin,
                                'Quantity': qty,
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': inv_base + ('PLV' if add_plv else '')
                            })
                            i += 1

                        # Fila de PROFORMA
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
                                'Invoice Number': ''
                            })
                        i += 1

            # Borrar temp file
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # Columnas fijas (salen en todas las filas, aunque algunas queden vacías)
        cols = [
            'Reference','Code EAN','Custom Code','Description',
            'Origin','Quantity','Unit Price','Total Price','Invoice Number'
        ]

        # Crear workbook y volcar datos
        wb = Workbook()
        ws = wb.active
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c, '') for c in cols])

        # Enviamos Excel resultante
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500
