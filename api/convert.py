"""
Convertidor PDF → Excel (factura / proforma) para Saint-Honoré.
Funciona como función serverless en Vercel: /api/convert (POST multipart/form-data).

FACTURA  ➜ columnas:
  Reference, Code EAN, Custom Code, Description, Origin,
  Quantity, Unit Price, Total Price, Invoice Number

PROFORMA ➜ columnas:
  Reference, Code EAN, Description, Origin,
  Quantity, Unit Price, Total Price
"""

# ─────────────────── Imports ───────────────────
import logging, re, tempfile, traceback
from io import BytesIO

from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)

app = Flask(__name__)

# ─────────────────── Helpers ───────────────────
def fnum(s: str) -> float:
    """'1.234,56' → 1234.56"""
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

def detect_doc_type(first_page: str) -> str:
    up = first_page.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('No pude determinar tipo (factura / proforma).')

# ────────────── Regex globales ──────────────
INV_PAT   = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', re.I)
ORIG_PAT  = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

ROW_FACT  = re.compile(
    r'^([A-Z]\w{3,11})\s+'        # Reference
    r'(\d{12,14})\s+'             # EAN
    r'(\d{6,9})\s+'               # Custom Code
    r'(\d[\d.,]*)\s+'             # Quantity
    r'([\d.,]+)\s+'               # Unit Price
    r'([\d.,]+)\s*$'              # Total
)

ROW_PROF  = re.compile(
    r'([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

# ─────────────────── Endpoint ───────────────────
@app.route('/', methods=['POST'])            # útil para pruebas locales
@app.route('/api/convert', methods=['POST'])  # ruta oficial en Vercel
def convert():
    try:
        # ---------- Validación ----------
        if 'file' not in request.files:
            return 'No file uploaded', 400

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        request.files['file'].save(tmp.name)

        first_txt = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type  = detect_doc_type(first_txt)

        # ---------- Inicializar contexto ----------
        m_inv = INV_PAT.search(first_txt)
        current_invbase = m_inv.group(1) if m_inv else ''
        add_plv         = 'FACTURE SANS PAIEMENT' in first_txt.upper()

        m_org = ORIG_PAT.search(first_txt)
        current_origin = m_org.group(1).strip() if m_org else ''

        records = []

        # ---------- Recorrido de páginas ----------
        with pdfplumber.open(tmp.name) as pdf:
            for pg in pdf.pages:
                text  = pg.extract_text() or ''
                lines = text.split('\n')
                up    = text.upper()

                # cabecera invoice en esta página
                m_new_inv = INV_PAT.search(text)
                if m_new_inv:
                    current_invbase = m_new_inv.group(1)
                    add_plv = 'FACTURE SANS PAIEMENT' in up

                # país de origen en esta página
                m_new_org = ORIG_PAT.search(text)
                if m_new_org:
                    current_origin = m_new_org.group(1).strip()

                i = 0
                while i < len(lines):
                    line = lines[i].strip()

                    # ---------- FACTURA ----------
                    if doc_type == 'factura':
                        mo = ROW_FACT.match(line)
                        if mo:
                            ref, ean, custom, qty_s, unit_s, tot_s = mo.groups()
                            # descripción = línea siguiente si no es otra fila
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            qty = int(qty_s.replace('.', '').replace(',', ''))
                            records.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': current_origin,
                                'Quantity': qty,
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': (current_invbase + 'PLV') if add_plv else current_invbase
                            })
                            i += 1  # saltar descripción

                    # ---------- PROFORMA ----------
                    else:
                        mp = ROW_PROF.search(line)
                        if mp:
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.', '').replace(',', ''))
                            unit = fnum(unit_s)
                            records.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Description': desc,
                                'Origin': current_origin,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit * qty
                            })
                    i += 1

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        # ---------- Exportar a Excel ----------
        headers = (['Reference','Code EAN','Custom Code','Description','Origin',
                    'Quantity','Unit Price','Total Price','Invoice Number']
                   if 'Invoice Number' in records[0]
                   else ['Reference','Code EAN','Description','Origin',
                         'Quantity','Unit Price','Total Price'])

        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for r in records:
            ws.append([r.get(h, '') for h in headers])

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
        return f'❌ Error interno:\n{traceback.format_exc()}', 500

