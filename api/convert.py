import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

# ─── Configuración de logging ────────────────────────────────
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── 1) REGEX ────────────────────────────────────────────────
# Nº de factura después de FACTURE/INVOICE (hasta 800 caracteres + saltos de línea)
INV_PAT = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,800}?(\d{6,})', re.I | re.S)

# Indicador de factura sin pago (añade sufijo PLV)
PLV_PAT = re.compile(r'(?:FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT)', re.I)

# Líneas de detalle ‑ FACTURE
ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+'      # Reference
    r'(\d{12,14})\s+'           # EAN
    r'(\d{6,9})\s+'             # Custom Code
    r'(\d[\d.,]*)\s+'          # Qty
    r'([\d.,]+)\s+'             # Unit Price
    r'([\d.,]+)\s*$'            # Total Price
)
# Líneas de detalle ‑ PROFORMA
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+'      # Reference
    r'(\d{12,14})\s+'           # EAN
    r'([\d.,]+)\s+'             # Unit Price
    r'([\d.,]+)\s*$'            # Qty
)

# ─── 2) HELPERS ──────────────────────────────────────────────

def fnum(s: str) -> float:
    """Convierte texto numérico (.,) a float"""
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def doc_kind(text: str) -> str:
    """Determina si el documento es factura o proforma"""
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('Tipo de documento no reconocido.')

# ─── 3) ENDPOINT ─────────────────────────────────────────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert_pdf():
    logging.info('Inicio de conversión')
    try:
        pdfs = request.files.getlist('file')
        if not pdfs:
            return 'No file(s) uploaded', 400

        rows = []
        current_inv_global = ''   # por si una página no trae encabezado
        current_org_global = ''

        for idx_file, pdf_file in enumerate(pdfs, 1):
            logging.info(f'Archivo {idx_file}/{len(pdfs)}: {pdf_file.filename}')
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                kind = doc_kind(pdf.pages[0].extract_text() or '')

                for idx_page, page in enumerate(pdf.pages, 1):
                    text = page.extract_text() or ''
                    lines = text.split('\n')

                    # 3.1 Numero de factura y sufijo PLV para ESTA página
                    m_inv = INV_PAT.search(text)
                    if m_inv:
                        current_inv_global = m_inv.group(1)
                    add_plv_page = bool(PLV_PAT.search(text))
                    invoice_full_page = current_inv_global + ('PLV' if add_plv_page else '')

                    # 3.2 Detectar país (último encontrado en la página)
                    current_org_page = current_org_global  # heredamos
                    for ln in lines:
                        if "PAYS D'ORIGINE" in ln.upper():
                            parts = ln.split(':', 1)
                            org = parts[1].strip() if len(parts) > 1 else ''
                            if org:
                                current_org_page = org
                    current_org_global = current_org_page  # persiste al siguiente bloque

                    # 3.3 Extraer filas
                    for i, raw in enumerate(lines):
                        line = raw.strip()
                        # FACTURA
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i + 1 < len(lines) and not ROW_FACT.match(lines[i + 1]):
                                desc = lines[i + 1].strip()
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': current_org_page,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full_page
                            })
                        # PROFORMA
                        elif kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if i + 1 < len(lines):
                                desc = lines[i + 1].strip()
                            qty = int(qty_s.replace('.', '').replace(',', ''))
                            unit = fnum(unit_s)
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': '',
                                'Description': desc,
                                'Origin': current_org_page,
                                'Quantity': qty,
                                'Unit Price': unit,
                                'Total Price': unit * qty,
                                'Invoice Number': invoice_full_page
                            })
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # 4) Rellenar ORIGIN si está vacío y único por factura
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_to_org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv_to_org[r['Invoice Number']]))

        # 5) Exportar a Excel
        cols = ['Reference', 'Code EAN', 'Custom Code', 'Description',
                'Origin', 'Quantity', 'Unit Price', 'Total Price', 'Invoice Number']
        wb = Workbook(); ws = wb.active; ws.append(cols)
        for r in rows:
            ws.append([r.get(c, '') for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        logging.info('Conversión completada')
        return send_file(buf,
                         as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

