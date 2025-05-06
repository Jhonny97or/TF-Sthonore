import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

# Configuración de logging para ver avances en consola
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── 1) PATRONES ────────────────────────────────────────────
# Captura número de factura con posible sufijo PLV inline, permite salto de línea
INV_PAT = re.compile(
    r'(?:FACTURE|INVOICE)'                  # palabra clave
    r'(?:\s+(?:SANS\s+PAIEMENT|WITHOUT\s+PAYMENT))?'  # sufijo PLV inline opcional
    r'[^
\d]{0,800}?'+                     # hasta 800 chars no dígito ni newline
    r'(\d{6,})',                           # captura número
    re.I | re.S
)
# Patrón para detectar PLV en el texto de página
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

# Determina tipo de documento a partir de un texto
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
    logging.info('Inicio de conversión de PDF a Excel')
    try:
        pdfs = request.files.getlist('file')
        if not pdfs:
            logging.warning('No se recibieron archivos')
            return 'No file(s) uploaded', 400

        rows = []

        for idx_file, pdf_file in enumerate(pdfs, start=1):
            logging.info(f'Procesando archivo {idx_file}/{len(pdfs)}: {pdf_file.filename}')
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                kind = doc_kind(pdf.pages[0].extract_text() or '')
                logging.info(f'Tipo de documento: {kind}')

                # Estado por PDF
                current_inv = ''
                add_plv = False
                current_org = ''
                pending = []

                for idx_page, page in enumerate(pdf.pages, start=1):
                    logging.info(f'  Página {idx_page}/{len(pdf.pages)}')
                    text = page.extract_text() or ''
                    lines = text.split('\n')

                    # 1) Detectar invoice y PLV en la página completa
                    m = INV_PAT.search(text)
                    if m:
                        current_inv = m.group(1)
                        add_plv = bool(PLV_PAT.search(text))
                        logging.info(f'    Invoice detectado: {current_inv} (PLV={add_plv})')
                        # Rellenar pendientes
                        for r in pending:
                            r['Invoice Number'] = current_inv + ('PLV' if add_plv else '')
                        pending.clear()

                    # 2) Actualizar país por línea
                    for ln in lines:
                        if "PAYS D'ORIGINE" in ln.upper():
                            part = ln.split(':',1)
                            org = part[1].strip() if len(part)>1 else ''
                            if org:
                                current_org = org
                                logging.info(f'    Origin detectado: {current_org}')

                    # 3) Extraer filas
                    for i, raw in enumerate(lines):
                        line = raw.strip()
                        # Factura
                        if kind=='factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = lines[i+1].strip() if i+1<len(lines) and not ROW_FACT.match(lines[i+1]) else ''
                            invoice_full = current_inv + ('PLV' if add_plv else '')
                            row = {
                                'Reference': ref,'Code EAN': ean,'Custom Code': custom,
                                'Description': desc,'Origin': current_org,
                                'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                'Unit Price': fnum(unit_s),'Total Price': fnum(tot_s),
                                'Invoice Number': invoice_full
                            }
                            rows.append(row)
                            if not current_inv:
                                pending.append(row)
                        # Proforma
                        elif kind=='proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1<len(lines) else ''
                            row = {'Reference': ref,'Code EAN': ean,'Custom Code':'',
                                   'Description': desc,'Origin': current_org,
                                   'Quantity': int(qty_s.replace('.', '').replace(',', '')),
                                   'Unit Price': fnum(unit_s),'Total Price': fnum(unit_s)*int(qty_s.replace('.', '').replace(',', '')),
                                   'Invoice Number': current_inv + ('PLV' if add_plv else '')}
                            rows.append(row)
                            if not current_inv:
                                pending.append(row)
            os.unlink(tmp.name)

        if not rows:
            logging.warning('No se extrajeron filas')
            return 'Sin registros extraídos', 400

        # 4) Forward-fill Invoice Number
        last = ''
        for r in rows:
            if r['Invoice Number']:
                last = r['Invoice Number']
            else:
                r['Invoice Number'] = last

        # 5) Completar Origin si único
        inv_to_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_to_org[r['Invoice Number']])==1:
                r['Origin']=next(iter(inv_to_org[r['Invoice Number']]))

        # 6) Generar Excel
        logging.info('Generando archivo Excel')
        cols=['Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number']
        wb=Workbook(); ws=wb.active; ws.append(cols)
        for r in rows:
            ws.append([r.get(c,'') for c in cols])

        buf=BytesIO(); wb.save(buf); buf.seek(0)
        logging.info('Conversión completada correctamente')
        return send_file(buf,as_attachment=True,download_name='extracted_data.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500

