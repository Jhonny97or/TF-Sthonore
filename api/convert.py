import logging
# Suprime los warnings de pdfminer sobre CropBox
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfpage").setLevel(logging.ERROR)

from flask import Flask, request, send_file
from io import BytesIO
import re, warnings, tempfile, traceback
from pdfminer.high_level import extract_text
from openpyxl import Workbook

warnings.filterwarnings('ignore', category=UserWarning)

# Creamos la app WSGI que Vercel va a exponer
app = Flask(__name__)

# Atiende POST en “/” y en “/api/convert”
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        # Validar upload
        if 'file' not in request.files:
            return "No file uploaded", 400
        file = request.files['file']
        doc_type = request.form.get('type', 'auto')

        # Guardar PDF en temporal
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        file.save(tmp.name)
        pdf_path = tmp.name

        # Leer primera página
        first = extract_text(pdf_path, page_numbers=[0]).upper()

        # Detectar tipo si es “auto”
        if doc_type == 'auto':
            if ('ACCUSE' in first and 'RECEPTION' in first) or 'ACKNOWLEDGE' in first:
                doc_type = 'proforma'
            elif 'FACTURE' in first or 'INVOICE' in first:
                doc_type = 'factura'
            else:
                return "No pude determinar tipo", 400

        def num(s):
            s = (s or '').strip()
            return float(s.replace('.','').replace(',','.')) if s else 0.0

        records = []
        # — FACTURA —
        if doc_type == 'factura':
            m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
            base = m.group(1) if m else None
            if not base:
                return "No número de factura", 400

            pat = re.compile(r'^([A-Z]\d{5,7})\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$')
            origin = ''
            full = extract_text(pdf_path)
            for ln in full.split('\n'):
                up = ln.upper()
                if "PAYS D'ORIGINE" in ln:
                    origin = ln.split(':',1)[-1].strip()
                mm = pat.match(ln.strip())
                if mm:
                    ref, ean, custom, qty_s, unit_s, tot_s = mm.groups()
                    inv = base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else base
                    records.append({
                        'Reference': ref,
                        'Code EAN': ean,
                        'Custom Code': custom,
                        'Description': '',
                        'Origin': origin,
                        'Quantity': int(qty_s),
                        'Unit Price': num(unit_s),
                        'Total Price': num(tot_s),
                        'Invoice Number': inv
                    })
            headers = ['Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number']

        # — PROFORMA —
        else:
            pat = re.compile(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')
            full = extract_text(pdf_path)
            lines = full.split('\n')
            i = 0
            while i < len(lines):
                mm = pat.search(lines[i])
                if mm:
                    ref, ean, price_s, qty_s = mm.groups()
                    desc = lines[i+1].strip() if i+1<len(lines) else ''
                    qty = int(qty_s.replace('.','').replace(',',''))
                    unit = num(price_s)
                    records.append({
                        'Reference': ref,
                        'Code EAN': ean,
                        'Description': desc,
                        'Quantity': qty,
                        'Unit Price': unit,
                        'Total Price': unit*qty
                    })
                    i += 2
                else:
                    i += 1
            headers = ['Reference','Code EAN','Description','Quantity','Unit Price','Total Price']

        if not records:
            return "Sin registros extraídos.", 400

        # Generar Excel en memoria
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for r in records:
            ws.append([r[h] for h in headers])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        # Devolver como descarga
        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        tb = traceback.format_exc()
        return f"❌ Error interno en la función:\n{tb}", 500

