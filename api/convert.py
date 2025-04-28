import logging
from flask import Flask, request, send_file
from io import BytesIO
import re, warnings, tempfile, traceback

# Suprime los warnings de pdfminer
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfpage").setLevel(logging.ERROR)
warnings.filterwarnings('ignore', category=UserWarning)

from pdfminer.high_level import extract_text
import pdfplumber
from openpyxl import Workbook

app = Flask(__name__)

@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        # 1) Validación de upload
        if 'file' not in request.files:
            return "No file uploaded", 400
        file = request.files['file']
        doc_type = request.form.get('type', 'auto')

        # 2) Guardar PDF en archivo temporal
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        file.save(tmp.name)
        pdf_path = tmp.name

        # 3) Extraer texto de primera página para detección
        text_first = extract_text(pdf_path, page_numbers=[0]) or ""
        first = text_first.upper()

        # 4) Logs para depurar
        print(">>> DOC_TYPE inicial:", doc_type)
        print(">>> Muestra de primeras 10 líneas:", text_first.split("\n")[:10])

        # 5) Detección automática de tipo
        if doc_type == 'auto':
            if any(k in first for k in ('ACCUSE', 'RECEPTION', 'ACKNOWLEDGE')):
                doc_type = 'proforma'
            elif 'FACTURE' in first or 'INVOICE' in first:
                doc_type = 'factura'
            else:
                return "No pude determinar tipo", 400

        # 6) Helper para convertir string a float
        def num(s):
            return float(s.replace('.', '').replace(',', '.')) if s else 0.0

        records = []
        headers = []

        # 7) Intentar extraer tabla con pdfplumber (solo primera página)
        with pdfplumber.open(pdf_path) as pdf:
            page0 = pdf.pages[0]
            table = page0.extract_table()

        if doc_type == 'factura' and table:
            # Asumimos que table[0] son headers
            raw_headers = table[0]
            for row in table[1:]:
                # Ajusta índices si tu tabla tiene más columnas
                ref, ean, custom, qty_s, unit_s, tot_s = (row + ['']*6)[:6]
                records.append({
                    'Reference': ref.strip(),
                    'Code EAN': ean.strip(),
                    'Custom Code': custom.strip(),
                    'Quantity': int(qty_s.replace('.', '').strip()),
                    'Unit Price': num(unit_s),
                    'Total Price': num(tot_s),
                    'Invoice Number': ''  # pondremos después
                })
            # Extraer número de factura del texto
            m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
            inv_num = m.group(1) if m else ''
            for r in records:
                r['Invoice Number'] = inv_num

            headers = [
                'Reference','Code EAN','Custom Code',
                'Quantity','Unit Price','Total Price','Invoice Number'
            ]

        else:
            # Fallback basado en regex genérica para línea por línea
            pat = re.compile(
                r'^\s*([A-Z0-9]{5,10})\s+'   # Reference
                r'(\d{8,14})\s+'             # EAN
                r'([A-Z0-9\-]+)\s+'          # Custom Code
                r'(\d+)\s+'                  # Quantity
                r'([\d.,]+)\s+'              # Unit Price
                r'([\d.,]+)\s*$'             # Total Price
            )
            full = extract_text(pdf_path) or ""
            for ln in full.split('\n'):
                mm = pat.match(ln)
                if mm:
                    ref, ean, custom, qty_s, unit_s, tot_s = mm.groups()
                    records.append({
                        'Reference': ref,
                        'Code EAN': ean,
                        'Custom Code': custom,
                        'Quantity': int(qty_s),
                        'Unit Price': num(unit_s),
                        'Total Price': num(tot_s),
                        'Invoice Number': ''
                    })
            # Número de factura/proforma
            m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
            inv_num = m.group(1) if m else ''
            for r in records:
                r['Invoice Number'] = inv_num

            headers = [
                'Reference','Code EAN','Custom Code',
                'Quantity','Unit Price','Total Price','Invoice Number'
            ]

        # 8) Si no halló nada, preview para depurar
        if not records:
            preview = "\n".join(full.split('\n')[:100])
            return (
                "Sin registros extraídos.\n"
                "--- Preview primeras 100 líneas del PDF ---\n"
                f"{preview}"
            ), 400

        # 9) Generar Excel en memoria
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for r in records:
            ws.append([r.get(h, '') for h in headers])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        # 10) Enviar archivo al cliente
        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        tb = traceback.format_exc()
        return f"❌ Error interno en la función:\n{tb}", 500

if __name__ == '__main__':
    app.run(debug=True)
