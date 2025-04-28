import logging
import tempfile
import traceback
import re
from io import BytesIO

from flask import Flask, request, send_file
import pdfplumber
from pdfminer.high_level import extract_text
from openpyxl import Workbook

# 1) Suprime warnings de pdfminer
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfpage").setLevel(logging.ERROR)

app = Flask(__name__)

@app.route("/", methods=["POST"])
def convert():
    try:
        # — Validar que venga el PDF —
        if 'file' not in request.files:
            return "No file uploaded", 400
        f = request.files['file']
        doc_type = request.form.get('type', 'auto')

        # — Guardar PDF en tmp —
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        f.save(tmp.name)
        pdf_path = tmp.name

        # — Leer primera página para detección —
        text_first = extract_text(pdf_path, page_numbers=[0]) or ""
        first = text_first.upper()

        # — Auto-detectar tipo —
        if doc_type == 'auto':
            if any(k in first for k in ('ACCUSE','RECEPTION','ACKNOWLEDGE')):
                doc_type = 'proforma'
            elif 'FACTURE' in first or 'INVOICE' in first:
                doc_type = 'factura'
            else:
                return "No pude determinar tipo", 400

        # — Helper para parsear números —
        def num(s):
            return float(s.replace('.', '').replace(',', '.')) if s else 0.0

        records = []

        # — Primero, intento con pdfplumber —
        with pdfplumber.open(pdf_path) as pdf:
            table = pdf.pages[0].extract_table()

        if table and doc_type == 'factura':
            # Asumimos que fila 0 es header y el resto son datos
            for row in table[1:]:
                ref, ean, custom, qty_s, unit_s, tot_s = (row + ['']*6)[:6]
                records.append({
                    'Reference': ref.strip(),
                    'Code EAN': ean.strip(),
                    'Custom Code': custom.strip(),
                    'Quantity': int(qty_s.replace('.','').strip()),
                    'Unit Price': num(unit_s),
                    'Total Price': num(tot_s),
                    'Invoice Number': ''
                })
            # Extraer número de factura
            m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
            inv = m.group(1) if m else ''
            for r in records:
                r['Invoice Number'] = inv

        else:
            # — Fallback: regex línea a línea —
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
            # Extraer número de factura/proforma
            m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
            inv = m.group(1) if m else ''
            for r in records:
                r['Invoice Number'] = inv

        # — Si no hay registros, devolver preview —
        if not records:
            preview = "\n".join(full.split('\n')[:100])
            return (
                "Sin registros extraídos.\n"
                "--- Preview primeras 100 líneas del PDF ---\n" +
                preview
            ), 400

        # — Generar Excel en memoria —
        headers = list(records[0].keys())
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for r in records:
            ws.append([r[h] for h in headers])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        # — Enviar al cliente —
        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception:
        tb = traceback.format_exc()
        return f"❌ Error interno:\n{tb}", 500
