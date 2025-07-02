import logging
import re
import tempfile
import os
import traceback
from io import BytesIO

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook
from collections import defaultdict

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── CONSTANTES Y REGEX ───────────────────────────────────────────
INV_PAT       = re.compile(r'(?:FACTURE|INVOICE)\D*(\d{6,})', re.I)
PROF_PAT      = re.compile(r'PROFORMA[\s\S]*?(\d{6,})', re.I)
ORDER_PAT_EN  = re.compile(r'ORDER\s+NUMBER\D*(\d{6,})', re.I)
ORDER_PAT_FR  = re.compile(r'N°\s*DE\s*COMMANDE\D*(\d{6,})', re.I)
PLV_PAT       = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT       = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)
HEADER_PAT    = re.compile(r'^No\.\s+Description', re.I)
SUMMARY_PAT   = re.compile(r'^\s*Total\s+before\s+discount', re.I)
ROW_START_PAT = re.compile(r'^\d{5,6}\s')       # Inicio de renglón real

ROW_FULL = re.compile(
    r'^(?P<ref>\d{5,6})\s+'
    r'(?P<desc>.+?)\s+'
    r'(?P<upc>\d{12,14})\s+'
    r'(?P<ctry>[A-Z]{2})\s+'
    r'(?P<hs>\d{4}\.\d{2}\.\d{4})\s+'
    r'(?P<qty>[\d.,]+)\s+'
    r'(?P<unit>[\d.]+)\s+'
    r'(?:-|[\d.,]+)\s+'
    r'(?P<total>[\d.,]+)$'
)

COLS = [
    'Reference','Code EAN','Custom Code','Description',
    'Origin','Quantity','Unit Price','Total Price','Invoice Number'
]

# ─── UTILIDAD ─────────────────────────────────────────────────────

def fnum(s: str) -> float:
    """Convierte una cadena numérica (US/EU) a float."""
    s = s.replace('\u202f', '').strip()
    if not s:
        return 0.0
    # si contiene ambos separadores
    if ',' in s and '.' in s:
        return float(s.replace(',', '')) if s.find(',') < s.find('.') else float(s.replace('.', '').replace(',', '.'))
    # solo comas (formato EU)
    if ',' in s:
        return float(s.replace('.', '').replace(',', '.'))
    return float(s.replace(',', ''))


def parse_qty(q: str) -> int:
    """Elimina separadores de miles en cantidades."""
    return int(q.replace(',', '').replace('.', ''))


def doc_kind(text: str) -> str:
    up = text.upper()
    return 'proforma' if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up) else 'factura'

# ─── PARSER PRINCIPAL ─────────────────────────────────────────────


def process_chunk(raw: str, origin: str, invoice_no: str, out_rows: list):
    """Intenta convertir un bloque de texto en fila y añadirlo a out_rows."""
    clean = ' '.join(raw.split())             # colapsa espacios y newlines
    clean = clean.replace(' Each ', ' ')      # quita palabra suelta
    m = ROW_FULL.match(clean)
    if not m:
        logging.debug('No match para fila: %s', clean)
        return
    gd = m.groupdict()
    out_rows.append({
        'Reference': gd['ref'],
        'Code EAN': gd['upc'],
        'Custom Code': gd['hs'],
        'Description': gd['desc'],
        'Origin': gd['ctry'] or origin,
        'Quantity': parse_qty(gd['qty']),
        'Unit Price': fnum(gd['unit']),
        'Total Price': fnum(gd['total']),
        'Invoice Number': invoice_no
    })

# ─── FLASK ENDPOINT ───────────────────────────────────────────────

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
                full_txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)
                kind = doc_kind(full_txt)

                # ── Invoice number global ───────────────────────
                inv_no = ''
                if kind == 'factura':
                    if m := INV_PAT.search(full_txt):
                        inv_no = m.group(1)
                else:
                    m = PROF_PAT.search(full_txt) or ORDER_PAT_EN.search(full_txt) or ORDER_PAT_FR.search(full_txt)
                    inv_no = m.group(1) if m else ''
                invoice_full = inv_no + ('PLV' if PLV_PAT.search(full_txt) else '')

                origin_global = ''
                stop_all = False

                # ── Recorremos páginas ─────────────────────────
                for page in pdf.pages:
                    if stop_all:
                        break
                    lines = (page.extract_text() or '').split('\n')

                    # actualizar origen global
                    for ln in lines:
                        if mo := ORG_PAT.search(ln):
                            origin_global = mo.group(1).strip() or origin_global

                    state = 'idle'
                    chunk_lines = []
                    for ln in lines:
                        ln_strip = ln.strip()
                        if SUMMARY_PAT.search(ln_strip):
                            # llegó a resumen / fin de ítems
                            if state == 'building':
                                process_chunk(' '.join(chunk_lines), origin_global, invoice_full, rows)
                            stop_all = True
                            break
                        # saltar cabecera
                        if HEADER_PAT.match(ln_strip):
                            continue

                        if state == 'idle':
                            if ROW_START_PAT.match(ln_strip):
                                state = 'building'
                                chunk_lines = [ln_strip]
                        else:  # building
                            if ROW_START_PAT.match(ln_strip):
                                # comienza nueva fila -> procesar anterior
                                process_chunk(' '.join(chunk_lines), origin_global, invoice_full, rows)
                                chunk_lines = [ln_strip]
                            else:
                                chunk_lines.append(ln_strip)
                    # fin for líneas página
                    if state == 'building' and chunk_lines:
                        process_chunk(' '.join(chunk_lines), origin_global, invoice_full, rows)
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ── Generar Excel ───────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in rows:
            ws.append([r[c] for c in COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf,
                         as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.exception('Error en /convert')
        return f'<pre>{traceback.format_exc()}</pre>', 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')


