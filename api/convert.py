# ─── 1) IMPORTS ────────────────────────────────────────────────────
import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 2) REGEX Y HELPERS ───────────────────────────────────────────
INV_PAT  = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', re.I)
PLV_PAT  = re.compile(r'FACTURE SANS PAIEMENT', re.I)

ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$'
)
ROW_PROF = re.compile(
    r'([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)'
)

def fnum(s: str) -> float:
    s = (s or '').strip().replace('.', '').replace(',', '.')
    return float(s) if s else 0.0

def doc_kind(text: str) -> str:
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE'  in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('Tipo de documento no reconocido.')

# ─── 3) ENDPOINT ──────────────────────────────────────────────────
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
                kind = doc_kind(pdf.pages[0].extract_text() or '')

            current_inv    = ''
            add_plv_flag   = False
            current_origin = ''
            pending_rows   = []

            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    # ── 3.1 Nº de factura en esta página ───────────────
                    if (m_inv := INV_PAT.search(txt)):
                        current_inv  = m_inv.group(1)
                        add_plv_flag = bool(PLV_PAT.search(txt))
                        for r in pending_rows:
                            r['Invoice Number'] = current_inv + ('PLV' if add_plv_flag else '')
                        pending_rows.clear()

                    # ── 3.2 País de origen (manejo línea + siguiente) ──
                    for idx, ln in enumerate(lines):
                        up_ln = ln.upper()
                        if "PAYS D'ORIGINE" in up_ln:
                            after = ln.split(':', 1)[1].strip() if ':' in ln else ''
                            if not after and idx + 1 < len(lines):
                                after = lines[idx + 1].strip()
                            if after:
                                current_origin = after
                            break   # ya lo encontramos en esta página

                    # ── 3.3 Explorar filas de artículos ────────────────
                    i = 0
                    while i < len(lines):
                        line = lines[i].strip()

                        # FACTURA
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i + 1 < len(lines) and not ROW_FACT.match(lines[i + 1]):
                                desc = lines[i + 1].strip()

                            row = {
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : custom,
                                'Description'   : desc,
                                'Origin'        : current_origin,
                                'Quantity'      : int(qty_s.replace('.','').replace(',','')),
                                'Unit Price'    : fnum(unit_s),
                                'Total Price'   : fnum(tot_s),
                                'Invoice Number': current_inv + ('PLV' if add_plv_flag else '') if current_inv else ''
                            }
                            rows.append(row)
                            if not current_inv:
                                pending_rows.append(row)
                            i += 1

                        # PROFORMA
                        elif kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if i + 1 < len(lines):
                                desc = lines[i + 1].strip()
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)

                            row = {
                                'Reference'     : ref,
                                'Code EAN'      : ean,
                                'Custom Code'   : '',
                                'Description'   : desc,
                                'Origin'        : current_origin,
                                'Quantity'      : qty,
                                'Unit Price'    : unit,
                                'Total Price'   : unit*qty,
                                'Invoice Number': current_inv + ('PLV' if add_plv_flag else '') if current_inv else ''
                            }
                            rows.append(row)
                            if not current_inv:
                                pending_rows.append(row)
                        i += 1

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ─── 4) Completar los Origin vacíos si es único por factura ───
        inv_to_origin = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_to_origin[r['Invoice Number']].add(r['Origin'])

        for r in rows:
            if not r['Origin']:
                origins = inv_to_origin.get(r['Invoice Number'], set())
                if len(origins) == 1:
                    r['Origin'] = next(iter(origins))

        # ─── 5) Generar Excel ────────────────────────────────────────
        cols = [
            'Reference','Code EAN','Custom Code','Description',
            'Origin','Quantity','Unit Price','Total Price','Invoice Number'
        ]
        wb = Workbook(); ws = wb.active
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c,'') for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name='extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error:\n{traceback.format_exc()}', 500


