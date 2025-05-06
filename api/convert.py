# ─── 1) IMPORTS ────────────────────────────────────────────────────
import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook

logging.getLogger("pdfminer").setLevel(logging.ERROR)
app = Flask(__name__)

# ─── 2) REGEX Y HELPERS ───────────────────────────────────────────
INV_PAT  = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', re.I)
ORIG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)
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
            # Guardar temporalmente
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            # ===== PASADA 1: construir índice {factura: {origin, plv}} =====
            meta = {}          # p.e. {'102555272': {'origin':'Union …', 'plv':True}}
            cur_inv = ''

            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ''

                    if m := INV_PAT.search(txt):
                        cur_inv = m.group(1)
                        meta.setdefault(cur_inv, {'origin':'', 'plv':False})

                    if cur_inv:
                        if ORIG_PAT.search(txt):
                            val = ORIG_PAT.search(txt).group(1).strip()
                            if val:
                                meta[cur_inv]['origin'] = val
                        if PLV_PAT.search(txt):
                            meta[cur_inv]['plv'] = True

            # ===== PASADA 2: extraer filas con datos completos ============
            kind = doc_kind(pdfplumber.open(tmp.name).pages[0].extract_text())

            with pdfplumber.open(tmp.name) as pdf:
                cur_inv = ''
                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    # Actualizar número de factura / flag PLV al entrar a página
                    if m := INV_PAT.search(txt):
                        cur_inv = m.group(1)

                    # Si aún no hemos visto inv, toma el primero del índice
                    if not cur_inv and meta:
                        cur_inv = next(iter(meta))

                    inv_info = meta.get(cur_inv, {'origin':'','plv':False})
                    origin   = inv_info['origin']
                    add_plv  = inv_info['plv']

                    i = 0
                    while i < len(lines):
                        line = lines[i].strip()

                        # ── Factura ───────────────────────────────
                        if kind == 'factura' and (mf := ROW_FACT.match(line)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            # descripción = línea siguiente si no es otra fila
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append({
                                'Reference': ref,
                                'Code EAN': ean,
                                'Custom Code': custom,
                                'Description': desc,
                                'Origin': origin,
                                'Quantity': int(qty_s.replace('.','').replace(',','')),
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': cur_inv + ('PLV' if add_plv else '')
                            })
                            i += 1

                        # ── Proforma ──────────────────────────────
                        elif kind == 'proforma' and (mp := ROW_PROF.match(line)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if i+1 < len(lines):
                                desc = lines[i+1].strip()
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
                                'Total Price': unit*qty,
                                'Invoice Number': cur_inv + ('PLV' if add_plv else '')
                            })
                        i += 1

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ===== Generar Excel ============================================
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

