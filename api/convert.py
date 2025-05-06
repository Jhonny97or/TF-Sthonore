import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── REGEX ───────────────────────────────────────────────────
# • Invoice number up to 1000 chars after FACTURE/INVOICE (including new‑lines)
INV_PAT = re.compile(r'(?:FACTURE|INVOICE)[\s\S]{0,1000}?(\d{6,})', re.I)
# • "Invoice without payment" → add suffix PLV
PLV_PAT = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
# • Country of origin line
ORG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)

# Detail rows (invoice)
ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+'          # ref
    r'(\d{12,14})\s+'               # EAN
    r'(\d{6,9})\s+'                 # custom
    r'(\d[\d.,]*)\s+'              # qty
    r'([\d.,]+)\s+'                 # unit
    r'([\d.,]+)\s*$'                # total
)
# Detail rows (proforma)
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+'          # ref
    r'(\d{12,14})\s+'               # EAN
    r'([\d.,]+)\s+'                 # unit
    r'([\d.,]+)\s*$'                # qty
)

# ─── HELPERS ────────────────────────────────────────────────

def fnum(s: str) -> float:
    """Convert "1.234,56" or "1,234.56" to float"""
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def doc_kind(text: str) -> str:
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    return 'factura'

# ─── MAIN ENDPOINT ──────────────────────────────────────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        pdfs = request.files.getlist('file')
        if not pdfs:
            return 'No file(s) uploaded', 400

        COLS = ['Reference','Code EAN','Custom Code','Description',
                'Origin','Quantity','Unit Price','Total Price','Invoice Number']
        rows = []

        for pdf_index, pdf_file in enumerate(pdfs, 1):
            logging.info(f'Processing PDF {pdf_index}/{len(pdfs)}: {pdf_file.filename}')
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                doc_type = doc_kind(pdf.pages[0].extract_text() or '')
                inv_global, org_global = '', ''

                for page_index, page in enumerate(pdf.pages, 1):
                    text  = page.extract_text() or ''
                    lines = text.split('\n')

                    # PASS 1 ─ detect invoice / PLV / origin for the whole page
                    inv_page = inv_global
                    if (m := INV_PAT.search(text)):
                        inv_page = m.group(1)
                    plv_page = bool(PLV_PAT.search(text))
                    org_page = org_global
                    for ln in lines:
                        if (mo := ORG_PAT.search(ln)):
                            val = mo.group(1).strip()
                            if val:
                                org_page = val
                    # persist for next pages
                    inv_global, org_global = inv_page, org_page
                    invoice_full = inv_page + ('PLV' if plv_page else '')

                    # PASS 2 ─ extract rows
                    for i, raw in enumerate(lines):
                        ln = raw.strip()
                        if doc_type == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append(dict(zip(COLS,[ref,ean,custom,desc,org_page,
                                int(qty_s.replace('.','').replace(',','')),
                                fnum(unit_s), fnum(tot_s), invoice_full])))
                        elif doc_type == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1<len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            rows.append(dict(zip(COLS,[ref,ean,'',desc,org_page,
                                qty, unit, unit*qty, invoice_full])))
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # Rellenar Origin vacío si es único por factura
        from collections import defaultdict
        inv_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_org[r['Invoice Number']])==1:
                r['Origin']=next(iter(inv_org[r['Invoice Number']]))

        # Export to Excel
        wb, ws = Workbook(), Workbook().active
        ws.append(COLS)
        for r in rows:
            ws.append([r[c] for c in COLS])
        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        # Devuelve la traza para depuración rápida
        return f'<pre>{traceback.format_exc()}</pre>', 500
