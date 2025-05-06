import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook
from collections import defaultdict

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── REGEX ─────────────────────────────────────────────────
INV_PAT = re.compile(r'(?:FACTURE|INVOICE)[\s\S]{0,1000}?(\d{6,})', re.I)
PLV_PAT = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)

ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+'      # ref
    r'(\d{12,14})\s+'           # ean
    r'(\d{6,9})\s+'             # custom
    r'(\d[\d.,]*)\s+'          # qty
    r'([\d.,]+)\s+'             # unit
    r'([\d.,]+)\s*$'            # total
)
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+'      # ref
    r'(\d{12,14})\s+'           # ean
    r'([\d.,]+)\s+'             # unit
    r'([\d.,]+)\s*$'            # qty
)

# ─── HELPERS ──────────────────────────────────────────────

def fnum(s: str) -> float:
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def doc_kind(txt: str) -> str:
    up = txt.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    return 'factura'

# ─── ENDPOINT ─────────────────────────────────────────────
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

        for pdf_idx, pdf_file in enumerate(pdfs, 1):
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                kind = doc_kind(pdf.pages[0].extract_text() or '')
                inv_global = ''   # persists across pages until changed
                org_global = ''

                for page in pdf.pages:
                    txt = page.extract_text() or ''
                    lines = txt.split('\n')

                    # Detect invoice on this page (can override previous)
                    if (m := INV_PAT.search(txt)):
                        inv_global = m.group(1)
                    plv_page = bool(PLV_PAT.search(txt))
                    invoice_full = inv_global + ('PLV' if plv_page else '')

                    # Running origin: start with previous page value
                    current_org = org_global

                    for i, raw in enumerate(lines):
                        ln = raw.strip()

                        # Update origin when line declares it
                        if (mo := ORG_PAT.search(ln)):
                            val = mo.group(1).strip()
                            if not val and i+1 < len(lines):
                                val = lines[i+1].strip()
                            if val:
                                current_org = val
                                org_global   = val   # carry to next page
                            continue

                        # Match detail rows
                        if kind == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append(dict(zip(COLS,[ref,ean,custom,desc,current_org,
                                int(qty_s.replace('.','').replace(',','')),
                                fnum(unit_s), fnum(tot_s), invoice_full])))
                            continue

                        if kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            rows.append(dict(zip(COLS,[ref,ean,'',desc,current_org,
                                qty, unit, unit*qty, invoice_full])))

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # Rellenar Origin vacío si único por factura
        inv2org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv2org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv2org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv2org[r['Invoice Number']]))

        wb = Workbook(); ws = wb.active; ws.append(COLS)
        for r in rows:
            ws.append([r[c] for c in COLS])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception:
        return f'<pre>{traceback.format_exc()}</pre>', 500
