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
INV_PAT   = re.compile(r'(?:FACTURE|INVOICE)[\s\S]{0,1000}?(\d{6,})', re.I)
PLV_PAT   = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT   = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)
ORDER_PAT = re.compile(r'ORDER\s+NUMBER\s*:?\s*(\d{6,})', re.I)  # proforma

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

COLS = ['Reference','Code EAN','Custom Code','Description',
        'Origin','Quantity','Unit Price','Total Price','Invoice Number']

# ─── HELPERS ──────────────────────────────────────────────

def fnum(s: str) -> float:
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def doc_kind(text: str) -> str:
    up = text.upper()
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

        rows = []
        for pdf_file in pdfs:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)
            with pdfplumber.open(tmp.name) as pdf:
                kind = doc_kind(pdf.pages[0].extract_text() or '')
                inv_global = ''
                org_global = ''

                for page in pdf.pages:
                    txt   = page.extract_text() or ''
                    lines = txt.split('\n')

                    # ── número de factura u ORDER NUMBER ─────────────
                    if kind == 'factura':
                        if (m := INV_PAT.search(txt)):
                            inv_global = m.group(1)
                        plv_page = bool(PLV_PAT.search(txt))
                    else:  # proforma
                        if (o := ORDER_PAT.search(txt)):
                            inv_global = o.group(1)
                        plv_page = False  # nunca PLV en proforma
                    invoice_full = inv_global + ('PLV' if plv_page else '')

                    # ── origen (último de la página) ─────────────────
                    cur_org = org_global
                    for ln in lines:
                        if (mo := ORG_PAT.search(ln)):
                            val = mo.group(1).strip()
                            if not val and (idx:=lines.index(ln))+1 < len(lines):
                                val = lines[idx+1].strip()
                            if val:
                                cur_org = val
                    org_global = cur_org

                    # ── extraer filas ───────────────────────────────
                    for i, raw in enumerate(lines):
                        ln = raw.strip()
                        if kind == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append(dict(zip(COLS,[ref,ean,custom,desc,cur_org,
                                int(qty_s.replace('.','').replace(',','')),
                                fnum(unit_s), fnum(tot_s), invoice_full])))
                        elif kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            rows.append(dict(zip(COLS,[ref,ean,'',desc,cur_org,
                                qty, unit, unit*qty, invoice_full])))
            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # Rellenar Origin único por factura/orden
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
