import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')
app = Flask(__name__)

# ─── patrones ───────────────────────────────────────────────
INV_PAT = re.compile(r'(?:FACTURE|INVOICE)[\s\S]{0,1000}?(\d{6,})', re.I)
PLV_PAT = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)
ORG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)

ROW_FACT = re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$')
ROW_PROF = re.compile(r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')

def fnum(s): return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0
def kind(t): return 'proforma' if 'PROFORMA' in t.upper() else 'factura'

# ─── endpoint ───────────────────────────────────────────────
@app.route('/', methods=['POST'])
def convert():
    try:
        pdfs = request.files.getlist('file')
        if not pdfs:
            return 'No file(s) uploaded', 400

        cols = ['Reference','Code EAN','Custom Code','Description',
                'Origin','Quantity','Unit Price','Total Price','Invoice Number']
        rows = []

        for pdf_file in pdfs:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                doc_kind = kind(pdf.pages[0].extract_text() or '')
                cur_inv_global, cur_org_global = '', ''

                for page in pdf.pages:
                    text  = page.extract_text() or ''
                    lines = text.split('\n')

                    # ── PASO 1: detectar invoice / PLV / origen ──────
                    cur_inv_page = cur_inv_global
                    if (m := INV_PAT.search(text)):
                        cur_inv_page = m.group(1)
                    add_plv_page = bool(PLV_PAT.search(text))
                    origin_page  = cur_org_global
                    for ln in lines:
                        if (o := ORG_PAT.search(ln)):
                            val = o.group(1).strip() or origin_page
                            origin_page = val
                    cur_inv_global, cur_org_global = cur_inv_page, origin_page
                    invoice_full = cur_inv_page + ('PLV' if add_plv_page else '')

                    # ── PASO 2: extraer filas con datos ya conocidos ──
                    for i, raw in enumerate(lines):
                        ln = raw.strip()
                        if doc_kind == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            rows.append(dict(zip(cols, [
                                ref, ean, custom, desc, origin_page,
                                int(qty_s.replace('.','').replace(',','')),
                                fnum(unit_s), fnum(tot_s), invoice_full ])))
                        elif doc_kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',', ''))
                            unit = fnum(unit_s)
                            rows.append(dict(zip(cols, [
                                ref, ean, '', desc, origin_page,
                                qty, unit, unit*qty, invoice_full ])))

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # rellenar Origin vacío cuando sea único por factura
        from collections import defaultdict
        inv_org = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_org[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_org[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv_org[r['Invoice Number']]))

        wb, ws = Workbook(), Workbook().active
        ws.append(cols)
        for r in rows:
            ws.append([r[c] for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception:
        logging.error(traceback.format_exc())
        return '❌ Error interno: revisa logs.', 500
