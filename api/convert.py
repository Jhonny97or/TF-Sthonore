import logging, re, tempfile, os, traceback
from io import BytesIO
from flask import Flask, request, send_file
import pdfplumber
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')

app = Flask(__name__)

# ─── REGEX ────────────────────────────────────────────────────
INV_PAT  = re.compile(r'(?:FACTURE|INVOICE)[\s\S]{0,1000}?(\d{6,})', re.I)
PLV_PAT  = re.compile(r'FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT', re.I)

ROW_FACT = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$')
ROW_PROF = re.compile(
    r'^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')

ORG_PAT  = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.*)", re.I)

# ─── HELPERS ─────────────────────────────────────────────────
def fnum(s: str) -> float:
    return float(s.strip().replace('.', '').replace(',', '.')) if s.strip() else 0.0

def kind(text: str) -> str:
    up = text.upper()
    if 'PROFORMA' in up or ('ACKNOWLEDGE' in up and 'RECEPTION' in up):
        return 'proforma'
    if 'FACTURE' in up or 'INVOICE' in up:
        return 'factura'
    raise ValueError('Tipo de documento no reconocido.')

# ─── MAIN ENDPOINT ───────────────────────────────────────────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        pdfs = request.files.getlist('file')
        if not pdfs:
            return 'No file(s) uploaded', 400

        cols  = ['Reference','Code EAN','Custom Code','Description',
                 'Origin','Quantity','Unit Price','Total Price','Invoice Number']
        rows  = []

        for f_idx, pdf_file in enumerate(pdfs, 1):
            logging.info(f'PDF {f_idx}/{len(pdfs)}  →  {pdf_file.filename}')
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            pdf_file.save(tmp.name)

            with pdfplumber.open(tmp.name) as pdf:
                doc_kind = kind(pdf.pages[0].extract_text() or '')

                # Current state
                cur_inv  = ''
                cur_plv  = False
                cur_org  = ''
                pending  = []  # rows without invoice yet

                for p_idx, page in enumerate(pdf.pages, 1):
                    lines = (page.extract_text() or '').split('\n')

                    for i, raw in enumerate(lines):
                        ln = raw.strip()

                        # 1️⃣  Header lines
                        if (m := INV_PAT.search(ln)):
                            cur_inv = m.group(1)
                            cur_plv = bool(PLV_PAT.search(ln))
                            # Maybe PLV is on the same page but not same line
                            if not cur_plv and any(PLV_PAT.search(x) for x in lines[i:i+3]):
                                cur_plv = True
                            for r in pending:
                                r['Invoice Number'] = cur_inv + ('PLV' if cur_plv else '')
                            pending.clear()
                            continue

                        # 2️⃣  Origin line
                        if (m := ORG_PAT.search(ln)):
                            found = m.group(1).strip()
                            if not found and i+1 < len(lines):
                                found = lines[i+1].strip()
                            if found:
                                cur_org = found
                            continue

                        # 3️⃣  Detail rows
                        if doc_kind == 'factura' and (mf := ROW_FACT.match(ln)):
                            ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            row = dict(zip(cols, [
                                ref, ean, custom, desc, cur_org,
                                int(qty_s.replace('.','').replace(',','')),
                                fnum(unit_s), fnum(tot_s),
                                cur_inv + ('PLV' if cur_plv else '')
                            ]))
                            rows.append(row)
                            if not cur_inv:
                                pending.append(row)
                            continue

                        if doc_kind == 'proforma' and (mp := ROW_PROF.match(ln)):
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = ''
                            if i+1 < len(lines):
                                desc = lines[i+1].strip()
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            unit = fnum(unit_s)
                            row  = dict(zip(cols, [
                                ref, ean, '', desc, cur_org,
                                qty, unit, unit*qty,
                                cur_inv + ('PLV' if cur_plv else '')
                            ]))
                            rows.append(row)
                            if not cur_inv:
                                pending.append(row)

            os.unlink(tmp.name)

        if not rows:
            return 'Sin registros extraídos', 400

        # ── Completar ORIGIN cuando sea único por factura ──────
        from collections import defaultdict
        inv_orgs = defaultdict(set)
        for r in rows:
            if r['Origin']:
                inv_orgs[r['Invoice Number']].add(r['Origin'])
        for r in rows:
            if not r['Origin'] and len(inv_orgs[r['Invoice Number']]) == 1:
                r['Origin'] = next(iter(inv_orgs[r['Invoice Number']]))

        # ── Generar Excel ──────────────────────────────────────
        wb = Workbook(); ws = wb.active; ws.append(cols)
        for r in rows:
            ws.append([r[c] for c in cols])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return '❌ Error interno – revisa logs.', 500


