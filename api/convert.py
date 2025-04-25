import re, warnings, tempfile, base64
from pdfminer.high_level import extract_text
from openpyxl import Workbook

warnings.filterwarnings('ignore', category=UserWarning)

def handler(request):
    if request.method != 'POST':
        return {'statusCode': 405, 'body': 'Only POST allowed'}

    # parse multipart form
    import cgi
    form = cgi.FieldStorage(
        fp=request.environ['wsgi.input'],
        environ=request.environ,
        keep_blank_values=True
    )
    fileitem = form['file']
    if not fileitem.filename:
        return {'statusCode': 400, 'body': 'No file uploaded'}
    doc_type = form.getvalue('type') or 'auto'

    # save uploaded PDF to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(fileitem.file.read())
        pdf_path = tmp.name

    # extract full text of first page
    first = extract_text(pdf_path, page_numbers=[0]).upper()

    # determine type
    if doc_type == 'auto':
        if ('ACCUSE' in first and 'RECEPTION' in first) or 'ACKNOWLEDGE' in first:
            doc_type = 'proforma'
        elif 'FACTURE' in first or 'INVOICE' in first:
            doc_type = 'factura'
        else:
            return {'statusCode':400,'body':'No pude determinar tipo'}

    def num(s):
        s = (s or '').strip()
        return float(s.replace('.','').replace(',','.')) if s else 0.0

    records = []
    if doc_type == 'factura':
        m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
        base = m.group(1) if m else None
        if not base:
            return {'statusCode':400,'body':'No número de factura'}
        line_pat = re.compile(r'^([A-Z]\d{5,7})\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$')
        origin = ''
        # extraer texto completo
        full = extract_text(pdf_path)
        for ln in full.split('\n'):
            up = ln.upper()
            if "PAYS D'ORIGINE" in ln:
                origin = ln.split(':',1)[-1].strip()
            mm = line_pat.match(ln.strip())
            if mm:
                ref, ean, custom, qty_s, unit_s, tot_s = mm.groups()
                inv = base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else base
                records.append({
                    'Reference': ref,
                    'Code EAN': ean,
                    'Custom Code': custom,
                    'Description': '',
                    'Origin': origin,
                    'Quantity': int(qty_s),
                    'Unit Price': num(unit_s),
                    'Total Price': num(tot_s),
                    'Invoice Number': inv
                })
        headers = ['Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number']

    else:  # proforma
        line_pat = re.compile(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')
        full = extract_text(pdf_path)
        lines = full.split('\n')
        i = 0
        while i < len(lines):
            mm = line_pat.search(lines[i])
            if mm:
                ref, ean, price_s, qty_s = mm.groups()
                desc = lines[i+1].strip() if i+1 < len(lines) else ''
                qty = int(qty_s.replace('.','').replace(',',''))
                unit = num(price_s)
                records.append({
                    'Reference': ref,
                    'Code EAN': ean,
                    'Description': desc,
                    'Quantity': qty,
                    'Unit Price': unit,
                    'Total Price': unit * qty
                })
                i += 2
            else:
                i += 1
        headers = ['Reference','Code EAN','Description','Quantity','Unit Price','Total Price']

    if not records:
        return {'statusCode':400,'body':'Sin registros extraídos.'}

    # crear Excel en memoria
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in records:
        ws.append([r[h] for h in headers])

    # volcar a base64
    from io import BytesIO
    buf = BytesIO()
    wb.save(buf)
    return {
        'statusCode': 200,
        'isBase64Encoded': True,
        'headers': {
            'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition':'attachment; filename="extracted_data.xlsx"'
        },
        'body': base64.b64encode(buf.getvalue()).decode()
    }

