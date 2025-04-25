import re, warnings, tempfile, base64
import pdfplumber, pandas as pd
warnings.filterwarnings('ignore', category=UserWarning)

def handler(request):
    if request.method != 'POST':
        return {'statusCode':405,'body':'Only POST'}
    # parse multipart
    import cgi
    form = cgi.FieldStorage(fp=request.environ['wsgi.input'], environ=request.environ)
    fileitem = form['file']
    if not fileitem.filename:
        return {'statusCode':400,'body':'No file uploaded'}
    doc_type = form.getvalue('type') or 'auto'

    # save upload to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(fileitem.file.read())
        path = tmp.name

    # extraction logic (igual al tuyo, adaptado)
    records = []
    with pdfplumber.open(path) as pdf:
        first = (pdf.pages[0].extract_text() or '').upper()
        if doc_type == 'auto':
            if ('ACCUSE' in first and 'RECEPTION' in first) or 'ACKNOWLEDGE' in first:
                doc_type = 'proforma'
            elif 'FACTURE' in first or 'INVOICE' in first:
                doc_type = 'factura'
            else:
                return {'statusCode':400,'body':'No pude determinar tipo'}
        def num(s): return float(s.replace('.','').replace(',','.')) if s else 0.0

        if doc_type == 'factura':
            m = re.search(r'(?:FACTURE|INVOICE)[^\d]{0,40}(\d{6,})', first)
            base = m.group(1) if m else None
            if not base: return {'statusCode':400,'body':'No número factura'}
            pat = re.compile(r'^([A-Z]\d{5,7})\s+(\d{13})\s+(\d{6,8})\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)$')
            origin = ''
            for p in pdf.pages:
                txt = p.extract_text() or ''
                up = txt.upper()
                inv = base + 'PLV' if 'FACTURE SANS PAIEMENT' in up else base
                for ln in txt.split('\n'):
                    if "PAYS D'ORIGINE" in ln:
                        origin = ln.split(':',1)[-1].strip()
                    mm = pat.match(ln.strip())
                    if mm:
                        r, e, c, q, u, t = mm.groups()
                        records.append({
                          'Reference':r,'Code EAN':e,'Custom Code':c,
                          'Description':'','Origin':origin,
                          'Quantity':int(q),'Unit Price':num(u),
                          'Total Price':num(t),'Invoice Number':inv
                        })
            cols = ['Reference','Code EAN','Custom Code','Description','Origin','Quantity','Unit Price','Total Price','Invoice Number']
        else:
            pat = re.compile(r'([A-Z]\d{5,7})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)')
            for p in pdf.pages:
                lines = (p.extract_text() or '').split('\n')
                i=0
                while i<len(lines):
                    mm = pat.search(lines[i])
                    if mm:
                        r,e,ps,qs = mm.groups()
                        desc = lines[i+1] if i+1<len(lines) else ''
                        q = int(qs.replace('.','').replace(',',''))
                        u = num(ps)
                        records.append({
                          'Reference':r,'Code EAN':e,
                          'Description':desc,'Quantity':q,
                          'Unit Price':u,'Total Price':u*q
                        })
                        i+=2
                    else: i+=1
            cols = ['Reference','Code EAN','Description','Quantity','Unit Price','Total Price']

    if not records:
        return {'statusCode':400,'body':'Sin registros extraídos.'}

    df = pd.DataFrame(records)[cols]
    from io import BytesIO
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return {
      'statusCode':200,
      'isBase64Encoded': True,
      'headers': {
        'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':'attachment; filename="extracted_data.xlsx"'
      },
      'body': base64.b64encode(buf.getvalue()).decode()
    }
