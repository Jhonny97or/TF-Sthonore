# … imports y helpers iguales …

INV_PAT  = re.compile(r'(?:FACTURE|INVOICE)[^\d]{0,60}(\d{6,})', re.I)
ORIG_PAT = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

# ─────── endpoint ───────
@app.route('/', methods=['POST'])
@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        request.files['file'].save(tmp.name)

        first_txt = extract_text(tmp.name, page_numbers=[0]) or ''
        doc_type  = detect_doc_type(first_txt)

        # ── INIT: invoice y origin vistas en la 1ª página ──
        base_match = INV_PAT.search(first_txt)
        current_invbase = base_match.group(1) if base_match else ''
        add_plv         = 'FACTURE SANS PAIEMENT' in first_txt.upper()

        o_match = ORIG_PAT.search(first_txt)
        current_origin = o_match.group(1).strip() if o_match else ''

        records = []

        with pdfplumber.open(tmp.name) as pdf:
            for pg in pdf.pages:
                text  = pg.extract_text() or ''
                lines = text.split('\n')
                up    = text.upper()

                # si esta página declara un NUEVO invoice, actualizamos
                m_inv = INV_PAT.search(text)
                if m_inv:
                    current_invbase = m_inv.group(1)
                    add_plv = 'FACTURE SANS PAIEMENT' in up

                # si esta página declara un NUEVO origin, actualizamos
                m_org = ORIG_PAT.search(text)
                if m_org:
                    current_origin = m_org.group(1).strip()

                i = 0
                while i < len(lines):
                    line = lines[i].strip()

                    if doc_type == 'factura':
                        mo = ROW_FACT.match(line)
                        if mo:
                            ref, ean, custom, qty_s, unit_s, tot_s = mo.groups()
                            desc = ''
                            if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                                desc = lines[i+1].strip()
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            records.append({
                                'Reference': ref, 'Code EAN': ean, 'Custom Code': custom,
                                'Description': desc,
                                'Origin': current_origin,
                                'Quantity': qty,
                                'Unit Price': fnum(unit_s),
                                'Total Price': fnum(tot_s),
                                'Invoice Number': (current_invbase + 'PLV') if add_plv else current_invbase
                            })
                            i += 1
                    else:  # proforma
                        mp = ROW_PROF.search(line)
                        if mp:
                            ref, ean, unit_s, qty_s = mp.groups()
                            desc = lines[i+1].strip() if i+1 < len(lines) else ''
                            qty  = int(qty_s.replace('.','').replace(',',''))
                            records.append({
                                'Reference': ref, 'Code EAN': ean,
                                'Description': desc, 'Origin': current_origin,
                                'Quantity': qty, 'Unit Price': fnum(unit_s),
                                'Total Price': fnum(unit_s)*qty
                            })
                    i += 1

        if not records:
            return 'Sin registros extraídos; revisa el PDF.', 400

        headers = (['Reference','Code EAN','Custom Code','Description','Origin',
                    'Quantity','Unit Price','Total Price','Invoice Number']
                   if 'Invoice Number' in records[0]
                   else ['Reference','Code EAN','Description','Origin',
                         'Quantity','Unit Price','Total Price'])

        wb = Workbook(); ws = wb.active; ws.append(headers)
        for r in records: ws.append([r.get(h,'') for h in headers])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='extracted_data.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception:
        logging.error(traceback.format_exc())
        return f'❌ Error interno:\n{traceback.format_exc()}', 500
