# app.py  ── listo para Vercel o ejecución local
import logging
import os
import re
import tempfile
import traceback
from collections import defaultdict
from io import BytesIO
from typing import Dict, List

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

# ──────────────────────────────  CONFIG GLOBAL  ─────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

COLS = [
    "Reference", "Code EAN", "Custom Code", "Description",
    "Origin", "Quantity", "Unit Price", "Total Price", "Invoice Number"
]

# ───────────────────────  PATRONES PARA CÓDIGOS  ────────────────────────────
HTS_PAT = re.compile(r"^\d{6,10}$")
UPC_PAT = re.compile(r"^\d{11,14}$")

# ─────────────────────  EXTRACTOR 1  (facturas clásicas)  ───────────────────
INV_PAT      = re.compile(r"(?:FACTURE|INVOICE)\D*(\d{6,})", re.I)
PROF_PAT     = re.compile(r"PROFORMA[\s\S]*?(\d{6,})", re.I)
ORDER_PAT_EN = re.compile(r"ORDER\s+NUMBER\D*(\d{6,})", re.I)
ORDER_PAT_FR = re.compile(r"N°\s*DE\s*COMMANDE\D*(\d{6,})", re.I)
PLV_PAT      = re.compile(r"FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT", re.I)
ORG_PAT      = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

ROW_FACT = re.compile(
    r"^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$"
)
ROW_PROF_DIOR = re.compile(
    r"^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,10})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)\s*$"
)
ROW_PROF = re.compile(
    r"^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)\s*$"
)

def fnum(s: str) -> float:
    return float(s.strip().replace(".", "").replace(",", ".")) if s and s.strip() else 0.0

def doc_kind(text: str) -> str:
    up = text.upper()
    return "proforma" if "PROFORMA" in up or ("ACKNOWLEDGE" in up and "RECEPTION" in up) else "factura"

def extract_original(pdf_path: str) -> List[dict]:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        all_txt = "\n".join(page.extract_text() or "" for page in pdf.pages)
        kind = doc_kind(all_txt)

        inv_global = ""
        plv_global = False
        if kind == "factura":
            if m := INV_PAT.search(all_txt):
                inv_global = m.group(1)
            if PLV_PAT.search(all_txt):
                plv_global = True
        else:
            if m := PROF_PAT.search(all_txt):
                inv_global = m.group(1)
            elif m := ORDER_PAT_EN.search(all_txt):
                inv_global = m.group(1)
            elif m := ORDER_PAT_FR.search(all_txt):
                inv_global = m.group(1)

        invoice_full = inv_global + ("PLV" if plv_global else "")
        org_global = ""

        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            # país de origen
            for ln in lines:
                if mo := ORG_PAT.search(ln):
                    val = mo.group(1).strip()
                    if val:
                        org_global = val

            for i, raw in enumerate(lines):
                ln = raw.strip()
                if kind == "factura" and (mf := ROW_FACT.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]) else ""
                    rows.append({
                        "Reference": ref,
                        "Code EAN": ean,
                        "Custom Code": custom,
                        "Description": desc,
                        "Origin": org_global,
                        "Quantity": int(qty_s.replace(".", "").replace(",", "")),
                        "Unit Price": fnum(unit_s),
                        "Total Price": fnum(tot_s),
                        "Invoice Number": invoice_full,
                    })
                elif kind == "proforma" and (mpd := ROW_PROF_DIOR.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = mpd.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append({
                        "Reference": ref,
                        "Code EAN": ean,
                        "Custom Code": custom,
                        "Description": desc,
                        "Origin": org_global,
                        "Quantity": int(qty_s.replace(".", "").replace(",", "")),
                        "Unit Price": fnum(unit_s),
                        "Total Price": fnum(tot_s),
                        "Invoice Number": invoice_full,
                    })
                elif kind == "proforma" and (mp := ROW_PROF.match(ln)):
                    ref, ean, unit_s, qty_s = mp.groups()
                    qty = int(qty_s.replace(".", "").replace(",", ""))
                    unit = fnum(unit_s)
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append({
                        "Reference": ref,
                        "Code EAN": ean,
                        "Custom Code": "",
                        "Description": desc,
                        "Origin": org_global,
                        "Quantity": qty,
                        "Unit Price": unit,
                        "Total Price": unit * qty,
                        "Invoice Number": invoice_full,
                    })

    # completar Origin si hay uno solo por invoice
    inv2org = defaultdict(set)
    for r in rows:
        if r["Origin"]:
            inv2org[r["Invoice Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2org[r["Invoice Number"]]) == 1:
            r["Origin"] = next(iter(inv2org[r["Invoice Number"]]))
    return rows

# ─────────────────────  EXTRACTOR 2  (por coordenadas)  ──────────────────────
COL_BOUNDS = {
    "ref":   (0,   70),
    "desc":  (70, 340),
    "upc":   (340,430),
    "ctry":  (430,465),
    "hs":    (465,535),
    "qty":   (535,585),
    "unit":  (585,635),
    "total": (635,725),
}
REF_PAT = re.compile(r"^\d{5,6}[A-Z]?$")
NUM_PAT = re.compile(r"[0-9]")
SKIP_SNIPPETS = {
    "No. Description","Total before","Bill To Ship","CIF CHILE",
    "Invoice","Ship From","Ship To","VAT/Tax","Shipping Te"
}

def clean(txt: str) -> str:
    return txt.replace("\u202f"," ").strip()

def to_float2(txt: str) -> float:
    t = txt.replace("\u202f","").replace(" ","")
    if t.count(",")==1 and t.count(".")==0:
        t = t.replace(",",".")
    elif t.count(".")>1:
        t = t.replace(".","")
    return float(t or 0)

def to_int2(txt: str) -> int:
    return int(txt.replace(",","").replace(".","") or 0)

def rows_from_page(page) -> List[Dict[str,str]]:
    rows=[]
    grouped={}
    for ch in page.chars:
        grouped.setdefault(round(ch["top"],1),[]).append(ch)
    for _,chs in sorted(grouped.items()):
        line_txt="".join(c["text"] for c in sorted(chs,key=lambda c:c["x0"]))
        if not line_txt.strip() or any(sn in line_txt for sn in SKIP_SNIPPETS):
            continue
        cols={k:"" for k in COL_BOUNDS}
        for c in sorted(chs,key=lambda c:c["x0"]):
            xm=(c["x0"]+c["x1"])/2
            for key,(x0,x1) in COL_BOUNDS.items():
                if x0<=xm<x1:
                    cols[key]+=c["text"]
                    break
        cols={k:clean(v) for k,v in cols.items()}
        if not cols["ref"]:
            if rows: rows[-1]["desc"]+=" "+cols["desc"]
            continue
        if not REF_PAT.match(cols["ref"]) or not NUM_PAT.search(cols["qty"]):
            continue
        rows.append(cols)
    return rows

def extract_slice(pdf_path: str, inv_number: str) -> List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for r in rows_from_page(page):
                rows.append({
                    "Reference": r["ref"],
                    "Code EAN": r["upc"],
                    "Custom Code": r["hs"],
                    "Description": r["desc"],
                    "Origin": r["ctry"],
                    "Quantity": to_int2(r["qty"]),
                    "Unit Price": to_float2(r["unit"]),
                    "Total Price": to_float2(r["total"]),
                    "Invoice Number": inv_number
                })
    return rows

# ─────────────────────  EXTRACTOR 3  (proveedor nuevo)  ──────────────────────
pattern_full = re.compile(r"""
    ^\s*
    (?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+
    Each\s+
    (?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+
    (?P<total>[\d.,]+)
    """, re.VERBOSE)

pattern_nohs = re.compile(r"""
    ^\s*
    (?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<qty>[\d,]+)\s+
    Each\s+
    (?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+
    (?P<total>[\d.,]+)
    """, re.VERBOSE)

pattern_basic = re.compile(r"""
    ^\s*
    (?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+
    Each\s+
    (?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+
    (?P<total>[\d.,]+)
    """, re.VERBOSE)

def extract_new_provider(pdf_path: str, inv_number: str) -> List[dict]:
    def new_fnum(s: str) -> float:
        return float(s.replace(",", "")) if s.strip() else 0.0

    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if "No. Description" not in txt:
                continue
            pending_desc=None
            for ln in txt.split("\n"):
                ln_s=ln.strip()
                if not ln_s or ln_s.startswith(("No. Description","Invoice")):
                    continue

                # 1) línea completa con HS
                if m := pattern_full.match(ln):
                    d=m.groupdict()
                    rows.append({
                        "Reference": d["ref"],
                        "Code EAN": d["upc"],
                        "Custom Code": d["hs"],
                        "Description": d["desc"].strip(),
                        "Origin": d["ctry"],
                        "Quantity": int(d["qty"].replace(",","")),
                        "Unit Price": new_fnum(d["unit"]),
                        "Total Price": new_fnum(d["total"]),
                        "Invoice Number": inv_number
                    })
                    pending_desc=None
                    continue

                # 2) línea completa sin HS
                if m2 := pattern_nohs.match(ln):
                    d=m2.groupdict()
                    rows.append({
                        "Reference": d["ref"],
                        "Code EAN": d["upc"],
                        "Custom Code": "",
                        "Description": d["desc"].strip(),
                        "Origin": d["ctry"],
                        "Quantity": int(d["qty"].replace(",","")),
                        "Unit Price": new_fnum(d["unit"]),
                        "Total Price": new_fnum(d["total"]),
                        "Invoice Number": inv_number
                    })
                    pending_desc=None
                    continue

                # 3) línea básica (solo números tras desc previa)
                if mb := pattern_basic.match(ln):
                    if pending_desc:
                        d=mb.groupdict()
                        rows.append({
                            "Reference": d["ref"],
                            "Code EAN": d["upc"],
                            "Custom Code": d["hs"],
                            "Description": pending_desc.strip(),
                            "Origin": d["ctry"],
                            "Quantity": int(d["qty"].replace(",","")),
                            "Unit Price": new_fnum(d["unit"]),
                            "Total Price": new_fnum(d["total"]),
                            "Invoice Number": inv_number
                        })
                        pending_desc=None
                    continue

                # 4) acumular descripción multi-línea
                if re.search(r"[A-Za-z]", ln_s):
                    skip_pref=("Country of","Customer PO","Order No",
                               "Shipping Terms","Bill To","Finance",
                               "Total","CIF","Ship To")
                    if not any(ln_s.startswith(p) for p in skip_pref):
                        pending_desc=(pending_desc+" "+ln_s) if pending_desc else ln_s
    return rows

# ──────────────────  EXTRACTOR 4 v3 (Interparfums por ventana)  ─────────────
HEAD_CAND_PAT = re.compile(
    r"^(?P<ref>(?:[A-Z]{3}\d{3,6}[A-Z]?|\d{4,6}[A-Z]?))\s+(?P<desc>.+)$"
)
HS_ORG_PAT = re.compile(r"HS\s*Code:\s*(?P<hs>\d{8,14})\s*,\s*Origin:\s*(?P<org>[A-Z]{2})", re.I)
EAN_PAT = re.compile(r"EAN\s*Code:\s*(?P<ean>\d{12,14})", re.I)
TOTAL_LINE_PAT = re.compile(
    r"""^
    (?P<qty>[\d\.]+)\s+PZ\s+
    (?P<unit>[\d\.,]+)\s+
    (?P<gross>[\d\.,]+)
    (?:\s+(?P<disc>-?\d+%)\s+(?P<net>[\d\.,]+))?
    \s+(?P<vat>[A-Z]{2})
    $""", re.X | re.I
)

BAD_HEAD_PREFIXES = (
    "Invoice", "Document Date", "Invoice No.", "Location Page", "DESTINATION",
    "Sell-to", "Ship-to", "Registered office", "Administrative headquarters",
    "Share Capital", "Company belonging", "CIF", "Shipment Method",
    "Your Reference", "Payment Bank", "Bonifico", "SWIFT", "Due Date",
    "TOTAL GROSS AMOUNT", "VAT BASE", "NET TO PAY", "Customs Code",
    "VALUES FOR CUSTOMS", "Preferential Origin", "No Preferential Origin",
    "INVOICING"
)

def extract_interparfums_blocks(pdf_path: str, invoice_number: str) -> List[dict]:
    rows: List[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            # normaliza pequeñas dobles-espacios
            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

            # recorre y cuando encuentre TOTALES, sube a buscar EAN/HS/HEAD
            for i, ln in enumerate(lines):
                mt = TOTAL_LINE_PAT.match(ln)
                if not mt:
                    continue

                # 1) totales
                qty  = int(mt.group("qty").replace(".","").replace(",",""))
                unit = fnum(mt.group("unit"))
                gross = fnum(mt.group("gross"))
                total = fnum(mt.group("net")) if mt.group("net") else gross

                # 2) ventana hacia arriba (5-8 líneas suelen bastar)
                win_start = max(0, i-8)
                window = lines[win_start:i]

                ref = desc = ean = hs = org = ""

                # Busca EAN y HS/Origin primero (pueden venir en cualquier orden)
                for w in reversed(window):
                    if not ean:
                        me = EAN_PAT.search(w)
                        if me: ean = me.group("ean")
                    if not hs or not org:
                        mh = HS_ORG_PAT.search(w)
                        if mh:
                            hs = mh.group("hs")
                            org = mh.group("org")
                    if ean and hs and org:
                        break

                # Busca cabecera válida inmediatamente por encima
                for w in reversed(window):
                    # evita líneas de metadatos
                    if any(w.startswith(p) for p in BAD_HEAD_PREFIXES):
                        continue
                    mh = HEAD_CAND_PAT.match(w)
                    if mh:
                        ref  = mh.group("ref")
                        desc = mh.group("desc").strip()
                        break

                # Si no encontró cabecera, intenta heurística: toma el primer token de la línea previa
                if not ref and window:
                    tokens = window[-1].split()
                    if tokens:
                        tok = tokens[0]
                        if re.fullmatch(r"(?:[A-Z]{3}\d{3,6}[A-Z]?|\d{4,6}[A-Z]?)", tok):
                            ref = tok
                            desc = window[-1][len(tok):].strip()

                # Si aún no hay ref/desc, omite (ruido)
                if not ref or not desc:
                    continue

                rows.append({
                    "Reference": ref,
                    "Code EAN": ean,
                    "Custom Code": hs,
                    "Description": desc,
                    "Origin": org,
                    "Quantity": qty,
                    "Unit Price": unit,
                    "Total Price": total,
                    "Invoice Number": invoice_number
                })

    return rows


# ───────────  COMPLEMENTO: detectar Invoice No. dentro del PDF  ─────────────
INVNO_PAT = re.compile(r"Invoice\s+No\.\s*([A-Z0-9\-\/]+)", re.I)

def parse_invoice_number_from_pdf(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if m := INVNO_PAT.search(txt):
                return m.group(1).strip()
    return ""

# ────────────────  COMPLEMENTO: llenar HTS / UPC faltantes  ────────────────
def complete_missing_codes(pdf_path: str, rows: List[dict]) -> None:
    """Rellena in-place cualquier fila sin HTS o UPC."""
    lines=[]
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            txt=pg.extract_text(x_tolerance=1.5) or ""
            lines.extend(txt.split("\n"))
    lines=[re.sub(r"\s{2,}"," ",ln.strip()) for ln in lines if ln.strip()]

    # mapa Reference → índice aproximado
    ref_idx={}
    for idx,ln in enumerate(lines):
        m=re.match(r"^([A-Z0-9]{3,})\s+[A-Z]{3}\s",ln)
        if m:
            ref_idx.setdefault(m.group(1), idx)

    for r in rows:
        if r["Custom Code"] and r["Code EAN"]:
            continue
        start=ref_idx.get(r["Reference"])
        if start is None:
            continue
        end=start+1
        while end<len(lines) and end-start<20:
            if re.match(r"^[A-Z0-9]{3,}\s+[A-Z]{3}\s",lines[end]):
                break
            end+=1
        snippet=" ".join(lines[start:end])
        seqs=re.findall(r"\d{6,14}", snippet)
        hts=[s for s in seqs if HTS_PAT.match(s)]
        upc=[s for s in seqs if UPC_PAT.match(s)]
        if hts and not r["Custom Code"]:
            r["Custom Code"]=hts[0]
        if upc and not r["Code EAN"]:
            r["Code EAN"]=upc[0]

# ─────────────────────────────  ENDPOINT  ────────────────────────────────────
@app.post("/api/convert")
@app.post("/")
def convert():
    try:
        pdfs=request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded",400

        all_rows=[]
        for pdf in pdfs:
            with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
                pdf.save(tmp.name)

                # 1) intenta desde el nombre (SIP…), 2) si no, desde el PDF (Invoice No.)
                inv_num=(m.group(1) if (m:=re.search(r"SIP(\d+)", pdf.filename or "")) else "")
                if not inv_num:
                    inv_num = parse_invoice_number_from_pdf(tmp.name)

                # 1-4) extraemos con cada estrategia
                rows1=extract_original(tmp.name)
                rows2=extract_slice(tmp.name,inv_num)
                rows3=extract_new_provider(tmp.name,inv_num)
                rows4=extract_interparfums_blocks(tmp.name,inv_num)

                combo=rows1+rows2+rows3+rows4

                # eliminar duplicados por (Reference, EAN, Invoice)
                seen=set(); uniq=[]
                for r in combo:
                    key=(r["Reference"], r["Code EAN"], r["Invoice Number"])
                    if key not in seen:
                        seen.add(key); uniq.append(r)

                # rellenar cualquier HTS / UPC faltante
                complete_missing_codes(tmp.name, uniq)

                all_rows.extend(uniq)
            os.unlink(tmp.name)

        if not all_rows:
            return "Sin registros extraídos",400

        wb=Workbook(); ws=wb.active; ws.append(COLS)
        for r in all_rows:
            ws.append([r.get(c, "") for c in COLS])
        buf=BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>",500

if __name__=="__main__":
    app.run(debug=True,host="0.0.0.0")



