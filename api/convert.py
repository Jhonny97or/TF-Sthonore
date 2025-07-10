# app.py  – MISMO COMPORTAMIENTO QUE TENÍAS EN COLAB + EXTRACTORES EXTRA

import logging, os, re, tempfile, traceback
from collections import defaultdict
from io import BytesIO
from typing import Dict, List

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

COLS = [
    "Reference", "Code EAN", "Custom Code", "Description",
    "Origin", "Quantity", "Unit Price", "Total Price", "Invoice Number"
]

# ──────────────  VALIDADORES DE CÓDIGOS  ─────────────────────────────────────
HTS_PAT = re.compile(r"^\d{6,10}$")
UPC_PAT = re.compile(r"^\d{11,14}$")

# ─────────────────────  EXTRACTOR 1 (facturas clásicas)  ────────────────────
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
    return float(s.replace(".", "").replace(",", ".")) if s.strip() else 0.0

def doc_kind(text: str) -> str:
    up = text.upper()
    return "proforma" if "PROFORMA" in up or ("ACKNOWLEDGE" in up and "RECEPTION" in up) else "factura"

def extract_original(pdf_path: str) -> List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        full_txt="\n".join(p.extract_text() or "" for p in pdf.pages)
        kind=doc_kind(full_txt)

        inv=""; plv=False
        if kind=="factura":
            if m:=INV_PAT.search(full_txt): inv=m.group(1)
            if PLV_PAT.search(full_txt):    plv=True
        else:
            for pat in (PROF_PAT, ORDER_PAT_EN, ORDER_PAT_FR):
                if m:=pat.search(full_txt): inv=m.group(1); break
        invoice=inv+("PLV" if plv else "")

        org=""
        for pg in pdf.pages:
            lines=(pg.extract_text() or "").split("\n")
            for ln in lines:
                if mo:=ORG_PAT.search(ln): org=mo.group(1).strip() or org
            for i,ln in enumerate(lines):
                ln_s=ln.strip()
                if kind=="factura" and (m:=ROW_FACT.match(ln_s)):
                    r,e,h,q,u,t=m.groups(); desc=lines[i+1].strip() if i+1<len(lines) else ""
                    rows.append({"Reference":r,"Code EAN":e,"Custom Code":h,"Description":desc,
                                 "Origin":org,"Quantity":int(q.replace(",","").replace(".","")),
                                 "Unit Price":fnum(u),"Total Price":fnum(t),
                                 "Invoice Number":invoice})
                elif kind=="proforma" and (m:=ROW_PROF_DIOR.match(ln_s)):
                    r,e,h,q,u,t=m.groups(); desc=lines[i+1].strip() if i+1<len(lines) else ""
                    rows.append({"Reference":r,"Code EAN":e,"Custom Code":h,"Description":desc,
                                 "Origin":org,"Quantity":int(q.replace(",","").replace(".","")),
                                 "Unit Price":fnum(u),"Total Price":fnum(t),
                                 "Invoice Number":invoice})
                elif kind=="proforma" and (m:=ROW_PROF.match(ln_s)):
                    r,e,u,q=m.groups(); desc=lines[i+1].strip() if i+1<len(lines) else ""
                    qty=int(q.replace(",","").replace(".","")); unit=fnum(u)
                    rows.append({"Reference":r,"Code EAN":e,"Custom Code":"",
                                 "Description":desc,"Origin":org,"Quantity":qty,
                                 "Unit Price":unit,"Total Price":unit*qty,
                                 "Invoice Number":invoice})
    inv2org=defaultdict(set)
    for r in rows:
        if r["Origin"]: inv2org[r["Invoice Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2org[r["Invoice Number"]])==1:
            r["Origin"]=next(iter(inv2org[r["Invoice Number"]]))
    return rows

# ──────────────────── EXTRACTOR 2 (tu script de coordenadas) ────────────────
COL_BOUNDS: Dict[str, tuple] = {
    "ref": (0,70),"desc":(70,340),"upc":(340,430),"ctry":(430,465),
    "hs": (465,535),"qty":(535,585),"unit":(585,635),"total":(635,725)
}
REF_PAT  = re.compile(r"^[A-Z0-9]{3,}$")
NUM_PAT  = re.compile(r"[0-9]")
SKIP_SNIPPETS = {
    "No. Description","Total before","Bill To Ship","CIF CHILE",
    "Invoice","Ship From","Ship To","VAT/Tax","Shipping Te"
}

def clean(t:str)->str: return t.replace("\u202f"," ").strip()
def to_float2(t:str)->float:
    t=t.replace("\u202f","").replace(" ","")
    if t.count(",")==1 and t.count(".")==0: t=t.replace(",",".")
    elif t.count(".")>1: t=t.replace(".","")
    return float(t or 0)
def to_int2(t:str)->int: return int(t.replace(",","").replace(".","") or 0)

def rows_from_page(page)->List[Dict[str,str]]:
    rows=[]; grouped={}
    for ch in page.chars: grouped.setdefault(round(ch["top"],1),[]).append(ch)
    for _,chs in sorted(grouped.items()):
        txt="".join(c["text"] for c in sorted(chs,key=lambda c:c["x0"]))
        if not txt.strip() or any(s in txt for s in SKIP_SNIPPETS): continue
        cols={k:"" for k in COL_BOUNDS}
        for c in sorted(chs,key=lambda c:c["x0"]):
            xm=(c["x0"]+c["x1"])/2
            for k,(x0,x1) in COL_BOUNDS.items():
                if x0<=xm<x1: cols[k]+=c["text"]; break
        cols={k:clean(v) for k,v in cols.items()}

        if not cols["ref"]:
            if rows: rows[-1]["desc"]+=" "+cols["desc"]
            continue
        if not REF_PAT.match(cols["ref"]) or not UPC_PAT.match(cols["upc"]):
            pass  # mantenemos la fila: se completará después si falta el UPC/HS
        if not NUM_PAT.search(cols["qty"]):
            continue
        rows.append(cols)
    return rows

def extract_slice(pdf, inv_number)->List[dict]:
    out=[]
    with pdfplumber.open(pdf) as pdfdoc:
        for pg in pdfdoc.pages:
            for r in rows_from_page(pg):
                out.append({
                    "Reference": r["ref"], "Code EAN": r["upc"],
                    "Custom Code": r["hs"], "Description": r["desc"],
                    "Origin": r["ctry"], "Quantity": to_int2(r["qty"]),
                    "Unit Price": to_float2(r["unit"]),
                    "Total Price": to_float2(r["total"]),
                    "Invoice Number": inv_number
                })
    return out

# ───────────────── EXTRACTOR 3 (proveedor nuevo) ────────────────────────────
# ... (sin cambios, igual que antes: pattern_full, pattern_nohs, pattern_basic)
# (por brevedad el bloque se omite aquí, copia exactamente tu versión que funciona)

# ───────  COMPLETAR CÓDIGOS FALTANTES (sin tocar tu lógica) ────────────────
def complete_missing(pdf_path:str, rows:List[dict])->None:
    lines=[]
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            txt=pg.extract_text(x_tolerance=1.5) or ""
            lines.extend(txt.split("\n"))
    lines=[re.sub(r"\s{2,}"," ",ln.strip()) for ln in lines if ln.strip()]
    idx_map={}
    for i,ln in enumerate(lines):
        if m:=re.match(r"^([A-Z0-9]{3,})\s+[A-Z]{3}\s",ln):
            idx_map.setdefault(m.group(1),i)
    for r in rows:
        if r["Custom Code"] and r["Code EAN"]: continue
        start=idx_map.get(r["Reference"]); end=start
        if start is None: continue
        while end<len(lines) and end-start<20 and not re.match(r"^[A-Z0-9]{3,}\s+[A-Z]{3}\s",lines[end]): end+=1
        snippet=" ".join(lines[start:end])
        seqs=re.findall(r"\d{6,14}",snippet)
        hts=[s for s in seqs if HTS_PAT.fullmatch(s)]
        upc=[s for s in seqs if UPC_PAT.fullmatch(s)]
        if hts and not r["Custom Code"]: r["Custom Code"]=hts[0]
        if upc and not r["Code EAN"]:    r["Code EAN"]=upc[0]

# ─────────────────────────  ENDPOINT  ───────────────────────────────────────
@app.post("/api/convert")
@app.post("/")
def convert():
    try:
        pdfs=request.files.getlist("file")
        if not pdfs: return "No file(s) uploaded",400
        all_rows=[]
        for pdf in pdfs:
            with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
                pdf.save(tmp.name)
                inv=(m.group(1) if (m:=re.search(r"SIP(\d+)",pdf.filename or "")) else "")
                r1=extract_original(tmp.name)
                r2=extract_slice(tmp.name,inv)
                r3=extract_new_provider(tmp.name,inv)  # usa tu versión existente
                combo=r1+r2+r3
                seen=set(); uniq=[]
                for r in combo:
                    k=(r["Reference"],r["Code EAN"],r["Invoice Number"])
                    if k not in seen: seen.add(k); uniq.append(r)
                complete_missing(tmp.name, uniq)
                all_rows.extend(uniq)
            os.unlink(tmp.name)
        if not all_rows: return "Sin registros extraídos",400
        wb=Workbook(); ws=wb.active; ws.append(COLS)
        for r in all_rows: ws.append([r[c] for c in COLS])
        buf=BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf,as_attachment=True,
                         download_name="extracted_data.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>",500

if __name__=="__main__":
    app.run(debug=True,host="0.0.0.0")

