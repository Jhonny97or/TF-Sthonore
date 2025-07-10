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

# ────────────────────────────── CONFIG GLOBAL ───────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

COLS = [
    "Reference", "Code EAN", "Custom Code", "Description",
    "Origin", "Quantity", "Unit Price", "Total Price", "Invoice Number"
]

# ─────────── Patrones genéricos para validar códigos numéricos ──────────────
HTS_PAT = re.compile(r"^\d{6,10}$")   # 6-10 dígitos
UPC_PAT = re.compile(r"^\d{11,14}$")  # 11-14 dígitos

# ───────────────────────── EXTRACTOR 1 (facturas) ───────────────────────────
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
    return float(s.strip().replace(".", "").replace(",", ".")) if s.strip() else 0.0

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
            if m := INV_PAT.search(all_txt): inv_global = m.group(1)
            if PLV_PAT.search(all_txt):      plv_global = True
        else:
            if m := PROF_PAT.search(all_txt):            inv_global = m.group(1)
            elif m := ORDER_PAT_EN.search(all_txt):      inv_global = m.group(1)
            elif m := ORDER_PAT_FR.search(all_txt):      inv_global = m.group(1)

        invoice_full = inv_global + ("PLV" if plv_global else "")
        org_global = ""

        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            for ln in lines:  # origen
                if mo := ORG_PAT.search(ln):
                    org_global = mo.group(1).strip() or org_global

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
                    qty  = int(qty_s.replace(".", "").replace(",", ""))
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

    inv2org = defaultdict(set)
    for r in rows:
        if r["Origin"]:
            inv2org[r["Invoice Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2org[r["Invoice Number"]]) == 1:
            r["Origin"] = next(iter(inv2org[r["Invoice Number"]]))
    return rows

# ─────────────────────  EXTRACTOR 2 (coordenadas)  ───────────────────────────
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
REF_PAT  = re.compile(r"^\d{5,6}[A-Z]?$")
NUM_PAT  = re.compile(r"[0-9]")
SKIP_SN  = {
    "No. Description","Total before","Bill To Ship","CIF CHILE",
    "Invoice","Ship From","Ship To","VAT/Tax","Shipping Te"
}

def clean(txt: str) -> str:
    return txt.replace("\u202f"," ").strip()

def to_float2(txt: str) -> float:
    t = txt.replace("\u202f","").replace(" ","")
    if t.count(",")==1 and t.count(".")==0: t = t.replace(",",".")
    elif t.count(".")>1:                    t = t.replace(".","")
    return float(t or 0)

def to_int2(txt: str) -> int:
    return int(txt.replace(",","").replace(".","") or 0)

def rows_from_page(page) -> List[Dict[str,str]]:
    rows=[]
    grouped={}
    for ch in page.chars:
        grouped.setdefault(round(ch["top"],1),[]).append(ch)

    for _,chs in sorted(grouped.items()):
        line="".join(c["text"] for c in sorted(chs,key=lambda c:c["x0"]))
        if not line.strip() or any(sn in line for sn in SKIP_SN):
            continue

        cols={k:"" for k in COL_BOUNDS}
        for c in sorted(chs,key=lambda c:c["x0"]):
            xm=(c["x0"]+c["x1"])/2
            for key,(x0,x1) in COL_BOUNDS.items():
                if x0<=xm<x1:
                    cols[key]+=c["text"]
                    break
        cols={k:clean(v) for k,v in cols.items()}

        # ─── Validaciones corregidas ───────────────────────────────────────
        if not REF_PAT.match(cols["ref"]):
            continue                      # referencia malformada

        # Si UPC / HS no son válidos, los vaciamos (luego se completan)
        if not UPC_PAT.match(cols["upc"]): cols["upc"] = ""
        if not HTS_PAT.match(cols["hs"]):  cols["hs"]  = ""

        if not NUM_PAT.search(cols["qty"]):
            continue                      # sin cantidad → no es fila de ítem

        if not cols["ref"]:
            if rows: rows[-1]["desc"] += " " + cols["desc"]
            continue
        rows.append(cols)
    return rows

def extract_slice(pdf_path: str, inv_number: str) -> List[dict]:
    out=[]
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            for r in rows_from_page(pg):
                out.append({
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
    return out

# ─────────────────────  EXTRACTOR 3 (proveedor nuevo)  ───────────────────────
pattern_full = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+(?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+(?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
    """, re.VERBOSE)

pattern_nohs = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+(?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+(?P<ctry>[A-Z]{2})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
    """, re.VERBOSE)

pattern_basic = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+(?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+(?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
    """, re.VERBOSE)

def extract_new_provider(pdf_path: str, inv_number: str) -> List[dict]:
    def fn(s: str) -> float: return float(s.replace(",", "")) if s.strip() else 0.0
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            txt = pg.extract_text() or ""
            if "No. Description" not in txt: continue
            pending=None
            for ln in txt.split("\n"):
                ls = ln.strip()
                if not ls or ls.startswith(("No. Description","Invoice")): continue

                if m := pattern_full.match(ln):
                    d=m.groupdict(); rows.append({
                        "Reference": d["ref"], "Code EAN": d["upc"],
                        "Custom Code": d["hs"], "Description": d["desc"].strip(),
                        "Origin": d["ctry"], "Quantity": int(d["qty"].replace(",","")),
                        "Unit Price": fn(d["unit"]),
                        "Total Price": fn(d["total"]),
                        "Invoice Number": inv_number
                    }); pending=None; continue

                if m2 := pattern_nohs.match(ln):
                    d=m2.groupdict(); rows.append({
                        "Reference": d["ref"], "Code EAN": d["upc"],
                        "Custom Code": "", "Description": d["desc"].strip(),
                        "Origin": d["ctry"], "Quantity": int(d["qty"].replace(",","")),
                        "Unit Price": fn(d["unit"]),
                        "Total Price": fn(d["total"]),
                        "Invoice Number": inv_number
                    }); pending=None; continue

                if mb := pattern_basic.match(ln):
                    if pending:
                        d=mb.groupdict(); rows.append({
                            "Reference": d["ref"], "Code EAN": d["upc"],
                            "Custom Code": d["hs"], "Description": pending.strip(),
                            "Origin": d["ctry"], "Quantity": int(d["qty"].replace(",","")),
                            "Unit Price": fn(d["unit"]),
                            "Total Price": fn(d["total"]),
                            "Invoice Number": inv_number
                        }); pending=None; continue

                if re.search(r"[A-Za-z]", ls):   # desc multi-línea
                    if not any(ls.startswith(p) for p in ("Country of","Customer PO",
                                                          "Order No","Shipping Terms",
                                                          "Bill To","Finance","Total",
                                                          "CIF","Ship To")):
                        pending = (pending+" "+ls) if pending else ls
    return rows

# ───────  COMPLEMENTO: rellena HTS / UPC faltantes con búsqueda libre ───────
def complete_missing_codes(pdf_path: str, rows: List[dict]) -> None:
    lines=[]
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            txt=pg.extract_text(x_tolerance=1.5) or ""
            lines.extend(txt.split("\n"))
    lines=[re.sub(r"\s{2,}"," ",ln.strip()) for ln in lines if ln.strip()]

    ref_idx={}
    for i,ln in enumerate(lines):
        if m:=re.match(r"^([A-Z0-9]{3,})\s+[A-Z]{3}\s",ln):
            ref_idx.setdefault(m.group(1), i)

    for r in rows:
        if r["Custom Code"] and r["Code EAN"]: continue
        start=ref_idx.get(r["Reference"]);   end=start
        if start is None: continue
        while end<len(lines) and end-start<20 and not re.match(r"^[A-Z0-9]{3,}\s+[A-Z]{3}\s", lines[end]):
            end+=1
        snippet=" ".join(lines[start:end])
        seqs=re.findall(r"\d{6,14}", snippet)
        hts=[s for s in seqs if HTS_PAT.match(s)]
        upc=[s for s in seqs if UPC_PAT.match(s)]
        if hts and not r["Custom Code"]: r["Custom Code"]=hts[0]
        if upc and not r["Code EAN"]:    r["Code EAN"]=upc[0]

# ────────────────────────────  ENDPOINT  ─────────────────────────────────────
@app.post("/api/convert")
@app.post("/")
def convert():
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded", 400

        all_rows=[]
        for p in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                p.save(tmp.name)
                inv_num = (m.group(1) if (m:=re.search(r"SIP(\d+)", p.filename or "")) else "")

                rows = (
                    extract_original(tmp.name) +
                    extract_slice(tmp.name, inv_num) +
                    extract_new_provider(tmp.name, inv_num)
                )

                # quita duplicados
                seen=set(); uniq=[]
                for r in rows:
                    key=(r["Reference"], r["Code EAN"], r["Invoice Number"])
                    if key not in seen:
                        seen.add(key); uniq.append(r)

                # completar códigos faltantes
                complete_missing_codes(tmp.name, uniq)

                all_rows.extend(uniq)
            os.unlink(tmp.name)

        if not all_rows:
            return "Sin registros extraídos", 400

        wb = Workbook(); ws = wb.active; ws.append(COLS)
        for r in all_rows:
            ws.append([r[c] for c in COLS])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name="extracted_data.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
