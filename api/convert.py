#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TF‑StHonoré – Conversor PDF → Excel (v2025‑07)
· extract_original
· extract_slice
· extract_new_provider
· extract_tepf_scalp
· extract_new_layout
"""

import logging
import os
import re
import tempfile
import traceback
import time
from collections import defaultdict
from io import BytesIO
from typing import Dict, List

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

# ─────────────────────────────────────────────────────────────────────────────
# Configuración logging y supresión de warnings molestos de pdfplumber
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
# Suprimir únicamente las advertencias de CropBox
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Cabecera Excel
# ─────────────────────────────────────────────────────────────────────────────
COLS = [
    "Reference", "Code EAN", "Custom Code", "Description", "Origin",
    "Quantity", "GrossUnitPrice", "NetUnitPrice", "POSM FOC",
    "GrossTotalExclVAT", "TotalAI", "Invoice Number", "Order Name", "gencod"
]

# Patrones globales
ORD_NAME_PAT = re.compile(r"V\/CDE[^\n]*?ORD(?:ER)?\s*[:\-]\s*(.+)", re.I)
FC_PAT       = re.compile(r"FC-\d{3}-\d{2}-\d{5}")

# Helpers numéricos
def fnum(txt: str) -> float:
    return float(txt.replace(".", "").replace(",", ".").strip() or 0)

def to_float2(txt: str) -> float:
    t = txt.replace("\u202f","").replace(" ","")
    if t.count(",")==1 and t.count(".")==0: t = t.replace(",",".")
    elif t.count(".")>1:                t = t.replace(".","")
    return float(t or 0)

def to_int2(txt: str) -> int:
    digits = re.sub(r"[^\d]","", txt or "")
    return int(digits) if digits else 0

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 1: facturas / proformas clásicas
# ─────────────────────────────────────────────────────────────────────────────
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

def doc_kind(text: str) -> str:
    up = text.upper()
    return "proforma" if "PROFORMA" in up or ("ACKNOWLEDGE" in up and "RECEPTION" in up) else "factura"

def extract_original(pdf_path: str) -> List[dict]:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        all_txt = "\n".join(page.extract_text() or "" for page in pdf.pages)
        kind = doc_kind(all_txt)

        inv_global, plv_global = "", False
        if kind == "factura":
            if m := INV_PAT.search(all_txt): inv_global = m.group(1)
            if PLV_PAT.search(all_txt):     plv_global = True
        else:
            if m := PROF_PAT.search(all_txt):       inv_global = m.group(1)
            elif m := ORDER_PAT_EN.search(all_txt): inv_global = m.group(1)
            elif m := ORDER_PAT_FR.search(all_txt): inv_global = m.group(1)

        invoice_full = inv_global + ("PLV" if plv_global else "")
        org_global = ""

        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            for ln in lines:
                if mo := ORG_PAT.search(ln) and mo.group(1).strip():
                    org_global = mo.group(1).strip()
            for i, raw in enumerate(lines):
                ln = raw.strip()
                # Factura clásica
                if kind == "factura" and (mf := ROW_FACT.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                    desc = ""
                    if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                        desc = lines[i+1].strip()
                    rows.append({
                        "Reference":ref, "Code EAN":ean, "Custom Code":custom,
                        "Description":desc, "Origin":org_global,
                        "Quantity":to_int2(qty_s),
                        "GrossUnitPrice":fnum(unit_s),
                        "NetUnitPrice":fnum(unit_s),
                        "POSM FOC":"",
                        "GrossTotalExclVAT":fnum(tot_s),
                        "TotalAI":fnum(tot_s),
                        "Invoice Number":invoice_full
                    })
                # Proforma Dior
                elif kind=="proforma" and (mpd:=ROW_PROF_DIOR.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = mpd.groups()
                    desc = lines[i+1].strip() if i+1<len(lines) else ""
                    rows.append({
                        "Reference":ref, "Code EAN":ean, "Custom Code":custom,
                        "Description":desc, "Origin":org_global,
                        "Quantity":to_int2(qty_s),
                        "GrossUnitPrice":fnum(unit_s),
                        "NetUnitPrice":fnum(unit_s),
                        "POSM FOC":"",
                        "GrossTotalExclVAT":fnum(tot_s),
                        "TotalAI":fnum(tot_s),
                        "Invoice Number":invoice_full
                    })
                # Proforma genérico
                elif kind=="proforma" and (mp:=ROW_PROF.match(ln)):
                    ref, ean, unit_s, qty_s = mp.groups()
                    desc = lines[i+1].strip() if i+1<len(lines) else ""
                    qty, unit = to_int2(qty_s), fnum(unit_s)
                    rows.append({
                        "Reference":ref, "Code EAN":ean, "Custom Code":"",
                        "Description":desc, "Origin":org_global,
                        "Quantity":qty,
                        "GrossUnitPrice":unit,
                        "NetUnitPrice":unit,
                        "POSM FOC":"",
                        "GrossTotalExclVAT":unit*qty,
                        "TotalAI":unit*qty,
                        "Invoice Number":invoice_full
                    })
    # Rellenar Origin faltantes si todos los items comparten uno solo
    inv2org = defaultdict(set)
    for r in rows:
        if r["Origin"]:
            inv2org[r["Invoice Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2org[r["Invoice Number"]])==1:
            r["Origin"] = next(iter(inv2org[r["Invoice Number"]]))
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 2: slice por columnas fijas
# ─────────────────────────────────────────────────────────────────────────────
COL_BOUNDS: Dict[str, tuple] = {
    "ref":(0,70),"desc":(70,340),"upc":(340,430),"ctry":(430,465),
    "hs":(465,535),"qty":(535,585),"unit":(585,635),
    "posm":(635,675),"total":(675,755)
}
REF_PAT = re.compile(r"^\d{5,6}[A-Z]?$")
UPC_PAT = re.compile(r"^\d{12,14}$")
NUM_PAT = re.compile(r"[0-9]")
SKIP_SNIPPETS = {
    "No. Description","Total before","Bill To Ship","CIF CHILE",
    "Invoice","Ship From","Ship To","VAT/Tax","Shipping Te"
}

def clean(txt:str)->str:
    return txt.replace("\u202f"," ").strip()

def rows_from_page(page)->List[Dict[str,str]]:
    rows, grouped = [], {}
    for ch in page.chars:
        grouped.setdefault(round(ch["top"],1), []).append(ch)
    for _, chs in sorted(grouped.items()):
        raw = "".join(c["text"] for c in sorted(chs,key=lambda c:c["x0"]))
        if not raw.strip() or any(sn in raw for sn in SKIP_SNIPPETS):
            continue
        cols = {k:"" for k in COL_BOUNDS}
        for ch in sorted(chs,key=lambda c:c["x0"]):
            xm = (ch["x0"]+ch["x1"])/2
            for k,(x0,x1) in COL_BOUNDS.items():
                if x0<=xm<x1:
                    cols[k]+=ch["text"]
                    break
        cols = {k:clean(v) for k,v in cols.items()}
        if not cols["ref"]:
            if rows:
                rows[-1]["desc"] += " " + cols["desc"]
            continue
        if not (REF_PAT.match(cols["ref"]) and UPC_PAT.match(cols["upc"])):
            continue
        if not NUM_PAT.search(cols["qty"]):
            continue
        rows.append(cols)
    return rows

def extract_slice(pdf_path:str, inv_num:str)->List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            for r in rows_from_page(p):
                rows.append({
                    "Reference":         r["ref"],
                    "Code EAN":          r["upc"],
                    "Custom Code":       r["hs"],
                    "Description":       r["desc"],
                    "Origin":            r["ctry"],
                    "Quantity":          to_int2(r["qty"]),
                    "GrossUnitPrice":    to_float2(r["unit"]),
                    "NetUnitPrice":      to_float2(r["unit"]),
                    "POSM FOC":          to_float2(r["posm"]),
                    "GrossTotalExclVAT": to_float2(r["total"]),
                    "TotalAI":           to_float2(r["total"]),
                    "Invoice Number":    inv_num
                })
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 3: Dior “No. Description … Each”
# ─────────────────────────────────────────────────────────────────────────────
pat_full = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
""", re.VERBOSE)
pat_basic = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
""", re.VERBOSE)

def extract_new_provider(pdf_path:str, inv_num:str)->List[dict]:
    def nfloat(s:str)->float: return float(s.replace(",","") or 0)
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if "No. Description" not in txt:
                continue
            pend=None
            for ln in txt.split("\n"):
                ls=ln.strip()
                if not ls or ls.startswith(("No. Description","Invoice")):
                    continue
                if m:=pat_full.match(ln):
                    d=m.groupdict()
                    rows.append({
                        "Reference":         d["ref"],
                        "Code EAN":          d["upc"],
                        "Custom Code":       d["hs"],
                        "Description":       d["desc"].strip(),
                        "Origin":            d["ctry"],
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    nfloat(d["unit"]),
                        "NetUnitPrice":      nfloat(d["unit"]),
                        "POSM FOC":          nfloat(d.get("posm","")),
                        "GrossTotalExclVAT": nfloat(d["total"]),
                        "TotalAI":           nfloat(d["total"]),
                        "Invoice Number":    inv_num
                    })
                    pend=None
                    continue
                if (mb:=pat_basic.match(ln)) and pend:
                    d=mb.groupdict()
                    rows.append({
                        "Reference":         d["ref"],
                        "Code EAN":          d["upc"],
                        "Custom Code":       d["hs"],
                        "Description":       pend.strip(),
                        "Origin":            d["ctry"],
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    nfloat(d["unit"]),
                        "NetUnitPrice":      nfloat(d["unit"]),
                        "POSM FOC":          nfloat(d.get("posm","")),
                        "GrossTotalExclVAT": nfloat(d["total"]),
                        "TotalAI":           nfloat(d["total"]),
                        "Invoice Number":    inv_num
                    })
                    pend=None
                    continue
                if re.search(r"[A-Za-z]", ls) and not any(ls.startswith(p) for p in (
                    "Country of","Customer PO","Order No","Shipping Terms",
                    "Bill To","Finance","Total","CIF","Ship To"
                )):
                    pend=(pend+" "+ls) if pend else ls
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 4: TE/PF … UN xx … gencod
# ─────────────────────────────────────────────────────────────────────────────
TEPF_REGEX = re.compile(
    r'^(?P<art>(?:TE|PF)\d+)\s+'
    r'(?P<desc>.+?)\s+UN\s*(?P<qty>\d+)\s+'
    r'(?P<gup>[\d,]+)\s+'
    r'(?P<ntp>[\d,]+)\s+'
    r'(?P<gtx>[\d,]+)\s+'
    r'(?P<nta>[\d,]+)'
)

def extract_tepf_scalp(pdf_path:str)->List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines=(page.extract_text() or "").split("\n")
            for i,line in enumerate(lines):
                if m:=TEPF_REGEX.match(line):
                    d=m.groupdict()
                    gencod=""
                    for j in range(i+1, min(i+3,len(lines))):
                        if gm:=re.search(r'gencod\s*[:\-]\s*(\d{13})', lines[j], re.I):
                            gencod=gm.group(1)
                            break
                    rows.append({
                        "Reference":         d["art"],
                        "Code EAN":          "",
                        "Custom Code":       "",
                        "Description":       d["desc"].strip(),
                        "Origin":            "",
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    fnum(d["gup"]),
                        "NetUnitPrice":      fnum(d["ntp"]),
                        "POSM FOC":          "",
                        "GrossTotalExclVAT": fnum(d["gtx"]),
                        "TotalAI":           fnum(d["nta"]),
                        "Invoice Number":    "",
                        "gencod":            gencod
                    })
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 5: nuevo layout “No. Description UPC … Line Amount”
# ─────────────────────────────────────────────────────────────────────────────
COL5_HEADER = re.compile(
    r"No\.\s+Description\s+UPC\s+Country of\s+Origin\s+HS Code\s+Quantity\s+U\.of M\.\s+Unit Price\s+POSM FOC\s+Line Amount",
    re.I
)
COL5_LINE = re.compile(r"""
    ^\s*(?P<ref>\d+)\s+
    (?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<country>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+
    (?P<uom>\w+)\s+
    (?P<unit>[\d.,]+)\s+
    (?P<posm>[\d\.,\-]+)\s+
    (?P<line>[\d.,]+)
""", re.VERBOSE)

def extract_new_layout(pdf_path:str, inv_num:str)->List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt=page.extract_text() or ""
            if not COL5_HEADER.search(txt):
                continue
            for ln in txt.split("\n"):
                if m:=COL5_LINE.match(ln):
                    d=m.groupdict()
                    posm_raw=d["posm"].strip()
                    rows.append({
                        "Reference":         d["ref"],
                        "Code EAN":          d["upc"],
                        "Custom Code":       d["hs"],
                        "Description":       d["desc"].strip(),
                        "Origin":            d["country"],
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    fnum(d["unit"]),
                        "NetUnitPrice":      fnum(d["unit"]),
                        "POSM FOC":          "" if posm_raw in ("-","") else fnum(posm_raw),
                        "GrossTotalExclVAT": fnum(d["line"]),
                        "TotalAI":           fnum(d["line"]),
                        "Invoice Number":    inv_num,
                        "gencod":            ""
                    })
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# ENDPOINT /convert
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/api/convert", methods=["POST"])
@app.route("/", methods=["POST"])
def convert():
    logging.info(">> /convert start")
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            logging.warning("No files uploaded")
            return "No file(s) uploaded", 400

        all_rows = []
        for pdf in pdfs:
            # Guardar PDF temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf.save(tmp.name)
            logging.info(f"Saved temp PDF: {tmp.name}")

            # Leer texto completo
            with pdfplumber.open(tmp.name) as p:
                full_txt = "\n".join(pg.extract_text() or "" for pg in p.pages)

            # Determinar invoice number
            inv_num = ""
            if m := FC_PAT.search(full_txt):
                inv_num = m.group(0)
            elif m := re.search(r"SIP(\d+)", pdf.filename or ""):
                inv_num = m.group(1)

            # Determinar order name
            order_name = ""
            if m := ORD_NAME_PAT.search(full_txt):
                order_name = m.group(1).strip()

            # Ejecutar extractores con temporización
            tasks = [
                ("extract_original",     lambda: extract_original(tmp.name)),
                ("extract_slice",        lambda: extract_slice(tmp.name, inv_num)),
                ("extract_new_provider", lambda: extract_new_provider(tmp.name, inv_num)),
                ("extract_tepf_scalp",   lambda: extract_tepf_scalp(tmp.name)),
                ("extract_new_layout",   lambda: extract_new_layout(tmp.name, inv_num)),
            ]
            segments = []
            for name, fn in tasks:
                logging.info(f"–> {name}…")
                t0 = time.time()
                try:
                    seg = fn()
                    if name == "extract_tepf_scalp":
                        for r in seg:
                            r["Invoice Number"] = inv_num
                except Exception:
                    logging.exception(f"Error in {name}")
                    seg = []
                dt = time.time() - t0
                logging.info(f"<– {name} returned {len(seg)} rows in {dt:.2f}s")
                segments.extend(seg)

            # Desduplicar y añadir Order Name
            seen, uniq = set(), []
            for r in segments:
                key = (r["Reference"], r.get("Code EAN",""), r["Invoice Number"])
                if key not in seen:
                    seen.add(key)
                    r["Order Name"] = order_name
                    uniq.append(r)
            all_rows.extend(uniq)

            os.unlink(tmp.name)
            logging.info(f"Deleted temp PDF: {tmp.name}")

        if not all_rows:
            logging.warning("No rows extracted")
            return "Sin registros extraídos", 400

        # Normalización final
        for r in all_rows:
            if not r.get("NetUnitPrice"):
                r["NetUnitPrice"] = r.get("GrossUnitPrice", 0)
            if not r.get("GrossTotalExclVAT"):
                r["GrossTotalExclVAT"] = r.get("TotalAI", 0)
            if not r.get("Code EAN") and r.get("gencod"):
                r["Code EAN"] = r["gencod"]

        # Escribir Excel
        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in all_rows:
            ws.append([r.get(c,"") for c in COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        logging.info("<< /convert done")
        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception:
        logging.exception("Unhandled error in /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")


