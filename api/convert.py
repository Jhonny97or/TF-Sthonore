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
# Configuración de logging y supresión de warnings de CropBox
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Cabecera de columnas para el Excel
# ─────────────────────────────────────────────────────────────────────────────
COLS = [
    "Reference", "Code EAN", "Custom Code", "Description", "Origin",
    "Quantity", "GrossUnitPrice", "NetUnitPrice", "POSM FOC",
    "GrossTotalExclVAT", "TotalAI", "Invoice Number", "Order Name", "gencod"
]

# Patrones globales
ORD_NAME_PAT = re.compile(r"V\/CDE[^\n]*?ORD(?:ER)?\s*Nr\s*[:\-]\s*(.+)", re.I)
FC_PAT       = re.compile(r"FC-\d{3}-\d{2}-\d{5}")

# ─────────────────────────────────────────────────────────────────────────────
# Helpers numéricos
# ─────────────────────────────────────────────────────────────────────────────
def fnum(txt: str) -> float:
    return float(txt.replace(".", "").replace(",", ".").strip() or 0)

def to_float2(txt: str) -> float:
    t = txt.replace("\u202f", "").replace(" ", "")
    if t.count(",") == 1 and t.count(".") == 0:
        t = t.replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    return float(t or 0)

def to_int2(txt: str) -> int:
    digits = re.sub(r"[^\d]", "", txt or "")
    return int(digits) if digits else 0

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 1: facturas / proformas clásicas
# ─────────────────────────────────────────────────────────────────────────────
INV_PAT      = re.compile(r"(?:FACTURE|INVOICE)\D*(\d{6,})", re.I)
PROF_PAT     = re.compile(r"PROFORMA[\s\S]*?(\d{6,})", re.I)
ORDER_EN     = re.compile(r"ORDER\s+NUMBER\D*(\d{6,})", re.I)
ORDER_FR     = re.compile(r"N°\s*DE\s*COMMANDE\D*(\d{6,})", re.I)
PLV_PAT      = re.compile(r"FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT", re.I)
ORG_PAT      = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

ROW_FACT     = re.compile(r"^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,9})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)$")
ROW_PROF_D   = re.compile(r"^([A-Z]\w{3,11})\s+(\d{12,14})\s+(\d{6,10})\s+(\d[\d.,]*)\s+([\d.,]+)\s+([\d.,]+)$")
ROW_PROF     = re.compile(r"^([A-Z]\w{3,11})\s+(\d{12,14})\s+([\d.,]+)\s+([\d.,]+)$")

def doc_kind(text: str) -> str:
    up = text.upper()
    return "proforma" if "PROFORMA" in up or ("ACKNOWLEDGE" in up and "RECEPTION" in up) else "factura"

def extract_original(pdf_path: str) -> List[dict]:
    rows: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        txt_all = "\n".join(p.extract_text() or "" for p in pdf.pages)
        kind = doc_kind(txt_all)
        inv, plv = "", False
        if kind == "factura":
            if m := INV_PAT.search(txt_all): inv = m.group(1)
            if PLV_PAT.search(txt_all): plv = True
        else:
            if m := PROF_PAT.search(txt_all): inv = m.group(1)
            elif m := ORDER_EN.search(txt_all): inv = m.group(1)
            elif m := ORDER_FR.search(txt_all): inv = m.group(1)
        invoice_no = inv + ("PLV" if plv else "")
        origin_global = ""
        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            # detectar Origin
            for ln in lines:
                if mo := ORG_PAT.search(ln):
                    if mo.group(1).strip():
                        origin_global = mo.group(1).strip()
            # extraer líneas
            for i, ln in enumerate(lines):
                line = ln.strip()
                if kind=="factura" and (m:=ROW_FACT.match(line)):
                    ref, ean, custom, qty, unit, total = m.groups()
                    desc = ""
                    if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                        desc = lines[i+1].strip()
                    rows.append({
                        "Reference": ref, "Code EAN": ean, "Custom Code": custom,
                        "Description": desc, "Origin": origin_global,
                        "Quantity": to_int2(qty),
                        "GrossUnitPrice": fnum(unit), "NetUnitPrice": fnum(unit),
                        "POSM FOC": "", "GrossTotalExclVAT": fnum(total),
                        "TotalAI": fnum(total), "Invoice Number": invoice_no
                    })
                elif kind=="proforma" and (m:=ROW_PROF_D.match(line)):
                    ref, ean, unit, qty = m.groups()
                    qty_i = to_int2(qty); unit_f = fnum(unit)
                    desc = lines[i+1].strip() if i+1<len(lines) else ""
                    rows.append({
                        "Reference": ref, "Code EAN": ean, "Custom Code": "",
                        "Description": desc, "Origin": origin_global,
                        "Quantity": qty_i,
                        "GrossUnitPrice": unit_f, "NetUnitPrice": unit_f,
                        "POSM FOC": "", "GrossTotalExclVAT": unit_f*qty_i,
                        "TotalAI": unit_f*qty_i, "Invoice Number": invoice_no
                    })
                elif kind=="proforma" and (m:=ROW_PROF_D.match(line)):
                    ref, ean, custom, qty, unit, total = m.groups()
                    desc = lines[i+1].strip() if i+1<len(lines) else ""
                    rows.append({
                        "Reference": ref, "Code EAN": ean, "Custom Code": custom,
                        "Description": desc, "Origin": origin_global,
                        "Quantity": to_int2(qty),
                        "GrossUnitPrice": fnum(unit), "NetUnitPrice": fnum(unit),
                        "POSM FOC": "", "GrossTotalExclVAT": fnum(total),
                        "TotalAI": fnum(total), "Invoice Number": invoice_no
                    })
    # rellenar Origin si es único
    inv2orig = defaultdict(set)
    for r in rows:
        if r["Origin"]:
            inv2orig[r["Invoice Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2orig[r["Invoice Number"]])==1:
            r["Origin"] = next(iter(inv2orig[r["Invoice Number"]]))
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 2: layout por columnas fijas
# ─────────────────────────────────────────────────────────────────────────────
COL_BOUNDS = {
    "ref":(0,70),"desc":(70,340),"upc":(340,430),"ctry":(430,465),
    "hs":(465,535),"qty":(535,585),"unit":(585,635),
    "posm":(635,675),"total":(675,755)
}
REF_RE = re.compile(r"^\d{5,6}[A-Z]?$")
UPC_RE = re.compile(r"^\d{12,14}$")
NUM_RE = re.compile(r"[0-9]")
SKIP = {"No. Description","Total before","Bill To Ship","CIF CHILE","Invoice",
        "Ship From","Ship To","VAT/Tax","Shipping Te"}

def rows_from_page(p) -> List[Dict[str,str]]:
    grouped={}
    for ch in p.chars:
        grouped.setdefault(round(ch["top"],1), []).append(ch)
    rows=[]
    for _,chs in sorted(grouped.items()):
        text = "".join(c["text"] for c in sorted(chs,key=lambda c:c["x0"]))
        if not text.strip() or any(sn in text for sn in SKIP): continue
        cols={k:"" for k in COL_BOUNDS}
        for c in sorted(chs,key=lambda c:c["x0"]):
            mid=(c["x0"]+c["x1"])/2
            for k,(x0,x1) in COL_BOUNDS.items():
                if x0<=mid<x1:
                    cols[k]+=c["text"]; break
        cols={k:cols[k].strip() for k in cols}
        if not cols["ref"]:
            if rows: rows[-1]["desc"] += " " + cols["desc"]
            continue
        if not (REF_RE.match(cols["ref"]) and UPC_RE.match(cols["upc"])): continue
        if not NUM_RE.search(cols["qty"]): continue
        rows.append(cols)
    return rows

def extract_slice(pdf_path:str, inv_num:str) -> List[dict]:
    out=[]
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            for r in rows_from_page(p):
                out.append({
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
    return out

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 3: Dior “No. Description … Each”
# ─────────────────────────────────────────────────────────────────────────────
DIOR_FULL = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
""", re.VERBOSE)
DIOR_BAS  = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
""", re.VERBOSE)

def extract_new_provider(pdf_path:str, inv_num:str) -> List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            txt = p.extract_text() or ""
            if "No. Description" not in txt: continue
            pend=None
            for ln in txt.split("\n"):
                ls=ln.strip()
                if not ls or ls.startswith(("No. Description","Invoice")): continue
                if m:=DIOR_FULL.match(ln):
                    d=m.groupdict()
                    posm = float(d.get("posm","0").replace(",","")) if d.get("posm") else 0
                    rows.append({
                        "Reference": d["ref"], "Code EAN": d["upc"], "Custom Code": d["hs"],
                        "Description": d["desc"].strip(), "Origin": d["ctry"],
                        "Quantity": to_int2(d["qty"]),
                        "GrossUnitPrice": float(d["unit"].replace(",","")),
                        "NetUnitPrice":   float(d["unit"].replace(",","")),
                        "POSM FOC":       posm,
                        "GrossTotalExclVAT": float(d["total"].replace(",","")),
                        "TotalAI":           float(d["total"].replace(",","")),
                        "Invoice Number": inv_num
                    })
                    pend=None; continue
                if (m:=DIOR_BAS.match(ln)) and pend:
                    d=m.groupdict()
                    rows.append({
                        "Reference": d["ref"], "Code EAN": d["upc"], "Custom Code": d["hs"],
                        "Description": pend.strip(), "Origin": d["ctry"],
                        "Quantity": to_int2(d["qty"]),
                        "GrossUnitPrice": float(d["unit"].replace(",","")),
                        "NetUnitPrice":   float(d["unit"].replace(",","")),
                        "POSM FOC":       0,
                        "GrossTotalExclVAT": float(d["total"].replace(",","")),
                        "TotalAI":           float(d["total"].replace(",","")),
                        "Invoice Number": inv_num
                    })
                    pend=None; continue
                if re.search(r"[A-Za-z]", ls) and not any(ls.startswith(x) for x in (
                    "Country of","Customer PO","Order No","Shipping Terms",
                    "Bill To","Finance","Total","CIF","Ship To"
                )):
                    pend = (pend+" "+ls) if pend else ls
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 4: TE/PF … UN xx … gencod
# ─────────────────────────────────────────────────────────────────────────────
TEPF_RE = re.compile(
    r'^(?P<art>(?:TE|PF)\d+)\s+'
    r'(?P<desc>.+?)\s+UN\s*(?P<qty>\d+)\s+'
    r'(?P<gup>[\d,]+)\s+'
    r'(?P<ntp>[\d,]+)\s+'
    r'(?P<gtx>[\d,]+)\s+'
    r'(?P<nta>[\d,]+)'
)

def extract_tepf_scalp(pdf_path:str) -> List[dict]:
    rows=[]
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            for i,ln in enumerate((p.extract_text() or "").split("\n")):
                if m:=TEPF_RE.match(ln):
                    d=m.groupdict()
                    gencod=""
                    for j in range(i+1, min(i+3,len(ln))):
                        if gm:=re.search(r'gencod\s*[:\-]\s*(\d{13})', ln, re.I):
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
# EXTRACTOR 5: nuevo layout “No. Description UPC ... Line Amount”
# ─────────────────────────────────────────────────────────────────────────────
LAYOUT5_RE = re.compile(r"""
    ^\s*(?P<ref>\d+)\s+
    (?P<desc>.+?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<country>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+
    (?P<uom>\w+)\s+
    (?P<unit>[\d.,]+)\s+
    (?P<posm>-|[\d.,]+)\s+
    (?P<line>[\d.,]+)
""", re.VERBOSE)

def extract_new_layout(pdf_path:str, inv_num:str) -> List[dict]:
    rows=[]
    logging.info("  > extract_new_layout")
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            for ln in (p.extract_text() or "").split("\n"):
                if m:=LAYOUT5_RE.match(ln):
                    d=m.groupdict()
                    posm_raw=d["posm"]
                    posm_val = "" if posm_raw=="-" else fnum(posm_raw)
                    rows.append({
                        "Reference":         d["ref"],
                        "Code EAN":          d["upc"],
                        "Custom Code":       d["hs"],
                        "Description":       d["desc"].strip(),
                        "Origin":            d["country"],
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    fnum(d["unit"]),
                        "NetUnitPrice":      fnum(d["unit"]),
                        "POSM FOC":          posm_val,
                        "GrossTotalExclVAT": fnum(d["line"]),
                        "TotalAI":           fnum(d["line"]),
                        "Invoice Number":    inv_num,
                        "gencod":            ""
                    })
    logging.info(f"  < extract_new_layout: {len(rows)} rows")
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

        all_rows: List[dict] = []
        for pdf in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf.save(tmp.name)
            logging.info(f"Saved {tmp.name}")

            # texto completo
            with pdfplumber.open(tmp.name) as p:
                full_txt = "\n".join(pg.extract_text() or "" for pg in p.pages)

            inv_num = ""
            if m:=FC_PAT.search(full_txt):
                inv_num = m.group(0)
            elif m:=re.search(r"SIP(\d+)", pdf.filename or ""):
                inv_num = m.group(1)

            order_name = ORD_NAME_PAT.search(full_txt).group(1).strip() if ORD_NAME_PAT.search(full_txt) else ""

            # ejecutar extractores
            tasks = [
                ("original", lambda: extract_original(tmp.name)),
                ("slice",    lambda: extract_slice(tmp.name, inv_num)),
                ("provider", lambda: extract_new_provider(tmp.name, inv_num)),
                ("tepf",     lambda: extract_tepf_scalp(tmp.name)),
                ("layout5",  lambda: extract_new_layout(tmp.name, inv_num)),
            ]
            segs=[]
            for name, fn in tasks:
                logging.info(f"–> {name}…")
                t0=time.time()
                try:
                    part = fn()
                    if name=="tepf":
                        for r in part: r["Invoice Number"]=inv_num
                except Exception:
                    logging.exception(f"Error in {name}")
                    part=[]
                logging.info(f"<– {name} returned {len(part)} in {time.time()-t0:.2f}s")
                segs.extend(part)

            # desduplicar + order_name
            seen, uniq = set(), []
            for r in segs:
                key=(r["Reference"], r.get("Code EAN",""), r["Invoice Number"])
                if key not in seen:
                    seen.add(key)
                    r["Order Name"]=order_name
                    uniq.append(r)
            all_rows.extend(uniq)

            os.unlink(tmp.name)
            logging.info(f"Deleted {tmp.name}")

        if not all_rows:
            logging.warning("No rows extracted")
            return "Sin registros extraídos", 400

        # normalizar
        for r in all_rows:
            if not r.get("NetUnitPrice"):
                r["NetUnitPrice"]=r.get("GrossUnitPrice",0)
            if not r.get("GrossTotalExclVAT"):
                r["GrossTotalExclVAT"]=r.get("TotalAI",0)
            if not r.get("Code EAN") and r.get("gencod"):
                r["Code EAN"]=r["gencod"]

        # generar Excel
        wb=Workbook(); ws=wb.active; ws.append(COLS)
        for r in all_rows:
            ws.append([r.get(c,"") for c in COLS])

        buf=BytesIO(); wb.save(buf); buf.seek(0)
        logging.info("<< /convert done")
        return send_file(
            buf, as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception:
        logging.exception("Unhandled error in /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")


