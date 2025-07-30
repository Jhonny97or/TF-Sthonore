#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TF-StHonoré – Conversor PDF → Excel (v2025-07)
· extract_original          (facturas & proformas clásicas)
· extract_slice             (layout por columnas fijas)
· extract_new_provider      (proveedor Dior “No. Description … Each”)
· extract_tepf_scalp        (líneas TE/PF … UN xx … gencod)
· extract_new_layout        (nuevo layout No. Description UPC … usando posiciones fijas)
"""

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

# ─────────────────────────────────────────────────────────────────────────────
# Configuración logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
# Suprimir solo los warnings de CropBox de pdfplumber
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Columnas finales para Excel
# ─────────────────────────────────────────────────────────────────────────────
COLS = [
    "Reference","Code EAN","Custom Code","Description","Origin",
    "Quantity","GrossUnitPrice","NetUnitPrice","POSM FOC",
    "GrossTotalExclVAT","TotalAI","Invoice Number","Order Name","gencod"
]

# ─────────────────────────────────────────────────────────────────────────────
# Patrones globales
# ─────────────────────────────────────────────────────────────────────────────
ORD_NAME_PAT = re.compile(r"V\/CDE[^\n]*?ORD(?:ER)?\s*Nr\s*[:\-]\s*(.+)", re.I)
FC_PAT       = re.compile(r"FC-\d{3}-\d{2}-\d{5}")

# ─────────────────────────────────────────────────────────────────────────────
# Helpers numéricos
# ─────────────────────────────────────────────────────────────────────────────
def fnum(txt: str) -> float:
    return float(txt.replace(".", "").replace(",", ".").strip() or 0)

def to_int2(txt: str) -> int:
    d = re.sub(r"[^\d]", "", txt or "")
    return int(d) if d else 0

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 1: facturas & proformas clásicas
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
    if "PROFORMA" in up or ("ACKNOWLEDGE" in up and "RECEPTION" in up):
        return "proforma"
    return "factura"

def extract_original(pdf_path: str) -> List[dict]:
    rows: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        txt_all = "\n".join(p.extract_text() or "" for p in pdf.pages)
        kind = doc_kind(txt_all)

        inv, plv = "", False
        if kind == "factura":
            if m := INV_PAT.search(txt_all): inv = m.group(1)
            if PLV_PAT.search(txt_all):       plv = True
        else:
            if m := PROF_PAT.search(txt_all):       inv = m.group(1)
            elif m := ORDER_EN.search(txt_all):     inv = m.group(1)
            elif m := ORDER_FR.search(txt_all):     inv = m.group(1)

        invoice_no = inv + ("PLV" if plv else "")
        origin_global = ""

        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            # detectar origen
            for ln in lines:
                if mo := ORG_PAT.search(ln):
                    if mo.group(1).strip():
                        origin_global = mo.group(1).strip()
            # extraer líneas
            for i, ln in enumerate(lines):
                line = ln.strip()
                if kind == "factura" and (m := ROW_FACT.match(line)):
                    ref, ean, custom, qty, unit, total = m.groups()
                    desc = ""
                    if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]):
                        desc = lines[i+1].strip()
                    rows.append({
                        "Reference":         ref,
                        "Code EAN":          ean,
                        "Custom Code":       custom,
                        "Description":       desc,
                        "Origin":            origin_global,
                        "Quantity":          to_int2(qty),
                        "GrossUnitPrice":    fnum(unit),
                        "NetUnitPrice":      fnum(unit),
                        "POSM FOC":          "",
                        "GrossTotalExclVAT": fnum(total),
                        "TotalAI":           fnum(total),
                        "Invoice Number":    invoice_no,
                    })
                elif kind == "proforma" and (m := ROW_PROF_D.match(line)):
                    ref, ean, custom, qty, unit, total = m.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append({
                        "Reference":         ref,
                        "Code EAN":          ean,
                        "Custom Code":       custom,
                        "Description":       desc,
                        "Origin":            origin_global,
                        "Quantity":          to_int2(qty),
                        "GrossUnitPrice":    fnum(unit),
                        "NetUnitPrice":      fnum(unit),
                        "POSM FOC":          "",
                        "GrossTotalExclVAT": fnum(total),
                        "TotalAI":           fnum(total),
                        "Invoice Number":    invoice_no,
                    })
                elif kind == "proforma" and (m := ROW_PROF.match(line)):
                    ref, ean, unit, qty = m.groups()
                    qty_i, unit_f = to_int2(qty), fnum(unit)
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append({
                        "Reference":         ref,
                        "Code EAN":          ean,
                        "Custom Code":       "",
                        "Description":       desc,
                        "Origin":            origin_global,
                        "Quantity":          qty_i,
                        "GrossUnitPrice":    unit_f,
                        "NetUnitPrice":      unit_f,
                        "POSM FOC":          "",
                        "GrossTotalExclVAT": unit_f * qty_i,
                        "TotalAI":           unit_f * qty_i,
                        "Invoice Number":    invoice_no,
                    })
    # si todos los ítems de una factura comparten un solo origin, rellenar
    inv2orig = defaultdict(set)
    for r in rows:
        if r["Origin"]:
            inv2orig[r["Invoice Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2orig[r["Invoice Number"]]) == 1:
            r["Origin"] = next(iter(inv2orig[r["Invoice Number"]]))
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 2: slice por columnas fijas
# ─────────────────────────────────────────────────────────────────────────────
COL_BOUNDS = {
    "ref":    (0, 70),
    "desc":   (70, 340),
    "upc":    (340, 430),
    "ctry":   (430, 465),
    "hs":     (465, 535),
    "qty":    (535, 585),
    "unit":   (585, 635),
    "posm":   (635, 675),
    "total":  (675, 755),
}
REF_RE  = re.compile(r"^\d{5,6}[A-Z]?$")
UPC_RE  = re.compile(r"^\d{12,14}$")
NUM_RE  = re.compile(r"[0-9]")
SKIP_SN = {
    "No. Description","Total before","Bill To Ship","CIF CHILE",
    "Invoice","Ship From","Ship To","VAT/Tax","Shipping Te"
}

def rows_from_page(page) -> List[Dict[str,str]]:
    grouped = {}
    for ch in page.chars:
        y = round(ch["top"],1)
        grouped.setdefault(y, []).append(ch)
    rows = []
    for _, chs in sorted(grouped.items()):
        line = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if not line.strip() or any(skip in line for skip in SKIP_SN):
            continue
        cols = {k:"" for k in COL_BOUNDS}
        for c in sorted(chs, key=lambda c: c["x0"]):
            xm = (c["x0"] + c["x1"]) / 2
            for k, (x0, x1) in COL_BOUNDS.items():
                if x0 <= xm < x1:
                    cols[k] += c["text"]
                    break
        cols = {k:cols[k].strip() for k in cols}
        if not REF_RE.match(cols["ref"]): continue
        if not UPC_RE.match(cols["upc"]): continue
        if not NUM_RE.search(cols["qty"]): continue
        rows.append(cols)
    return rows

def extract_slice(pdf_path: str, inv_num: str) -> List[dict]:
    out: List[dict] = []
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
                    "GrossUnitPrice":    fnum(r["unit"]),
                    "NetUnitPrice":      fnum(r["unit"]),
                    "POSM FOC":          fnum(r["posm"]),
                    "GrossTotalExclVAT": fnum(r["total"]),
                    "TotalAI":           fnum(r["total"]),
                    "Invoice Number":    inv_num,
                })
    return out

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 3: proveedor Dior “No. Description … Each”
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
DIOR_BAS = re.compile(r"""
    ^\s*(?P<ref>\d{5,6}[A-Z]?)\s+
    (?P<upc>\d{12,14})\s+
    (?P<ctry>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4})\s+
    (?P<qty>[\d,]+)\s+Each\s+(?P<unit>[\d.,]+)\s+
    (?:-|(?P<posm>[\d.,]+))\s+(?P<total>[\d.,]+)
""", re.VERBOSE)

def extract_new_provider(pdf_path: str, inv_num: str) -> List[dict]:
    rows: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "No. Description" not in text:
                continue
            pend = None
            for ln in text.split("\n"):
                ls = ln.strip()
                if not ls or ls.startswith(("No. Description","Invoice")):
                    continue
                if m := DIOR_FULL.match(ln):
                    d = m.groupdict()
                    unit  = float(d["unit"].replace(",", ""))
                    total = float(d["total"].replace(",", ""))
                    posm  = float(d.get("posm", "0").replace(",", "")) if d.get("posm") else 0
                    rows.append({
                        "Reference":         d["ref"],
                        "Code EAN":          d["upc"],
                        "Custom Code":       d["hs"],
                        "Description":       d["desc"].strip(),
                        "Origin":            d["ctry"],
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    unit,
                        "NetUnitPrice":      unit,
                        "POSM FOC":          posm,
                        "GrossTotalExclVAT": total,
                        "TotalAI":           total,
                        "Invoice Number":    inv_num,
                    })
                    pend = None
                    continue
                if (m2 := DIOR_BAS.match(ln)) and pend:
                    d = m2.groupdict()
                    unit  = float(d["unit"].replace(",", ""))
                    total = float(d["total"].replace(",", ""))
                    rows.append({
                        "Reference":         d["ref"],
                        "Code EAN":          d["upc"],
                        "Custom Code":       d["hs"],
                        "Description":       pend.strip(),
                        "Origin":            d["ctry"],
                        "Quantity":          to_int2(d["qty"]),
                        "GrossUnitPrice":    unit,
                        "NetUnitPrice":      unit,
                        "POSM FOC":          0,
                        "GrossTotalExclVAT": total,
                        "TotalAI":           total,
                        "Invoice Number":    inv_num,
                    })
                    pend = None
                    continue
                if re.search(r"[A-Za-z]", ls) and not any(ls.startswith(x) for x in (
                    "Country of","Customer PO","Order No","Shipping Terms",
                    "Bill To","Finance","Total","CIF","Ship To"
                )):
                    pend = (pend + " " + ls) if pend else ls
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

def extract_tepf_scalp(pdf_path: str) -> List[dict]:
    rows: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            for i, ln in enumerate(lines):
                if m := TEPF_RE.match(ln):
                    d = m.groupdict()
                    genc = ""
                    for j in range(i+1, min(i+3, len(lines))):
                        if gm := re.search(r'gencod\s*[:\-]\s*(\d{13})', lines[j], re.I):
                            genc = gm.group(1)
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
                        "gencod":            genc,
                    })
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR 5: layout No. Description UPC … usando posiciones fijas
# ─────────────────────────────────────────────────────────────────────────────
COL5_BOUNDS = {
    "ref":     (0,   60),
    "desc":    (60, 350),
    "upc":     (350,450),
    "country": (450,490),
    "hs":      (490,580),
    "qty":     (580,640),
    "uom":     (640,700),
    "unit":    (700,760),
    "posm":    (760,820),
    "line":    (820,900),
}
REF5_RE    = re.compile(r"^\d+$")
UPC_RE5    = re.compile(r"^\d{12,14}$")
NUM_RE5    = re.compile(r"\d")
SKIP5      = {"No.","Description","UPC","Country","HS Code"}

def rows5_from_page(page) -> List[Dict[str,str]]:
    grouped = {}
    for ch in page.chars:
        y = round(ch["top"],1)
        grouped.setdefault(y, []).append(ch)
    rows = []
    for _, chs in sorted(grouped.items()):
        text = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if any(skip in text for skip in SKIP5):
            continue
        cols = {k:"" for k in COL5_BOUNDS}
        for c in sorted(chs, key=lambda c: c["x0"]):
            xm = (c["x0"] + c["x1"]) / 2
            for k,(x0,x1) in COL5_BOUNDS.items():
                if x0 <= xm < x1:
                    cols[k] += c["text"]
                    break
        cols = {k:cols[k].strip() for k in cols}
        if not REF5_RE.match(cols["ref"]): continue
        if not UPC_RE5.match(cols["upc"]): continue
        if not NUM_RE5.search(cols["qty"]): continue
        rows.append(cols)
    return rows

def extract_new_layout(pdf_path: str, inv_num: str) -> List[dict]:
    out: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            for r in rows5_from_page(p):
                out.append({
                    "Reference":         r["ref"],
                    "Code EAN":          r["upc"],
                    "Custom Code":       r["hs"],
                    "Description":       r["desc"],
                    "Origin":            r["country"],
                    "Quantity":          to_int2(r["qty"]),
                    "GrossUnitPrice":    fnum(r["unit"]),
                    "NetUnitPrice":      fnum(r["unit"]),
                    "POSM FOC":          fnum(r["posm"]),
                    "GrossTotalExclVAT": fnum(r["line"]),
                    "TotalAI":           fnum(r["line"]),
                    "Invoice Number":    inv_num,
                    "gencod":            ""
                })
    return out

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
            return "No file(s) uploaded", 400

        all_rows: List[dict] = []
        for pdf in pdfs:
            # guardar PDF temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf.save(tmp.name)
            logging.info(f"Saved temp PDF: {tmp.name}")

            # extraer texto completo
            with pdfplumber.open(tmp.name) as p:
                full_txt = "\n".join(pg.extract_text() or "" for pg in p.pages)

            # determinar invoice & order name
            inv_num = ""
            if m := FC_PAT.search(full_txt):
                inv_num = m.group(0)
            elif m := re.search(r"SIP(\d+)", pdf.filename or ""):
                inv_num = m.group(1)
            order_name = ORD_NAME_PAT.search(full_txt).group(1).strip() if ORD_NAME_PAT.search(full_txt) else ""

            # ejecutar todos los extractores
            o1 = extract_original(tmp.name)
            o2 = extract_slice(tmp.name, inv_num)
            o3 = extract_new_provider(tmp.name, inv_num)
            o4 = extract_tepf_scalp(tmp.name)
            for r in o4: r["Invoice Number"] = inv_num
            o5 = extract_new_layout(tmp.name, inv_num)

            # combinar y deduplicar
            combined = o1 + o2 + o3 + o4 + o5
            seen, uniq = set(), []
            for r in combined:
                key = (r["Reference"], r.get("Code EAN",""), r["Invoice Number"])
                if key not in seen:
                    seen.add(key)
                    r["Order Name"] = order_name
                    uniq.append(r)
            all_rows.extend(uniq)

            os.unlink(tmp.name)
            logging.info(f"Deleted temp PDF: {tmp.name}")

        if not all_rows:
            return "Sin registros extraídos", 400

        # normalización final
        for r in all_rows:
            if not r.get("NetUnitPrice"):
                r["NetUnitPrice"] = r.get("GrossUnitPrice", 0)
            if not r.get("GrossTotalExclVAT"):
                r["GrossTotalExclVAT"] = r.get("TotalAI", 0)
            if not r.get("Code EAN") and r.get("gencod"):
                r["Code EAN"] = r["gencod"]

        # escribir Excel
        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in all_rows:
            ws.append([r.get(c, "") for c in COLS])

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
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")

