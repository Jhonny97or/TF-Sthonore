# app.py  ── listo para Vercel o ejecución local JHONNY 
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
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

COLS = [
    "Reference", "Code EAN", "Custom Code", "Description",
    "Origin", "Quantity", "Unit Price", "Total Price", "Invoice Number"
]

# ───────────────  UTIL: sacar número de invoice del PDF  ───────────────
INV_RE = re.compile(r"(?:INVOICE|FACTURE|FACTURA)\s*(?:NO\.?|N°|NUMBER)?\s*[:\-]?\s*(\w[\w\-\/]{4,})", re.I)
SIP_RE = re.compile(r"\bSIP(\d{6,})\b", re.I)
PO_RE  = re.compile(r"(?:ORDER|PO)\s*(?:NO\.?|N°|NUMBER)?\s*[:\-]?\s*(\w[\w\-\/]{4,})", re.I)

def parse_invoice_number_from_pdf(pdf_path: str) -> str:
    def _clean(tok: str) -> str:
        return re.sub(r"(^[^A-Z0-9]+|[^A-Z0-9\/\-]+$)", "", tok.strip(), flags=re.I)

    def _valid(tok: str) -> bool:
        if not tok:
            return False
        if re.fullmatch(r"fech[aá]?", tok, flags=re.I):  # evita FECHA
            return False
        return any(ch.isdigit() for ch in tok) and len(tok) >= 4

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full = "\n".join((p.extract_text() or "") for p in pdf.pages)
        lines = [ln.strip() for ln in full.split("\n") if ln.strip()]

        hdr_pat = re.compile(r"(INVOICE|FACTURA|FACTURE)", re.I)
        no_pat  = re.compile(r"(?:No\.?|N°|Number|Num\.?)\s*[:\-]?\s*(\S+)", re.I)
        for ln in lines:
            if not hdr_pat.search(ln):
                continue
            m = no_pat.search(ln)
            if m:
                cand = _clean(m.group(1))
                if _valid(cand):
                    return cand
            for tok in re.split(r"\s+", ln):
                tokc = _clean(tok)
                if _valid(tokc):
                    return tokc

        m = re.search(r"(?:INVOICE|FACTURA|FACTURE)[^\n]{0,40}?([A-Z]*\d[\w\-\/]*)", full, re.I)
        if m:
            cand = _clean(m.group(1))
            if _valid(cand):
                return cand
    except Exception:
        pass
    return ""


# ───────────────────────  PATRONES PARA CÓDIGOS  ────────────────────────────
HTS_PAT = re.compile(r"^\d{6,10}$")
UPC_PAT = re.compile(r"^\d{11,14}$")

# ─────────────────────  EXTRACTOR 1  (facturas clásicas + Dior Proforma)  ───────────────────
INV_PAT      = re.compile(r"(?:FACTURE|INVOICE)\D*(\d{6,})", re.I)
PROF_PAT     = re.compile(r"PROFORMA[\s\S]*?(\d{6,})", re.I)
ORDER_PAT_EN = re.compile(r"ORDER\s+NUMBER\D*(\d{6,})", re.I)
ORDER_PAT_FR = re.compile(r"N°\s*DE\s*COMMANDE\D*(\d{6,})", re.I)
PLV_PAT      = re.compile(r"FACTURE\s+SANS\s+PAIEMENT|INVOICE\s+WITHOUT\s+PAYMENT", re.I)
ORG_PAT      = re.compile(r"PAYS D['’]?ORIGINE[^:]*:\s*(.+)", re.I)

# NUEVO: patrón para capturar el Your Order Nr en las proformas Dior
ORDER_NR_PAT = re.compile(r"V/CDE[-\s]?Y/ORD\s*Nr\s*:\s*(.+)", re.I)

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
        your_order_nr = ""   # <── nuevo campo

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

        # Buscar el Your Order Nr en todo el texto
        if mo := ORDER_NR_PAT.search(all_txt):
            your_order_nr = mo.group(1).strip()

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
                        "Your Order Nr": your_order_nr,   # <── agregado
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
                        "Your Order Nr": your_order_nr,
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
                        "Your Order Nr": your_order_nr,
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


# ─────────────────────  EXTRACTOR 2  (por coordenadas: LVMH)  ──────────────────────
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

ORDER_NR_PAT2 = re.compile(r"(?:YOUR\s+ORDER\s+Nr|V/CDE)\s*[:\-]?\s*(.+)", re.I)

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

        # Caso 1: fila normal con referencia y cantidad
        if cols["ref"] and REF_PAT.match(cols["ref"]) and NUM_PAT.search(cols["qty"]):
            rows.append(cols)

        # Caso 2: línea sin referencia → se trata como descripción extendida
        elif not cols["ref"] and rows:
            rows[-1]["desc"] = (rows[-1]["desc"] + " " + line_txt.strip()).strip()

    return rows

def extract_slice(pdf_path: str, inv_number: str) -> List[dict]:
    rows=[]
    your_order_nr=""

    with pdfplumber.open(pdf_path) as pdf:
        full_txt="\n".join(page.extract_text() or "" for page in pdf.pages)
        if mo := ORDER_NR_PAT2.search(full_txt):
            your_order_nr = mo.group(1).strip()

        for page in pdf.pages:
            for r in rows_from_page(page):
                rows.append({
                    "Reference": r.get("ref",""),
                    "Code EAN": r.get("upc",""),
                    "Custom Code": r.get("hs",""),
                    "Description": r.get("desc",""),
                    "Origin": r.get("ctry",""),
                    "Quantity": to_int2(r.get("qty","0")),
                    "Unit Price": to_float2(r.get("unit","0")),
                    "Total Price": to_float2(r.get("total","0")),
                    "Invoice Number": inv_number,
                    "Your Order Nr": your_order_nr  # <── agregado fijo
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

                if re.search(r"[A-Za-z]", ln_s):
                    skip_pref=("Country of","Customer PO","Order No",
                               "Shipping Terms","Bill To","Finance",
                               "Total","CIF","Ship To")
                    if not any(ln_s.startswith(p) for p in skip_pref):
                        pending_desc=(pending_desc+" "+ln_s) if pending_desc else ln_s
    return rows

# ──────────────────  EXTRACTOR 4 (Interparfums Italia: totales inline)  ─────
HS_ORG_PAT = re.compile(r"HS\s*Code:\s*(?P<hs>\d{8,14})\s*,\s*Origin:\s*(?P<org>[A-Z]{2})", re.I)
EAN_PAT    = re.compile(r"EAN\s*Code:\s*(?P<ean>\d{12,14})", re.I)

HEAD_INLINE_PAT = re.compile(
    r"""^
    (?P<ref>[A-Z0-9]{3,}\w*)\s+
    (?P<desc>.+?)\s+
    (?P<qty>[\d\.\s]+)\s+PZ\s+
    (?P<unit>[\d\.,]+)\s+
    (?P<gross>[\d\.,]+)
    (?:\s+(?P<disc>-?\d+%)\s+(?P<net>[\d\.,]+)
       |\s+(?P<net2>[\d\.,]+)
    )
    \s+(?P<vat>[A-Z]{2})\s*$
    """,
    re.X | re.I
)

def _fnum_euro(s: str) -> float:
    if not s: return 0.0
    t = s.replace("\u202f","").replace(" ","")
    if t.count(",")==1:
        t = t.replace(".","").replace(",",".")
    else:
        t = t.replace(",","")
    try:
        return float(t)
    except:
        return 0.0

def _qty_to_int(s: str) -> int:
    return int(s.replace("\u202f","").replace(" ","").replace(".","").replace(",","") or 0)

def extract_interparfums_blocks(pdf_path: str, invoice_number: str) -> List[dict]:
    rows: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = [ (page.extract_text() or "").replace("\u202f"," ").split("\n") ][0]
            for i, raw in enumerate(lines):
                line = raw.strip()
                if not line:
                    continue
                m = HEAD_INLINE_PAT.match(line)
                if not m:
                    continue

                gd = m.groupdict()
                ref  = gd["ref"].strip()
                desc = gd["desc"].strip()
                qty  = _qty_to_int(gd["qty"])
                unit = _fnum_euro(gd["unit"])
                gross = _fnum_euro(gd["gross"])
                net_s = gd.get("net") or gd.get("net2")
                total = _fnum_euro(net_s) if net_s else gross

                hs = org = ean = ""
                lookahead = lines[i+1:i+6]
                for w in lookahead:
                    if not hs or not org:
                        mh = HS_ORG_PAT.search(w)
                        if mh:
                            hs  = mh.group("hs")
                            org = mh.group("org")
                    if not ean:
                        me = EAN_PAT.search(w)
                        if me:
                            ean = me.group("ean")
                    if hs and org and ean:
                        break

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

# ─────────────────────  EXTRACTOR 5 (COTY, robusto) ─────────────────────
# Soporta:
#  • ES: "Ref. No. / EAN Code Material Ctdad. Precio USD"
#  • EN: "Ref. No. / EAN Code Article Qty Price USD"
#  • Ítems en una sola línea:   ref  ean  desc  qty  unit  total
#  • Ítems en dos líneas:       (cabecera) + (qty unit total)
#  • Totales con * o ** (FOC)
import re

# Cabeceras de tabla y cortes
_COTY_TABLE_HDR = re.compile(
    r"^(Ref\.\s*No\.|Ref\.\s*No\.\s*Customer|EAN\s*Code\s*Article|EAN\s*Code\s*Material)",
    re.I
)
_COTY_END_ROW   = re.compile(r"^(subtotal|total|carry\s*forward)", re.I)

# Una sola línea con todo (ref/ean/desc/qty/unit/total)
_COTY_ONE_LINE = re.compile(
    r"^\s*(?P<ref>\d{8,14})\s+(?P<ean>\d{12,14})\s+(?P<desc>.+?)\s+"
    r"(?P<qty>\d{1,6})\s+(?P<unit>[\d\.,\s]+?)\s+(?P<total>[\d\.,]+)(?:\*+)?\s*$"
)

# Variante en dos líneas: primero solo cabecera ref/ean/desc…
_COTY_HEAD_ONLY = re.compile(
    r"^\s*(?P<ref>\d{8,14})\s+(?P<ean>\d{12,14})\s+(?P<desc>.+?)\s*$"
)

# …y después cantidades/precio/total
_COTY_NUMS = re.compile(
    r"^\s*(?P<qty>\d{1,6})\s+(?P<unit>[\d\.\,\s]+)\s+(?P<total>[\d\.,]+)(?:\*+)?\s*$"
)

# HS y Origen (pueden aparecer entre cabecera y números, o tras la línea completa)
_COTY_HS    = re.compile(r"\(\s*H\s*S\s*No\.?\s*(?P<hs>\d{6,14})\s*\)", re.I)
_COTY_ORGES = re.compile(r"Pa[ií]s\s+de\s+origen:\s*(?P<org>.+)", re.I)
_COTY_ORGEN = re.compile(r"Country\s+of\s+origin:\s*(?P<org>.+)", re.I)

def _coty_num(s: str) -> float:
    if not s:
        return 0.0
    t = s.replace("\u202f", "").replace(" ", "")
    if t.count(",") == 1:
        t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", "")
    try:
        return float(t)
    except:
        return 0.0

def _coty_qty(s: str) -> int:
    return int(s.replace("\u202f","").replace(" ","").replace(".","").replace(",","") or 0)

def extract_coty(pdf_path: str, invoice_number: str) -> List[dict]:
    rows: List[dict] = []
    LOOKAHEAD = 10  # líneas a mirar para HS/Origen después de detectar un ítem

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = [ln.strip() for ln in (page.extract_text(x_tolerance=1.2) or "").split("\n") if ln.strip()]
            in_table = False
            i = 0

            while i < len(lines):
                ln = lines[i]

                # entrar/salir de tabla
                if _COTY_TABLE_HDR.match(ln):
                    in_table = True
                    i += 1
                    continue
                if not in_table:
                    i += 1
                    continue
                if _COTY_END_ROW.match(ln):
                    in_table = False
                    i += 1
                    continue

                # 1) caso "una sola línea"
                m1 = _COTY_ONE_LINE.match(ln)
                if m1:
                    gd = m1.groupdict()
                    # busca HS/Origen en las siguientes N líneas
                    hs = ""
                    org = ""
                    for w in lines[i+1:i+1+LOOKAHEAD]:
                        if not hs:
                            mh = _COTY_HS.search(w)
                            if mh: hs = mh.group("hs")
                        if not org:
                            mo = _COTY_ORGES.search(w) or _COTY_ORGEN.search(w)
                            if mo: org = mo.group("org").strip()
                        if hs and org:
                            break

                    rows.append({
                        "Reference": gd["ref"],
                        "Code EAN": gd["ean"],
                        "Custom Code": hs,
                        "Description": gd["desc"],
                        "Origin": org,
                        "Quantity": _coty_qty(gd["qty"]),
                        "Unit Price": _coty_num(gd["unit"]),
                        "Total Price": _coty_num(gd["total"]),
                        "Invoice Number": invoice_number
                    })
                    i += 1
                    continue

                # 2) caso "dos líneas": cabecera + línea numérica
                mh = _COTY_HEAD_ONLY.match(ln)
                if mh and i + 1 < len(lines) and _COTY_NUMS.match(lines[i+1]):
                    gd = mh.groupdict()
                    mn = _COTY_NUMS.match(lines[i+1])

                    hs = ""
                    org = ""
                    for w in lines[i+2:i+2+LOOKAHEAD]:
                        if not hs:
                            mh2 = _COTY_HS.search(w)
                            if mh2: hs = mh2.group("hs")
                        if not org:
                            mo2 = _COTY_ORGES.search(w) or _COTY_ORGEN.search(w)
                            if mo2: org = mo2.group("org").strip()
                        if hs and org:
                            break

                    rows.append({
                        "Reference": gd["ref"],
                        "Code EAN": gd["ean"],
                        "Custom Code": hs,
                        "Description": gd["desc"],
                        "Origin": org,
                        "Quantity": _coty_qty(mn.group("qty")),
                        "Unit Price": _coty_num(mn.group("unit")),
                        "Total Price": _coty_num(mn.group("total")),
                        "Invoice Number": invoice_number
                    })
                    i += 2
                    continue

                # si no hizo match, sigue
                i += 1

    return rows
# ─────────────────────  EXTRACTOR 6 (Bulgari ASN: Pos/Ref/Q.ty…)  ─────────────────────
ASN_HEAD = re.compile(r"^\s*Pos\.\s*Reference\s*-\s*Cust\.\s*Material", re.I)
ASN_END  = re.compile(r"^(SUBTOTAL|TOTAL)\s*:", re.I)

# Línea numérica (pos ref qty unit hs netw KG unit total)
ASN_NUM = re.compile(
    r"""^\s*
    (?P<pos>\d{1,4})\s+
    (?P<ref>\d{3,})\s+
    (?P<qty>\d{1,6})\s+
    (?P<unit>[A-Z]{2,4})\s+
    (?P<hs>\d{8,10})\s+
    (?P<netw>[\d\.,]+)\s+KG\s+
    (?P<uprice>[\d\.,]+)\s+
    (?P<total>[\d\.,]+)(?:\*+)?\s*$
    """, re.X | re.I
)

ASN_ORIGIN = re.compile(r"^Origin:\s*(?P<org>.+)$", re.I)

def _eu_to_float(s: str) -> float:
    if not s: return 0.0
    t = s.replace("\u202f","").replace(" ","")
    if t.count(",")==1:
        t = t.replace(".","").replace(",",".")
    else:
        t = t.replace(",","")
    try:
        return float(t)
    except:
        return 0.0

def _to_int(s: str) -> int:
    return int(s.replace("\u202f","").replace(" ","").replace(".","").replace(",","") or 0)

def extract_bulgari_asn(pdf_path: str, invoice_number: str) -> List[dict]:
    """
    Lee ítems en bloques de 3 líneas:
      1) numérica: pos ref qty unit hs netw KG unit total
      2) descripción (puede tener guiones)
      3) 'Origin: <pais>'
    """
    rows: List[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = [ln.strip() for ln in (page.extract_text(x_tolerance=1.2) or "").split("\n") if ln.strip()]

            in_table = False
            i = 0
            while i < len(lines):
                ln = lines[i]

                if ASN_HEAD.search(ln):
                    in_table = True
                    i += 1
                    continue
                if not in_table:
                    i += 1
                    continue
                if ASN_END.match(ln):
                    in_table = False
                    i += 1
                    continue

                m = ASN_NUM.match(ln)
                if not m:
                    i += 1
                    continue

                gd = m.groupdict()
                # descripción = siguiente línea no vacía ni encabezado
                desc = ""
                org  = ""
                if i + 1 < len(lines):
                    desc = lines[i+1].strip()
                if i + 2 < len(lines) and ASN_ORIGIN.match(lines[i+2]):
                    org = ASN_ORIGIN.match(lines[i+2]).group("org").strip()
                    step = 3
                else:
                    # a veces "Origin:" podría venir más abajo; busca hasta 3 líneas
                    step = 2
                    for k in range(i+2, min(i+5, len(lines))):
                        mo = ASN_ORIGIN.match(lines[k])
                        if mo:
                            org = mo.group("org").strip()
                            step = (k - i + 1)
                            break

                rows.append({
                    "Reference": gd["ref"],
                    "Code EAN": "",                # este layout no trae EAN
                    "Custom Code": gd["hs"],
                    "Description": desc,
                    "Origin": org,
                    "Quantity": _to_int(gd["qty"]),
                    "Unit Price": _eu_to_float(gd["uprice"]),
                    "Total Price": _eu_to_float(gd["total"]),
                    "Invoice Number": invoice_number
                })

                i += step
    return rows
# ─────────────────────  EXTRACTOR 7 (Interparfums USA: Order Confirmation)  ─────────────────────
IPUSA_HEAD = re.compile(r"^\s*No\.\s*Description", re.I)
IPUSA_END  = re.compile(r"^(Subtotal|Grand\s+Total|\$?\s*Grand\s+Total)", re.I)

# Caso "una sola línea":   ref desc ... IT 3303.00.0000 720 720 Each 26.25 - 18,900.00
IPUSA_ONE = re.compile(
    r"""^\s*
    (?P<ref>[A-Z0-9]{4,})\s+
    (?P<desc>.+?)\s+
    (?P<org>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4}|\d{6,10})\s+
    (?P<qty>[\d,]+)\s+
    (?P<res>[\d,]+)\s+
    (?P<uom>[A-Za-z]+)\s+
    (?P<unit>[\d\.,]+)\s+
    (?P<posm>-|[\d\.,]+)\s+
    (?P<total>[\d\.,]+)\s*$
    """, re.X
)

# Caso "dos líneas": 1) ref + desc (solo)   2) IT HS qty res uom unit posm total
IPUSA_HEAD_ONLY = re.compile(r"^\s*(?P<ref>[A-Z0-9]{4,})\s+(?P<desc>.+?)\s*$")
IPUSA_NUM_LINE  = re.compile(
    r"""^\s*
    (?P<org>[A-Z]{2})\s+
    (?P<hs>\d{4}\.\d{2}\.\d{4}|\d{6,10})\s+
    (?P<qty>[\d,]+)\s+
    (?P<res>[\d,]+)\s+
    (?P<uom>[A-Za-z]+)\s+
    (?P<unit>[\d\.,]+)\s+
    (?P<posm>-|[\d\.,]+)\s+
    (?P<total>[\d\.,]+)\s*$
    """, re.X
)

IPUSA_UPC = re.compile(r"^UPC\s*:\s*(?P<ean>\d{11,14})\s*$", re.I)

def _us_to_float(s: str) -> float:
    if not s: return 0.0
    t = s.replace("\u202f","").replace(" ","")
    # en estos SO suele venir 18,900.00 (coma miles, punto decimal)
    # quitamos comas miles; dejamos punto decimal
    t = t.replace(",", "")
    try:
        return float(t)
    except:
        # fallback estilo europeo
        if t.count(",")==1 and t.count(".")==0:
            try:
                return float(t.replace(",", "."))
            except:
                return 0.0
        return 0.0

def _to_int_clean(s: str) -> int:
    return int(s.replace("\u202f","").replace(" ","").replace(",","").replace(".","") or 0)

def extract_ipusa_order_conf(pdf_path: str, invoice_number: str) -> List[dict]:
    """
    Interparfums USA - Order Confirmation (SO…):
      - Acepta una línea o dos líneas (desc multilínea).
      - Lee UPC en la/s línea/s siguiente/s ("UPC: 0857…") y lo usa como Code EAN.
    """
    rows: List[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # x_tolerance bajo para mantener el orden natural
            lines = [ln.strip() for ln in (page.extract_text(x_tolerance=1.2) or "").split("\n") if ln.strip()]

            in_table = False
            i = 0
            while i < len(lines):
                ln = lines[i]

                if IPUSA_HEAD.match(ln):
                    in_table = True
                    i += 1
                    continue
                if not in_table:
                    i += 1
                    continue
                if IPUSA_END.match(ln):
                    in_table = False
                    i += 1
                    continue

                # 1) Intento "una sola línea"
                m1 = IPUSA_ONE.match(ln)
                if m1:
                    gd = m1.groupdict()
                    ean = ""
                    # buscar UPC hacia adelante (hasta 3 líneas)
                    for k in range(i+1, min(i+4, len(lines))):
                        mu = IPUSA_UPC.match(lines[k])
                        if mu:
                            ean = mu.group("ean")
                            break

                    rows.append({
                        "Reference": gd["ref"],
                        "Code EAN": ean,
                        "Custom Code": gd["hs"].replace(".", ""),
                        "Description": gd["desc"],
                        "Origin": gd["org"],
                        "Quantity": _to_int_clean(gd["qty"]),
                        "Unit Price": _us_to_float(gd["unit"]),
                        "Total Price": _us_to_float(gd["total"]),
                        "Invoice Number": invoice_number
                    })
                    i += 1
                    continue

                # 2) Intento "dos líneas": cabecera (ref+desc) + línea numérica
                mh = IPUSA_HEAD_ONLY.match(ln)
                if mh:
                    ref = mh.group("ref")
                    desc = mh.group("desc")

                    # la numérica puede estar en la siguiente o dos más abajo si la desc se parte
                    j = i + 1
                    # acumula descripción en caso de salto de línea (hasta encontrar numérica o UPC)
                    while j < len(lines) and not IPUSA_NUM_LINE.match(lines[j]) and not IPUSA_UPC.match(lines[j]):
                        # si es una línea tipo "SPRAY" separada, concaténala
                        if not IPUSA_HEAD_ONLY.match(lines[j]):
                            desc = (desc + " " + lines[j]).strip()
                        j += 1

                    if j < len(lines) and IPUSA_NUM_LINE.match(lines[j]):
                        gn = IPUSA_NUM_LINE.match(lines[j]).groupdict()
                        ean = ""
                        # busca UPC en las 3 líneas siguientes a la numérica
                        for k in range(j+1, min(j+4, len(lines))):
                            mu = IPUSA_UPC.match(lines[k])
                            if mu:
                                ean = mu.group("ean")
                                break

                        rows.append({
                            "Reference": ref,
                            "Code EAN": ean,
                            "Custom Code": gn["hs"].replace(".", ""),
                            "Description": desc,
                            "Origin": gn["org"],
                            "Quantity": _to_int_clean(gn["qty"]),
                            "Unit Price": _us_to_float(gn["unit"]),
                            "Total Price": _us_to_float(gn["total"]),
                            "Invoice Number": invoice_number
                        })
                        i = j + 1
                        continue

                # si no hizo match, avanza
                i += 1

    return rows


# ────────────────  COMPLEMENTO: llenar HTS / UPC faltantes  ────────────────
def complete_missing_codes(pdf_path: str, rows: List[dict]) -> None:
    """Rellena in-place cualquier fila sin HTS o UPC."""
    lines=[]
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            txt=pg.extract_text(x_tolerance=1.5) or ""
            lines.extend(txt.split("\n"))
    lines=[re.sub(r"\s{2,}"," ",ln.strip()) for ln in lines if ln.strip()]

    # mapa Reference → índice aproximado (solo layouts con país abreviado)
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

                logging.info("Procesando %s (inv=%s)", pdf.filename, inv_num)

                # 1-5) extraemos con cada estrategia
                rows1=extract_original(tmp.name);                         logging.info("r1=%d", len(rows1))
                rows2=extract_slice(tmp.name,inv_num);                    logging.info("r2=%d", len(rows2))
                rows3=extract_new_provider(tmp.name,inv_num);             logging.info("r3=%d", len(rows3))
                rows4=extract_interparfums_blocks(tmp.name,inv_num);      logging.info("r4=%d", len(rows4))
                rows5=extract_coty(tmp.name, inv_num);                    logging.info("r5=%d", len(rows5))
                rows6=extract_bulgari_asn(tmp.name, inv_num);             logging.info("r6=%d", len(rows6))
                rows7 = extract_ipusa_order_conf(tmp.name, inv_num);        logging.info("r7=%d", len(rows7))

                combo = rows1 + rows2 + rows3 + rows4 + rows5 + rows6 + rows7
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


