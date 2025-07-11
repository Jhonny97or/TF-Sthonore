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

# ────────────────────────────── CONFIG GLOBAL ─────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

# Regex mejorada para capturar "V/CDE… ORD Nr" o "V/CDE… ORDER Nr"
ORD_NAME_PAT = re.compile(
    r"V\/CDE[^\n]*?ORD(?:ER)?\s*Nr\s*[:\-]\s*(.+)", re.I
)

# Columnas de salida (dejamos sólo "Order Name")
COLS = [
    "Reference", "Code EAN", "Custom Code", "Description",
    "Origin", "Quantity", "Unit Price", "POSM FOC", "Line Amount",
    "Invoice Number", "Order Name"
]

# ───────────────────── EXTRACTOR 1 (original) ─────────────────────
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
            for ln in lines:
                if mo := ORG_PAT.search(ln):
                    val = mo.group(1).strip()
                    if val:
                        org_global = val

            for i, raw in enumerate(lines):
                ln = raw.strip()
                if kind == "factura" and (mf := ROW_FACT.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = mf.groups()
                    desc = (
                        lines[i+1].strip()
                        if i+1 < len(lines) and not ROW_FACT.match(lines[i+1])
                        else ""
                    )
                    rows.append({
                        "Reference":     ref,
                        "Code EAN":      ean,
                        "Custom Code":   custom,
                        "Description":   desc,
                        "Origin":        org_global,
                        "Quantity":      int(qty_s.replace(".", "").replace(",", "")),
                        "Unit Price":    fnum(unit_s),
                        "POSM FOC":      "",
                        "Line Amount":   fnum(tot_s),
                        "Invoice Number": invoice_full,
                    })
                elif kind == "proforma" and (mpd := ROW_PROF_DIOR.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = mpd.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append({
                        "Reference":     ref,
                        "Code EAN":      ean,
                        "Custom Code":   custom,
                        "Description":   desc,
                        "Origin":        org_global,
                        "Quantity":      int(qty_s.replace(".", "").replace(",", "")),
                        "Unit Price":    fnum(unit_s),
                        "POSM FOC":      "",
                        "Line Amount":   fnum(tot_s),
                        "Invoice Number": invoice_full,
                    })
                elif kind == "proforma" and (mp := ROW_PROF.match(ln)):
                    ref, ean, unit_s, qty_s = mp.groups()
                    qty = int(qty_s.replace(".", "").replace(",", ""))
                    unit = fnum(unit_s)
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append({
                        "Reference":     ref,
                        "Code EAN":      ean,
                        "Custom Code":   "",
                        "Description":   desc,
                        "Origin":        org_global,
                        "Quantity":      qty,
                        "Unit Price":    unit,
                        "POSM FOC":      "",
                        "Line Amount":   unit * qty,
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

# ───────────────────── EXTRACTOR 2 (slice) ─────────────────────
COL_BOUNDS: Dict[str, tuple] = {
    "ref":   (  0,  70),
    "desc":  ( 70, 340),
    "upc":   (340, 430),
    "ctry":  (430, 465),
    "hs":    (465, 535),
    "qty":   (535, 585),
    "unit":  (585, 635),
    "posm":  (635, 675),
    "total": (675, 755),
}
REF_PAT = re.compile(r"^\d{5,6}[A-Z]?$")
UPC_PAT = re.compile(r"^\d{12,14}$")
NUM_PAT = re.compile(r"[0-9]")
SKIP_SNIPPETS = {
    "No. Description","Total before","Bill To Ship","CIF CHILE",
    "Invoice","Ship From","Ship To","VAT/Tax","Shipping Te"
}

def clean(txt: str) -> str:
    return txt.replace("\u202f", " ").strip()

def to_float2(txt: str) -> float:
    t = txt.replace("\u202f", "").replace(" ", "")
    if t.count(",") == 1 and t.count(".") == 0:
        t = t.replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    return float(t or 0)

def to_int2(txt: str) -> int:
    return int(txt.replace(",", "").replace(".", "") or 0)

def rows_from_page(page) -> List[Dict[str, str]]:
    rows = []
    grouped: Dict[float, List[dict]] = {}
    for ch in page.chars:
        grouped.setdefault(round(ch["top"], 1), []).append(ch)

    for _, chs in sorted(grouped.items()):
        line_txt = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if not line_txt.strip() or any(sn in line_txt for sn in SKIP_SNIPPETS):
            continue

        cols = {k: "" for k in COL_BOUNDS}
        for ch in sorted(chs, key=lambda c: c["x0"]):
            xm = (ch["x0"] + ch["x1"]) / 2
            for key, (x0, x1) in COL_BOUNDS.items():
                if x0 <= xm < x1:
                    cols[key] += ch["text"]
                    break
        cols = {k: clean(v) for k, v in cols.items()}

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

def extract_slice(pdf_path: str, inv_number: str) -> List[dict]:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for r in rows_from_page(page):
                rows.append({
                    "Reference":     r["ref"],
                    "Code EAN":      r["upc"],
                    "Custom Code":   r["hs"],
                    "Description":   r["desc"],
                    "Origin":        r["ctry"],
                    "Quantity":      to_int2(r["qty"]),
                    "Unit Price":    to_float2(r["unit"]),
                    "POSM FOC":      to_float2(r["posm"]),
                    "Line Amount":   to_float2(r["total"]),
                    "Invoice Number": inv_number,
                })
    return rows

# ───────────────────── EXTRACTOR 3 (nuevo proveedor) ─────────────────────
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
        txt = s.replace(",", "")
        return float(txt) if txt else 0.0

    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if "No. Description" not in txt:
                continue

            pending_desc = None
            for ln in txt.split("\n"):
                ln_s = ln.strip()
                if not ln_s or ln_s.startswith(("No. Description", "Invoice")):
                    continue

                m = pattern_full.match(ln)
                if m:
                    d = m.groupdict()
                    rows.append({
                        "Reference":     d["ref"],
                        "Code EAN":      d["upc"],
                        "Custom Code":   d["hs"],
                        "Description":   d["desc"].strip(),
                        "Origin":        d["ctry"],
                        "Quantity":      int(d["qty"].replace(",", "")),
                        "Unit Price":    new_fnum(d["unit"]),
                        "POSM FOC":      new_fnum(d.get("posm", "")),
                        "Line Amount":   new_fnum(d["total"]),
                        "Invoice Number": inv_number,
                    })
                    pending_desc = None
                    continue

                mb = pattern_basic.match(ln)
                if mb and pending_desc is not None:
                    d = mb.groupdict()
                    rows.append({
                        "Reference":     d["ref"],
                        "Code EAN":      d["upc"],
                        "Custom Code":   d["hs"],
                        "Description":   pending_desc.strip(),
                        "Origin":        d["ctry"],
                        "Quantity":      int(d["qty"].replace(",", "")),
                        "Unit Price":    new_fnum(d["unit"]),
                        "POSM FOC":      new_fnum(d.get("posm", "")),
                        "Line Amount":   new_fnum(d["total"]),
                        "Invoice Number": inv_number,
                    })
                    pending_desc = None
                    continue

                if re.search(r"[A-Za-z]", ln_s):
                    skip = (
                        "Country of","Customer PO","Order No","Shipping Terms",
                        "Bill To","Finance","Total","CIF","Ship To"
                    )
                    if not any(ln_s.startswith(p) for p in skip):
                        pending_desc = (pending_desc + " " + ln_s) if pending_desc else ln_s

    return rows

# ───────────────────────────── ENDPOINT ─────────────────────────────
@app.post("/api/convert")
@app.post("/")
def convert():
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded", 400

        all_rows: List[dict] = []
        for pdf in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf.save(tmp.name)
                inv_match = re.search(r"SIP(\d+)", pdf.filename or "")
                inv_num = inv_match.group(1) if inv_match else ""

                # extraer sólo Order Name
                order_name = ""
                with pdfplumber.open(tmp.name) as pdf_for_name:
                    txt_all = "\n".join(page.extract_text() or "" for page in pdf_for_name.pages)
                    if m := ORD_NAME_PAT.search(txt_all):
                        order_name = m.group(1).strip()

                o1 = extract_original(tmp.name)
                o2 = extract_slice(tmp.name, inv_num)
                o3 = extract_new_provider(tmp.name, inv_num)

                combined = o1 + o2 + o3
                seen = set()
                unique = []
                for r in combined:
                    key = (r["Reference"], r["Code EAN"], r["Invoice Number"])
                    if key not in seen:
                        seen.add(key)
                        r["Order Name"] = order_name
                        unique.append(r)

                all_rows.extend(unique)
            os.unlink(tmp.name)

        if not all_rows:
            return "Sin registros extraídos", 400

        wb = Workbook()
        ws = wb.active
        ws.append(COLS)
        for r in all_rows:
            ws.append([r.get(c, "") for c in COLS])

        buf = BytesIO()
        wb.save(buf); buf.seek(0)
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


