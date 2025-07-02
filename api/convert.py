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

##############################################################################
# 1.  EXTRACTOR ORIGINAL  (tus regex para facturas / proformas clásicas)     #
##############################################################################
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
        all_txt = "\n".join(p.extract_text() or "" for p in pdf.pages)
        kind = doc_kind(all_txt)

        inv = ""
        if kind == "factura":
            if m := INV_PAT.search(all_txt): inv = m.group(1)
            plv = bool(PLV_PAT.search(all_txt))
        else:
            if m := PROF_PAT.search(all_txt):  inv = m.group(1)
            elif m := ORDER_PAT_EN.search(all_txt): inv = m.group(1)
            elif m := ORDER_PAT_FR.search(all_txt): inv = m.group(1)
            plv = False
        inv_full = inv + ("PLV" if plv else "")

        org_global = ""
        for page in pdf.pages:
            txt = page.extract_text() or ""
            for ln in txt.splitlines():
                if mo := ORG_PAT.search(ln):
                    org_global = mo.group(1).strip() or org_global

            lines = txt.split("\n")
            for i, raw in enumerate(lines):
                ln = raw.strip()
                if kind == "factura" and (m := ROW_FACT.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = m.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) and not ROW_FACT.match(lines[i+1]) else ""
                    rows.append(
                        dict(Reference=ref, Code_EAN=ean, Custom_Code=custom,
                             Description=desc, Origin=org_global,
                             Quantity=int(qty_s.replace(".", "").replace(",", "")),
                             Unit_Price=fnum(unit_s), Total_Price=fnum(tot_s),
                             Invoice_Number=inv_full)
                    )
                elif kind == "proforma" and (m := ROW_PROF_DIOR.match(ln)):
                    ref, ean, custom, qty_s, unit_s, tot_s = m.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    rows.append(
                        dict(Reference=ref, Code_EAN=ean, Custom_Code=custom,
                             Description=desc, Origin=org_global,
                             Quantity=int(qty_s.replace(".", "").replace(",", "")),
                             Unit_Price=fnum(unit_s), Total_Price=fnum(tot_s),
                             Invoice_Number=inv_full)
                    )
                elif kind == "proforma" and (m := ROW_PROF.match(ln)):
                    ref, ean, unit_s, qty_s = m.groups()
                    desc = lines[i+1].strip() if i+1 < len(lines) else ""
                    qty  = int(qty_s.replace(".", "").replace(",", ""))
                    unit = fnum(unit_s)
                    rows.append(
                        dict(Reference=ref, Code_EAN=ean, Custom_Code="",
                             Description=desc, Origin=org_global,
                             Quantity=qty, Unit_Price=unit, Total_Price=unit*qty,
                             Invoice_Number=inv_full)
                    )

    inv2org = defaultdict(set)
    for r in rows:
        if r["Origin"]:
            inv2org[r["Invoice_Number"]].add(r["Origin"])
    for r in rows:
        if not r["Origin"] and len(inv2org[r["Invoice_Number"]]) == 1:
            r["Origin"] = next(iter(inv2org[r["Invoice_Number"]]))
    return rows


##############################################################################
# 2.  EXTRACTOR RÁPIDO  (layout "No. Description UPC ...")                   #
##############################################################################
FAST_RE = re.compile(
    r"^(?P<ref>\d{5,6}[A-Z]?)\s+"
    r"(?P<desc>.+?)\s+"
    r"(?P<upc>\d{12,14})\s+"
    r"(?P<ctry>[A-Z]{2,3})\s+"
    r"(?P<hs>\d{4}\.\d{2}\.\d{4})\s+"
    r"(?P<qty>[\d.,]+)\s+Each\s+"
    r"(?P<unit>[\d.,]+)\s+(?:-|\d[\d.,]+)?\s+"
    r"(?P<total>[\d.,]+)$"
)

def to_float(txt: str) -> float:
    txt = txt.replace(",", "").replace(" ", "")
    return float(txt) if txt else 0.0

def fast_extract(pdf_path: str, invoice: str) -> List[dict]:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                if m := FAST_RE.match(line.strip()):
                    d = m.groupdict()
                    rows.append(
                        dict(
                            Reference=d["ref"],
                            Code_EAN=d["upc"],
                            Custom_Code=d["hs"],
                            Description=d["desc"],
                            Origin=d["ctry"],
                            Quantity=int(float(d["qty"].replace(",", ""))),
                            Unit_Price=to_float(d["unit"]),
                            Total_Price=to_float(d["total"]),
                            Invoice_Number=invoice,
                        )
                    )
    return rows

##############################################################################
# 3.  FLASK ENDPOINT                                                         #
##############################################################################
OUTPUT_ORDER = ["Reference","Code_EAN","Custom_Code","Description","Origin",
                "Quantity","Unit_Price","Total_Price","Invoice_Number"]

@app.post("/api/convert")
@app.post("/")
def convert():
    try:
        files = request.files.getlist("file")
        if not files:
            return "No file(s) uploaded", 400

        out_rows = []
        for f in files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                f.save(tmp.name)
                inv = re.search(r"SIP(\d+)", f.filename or "")
                inv = inv.group(1) if inv else ""

                # ① Original extractor
                rows = extract_original(tmp.name)

                # ② If nothing found → fast extractor
                if not rows:
                    rows = fast_extract(tmp.name, inv)

                out_rows.extend(rows)
                os.unlink(tmp.name)

        if not out_rows:
            return "Sin registros extraídos", 400

        wb = Workbook(); ws = wb.active; ws.append(OUTPUT_ORDER)
        for r in out_rows:
            ws.append([r[c] for c in OUTPUT_ORDER])

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name="extracted_data.xlsx",
                         mimetype=("application/vnd.openxmlformats-"
                                   "officedocument.spreadsheetml.sheet"))
    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")

