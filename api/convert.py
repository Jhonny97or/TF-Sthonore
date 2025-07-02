import logging
import os
import re
import tempfile
import traceback
from io import BytesIO
from typing import Dict, List

import pdfplumber
from flask import Flask, request, send_file
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")
app = Flask(__name__)

# ─────────────────────────  AJUSTA SOLO ESTO SI CAMBIA TU PDF  ───────────────
COL_BOUNDS: Dict[str, tuple] = {
    "ref":   (  0,  70),
    "desc":  ( 70, 340),
    "upc":   (340, 430),
    "ctry":  (430, 465),
    "hs":    (465, 535),
    "qty":   (535, 585),
    "unit":  (585, 635),
    "total": (635, 725),
}
# ─────────────────────────────────────────────────────────────────────────────

OUTPUT_COLS = [
    "Reference", "Code EAN", "Custom Code", "Description", "Origin",
    "Quantity", "Unit Price", "Total Price", "Invoice Number",
]

REF_PAT  = re.compile(r"^[A-Z0-9]{3,}$")          # ahora acepta alfa-num
UPC_PAT  = re.compile(r"^\d{11,14}$")
HTS_PAT  = re.compile(r"^\d{6,10}$")
NUM_PAT  = re.compile(r"[0-9]")

SKIP_SNIPPETS = {
    "No. Description", "Total before", "Bill To Ship", "CIF CHILE",
    "Invoice", "Ship From", "Ship To", "VAT/Tax", "Shipping Te"
}

def clean(t: str) -> str:
    return t.replace("\u202f", " ").strip()

def to_float(t: str) -> float:
    t = t.replace("\u202f", "").replace(" ", "")
    if t.count(",") == 1 and t.count(".") == 0:
        t = t.replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    return float(t or 0)

def to_int(t: str) -> int:
    return int(t.replace(",", "").replace(".", "") or 0)

# ─────────────────────────  PARSER ORIGINAL POR COORDENADAS  ────────────────
def rows_from_page(page) -> List[Dict[str, str]]:
    rows = []
    grouped: Dict[float, List[dict]] = {}
    for ch in page.chars:
        y = round(ch["top"], 1)
        grouped.setdefault(y, []).append(ch)

    for _, chs in sorted(grouped.items()):
        text = "".join(c["text"] for c in sorted(chs, key=lambda c: c["x0"]))
        if not text.strip() or any(s in text for s in SKIP_SNIPPETS):
            continue

        cols = {k: "" for k in COL_BOUNDS}
        for ch in sorted(chs, key=lambda c: c["x0"]):
            xm = (ch["x0"] + ch["x1"]) / 2
            for key, (x0, x1) in COL_BOUNDS.items():
                if x0 <= xm < x1:
                    cols[key] += ch["text"]; break
        cols = {k: clean(v) for k, v in cols.items()}

        # ensamblamos filas completas
        if not cols["ref"]:
            if rows: rows[-1]["desc"] += " " + cols["desc"]
            continue
        if not REF_PAT.match(cols["ref"]) or not UPC_PAT.match(cols["upc"]):
            # si UPC falta, de todos modos guardamos y luego completaremos
            pass
        if not NUM_PAT.search(cols["qty"]):
            continue
        rows.append(cols)
    return rows

# ──────────────────  COMPLETAR HTS / UPC FALTANTES POR TEXTO  ───────────────
def _complete_codes(pdf_path: str, registros: List[Dict[str, str]]) -> None:
    """
    En cada registro sin HTS o UPC busca en el texto del PDF la línea con la
    referencia y extrae el par HTS / UPC más cercano.
    Actúa IN-PLACE sobre la lista 'registros'.
    """
    # 1) indexamos líneas de texto
    lines: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            txt = pg.extract_text(x_tolerance=1.5)
            if txt:
                lines.extend(txt.split("\n"))
    lines = [re.sub(r"\s{2,}", " ", l.strip()) for l in lines if l.strip()]

    # 2) mapa Reference → índice de línea
    idx_map: Dict[str, int] = {}
    for i, ln in enumerate(lines):
        m = re.match(r"^([A-Z0-9]{3,})\s+[A-Z]{3}\s", ln)
        if m:
            idx_map.setdefault(m.group(1), i)

    # 3) completamos cada registro pendiente
    for reg in registros:
        if reg["Custom Code"] and reg["Code EAN"]:
            continue  # ya completo

        ref = reg["Reference"]
        start = idx_map.get(ref)
        if start is None:
            continue

        # miramos hasta 20 líneas después (o hasta próximo artículo)
        end = start + 1
        while end < len(lines) and end - start < 20:
            if re.match(r"^[A-Z0-9]{3,}\s+[A-Z]{3}\s", lines[end]):
                break
            end += 1

        snippet = " ".join(lines[start:end])
        # buscamos todas las secuencias numéricas de 6-14 dígitos
        seqs = re.findall(r"\d{6,14}", snippet)
        hts_candidates = [s for s in seqs if HTS_PAT.fullmatch(s)]
        upc_candidates = [s for s in seqs if UPC_PAT.fullmatch(s)]

        # Elegimos la 1ª HTS y la 1ª UPC encontradas
        if hts_candidates and not reg["Custom Code"]:
            reg["Custom Code"] = hts_candidates[0]
        if upc_candidates and not reg["Code EAN"]:
            reg["Code EAN"] = upc_candidates[0]

# ─────────────────────────────  ENDPOINT FLASK  ──────────────────────────────
@app.route("/", methods=["POST"])
@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        pdfs = request.files.getlist("file")
        if not pdfs:
            return "No file(s) uploaded", 400

        extracted: List[Dict[str, str]] = []

        # procesamos cada PDF
        for f in pdfs:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                f.save(tmp.name)

            # Nº de factura como pista
            invoice_hint = os.path.splitext(os.path.basename(f.filename))[0]

            pdf_rows: List[Dict[str, str]] = []
            with pdfplumber.open(tmp.name) as pdf:
                for p in pdf.pages:
                    pdf_rows.extend(rows_from_page(p))

            # transformamos a formato final
            for r in pdf_rows:
                extracted.append(
                    {
                        "Reference": r["ref"],
                        "Code EAN": r["upc"],
                        "Custom Code": r["hs"],
                        "Description": r["desc"],
                        "Origin": r["ctry"],
                        "Quantity": to_int(r["qty"]),
                        "Unit Price": to_float(r["unit"]),
                        "Total Price": to_float(r["total"]),
                        "Invoice Number": invoice_hint,
                    }
                )

            # completamos códigos que falten (HTS / UPC)
            _complete_codes(tmp.name, extracted)

            os.unlink(tmp.name)

        if not extracted:
            return "Sin registros extraídos", 400

        # ─── Exportamos a Excel ────────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.append(OUTPUT_COLS)
        for r in extracted:
            ws.append([r[c] for c in OUTPUT_COLS])

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name="extracted_data.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    except Exception:
        logging.exception("Error en /convert")
        return f"<pre>{traceback.format_exc()}</pre>", 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")


