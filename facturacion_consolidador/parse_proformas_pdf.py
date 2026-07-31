import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber

# Uso: python parse_proformas_pdf.py carpeta_o_archivos.pdf [...] -o salida.xlsx
#
# Lee proformas en PDF (formato "PIRELLI TYRE (SUISSE) SA") y arma una tabla con
# las mismas columnas que la pestaña Suiza_DE_CN_ID:
# cliente, Tscode, pf, Ip Code, Description, Brand Line, Quantity, Net Price, Amount, Goods Origin

COUNTRY_TO_CODE = {
    "INDONESIA": "ID",
    "CHINA": "CN",
    "GERMANY": "DE",
    "ALEMANIA": "DE",
    "BRAZIL": "BR",
    "BRASIL": "BR",
}

LINE_ITEM_RE = re.compile(
    r"^(\S+)\s+(.+?)\s+(\S+\s+\S+)\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([A-Z]{2,3})$"
)


def parse_number(raw):
    return float(raw.replace(".", "").replace(",", "."))


def parse_pdf(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    lines = text.split("\n")

    tscode = cliente = pf = goods_origin = None
    for i, line in enumerate(lines):
        if line.strip() == "Payer" and i + 2 < len(lines):
            tscode = lines[i + 1].strip()
            cliente = lines[i + 2].strip()
        m = re.search(r"Dispatching country\s+(\S+)", line)
        if m:
            goods_origin = COUNTRY_TO_CODE.get(m.group(1).upper(), m.group(1).upper())
        m = re.search(r"Order N°\s+(\d+)", line)
        if m:
            pf = m.group(1)

    if not (tscode and cliente and pf):
        raise ValueError(f"{path}: no se pudo leer Tscode/cliente/pf del encabezado")

    rows = []
    for line in lines:
        m = LINE_ITEM_RE.match(line.strip())
        if not m:
            continue
        ip_code, description, brand_line, qty, net_price, amount, origin = m.groups()
        rows.append({
            "cliente": cliente,
            "Tscode": tscode,
            "pf": pf,
            "Ip Code": ip_code,
            "Description": description,
            "Brand Line": brand_line,
            "Quantity": int(qty),
            "Net Price": parse_number(net_price),
            "Amount": parse_number(amount),
            "Goods Origin": origin if origin else goods_origin,
        })

    if not rows:
        raise ValueError(f"{path}: no se encontraron líneas de producto")
    return rows


def main():
    args = sys.argv[1:]
    if "-o" in args:
        idx = args.index("-o")
        out_path = args[idx + 1]
        args = args[:idx] + args[idx + 2:]
    else:
        out_path = "suiza_desde_pdf.xlsx"

    pdf_paths = []
    for a in args:
        p = Path(a)
        if p.is_dir():
            pdf_paths.extend(sorted(p.glob("*.pdf")))
        else:
            pdf_paths.append(p)

    if not pdf_paths:
        print("Uso: python parse_proformas_pdf.py carpeta_o_archivos.pdf [...] -o salida.xlsx")
        sys.exit(1)

    all_rows = []
    for p in pdf_paths:
        all_rows.extend(parse_pdf(p))

    df = pd.DataFrame(all_rows)
    df.to_excel(out_path, index=False)
    print(f"saved {out_path}: {len(df)} filas de {len(pdf_paths)} PDF(s)")


if __name__ == "__main__":
    main()
