import sys

from openpyxl import load_workbook

# Uso: python resize_capacity.py Consolidador_Facturacion.xlsx nuevo_brasil_cap nuevo_suiza_cap
#
# Regenera las formulas de Suiza_DE_CN_ID (columna Tscode automatico) y de Consolidado
# (las 15 columnas, bloque Brasil + bloque Suiza) para una capacidad mayor, usando el
# Gamma_Catalogo y la tabla Mercados que YA existen en el archivo (no hace falta el
# Gamma.xls ni el mercados.xlsx originales). Pensado para correrlo sobre la plantilla
# base (en blanco), no sobre un archivo ya lleno de datos del mes.

WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "Consolidador_Facturacion.xlsx"
BRASIL_CAP = int(sys.argv[2]) if len(sys.argv) > 2 else 900
SUIZA_CAP = int(sys.argv[3]) if len(sys.argv) > 3 else 800

wb = load_workbook(WORKBOOK)

GAMMA_LAST_ROW = wb["Gamma_Catalogo"].max_row
MERCADOS_LAST_ROW = wb["Mercados"].max_row

gcat_code = f"Gamma_Catalogo!$A$3:$A${GAMMA_LAST_ROW}"
gcat_rc = f"Gamma_Catalogo!$C$3:$C${GAMMA_LAST_ROW}"
gcat_cat = f"Gamma_Catalogo!$D$3:$D${GAMMA_LAST_ROW}"
gcat_mks = f"Gamma_Catalogo!$E$3:$E${GAMMA_LAST_ROW}"
tscode_col = f"Mercados!$A$2:$A${MERCADOS_LAST_ROW}"
alias_col = f"Mercados!$B$2:$B${MERCADOS_LAST_ROW}"
mercado_col = f"Mercados!$C$2:$C${MERCADOS_LAST_ROW}"

BRASIL_LAST_ROW = 1 + BRASIL_CAP
SUIZA_LAST_ROW = 1 + SUIZA_CAP

# ---- Suiza_DE_CN_ID: columna Tscode (automatico) ----
ws = wb["Suiza_DE_CN_ID"]
for r in range(2, SUIZA_LAST_ROW + 1):
    a = f"A{r}"
    ws.cell(
        row=r, column=2,
        value=f'=IF({a}="","",IFERROR(INDEX({tscode_col},MATCH({a},{alias_col},0)),"Revisar alias"))'
    )

# ---- Consolidado: regenerar todo el bloque de formulas ----
ws = wb["Consolidado"]
out_row = 2

for src_row in range(2, BRASIL_LAST_ROW + 1):
    b = f"Brasil!C{src_row}"
    ts = f"Brasil!B{src_row}"
    ip = f"Brasil!E{src_row}"
    ws.cell(row=out_row, column=1,
            value=f'=IF({b}="","",IFERROR(INDEX({mercado_col},MATCH({ts},{tscode_col},0)),"Revisar Tscode"))')
    ws.cell(row=out_row, column=2, value=f'=IF({b}="","",{b})')
    ws.cell(row=out_row, column=3, value=f'=IF({b}="","",Brasil!A{src_row})')
    ws.cell(row=out_row, column=4, value=f'=IF({b}="","",{ip})')
    ws.cell(row=out_row, column=5, value=f'=IF({b}="","",Brasil!F{src_row})')
    ws.cell(row=out_row, column=6, value=f'=IF({b}="","",Brasil!D{src_row})')
    ws.cell(row=out_row, column=9, value=f'=IF({b}="","",Brasil!H{src_row})')
    ws.cell(row=out_row, column=10, value=f'=IF({b}="","",Brasil!D{src_row}*Brasil!H{src_row})')
    ws.cell(row=out_row, column=11, value=f'=IF({b}="","","BR")')
    ws.cell(row=out_row, column=12,
            value=f'=IF({b}="","",IFERROR(INDEX({gcat_cat},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=13,
            value=f'=IF({b}="","",IFERROR(INDEX({gcat_rc},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=14,
            value=f'=IF({b}="","",IFERROR(INDEX({gcat_mks},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=15, value=f'=IF({b}="","","Brasil")')
    out_row += 1

for src_row in range(2, SUIZA_LAST_ROW + 1):
    s = f"Suiza_DE_CN_ID!A{src_row}"
    ts = f"Suiza_DE_CN_ID!B{src_row}"
    ip = f"Suiza_DE_CN_ID!D{src_row}"
    origen = f"UPPER(Suiza_DE_CN_ID!J{src_row})"
    ws.cell(row=out_row, column=1,
            value=f'=IF({s}="","",IFERROR(INDEX({mercado_col},MATCH({ts},{tscode_col},0)),"Revisar Tscode"))')
    ws.cell(row=out_row, column=2, value=f'=IF({s}="","",{s})')
    ws.cell(row=out_row, column=3, value=f'=IF({s}="","",Suiza_DE_CN_ID!C{src_row})')
    ws.cell(row=out_row, column=4, value=f'=IF({s}="","",{ip})')
    ws.cell(row=out_row, column=5, value=f'=IF({s}="","",Suiza_DE_CN_ID!E{src_row})')
    ws.cell(row=out_row, column=6, value=f'=IF({s}="","",Suiza_DE_CN_ID!G{src_row})')
    ws.cell(row=out_row, column=9, value=f'=IF({s}="","",Suiza_DE_CN_ID!H{src_row})')
    ws.cell(row=out_row, column=10, value=f'=IF({s}="","",Suiza_DE_CN_ID!I{src_row})')
    ws.cell(row=out_row, column=11, value=f'=IF({s}="","",{origen})')
    ws.cell(row=out_row, column=12,
            value=f'=IF({s}="","",IFERROR(INDEX({gcat_cat},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=13,
            value=f'=IF({s}="","",IFERROR(INDEX({gcat_rc},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=14,
            value=f'=IF({s}="","",IFERROR(INDEX({gcat_mks},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=15, value=f'=IF({s}="","","Suiza")')
    out_row += 1

ws.auto_filter.ref = f"A1:O{out_row - 1}"

wb.save(WORKBOOK)
print(f"listo: capacidad ampliada a Brasil={BRASIL_CAP}, Suiza={SUIZA_CAP} filas "
      f"(Consolidado hasta la fila {out_row - 1})")
