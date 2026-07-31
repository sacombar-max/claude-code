import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Uso: python build_consolidador.py [Gamma.xls] [mercados.xlsx] [salida.xlsx]
GAMMA_XLS = sys.argv[1] if len(sys.argv) > 1 else "Gamma.xls"
MERCADOS_XLSX = sys.argv[2] if len(sys.argv) > 2 else "mercados.xlsx"
OUT = sys.argv[3] if len(sys.argv) > 3 else "Consolidador_Facturacion.xlsx"

BRASIL_CAP = 650   # filas de datos disponibles para Brasil (~517 filas/mes observadas)
SUIZA_CAP = 450    # filas de datos disponibles para Suiza (~343 filas/mes observadas)
GAMMA_HEADER_ROW = 2  # fila de encabezados en Gamma_Catalogo (row 1 = titulo)

# ---------- cargar catalogo Gamma ----------
gamma_raw = pd.read_excel(GAMMA_XLS, sheet_name="Gamma", header=7)
gamma = gamma_raw[["IP7", "Prod Description", "R/C", "Category", "Mks description"]].dropna(subset=["IP7"])
gamma = gamma.drop_duplicates(subset=["IP7"]).sort_values("IP7")
gamma["IP7"] = gamma["IP7"].astype(int)
GAMMA_ROWS = len(gamma)

# ---------- cargar tabla Tscode -> Mercado ----------
mercados_raw = pd.read_excel(MERCADOS_XLSX, sheet_name="Sheet2", header=2)
mercados_raw["Market"] = mercados_raw["Market"].ffill()
mercados = mercados_raw.dropna(subset=["Customer Hierarchy 3 Code"])[["Customer Hierarchy 3 Code", "Market"]]
mercados.columns = ["Tscode", "Mercado"]
mercados = mercados.drop_duplicates(subset=["Tscode"]).sort_values("Tscode")
MERCADOS_ROWS = len(mercados)

# Alias de cliente (tal como aparece en la columna 'cliente' del archivo de Suiza) -> Tscode.
# Con esto Suiza no necesita traer el Tscode manualmente: se autocompleta buscando el alias aquí.
# Para un cliente nuevo, basta con agregar su alias y Tscode en la pestaña 'Mercados'.
ALIAS_TSCODE = {
    "PETIN": "TS01696",
    "YAMAHA": "TS01886",
    "ITALCAUCHOS": "TS01916",
    "AUTOCENTRO": "TS01927",
    "CENEU": "TS01928",
    "DESERT": "TS01932",
    "ISN S.": "TS02378",
    "MOTORALMOR": "TS02414",
    "FAHONDA": "TS02459",
    "AKT": "TS02472",
    "MOAUTO": "TS02592",
    "AUTECO": "TS02793",
    "GEES": "TS02815",
    "ROLPARTS": "TS02816",
    "RENDILLANTAS": "TS02839",
    "PRESTIGE": "TS02848",
    "INTEGRANDO": "TS02861",
    "BRIFABRI": "TS02865",
    "FANALCA": "TS03434",
    "MULTILLANTA": "TS06007",
}
alias_by_tscode = {}
for alias_name, ts in ALIAS_TSCODE.items():
    alias_by_tscode.setdefault(ts, []).append(alias_name)

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=13, color="1F4E78")
NOTE_FONT = Font(name=FONT_NAME, italic=True, color="7F7F7F")
EXAMPLE_FILL = PatternFill("solid", fgColor="FFF2CC")
BLUE_INPUT = Font(name=FONT_NAME, color="0000FF")
BLACK = Font(name=FONT_NAME, color="000000")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ================= Instrucciones =================
ws = wb.active
ws.title = "Instrucciones"
ws["A1"] = "Consolidador de Facturación Mensual — Brasil / China / Alemania / Indonesia"
ws["A1"].font = TITLE_FONT
lines = [
    "",
    "Cómo usar este archivo cada fin de mes:",
    "1. Pestaña 'Brasil': borra la fila de ejemplo (fondo amarillo) y pega debajo los datos del mes tal como llegan de Brasil,",
    "   respetando las columnas: Proforma, Tscode, Cliente, Qty, Ipcode, Description, Sales force, Price USD.",
    "2. Pestaña 'Suiza_DE_CN_ID': borra la fila de ejemplo y pega los datos del archivo que envía Suiza (agrupa Alemania,",
    "   China e Indonesia), respetando las columnas: cliente, pf, Ip Code, Description, Brand Line, Quantity, Net Price,",
    "   Amount, Goods Origin. La columna 'Tscode (automático)' NO se pega: se autocompleta sola buscando el nombre de",
    "   'cliente' en la columna Alias de la pestaña 'Mercados'. Si un cliente nuevo no aparece, agrégalo ahí (ver punto 6).",
    "3. La pestaña 'Consolidado' se arma sola con fórmulas. NO está protegida a propósito, para poder armar tablas",
    "   dinámicas y gráficos directamente sobre ella. Por eso mismo no la edites/arrastres/borres filas a mano ahí",
    "   (rompe las fórmulas) — para otra vista, créala en una pestaña nueva. Para borrar la fila de ejemplo en 'Brasil'",
    "   o 'Suiza_DE_CN_ID': selecciona las celdas y presiona Supr/Delete. NO uses clic derecho > 'Eliminar fila'",
    "   (Delete Row), porque eso corre las fórmulas de Consolidado y las rompe.",
    "4. 'mercado' = país del CLIENTE (no el país de fabricación): se busca el Tscode de cada fila en la pestaña 'Mercados'.",
    "   'Goods Origin' sigue siendo el país de fabricación del producto (BR para Brasil, o el código que traiga cada fila de Suiza).",
    "   'Fuente' indica de dónde vino la fila: 'Brasil' (pestaña Brasil, en Excel) o 'Suiza' (pestaña Suiza_DE_CN_ID,",
    "   que agrupa Alemania/China/Indonesia, ya sea desde el Excel de Suiza o desde los PDF de proforma).",
    "5. La pestaña 'Gamma_Catalogo' es la tabla de referencia (catálogo de productos) usada para completar Category, R/C y",
    "   Mks description por Ip Code. Cuando recibas una versión nueva del archivo Gamma, reemplaza estos datos.",
    "6. La pestaña 'Mercados' tiene 3 columnas: Tscode | Alias | Mercado. Si aparece un Tscode nuevo (cliente de Brasil) o",
    "   un cliente nuevo de Suiza sin alias todavía, agrega una fila con su Tscode, su alias (tal como aparece en la columna",
    "   'cliente' de Suiza) y su mercado. Escribe el texto directamente — no copies celdas desde otro archivo Excel abierto,",
    "   porque eso puede pegar un vínculo roto a ese otro archivo en vez del texto.",
    "",
    f"Capacidad actual: {BRASIL_CAP} filas de datos para Brasil y {SUIZA_CAP} filas para Suiza (con margen sobre el volumen",
    "mensual observado: ~517 filas Brasil, ~343 filas Suiza). Si algún mes se supera la capacidad, avisa para ampliar las filas.",
    "",
    "Si en 'Consolidado' una fila muestra 'Revisar Tscode', el Tscode no está en la pestaña 'Mercados' (agrégalo ahí).",
    "Si en 'Suiza_DE_CN_ID' la columna Tscode muestra 'Revisar alias', el nombre de 'cliente' no coincide con ningún Alias",
    "de 'Mercados' (agrega el alias, o revisa que esté escrito igual).",
    "Si muestra 'Revisar Ip Code' en Category/R-C/Mks description, el Ip Code no está en 'Gamma_Catalogo' (código nuevo o catálogo desactualizado).",
]
for i, txt in enumerate(lines, start=2):
    ws.cell(row=i, column=1, value=txt).font = NOTE_FONT if txt else NOTE_FONT
autofit(ws, [130])

# ================= Mercados (Tscode -> Mercado del cliente, + Alias -> Tscode) =================
ws = wb.create_sheet("Mercados")
ws["A1"] = "Tscode"
ws["B1"] = "Alias"
ws["C1"] = "Mercado"
style_header(ws, 1, 3)
r = 2
for _, row in mercados.iterrows():
    ts = row["Tscode"]
    aliases = alias_by_tscode.get(ts, [])
    ws.cell(row=r, column=1, value=ts).font = BLUE_INPUT
    ws.cell(row=r, column=2, value=", ".join(aliases) if aliases else None).font = BLUE_INPUT
    ws.cell(row=r, column=3, value=row["Mercado"]).font = BLUE_INPUT
    r += 1
MERCADOS_LAST_ROW = r - 1
autofit(ws, [14, 18, 24])
ws.freeze_panes = "A2"

# ================= Gamma_Catalogo =================
ws = wb.create_sheet("Gamma_Catalogo")
ws["A1"] = "Catálogo de productos Gamma (referencia — actualizar cuando llegue versión nueva)"
ws["A1"].font = NOTE_FONT
headers = ["Ip Code", "Description", "R/C", "Category", "Mks description"]
for c, h in enumerate(headers, start=1):
    ws.cell(row=GAMMA_HEADER_ROW, column=c, value=h)
style_header(ws, GAMMA_HEADER_ROW, len(headers))
r = GAMMA_HEADER_ROW + 1
for _, row in gamma.iterrows():
    ws.cell(row=r, column=1, value=int(row["IP7"])).font = BLUE_INPUT
    ws.cell(row=r, column=2, value=row["Prod Description"]).font = BLUE_INPUT
    ws.cell(row=r, column=3, value=row["R/C"]).font = BLUE_INPUT
    ws.cell(row=r, column=4, value=row["Category"]).font = BLUE_INPUT
    ws.cell(row=r, column=5, value=row["Mks description"]).font = BLUE_INPUT
    r += 1
GAMMA_LAST_ROW = r - 1
autofit(ws, [12, 32, 8, 14, 26])
ws.freeze_panes = f"A{GAMMA_HEADER_ROW + 1}"

# ================= Brasil =================
ws = wb.create_sheet("Brasil")
b_headers = ["Proforma", "Tscode", "Cliente", "Qty", "Ipcode", "Description", "Sales force", "Price USD"]
for c, h in enumerate(b_headers, start=1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(b_headers))
example = [7530016455, "TS01916", "ITALCAUCHOS C.L. (EJEMPLO - BORRAR)", 19, 1342900,
           "110/70-17M/CTL 54H SpDemF", 11386369, 38.34]
for c, v in enumerate(example, start=1):
    cell = ws.cell(row=2, column=c, value=v)
    cell.fill = EXAMPLE_FILL
    cell.font = BLACK
autofit(ws, [14, 10, 26, 8, 10, 26, 14, 11])
ws.freeze_panes = "A2"
BRASIL_LAST_ROW = 1 + BRASIL_CAP

# ================= Suiza_DE_CN_ID =================
ws = wb.create_sheet("Suiza_DE_CN_ID")
s_headers = ["cliente", "Tscode (automático)", "pf", "Ip Code", "Description", "Brand Line", "Quantity", "Net Price", "Amount", "Goods Origin"]
for c, h in enumerate(s_headers, start=1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(s_headers))
example = ["akt", None, 1730134137, 2283700, "180/55ZR17M/CTL (73W)(M) Z8-R",
           "Moto METZELER", 3, 102, 327.3, "CN"]
for c, v in enumerate(example, start=1):
    cell = ws.cell(row=2, column=c, value=v)
    cell.fill = EXAMPLE_FILL
    cell.font = BLACK
_mercados_tscode = f"Mercados!$A$2:$A${MERCADOS_LAST_ROW}"
_mercados_alias = f"Mercados!$B$2:$B${MERCADOS_LAST_ROW}"
for r in range(2, 2 + SUIZA_CAP):
    a = f"A{r}"
    cell = ws.cell(
        row=r, column=2,
        value=f'=IF({a}="","",IFERROR(INDEX({_mercados_tscode},MATCH({a},{_mercados_alias},0)),"Revisar alias"))'
    )
    if r == 2:
        cell.fill = EXAMPLE_FILL
autofit(ws, [24, 16, 14, 10, 32, 16, 10, 11, 11, 13])
ws.freeze_panes = "A2"
SUIZA_LAST_ROW = 1 + SUIZA_CAP

# ================= Consolidado =================
ws = wb.create_sheet("Consolidado")
c_headers = ["mercado", "cliente", "pf", "Ip Code", "Description", "Quantity", "Net Price",
             "Amount", "Goods Origin", "Category", "R/C", "Mks description", "Fuente"]
for c, h in enumerate(c_headers, start=1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(c_headers))

gcat_code = f"Gamma_Catalogo!$A$3:$A${GAMMA_LAST_ROW}"
gcat_rc = f"Gamma_Catalogo!$C$3:$C${GAMMA_LAST_ROW}"
gcat_cat = f"Gamma_Catalogo!$D$3:$D${GAMMA_LAST_ROW}"
gcat_mks = f"Gamma_Catalogo!$E$3:$E${GAMMA_LAST_ROW}"
tscode_col = f"Mercados!$A$2:$A${MERCADOS_LAST_ROW}"
alias_col = f"Mercados!$B$2:$B${MERCADOS_LAST_ROW}"
mercado_col = f"Mercados!$C$2:$C${MERCADOS_LAST_ROW}"

out_row = 2

# --- bloque Brasil ---
for src_row in range(2, BRASIL_LAST_ROW + 1):
    b = f"Brasil!C{src_row}"  # Cliente, usada como bandera de "hay dato"
    ts = f"Brasil!B{src_row}"
    ip = f"Brasil!E{src_row}"
    ws.cell(row=out_row, column=1,
            value=f'=IF({b}="","",IFERROR(INDEX({mercado_col},MATCH({ts},{tscode_col},0)),"Revisar Tscode"))')
    ws.cell(row=out_row, column=2, value=f'=IF({b}="","",{b})')
    ws.cell(row=out_row, column=3, value=f'=IF({b}="","",Brasil!A{src_row})')
    ws.cell(row=out_row, column=4, value=f'=IF({b}="","",{ip})')
    ws.cell(row=out_row, column=5, value=f'=IF({b}="","",Brasil!F{src_row})')
    ws.cell(row=out_row, column=6, value=f'=IF({b}="","",Brasil!D{src_row})')
    ws.cell(row=out_row, column=7, value=f'=IF({b}="","",Brasil!H{src_row})')
    ws.cell(row=out_row, column=8, value=f'=IF({b}="","",Brasil!D{src_row}*Brasil!H{src_row})')
    ws.cell(row=out_row, column=9, value=f'=IF({b}="","","BR")')
    ws.cell(row=out_row, column=10,
            value=f'=IF({b}="","",IFERROR(INDEX({gcat_cat},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=11,
            value=f'=IF({b}="","",IFERROR(INDEX({gcat_rc},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=12,
            value=f'=IF({b}="","",IFERROR(INDEX({gcat_mks},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=13, value=f'=IF({b}="","","Brasil")')
    out_row += 1

# --- bloque Suiza (Alemania + China + Indonesia) ---
for src_row in range(2, SUIZA_LAST_ROW + 1):
    s = f"Suiza_DE_CN_ID!A{src_row}"  # cliente, usada como bandera de "hay dato"
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
    ws.cell(row=out_row, column=7, value=f'=IF({s}="","",Suiza_DE_CN_ID!H{src_row})')
    ws.cell(row=out_row, column=8, value=f'=IF({s}="","",Suiza_DE_CN_ID!I{src_row})')
    ws.cell(row=out_row, column=9, value=f'=IF({s}="","",{origen})')
    ws.cell(row=out_row, column=10,
            value=f'=IF({s}="","",IFERROR(INDEX({gcat_cat},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=11,
            value=f'=IF({s}="","",IFERROR(INDEX({gcat_rc},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=12,
            value=f'=IF({s}="","",IFERROR(INDEX({gcat_mks},MATCH({ip},{gcat_code},0)),"Revisar Ip Code"))')
    ws.cell(row=out_row, column=13, value=f'=IF({s}="","","Suiza")')
    out_row += 1

autofit(ws, [12, 24, 14, 10, 32, 10, 11, 12, 13, 14, 8, 26, 10])
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:M{out_row - 1}"
# Sin protección: se deja libre para poder armar tablas dinámicas y gráficos desde aquí.

wb.save(OUT)
print("saved", OUT, "gamma rows:", GAMMA_ROWS, "gamma last row:", GAMMA_LAST_ROW,
      "brasil last row:", BRASIL_LAST_ROW, "suiza last row:", SUIZA_LAST_ROW,
      "consolidado rows written:", out_row - 2)
