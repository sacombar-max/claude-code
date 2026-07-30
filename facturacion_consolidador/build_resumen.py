import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Uso: python build_resumen.py Consolidador_Facturacion.xlsx
# Requiere que el libro ya tenga Brasil/Suiza_DE_CN_ID con datos y haya sido
# recalculado (recalc.py) al menos una vez, para poder leer los valores de Consolidado.
IN_OUT = sys.argv[1] if len(sys.argv) > 1 else "Consolidador_Facturacion.xlsx"

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


df_full = pd.read_excel(IN_OUT, sheet_name="Consolidado", header=0)
CONSOL_LAST_ROW = 1 + len(df_full)
df = df_full[df_full["cliente"].notna()]

grouped = (
    df.groupby(["mercado", "Goods Origin", "Category", "R/C", "Mks description"], dropna=False)
    .size()
    .reset_index()[["mercado", "Goods Origin", "Category", "R/C", "Mks description"]]
    .sort_values(["mercado", "Goods Origin", "Category", "R/C", "Mks description"])
)

wb = load_workbook(IN_OUT)
if "Resumen" in wb.sheetnames:
    del wb["Resumen"]
ws = wb.create_sheet("Resumen")

headers = ["mercado", "Quantity", "Amount", "Goods Origin", "Category", "R/C", "Mks description"]
for c, h in enumerate(headers, start=1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))

def col_rng(col):
    return f"Consolidado!${col}$2:${col}${CONSOL_LAST_ROW}"

mercado_rng = col_rng("A")
qty_rng = col_rng("F")
amount_rng = col_rng("H")
origen_rng = col_rng("I")
cat_rng = col_rng("J")
rc_rng = col_rng("K")
mks_rng = col_rng("L")

r = 2
for _, row in grouped.iterrows():
    ws.cell(row=r, column=1, value=row["mercado"])
    ws.cell(row=r, column=4, value=row["Goods Origin"])
    ws.cell(row=r, column=5, value=row["Category"])
    ws.cell(row=r, column=6, value=row["R/C"])
    ws.cell(row=r, column=7, value=row["Mks description"])
    ws.cell(
        row=r, column=2,
        value=f'=SUMIFS({qty_rng},{mercado_rng},A{r},{origen_rng},D{r},{cat_rng},E{r},{rc_rng},F{r},{mks_rng},G{r})'
    )
    ws.cell(
        row=r, column=3,
        value=f'=SUMIFS({amount_rng},{mercado_rng},A{r},{origen_rng},D{r},{cat_rng},E{r},{rc_rng},F{r},{mks_rng},G{r})'
    )
    r += 1
last_row = r - 1

total_row = r
ws.cell(row=total_row, column=1, value="Total").font = Font(name=FONT_NAME, bold=True)
ws.cell(row=total_row, column=2, value=f"=SUM(B2:B{last_row})").font = Font(name=FONT_NAME, bold=True)
ws.cell(row=total_row, column=3, value=f"=SUM(C2:C{last_row})").font = Font(name=FONT_NAME, bold=True)

for i, w in enumerate([16, 11, 13, 13, 14, 8, 26], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:G{last_row}"

wb.save(IN_OUT)
print("saved", IN_OUT, "resumen rows:", last_row - 1)
