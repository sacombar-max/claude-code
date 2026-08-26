# Consolidador de Facturación Multipaís

Libro Excel (`Consolidador_Facturacion.xlsx`) para consolidar la facturación mensual de
Brasil, China, Alemania e Indonesia en un solo formato:

`mercado | cliente | pf | Ip Code | Description | Quantity | FAC | SEP | Net Price | Amount | Goods Origin | Category | R/C | Mks description | Fuente`

`mercado` es el país del **cliente** (se busca por Tscode en la pestaña `Mercados`), no el país
de fabricación — ese es `Goods Origin`. `Fuente` indica de dónde vino la fila: `Brasil` (pestaña
Brasil) o `Suiza` (pestaña Suiza_DE_CN_ID, ya sea desde el Excel de Suiza o desde PDFs de proforma).
`FAC` (piezas ya facturadas este mes) y `SEP` (piezas que quedan para el mes siguiente) son las
**únicas** columnas de Consolidado pensadas para llenarse a mano, fila por fila — todas las demás
son fórmula y no deben tocarse.

La pestaña **Mercados** tiene 3 columnas: `Tscode | Alias | Mercado`. Brasil ya trae su propio
Tscode; Suiza no, así que su Tscode se **autocompleta por fórmula** buscando el nombre de
`cliente` contra la columna `Alias`.

## Uso mensual

1. Pega los datos de Brasil en la pestaña **Brasil** (reemplaza la fila de ejemplo).
2. Pega los datos del archivo de Suiza (agrupa Alemania, China e Indonesia) en la
   pestaña **Suiza_DE_CN_ID** — sin Tscode, esa columna se calcula sola.
3. La pestaña **Consolidado** se arma sola con fórmulas. No está protegida a propósito: se
   puede armar tablas dinámicas y gráficos directamente sobre ella. Como contrapartida, no se
   debe editar/arrastrar/borrar filas manualmente ahí (rompe las fórmulas) — si necesitas otra
   vista, créala en una pestaña nueva en vez de tocar Consolidado.
4. Si aparece "Revisar Tscode", falta ese Tscode en **Mercados**. Si en Suiza aparece
   "Revisar alias", el nombre de `cliente` no coincide con ningún Alias — agrega una fila
   en **Mercados** (Tscode | Alias | Mercado) escribiendo el texto directamente, sin copiar
   celdas desde otro Excel abierto (eso puede pegar un vínculo roto en vez del texto).
5. Si llega una versión nueva del catálogo de productos Gamma, **no la pegues directamente en
   la pestaña Gamma_Catalogo** (aunque esté oculta) — eso puede traer fórmulas u otros formatos
   que no funcionan ahí y romper Consolidado. Usa en cambio:

   ```
   python3 update_gamma_catalog.py Consolidador_Facturacion.xlsx catalogo_nuevo.xlsx
   ```

   donde `catalogo_nuevo.xlsx` tiene solo estas 5 columnas, en este orden, como valores (no
   fórmulas): Ip Code, Description, R/C, Category, Mks description. El script reemplaza el
   catálogo y ajusta automáticamente el rango de las fórmulas de Consolidado que lo consultan.
   Recalcula después.

   Si en cambio lo que cambió es la tabla de mercados (`mercados.xlsx`), regenera el libro base
   completo (esto sí reconstruye todo desde cero, no solo el catálogo):

   ```
   python3 build_consolidador.py /ruta/a/Gamma.xls /ruta/a/mercados.xlsx Consolidador_Facturacion.xlsx
   ```

6. `build_resumen.py`, `hide_blank_rows.py` y `finalize_delivery.py` (ver abajo) cada uno abre
   el libro con openpyxl y lo vuelve a guardar — **eso borra los valores calculados de todas
   las fórmulas del libro**, no solo de lo que cada script toca. Como `hide_blank_rows.py`
   necesita leer esos valores para saber qué filas están vacías, y `build_resumen.py` los
   necesita para armar la tabla, **hay que recalcular después de cada uno de estos pasos**, no
   solo al final:

   ```
   python3 recalc.py Consolidador_Facturacion.xlsx      # (script del skill de xlsx, o tu copia)

   python3 build_resumen.py Consolidador_Facturacion.xlsx     # agrega/actualiza Resumen
   python3 recalc.py Consolidador_Facturacion.xlsx

   python3 hide_blank_rows.py Consolidador_Facturacion.xlsx   # oculta filas sin datos en Consolidado
   python3 finalize_delivery.py Consolidador_Facturacion.xlsx # quita Instrucciones, oculta Mercados/Gamma_Catalogo
   python3 recalc.py Consolidador_Facturacion.xlsx      # deja cacheados los valores finales
   ```

   `hide_blank_rows.py` solo oculta filas (no las borra); si el mes siguiente cambia la
   cantidad de filas con datos, hay que volver a correrlo. `finalize_delivery.py` no lee
   valores calculados, así que puede ir después de `hide_blank_rows.py` sin recalcular entre
   los dos — pero sí hay que recalcular una última vez después de ambos.

Instrucciones detalladas dentro del propio libro, pestaña **Instrucciones** (mientras no se
haya corrido `finalize_delivery.py`).

## Cuando los otros orígenes llegan en PDF (proforma individual) en vez de Excel

Si en lugar del Excel de Suiza llegan proformas en PDF (formato "PIRELLI TYRE (SUISSE) SA",
una por cliente/pedido), `parse_proformas_pdf.py` las lee y arma una tabla con las mismas
columnas de `Suiza_DE_CN_ID` (cliente, Tscode, pf, Ip Code, Description, Brand Line, Quantity,
Net Price, Amount, Goods Origin) — el Tscode y el país de despacho salen del propio PDF, no
hace falta alias:

```
pip install -r requirements.txt   # una sola vez: pandas, openpyxl, pdfplumber
```

Con 15-20 PDFs al mes, lo más simple es que se escriban directo en el libro (reemplaza lo que
haya en `Suiza_DE_CN_ID`, no hace falta copiar/pegar):

```
python3 load_pdfs_into_suiza.py Consolidador_Facturacion.xlsx carpeta_con_pdfs/
python3 recalc.py Consolidador_Facturacion.xlsx      # desde el skill de xlsx, o tu copia
```

Y sigue con el resto del flujo del punto 6 de arriba (build_resumen → recalc → hide_blank_rows
→ finalize_delivery → recalc).

Si prefieres revisar los datos en un Excel aparte antes de pegarlos, `parse_proformas_pdf.py`
hace lo mismo pero deja el resultado en un archivo suelto en vez de escribirlo en el libro:

```
python3 parse_proformas_pdf.py carpeta_con_pdfs/ -o suiza_desde_pdf.xlsx
```

Ambos aceptan varios PDFs sueltos o una carpeta entera.
