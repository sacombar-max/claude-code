# Consolidador de Facturación Multipaís

Libro Excel (`Consolidador_Facturacion.xlsx`) para consolidar la facturación mensual de
Brasil, China, Alemania e Indonesia en un solo formato:

`mercado | cliente | pf | Ip Code | Description | Quantity | Net Price | Amount | Goods Origin | Category | R/C | Mks description`

`mercado` es el país del **cliente** (se busca por Tscode en la pestaña `Mercados`), no el país
de fabricación — ese es `Goods Origin`.

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
5. Si llega una versión nueva del catálogo de productos Gamma o de la tabla de mercados,
   regenera el libro corriendo:

   ```
   python3 build_consolidador.py /ruta/a/Gamma.xls /ruta/a/mercados.xlsx Consolidador_Facturacion.xlsx
   ```

6. Para agregar/actualizar la pestaña **Resumen** (tabla cruzada: mercado en filas, piezas
   facturadas por cada segmento/Mks description en columnas, y Amount Total por mercado al
   final) una vez que el libro ya tiene los datos del mes pegados y recalculados:

   ```
   python3 build_resumen.py Consolidador_Facturacion.xlsx
   ```

Instrucciones detalladas dentro del propio libro, pestaña **Instrucciones**.

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
python3 build_resumen.py Consolidador_Facturacion.xlsx   # opcional, actualiza el Resumen
```

Si prefieres revisar los datos en un Excel aparte antes de pegarlos, `parse_proformas_pdf.py`
hace lo mismo pero deja el resultado en un archivo suelto en vez de escribirlo en el libro:

```
python3 parse_proformas_pdf.py carpeta_con_pdfs/ -o suiza_desde_pdf.xlsx
```

Ambos aceptan varios PDFs sueltos o una carpeta entera.
