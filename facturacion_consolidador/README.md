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
