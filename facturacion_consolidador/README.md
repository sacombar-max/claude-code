# Consolidador de Facturación Multipaís

Libro Excel (`Consolidador_Facturacion.xlsx`) para consolidar la facturación mensual de
Brasil, China, Alemania e Indonesia en un solo formato:

`mercado | cliente | pf | Ip Code | Description | Quantity | Net Price | Amount | Goods Origin | Category | R/C | Mks description`

`mercado` es el país del **cliente** (se busca por Tscode en la pestaña `Mercados`), no el país
de fabricación — ese es `Goods Origin`.

## Uso mensual

1. Pega los datos de Brasil en la pestaña **Brasil** (reemplaza la fila de ejemplo).
2. Pega los datos del archivo de Suiza (agrupa Alemania, China e Indonesia) en la
   pestaña **Suiza_DE_CN_ID**, incluyendo su columna Tscode.
3. La pestaña **Consolidado** se arma sola con fórmulas.
4. Si aparece "Revisar Tscode", agrega ese Tscode y su mercado en la pestaña **Mercados**.
5. Si llega una versión nueva del catálogo de productos Gamma o de la tabla de mercados,
   regenera el libro corriendo:

   ```
   python3 build_consolidador.py /ruta/a/Gamma.xls /ruta/a/mercados.xlsx Consolidador_Facturacion.xlsx
   ```

Instrucciones detalladas dentro del propio libro, pestaña **Instrucciones**.
