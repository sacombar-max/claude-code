# Analizador de portafolio BVC

Genera un informe HTML + CSV de tu portafolio de acciones de la Bolsa de Valores de Colombia:
posición %, precio, variación, rango del día, ganancia/pérdida, valor intrínseco (Graham),
consenso, precio objetivo, Piotroski, BPA, P/E y capitalización.

## Uso

```bash
pip install -r ../requirements.txt     # yfinance, pandas
python analizador.py                    # lee portafolio.csv → reportes/portafolio_AAAA-MM-DD.html
python analizador.py --solo-manual      # sin internet: solo precios_manuales.csv
```

## Archivos de entrada

`portafolio.csv` (las filas incluidas son solo un ejemplo; reemplázalas por tus posiciones):

| simbolo | cantidad | costo_promedio | yahoo (opcional) |
|---|---|---|---|
| PFBCOLOM | 100 | 32000 | |

Guarda tus posiciones reales en `mi_portafolio.csv` (mismo formato): git lo ignora y el programa lo usa automáticamente si existe.

Si `yahoo` está vacío se usa `SIMBOLO.CL`. Si Yahoo no reconoce un ticker, escribe aquí el símbolo correcto.

`precios_manuales.csv`: precios que digitas tú (p. ej. copiados de bvc.com.co). Cuando existen, prevalecen sobre Yahoo
y la columna "Fuente" dice "Manual".

## Limitaciones conocidas

- La cobertura de Yahoo Finance para acciones colombianas es irregular: consenso, precio objetivo y estados financieros
  suelen faltar en emisores pequeños. En ese caso se muestra **N/D**; nunca se estiman ni se rellenan datos.
- Piotroski se muestra como `cumplidos/evaluables`. En bancos varios criterios no aplican (no hay activos corrientes ni
  utilidad bruta), así que el denominador puede ser menor que 9.
- No se incluye una "valoración" compuesta propietaria como la de algunos screeners, porque su metodología no es pública.
- No es una recomendación de inversión.
