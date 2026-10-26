#!/usr/bin/env python3
"""
Analizador de portafolio para acciones de la Bolsa de Valores de Colombia (BVC).

Lee las posiciones de portafolio.csv, obtiene precios y fundamentales vía
yfinance (sufijo .CL por defecto) y genera un informe HTML + CSV con columnas
similares a las de un screener: posición %, precio, variación, rango del día,
valor intrínseco, consenso, precio objetivo, Piotroski, BPA, P/E y
capitalización.

Principio: ningún dato se inventa. Si una fuente no entrega un valor, la
celda queda "N/D" y el informe indica de dónde salió cada cifra.

Uso:
    python analizador.py                      # usa portafolio.csv
    python analizador.py -p mi_portafolio.csv -o reports/
    python analizador.py --solo-manual        # sin internet, solo precios_manuales.csv
"""

import argparse
import csv
import datetime
import html
import math
import os
import sys

SUFIJO_YAHOO = ".CL"
DIR_BASE = os.path.dirname(os.path.abspath(__file__))


# ──────────────────────────────────────────────
# Lectura de archivos
# ──────────────────────────────────────────────
def leer_portafolio(ruta):
    posiciones = []
    with open(ruta, newline="", encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            simbolo = (fila.get("simbolo") or "").strip().upper()
            if not simbolo:
                continue
            posiciones.append({
                "simbolo": simbolo,
                "cantidad": float(fila["cantidad"]),
                "costo_promedio": _num(fila.get("costo_promedio")),
                "yahoo": (fila.get("yahoo") or "").strip() or simbolo + SUFIJO_YAHOO,
            })
    return posiciones


def leer_precios_manuales(ruta):
    """Precios digitados por el usuario (p. ej. copiados de bvc.com.co)."""
    if not os.path.exists(ruta):
        return {}
    precios = {}
    with open(ruta, newline="", encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            simbolo = (fila.get("simbolo") or "").strip().upper()
            if simbolo and _num(fila.get("precio")) is not None:
                precios[simbolo] = {
                    "precio": _num(fila.get("precio")),
                    "cierre_anterior": _num(fila.get("cierre_anterior")),
                    "minimo_dia": _num(fila.get("minimo_dia")),
                    "maximo_dia": _num(fila.get("maximo_dia")),
                }
    return precios


def _num(valor):
    if valor is None:
        return None
    try:
        x = float(str(valor).strip())
    except ValueError:
        return None
    return None if math.isnan(x) else x


# ──────────────────────────────────────────────
# Obtención de datos (yfinance)
# ──────────────────────────────────────────────
def obtener_datos_yahoo(ticker_yahoo):
    """Devuelve un dict con datos crudos de Yahoo; claves ausentes = no disponible."""
    import yfinance as yf

    t = yf.Ticker(ticker_yahoo)
    datos = {}
    try:
        fi = t.fast_info
        datos["precio"] = _num(getattr(fi, "last_price", None))
        datos["cierre_anterior"] = _num(getattr(fi, "previous_close", None))
        datos["minimo_dia"] = _num(getattr(fi, "day_low", None))
        datos["maximo_dia"] = _num(getattr(fi, "day_high", None))
        datos["capitalizacion"] = _num(getattr(fi, "market_cap", None))
        datos["moneda"] = getattr(fi, "currency", None)
    except Exception as exc:
        print(f"  [aviso] {ticker_yahoo} fast_info: {exc}", file=sys.stderr)

    try:
        info = t.info or {}
    except Exception as exc:
        info = {}
        print(f"  [aviso] {ticker_yahoo} info: {exc}", file=sys.stderr)
    datos["bpa"] = _num(info.get("trailingEps"))
    datos["valor_libros"] = _num(info.get("bookValue"))
    datos["consenso"] = info.get("recommendationKey")
    datos["precio_objetivo"] = _num(info.get("targetMeanPrice"))
    datos["n_analistas"] = info.get("numberOfAnalystOpinions")
    datos["nombre"] = info.get("shortName") or info.get("longName")
    if datos.get("capitalizacion") is None:
        datos["capitalizacion"] = _num(info.get("marketCap"))

    try:
        datos["piotroski"] = calcular_piotroski(t.financials, t.balance_sheet, t.cashflow)
    except Exception as exc:
        datos["piotroski"] = None
        print(f"  [aviso] {ticker_yahoo} estados financieros: {exc}", file=sys.stderr)
    return datos


# ──────────────────────────────────────────────
# Métricas
# ──────────────────────────────────────────────
def valor_graham(bpa, valor_libros):
    """Número de Graham: sqrt(22.5 × BPA × valor en libros por acción).

    Solo es válido con BPA y valor en libros positivos; si no, None.
    """
    if bpa is None or valor_libros is None or bpa <= 0 or valor_libros <= 0:
        return None
    return math.sqrt(22.5 * bpa * valor_libros)


def _fila(df, *nombres, col=0):
    """Valor de la primera fila existente entre `nombres`, columna `col` (0 = año más reciente)."""
    if df is None or getattr(df, "empty", True) or df.shape[1] <= col:
        return None
    for n in nombres:
        if n in df.index:
            return _num(df.loc[n].iloc[col])
    return None


def calcular_piotroski(ingresos, balance, flujo):
    """F-Score de Piotroski (9 criterios) con los dos últimos años anuales.

    Devuelve {"puntaje", "evaluados", "detalle"}; los criterios sin datos
    (frecuente en bancos: sin activos corrientes ni utilidad bruta) no suman
    ni se cuentan como evaluados.
    """
    def g(df, col, *n):
        return _fila(df, *n, col=col)

    ni0, ni1 = g(ingresos, 0, "Net Income"), g(ingresos, 1, "Net Income")
    ta0, ta1 = g(balance, 0, "Total Assets"), g(balance, 1, "Total Assets")
    ta2 = g(balance, 2, "Total Assets")
    cfo0 = g(flujo, 0, "Operating Cash Flow")
    ltd0 = g(balance, 0, "Long Term Debt")
    ltd1 = g(balance, 1, "Long Term Debt")
    ca0, ca1 = g(balance, 0, "Current Assets"), g(balance, 1, "Current Assets")
    cl0 = g(balance, 0, "Current Liabilities")
    cl1 = g(balance, 1, "Current Liabilities")
    sh0 = g(balance, 0, "Ordinary Shares Number", "Share Issued")
    sh1 = g(balance, 1, "Ordinary Shares Number", "Share Issued")
    gp0, gp1 = g(ingresos, 0, "Gross Profit"), g(ingresos, 1, "Gross Profit")
    rv0, rv1 = g(ingresos, 0, "Total Revenue"), g(ingresos, 1, "Total Revenue")

    def div(a, b):
        return a / b if a is not None and b not in (None, 0) else None

    # ROA usa activos promedio cuando hay tres balances; si no, activos de cierre.
    roa0 = div(ni0, (ta0 + ta1) / 2 if ta0 is not None and ta1 is not None else ta0)
    roa1 = div(ni1, (ta1 + ta2) / 2 if ta1 is not None and ta2 is not None else ta1)

    def cmp(a, b, mayor=True):
        if a is None or b is None:
            return None
        return a > b if mayor else a < b

    criterios = {
        "ROA positivo": None if roa0 is None else roa0 > 0,
        "Flujo operativo positivo": None if cfo0 is None else cfo0 > 0,
        "ROA creciente": cmp(roa0, roa1),
        "Flujo operativo > utilidad neta": cmp(cfo0, ni0),
        "Menor apalancamiento LP": cmp(div(ltd0, ta0), div(ltd1, ta1), mayor=False),
        "Mayor razón corriente": cmp(div(ca0, cl0), div(ca1, cl1)),
        "Sin emisión de acciones": None if sh0 is None or sh1 is None else sh0 <= sh1,
        "Mayor margen bruto": cmp(div(gp0, rv0), div(gp1, rv1)),
        "Mayor rotación de activos": cmp(div(rv0, ta0), div(rv1, ta1)),
    }
    evaluados = [v for v in criterios.values() if v is not None]
    if not evaluados:
        return None
    return {"puntaje": sum(evaluados), "evaluados": len(evaluados), "detalle": criterios}


def analizar(posiciones, precios_manuales, usar_yahoo=True):
    filas = []
    for p in posiciones:
        print(f"→ {p['simbolo']} ({p['yahoo']})")
        datos, fuente = {}, "N/D"
        if usar_yahoo:
            try:
                datos = obtener_datos_yahoo(p["yahoo"])
                if datos.get("precio") is not None:
                    fuente = "Yahoo Finance"
            except Exception as exc:
                print(f"  [aviso] {p['yahoo']}: {exc}", file=sys.stderr)
        # Un precio manual siempre prevalece: es lo que el usuario verificó.
        manual = precios_manuales.get(p["simbolo"])
        if manual:
            datos.update({k: v for k, v in manual.items() if v is not None})
            fuente = "Manual"
        filas.append(construir_fila(p, datos, fuente))

    total = sum(f["valor_mercado"] for f in filas if f["valor_mercado"] is not None)
    for f in filas:
        f["posicion_pct"] = f["valor_mercado"] / total * 100 if total and f["valor_mercado"] is not None else None
    return filas, total


def construir_fila(p, d, fuente):
    precio = d.get("precio")
    prev = d.get("cierre_anterior")
    var = precio - prev if precio is not None and prev is not None else None
    var_pct = var / prev * 100 if var is not None and prev else None
    bpa = d.get("bpa")
    valor_mercado = precio * p["cantidad"] if precio is not None else None
    costo_total = p["costo_promedio"] * p["cantidad"] if p["costo_promedio"] is not None else None
    ganancia = valor_mercado - costo_total if valor_mercado is not None and costo_total is not None else None
    return {
        "simbolo": p["simbolo"],
        "yahoo": p["yahoo"],
        "nombre": d.get("nombre"),
        "cantidad": p["cantidad"],
        "costo_promedio": p["costo_promedio"],
        "precio": precio,
        "variacion": var,
        "variacion_pct": var_pct,
        "minimo_dia": d.get("minimo_dia"),
        "maximo_dia": d.get("maximo_dia"),
        "valor_intrinseco": valor_graham(bpa, d.get("valor_libros")),
        "consenso": (d.get("consenso") or "").upper() or None,
        "n_analistas": d.get("n_analistas"),
        "precio_objetivo": d.get("precio_objetivo"),
        "piotroski": d.get("piotroski"),
        "bpa": bpa,
        "pe": precio / bpa if precio is not None and bpa and bpa > 0 else None,
        "capitalizacion": d.get("capitalizacion"),
        "valor_mercado": valor_mercado,
        "ganancia": ganancia,
        "ganancia_pct": ganancia / costo_total * 100 if ganancia is not None and costo_total else None,
        "fuente": fuente,
    }


# ──────────────────────────────────────────────
# Formato y salida
# ──────────────────────────────────────────────
def fmt(x, dec=2, signo=False, sufijo=""):
    if x is None:
        return "N/D"
    s = f"{x:+,.{dec}f}" if signo else f"{x:,.{dec}f}"
    # Formato colombiano: punto de miles, coma decimal.
    return s.replace(",", "_").replace(".", ",").replace("_", ".") + sufijo


def fmt_cap(x):
    if x is None:
        return "N/D"
    for limite, letra in ((1e12, "B"), (1e9, "MM")):  # billones / miles de millones (COP)
        if abs(x) >= limite:
            return fmt(x / limite, 2) + " " + letra
    return fmt(x, 0)


def fmt_piotroski(p):
    return "N/D" if not p else f"{p['puntaje']}/{p['evaluados']}"


def escribir_csv(filas, ruta):
    campos = ["simbolo", "yahoo", "cantidad", "costo_promedio", "precio", "variacion",
              "variacion_pct", "minimo_dia", "maximo_dia", "posicion_pct", "valor_mercado",
              "ganancia", "ganancia_pct", "valor_intrinseco", "consenso", "precio_objetivo",
              "piotroski", "bpa", "pe", "capitalizacion", "fuente"]
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
        w.writeheader()
        for fila in filas:
            fila = dict(fila, piotroski=fmt_piotroski(fila["piotroski"]))
            w.writerow({k: ("" if fila.get(k) is None else fila.get(k)) for k in campos})


def _rango_html(lo, hi, precio):
    if lo is None or hi is None:
        return "N/D"
    ancho = 100 if hi == lo or precio is None else max(0, min(100, (precio - lo) / (hi - lo) * 100))
    return (f'{fmt(lo)} – {fmt(hi)}<div class="barra"><span style="width:{ancho:.0f}%"></span></div>')


def generar_html(filas, total, fecha):
    def clase(x):
        return "" if x is None else ("pos" if x > 0 else "neg" if x < 0 else "")

    cuerpo = []
    for f in sorted(filas, key=lambda r: r["variacion_pct"] if r["variacion_pct"] is not None else -1e9, reverse=True):
        pio = f["piotroski"]
        pio_tip = "" if not pio else html.escape("; ".join(
            f"{k}: {'sí' if v else 'no' if v is not None else 'sin datos'}" for k, v in pio["detalle"].items()))
        cons = f["consenso"] or "N/D"
        if f["consenso"] and f["n_analistas"]:
            cons += f' <small>({f["n_analistas"]})</small>'
        cuerpo.append(f"""
<tr>
  <td class="sim" title="{html.escape(f['nombre'] or f['yahoo'])}">{html.escape(f['simbolo'])}</td>
  <td><b>{fmt(f['posicion_pct'], 1, sufijo='%')}</b></td>
  <td><b>{fmt(f['precio'])}</b></td>
  <td class="{clase(f['variacion'])}">{fmt(f['variacion'], signo=True)}</td>
  <td class="{clase(f['variacion_pct'])}">{fmt(f['variacion_pct'], signo=True, sufijo='%')}</td>
  <td>{_rango_html(f['minimo_dia'], f['maximo_dia'], f['precio'])}</td>
  <td class="{clase(f['ganancia'])}">{fmt(f['ganancia'], 0, signo=True)}<br><small>{fmt(f['ganancia_pct'], 1, signo=True, sufijo='%')}</small></td>
  <td>{fmt(f['valor_intrinseco'])}</td>
  <td>{cons}</td>
  <td>{fmt(f['precio_objetivo'])}</td>
  <td title="{pio_tip}">{fmt_piotroski(pio)}</td>
  <td>{fmt(f['bpa'])}</td>
  <td>{fmt(f['pe'])}</td>
  <td>{fmt_cap(f['capitalizacion'])}</td>
  <td><small>{f['fuente']}</small></td>
</tr>""")

    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Portafolio BVC</title>
<style>
:root {{ --fg:#1f2937; --muted:#6b7280; --line:#e5e7eb; --bg:#fff; --pos:#16a34a; --neg:#dc2626; --acc:#14b8a6; }}
@media (prefers-color-scheme: dark) {{ :root {{ --fg:#e5e7eb; --muted:#9ca3af; --line:#374151; --bg:#111827; }} }}
body {{ font-family: system-ui, sans-serif; color:var(--fg); background:var(--bg); margin:24px 16px; }}
h1 {{ font-size:1.3rem; margin:0 0 4px; }}
.meta {{ color:var(--muted); font-size:.85rem; margin-bottom:16px; }}
.wrap {{ overflow-x:auto; }}
table {{ border-collapse:collapse; width:100%; font-size:.9rem; }}
th {{ text-align:left; font-weight:600; padding:10px 8px; border-bottom:1px solid var(--line); vertical-align:bottom; }}
td {{ padding:12px 8px; border-bottom:1px solid var(--line); white-space:nowrap; }}
td.sim {{ color:var(--acc); font-weight:600; font-size:1rem; }}
.pos {{ color:var(--pos); }} .neg {{ color:var(--neg); }}
.barra {{ height:4px; background:var(--line); border-radius:2px; margin-top:4px; width:120px; }}
.barra span {{ display:block; height:100%; background:var(--acc); border-radius:2px; }}
.notas {{ color:var(--muted); font-size:.8rem; margin-top:20px; line-height:1.5; max-width:900px; }}
</style></head><body>
<h1>Portafolio BVC</h1>
<div class="meta">Generado: {fecha} · Valor de mercado total: <b>COP {fmt(total, 0)}</b></div>
<div class="wrap"><table>
<thead><tr><th>Símbolo</th><th>Posición</th><th>Precio</th><th>Variación</th><th>Variación %</th>
<th>Rango del día</th><th>Ganancia / pérdida</th><th>Valor intrínseco<br><small>(Graham)</small></th>
<th>Consenso</th><th>Precio objetivo</th><th>Piotroski</th><th>BPA</th><th>P/E</th><th>Capitalización</th><th>Fuente</th></tr></thead>
<tbody>{''.join(cuerpo)}
</tbody></table></div>
<div class="notas">
<b>Cómo se calcula cada columna</b><br>
Precio, cierre anterior, rango, BPA (trailingEps), valor en libros, consenso (recommendationKey), precio objetivo
(targetMeanPrice) y capitalización: Yahoo Finance vía yfinance, o precios_manuales.csv cuando la fuente dice "Manual".<br>
Valor intrínseco = √(22,5 × BPA × valor en libros por acción) (número de Graham); N/D si alguno es ≤ 0 o falta.<br>
Piotroski = criterios cumplidos / criterios evaluables con los dos últimos estados financieros anuales
(pase el cursor para ver el detalle). En bancos varios criterios no aplican, por eso el denominador puede ser &lt; 9.<br>
P/E = precio ÷ BPA. Capitalización: B = billones, MM = miles de millones (en la moneda que reporte la fuente).<br>
N/D = la fuente no entregó el dato; no se estima ni se rellena. Verifique cifras críticas en bvc.com.co
y en los estados financieros publicados por cada emisor. Esto no es una recomendación de inversión.
</div>
</body></html>"""


def main():
    ap = argparse.ArgumentParser(description="Analizador de portafolio BVC")
    ap.add_argument("-p", "--portafolio", default=os.path.join(DIR_BASE, "portafolio.csv"))
    ap.add_argument("-m", "--manuales", default=os.path.join(DIR_BASE, "precios_manuales.csv"))
    ap.add_argument("-o", "--salida", default=os.path.join(DIR_BASE, "reportes"))
    ap.add_argument("--solo-manual", action="store_true", help="no consultar Yahoo Finance")
    args = ap.parse_args()

    posiciones = leer_portafolio(args.portafolio)
    if not posiciones:
        sys.exit("El portafolio está vacío.")
    filas, total = analizar(posiciones, leer_precios_manuales(args.manuales), not args.solo_manual)

    os.makedirs(args.salida, exist_ok=True)
    ahora = datetime.datetime.now()
    base = os.path.join(args.salida, f"portafolio_{ahora:%Y-%m-%d}")
    with open(base + ".html", "w", encoding="utf-8") as f:
        f.write(generar_html(filas, total, ahora.strftime("%Y-%m-%d %H:%M")))
    escribir_csv(filas, base + ".csv")

    sin_precio = [f["simbolo"] for f in filas if f["precio"] is None]
    print(f"\nInforme: {base}.html\nDatos:   {base}.csv")
    if sin_precio:
        print("Sin precio (agréguelos a precios_manuales.csv o revise el ticker Yahoo): "
              + ", ".join(sin_precio))


if __name__ == "__main__":
    main()
