# Mis Finanzas — Context for Claude

## What this is
Personal finance web app. Backend: Google Apps Script. Frontend: SPA served via `doGet()`. Database: Google Sheets (4 tabs).

## File structure (copy all to Apps Script editor)

| File | Lines | Purpose |
|------|-------|---------|
| `Code.gs` | ~330 | All server-side logic |
| `index.html` | ~185 | Apps Script template — HTML skeleton only, no JS |
| `styles.html` | ~260 | `<style>` block — all CSS |
| `script.html` | ~380 | Core JS: CONFIG, STATE, UTILS, NAVIGATION, DATA, RENDER for Resumen/Registrar/Historial/Presupuesto/Tendencias, INIT |
| `script-inversiones.html` | ~430 | TICKERS_BVC, TICKERS_US, all INVERSIONES functions |

**Rule: only read the file relevant to the task.** Bug in budget? Read `script.html`. Bug in portfolio tabs? Read `script-inversiones.html`. CSS issue? Read `styles.html`. Backend issue? Read `Code.gs`.

## Architecture

```
Browser → doGet() → HtmlService.createTemplateFromFile('index').evaluate()
                     → include('styles')            → styles.html content
                     → include('script')            → script.html content
                     → include('script-inversiones') → script-inversiones.html content

Client calls: google.script.run.withSuccessHandler(ok).withFailureHandler(fail).functionName(args)
Wrapped by: api(fn, ...args) → Promise  (defined in script.html)
```

## Sheet schemas (SHEETS constant in Code.gs)

```
Transacciones: id | fecha | tipo | descripcion | categoria | monto | tipo_gasto | cuenta
Presupuesto:   mes | categoria | monto
Me_Deben:      id | fecha | persona | descripcion | monto | tasa | estado
Inversiones:   id | fecha | tipo | nombre | moneda | cantidad | precio_entrada | valor_cop_entrada | tasa | fecha_cierre | precio_cierre | valor_cop_cierre | estado | nota
```

## State object (S) — defined in script.html

```js
const S = {
  view, mes, anio, txs, budget, medeben, trends, inversiones,
  invView, invTab, invAnio, rtype, tgasto, cuenta, filtro, chart
}
```

## Key patterns

### Adding a new field to a sheet
1. Add column name to `SHEETS[key].headers` in `Code.gs`
2. `getSheet()` auto-adds missing columns on next call (no manual migration needed)
3. `sheetToObjects()` maps all headers automatically

### Budget months bug
Google Sheets auto-converts "2026-05" strings to Date objects. Fix: `mesStr()` in Code.gs normalizes everything to `YYYY-MM`. Column A of Presupuesto sheet has `@` format applied by `getSheet()`.

### Real-time prices (getPrecioActual)
Uses a hidden `_cache_` sheet. Writes a GOOGLEFINANCE formula, calls `SpreadsheetApp.flush()`, sleeps 2500-3000ms, reads the result. Works for BVC tickers (`BVC:ECOPETL`), US tickers (`AAPL`), and FX (`CURRENCY:USDCOP`).

### Ticker input
**Never free-text.** BVC uses `<select id="inv-nombre-bvc">` built from `TICKERS_BVC`. US uses search filter `<input id="inv-us-filter">` + `<select id="inv-nombre-us">`. The stored `nombre` field IS the ticker string passed to GOOGLEFINANCE.

### importarDesdeHoja
Uses `sheet.getRange(...).setValues(filas)` batch (not appendRow loop — too slow).

### testConexion
Uses `Logger.log()` not `getUi().alert()` — cannot call getUi() from Apps Script editor.

## Investment types

```js
const INV_TIPOS = {
  accion_nacional:      'Acción Nacional',       // BVC Colombia
  accion_internacional: 'Acción Internacional',  // NYSE/NASDAQ
  prestamo:             'Préstamo',
  divisa:               'Divisa'
};

const TAB_TIPOS = {
  colombia:  ['accion_nacional'],
  eeuu:      ['accion_internacional'],
  divisas:   ['divisa'],
  prestamos: ['prestamo']
};
```

Portfolio view groups investments by exchange tab (`S.invTab`).

## Deployment

1. Open [script.google.com](https://script.google.com), open the project
2. Paste/sync all 5 files (Code.gs + 4 HTML files)
3. Deploy → New deployment → Web app → "Anyone" access
4. Copy the deployment URL

Each code change requires a **new deployment** (or update existing deployment).

## Known tickers

`TICKERS_BVC` (20 entries) — format: `BVC:TICKER` — defined in `script-inversiones.html`
`TICKERS_US` (~123 entries, stocks + ETFs) — format: `TICKER` — defined in `script-inversiones.html`

If a BVC ticker fails with getPrecioActual, try removing the `BVC:` prefix or check GOOGLEFINANCE docs for the correct symbol.
