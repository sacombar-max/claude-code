// ============================================================
// MIS FINANZAS — Google Apps Script Backend
// ============================================================

const SHEETS = {
  TX:         { name: 'Transacciones', headers: ['id','fecha','tipo','descripcion','categoria','monto','tipo_gasto','cuenta'] },
  BUDGET:     { name: 'Presupuesto',   headers: ['mes','categoria','monto'] },
  MEDEBEN:    { name: 'Me_Deben',      headers: ['id','fecha','persona','descripcion','monto','tasa','estado'] },
  INVERSIONES:{ name: 'Inversiones',   headers: ['id','fecha','tipo','nombre','moneda','cantidad','precio_entrada','valor_cop_entrada','tasa','fecha_cierre','precio_cierre','valor_cop_cierre','estado','nota'] }
};

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Mis Finanzas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---- Helpers ----

function getSheet(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { name, headers } = SHEETS[key];
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const range = sheet.getRange(1, 1, 1, headers.length);
    range.setValues([headers]);
    range.setFontWeight('bold');
    range.setBackground('#E8F5E9');
  } else {
    // Agrega columnas nuevas si la hoja ya existía sin ellas
    const lastCol = sheet.getLastColumn();
    if (lastCol < headers.length) {
      for (let i = lastCol; i < headers.length; i++) {
        sheet.getRange(1, i + 1).setValue(headers[i]).setFontWeight('bold').setBackground('#E8F5E9');
      }
    }
  }
  if (key === 'BUDGET') sheet.getRange('A:A').setNumberFormat('@');
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const tz = Session.getScriptTimeZone();
  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const v = row[i];
        obj[h] = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v;
      });
      return obj;
    });
}

// ---- Transacciones ----

function getTransacciones(mes) {
  return sheetToObjects(getSheet('TX')).filter(t => String(t.fecha).startsWith(mes));
}

function getTransaccionesPorAnio(anio) {
  return sheetToObjects(getSheet('TX')).filter(t => String(t.fecha).startsWith(anio));
}

function addTransaccion(data) {
  const id = Utilities.getUuid();
  getSheet('TX').appendRow([
    id, data.fecha, data.tipo, data.descripcion,
    data.categoria, Number(data.monto), data.tipo_gasto || '', data.cuenta || ''
  ]);
  return { success: true, id };
}

function deleteTransaccion(id) {
  const sheet = getSheet('TX');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ---- Presupuesto ----

function mesStr(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(val).substring(0, 7);
}

function getPresupuesto(mes) {
  return sheetToObjects(getSheet('BUDGET')).filter(p => mesStr(p.mes) === mes);
}

function setPresupuesto(mes, categoria, monto) {
  const sheet = getSheet('BUDGET');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (mesStr(values[i][0]) === mes && values[i][1] === categoria) {
      sheet.getRange(i + 1, 3).setValue(Number(monto));
      return { success: true };
    }
  }
  sheet.appendRow([mes, categoria, Number(monto)]);
  return { success: true };
}

function deletePresupuesto(mes, categoria) {
  const sheet = getSheet('BUDGET');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (mesStr(values[i][0]) === mes && values[i][1] === categoria) {
      sheet.deleteRow(i + 1); return { success: true };
    }
  }
  return { success: false };
}

function copiarPresupuestoMesAnterior(mes) {
  const [y, m] = mes.split('-').map(Number);
  const prev = new Date(y, m - 2, 1);
  const prevMes = `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, '0')}`;
  const prevBudget = getPresupuesto(prevMes);
  if (prevBudget.length === 0) return { success: false, message: 'Sin presupuesto en el mes anterior' };
  prevBudget.forEach(b => setPresupuesto(mes, b.categoria, b.monto));
  return { success: true, count: prevBudget.length };
}

// ---- Me Deben ----

function getMeDeben() {
  return sheetToObjects(getSheet('MEDEBEN'));
}

function addMeDeben(data) {
  const id = Utilities.getUuid();
  getSheet('MEDEBEN').appendRow([id, data.fecha, data.persona, data.descripcion, Number(data.monto), Number(data.tasa || 0), 'pendiente']);
  return { success: true, id };
}

function updateMeDebenEstado(id, estado) {
  const sheet = getSheet('MEDEBEN');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) { sheet.getRange(i + 1, 6).setValue(estado); return { success: true }; }
  }
  return { success: false };
}

function deleteMeDeben(id) {
  const sheet = getSheet('MEDEBEN');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ---- Inversiones ----

function getInversiones() {
  return sheetToObjects(getSheet('INVERSIONES'));
}

function addInversion(data) {
  const id = Utilities.getUuid();
  getSheet('INVERSIONES').appendRow([
    id, data.fecha, data.tipo, data.nombre,
    data.moneda || 'COP', Number(data.cantidad),
    Number(data.precio_entrada), Number(data.valor_cop_entrada),
    Number(data.tasa || 0),
    '', '', '', 'abierta', data.nota || ''
  ]);
  return { success: true, id };
}

function cerrarInversion(id, datos) {
  const sheet = getSheet('INVERSIONES');
  const values = sheet.getDataRange().getValues();
  const h = values[0];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      const r = i + 1;
      sheet.getRange(r, h.indexOf('fecha_cierre') + 1).setValue(datos.fecha_cierre);
      sheet.getRange(r, h.indexOf('precio_cierre') + 1).setValue(Number(datos.precio_cierre || 0));
      sheet.getRange(r, h.indexOf('valor_cop_cierre') + 1).setValue(Number(datos.valor_cop_cierre));
      sheet.getRange(r, h.indexOf('estado') + 1).setValue('cerrada');
      return { success: true };
    }
  }
  return { success: false };
}

function deleteInversion(id) {
  const sheet = getSheet('INVERSIONES');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function getPrecioActual(tipo, nombre, moneda) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let cache = ss.getSheetByName('_cache_');
    if (!cache) { cache = ss.insertSheet('_cache_'); cache.hideSheet(); }

    if (tipo === 'divisa') {
      const sym = moneda === 'EUR' ? 'EURCOP' : 'USDCOP';
      cache.getRange('A1').setFormula(`=GOOGLEFINANCE("CURRENCY:${sym}")`);
      SpreadsheetApp.flush(); Utilities.sleep(2500);
      const v = cache.getRange('A1').getValue();
      cache.getRange('A1').clearContent();
      return typeof v === 'number' ? { tasa: v } : null;
    }
    if (tipo === 'accion_nacional') {
      cache.getRange('A1').setFormula(`=GOOGLEFINANCE("${nombre}","price")`);
      SpreadsheetApp.flush(); Utilities.sleep(2500);
      const v = cache.getRange('A1').getValue();
      cache.getRange('A1').clearContent();
      return typeof v === 'number' ? { precio: v } : null;
    }
    if (tipo === 'accion_internacional') {
      cache.getRange('A1').setFormula(`=GOOGLEFINANCE("${nombre}","price")`);
      cache.getRange('A2').setFormula('=GOOGLEFINANCE("CURRENCY:USDCOP")');
      SpreadsheetApp.flush(); Utilities.sleep(3000);
      const precio = cache.getRange('A1').getValue();
      const tasa   = cache.getRange('A2').getValue();
      cache.getRange('A1:A2').clearContent();
      if (typeof precio === 'number' && typeof tasa === 'number') {
        return { precioUSD: precio, tasaCOP: tasa, precioCOP: precio * tasa };
      }
      return null;
    }
    return null;
  } catch(e) {
    Logger.log('getPrecioActual: ' + e.message);
    return null;
  }
}

// ---- Diagnóstico (correr desde el editor para verificar que todo funciona) ----

function testConexion() {
  try {
    setPresupuesto('2026-01', 'TEST-categoria', 1);
    deletePresupuesto('2026-01', 'TEST-categoria');
    Logger.log('✅ Escritura en Presupuesto: OK');
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
    throw e;
  }
}

// ---- Importación masiva desde hoja "Importar" ----
// Columnas esperadas: mes (YYYY-MM) | tipo | descripcion | categoria | monto | tipo_gasto

function importarDesdeHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('Importar');
  if (!importSheet) {
    SpreadsheetApp.getUi().alert('No encontré una hoja llamada "Importar".\nCréala y llena los datos según la plantilla.');
    return;
  }

  const data = importSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('La hoja "Importar" está vacía o solo tiene encabezados.');
    return;
  }

  const filas = [];
  let errores = 0;
  const errDetail = [];

  for (let i = 1; i < data.length; i++) {
    const [mes, tipo, descripcion, categoria, monto, tipo_gasto] = data[i];
    if (!mes && !descripcion && !monto) continue; // fila vacía

    const mesStr   = String(mes).trim();
    const tipoStr  = String(tipo).trim().toLowerCase();
    const descStr  = String(descripcion).trim();
    const catStr   = String(categoria).trim();
    const montoNum = Number(String(monto).replace(/[^0-9.]/g, ''));
    const tgStr    = String(tipo_gasto || '').trim().toUpperCase();

    if (!/^\d{4}-\d{2}$/.test(mesStr)) {
      errores++; errDetail.push(`Fila ${i+1}: mes inválido "${mes}" — usa formato YYYY-MM`); continue;
    }
    if (!['ingreso','egreso'].includes(tipoStr)) {
      errores++; errDetail.push(`Fila ${i+1}: tipo inválido "${tipo}" — usa "ingreso" o "egreso"`); continue;
    }
    if (!descStr) {
      errores++; errDetail.push(`Fila ${i+1}: descripción vacía`); continue;
    }
    if (isNaN(montoNum) || montoNum <= 0) {
      errores++; errDetail.push(`Fila ${i+1}: monto inválido "${monto}"`); continue;
    }

    filas.push([Utilities.getUuid(), mesStr + '-01', tipoStr, descStr, catStr, montoNum, tgStr]);
  }

  // Escribe todas las filas de una sola vez (mucho más rápido que appendRow en bucle)
  if (filas.length > 0) {
    const txSheet = getSheet('TX');
    const lastRow = txSheet.getLastRow();
    txSheet.getRange(lastRow + 1, 1, filas.length, 7).setValues(filas);
  }

  let msg = `✅ ${filas.length} registro(s) importados correctamente.`;
  if (errores > 0) msg += `\n\n⚠️ ${errores} fila(s) con errores:\n` + errDetail.slice(0, 10).join('\n');
  SpreadsheetApp.getUi().alert(msg);
}
