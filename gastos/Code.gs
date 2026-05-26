// ============================================================
// MIS FINANZAS — Google Apps Script Backend
// ============================================================

const SHEETS = {
  TX:      { name: 'Transacciones', headers: ['id','fecha','tipo','descripcion','categoria','monto','tipo_gasto'] },
  BUDGET:  { name: 'Presupuesto',   headers: ['mes','categoria','monto'] },
  MEDEBEN: { name: 'Me_Deben',      headers: ['id','fecha','persona','descripcion','monto','estado'] }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Mis Finanzas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  }
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
    data.categoria, Number(data.monto), data.tipo_gasto || ''
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

function getPresupuesto(mes) {
  return sheetToObjects(getSheet('BUDGET')).filter(p => p.mes === mes);
}

function setPresupuesto(mes, categoria, monto) {
  const sheet = getSheet('BUDGET');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === mes && values[i][1] === categoria) {
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
    if (values[i][0] === mes && values[i][1] === categoria) {
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
  getSheet('MEDEBEN').appendRow([id, data.fecha, data.persona, data.descripcion, Number(data.monto), 'pendiente']);
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
