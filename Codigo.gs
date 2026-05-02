// Configuración global
const SHEET_NAME = "Registros";
const TZ = "America/Mexico_City"; // Zona horaria Ciudad de México
const GOAL_PER_DAY = 63; // Meta diaria de llamadas (incluye transferencias)
const HEADERS = ["ID/Timestamp", "Fecha", "Mes", "Tipo", "Monto", "Pedido"];
const VALID_TYPES = ["Llamada", "Transferencia", "Venta"];
const CALL_TYPES = ["Llamada", "Transferencia"]; // Cuentan para la meta

/**
 * Crea la hoja si no existe con sus encabezados (6 columnas).
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#f3f4f6");
  }
  return sheet;
}

/**
 * Sirve la aplicación web HTML.
 */
function doGet() {
  setupSheet();
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Dashboard de Productividad')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper interno: obtiene la hoja, creándola si hace falta.
 */
function _getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || setupSheet();
}

/**
 * Normaliza una celda de fecha a string "yyyy-MM-dd".
 */
function _toDateStr_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, TZ, "yyyy-MM-dd");
  return String(value || "");
}

/**
 * Normaliza una celda de mes a string "yyyy-MM".
 */
function _toMonthStr_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, TZ, "yyyy-MM");
  return String(value || "");
}

/**
 * Respuesta vacía cuando aún no hay datos.
 */
function _emptyResponse_() {
  return {
    todayCalls: 0,
    monthCalls: 0,
    monthTransfers: 0,
    diasLaborados: 1,
    productividad: 0,
    estatus: "Debemos esforzarnos",
    metaTotalMensual: GOAL_PER_DAY,
    goalPerDay: GOAL_PER_DAY,
    faltantesHoy: GOAL_PER_DAY,
    promDiario: 0,
    recentHistory: [],
    streakCount: 0,
    peakText: "Sin datos"
  };
}

/**
 * Obtiene y calcula todos los datos de productividad actuales.
 * Lee solo las columnas necesarias (A-D) para mejor performance.
 */
function getData() {
  const sheet = _getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return _emptyResponse_();

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  const now = new Date();
  const todayStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  const monthStr = Utilities.formatDate(now, TZ, "yyyy-MM");

  let todayCalls = 0;
  let monthCalls = 0;
  let monthTransfers = 0;
  const daysWorkedSet = new Set();
  const todaysRecords = [];

  data.forEach(row => {
    const rowDateStr = _toDateStr_(row[1]);
    const rowMonthStr = _toMonthStr_(row[2]);
    const type = String(row[3] || "");

    if (rowMonthStr === monthStr) {
      // Cualquier registro válido marca el día como laborado.
      if (VALID_TYPES.indexOf(type) !== -1) {
        daysWorkedSet.add(rowDateStr);
      }
      // Llamadas + Transferencias suman para la meta.
      if (CALL_TYPES.indexOf(type) !== -1) {
        monthCalls++;
      }
      if (type === "Transferencia") {
        monthTransfers++;
      }
    }

    if (rowDateStr === todayStr) {
      todaysRecords.push({ timestamp: Number(row[0]), type: type });
      if (CALL_TYPES.indexOf(type) !== -1) {
        todayCalls++;
      }
    }
  });

  const diasLaborados = daysWorkedSet.size > 0 ? daysWorkedSet.size : 1;
  const metaTotalMensual = GOAL_PER_DAY * diasLaborados;
  const productividad = (monthCalls / metaTotalMensual) * 100;

  let estatus = "Debemos esforzarnos";
  if (productividad >= 90) estatus = "Buena";
  else if (productividad >= 70) estatus = "Promedio";

  const faltantesHoy = Math.max(0, GOAL_PER_DAY - todayCalls);
  const promDiario = Math.round((monthCalls / diasLaborados) * 10) / 10;

  const last3 = todaysRecords.slice(-3).reverse();
  const recentHistory = last3.map(r => {
    const timeStr = Utilities.formatDate(new Date(r.timestamp), TZ, "HH:mm");
    return `${timeStr} - ${r.type}`;
  });

  const oneHourAgo = now.getTime() - (60 * 60 * 1000);
  const streakCount = todaysRecords.filter(r =>
    CALL_TYPES.indexOf(r.type) !== -1 && r.timestamp >= oneHourAgo
  ).length;

  const hoursCount = {};
  todaysRecords.forEach(r => {
    if (CALL_TYPES.indexOf(r.type) !== -1) {
      const h = Utilities.formatDate(new Date(r.timestamp), TZ, "HH");
      hoursCount[h] = (hoursCount[h] || 0) + 1;
    }
  });

  let peakHour = null;
  let maxCallsInHour = 0;
  for (const h in hoursCount) {
    if (hoursCount[h] > maxCallsInHour) {
      maxCallsInHour = hoursCount[h];
      peakHour = h;
    }
  }
  let peakText = "Sin datos";
  if (peakHour) {
    const nextHour = String((Number(peakHour) + 1) % 24).padStart(2, '0');
    peakText = `${peakHour}:00 - ${nextHour}:00`;
  }

  return {
    todayCalls,
    monthCalls,
    monthTransfers,
    diasLaborados,
    productividad: Math.round(productividad * 100) / 100,
    estatus,
    metaTotalMensual,
    goalPerDay: GOAL_PER_DAY,
    faltantesHoy,
    promDiario,
    recentHistory,
    streakCount,
    peakText
  };
}

/**
 * Añade un nuevo registro desde el frontend.
 * Valida el tipo y escribe las 6 columnas para mantener consistencia con la hoja.
 */
function addRecord(type) {
  if (VALID_TYPES.indexOf(type) === -1) {
    throw new Error("Tipo de registro no válido: " + type);
  }

  const sheet = _getSheet_();
  const now = new Date();
  const dateStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  const monthStr = Utilities.formatDate(now, TZ, "yyyy-MM");

  sheet.appendRow([now.getTime(), dateStr, monthStr, type, 0, ""]);
  return getData();
}

/**
 * Elimina el último registro (Deshacer).
 * Valida que el registro pertenezca al día actual antes de borrarlo.
 */
function undoLast() {
  const sheet = _getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return getData();

  const lastRowData = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];
  const rowDateStr = _toDateStr_(lastRowData[1]);
  const todayStr = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");

  if (rowDateStr !== todayStr) {
    throw new Error("Solo puedes deshacer registros del día actual.");
  }

  sheet.deleteRow(lastRow);
  return getData();
}
