// Configuración global
const SHEET_NAME = "Registros";
const TZ = "America/Mexico_City"; // Zona horaria Ciudad de México
const GOAL_PER_DAY = 63; // Meta diaria de llamadas

/**
 * Crea la hoja si no existe con sus encabezados
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ID/Timestamp", "Fecha", "Mes", "Tipo"]);
    // Congelar la primera fila
    sheet.setFrozenRows(1);
    // Formatear los encabezados
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f4f6");
  }
}

/**
 * Sirve la aplicación web HTML
 */
function doGet() {
  setupSheet();
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Dashboard de Productividad')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Obtiene y calcula todos los datos de productividad actuales
 */
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setupSheet(); sheet = ss.getSheetByName(SHEET_NAME); }

  const data = sheet.getDataRange().getValues();
  if (data.length > 0) data.shift(); // Remover encabezados

  const now = new Date();
  const todayStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  const monthStr = Utilities.formatDate(now, TZ, "yyyy-MM");

  let todayCalls = 0;
  let monthCalls = 0;
  let daysWorkedSet = new Set();
  let todaysRecords = []; // Para racha y pico

  data.forEach(row => {
    // Transformar la celda a texto pase lo que pase
    let rowDateStr = row[1];
    if (rowDateStr instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDateStr, TZ, "yyyy-MM-dd");
    } else {
      rowDateStr = String(rowDateStr); 
    }

    let rowMonthStr = row[2];
    if (rowMonthStr instanceof Date) {
      rowMonthStr = Utilities.formatDate(rowMonthStr, TZ, "yyyy-MM");
    } else {
      rowMonthStr = String(rowMonthStr);
    }

    const type = String(row[3]);

    // Cálculos mensuales
    if (rowMonthStr === monthStr) {
      if (type === "Llamada") {
        monthCalls++;
        daysWorkedSet.add(rowDateStr);
      }
    }

    // Cálculos diarios
    if (rowDateStr === todayStr) {
      todaysRecords.push({ timestamp: Number(row[0]), type: type });
      if (type === "Llamada") {
        todayCalls++;
      }
    }
  });

  // Cálculo de la productividad
  const diasLaborados = daysWorkedSet.size > 0 ? daysWorkedSet.size : 1; 
  const metaTotalMensual = GOAL_PER_DAY * diasLaborados;
  let productividad = (monthCalls / metaTotalMensual) * 100;

  // Determinar estatus
  let estatus = "Debemos esforzarnos";
  if (productividad >= 90) estatus = "Buena";
  else if (productividad >= 70) estatus = "Promedio";

  // Nuevas métricas derivadas
  let faltantesHoy = Math.max(0, GOAL_PER_DAY - todayCalls);
  let promDiario = Math.round((monthCalls / diasLaborados) * 10) / 10; // Redondeado a 1 decimal

  // Historial, Racha y Pico
  let last3 = todaysRecords.slice(-3).reverse();
  let recentHistory = last3.map(r => {
    let timeStr = Utilities.formatDate(new Date(r.timestamp), TZ, "HH:mm");
    return `${timeStr} - ${r.type}`;
  });

  const oneHourAgo = now.getTime() - (60 * 60 * 1000);
  let streakCount = todaysRecords.filter(r => r.type === 'Llamada' && r.timestamp >= oneHourAgo).length;

  let hoursCount = {};
  todaysRecords.forEach(r => {
      if (r.type === 'Llamada') {
          let h = Utilities.formatDate(new Date(r.timestamp), TZ, "HH");
          hoursCount[h] = (hoursCount[h] || 0) + 1;
      }
  });
  
  let peakHour = null;
  let maxCallsInHour = 0;
  for (let h in hoursCount) {
      if (hoursCount[h] > maxCallsInHour) {
          maxCallsInHour = hoursCount[h];
          peakHour = h;
      }
  }
  let peakText = "Sin datos";
  if (peakHour) {
      let nextHour = String((Number(peakHour) + 1) % 24).padStart(2, '0');
      peakText = `${peakHour}:00 - ${nextHour}:00`;
  }

  return {
    todayCalls,
    monthCalls,
    diasLaborados,
    productividad: Math.round(productividad * 100) / 100,
    estatus,
    metaTotalMensual,
    faltantesHoy,
    promDiario,
    recentHistory,
    streakCount,
    peakText
  };
}

/**
 * Añade un nuevo registro desde el frontend
 */
function addRecord(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) { setupSheet(); sheet = ss.getSheetByName(SHEET_NAME); }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  const monthStr = Utilities.formatDate(now, TZ, "yyyy-MM");

  sheet.appendRow([now.getTime(), dateStr, monthStr, type]);
  return getData();
}

/**
 * Elimina el último registro (Deshacer)
 */
function undoLast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) { // Evita borrar el encabezado
    sheet.deleteRow(lastRow);
  }
  return getData();
}
