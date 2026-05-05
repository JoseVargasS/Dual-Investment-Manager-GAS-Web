/**
 * Aplicación Web de Gestión de Operaciones de Inversión Dual
 * Google Apps Script - Bound Script (Extensiones → Apps Script desde la hoja)
 */

// ============================================
// CONFIGURACIÓN
// ============================================
const SPREADSHEET_ID = "1g9JBdaZ7eAAhaEjUf5pQs2YEzKAR2QYANXAJHMz0oJc";
const DEFAULT_SHEET_NAME = "CAPITAL 1";

// Los datos individuales empiezan en fila 16
const FIRST_DATA_ROW = 16;

// Columnas (1-indexado)
// E=FECHA INICIO, F=FECHA FIN, G=CEX, H=MONTO, I=MONEDA
// J=%APR, K=TIPO, L=PRECIO OBJ.
// M=TIEMPO CEX, N=TIEMPO DÍAS, O=DURACIÓN
// Q=FINAL OBTENIDO, R=INTERÉS, S=TOTAL
// U=MONEDA FINAL, V=APR ACUMULADO, W=APR EFECTIVO DIARIO
const COL = {
  FECHA_INICIO: 5, // E
  FECHA_FIN: 6, // F
  CEX: 7, // G
  MONTO: 8, // H
  MONEDA: 9, // I
  APR: 10, // J
  TIPO: 11, // K
  PRECIO_OBJ: 12, // L
  TIEMPO_D_VAL: 13, // M (Días calc)
  TIEMPO_D_TXT: 14, // N ("dia")
  TIEMPO_H_VAL: 15, // O (Hrs calc)
  TIEMPO_H_TXT: 16, // P ("hrs")
  TIEMPO_DIAS: 17, // Q (Tiempo días decimal)
  INTERES: 18, // R
  FINAL_CALC: 19, // S
  FINAL_OBT_VAL: 20, // T
  FINAL_OBT_MON: 21, // U
  APR_ACUM: 22, // V
  APR_EFECTIVO: 23, // W
};

const NUM_COLS = 19; // Desde E hasta W (23 - 5 + 1)
const COLS_INPUT = 8; // Columnas E-L (entrada manual)
const OP_FONT_FAMILY = "Roboto";
const OP_FONT_SIZE = 10;
const OP_HEADER_ROW = 15;
const OP_HEADER_COL_WIDTHS = [
  105, // E FECHA INICIO
  105, // F FECHA FIN
  50, // G CEX
  70, // H MONTO
  62, // I MONEDA
  76, // J % APR EN CEX
  92, // K TIPO OPERACION
  62, // L PRECIO OBJ.
  38, // M DIAS
  26, // N dia
  38, // O HRS
  26, // P hrs
  74, // Q Tiempo (dias)
  66, // R INTERES
  102, // S FINAL CALCULADO
  68, // T FINAL VAL
  58, // U FINAL MON
  118, // V APR ACUMULADO
  118, // W APR EFECTIVO DIARIO
];

// ============================================
// HELPERS DE FECHA
// ============================================
// Convierte fecha ISO (from datetime-local input) a Date object
function parseFechaISO(fechaISO, tz) {
  const fechaStr = fechaISO.replace("T", " ").substring(0, 16) + ":00";
  return Utilities.parseDate(fechaStr, tz, "yyyy-MM-dd HH:mm:ss");
}

// Calcula fecha de fin según el exchange
function calcFechaFin(fechaInicio, tz, cex) {
  // Bybit: siempre 8 horas (o 1 día + 8 horas si la operación supera 1 día)
  if (cex && cex.toString().toLowerCase() === "bybit") {
    return new Date(fechaInicio.getTime() + 8 * 60 * 60 * 1000);
  }
  // Comportamiento actual: día siguiente a las 3:00 AM
  const diaStr = Utilities.formatDate(
    new Date(fechaInicio.getTime() + 86400000),
    tz,
    "yyyy-MM-dd",
  );
  return Utilities.parseDate(diaStr + " 03:00:00", tz, "yyyy-MM-dd HH:mm:ss");
}

// Formatea porcentaje (decimal → string con %)
function fmtPct(raw) {
  if (typeof raw === "number" && isFinite(raw)) return (raw * 100).toFixed(2);
  const n = parseFloat(String(raw));
  return isNaN(n) ? "0.00" : (n * 100).toFixed(2);
}

// ============================================
// FUNCIONES DE PLANTILLA
// ============================================
// Incluir archivos HTML (Index, Styles, Scripts)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// PUNTO DE ENTRADA WEB
// ============================================
// Se ejecuta al abrir la URL de la web app
function doGet(e) {
  const template = HtmlService.createTemplateFromFile("Index");
  return template
    .evaluate()
    .setTitle("Inversión Dual · Crypto Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

// ============================================
// HELPER: Obtener el spreadsheet
// ============================================
function getSpreadsheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss;
}

// ============================================
// HELPER: Obtener la hoja
// ============================================
// Retorna {sheet, ss} para operar con la hoja de cálculo
function getSheet(sheetName) {
  const name = sheetName || DEFAULT_SHEET_NAME;
  const ss = getSpreadsheet();
  if (!ss) throw new Error("No se pudo obtener la hoja de cálculo");

  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('No se encontró la pestaña "' + name + '"');

  return { sheet, ss };
}

// Retorna lista de nombres de pestañas (Capitales)
function getCapitalsList() {
  try {
    const ss = getSpreadsheet();
    if (!ss) return [DEFAULT_SHEET_NAME];

    const sheets = ss
      .getSheets()
      .filter((s) => !s.isSheetHidden());
    sheets.forEach((sheet) => aplicarFormatoEstructuraCapital(sheet));
    return sheets.map((s) => s.getName());
  } catch (e) {
    Logger.log("Error en getCapitalsList: " + e);
    return [DEFAULT_SHEET_NAME];
  }
}

function aplicarFormatoEstructuraCapital(sheet) {
  if (!sheet) return;

  sheet.setColumnWidths(1, 4, 9);
  for (let row = 9; row <= 14; row++) {
    sheet.setRowHeight(row, 9);
  }
  sheet.setRowHeight(OP_HEADER_ROW, 40);

  OP_HEADER_COL_WIDTHS.forEach((width, idx) => {
    sheet.setColumnWidth(COL.FECHA_INICIO + idx, width);
  });

  sheet
    .getRange(OP_HEADER_ROW, COL.FECHA_INICIO, 1, NUM_COLS)
    .setFontFamily(OP_FONT_FAMILY)
    .setFontSize(OP_FONT_SIZE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  const lastRow = Math.max(sheet.getLastRow(), FIRST_DATA_ROW);
  if (lastRow >= FIRST_DATA_ROW) {
    aplicarFormatoRangoOperaciones(sheet, FIRST_DATA_ROW, lastRow - FIRST_DATA_ROW + 1);
  }
}

function aplicarFormatoRangoOperaciones(sheet, startRow, numRows) {
  if (!sheet || numRows <= 0) return;
  sheet
    .getRange(startRow, COL.FECHA_INICIO, numRows, NUM_COLS)
    .setFontFamily(OP_FONT_FAMILY)
    .setFontSize(OP_FONT_SIZE)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function aplicarFormatoFilaOperacion(sheet, row, moneda, monedaFinal) {
  aplicarFormatoRangoOperaciones(sheet, row, 1);

  sheet.getRange(row, COL.FECHA_INICIO, 1, 2).setNumberFormat("dd/mm/yy hh:mm");
  sheet.getRange(row, COL.APR).setNumberFormat("0.00%");
  sheet.getRange(row, COL.TIEMPO_D_VAL).setNumberFormat("0");
  sheet.getRange(row, COL.TIEMPO_H_VAL).setNumberFormat("0");
  sheet.getRange(row, COL.TIEMPO_DIAS).setNumberFormat("0.00");
  sheet.getRange(row, COL.INTERES, 1, 2).setNumberFormat("$#,##0.00");
  sheet.getRange(row, COL.APR_ACUM, 1, 2).setNumberFormat("0.00%");

  const fmtMonto = moneda === "USDT" ? "$#,##0.00" : "0.000000";
  sheet.getRange(row, COL.MONTO).setNumberFormat(fmtMonto);

  const finalFormat = monedaFinal
    ? monedaFinal === "USDT"
      ? "$#,##0.00"
      : "0.000000"
    : "$#,##0.00";
  sheet.getRange(row, COL.FINAL_OBT_VAL).setNumberFormat(finalFormat);
}

function encontrarPrimeraFilaVacia(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), FIRST_DATA_ROW);
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const values = sheet
    .getRange(FIRST_DATA_ROW, COL.FECHA_INICIO, numRows, 1)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const val = values[i][0];
    if (!val || val.toString().trim() === "") return FIRST_DATA_ROW + i;
  }

  return lastRow + 1;
}

// ============================================
// HELPER: Serializar valor para JSON
// ============================================
// Convierte valores de celda a tipos JSON válidos (Date → string)
function serCell(val, sheetName, tz) {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    if (tz) return Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm:ss");
    try {
      const { ss } = getSheet(sheetName);
      return Utilities.formatDate(
        val,
        ss.getSpreadsheetTimeZone(),
        "yyyy-MM-dd HH:mm:ss",
      );
    } catch (e) {
      return Utilities.formatDate(val, "GMT-5", "yyyy-MM-dd HH:mm:ss");
    }
  }
  if (typeof val === "number") return isNaN(val) || !isFinite(val) ? null : val;
  if (typeof val === "boolean") return val;
  return String(val);
}

// ============================================
// AGREGAR OPERACIÓN
// ============================================
// Crea una nueva fila con los datos de la operación (solo E-L)
// Las columnas M-W se calculan automáticamente en la hoja
function agregarOperacion(datos, sheetName) {
  try {
    const { sheet, ss } = getSheet(sheetName);
    const tz = ss.getSpreadsheetTimeZone();

    // Buscar primera fila vacía desde FIRST_DATA_ROW
    const nuevaFila = encontrarPrimeraFilaVacia(sheet);

    // Calcular fechas de inicio y fin
    const fechaInicio = parseFechaISO(datos.fechaInicio, tz);
    const fechaFin = calcFechaFin(fechaInicio, tz, datos.cex);

    // Datos para columnas E-L (entrada manual del usuario)
    const rowData = [
      fechaInicio,
      fechaFin,
      datos.cex,
      datos.monto,
      datos.moneda,
      datos.apr / 100, // APR como decimal
      datos.tipoOperacion,
      datos.precioObjetivo,
    ];

    // Escribir en la hoja
    const range = sheet.getRange(nuevaFila, COL.FECHA_INICIO, 1, COLS_INPUT);
    range.setValues([rowData]);

    // Insertar fórmulas y valores automáticos para columnas M a W (13 a 23)
    const r = nuevaFila;
    const rowValues = [
      `=INT(Q${r})`, // M: Días calc
      "dia", // N: Texto "dia"
      `=ROUND((Q${r}-M${r})*24, 0)`, // O: Horas calc
      "hrs", // P: Texto "hrs"
      `=F${r}-E${r}`, // Q: Tiempo días decimal
      `=IF(I${r}="USDT", H${r}*J${r}*Q${r}/365, H${r}*J${r}*Q${r}*L${r}/365)`, // R: Interés
      `=IF(I${r}="USDT", H${r}+R${r}, H${r}*L${r}+R${r})`, // S: Final Calc
      "", // T: Final Obtenido Valor (Manual)
      "", // U: Final Obtenido Moneda (Manual)
    ];

    // APR Acumulado (V) y APR Efectivo (W)
    const capInitCell = sheetName === DEFAULT_SHEET_NAME ? "J2" : "F3";
    rowValues.push(
      `=IFERROR((S${r}-${capInitCell})*365/${capInitCell}/CEILING(F${r}-MIN(E:E), 1), 0)`,
    ); // V: APR Acum
    rowValues.push(`=IF(I${r}="USDT", R${r}*365/H${r}, R${r}*365/H${r}/L${r})`); // W: APR Ef

    sheet.getRange(r, 13, 1, rowValues.length).setValues([rowValues]);

    aplicarFormatoEstructuraCapital(sheet);
    aplicarFormatoFilaOperacion(sheet, r, datos.moneda);
    sheet.getRange(r, COL.MONEDA).setValue(datos.moneda);

    return {
      success: true,
      message: "Operación agregada en fila " + nuevaFila,
      fila: nuevaFila,
    };
  } catch (error) {
    Logger.log("Error en agregarOperacion: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// OBTENER OPERACIONES
// ============================================
// Lee todas las operaciones de la hoja y las separa en activas y completadas
function obtenerOperaciones(sheetName) {
  try {
    const { sheet, ss } = getSheet(sheetName);
    const tz = ss.getSpreadsheetTimeZone();
    const lastRow = sheet.getLastRow();

    Logger.log(
      "obtenerOperaciones: lastRow=" +
        lastRow +
        ", FIRST_DATA_ROW=" +
        FIRST_DATA_ROW,
    );

    if (lastRow < FIRST_DATA_ROW)
      return { success: true, activas: [], completadas: [] };

    // Leer todas las filas con datos
    const numRows = lastRow - FIRST_DATA_ROW + 1;
    const values = sheet
      .getRange(FIRST_DATA_ROW, COL.FECHA_INICIO, numRows, NUM_COLS)
      .getValues();

    const activas = [];
    const completadas = [];

    // Procesar cada fila
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const fechaVal = row[0];
      if (!fechaVal || fechaVal.toString().trim() === "") continue;

      // Calcular tiempo formateado
      const daysInt = Math.round(parseFloat(row[8])) || 0;
      const hrsInt = Math.round(parseFloat(row[10])) || 0;
      let tStr = "";
      if (daysInt > 0) tStr += daysInt + (daysInt === 1 ? " día " : " días ");
      if (hrsInt > 0 || (daysInt === 0 && hrsInt === 0))
        tStr += hrsInt + (hrsInt === 1 ? " hora" : " horas");

      // Crear objeto operación con los datos de la fila
      const operacion = {
        fila: FIRST_DATA_ROW + i,
        fechaInicio: serCell(row[0], sheetName, tz),
        fechaFin: serCell(row[1], sheetName, tz),
        cex: serCell(row[2], sheetName, tz),
        monto: serCell(row[3], sheetName, tz),
        moneda: serCell(row[4], sheetName, tz),
        apr: fmtPct(row[5]),
        tipoOperacion: serCell(row[6], sheetName, tz),
        precioObjetivo: serCell(row[7], sheetName, tz),
        tiempoCex: tStr,
        tiempoDias: serCell(row[12], sheetName, tz),
        interes: serCell(row[13], sheetName, tz),
        total: serCell(row[14], sheetName, tz),
        final: serCell(row[15], sheetName, tz),
        monedaFinal: serCell(row[16], sheetName, tz),
        aprAcum: fmtPct(row[17]),
        aprEfectivo: fmtPct(row[18]),
      };

      // Clasificar: tiene MONEDA_FINAL → completada, si no → activa
      const mf = row[16];
      if (mf !== null && mf !== undefined && mf.toString().trim() !== "") {
        completadas.push(operacion);
      } else {
        activas.push(operacion);
      }
    }

    Logger.log(
      "Activas: " + activas.length + ", Completadas: " + completadas.length,
    );
    return { success: true, activas: activas, completadas: completadas };
  } catch (error) {
    Logger.log("Error en obtenerOperaciones: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// ELIMINAR OPERACIÓN
// ============================================
// Elimina la fila de la operación
function eliminarOperacionFila(fila, sheetName) {
  try {
    const { sheet } = getSheet(sheetName);
    sheet.deleteRow(fila);
    SpreadsheetApp.flush();
    return { success: true, message: "Operación eliminada" };
  } catch (error) {
    Logger.log("Error eliminarOperacionFila: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// ACTUALIZAR OPERACIÓN
// ============================================
// Actualiza los datos de una operación existente (solo E-L)
function actualizarOperacionFila(datos, sheetName) {
  try {
    const { sheet, ss } = getSheet(sheetName);
    const tz = ss.getSpreadsheetTimeZone();

    // Calcular nuevas fechas
    const fechaInicio = parseFechaISO(datos.fechaInicio, tz);
    const fechaFin = calcFechaFin(fechaInicio, tz, datos.cex);

    const f = datos.fila;
    const rowData = [
      fechaInicio,
      fechaFin,
      datos.cex,
      datos.monto,
      datos.moneda,
      datos.apr / 100,
      datos.tipoOperacion,
      datos.precioObjetivo,
    ];

    const range = sheet.getRange(f, COL.FECHA_INICIO, 1, COLS_INPUT);
    range.setValues([rowData]);

    // Re-insertar fórmulas por si acaso se borraron
    const r = f;
    const formulas = [
      `=F${r}-E${r}`, // M: Tiempo CEX
      `=F${r}-E${r}`, // N: Tiempo Días
      `=INT(N${r}) & "d " & ROUND((N${r}-INT(N${r}))*24,0) & "h"`, // O: Duración
      "", // P: (vacía)
      `=IF(U${r}="", "", IF(U${r}="USDT", S${r}, S${r}/L${r}))`, // Q: Final Obtenido
      `=IF(I${r}="USDT", H${r}*J${r}*N${r}/365, H${r}*J${r}*N${r}*L${r}/365)`, // R: Interés
      `=IF(I${r}="USDT", H${r}+R${r}, H${r}*L${r}+R${r})`, // S: Total
      "", // T: (vacía)
      "", // U: Moneda Final (manual al confirmar)
    ];
    const capInitCell = sheetName === DEFAULT_SHEET_NAME ? "J2" : "F3";
    formulas.push(
      `=IFERROR((S${r}-${capInitCell})*365/${capInitCell}/(F${r}-MIN(E:E)), 0)`,
    ); // V: APR Acum
    formulas.push(`=IF(I${r}="USDT", R${r}*365/H${r}, R${r}*365/H${r}/L${r})`); // W: APR Ef

    sheet.getRange(r, 13, 1, formulas.length).setFormulas([formulas]);

    aplicarFormatoEstructuraCapital(sheet);
    aplicarFormatoFilaOperacion(sheet, r, datos.moneda);

    SpreadsheetApp.flush();
    return { success: true, message: "Operación actualizada correctamente" };
  } catch (error) {
    Logger.log("Error actualizarOperacionFila: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// COMPLETAR OPERACIÓN
// ============================================
// Marca la operación como completada ingresando la moneda final (columna U)
// Las columnas M-W se calculan automáticamente en la hoja
function completarOperacion(fila, monedaFinal, sheetName) {
  try {
    const { sheet } = getSheet(sheetName);
    const r = fila;
    // T: Valor, U: Moneda
    const valFormula = monedaFinal === "USDT" ? `=S${r}` : `=S${r}/L${r}`;

    sheet.getRange(r, COL.FINAL_OBT_VAL).setFormula(valFormula);
    sheet.getRange(r, COL.FINAL_OBT_MON).setValue(monedaFinal);

    aplicarFormatoEstructuraCapital(sheet);
    aplicarFormatoFilaOperacion(sheet, r, sheet.getRange(r, COL.MONEDA).getValue(), monedaFinal);

    SpreadsheetApp.flush();
    return { success: true, message: "Operación completada correctamente" };
  } catch (error) {
    Logger.log("Error en completarOperacion: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// OBTENER RESUMEN
// ============================================
// Obtiene los datos de los cuadros de resumen de la hoja
function obtenerResumen(sheetName) {
  try {
    const { sheet } = getSheet(sheetName);

    let cuadro1 = [],
      cuadro2 = [],
      cuadro3 = [];
    let notaB6 = "",
      formulaB6 = "",
      notaB4 = "",
      formulaB4 = "";

    // Actualizar dólar antes de leer (para valores frescos)
    GetDollarHouse(sheetName);
    SpreadsheetApp.flush();

    // Obtener valores formateados de cada cuadro
    try {
      cuadro1 = sheet.getRange("A3:B11").getDisplayValues();
    } catch (e) {
      Logger.log("Error A3:B11: " + e);
    }
    try {
      cuadro2 = sheet.getRange("E2:F5").getDisplayValues();
    } catch (e) {
      Logger.log("Error E2:F5: " + e);
    }
    try {
      if (sheet.getName() === DEFAULT_SHEET_NAME) {
        cuadro3 = sheet.getRange("H2:J8").getDisplayValues();
      } else {
        // Para nuevas pestañas, el resumen está en E2:F8
        cuadro3 = sheet.getRange("E2:F8").getDisplayValues();
      }
    } catch (e) {
      Logger.log("Error obteniendo cuadro3: " + e);
    }

    // Obtener notas y fórmulas de celdas editables
    try {
      const cellB6 = sheet.getRange("B6");
      notaB6 = cellB6.getNote() || "";
      formulaB6 = cellB6.getFormula() || "";
    } catch (e) {}
    try {
      const cellB4 = sheet.getRange("B4");
      notaB4 = cellB4.getNote() || "";
      formulaB4 = cellB4.getFormula() || "";
    } catch (e) {}

    Logger.log(
      "Resumen: cuadro1=" +
        cuadro1.length +
        ", cuadro2=" +
        cuadro2.length +
        ", quadro3=" +
        cuadro3.length,
    );

    return {
      success: true,
      cuadro1,
      cuadro2,
      cuadro3,
      notaB6,
      notaB4,
      formulaB6,
      formulaB4,
    };
  } catch (error) {
    Logger.log("Error en obtenerResumen: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// DOLLARHOUSE
// ============================================
// Extrae texto entre dos delimitadores
function extractBetween(text, startStr, endStr) {
  const startIdx = text.indexOf(startStr);
  if (startIdx === -1) return null;
  return text
    .substring(
      startIdx + startStr.length,
      text.indexOf(endStr, startIdx + startStr.length),
    )
    .trim();
}

// Obtiene el tipo de cambio del dólar desde DollarHouse
// Usa caché para evitar solicitudes frecuentes
function GetDollarHouse(sheetName) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("dollar_rate");
  if (cached) return parseFloat(cached);

  const { sheet } = getSheet(sheetName);
  if (!sheet) return null;

  const dollarCell = sheet.getName() === DEFAULT_SHEET_NAME ? "F2" : "F7";

  try {
    const resp = UrlFetchApp.fetch("https://app.dollarhouse.pe/", {
      method: "get",
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" },
    });
    if (resp.getResponseCode() !== 200) {
      sheet.getRange(dollarCell).setValue("N/A");
      return null;
    }

    const html = resp.getContentText();
    const buy = extractBetween(html, 'id="buy-exchange-rate">', "<");
    const buyNum = buy ? parseFloat(buy.replace(/[^0-9.]/g, "")) : null;

    if (buyNum && !isNaN(buyNum)) {
      sheet.getRange(dollarCell).setValue(buyNum);
      cache.put("dollar_rate", buyNum.toString(), 60);
      return buyNum;
    } else {
      sheet.getRange(dollarCell).setValue("N/A");
      return null;
    }
  } catch (e) {
    Logger.log("Error en GetDollarHouse: " + e);
    return null;
  }
}

// ============================================
// ACTUALIZAR CELDA
// ============================================
// Actualiza una celda específica con valor o fórmula
function actualizarCelda(celda, valor, nota, sheetName) {
  try {
    const { sheet } = getSheet(sheetName);
    const rango = sheet.getRange(celda || "B6");

    // Si empieza con "=", guardar como fórmula
    if (typeof valor === "string" && valor.startsWith("=")) {
      rango.setFormula(valor);
    } else {
      const numVal = parseFloat(String(valor).replace(/[$,\s]/g, ""));
      rango.setValue(isNaN(numVal) ? valor : numVal);
    }
    if (nota !== undefined) rango.setNote(nota);
    SpreadsheetApp.flush();
    return { success: true, message: "Valor actualizado" };
  } catch (error) {
    Logger.log("Error en actualizarCelda: " + error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// CREAR NUEVO CAPITAL (NUEVA PESTAÑA)
// ============================================
function crearNuevoCapital(datos) {
  try {
    const ss = getSpreadsheet();
    if (!ss) throw new Error("No se pudo acceder a la hoja de cálculo");

    const nombre = datos.nombre;
    if (ss.getSheetByName(nombre)) {
      return {
        success: false,
        message: "Ya existe una pestaña con ese nombre",
      };
    }

    const sheet = ss.insertSheet(nombre);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(ss.getNumSheets());

    // 1. Estilo General de la Hoja (Fondo oscuro total)
    sheet
      .getRange("1:1000")
      .setBackground("#111111")
      .setFontColor("#e0e0e0")
      .setFontFamily("Roboto");

    // 2. Configurar resumen en E2:F8
    const summaryData = [
      ["Capital Inicial (S/)", datos.soles],
      ["Capital Inicial ($)", datos.dolares],
      [
        "NET PROFIT",
        "=ROUND(IFERROR(INDEX(S:S,MATCH(9.99999999999999E+307,S:S))-F3,0), 2)",
      ],
      ["Promedio Diario", "=IFERROR(F4 / CEILING(MAX(E:E)-MIN(E:E), 1), 0)"],
      ["%APR Promedio", "=AVERAGE(W16:W1000)"],
      ["Tipo Cambio Venta", ""],
      ["Capital Final (S/)", "=(F3+F4)*F7"],
    ];
    const summaryRange = sheet.getRange("E2:F8");
    summaryRange.setValues(summaryData);
    summaryRange
      .setBackground("#1a1a1a")
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#333333",
        SpreadsheetApp.BorderStyle.SOLID,
      );

    // Resaltar Net Profit
    sheet.getRange("E4:F4").setFontColor("#00ff94").setFontWeight("bold");
    sheet.getRange("F4").setFontSize(14);

    sheet.getRange("F2").setNumberFormat('"S/"#,##0.00');
    sheet.getRange("F3").setNumberFormat("$#,##0.00");
    sheet.getRange("F4:F5").setNumberFormat("$#,##0.00");
    sheet.getRange("F6").setNumberFormat("0.00%");
    sheet.getRange("F7").setNumberFormat("0.000");
    sheet.getRange("F8").setNumberFormat('"S/"#,##0.00');

    // 3. Configurar encabezados de tabla (E15:W15)
    const headers = [
      "FECHA INICIO",
      "FECHA FIN",
      "CEX",
      "MONTO",
      "MONEDA",
      "% APR EN CEX",
      "TIPO OPERACIÓN",
      "PRECIO OBJ.",
      "DIAS",
      "dia",
      "HRS",
      "hrs",
      "Tiempo (días)",
      "INTERÉS",
      "FINAL CALCULADO",
      "FINAL VAL",
      "FINAL MON",
      "APR ACUMULADO",
      "APR EFECTIVO DIARIO",
    ];
    const headerRange = sheet.getRange(15, 5, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange
      .setBackground("#222222")
      .setFontColor("#00ff94")
      .setFontWeight("bold");
    headerRange.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#444444",
      SpreadsheetApp.BorderStyle.SOLID,
    );
    headerRange.setHorizontalAlignment("center");

    // Bordes para el área de datos
    sheet
      .getRange("E16:W1000")
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#222222",
        SpreadsheetApp.BorderStyle.SOLID,
      );

    aplicarFormatoEstructuraCapital(sheet);

    // Actualizar tipo de cambio inicial
    GetDollarHouse(nombre);

    SpreadsheetApp.flush();
    return {
      success: true,
      message: "Pestaña '" + nombre + "' creada correctamente",
    };
  } catch (error) {
    Logger.log("Error en crearNuevoCapital: " + error);
    return { success: false, message: error.toString() };
  }
}
