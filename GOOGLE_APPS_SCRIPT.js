/**
 * Backend para el formulario de evaluación UCA-FACC
 *
 * INSTRUCCIONES:
 * 1. Pegá TODO este archivo en tu Apps Script (reemplazá lo que había).
 * 2. Ctrl+S para guardar.
 * 3. En la barra superior, seleccioná la función "setupSheet" y clickeá ▶ Ejecutar.
 *    Te va a pedir permisos: Autorizar → tu cuenta → "Configuración avanzada" → Ir a…
 *    Esto renombra las columnas viejas (q1, q2, ...) a los títulos reales de las
 *    preguntas y aplica el formato lindo (freeze, negrita, filas alternadas, filtros).
 * 4. Implementar → Administrar implementaciones → editá la actual (lápiz) → Versión:
 *    "Nueva versión" → Implementar. Sin esto, los cambios al doPost NO se aplican.
 */

// Mapa qID → título de la pregunta (deben coincidir con index.html)
const QUESTIONS = {
  q1:  '¿Considera que el contenido del Programa es relevante?',
  q2:  '¿Cómo califica el aporte del Programa?',
  q3:  '¿Considera que este Programa contribuye a mejorar su conocimiento sobre la temática de la seguridad?',
  q4:  '¿Está satisfecho con el conocimiento adquirido durante el Programa?',
  q5:  '¿Cuál cree que fue la cuestión que contribuyó MÁS?',
  q6:  '¿Cuál cree que fue la cuestión que contribuyó MENOS?',
  q7:  '¿Hay algún tema que debería abordarse en futuras ocasiones?',
  q8:  '¿Le interesaría participar en otros cursos sobre el mismo tema?',
  q9:  '¿El tiempo asignado para la reflexión grupal fue suficiente?',
  q10: '¿El tiempo asignado a cada tema fue suficiente?',
  q11: 'Utilidad de los conocimientos',
  q12: 'Calificación general de las exposiciones docentes',
  q13: 'Prof. Florencia Millet (Clase 1)',
  q14: 'Prof. María Belén Hidalgo (Clase 1)',
  q15: 'Prof. Eduardo Valobra (Clase 1)',
  q16: 'Prof. Arturo Castro (Clase 2)',
  q17: 'Prof. Daniel Viola (Clase 2)',
  q18: 'Prof. Enrique del Carril (Clase 3)',
  q19: 'Prof. Carlos Cofiño (Clases 3 y 4)',
  q20: 'Prof. Jorge de Lucio (Clases 3 y 4)',
  q21: 'Prof. Gabriel Curi (Clase 4)',
  q22: 'Prof. Hernán Cappiello (Clase 4)',
  q23: 'Prof. Ignacio Forconi (Clase 5)',
  q24: 'Prof. Victoria García Huidobro (Clase 5)',
  q25: 'Prof. Lucas Ramírez Bosco (Clase 5)',
  q26: 'Prof. Claudia Gómez Prieto (Clase 6)',
  q27: 'Prof. Jorge Vitti (Clase 6)',
  q28: 'Prof. Martín Rosenbaum (Clase 7)',
  q29: 'Prof. Ana Lamas (Clase 7)',
  q30: 'Prof. Daniela Rojo (Clase 7)',
  q31: 'Prof. Leila Devia (Clase 7)',
  q32: 'Prof. Maximiliano Méndez (Clase 8)',
  q33: 'Prof. Ricardo Nievas (Clase 9)',
  q34: 'Prof. Facundo Améndola (Clase 9)',
  q35: '¿El Programa cumplió con sus expectativas?',
  q36: '¿Qué le pareció la extensión del curso?',
  q37: 'Evaluación general del curso',
  q38: 'Calificación de la coordinación del curso',
  q39: '¿Recomendaría este curso? (probabilidad)',
  q41: '¿Qué tipo de cursada prefiere?',
  q42: '¿Qué horario de cursada prefiere?',
  q40: 'Comentarios adicionales'
};

const TIMESTAMP_HEADER = 'Fecha';
const HEADER_ORDER = ['timestamp'].concat(Object.keys(QUESTIONS)); // qIDs internos
const DISPLAY_HEADERS = [TIMESTAMP_HEADER].concat(HEADER_ORDER.slice(1).map(k => QUESTIONS[k]));

// -------- Endpoints --------

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return _json([]);
  }
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return _json(rows);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Asegurar headers correctos (títulos, no qN)
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, DISPLAY_HEADERS.length).setValues([DISPLAY_HEADERS]);
      _formatHeader(sheet);
    }

    // Escribir fila alineada al orden fijo
    const row = HEADER_ORDER.map(key => {
      if (key === 'timestamp') {
        const ts = payload.timestamp || new Date().toISOString();
        return new Date(ts);
      }
      return payload[key] !== undefined ? payload[key] : '';
    });
    sheet.appendRow(row);

    return _json({ok: true});
  } catch (err) {
    return _json({ok: false, error: err.message});
  }
}

// -------- Migración + formato (correr una vez desde el editor) --------

function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCol = Math.max(sheet.getLastColumn(), DISPLAY_HEADERS.length);
  const lastRow = Math.max(sheet.getLastRow(), 1);

  // 1. Escribir headers con títulos
  sheet.getRange(1, 1, 1, DISPLAY_HEADERS.length).setValues([DISPLAY_HEADERS]);

  // 2. Limpiar cualquier header extra a la derecha
  if (lastCol > DISPLAY_HEADERS.length) {
    sheet.getRange(1, DISPLAY_HEADERS.length + 1, 1, lastCol - DISPLAY_HEADERS.length).clearContent();
  }

  // 3. Formatear columna Fecha (col 1) como fecha ISO amigable
  if (lastRow > 1) {
    const dateRange = sheet.getRange(2, 1, lastRow - 1, 1);
    // Convertir strings ISO a Date reales
    const vals = dateRange.getValues();
    const converted = vals.map(r => {
      if (r[0] instanceof Date) return r;
      if (typeof r[0] === 'string' && r[0].length > 10) {
        const d = new Date(r[0]);
        return isNaN(d.getTime()) ? r : [d];
      }
      return r;
    });
    dateRange.setValues(converted);
    dateRange.setNumberFormat('dd/mm/yyyy hh:mm');
  }

  _formatHeader(sheet);
  _formatBody(sheet);
}

function _formatHeader(sheet) {
  const range = sheet.getRange(1, 1, 1, DISPLAY_HEADERS.length);
  range
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f3864')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 60);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function _formatBody(sheet) {
  const lastRow = sheet.getLastRow();
  const nCols = DISPLAY_HEADERS.length;

  // Anchos: primero fecha ancha, resto 220 excepto columnas de texto libre 380
  const widths = {1: 150, 2: 260}; // Fecha, q1
  for (let c = 3; c <= nCols; c++) widths[c] = 200;
  // Preguntas de texto libre (más ancho)
  const textCols = ['q5','q6','q7','q40'].map(k => HEADER_ORDER.indexOf(k) + 1);
  textCols.forEach(c => widths[c] = 360);
  Object.keys(widths).forEach(c => sheet.setColumnWidth(Number(c), widths[c]));

  // Wrap en las columnas de texto libre
  if (lastRow >= 2) {
    textCols.forEach(c => {
      sheet.getRange(2, c, lastRow - 1, 1).setWrap(true).setVerticalAlignment('top');
    });
  }

  // Banded rows (filas alternadas)
  const existingBandings = sheet.getBandings();
  existingBandings.forEach(b => b.remove());
  if (lastRow >= 1) {
    const banding = sheet.getRange(1, 1, lastRow, nCols)
      .applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
    banding.setHeaderRowColor('#1f3864');
    banding.setFirstRowColor('#eef2f7');
    banding.setSecondRowColor('#ffffff');
  }

  // Filtros
  if (sheet.getFilter()) sheet.getFilter().remove();
  if (lastRow >= 1) {
    sheet.getRange(1, 1, Math.max(lastRow, 1), nCols).createFilter();
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
