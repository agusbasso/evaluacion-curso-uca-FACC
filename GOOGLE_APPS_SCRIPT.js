/**
 * Backend para el formulario de evaluación UCA-FACC
 *
 * SETUP (2 minutos):
 * 1. Andá a https://sheets.google.com/ y creá una hoja nueva. Llamala como quieras.
 * 2. En la hoja: Extensiones → Apps Script.
 * 3. Borrá el código que aparece y pegá TODO este archivo.
 * 4. Guardá (Ctrl+S) y dale un nombre al proyecto.
 * 5. Arriba a la derecha: Implementar → Nueva implementación.
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Acceso: CUALQUIERA
 *    - Implementar.
 * 6. Copiá la "URL de la aplicación web" que te da (termina en /exec).
 * 7. En la página de la encuesta, entrá al Dashboard como admin y pegá esa URL
 *    en el recuadro "Backend compartido (Google Apps Script URL)" y Guardar.
 *
 * Listo. Cada respuesta nueva se guarda en la hoja y en el Dashboard se ven todas.
 */

// GET: devuelve todas las respuestas como JSON (usado por el Dashboard)
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  }
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST: recibe una respuesta del formulario y la agrega como fila nueva
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Si la hoja está vacía, escribí headers en base a las keys del primer POST
    let headers;
    if (sheet.getLastRow() === 0) {
      headers = Object.keys(payload);
      sheet.appendRow(headers);
    } else {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      // Agregar columnas nuevas si aparecen keys que no estaban
      const newKeys = Object.keys(payload).filter(k => !headers.includes(k));
      if (newKeys.length) {
        newKeys.forEach(k => {
          sheet.getRange(1, headers.length + 1).setValue(k);
          headers.push(k);
        });
      }
    }

    const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
