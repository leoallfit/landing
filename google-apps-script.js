// ============================================================
//  ALLFIT — Google Apps Script
//  Guarda cada lead del formulario en Google Sheets
//  y te manda un email con los datos.
//
//  PASOS PARA ACTIVARLO (una sola vez, 3 minutos):
//  1. Abrí sheets.google.com → creá una hoja nueva → llamala "ALLFIT Leads"
//  2. En el menú: Extensiones → Apps Script
//  3. Borrá todo lo que hay y pegá este código completo
//  4. Clic en "Guardar" (ícono de disquete)
//  5. Clic en "Implementar" → "Nueva implementación"
//     - Tipo: Aplicación web
//     - Ejecutar como: Yo (tu cuenta)
//     - Quién tiene acceso: Cualquier usuario
//  6. Clic en "Implementar" → copiá la URL que aparece
//  7. En el admin panel de ALLFIT, pegá esa URL en "Webhook Google Sheets"
// ============================================================

const EMAIL_NOTIFICACION = "allfit.asesorias@gmail.com"; // tu email

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data  = JSON.parse(e.postData.contents);

    // Crear encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Fecha", "Nombre", "Edad", "Experiencia",
        "Objetivo", "Días/sem", "Obstáculo", "WhatsApp", "Estado"
      ]);
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#104a8e").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // Agregar fila con los datos del lead
    sheet.appendRow([
      new Date(),
      data["Nombre"]      || "",
      data["Edad"]        || "",
      data["Experiencia"] || "",
      data["Objetivo"]    || "",
      data["Días"]        || "",
      data["Obstáculo"]   || "",
      data["WhatsApp"]    || "",
      "Nuevo"  // Estado inicial
    ]);

    // Resaltar fila nueva en amarillo claro
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 9).setBackground("#fffde7");

    // Enviar email de notificación
    GmailApp.sendEmail(
      EMAIL_NOTIFICACION,
      "🏋️ Nuevo lead ALLFIT — " + (data["Nombre"] || "sin nombre"),
      "Nuevo contacto desde la landing:\n\n" +
      "Nombre: "      + (data["Nombre"]      || "") + "\n" +
      "Edad: "        + (data["Edad"]        || "") + "\n" +
      "Experiencia: " + (data["Experiencia"] || "") + "\n" +
      "Objetivo: "    + (data["Objetivo"]    || "") + "\n" +
      "Días/sem: "    + (data["Días"]        || "") + "\n" +
      "Obstáculo: "   + (data["Obstáculo"]   || "") + "\n" +
      "WhatsApp: "    + (data["WhatsApp"]    || "") + "\n\n" +
      "Ver todos los leads: https://docs.google.com/spreadsheets/d/" +
      SpreadsheetApp.getActiveSpreadsheet().getId()
    );

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
