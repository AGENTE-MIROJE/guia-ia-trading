/**
 * Base de datos de inscritos de la guía «IA en el trading, sin humo».
 * Va pegado a la hoja de cálculo «Inscritos · Guía IA» (Extensiones → Apps Script)
 * y se publica como aplicación web: Ejecutar como «Yo», acceso «Cualquier persona».
 * El formulario de la landing le envía los mismos campos que a FormSubmit.
 */
const HOJA = 'Inscritos';
const ENCABEZADO = ['N°', 'Fecha y hora (Bogotá)', 'Nombre', 'Correo', 'Autorización de datos', 'Origen'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p._honey) return respuesta_({ ok: true });            // robot: se descarta sin avisar
  const nombre = limpio_(p.nombre, 80);
  const correo = limpio_(p.email, 120);
  if (!nombre || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return respuesta_({ ok: false, error: 'datos incompletos' });

  const candado = LockService.getScriptLock();
  candado.waitLock(15000);                                  // dos registros a la vez no se pisan
  try {
    const hoja = hoja_();
    const n = hoja.getLastRow();                            // la fila 1 es el encabezado
    hoja.appendRow([
      n,
      Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss'),
      nombre,
      correo,
      limpio_(p.autorizacion_datos, 60),
      limpio_(p.origen, 60),
    ]);
    return respuesta_({ ok: true, n: n });
  } finally {
    candado.releaseLock();
  }
}

function doGet() {
  return respuesta_({ ok: true, servicio: 'inscritos guía IA' });
}

function hoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA, 0);
    hoja.appendRow(ENCABEZADO);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, ENCABEZADO.length).setFontWeight('bold');
  }
  return hoja;
}

// Recorta y evita que un texto que empiece por = + - @ se interprete como fórmula.
function limpio_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function respuesta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
