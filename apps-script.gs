// Grabados Carey: recibe los pedidos de la web y los anota en esta planilla.
// Va en Extensiones > Apps Script de la planilla "Pedidos GRABADOS".

const HOJA = ""; // nombre de la pestaña; vacío = la primera

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    if (p.web_extra) return ok(); // campo trampa: si viene lleno es un bot

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = HOJA ? ss.getSheetByName(HOJA) : ss.getSheets()[0];
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    const fecha = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "yyyy-MM-dd HH:mm:ss");
    const datos = {
      "Submission ID": "WEB-" + Date.now().toString(36).toUpperCase(),
      "Respondent ID": "web",
      "Submitted at": fecha,
      "¿Qué producto querés grabar?": limpiar(p.producto),
      "¿Cuántos?": limpiar(p.cantidad),
      "Tu idea de diseño": limpiar(p.diseno),
      "Envío": limpiar(p.envio),
      "Nombre": limpiar(p.nombre),
      "WhatsApp": limpiar(p.whatsapp)
    };

    // arma la fila respetando el orden de tus columnas (Pagado y Entregado quedan vacías)
    const fila = headers.map(h => datos[String(h).trim()] ?? "");
    sh.appendRow(fila);
    return ok();
  } finally {
    lock.releaseLock();
  }
}

function limpiar(v) {
  // corta textos larguísimos y evita que algo se interprete como fórmula
  let s = String(v || "").slice(0, 300).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function ok() {
  return ContentService.createTextOutput("ok");
}
