/**
 * Registro de la encuesta a estilistas (versión 7 en adelante).
 * Pega este código en la hoja "Encuesta estilistas" > Extensiones > Apps Script,
 * reemplazando todo el código anterior, y publica una NUEVA VERSIÓN de la implementación existente.
 *
 * Cada versión de la encuesta escribe en su propia pestaña ("Respuestas v7").
 * Los encabezados se crean solos a partir de los datos recibidos; si la encuesta
 * agrega una pregunta nueva, la columna se añade al final sin mover las demás.
 */
const MAX_LEN = 1000;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse(e.postData.contents || "{}");

    // Filtro anti-spam: campo oculto lleno = robot
    if (data.hp) return out_({ ok: true });

    // Evita registrar dos veces el mismo envío
    const cache = CacheService.getScriptCache();
    if (data.sid) {
      if (cache.get("sid_" + data.sid)) return out_({ ok: true, dup: true });
      cache.put("sid_" + data.sid, "1", 21600);
    }

    const version = String(data.v || "sin_version").replace(/[^\w.-]/g, "");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = "Respuestas v" + version;
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);

    delete data.hp;
    data.recibido = new Date();

    // Encabezados
    let headers = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    const nuevos = Object.keys(data).filter(k => headers.indexOf(k) === -1);
    if (nuevos.length) {
      sh.getRange(1, headers.length + 1, 1, nuevos.length).setValues([nuevos]);
      headers = headers.concat(nuevos);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    const row = headers.map(h => {
      const v = data[h];
      if (v === undefined || v === null) return "";
      if (v instanceof Date) return v;
      const s = String(v).slice(0, MAX_LEN);
      return /^[=+\-@]/.test(s) ? "'" + s : s; // evita fórmulas inyectadas
    });
    sh.appendRow(row);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  return out_({ ok: true, servicio: "Encuesta estilistas" });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
