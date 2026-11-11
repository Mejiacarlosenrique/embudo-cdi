/**
 * HOJA DE LEADS · EMBUDO JUAN IMPORTA
 *
 * Recibe los datos del embudo (sheetsWebhook) y los ordena en dos pestañas:
 *  - "Leads": una fila por persona (se actualiza a medida que avanza). Solo entra quien dejó WhatsApp o correo.
 *  - "Eventos": una fila por cada paso del embudo (para ver en qué paso se cae la gente).
 *
 * CÓMO INSTALARLO
 *  1. En la hoja: Extensiones > Apps Script. Borra lo que haya y pega todo este archivo. Guarda.
 *  2. Implementar > Nueva implementación > tipo "Aplicación web".
 *     Ejecutar como: "Yo". Quién tiene acceso: "Cualquier persona". Implementar y autorizar.
 *  3. Copia la URL que termina en /exec y pégala en CONFIG.sheetsWebhook del embudo.
 *  Si cambias este código, usa Implementar > Gestionar implementaciones > editar > "Nueva versión" (así la URL no cambia).
 */

var COLUMNAS_LEADS = [
  "fecha_registro", "ultima_actualizacion", "ultimo_evento",
  "nombre", "telefono", "email",
  "indice", "perfil", "ruta", "califica",
  "p_etapa", "p_freno", "p_ventas",
  "p_motivacion", "p_capital", "p_capital_exacto", "p_decision", "p_compromiso",
  "utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term",
  "fbclid", "gclid", "src", "sck", "pagina", "sid"
];
var COLUMNAS_EVENTOS = ["fecha", "evento", "sid", "nombre", "ruta", "utm_source", "utm_campaign", "utm_content"];

// Estos campos no se borran si un evento posterior llega vacío (ej. la fecha de registro)
var FIJOS = { fecha_registro: 1 };

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ahora = new Date();

    // 1) Registro de eventos
    var ev = hoja_(ss, "Eventos", COLUMNAS_EVENTOS);
    var cabEv = cabecera_(ev, null);
    ev.appendRow(cabEv.map(function (k) { return k === "fecha" ? ahora : (d[k] !== undefined ? d[k] : ""); }));

    // 2) Ficha del lead (una fila por persona, buscada por sid)
    var le = hoja_(ss, "Leads", COLUMNAS_LEADS);
    var cab = cabecera_(le, d);   // si el embudo manda un campo nuevo, se agrega como columna al final
    var fila = buscaSid_(le, cab, d.sid);
    var tieneContacto = d.telefono || d.email;
    if (!fila && !tieneContacto) return ok_();   // aún no deja sus datos: solo queda en "Eventos"

    var actual = fila ? le.getRange(fila, 1, 1, cab.length).getValues()[0] : cab.map(function () { return ""; });
    var nueva = cab.map(function (k, i) {
      if (k === "fecha_registro") return actual[i] || ahora;
      if (k === "ultima_actualizacion") return ahora;
      if (k === "ultimo_evento") return d.evento || actual[i];
      var v = d[k];
      if (FIJOS[k] && actual[i]) return actual[i];
      return v !== undefined && v !== "" ? v : actual[i];
    });
    if (fila) le.getRange(fila, 1, 1, cab.length).setValues([nueva]);
    else le.appendRow(nueva);
    return ok_();
  } catch (err) {
    return ContentService.createTextOutput("error: " + err);
  } finally {
    lock.releaseLock();
  }
}

// Para probar desde el editor: ejecuta "prueba" y revisa la pestaña "Leads"
function prueba() {
  doPost({ postData: { contents: JSON.stringify({
    evento: "lead", sid: "prueba123", nombre: "Lead De Prueba", telefono: "+573000000000", email: "prueba@ejemplo.com",
    indice: 49, perfil: "arranque", ruta: "A", p_etapa: "Nunca he importado",
    utm_source: "facebook", utm_medium: "paid", utm_campaign: "campana-prueba", utm_content: "anuncio-1"
  }) } });
}

function hoja_(ss, nombre, columnas) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.appendRow(columnas);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, columnas.length).setFontWeight("bold");
  }
  return sh;
}

// Devuelve la cabecera; con d, agrega al final las columnas nuevas que lleguen
var NO_COLUMNA = { fecha: 1, evento: 1 };
function cabecera_(sh, d) {
  var cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (d) Object.keys(d).forEach(function (k) {
    if (cab.indexOf(k) < 0 && !NO_COLUMNA[k]) { cab.push(k); sh.getRange(1, cab.length).setValue(k).setFontWeight("bold"); }
  });
  return cab;
}

function buscaSid_(sh, cab, sid) {
  if (!sid || sh.getLastRow() < 2) return 0;
  var col = cab.indexOf("sid") + 1;
  var f = sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(String(sid)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

function ok_() { return ContentService.createTextOutput("ok"); }
