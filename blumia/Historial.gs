/* Historial de quién toca cada cosa (como el de PAMI): una fila por acción en la pestaña «Historial»
   [fecha, hoja, candidato, tipo, detalle, quién]. «hoja» es el nombre de la pestaña de la búsqueda. */
const HOJA_HIST_ = 'Historial';
var QUIEN_ = ''; // quién dispara la acción en esta ejecución (lo setea cada función pública)

function quienHist_(q) {
  if (typeof SES_ !== 'undefined' && SES_ && SES_.nombre) return SES_.nombre; // el usuario que ingresó (no se puede falsear desde la página)
  q = String(q || '').trim() || String(QUIEN_ || '').trim();
  if (!q) { try { q = Session.getActiveUser().getEmail() || ''; } catch (e) { } }
  return q || 'sistema';
}

function registrarHistorial_(hoja, candidato, tipo, detalle, quien) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
    let sh = ss.getSheetByName(HOJA_HIST_);
    if (!sh) {
      sh = ss.insertSheet(HOJA_HIST_);
      sh.getRange(1, 1, 1, 6).setValues([['fecha', 'hoja', 'candidato', 'tipo', 'detalle', 'quien']]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#fff');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(hoja), String(candidato || ''), String(tipo), String(detalle || '').slice(0, 480), quienHist_(quien)]);
  } catch (e) { }
}

function leerHistorial_(hoja, candidato, limite) {
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(HOJA_HIST_);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(sh.getLastRow() - 1, 4000);
  const d = sh.getRange(sh.getLastRow() - n + 1, 1, n, 6).getValues();
  const out = [];
  for (let i = d.length - 1; i >= 0 && out.length < (limite || 100); i--) {
    if (d[i][1] !== hoja) continue;
    if (candidato && String(d[i][2]).trim() !== candidato) continue;
    out.push({ fecha: Utilities.formatDate(new Date(d[i][0]), 'GMT-3', 'dd/MM/yyyy HH:mm'), candidato: String(d[i][2]), tipo: String(d[i][3]), detalle: String(d[i][4]), quien: String(d[i][5]) });
  }
  return out;
}

function getHistorialCandidato_(perfilId, fila) {
  const sh = getSheet(getConfig(perfilId));
  return leerHistorial_(sh.getName(), String(sh.getRange(fila, 2).getValue() || '').trim(), 100);
}

function getHistorialBusqueda_(perfilId) {
  return leerHistorial_(getSheet(getConfig(perfilId)).getName(), '', 150);
}

// último movimiento de cada candidato de una hoja: {nombre: milisegundos}
function ultimosMovimientos_(hoja) {
  const out = {};
  try {
    const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(HOJA_HIST_);
    if (!sh || sh.getLastRow() < 2) return out;
    const n = Math.min(sh.getLastRow() - 1, 6000);
    sh.getRange(sh.getLastRow() - n + 1, 1, n, 4).getValues().forEach(r => {
      if (r[1] !== hoja || !r[2]) return;
      const t = new Date(r[0]).getTime();
      if (t > (out[String(r[2]).trim()] || 0)) out[String(r[2]).trim()] = t;
    });
  } catch (e) { }
  return out;
}
