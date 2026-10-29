/* Vaciar una búsqueda (pedido de Ariel 29/10): borra TODOS los candidatos de la búsqueda y con ellos su historial, contactos y evaluaciones,
   y manda a la papelera de Drive sus archivos (CV, PDF de screening y competencias, postulación e informe; se recuperan 30 días desde la papelera de Drive).
   Solo administradores, y hay que escribir el nombre de la búsqueda para confirmar. No toca la configuración, los formularios, el flyer, las carpetas ni el Banco de talento.
   Queda una sola línea en el Historial («Búsqueda vaciada…») como constancia. */

function vaciarBusquedaInfo_(perfilId) {
  const cfg = getConfig(perfilId), sh = getSheet(cfg);
  const n = sh && sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().filter(r => String(r[0] || '').trim()).length : 0;
  return { nombre: String(cfg.nombre || perfilId), candidatos: n };
}

function vaciarBusqueda_(perfilId, confirmacion, quien) {
  QUIEN_ = quien || '';
  const cfg = getConfig(perfilId), sh = getSheet(cfg);
  const nombre = String(cfg.nombre || perfilId);
  const norm = t => String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
  if (norm(confirmacion) !== norm(nombre)) throw new Error('Para vaciar la búsqueda hay que escribir su nombre tal cual: «' + nombre + '». No borré nada.');
  const ult = sh.getLastRow();
  if (ult < 2) return { ok: true, candidatos: 0, archivos: 0, historial: 0 };
  const nf = ult - 1, ncol = Math.max(sh.getLastColumn(), 40);
  const datos = sh.getRange(2, 1, nf, ncol).getValues();
  const candidatos = datos.filter(r => String(r[1] || '').trim()).length;

  // archivos de Drive de cada candidato (solo los que están enlazados en su fila)
  const ids = {};
  const idDe = u => { u = String(u || ''); const a = u.match(/[?&]id=([-\w]{25,})/) || u.match(/\/d\/([-\w]{25,})/) || u.match(/[-\w]{25,}/); return a ? (a[1] || a[0]) : ''; };
  [8, 13, 17, 31, 33].forEach(c => {
    try {
      const vals = datos.map(r => r[c - 1]);
      const urls = idsCVColumna_(sh, c, nf, vals);
      urls.forEach(u => { const id = idDe(u); if (id) ids[id] = 1; });
    } catch (e) { }
  });
  let archivos = 0;
  Object.keys(ids).forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); archivos++; } catch (e) { } });

  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const quitar = (nomHoja, fn, cols) => {
    const s = ss.getSheetByName(nomHoja);
    if (!s || s.getLastRow() < 2) return 0;
    const nc = s.getLastColumn(), v = s.getRange(2, 1, s.getLastRow() - 1, nc).getValues();
    const quedan = v.filter(r => !fn(r));
    const borradas = v.length - quedan.length;
    if (!borradas) return 0;
    s.getRange(2, 1, v.length, nc).clearContent();
    if (quedan.length) s.getRange(2, 1, quedan.length, nc).setValues(quedan);
    return borradas;
  };
  // evaluaciones: sus respuestas guardadas (.json) también van a la papelera de Drive
  try {
    const se = ss.getSheetByName(HOJA_EVAL_);
    if (se && se.getLastRow() > 1) se.getRange(2, 1, se.getLastRow() - 1, 7).getValues().forEach(r => { if (String(r[1]) === String(perfilId) && r[6]) { try { DriveApp.getFileById(String(r[6])).setTrashed(true); archivos++; } catch (e) { } } });
  } catch (e) { }
  quitar(HOJA_EVAL_, r => String(r[1]) === String(perfilId));
  quitar(HOJA_REG_CONT_, r => String(r[1]) === String(perfilId));
  const historial = quitar(HOJA_HIST_, r => String(r[1]) === String(sh.getName()));

  // la hoja de candidatos queda vacía pero con su formato y sus desplegables
  const rng = sh.getRange(2, 1, nf, sh.getLastColumn());
  rng.clearContent(); rng.clearNote(); rng.setBackground(null);

  try { const k = CacheService.getScriptCache(); k.remove('cfg_all'); k.remove('prev_v2_n'); } catch (e) { }
  registrarHistorial_(sh.getName(), '(búsqueda)', 'Búsqueda vaciada', candidatos + ' candidatos, ' + archivos + ' archivos de Drive a la papelera', quien);
  return { ok: true, candidatos: candidatos, archivos: archivos, historial: historial };
}
