// Clasificación única de estados (la usan el Informe de proceso, la tabla del Resumen y, copiada con toString(), el panel 📊 del kanban).
// lv = etapa más avanzada a la que llegó: 0 CV · 1 screening · 2 competencias · 3 screening con la empresa · 4 entrevista final · 5 contratado
function metEstado_(estado) {
  const l = String(estado || '').trim().toLowerCase();
  if (/^eliminado/.test(l)) return null;
  let t = 'act', lv, d = '';
  if (l === 'contratado') { t = 'cont'; lv = 5; }
  else if (/^descartad/.test(l)) {
    t = 'desc';
    if (/post screening/.test(l)) { lv = 1; d = 'post'; }
    else if (/en competencias/.test(l)) { lv = 2; d = 'comp'; }
    else if (/screening empresa|presencial empresa|en empresa/.test(l)) { lv = /screening empresa/.test(l) ? 3 : 4; d = 'emp'; }
    else { lv = 0; d = 'cv'; }
  } else if (/reserva/.test(l)) {
    t = 'res'; lv = /screening empresa/.test(l) ? 3 : /presencial empresa/.test(l) ? 4 : 2;
  } else {
    lv = /competencias/.test(l) ? 2 : /screening empresa/.test(l) ? 3 : /presencial empresa|avanza a empresa|entrevista empresa|en cliente/.test(l) ? 4 : /screening/.test(l) ? 1 : 0;
  }
  return { t: t, lv: lv, d: d, scr: lv >= 1, comp: lv >= 2, emp: lv >= 3, res: t === 'res' };
}

// contadores de una búsqueda con las mismas definiciones que el Informe de proceso
function contarBusqueda_(filas) {
  const r = { cv: 0, scr: 0, comp: 0, emp: 0, res: 0, dsin: 0, dpost: 0, dcomp: 0 };
  filas.forEach(f => {
    if (!String(f[1]).trim()) return;
    const m = metEstado_(f[8]);
    if (!m) return;
    r.cv++;
    if (m.scr) r.scr++;
    if (m.comp) r.comp++;
    if (m.emp) r.emp++;
    if (m.res) r.res++;
    if (m.d === 'cv') r.dsin++;
    else if (m.d === 'post') r.dpost++;
    else if (m.d === 'comp' || m.d === 'emp') r.dcomp++;
  });
  return r;
}

// una fila por búsqueda de la misma empresa que perfilId
function resumenBusquedas(perfilId) {
  const emp = String(getConfig(perfilId).empresa || '');
  const out = [];
  listarPerfiles().forEach(p => {
    let cfg, sh;
    try { cfg = getConfig(p.id); sh = getSheet(cfg); } catch (e) { return; }
    if (String(cfg.empresa || '') !== emp || !sh || sh.getLastRow() < 2) return;
    const c = contarBusqueda_(sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues());
    out.push({ id: p.id, nombre: String(cfg.nombre || p.nombre || p.id), cv: c.cv, scr: c.scr, comp: c.comp, emp: c.emp, res: c.res, desc: c.dsin + c.dpost + c.dcomp });
  });
  return out;
}

// escribe un campo de la fila de la búsqueda en Config (crea la columna si falta)
function setConfigCampo_(perfilId, campo, valor) {
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const data = sh.getDataRange().getValues();
  const fila = data.findIndex((r, i) => i > 0 && r[0] === perfilId) + 1;
  if (!fila) throw new Error('Perfil no encontrado: ' + perfilId);
  let col = data[0].indexOf(campo) + 1;
  if (!col) { col = sh.getLastColumn() + 1; sh.getRange(1, col).setValue(campo); }
  sh.getRange(fila, col).setNumberFormat('@').setValue(String(valor));
}

// banda salarial, industrias y fuentes que ya están cargadas en el Informe de proceso de esta búsqueda
function datosInformeBusqueda(perfilId) {
  const cfg = getConfig(perfilId);
  let banda = String(cfg.banda_salarial || '').replace(/\D/g, '');
  // Director Comercial (Energía y Minería): búsqueda de $ 13.000.000 mensuales; se carga sola la primera vez (cuando todavía no existe la columna) y se cambia desde ⚙ Configurar
  if (!banda && perfilId === 'director-comercial' && cfg.banda_salarial === undefined) { banda = '13000000'; try { setConfigCampo_(perfilId, 'banda_salarial', banda); } catch (e) { } }
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(hojaProceso_(empresaKey_(cfg)));
  const nombre = (listarPerfiles().filter(p => p.id === perfilId)[0] || {}).nombre;
  let industrias = '';
  if (sh && sh.getLastRow() > 1) {
    const f = sh.getDataRange().getValues().filter(r => r[0] === nombre)[0];
    if (f) { industrias = String(f[13] || ''); if (!banda) banda = String(f[9] || '').replace(/\D/g, ''); }
  }
  return { banda: banda, industrias: industrias };
}

// guarda la banda salarial de referencia (Config y, si ya existe, el Informe de proceso)
function guardarBandaSalarial(pin, perfilId, texto) {
  pinValido_(pin);
  const cfg = getConfig(perfilId);
  const num = Number(String(texto).replace(/\D/g, ''));
  if (!num || num < 1000) throw new Error('Escribí la banda en pesos, por ejemplo 13000000.');
  setConfigCampo_(perfilId, 'banda_salarial', num);
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(hojaProceso_(empresaKey_(cfg)));
  const nombre = (listarPerfiles().filter(p => p.id === perfilId)[0] || {}).nombre;
  if (sh && sh.getLastRow() > 1) {
    const i = sh.getDataRange().getValues().findIndex(r => r[0] === nombre);
    if (i > 0) sh.getRange(i + 1, 10).setNumberFormat('@').setValue(String(num));
  }
  return { ok: true, banda: String(num) };
}
