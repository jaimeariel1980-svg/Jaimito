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

// banda salarial, industrias y fuentes que ya están cargadas en el Informe de proceso de esta búsqueda
function datosInformeBusqueda(perfilId) {
  const cfg = getConfig(perfilId);
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(hojaProceso_(empresaKey_(cfg)));
  const nombre = (listarPerfiles().filter(p => p.id === perfilId)[0] || {}).nombre;
  if (!sh || sh.getLastRow() < 2) return { banda: '', industrias: '' };
  const f = sh.getDataRange().getValues().filter(r => r[0] === nombre)[0];
  return f ? { banda: String(f[9] || ''), industrias: String(f[13] || '') } : { banda: '', industrias: '' };
}

// guarda la banda salarial de referencia en el Informe de proceso (lo crea si hace falta)
function guardarBandaSalarial(pin, perfilId, texto) {
  pinValido_(pin);
  const cfg = getConfig(perfilId);
  const emp = empresaKey_(cfg);
  const nombre = (listarPerfiles().filter(p => p.id === perfilId)[0] || {}).nombre;
  const num = Number(String(texto).replace(/\D/g, ''));
  if (!num || num < 1000) throw new Error('Escribí la banda en pesos, por ejemplo 13000000.');
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const buscar = () => { const sh = ss.getSheetByName(hojaProceso_(emp)); if (!sh) return null; const d = sh.getDataRange().getValues(); const i = d.findIndex(r => r[0] === nombre); return i > 0 ? { sh: sh, fila: i + 1 } : null; };
  let u = buscar();
  if (!u) { generarInformeProceso(emp); u = buscar(); }
  if (!u) throw new Error('No pude ubicar la búsqueda en el Informe de proceso.');
  u.sh.getRange(u.fila, 10).setNumberFormat('@').setValue(String(num));
  return { ok: true, banda: String(num) };
}
