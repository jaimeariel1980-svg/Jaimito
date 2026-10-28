/* ============================================================
   🏠 INICIO: lo primero que se ve al ingresar (y a donde vuelve «🏠 Inicio» del tablero).
   Todas las búsquedas agrupadas por empresa, con ➕ Nueva búsqueda y la Administración
   (usuarios, copia de seguridad, accesos) fuera de cualquier búsqueda.
   ============================================================ */

const CACHE_CONTEO_ = 'inicio_conteo_v1';

// cantidad de candidatos por pestaña (cache de 2 minutos: es solo un número de orientación)
function conteoCandidatos_(filas) {
  const cache = CacheService.getScriptCache();
  try { const c = cache.get(CACHE_CONTEO_); if (c) return JSON.parse(c); } catch (e) { }
  const out = {};
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  filas.forEach(f => {
    try {
      const sh = ss.getSheetByName(f.hoja);
      const n = sh ? sh.getLastRow() - 1 : 0;
      out[f.id] = n > 0 ? sh.getRange(2, 2, n, 1).getValues().filter(r => String(r[0]).trim()).length : 0;
    } catch (e) { out[f.id] = 0; }
  });
  try { cache.put(CACHE_CONTEO_, JSON.stringify(out), 120); } catch (e) { }
  return out;
}

function inicioDatos_() {
  const rows = configValores_(), H = rows[0];
  const col = k => H.indexOf(k);
  const filas = rows.slice(1).filter(r => r[0] && r[0] !== 'operaciones').map(r => {
    const g = k => { const c = col(k); return c >= 0 ? String(r[c] == null ? '' : r[c]).trim() : ''; };
    return { id: String(r[0]), nombre: g('nombre') || g('nombre_perfil') || String(r[0]), empresa: g('empresa') || 'Sin empresa', hoja: g('hoja_nombre'), cerrada: g('estado_busqueda').toUpperCase() === 'CERRADA' };
  });
  const cnt = conteoCandidatos_(filas);
  filas.forEach(f => { f.n = cnt[f.id] || 0; delete f.hoja; });
  return filas;
}

function inicioHTML_() {
  const datos = inicioDatos_();
  // getRawContent: el HTML tal cual está en el archivo (createHtmlOutputFromFile le saca los comentarios y no se podrían reemplazar los marcadores)
  let html = HtmlService.createTemplateFromFile('InicioPagina').getRawContent();
  const admin = HtmlService.createTemplateFromFile('AdminPanel').getRawContent();
  html = html.replace('%%FUENTE%%', () => FUENTE_MARCA_()).replace('%%DATOS%%', () => JSON.stringify(datos).replace(/</g, '\\u003c'))
    .replace('%%URL%%', () => JSON.stringify(URL_BASE()))
    .replace('%%LOGO%%', () => JSON.stringify(LOGO_BLUMIA_))
    .replace('%%ASSETS%%', () => JSON.stringify(INICIO_ASSETS_))
    .replace('</body>', () => admin + '</body>');
  return HtmlService.createHtmlOutput(html).setTitle('Blumia · Atracción y Elección')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function copiaInfo_() {
  const out = { ultima: PropertiesService.getScriptProperties().getProperty('ULTIMA_COPIA_BLUMIA') || '', copias: [] };
  try {
    const it = DriveApp.getFoldersByName(COPIAS_CARPETA_);
    if (it.hasNext()) {
      const fs = it.next().getFiles(), L = [];
      while (fs.hasNext()) { const f = fs.next(); if (/^Blumia · copia /.test(f.getName())) L.push({ nombre: f.getName().replace(/^Blumia · copia /, ''), t: f.getDateCreated().getTime(), url: f.getUrl() }); }
      L.sort((a, b) => b.t - a.t);
      out.copias = L.slice(0, 14).map(x => ({ nombre: x.nombre, url: x.url }));
    }
  } catch (e) { }
  return out;
}

function listarAccesos_() {
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(HOJA_ACC_);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(sh.getLastRow() - 1, 60);
  return sh.getRange(sh.getLastRow() - n + 1, 1, n, 4).getValues().reverse().map(r => ({
    fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm') : String(r[0]),
    usuario: String(r[1] || ''), evento: String(r[2] || ''), detalle: String(r[3] || '')
  }));
}

// ⏱ Velocidad (Administración): cuánto tarda cada paso de armar el tablero de una búsqueda, para saber dónde se va el tiempo
function diagnosticoVelocidad_(perfilId) {
  CFG_MEMO_ = {}; TIMING_ = []; TIMING_T_ = Date.now();
  const t0 = TIMING_T_;
  let kb = 0, pasos = [];
  try {
    kb = Math.round(kanbanHTML_(perfilId).getContent().length / 1024);
    lapT_('Convertir la página a texto');
  } finally { pasos = TIMING_ || []; TIMING_ = null; CFG_MEMO_ = null; }
  return { pasos: pasos, total: Date.now() - t0, kb: kb, kbDatos: KANBAN_KB_ };
}

// Cerrar o reabrir una búsqueda desde el Inicio (solo administradores, ver registroApi_). Es lo mismo que «Estado de la búsqueda» en ⚙ Configurar → Datos.
function cambiarEstadoBusqueda_(perfilId, estado) {
  const e = String(estado).toUpperCase();
  if (e !== 'ABIERTA' && e !== 'CERRADA') throw new Error('El estado tiene que ser ABIERTA o CERRADA.');
  const cfg = getConfig(perfilId);
  const antes = estadoBusqueda_(cfg);
  if (antes === e) return { ok: true, estado: e };
  setConfigCampo_(perfilId, 'estado_busqueda', e);
  logForm_(SpreadsheetApp.openById(SHEET_MASTER_ID), (typeof SES_ !== 'undefined' && SES_ && SES_.nombre) || 'sin nombre', perfilId, 'Config', { estado_busqueda: antes });
  try { registrarHistorial_('(búsqueda)', String(cfg.nombre || perfilId), e === 'CERRADA' ? 'Búsqueda cerrada' : 'Búsqueda reabierta', 'Desde el Inicio', (typeof SES_ !== 'undefined' && SES_ && SES_.nombre) || 'sistema'); } catch (x) { }
  return { ok: true, estado: e };
}
